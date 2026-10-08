from __future__ import annotations

import html
from pathlib import Path

import pandas as pd
import streamlit as st

from core.indices import load_configuration


def render_holidays(*, active_cfg, weekdays_pt: list[str], page_intro, callout) -> None:
    page_intro(
        "Calendário",
        "Feriados",
        "Consulte o calendário usado pelos motores para vencimentos e incidências em dias úteis.",
    )
    holiday_df = pd.DataFrame([
        {"Data": d.strftime("%d/%m/%Y"), "Dia da semana": weekdays_pt[d.weekday()]}
        for d in sorted(active_cfg.holidays)
    ])
    if holiday_df.empty:
        st.info("A configuração ativa não possui feriados cadastrados.")
    else:
        st.dataframe(holiday_df, use_container_width=True, hide_index=True, height=520)
    callout(
        "Manutenção",
        "Os feriados continuam na configuração geral porque são parâmetros de calendário, não séries monetárias. A edição persistente pela interface só deve ser ativada após homologação da gravação permanente.",
        "warning",
    )


def render_configuration(
    *,
    active_cfg,
    export_configuration_bytes,
    page_intro,
    callout,
) -> None:
    page_intro(
        "Administração",
        "Configuração do sistema",
        "Modalidades, parâmetros globais e arquivo de configuração geral. Índices possuem base própria na Central de índices.",
    )
    tabs = st.tabs(["Modalidades", "Parâmetros gerais", "Importar / exportar"])
    with tabs[0]:
        st.dataframe(
            pd.DataFrame([{"Código": k, "Nome exibido": v} for k, v in active_cfg.modality_names.items()]),
            use_container_width=True,
            hide_index=True,
        )
    with tabs[1]:
        params_df = pd.DataFrame([{"Parâmetro": k, "Valor": v} for k, v in sorted(active_cfg.parameters.items())])
        if params_df.empty:
            st.info("Nenhum parâmetro global adicional foi encontrado na configuração ativa.")
        else:
            st.dataframe(params_df, use_container_width=True, hide_index=True)
        st.caption("Feriados também pertencem a este arquivo geral, mas possuem uma tela própria para consulta rápida.")
    with tabs[2]:
        conf_upload = st.file_uploader("Carregar configuração nesta sessão", type=["xlsx"], key="config_admin_upload")
        if conf_upload:
            try:
                load_configuration(conf_upload.getvalue())
                st.session_state["active_config_bytes"] = conf_upload.getvalue()
                st.success("Configuração temporária validada e carregada.")
                st.rerun()
            except Exception as exc:
                st.error(f"Arquivo inválido: {exc}")
        st.download_button(
            "Baixar configuração ativa",
            data=export_configuration_bytes(),
            file_name="configuracoes_validacao_atualizada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.info("Base de índices: **config/indices_motor_calculos.xlsx**. Controle de acesso: **config/parametros_acesso_motor_calculos.xlsx**.")
        callout(
            "Persistência",
            "Enquanto a gravação permanente do Databricks App não estiver homologada, preserve alterações exportando os arquivos atualizados e incorporando-os ao próximo deploy.",
            "warning",
        )


def render_access(
    *,
    app_dir: Path,
    access_config: Path,
    registry,
    identity,
    current_user_name: str,
    current_profile: str,
    page_intro,
    callout,
) -> None:
    page_intro(
        "Administração",
        "Acesso ao sistema",
        "Identidade do Databricks combinada com a planilha oficial de acesso do Motor.",
    )
    detected_path = Path(registry.source_path) if registry.source_path else access_config
    try:
        detected_display = str(detected_path.relative_to(app_dir))
    except Exception:
        detected_display = str(detected_path)

    c1, c2, c3 = st.columns(3, gap="medium")
    with c1:
        with st.container(border=True):
            st.caption("USUÁRIO LOGADO")
            st.markdown(f"**{html.escape(current_user_name)}**")
            st.caption(identity.email or identity.preferred_username or "Identidade não exposta")
    with c2:
        with st.container(border=True):
            st.caption("PERFIL NO MOTOR")
            st.markdown(f"**{html.escape(current_profile or 'Não cadastrado')}**")
            st.caption("Definido pela linha correspondente ao e-mail na planilha oficial.")
    with c3:
        with st.container(border=True):
            st.caption("CONTROLE DE ACESSO")
            if registry.enabled:
                st.markdown("**Ativo**")
                st.caption(f"{len(registry.active_entries)} usuário(s) ativo(s).")
            else:
                st.markdown("**Ainda não ativado**")
                st.caption("A planilha localizada não possui Login_Email preenchido.")

    st.markdown("### Planilha em uso")
    st.code(detected_display)
    st.caption("Aba: Controle_De_Acesso · colunas: Login_Email, Nome, Perfil/Pefil e Ativo.")

    if registry.enabled:
        access_rows = [
            {
                "Nome": item.name or "-",
                "Login / e-mail": item.email,
                "Perfil": item.profile,
                "Ativo": "Sim" if item.active else "Não",
            }
            for item in registry.entries if item.email.strip()
        ]
        if access_rows:
            st.dataframe(pd.DataFrame(access_rows), use_container_width=True, hide_index=True)
    else:
        callout(
            "Planilha sem e-mails",
            "O Motor encontrou um arquivo de acesso, mas o controle ainda não foi ativado com Login_Email preenchido.",
            "warning",
        )

    if st.button("Recarregar controle de acesso", use_container_width=True):
        st.rerun()

    st.markdown("### Regras aplicadas")
    st.markdown(
        "- **Elaboração:** preenchida automaticamente com o nome da pessoa logada.\n"
        "- **Validação:** lista com nomes ativos da planilha + **Sem validação**.\n"
        "- **Administrador:** possui também acesso às áreas administrativas.\n"
        "- **Usuário:** acessa a página inicial e os módulos de cálculo."
    )
