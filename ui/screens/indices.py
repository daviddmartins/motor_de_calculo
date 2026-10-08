from __future__ import annotations

from decimal import Decimal
from io import BytesIO
import re

import pandas as pd
import streamlit as st

from services.index_service import (
    ECONOMIC, JUDICIAL, FACTOR, PERCENT,
    IndexRepository, load_index_workbook, repository_to_xlsx_bytes,
    replace_series_values, upsert_series,
)


def _parse_editor_values(df: pd.DataFrame, value_column: str) -> tuple[dict[str, Decimal], list[str]]:
    values: dict[str, Decimal] = {}
    errors: list[str] = []
    for _, row in df.iterrows():
        comp = str(row.get("Competência") or "").strip()
        raw = row.get(value_column)
        if not comp and (raw is None or pd.isna(raw)):
            continue
        if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", comp):
            errors.append(f"Competência inválida: {comp or '-'} (use AAAA-MM).")
            continue
        try:
            text = str(raw).strip().replace(" ", "")
            if "," in text and "." in text:
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", ".")
            value = Decimal(text)
            values[comp] = value
        except Exception:
            errors.append(f"Valor inválido para {comp}.")
    return values, errors


def _series_dataframe(item) -> tuple[pd.DataFrame, str]:
    value_label = "Percentual mensal (%)" if item.value_type == PERCENT else "Fator"
    return pd.DataFrame([
        {"Competência": comp, value_label: float(value)}
        for comp, value in sorted(item.values.items())
    ]), value_label


def render_indices_center(
    *,
    repository: IndexRepository | None,
    source_label: str,
    set_navigation,
    page_intro,
    callout,
) -> None:
    page_intro(
        "Índices e parâmetros",
        "Central de índices",
        "Base única para índices econômicos e tabelas judiciais. Atualizações feitas aqui valem na sessão até a base ser exportada e incorporada ao deploy.",
    )

    if repository is None:
        st.error("A base única de índices não está disponível. Importe um arquivo indices_motor_calculos.xlsx válido abaixo.")
        upload = st.file_uploader("Importar base única de índices", type=["xlsx"], key="indices_bootstrap_upload")
        if upload:
            try:
                load_index_workbook(upload.getvalue())
                st.session_state["active_indices_bytes"] = upload.getvalue()
                st.success("Base validada. Recarregando a Central de índices...")
                st.rerun()
            except Exception as exc:
                st.error(f"Base inválida: {exc}")
        return

    with st.container(border=True):
        c1, c2, c3 = st.columns([1.7, 1.0, 1.0])
        with c1:
            st.caption("BASE EM USO")
            st.markdown(f"**{source_label or repository.source_path or 'Base carregada'}**")
            st.caption(f"{len(repository.active_definitions)} série(s) ativa(s) · {sum(x.record_count for x in repository.active_definitions.values())} registros")
        with c2:
            st.download_button(
                "Exportar base completa",
                data=repository_to_xlsx_bytes(repository),
                file_name="indices_motor_calculos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c3:
            if st.button("Restaurar base do deploy", use_container_width=True):
                st.session_state.pop("active_indices_bytes", None)
                st.rerun()

    upload = st.file_uploader("Importar uma base única atualizada (.xlsx)", type=["xlsx"], key="indices_full_upload")
    if upload:
        try:
            imported = load_index_workbook(upload.getvalue())
            st.session_state["active_indices_bytes"] = upload.getvalue()
            st.success(f"Base validada: {len(imported.active_definitions)} série(s) ativa(s).")
            st.rerun()
        except Exception as exc:
            st.error(f"Não foi possível ativar a base: {exc}")

    tab_overview, tab_series, tab_update, tab_new = st.tabs([
        "Visão geral", "Consultar série", "Atualizar série", "Criar série",
    ])

    with tab_overview:
        overview = pd.DataFrame(repository.status_rows())
        category = st.radio("Categoria", ["Todas", "Econômicos", "Judiciais"], horizontal=True, key="idx_category_filter")
        if category == "Econômicos":
            overview = overview[overview["Categoria"] == ECONOMIC]
        elif category == "Judiciais":
            overview = overview[overview["Categoria"] == JUDICIAL]
        st.dataframe(overview, use_container_width=True, hide_index=True, height=min(620, 80 + max(len(overview), 1) * 36))
        callout(
            "Uma única fonte de dados",
            "TJDFT continua sendo um perfil jurídico no código porque combina séries e vigências; os dados usados por ele (INPC/IPCA) ficam nesta base. TJSP, TJRJ e TJMG entram como séries do tipo FATOR.",
        )

    codes = sorted(repository.definitions)
    with tab_series:
        code = st.selectbox("Série", codes, key="indices_view_code")
        item = repository.get(code)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tipo", "Percentual" if item.value_type == PERCENT else "Fator")
        c2.metric("Primeira competência", item.first_competence)
        c3.metric("Última competência", item.last_competence)
        c4.metric("Registros", item.record_count)
        st.caption(f"Fonte: {item.source or '-'} · Status: {item.status or '-'}")
        df, value_label = _series_dataframe(item)
        st.dataframe(df, use_container_width=True, hide_index=True, height=430)
        st.download_button(
            "Exportar série em CSV",
            data=df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
            file_name=f"{code}.csv",
            mime="text/csv",
        )

    with tab_update:
        callout(
            "Alteração da sessão",
            "A edição abaixo não grava definitivamente no Databricks. Depois de validar, use Exportar base completa e substitua config/indices_motor_calculos.xlsx no deploy.",
            "warning",
        )
        code = st.selectbox("Série a atualizar", codes, key="indices_edit_code")
        item = repository.get(code)
        df, value_label = _series_dataframe(item)
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, hide_index=True, key=f"indices_editor_{code}")
        source_override = st.text_input("Fonte / referência da atualização", value=item.source, key=f"indices_source_{code}")
        if st.button("Aplicar atualização nesta sessão", type="primary", use_container_width=True, key=f"indices_apply_{code}"):
            values, errors = _parse_editor_values(edited, value_label)
            if errors:
                for err in errors:
                    st.error(err)
            else:
                try:
                    updated = replace_series_values(repository, code, values, source=source_override)
                    st.session_state["active_indices_bytes"] = repository_to_xlsx_bytes(updated)
                    st.success(f"{code} atualizado na sessão com {len(values)} competência(s).")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        with st.expander("Importar CSV para substituir a série"):
            csv_upload = st.file_uploader("CSV", type=["csv"], key=f"indices_csv_{code}")
            if csv_upload:
                try:
                    raw = csv_upload.getvalue().decode("utf-8-sig")
                    imported = pd.read_csv(BytesIO(raw.encode("utf-8")), sep=None, engine="python")
                    normalized = {str(c).strip().casefold(): c for c in imported.columns}
                    comp_col = normalized.get("competência") or normalized.get("competencia")
                    candidates = ["percentual mensal (%)", "percentual_mensal_pct", "fator", "valor"]
                    value_col = next((normalized.get(x) for x in candidates if normalized.get(x) is not None), None)
                    if comp_col is None or value_col is None:
                        raise ValueError("O CSV precisa conter Competência e uma coluna de valor (Valor, Fator ou Percentual mensal (%)).")
                    tmp = imported[[comp_col, value_col]].copy()
                    tmp.columns = ["Competência", value_label]
                    values, errors = _parse_editor_values(tmp, value_label)
                    if errors:
                        raise ValueError("; ".join(errors))
                    if st.button("Aplicar CSV nesta sessão", key=f"indices_apply_csv_{code}"):
                        updated = replace_series_values(repository, code, values, source=source_override)
                        st.session_state["active_indices_bytes"] = repository_to_xlsx_bytes(updated)
                        st.success(f"CSV aplicado à série {code}.")
                        st.rerun()
                except Exception as exc:
                    st.error(f"CSV inválido: {exc}")

    with tab_new:
        callout(
            "Cadastro controlado",
            "Criar uma série não altera regras jurídicas. Perfis que combinam índices por vigência continuam sendo parametrizados no código do Motor.",
        )
        c1, c2 = st.columns(2)
        with c1:
            new_code = st.text_input("Código", placeholder="Ex.: INDICE_X")
            new_name = st.text_input("Nome exibido", placeholder="Nome do índice ou tabela")
            new_category = st.selectbox("Categoria", [ECONOMIC, JUDICIAL])
            new_type = st.selectbox("Tipo de valor", [PERCENT, FACTOR])
        with c2:
            new_periodicity = st.text_input("Periodicidade", value="Mensal")
            new_source = st.text_input("Fonte / referência")
            new_status = st.text_input("Status", value="EM VALIDAÇÃO")
            new_active = st.checkbox("Ativo", value=True)
        new_observation = st.text_area("Observação")
        value_label = "Percentual mensal (%)" if new_type == PERCENT else "Fator"
        seed = pd.DataFrame({"Competência": [""], value_label: [None]})
        new_rows = st.data_editor(seed, num_rows="dynamic", use_container_width=True, hide_index=True, key="indices_new_rows")
        if st.button("Adicionar série nesta sessão", type="primary", use_container_width=True):
            values, errors = _parse_editor_values(new_rows, value_label)
            if errors:
                for err in errors:
                    st.error(err)
            else:
                try:
                    updated = upsert_series(
                        repository,
                        code=new_code,
                        name=new_name,
                        category=new_category,
                        value_type=new_type,
                        periodicity=new_periodicity,
                        source=new_source,
                        status=new_status,
                        active=new_active,
                        observation=new_observation,
                        values=values,
                    )
                    st.session_state["active_indices_bytes"] = repository_to_xlsx_bytes(updated)
                    st.success("Série adicionada à base da sessão.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
