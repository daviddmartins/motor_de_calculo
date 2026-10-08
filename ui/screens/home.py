from __future__ import annotations

import html

import pandas as pd
import streamlit as st

from services.health_service import HealthSnapshot, OK, WARNING, ERROR


def _module_card(card: dict, set_navigation) -> None:
    kpi_html = "".join(
        f"<div class='module-kpi-v5'><strong>{html.escape(str(value))}</strong><span>{html.escape(str(label))}</span></div>"
        for value, label in card["kpis"]
    )
    st.markdown(
        f"""
        <div class="module-card-v5">
          <div class="module-card-head-v5">
            <div class="module-icon-v5">{html.escape(card['number'])}</div>
            <div class="module-tag-v5">{html.escape(card['tag'])}</div>
          </div>
          <div class="module-title-v5">{html.escape(card['title'])}</div>
          <div class="module-text-v5">{html.escape(card['desc'])}</div>
          <div class="module-kpi-grid-v5">{kpi_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.button(
        "Acessar módulo →",
        key=f"home_{card['number']}",
        use_container_width=True,
        on_click=set_navigation,
        args=(card["target"], False),
    )


def render_home(
    *,
    set_navigation,
    modality_count: int,
    index_count: int,
    judicial_profile_count: int,
    health: HealthSnapshot | None,
    is_admin: bool,
) -> None:
    """Home deliberadamente enxuta: três módulos + saúde recolhida para administradores."""
    cards = [
        {
            "number": "01",
            "tag": "CONTRATO + PARECER",
            "title": "Evolução de contrato e parecer",
            "desc": "Evolução contratual com memória auditável e confecção do parecer técnico dentro do mesmo fluxo.",
            "target": "Evolução de contrato",
            "kpis": [(str(modality_count), "modalidades"), ("Diária", "memória"), ("PDF + XLSX", "saídas")],
        },
        {
            "number": "02",
            "tag": "MAJS",
            "title": "Recálculo MAJS",
            "desc": "Reconstrução do fluxo teórico, pagamentos, suspensões, amortizações extraordinárias e diferenças atualizadas.",
            "target": "Recálculo de diferenças",
            "kpis": [("2", "modalidades"), ("Pagamentos", "conciliação"), ("PDF + XLSX", "saídas")],
        },
        {
            "number": "03",
            "tag": "JUDICIAL",
            "title": "Atualização do saldo devedor",
            "desc": "Atualização monetária por índices econômicos ou critérios judiciais, com juros, multa e abatimentos quando aplicáveis.",
            "target": "Atualização do saldo devedor",
            "kpis": [(str(judicial_profile_count), "perfis judiciais"), (str(index_count), "séries ativas"), ("PDF + XLSX", "saídas")],
        },
    ]

    cols = st.columns(3, gap="large")
    for col, card in zip(cols, cards):
        with col:
            _module_card(card, set_navigation)

    if not is_admin or health is None:
        return

    icon = "✅" if health.overall == OK else ("⚠️" if health.overall == WARNING else "⛔")
    with st.expander(f"{icon} Saúde do Motor", expanded=False):
        if health.overall == OK:
            st.success("As verificações estruturais do Motor não indicaram pendências de configuração.")
        elif health.overall == WARNING:
            st.warning(f"Foram encontrados {health.warnings} ponto(s) que merecem conferência. Os cálculos continuam disponíveis quando a regra afetada possuir dados suficientes.")
        else:
            st.error(f"Foram encontrados {health.errors} erro(s) estruturais. Ajuste os itens indicados antes de usar as funcionalidades afetadas.")

        st.dataframe(pd.DataFrame(health.as_rows()), use_container_width=True, hide_index=True)

        actions = health.action_items
        if actions:
            st.markdown("**Ajustes recomendados**")
            for idx, item in enumerate(actions):
                c1, c2 = st.columns([3.3, 1.0])
                with c1:
                    st.write(f"**{item.label}** — {item.detail}")
                with c2:
                    st.button(
                        f"Ir para {item.target}",
                        key=f"health_go_{idx}_{item.key}",
                        use_container_width=True,
                        on_click=set_navigation,
                        args=(item.target, False),
                    )
        st.caption("A Central de Saúde verifica apenas dados e configurações que o Motor consegue confirmar. Ela não cria histórico de cálculos nem indicadores fictícios.")
