from __future__ import annotations

import base64
import html
from pathlib import Path
from typing import Iterable, Mapping

import streamlit as st

from .theme import LOGO_BLUE, LOGO_WHITE, STYLE_PATH


def _data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode("ascii")


def apply_design(style_path: Path = STYLE_PATH) -> None:
    css = style_path.read_text(encoding="utf-8") if style_path.exists() else ""
    if css:
        st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def sidebar_brand(version: str) -> None:
    logo_uri = _data_uri(LOGO_WHITE)
    logo = f'<img class="sidebar-brand-logo sidebar-brand-logo-white" src="{logo_uri}" alt="FUNCEF">' if logo_uri else "FUNCEF"
    st.sidebar.markdown(
        f"""
        <div class="sidebar-brand-shell">
          {logo}
          <div class="sidebar-brand-copy">
            <div class="sidebar-brand-title">Motor de Cálculos</div>
            <div class="sidebar-brand-version">Ambiente de validação · v{html.escape(version)}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def brand_header(version: str, *, title: str = "Motor de Cálculos", subtitle: str = "Cálculos contratuais e judiciais com memória técnica auditável") -> None:
    logo_uri = _data_uri(LOGO_WHITE)
    logo = f'<img class="brand-logo-white" src="{logo_uri}" alt="FUNCEF">' if logo_uri else ""
    st.markdown(
        f"""
        <div class="brand-shell-v3">
          <div class="brand-logo-wrap-v3">{logo}</div>
          <div class="brand-copy-v3">
            <div class="brand-kicker-v3">FUNCEF · OPERAÇÕES E CÁLCULOS</div>
            <div class="brand-title-v3">{html.escape(title)}</div>
            <div class="brand-subtitle-v3">{html.escape(subtitle)}</div>
          </div>
          <div class="brand-version-v3"><span>VALIDAÇÃO</span><b>v{html.escape(version)}</b></div>
          <div class="brand-orange-mark-v3"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def page_intro(eyebrow: str, title: str, description: str) -> None:
    st.markdown(
        f"""
        <div class="page-intro-v3">
          <div class="page-intro-accent-v3"></div>
          <div class="page-intro-copy-v3">
            <div class="eyebrow-v3">{html.escape(eyebrow)}</div>
            <div class="page-title-v3">{html.escape(title)}</div>
            <div class="page-desc-v3">{html.escape(description)}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def step_header(number: int, title: str, description: str = "") -> None:
    desc = f'<div class="step-desc-v3">{html.escape(description)}</div>' if description else ""
    st.markdown(
        f"""
        <div class="step-shell-v3">
          <div class="step-number-v3">{number:02d}</div>
          <div class="step-copy-v3"><div class="step-title-v3">{html.escape(title)}</div>{desc}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def workflow_strip(labels: Iterable[str], active: int = 1) -> None:
    parts = []
    for idx, label in enumerate(labels, start=1):
        cls = "workflow-node active" if idx == active else ("workflow-node done" if idx < active else "workflow-node")
        parts.append(
            f'<div class="{cls}"><span>{idx:02d}</span><b>{html.escape(label)}</b></div>'
        )
    st.markdown(f'<div class="workflow-strip">{"".join(parts)}</div>', unsafe_allow_html=True)


def profile_summary_card(*, tribunal: str, title: str, description: str, status: str, source_label: str, active: bool = False) -> None:
    cls = "court-card-v3 active" if active else "court-card-v3"
    st.markdown(
        f"""
        <div class="{cls}">
          <div class="court-card-bar"></div>
          <div class="court-top-v3"><span class="court-badge-v3">{html.escape(tribunal)}</span><span class="court-status-v3">{html.escape(status)}</span></div>
          <div class="court-title-v3">{html.escape(title)}</div>
          <div class="court-desc-v3">{html.escape(description)}</div>
          <div class="court-source-v3">Fonte: {html.escape(source_label)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def rule_timeline(rows: Iterable[Mapping[str, str]]) -> None:
    """Renderiza a linha do tempo sem indentação Markdown.

    O Streamlit pode interpretar HTML multilinha iniciado por espaços como bloco de
    código. Por isso cada card é montado em uma única sequência HTML, sem
    indentação à esquerda.
    """
    blocks: list[str] = []
    for idx, row in enumerate(rows, start=1):
        period = html.escape(str(row.get("period", "")))
        title = html.escape(str(row.get("title", "Regra")))
        meta = html.escape(str(row.get("meta", "")))
        blocks.append(
            f'<div class="timeline-item-v3">'
            f'<div class="timeline-index-v3">{idx:02d}</div>'
            f'<div class="timeline-body-v3">'
            f'<div class="timeline-period-v3">{period}</div>'
            f'<div class="timeline-title-v3">{title}</div>'
            f'<div class="timeline-meta-v3">{meta}</div>'
            f'</div></div>'
        )
    if blocks:
        payload = '<div class="timeline-shell-v3">' + ''.join(blocks) + '</div>'
        st.markdown(payload, unsafe_allow_html=True)


def result_hero(*, label: str, value: str, subtitle: str = "") -> None:
    st.markdown(
        f"""
        <div class="result-hero-v3">
          <div class="result-hero-copy-v3">
            <div class="result-hero-label-v3">{html.escape(label)}</div>
            <div class="result-hero-value-v3">{html.escape(value)}</div>
            <div class="result-hero-sub-v3">{html.escape(subtitle)}</div>
          </div>
          <div class="result-hero-mark-v3"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def copy_block_title(title: str, subtitle: str = "") -> None:
    sub = f'<div class="copy-panel-sub-v3">{html.escape(subtitle)}</div>' if subtitle else ""
    st.markdown(
        f'<div class="copy-panel-title-v3"><b>{html.escape(title)}</b>{sub}</div>',
        unsafe_allow_html=True,
    )


def callout(title: str, text: str, kind: str = "info") -> None:
    cls = {"warning": "callout-v3 warning", "success": "callout-v3 success"}.get(kind, "callout-v3")
    st.markdown(
        f'<div class="{cls}"><b>{html.escape(title)}</b><br>{html.escape(text)}</div>',
        unsafe_allow_html=True,
    )
