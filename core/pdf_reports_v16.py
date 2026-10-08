from __future__ import annotations

"""Pareceres Credplan V16.

Preserva integralmente a V14 aprovada (incluindo SAC explícito e defasagem
dinâmica no Variável) e insere a legenda das fórmulas diretamente no story
final do parecer, imediatamente antes do quadro de competências.

A inserção no story final evita depender da ordem de importação/reload das
rotinas internas de fórmulas no ambiente de produção.
"""

from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, Spacer, Table, TableStyle

from . import pdf_reports_v14 as v14
from . import pdf_reports_v13 as v13

_BASE_BUILD_STORY = v14._build_story


def _legend_rows(settings):
    if settings.modality_code == "MOD_002":
        return [
            ("i<sub>a</sub>", "taxa anual nominal", "i<sub>m</sub>", "taxa mensal equivalente"),
            ("PV", "valor presente / saldo inicial", "VF", "saldo evoluído antes da prestação"),
            ("d<sub>k</sub>", "dias corridos do segmento k", "DC<sub>k</sub>", "dias corridos da competência k"),
            ("PMT", "prestação fixa pela Tabela Price", "J<sub>d</sub>", "juros apropriados no dia"),
            ("J<sub>1</sub>", "juros acumulados da prestação", "A<sub>1</sub>", "amortização da prestação"),
            ("SD<sub>1</sub>", "saldo principal após a prestação", "n", "número de prestações do cálculo"),
        ]

    return [
        ("i<sub>a</sub>", "taxa anual nominal", "i<sub>m</sub>", "taxa mensal equivalente"),
        ("i<sub>d</sub>", "taxa diária equivalente de juros", "c<sub>m</sub>", "índice mensal de correção"),
        ("c<sub>d</sub>", "taxa diária equivalente de correção", "DC", "dias corridos da competência"),
        ("DU", "dias úteis da competência", "SD<sub>0</sub>", "saldo no início do dia"),
        ("SD<sub>J</sub>", "saldo após a incidência dos juros", "SD<sub>J+CM</sub>", "saldo após juros e correção"),
        ("J<sub>d</sub>", "juros apropriados no dia", "J<sub>período</sub>", "somatório dos juros do período"),
        ("Base<sub>A</sub>", "base principal de amortização", "A", "amortização do ciclo"),
        ("P", "valor da prestação", "n", "prestações remanescentes"),
    ]


def _formula_legend(settings, st):
    symbol_style = v13.ParagraphStyle(
        "v16_legend_symbol",
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9.2,
        textColor=v13.NAVY_DARK,
    )
    meaning_style = v13.ParagraphStyle(
        "v16_legend_meaning",
        fontName="Helvetica",
        fontSize=7.5,
        leading=9.2,
        textColor=v13.TEXT,
    )

    data = []
    for left_symbol, left_meaning, right_symbol, right_meaning in _legend_rows(settings):
        data.append([
            Paragraph(left_symbol, symbol_style),
            Paragraph(left_meaning, meaning_style),
            Paragraph(right_symbol, symbol_style),
            Paragraph(right_meaning, meaning_style),
        ])

    table = Table(
        data,
        colWidths=[15 * mm, 73.5 * mm, 15 * mm, 73.5 * mm],
        hAlign="LEFT",
    )
    commands = [
        ("BACKGROUND", (0, 0), (-1, -1), v13.BLUE_SOFT2),
        ("BOX", (0, 0), (-1, -1), 0.45, v13.LINE),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#E2EAF2")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4.5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.1),
    ]
    for row_idx in range(len(data)):
        if row_idx % 2:
            commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), v13.WHITE))
    table.setStyle(TableStyle(commands))

    return KeepTogether([
        Paragraph("Legenda das variáveis", st["subsection"]),
        table,
        Spacer(1, 1.1 * mm),
    ])


def _build_story(tmp_dir, settings, result, modality_name, identity, selected_rows, overrides, st):
    story = list(_BASE_BUILD_STORY(
        tmp_dir,
        settings,
        result,
        modality_name,
        identity,
        selected_rows,
        overrides,
        st,
    ))

    marker = "Competências e taxas efetivamente utilizadas na primeira prestação"
    insert_at = None
    for idx, flowable in enumerate(story):
        if isinstance(flowable, Paragraph):
            try:
                text = flowable.getPlainText()
            except Exception:
                text = ""
            if marker in text:
                insert_at = idx
                break

    legend = _formula_legend(settings, st)
    if insert_at is None:
        # Fallback defensivo: a legenda continua sendo incluída no documento,
        # mesmo se o texto-marco mudar futuramente.
        story.append(legend)
    else:
        story.insert(insert_at, legend)

    return story


# O gerador V14/V13 resolve _build_story no namespace do módulo V13.
# Fazemos a substituição no ponto final de montagem do documento para garantir
# que a legenda seja efetivamente renderizada em produção.
v13._build_story = _build_story

build_opinion_pdf = v14.build_opinion_pdf

__all__ = ["build_opinion_pdf"]
