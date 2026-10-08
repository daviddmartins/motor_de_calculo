from __future__ import annotations

"""Ajustes V15 dos pareceres Credplan.

Preserva integralmente a V14 aprovada e acrescenta apenas um bloco compacto de
legenda das variáveis matemáticas logo após a seção de fórmulas, com conteúdo
específico para Credplan Fixo e Credplan Variável.
"""

from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, Spacer, Table, TableStyle

from . import pdf_reports_v14 as v14  # aplica SAC explícito e defasagem dinâmica
from . import pdf_reports_v13 as v13  # identidade visual aprovada

_BASE_FORMULA_STORY = v13._formula_story


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
        "v15_legend_symbol",
        fontName="Helvetica-Bold",
        fontSize=7.0,
        leading=8.5,
        textColor=v13.NAVY_DARK,
    )
    meaning_style = v13.ParagraphStyle(
        "v15_legend_meaning",
        fontName="Helvetica",
        fontSize=6.3,
        leading=8.0,
        textColor=v13.TEXT,
    )
    title = Paragraph("Legenda das variáveis", st["subsection"])
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
        Spacer(1, 0.6 * mm),
        title,
        table,
        Spacer(1, 1.1 * mm),
    ])


def _formula_story(tmp_dir, settings, result, st):
    story = list(_BASE_FORMULA_STORY(tmp_dir, settings, result, st))
    story.append(_formula_legend(settings, st))
    return story


# A V14 usa o namespace visual/estrutural da V13 para montar o parecer.
# Substituímos somente a composição da seção de fórmulas; os demais elementos
# permanecem exatamente no padrão V14 aprovado.
v13._formula_story = _formula_story

build_opinion_pdf = v14.build_opinion_pdf

__all__ = ["build_opinion_pdf"]
