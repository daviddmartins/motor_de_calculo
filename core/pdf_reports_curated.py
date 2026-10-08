from __future__ import annotations

from collections import OrderedDict
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from typing import Mapping, Sequence

from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate,
    CondPageBreak,
    Frame,
    KeepTogether,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from . import pdf_reports as base

_BASE_BUILD = base.build_opinion_pdf


# ---------------------------------------------------------------------------
# Curadoria visual dos pareceres Credplan
# ---------------------------------------------------------------------------
# Esta camada parte do modelo já aprovado no Motor e altera somente pontos de
# apresentação: fórmulas, fluxo visual, competência 21-20, paginação dos
# detalhes e alinhamento dos textos. Os cálculos continuam vindo do mesmo
# EvolutionResult produzido pelo motor.


def _styles():
    styles = base._styles()
    # Textos corridos e explicações: alinhamento justificado. Títulos, valores,
    # fórmulas e células numéricas mantêm seus alinhamentos próprios.
    for key in ("body", "body_justified", "step_body", "callout", "note"):
        if key in styles:
            styles[key].alignment = TA_JUSTIFY
    return styles


def _competence_br(value: str) -> str:
    text = str(value or "").strip()
    if len(text) == 7 and text[4] == "-" and text[:4].isdigit() and text[5:].isdigit():
        return f"{text[5:7]}/{text[:4]}"
    return text or "-"


def _segment_summary(settings, result, styles) -> Table:
    first_rows = [row for row in result.daily_rows if row.installment_number == 1]
    grouped: OrderedDict[str, list] = OrderedDict()
    for row in first_rows:
        grouped.setdefault(row.competence, []).append(row)

    headers = [
        "Competência",
        "Dias processados",
        "Taxa diária de juros",
        "Índice de referência",
        "Taxa diária de correção",
    ]
    data = [[Paragraph(h, styles["table_head"]) for h in headers]]

    for competence, rows in grouped.items():
        interest_rate = rows[0].daily_interest_rate if rows else Decimal("0")
        correction_rates = [r.daily_correction_rate for r in rows if r.daily_correction_rate != 0]
        correction_rate = correction_rates[0] if correction_rates else Decimal("0")
        references = list(dict.fromkeys(r.index_reference for r in rows if r.index_reference))
        reference = " a ".join(_competence_br(item) for item in references) if references else "Não se aplica"
        data.append([
            Paragraph(_competence_br(competence), styles["table"]),
            Paragraph(str(len(rows)), styles["table"]),
            Paragraph(base._percent_br(interest_rate, settings.rounding.percentage_display_places), styles["table_right"]),
            Paragraph(reference, styles["table"]),
            Paragraph(base._percent_br(correction_rate, settings.rounding.percentage_display_places), styles["table_right"]),
        ])

    table = Table(
        data,
        colWidths=[29 * mm, 25 * mm, 39 * mm, 43 * mm, 40 * mm],
        repeatRows=1,
        splitByRow=1,
    )
    commands = [
        ("BACKGROUND", (0, 0), (-1, 0), base.PDF_BLUE),
        ("GRID", (0, 0), (-1, -1), 0.35, base.PDF_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3.2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.2),
    ]
    for idx in range(1, len(data)):
        if idx % 2 == 0:
            commands.append(("BACKGROUND", (0, idx), (-1, idx), HexColor("#F2F6FA")))
    table.setStyle(TableStyle(commands))
    return table


def _neutral_note(title: str, text: str, styles) -> Table:
    icon = base._approved_icon("info", 9 * mm)
    content = Paragraph(f"<b>{title}</b><br/>{text}", styles["callout"])
    table = Table([[icon, content]], colWidths=[12 * mm, base.CONTENT_WIDTH - 12 * mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), base.PDF_BLUE_LIGHT),
        ("BOX", (0, 0), (-1, -1), 0.45, base.PDF_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return table


def _process_strip(settings, styles) -> Table:
    """Fluxo sem cards coloridos: branco/azul claro e uma única hierarquia."""
    if settings.modality_code == "MOD_002":
        items = [
            ("money", "SALDO\nINICIAL"),
            ("percent", "JUROS DO\nPERÍODO"),
            ("target", "VF / SALDO\nANTES"),
            ("date", "PRESTAÇÃO\nPRICE"),
            ("chart", "AMORTIZAÇÃO"),
            ("money", "SALDO PRINCIPAL\nFINAL"),
        ]
        box_width = 23.5 * mm
        arrow_width = 5.5 * mm
    else:
        items = [
            ("money", "SALDO\nINICIAL"),
            ("percent", "JUROS DO\nDIA"),
            ("index", "CORREÇÃO EM\nDIA ÚTIL"),
            ("calendar", "EVENTO /\nAMORTIZAÇÃO"),
            ("chart", "SALDO\nFINAL"),
        ]
        box_width = 28 * mm
        arrow_width = 7 * mm

    cells = []
    widths = []
    for idx, (kind, label) in enumerate(items):
        txt = Paragraph(
            label.replace("\n", "<br/>"),
            ParagraphStyle(
                f"process_label_{settings.modality_code}_{idx}",
                fontName="Helvetica-Bold",
                fontSize=6.0,
                leading=7.1,
                textColor=base.PDF_BLUE,
                alignment=TA_CENTER,
            ),
        )
        box = Table(
            [[base._approved_icon(kind, 8 * mm)], [txt]],
            colWidths=[box_width],
            rowHeights=[10 * mm, 11.5 * mm],
        )
        box.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), base.WHITE),
            ("BOX", (0, 0), (-1, -1), 0.55, base.PDF_LINE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ]))
        cells.append(box)
        widths.append(box_width)
        if idx < len(items) - 1:
            cells.append(Paragraph(
                "→",
                ParagraphStyle(
                    f"process_arrow_{settings.modality_code}_{idx}",
                    fontName="Helvetica-Bold",
                    fontSize=11,
                    textColor=base.PDF_MUTED,
                    alignment=TA_CENTER,
                ),
            ))
            widths.append(arrow_width)

    table = Table([cells], colWidths=widths, hAlign="CENTER")
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return table


def _variable_method_steps(settings, result, styles) -> list:
    steps = [
        ("Identificar a competência", "A evolução começa no dia seguinte ao crédito e fecha na referência da prestação. A competência aplicável é determinada conforme a regra operacional da modalidade."),
        ("Converter a taxa", "A taxa anual é convertida em mensal equivalente e depois em taxa diária equivalente conforme os dias corridos da competência."),
        ("Apropriar juros diários", "Os juros remuneratórios são calculados diariamente sobre o saldo de abertura de cada dia."),
        ("Aplicar a correção monetária", "O INPC, observada a defasagem parametrizada, é diarizado pelos dias úteis da competência e aplicado após os juros do dia."),
        ("Processar eventos", "Eventuais amortizações extraordinárias são processadas após os encargos do dia e reduzem o saldo da etapa."),
        ("Apurar saldo e base", "No vencimento, apura-se o saldo antes da prestação e segregam-se os juros acumulados para formar a base principal de amortização."),
        ("Calcular amortização e prestação", "A base é dividida pelo número de prestações remanescentes. A prestação é formada pela amortização e pelos juros do período."),
        ("Fechar o saldo principal", "O saldo final do principal corresponde à base de amortização menos a amortização do ciclo e inicia a etapa seguinte."),
    ]
    cards = [base._step_card(i + 1, title, text, styles, width=85 * mm) for i, (title, text) in enumerate(steps)]
    grid = Table(
        [[cards[0], cards[1]], [cards[2], cards[3]], [cards[4], cards[5]], [cards[6], cards[7]]],
        colWidths=[88 * mm, 88 * mm],
        hAlign="CENTER",
    )
    grid.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
    ]))
    return [grid]


def _formula_card(number: int, title: str, formula: str, definitions, note: str, styles, width: float = 85.5 * mm):
    return base._approved_formula_card(number, title, formula, definitions, note, styles, width=width)


def _variable_formulas(settings, result, styles) -> list:
    first = result.installments[0]
    cards = [
        _formula_card(1, "Conversão da taxa anual em mensal equivalente", "i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1", [("i<sub>a</sub>", "taxa anual contratual"), ("i<sub>m</sub>", "taxa mensal equivalente")], f"taxa anual {base._percent_br(settings.annual_interest_rate, settings.rounding.percentage_display_places)} convertida para {base._percent_br(result.monthly_interest_rate, settings.rounding.percentage_display_places)}.", styles),
        _formula_card(2, "Taxa diária de juros", "i<sub>d</sub> = (1 + i<sub>m</sub>)<super>1/DC</super> - 1", [("i<sub>d</sub>", "taxa diária equivalente"), ("i<sub>m</sub>", "taxa mensal equivalente"), ("DC", "dias corridos da competência")], "a taxa diária é recalculada conforme os dias corridos de cada competência.", styles),
        _formula_card(3, "Correção monetária diária", "c<sub>d</sub> = (1 + c<sub>m</sub>)<super>1/DU</super> - 1", [("c<sub>d</sub>", "taxa diária de correção"), ("c<sub>m</sub>", "INPC mensal da referência"), ("DU", "dias úteis da competência")], f"índice INPC, com defasagem de {int(settings.index_lag_months or 0)} meses.", styles),
        _formula_card(4, "Juros do dia e juros do período", "J<sub>d</sub> = SD<sub>0</sub> × i<sub>d</sub><br/>J<sub>período</sub> = Σ J<sub>d</sub>", [("J<sub>d</sub>", "juros do dia"), ("SD<sub>0</sub>", "saldo de abertura do dia"), ("i<sub>d</sub>", "taxa diária de juros"), ("J<sub>período</sub>", "somatório dos juros diários")], f"o somatório da primeira prestação resultou em {base._money_br(first.interest_amount)}.", styles),
        _formula_card(5, "Evolução diária do saldo", "SD<sub>J</sub> = SD<sub>0</sub> × (1 + i<sub>d</sub>)<br/>SD<sub>J+CM</sub> = SD<sub>J</sub> × (1 + c<sub>d</sub>)", [("SD<sub>0</sub>", "saldo no início do dia"), ("SD<sub>J</sub>", "saldo após juros"), ("SD<sub>J+CM</sub>", "saldo após juros e correção"), ("c<sub>d</sub>", "taxa diária de correção")], "juros todos os dias; correção monetária apenas nos dias úteis.", styles),
        _formula_card(6, "Amortização e prestação do ciclo", "Base<sub>A</sub> = SD<sub>antes</sub> - J<sub>período</sub><br/>A = Base<sub>A</sub> / n<br/>P = A + J<sub>período</sub><br/>SD<sub>final</sub> = Base<sub>A</sub> - A", [("Base<sub>A</sub>", "base principal sem juros do período"), ("A", "amortização"), ("n", "prestações remanescentes"), ("P", "prestação")], f"1ª prestação: base {base._money_br(first.amortization_base)}, amortização {base._money_br(first.regular_amortization)}, prestação {base._money_br(first.installment_amount)}.", styles),
    ]
    grid = Table(
        [[cards[0], cards[1]], [cards[2], cards[3]], [cards[4], cards[5]]],
        colWidths=[88 * mm, 88 * mm],
        hAlign="CENTER",
    )
    grid.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    return [grid]


def _fixed_formulas(settings, result, styles) -> list:
    first = result.installments[0]
    monthly_pct = base._percent_br(result.monthly_interest_rate, settings.rounding.percentage_display_places)
    annual_pct = base._percent_br(settings.annual_interest_rate, settings.rounding.percentage_display_places)
    timing_factor = " × [1 / (1 + i<sub>m</sub>)]" if settings.payment_timing == 1 else ""
    timing_numeric = f" × [1 / (1 + {monthly_pct})]" if settings.payment_timing == 1 else ""
    payment_note = (
        "pagamento no início do período (tipo 1)" if settings.payment_timing == 1
        else "pagamento no fim do período (tipo 0)"
    )
    cards = [
        _formula_card(1, "Taxa mensal equivalente", "i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1", [("i<sub>m</sub>", "taxa mensal equivalente"), ("i<sub>a</sub>", "taxa anual do contrato")], f"i_m = (1 + {annual_pct})^(1/12) - 1 = {monthly_pct}.", styles),
        _formula_card(2, "Evolução da carência até a primeira prestação", "VF = PV × ∏ (1 + i<sub>m</sub>)<super>d/DC</super>", [("VF", "valor futuro na primeira prestação"), ("PV", "saldo original na data do crédito"), ("d", "dias efetivamente evoluídos"), ("DC", "dias corridos da competência")], f"PV = {base._money_br(settings.initial_balance)}; de {base._date_br(settings.credit_date)} até {base._date_br(first.reference_due_date)}; VF = {base._money_br(first.balance_before_installment)}.", styles),
        _formula_card(3, "Prestação pela Tabela Price", f"PMT = [VF × i<sub>m</sub> / (1 - (1 + i<sub>m</sub>)<super>-n</super>)]{timing_factor}", [("PMT", "prestação fixa"), ("VF", "saldo na data da primeira prestação"), ("i<sub>m</sub>", "taxa mensal equivalente"), ("n", "número de prestações")], f"{payment_note}. Substituição: [{base._money_br(first.balance_before_installment)} × {monthly_pct} / (1 - (1 + {monthly_pct})^(-{settings.term}))]{timing_numeric} = {base._money_br(first.installment_amount)}.", styles),
        _formula_card(4, "Composição da prestação e saldo principal", "J<sub>1</sub> = Σ J<sub>d</sub><br/>A<sub>1</sub> = PMT - J<sub>1</sub><br/>SD<sub>final</sub> = Base<sub>principal</sub> - A<sub>1</sub>", [("J<sub>1</sub>", "juros da primeira prestação"), ("A<sub>1</sub>", "amortização da primeira prestação"), ("Base<sub>principal</sub>", "saldo sem os juros do período"), ("SD<sub>final</sub>", "saldo principal após a amortização")], f"J_1 = {base._money_br(first.interest_amount)}; A_1 = {base._money_br(first.regular_amortization)}; saldo principal final = {base._money_br(first.closing_balance)}.", styles),
    ]
    grid = Table(
        [[cards[0], cards[1]], [cards[2], cards[3]]],
        colWidths=[88 * mm, 88 * mm],
        hAlign="CENTER",
    )
    grid.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    return [grid]


def _neutral_equation_line(text: str, styles) -> Table:
    table = Table([[Paragraph(text, styles["formula_small"])]], colWidths=[base.CONTENT_WIDTH - 10 * mm])
    table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.35, base.PDF_LINE),
        ("BACKGROUND", (0, 0), (-1, -1), base.WHITE),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 2.4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4),
    ]))
    return table


def _detail_title(row, styles, idx: int) -> Table:
    p = Paragraph(
        f"Prestação {row.installment_number} - referência {base._date_br(row.reference_due_date)}",
        ParagraphStyle(
            f"curated_detail_title_{idx}_{row.installment_number}",
            fontName="Helvetica-Bold",
            fontSize=11.1,
            leading=13.0,
            textColor=base.PDF_BLUE,
            spaceAfter=0,
        ),
    )
    table = Table([[p]], colWidths=[base.CONTENT_WIDTH])
    table.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.45, base.PDF_LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return table


def _detail_summary(row, overrides, places, styles, fixed: bool) -> Table:
    if fixed:
        fields = [
            ("Saldo antes", base._value(row, "balance_before_installment", overrides)),
            ("Juros do período", base._value(row, "interest_amount", overrides)),
            ("Base principal", Decimal(str(row.opening_balance))),
            ("Prestação Price", base._value(row, "installment_amount", overrides)),
            ("Amortização", base._value(row, "regular_amortization", overrides)),
            ("FGQC", base._value(row, "fgqc_amount", overrides)),
            ("Correção monetária", "Não se aplica"),
            ("Saldo final", base._value(row, "closing_balance", overrides)),
        ]
    else:
        fields = [
            ("Saldo antes", base._value(row, "balance_before_installment", overrides)),
            ("Juros do período", base._value(row, "interest_amount", overrides)),
            ("Correção monetária", base._value(row, "correction_amount", overrides)),
            ("Base de amortização", Decimal(str(row.amortization_base))),
            ("Amortização", base._value(row, "regular_amortization", overrides)),
            ("Prestação", base._value(row, "installment_amount", overrides)),
            ("FGQC", base._value(row, "fgqc_amount", overrides)),
            ("Saldo final", base._value(row, "closing_balance", overrides)),
        ]

    label_style = ParagraphStyle("curated_summary_label", fontName="Helvetica", fontSize=6.3, leading=7.5, textColor=base.PDF_TEXT)
    value_style = ParagraphStyle("curated_summary_value", fontName="Helvetica-Bold", fontSize=6.4, leading=7.6, textColor=base.PDF_BLUE, alignment=TA_RIGHT)

    rows = []
    for idx in range(0, 8, 2):
        left_label, left_value = fields[idx]
        right_label, right_value = fields[idx + 1]
        def fmt(value):
            return base._money_br(value, places) if isinstance(value, (Decimal, int, float)) else str(value)
        rows.append([
            Paragraph(left_label, label_style), Paragraph(fmt(left_value), value_style),
            Paragraph(right_label, label_style), Paragraph(fmt(right_value), value_style),
        ])

    table = Table(rows, colWidths=[34 * mm, 50 * mm, 34 * mm, 50 * mm])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.35, base.PDF_LINE),
        ("BACKGROUND", (0, 0), (-1, -1), base.WHITE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return table


def _detail_block(row, overrides: Mapping[int, Mapping[str, Decimal]], settings, styles, idx: int) -> list:
    places = settings.rounding.money_places
    fixed = settings.modality_code == "MOD_002"
    interest = base._value(row, "interest_amount", overrides)
    correction = base._value(row, "correction_amount", overrides)
    amort = base._value(row, "regular_amortization", overrides)
    installment = base._value(row, "installment_amount", overrides)
    closing = base._value(row, "closing_balance", overrides)
    before = base._value(row, "balance_before_installment", overrides)
    fgqc = base._value(row, "fgqc_amount", overrides)
    opening = Decimal(str(row.opening_balance))
    extra = Decimal(str(row.extraordinary_amortization))
    remaining = row.remaining_installments + 1

    if fixed:
        base_principal = before - interest
        narrative = (
            f"O saldo acumulado antes da prestação totalizou <b>{base._money_br(before, places)}</b>. Nesse montante estão refletidos os juros do período de "
            f"<b>{base._money_br(interest, places)}</b>. Para identificar o principal antes da amortização, os juros são segregados da base, obtendo-se "
            f"<b>{base._money_br(base_principal, places)}</b>. A prestação fixa pela Tabela Price totalizou <b>{base._money_br(installment, places)}</b>; a amortização, "
            f"obtida pela diferença entre a prestação e os juros, foi de <b>{base._money_br(amort, places)}</b>. Somente essa amortização reduz o principal, que passou a "
            f"<b>{base._money_br(closing, places)}</b>. O FGQC de <b>{base._money_br(fgqc, places)}</b> permanece demonstrado separadamente. Não há correção monetária."
        )
        calc = [
            ("Saldo antes da prestação", "saldo principal do ciclo + juros do período", f"{base._money_br(base_principal, places)} + {base._money_br(interest, places)}", before),
            ("Base principal", "saldo antes - juros do período", f"{base._money_br(before, places)} - {base._money_br(interest, places)}", base_principal),
            ("Prestação Price", "prestação fixa apurada pela Tabela Price", base._money_br(installment, places), installment),
            ("Amortização", "prestação Price - juros do período", f"{base._money_br(installment, places)} - {base._money_br(interest, places)}", amort),
            ("Saldo principal final", "base principal - amortização", f"{base._money_br(base_principal, places)} - {base._money_br(amort, places)}", closing),
        ]
        equations = [
            f"J_período = Σ Jd = {base._money_br(interest, places)}",
            f"A = PMT - J_período = {base._money_br(installment, places)} - {base._money_br(interest, places)} = {base._money_br(amort, places)}",
            f"SD_principal_final = Base_principal - A = {base._money_br(base_principal, places)} - {base._money_br(amort, places)} = {base._money_br(closing, places)}",
        ]
    else:
        amort_base = Decimal(str(row.amortization_base))
        fgqc_text = f" O FGQC informado foi de <b>{base._money_br(fgqc, places)}</b> e permanece demonstrado separadamente." if fgqc != 0 else ""
        extra_text = f" Também houve amortização extraordinária de <b>{base._money_br(extra, places)}</b>." if extra != 0 else ""
        narrative = (
            f"O saldo acumulado antes da prestação totalizou <b>{base._money_br(before, places)}</b>. Nesse montante estão refletidos os juros do período de "
            f"<b>{base._money_br(interest, places)}</b> e a correção monetária de <b>{base._money_br(correction, places)}</b>. Para apurar a amortização, os juros do período "
            f"são excluídos da base, obtendo-se <b>{base._money_br(amort_base, places)}</b>. A divisão pelo prazo remanescente resultou em amortização de "
            f"<b>{base._money_br(amort, places)}</b>. A prestação corresponde à soma da amortização com os juros e totalizou <b>{base._money_br(installment, places)}</b>. "
            f"Após o pagamento, o saldo principal passou a <b>{base._money_br(closing, places)}</b>.{fgqc_text}{extra_text}"
        )
        calc = [
            ("Saldo antes da prestação", "saldo inicial do ciclo + juros do período + correção monetária", f"{base._money_br(opening, places)} + {base._money_br(interest, places)} {base._signed_money(correction, places)}", before),
            ("Base de amortização", "saldo antes - juros do período", f"{base._money_br(before, places)} - {base._money_br(interest, places)}", amort_base),
            ("Amortização", "Base_A / prestações remanescentes", f"{base._money_br(amort_base, places)} / {remaining}", amort),
            ("Prestação", "amortização + juros do período", f"{base._money_br(amort, places)} + {base._money_br(interest, places)}", installment),
            ("Saldo final", "Base_A - amortização", f"{base._money_br(amort_base, places)} - {base._money_br(amort, places)}", closing),
        ]
        equations = [
            f"Base_A = {base._money_br(before, places)} - {base._money_br(interest, places)} = {base._money_br(amort_base, places)}",
            f"A = {base._money_br(amort_base, places)} / {remaining} = {base._money_br(amort, places)}",
            f"P = {base._money_br(amort, places)} + {base._money_br(interest, places)} = {base._money_br(installment, places)}",
        ]

    desc_style = ParagraphStyle(
        f"curated_detail_desc_{idx}",
        fontName="Helvetica",
        fontSize=6.2,
        leading=7.5,
        textColor=base.PDF_TEXT,
    )
    result_style = ParagraphStyle(
        f"curated_detail_result_{idx}",
        fontName="Helvetica-Bold",
        fontSize=6.4,
        leading=7.6,
        textColor=base.PDF_BLUE,
        alignment=TA_RIGHT,
    )
    data = [[
        Paragraph("Etapa", styles["table_head"]),
        Paragraph("Descrição", styles["table_head"]),
        Paragraph("Substituição numérica", styles["table_head"]),
        Paragraph("Resultado", styles["table_head"]),
    ]]
    for number, (label, rule, substitution, value) in enumerate(calc, 1):
        badge = Table([[Paragraph(
            str(number),
            ParagraphStyle(
                f"curated_badge_{idx}_{number}",
                fontName="Helvetica-Bold",
                fontSize=6.5,
                textColor=base.WHITE,
                alignment=TA_CENTER,
            ),
        )]], colWidths=[7 * mm], rowHeights=[7 * mm])
        badge.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), base.PDF_BLUE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        data.append([
            badge,
            Paragraph(f"<b>{label}</b><br/><font color='#667A8D'>{rule}</font>", desc_style),
            Paragraph(substitution, desc_style),
            Paragraph(base._money_br(value, places), result_style),
        ])

    calc_table = Table(
        data,
        colWidths=[14 * mm, 58 * mm, 70 * mm, 34 * mm],
        repeatRows=1,
        splitByRow=1,
    )
    commands = [
        ("BACKGROUND", (0, 0), (-1, 0), base.PDF_BLUE),
        ("GRID", (0, 0), (-1, -1), 0.35, base.PDF_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]
    for row_idx in range(1, len(data)):
        if row_idx % 2 == 0:
            commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), HexColor("#F2F6FA")))
    calc_table.setStyle(TableStyle(commands))

    parts = [
        _detail_title(row, styles, idx),
        Spacer(1, 1.8 * mm),
        Paragraph(narrative, styles["body_justified"]),
        Spacer(1, 1.4 * mm),
        calc_table,
        Spacer(1, 1.8 * mm),
        Paragraph(f"Resumo da Prestação {row.installment_number}", styles["subsection"]),
        _detail_summary(row, overrides, places, styles, fixed),
        Spacer(1, 1.4 * mm),
        _neutral_equation_line(equations[0], styles),
        Spacer(1, 0.6 * mm),
        _neutral_equation_line(equations[1], styles),
        Spacer(1, 0.6 * mm),
        _neutral_equation_line(equations[2], styles),
    ]
    return parts


def _build_story(settings, result, modality_name, identity, selected_rows, overrides, styles):
    first = result.installments[0]
    story = [Spacer(1, 1.5 * mm)]

    story.append(KeepTogether([
        base._section_title("01", "Objeto e parâmetros contratuais", None, styles),
        Spacer(1, 2 * mm),
        base._object_text(identity, settings, modality_name, styles),
    ]))
    story.extend([
        Spacer(1, 2.2 * mm),
        base._summary_cards(settings, result, modality_name, styles),
        Spacer(1, 3.6 * mm),
        base._parameters_table(identity, settings, result, styles),
    ])

    foundation = base._technical_foundation(settings, styles)
    story.extend([Spacer(1, 3.5 * mm), CondPageBreak(38 * mm)])
    story.append(KeepTogether([
        base._section_title("02", "Fundamentação técnico-financeira", None, styles),
        Spacer(1, 2 * mm),
        foundation[0],
    ]))
    story.extend(foundation[1:])

    story.extend([
        Spacer(1, 3.5 * mm),
        CondPageBreak(44 * mm),
        base._section_title("03", "Metodologia aplicada - passo a passo", None, styles),
        Spacer(1, 2 * mm),
    ])
    if settings.modality_code == "MOD_001":
        story.extend(_variable_method_steps(settings, result, styles))
    else:
        story.extend(base._method_steps(settings, result, styles))
    story.extend([Spacer(1, 2.4 * mm), CondPageBreak(28 * mm), _process_strip(settings, styles)])

    story.extend([
        Spacer(1, 3.5 * mm),
        CondPageBreak(44 * mm),
        base._section_title("04", "Fórmulas e aplicação numérica", None, styles),
        Spacer(1, 2 * mm),
    ])
    if settings.modality_code == "MOD_001":
        story.extend(_variable_formulas(settings, result, styles))
    elif settings.modality_code == "MOD_002":
        story.extend(_fixed_formulas(settings, result, styles))
    else:
        story.extend(base._formula_application(settings, result, styles))

    if settings.modality_code in {"MOD_001", "MOD_002"}:
        story.extend([
            Spacer(1, 2.4 * mm),
            Paragraph("Competências e taxas efetivamente utilizadas na primeira prestação", styles["subsection"]),
            _segment_summary(settings, result, styles),
            Spacer(1, 1.5 * mm),
        ])
        if settings.modality_code == "MOD_001":
            story.append(_neutral_note(
                "Nota sobre a competência 21-20",
                "No ciclo 21-20, a competência é identificada pelo mês em que se inicia. Assim, a competência dezembro corresponde ao período de 21/12 a 20/01. Com defasagem de 2 meses, utiliza-se o INPC de outubro, diarizado pelos dias úteis da competência de aplicação. Os juros são apropriados diariamente ao longo do ciclo; a correção monetária é aplicada somente em dias úteis.",
                styles,
            ))
        else:
            story.append(_neutral_note(
                "Leitura do quadro de competências",
                "O Credplan Fixo não utiliza correção monetária. O quadro demonstra apenas os segmentos de competência efetivamente processados e as taxas diárias de juros utilizadas na evolução até a prestação de referência.",
                styles,
            ))

    story.extend([
        Spacer(1, 3.2 * mm),
        CondPageBreak(40 * mm),
        base._section_title("05", "Demonstração da evolução", None, styles),
        Spacer(1, 2 * mm),
        Paragraph("Prestações de referência", styles["subsection"]),
        base._installments_table(selected_rows, overrides, settings.rounding.money_places, styles),
    ])

    for idx, row in enumerate(selected_rows[:2]):
        # O bloco completo de cada prestação deve permanecer visualmente íntegro.
        # Se não houver área suficiente, ele começa na página seguinte; não se
        # comprime tabela/summary apenas para acomodar a conclusão.
        story.extend([Spacer(1, 4 * mm), CondPageBreak(118 * mm)])
        story.extend(_detail_block(row, overrides, settings, styles, idx))

    conclusion = base._conclusion(settings, first, overrides, styles)
    story.extend([Spacer(1, 4 * mm), CondPageBreak(52 * mm)])
    story.append(KeepTogether([
        base._section_title("06", "Conclusão técnica", None, styles),
        Spacer(1, 2 * mm),
        conclusion[0],
    ]))
    story.extend(conclusion[1:])

    if identity.additional_note.strip():
        story.extend([
            Spacer(1, 3 * mm),
            Paragraph("Observação da elaboração", styles["subsection"]),
            Paragraph(identity.additional_note.strip().replace("\n", "<br/>"), styles["body_justified"]),
        ])
    return story


def build_opinion_pdf(
    *,
    settings,
    result,
    modality_name: str,
    identity,
    selected_installment_numbers: Sequence[int],
    overrides: Mapping[int, Mapping[str, Decimal]] | None = None,
    logo_path: str | Path | None = None,
) -> bytes:
    if not result.installments:
        raise ValueError("Não há prestação concluída para compor o parecer.")

    overrides = overrides or {}
    by_number = {row.installment_number: row for row in result.installments}
    selected_rows = [by_number[number] for number in selected_installment_numbers if number in by_number]
    if not selected_rows:
        raise ValueError("Selecione pelo menos uma prestação concluída para o parecer.")

    # Somente Credplan Variável e Credplan Fixo recebem esta curadoria. As
    # demais modalidades permanecem no gerador consolidado original.
    if settings.modality_code not in {"MOD_001", "MOD_002"}:
        return _BASE_BUILD(
            settings=settings,
            result=result,
            modality_name=modality_name,
            identity=identity,
            selected_installment_numbers=selected_installment_numbers,
            overrides=overrides,
            logo_path=logo_path,
        )

    styles = _styles()
    logo = Path(logo_path) if logo_path else None
    output = BytesIO()
    doc = BaseDocTemplate(
        output,
        pagesize=A4,
        leftMargin=base.PAGE_MARGIN,
        rightMargin=base.PAGE_MARGIN,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=f"Parecer técnico - {modality_name}",
        author="FUNCEF - Motor de Cálculos",
        subject="Evolução contratual",
    )

    first_frame = Frame(
        base.PAGE_MARGIN,
        18 * mm,
        base.CONTENT_WIDTH,
        A4[1] - 18 * mm - 31 * mm,
        id="first_page_curated",
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
    )
    later_frame = Frame(
        base.PAGE_MARGIN,
        18 * mm,
        base.CONTENT_WIDTH,
        A4[1] - 18 * mm - 18 * mm,
        id="later_pages_curated",
        leftPadding=0,
        rightPadding=0,
        topPadding=0,
        bottomPadding=0,
    )
    doc.addPageTemplates([
        PageTemplate(
            id="first_curated",
            frames=[first_frame],
            onPage=lambda canvas, d: base._header_footer(canvas, d, logo, modality_name, identity.contract_number),
            autoNextPageTemplate="later_curated",
        ),
        PageTemplate(
            id="later_curated",
            frames=[later_frame],
            onPage=lambda canvas, d: base._header_footer(canvas, d, logo, modality_name, identity.contract_number),
        ),
    ])

    story = _build_story(settings, result, modality_name, identity, selected_rows, overrides, styles)
    doc.build(story, canvasmaker=base._NumberedCanvas)
    return output.getvalue()
