from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from typing import Mapping, Sequence

from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    KeepTogether,
    CondPageBreak,
    PageTemplate,
    PageBreak,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from .calendar_rules import calendar_days_in_day_competence
from .models import ContractSettings, EvolutionResult, InstallmentRow, PRICE_MODALITIES
from . import pdf_funcef_template as funcef
from .pdf_theme import (
    PETROLEUM, PETROLEUM_DARK, PETROLEUM_DEEP, TEAL, FUNCEF_BLUE, ORANGE,
    INK, MUTED, LINE, SOFT, WHITE, SUCCESS, WARNING,
)


SOFT_BLUE = HexColor("#EEF5F7")
SOFT_ORANGE = HexColor("#FFF4EA")
SOFT_GREEN = HexColor("#EAF6EF")

PAGE_MARGIN = 17 * mm
CONTENT_WIDTH = A4[0] - (2 * PAGE_MARGIN)
FORMULA_INNER_WIDTH = CONTENT_WIDTH - (12 * mm)
FORMULA_LEGEND_WIDTH = FORMULA_INNER_WIDTH - (4 * mm)

@dataclass(frozen=True)
class ReportIdentity:
    contract_number: str
    participant_name: str
    registration: str
    cpf: str
    request_date: date
    margin_amount: Decimal = Decimal("0")
    fgqc_concession: Decimal = Decimal("0")
    iof_amount: Decimal = Decimal("0")
    administrative_fee: Decimal = Decimal("0")
    settled_loan_balance: Decimal = Decimal("0")
    net_amount: Decimal = Decimal("0")
    elaborator: str = ""
    validator: str = ""
    additional_note: str = ""
    header: funcef.ManifestationHeader | None = None


REPORT_FIELDS = (
    "balance_before_installment",
    "correction_amount",
    "interest_amount",
    "regular_amortization",
    "installment_amount",
    "closing_balance",
    "fgqc_amount",
)


def _decimal(value) -> Decimal:
    return value if isinstance(value, Decimal) else Decimal(str(value))


def _number_br(value, places: int = 2) -> str:
    number = _decimal(value)
    text = f"{number:,.{places}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def _money_br(value, places: int = 2) -> str:
    number = _decimal(value)
    prefix = "-R$ " if number < 0 else "R$ "
    return prefix + _number_br(abs(number), places)


def _percent_br(rate, places: int = 4) -> str:
    return f"{_number_br(_decimal(rate) * Decimal('100'), places)}%"


def _date_br(value: date | datetime | None) -> str:
    if value is None:
        return "-"
    if isinstance(value, datetime):
        value = value.date()
    return value.strftime("%d/%m/%Y")


def _sanitize_filename(text: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in text.strip())
    return cleaned.strip("_") or "parecer"


def pdf_filename(modality_name: str, contract_number: str) -> str:
    return f"Parecer_{_sanitize_filename(modality_name)}_{_sanitize_filename(contract_number)}.pdf"



def _styles():
    sample = getSampleStyleSheet()
    return {
        "body": ParagraphStyle(
            "body", parent=sample["BodyText"], fontName="Helvetica", fontSize=9.0,
            leading=13.2, textColor=INK, spaceAfter=6.2, alignment=TA_LEFT,
        ),
        "body_justified": ParagraphStyle(
            "body_justified", parent=sample["BodyText"], fontName="Helvetica", fontSize=9.0,
            leading=13.3, textColor=INK, spaceAfter=6.2, alignment=4,
        ),
        "body_compact": ParagraphStyle(
            "body_compact", parent=sample["BodyText"], fontName="Helvetica", fontSize=8.45,
            leading=11.8, textColor=INK, spaceAfter=3.5,
        ),
        "small": ParagraphStyle(
            "small", parent=sample["BodyText"], fontName="Helvetica", fontSize=7.3,
            leading=9.8, textColor=MUTED,
        ),
        "micro": ParagraphStyle(
            "micro", parent=sample["BodyText"], fontName="Helvetica", fontSize=6.4,
            leading=8.1, textColor=MUTED,
        ),
        "label": ParagraphStyle(
            "label", parent=sample["BodyText"], fontName="Helvetica-Bold", fontSize=6.8,
            leading=8.2, textColor=MUTED,
        ),
        "card_value": ParagraphStyle(
            "card_value", parent=sample["BodyText"], fontName="Helvetica-Bold", fontSize=11.3,
            leading=13.4, textColor=PETROLEUM_DARK,
        ),
        "hero_title": ParagraphStyle(
            "hero_title", parent=sample["Title"], fontName="Helvetica-Bold", fontSize=18.0,
            leading=20.5, textColor=PETROLEUM_DARK, spaceAfter=2,
        ),
        "hero_subtitle": ParagraphStyle(
            "hero_subtitle", parent=sample["BodyText"], fontName="Helvetica", fontSize=8.7,
            leading=11.2, textColor=MUTED,
        ),
        "section": ParagraphStyle(
            "section", parent=sample["Heading2"], fontName="Helvetica-Bold", fontSize=12.5,
            leading=14.8, textColor=PETROLEUM_DARK, spaceBefore=0, spaceAfter=0,
        ),
        "section_line": ParagraphStyle(
            "section_line", parent=sample["BodyText"], fontName="Helvetica-Bold", fontSize=12.5,
            leading=14.8, textColor=PETROLEUM_DARK, spaceBefore=0, spaceAfter=0,
        ),
        "section_subtitle": ParagraphStyle(
            "section_subtitle", parent=sample["BodyText"], fontName="Helvetica", fontSize=7.3,
            leading=9.8, textColor=MUTED, leftIndent=14 * mm, spaceBefore=1.2, spaceAfter=0,
        ),
        "subsection": ParagraphStyle(
            "subsection", parent=sample["Heading3"], fontName="Helvetica-Bold", fontSize=10.1,
            leading=12.3, textColor=PETROLEUM, spaceBefore=5, spaceAfter=4,
        ),
        "step_number": ParagraphStyle(
            "step_number", parent=sample["BodyText"], fontName="Helvetica-Bold", fontSize=12,
            leading=13, textColor=WHITE, alignment=TA_CENTER,
        ),
        "step_title": ParagraphStyle(
            "step_title", parent=sample["BodyText"], fontName="Helvetica-Bold", fontSize=8.4,
            leading=10.3, textColor=PETROLEUM_DARK,
        ),
        "step_body": ParagraphStyle(
            "step_body", parent=sample["BodyText"], fontName="Helvetica", fontSize=7.35,
            leading=9.6, textColor=INK,
        ),
        "formula": ParagraphStyle(
            "formula", parent=sample["BodyText"], fontName="Helvetica", fontSize=10.0,
            leading=14, textColor=PETROLEUM_DARK, alignment=TA_CENTER,
        ),
        "formula_big": ParagraphStyle(
            "formula_big", parent=sample["BodyText"], fontName="Helvetica", fontSize=11.0,
            leading=15, textColor=PETROLEUM_DARK, alignment=TA_CENTER,
        ),
        "formula_small": ParagraphStyle(
            "formula_small", parent=sample["BodyText"], fontName="Helvetica", fontSize=8.2,
            leading=10.2, textColor=PETROLEUM_DARK, alignment=TA_CENTER,
        ),
        "formula_title": ParagraphStyle(
            "formula_title", parent=sample["BodyText"], fontName="Helvetica-Bold", fontSize=7.5,
            leading=9.0, textColor=INK,
        ),
        "formula_note": ParagraphStyle(
            "formula_note", parent=sample["BodyText"], fontName="Helvetica", fontSize=7.2,
            leading=9.4, textColor=MUTED,
        ),
        "table_head": ParagraphStyle(
            "table_head", parent=sample["BodyText"], fontName="Helvetica-Bold", fontSize=6.3,
            leading=7.3, textColor=WHITE, alignment=TA_CENTER,
        ),
        "table": ParagraphStyle(
            "table", parent=sample["BodyText"], fontName="Helvetica", fontSize=6.5,
            leading=8, textColor=INK,
        ),
        "table_right": ParagraphStyle(
            "table_right", parent=sample["BodyText"], fontName="Helvetica", fontSize=6.5,
            leading=8, textColor=INK, alignment=TA_RIGHT,
        ),
        "note": ParagraphStyle(
            "note", parent=sample["BodyText"], fontName="Helvetica", fontSize=8.0,
            leading=11.2, textColor=WARNING,
        ),
        "callout": ParagraphStyle(
            "callout", parent=sample["BodyText"], fontName="Helvetica", fontSize=8.15,
            leading=11.3, textColor=PETROLEUM_DARK,
        ),
    }

def _value(row: InstallmentRow, field: str, overrides: Mapping[int, Mapping[str, Decimal]]) -> Decimal:
    per_row = overrides.get(row.installment_number, {})
    if field in per_row and per_row[field] is not None:
        return _decimal(per_row[field])
    if field == "fgqc_amount":
        return Decimal("0")
    return _decimal(getattr(row, field))



def _hero(identity: ReportIdentity, settings: ContractSettings, modality_name: str, result: EvolutionResult, styles) -> Table:
    status = "CÁLCULO COMPLETO" if result.is_complete else "CÁLCULO PARCIAL"
    status_bg = SUCCESS if result.is_complete else WARNING
    status_table = Table(
        [[Paragraph(status, ParagraphStyle("hero_status", fontName="Helvetica-Bold", fontSize=7.0, textColor=WHITE, alignment=TA_CENTER))]],
        colWidths=[32 * mm], rowHeights=[6.2 * mm],
    )
    status_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), status_bg),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0, status_bg),
    ]))
    title = Table([
        [Paragraph("PARECER TÉCNICO", styles["hero_title"]), status_table],
        [Paragraph("Evolução contratual e demonstração da metodologia aplicada", styles["hero_subtitle"]), ""],
        [Paragraph(f"{modality_name}  |  Contrato {identity.contract_number}  |  Emitido em {_date_br(date.today())}", styles["small"]), ""],
    ], colWidths=[139 * mm, 34 * mm])
    title.setStyle(TableStyle([
        ("SPAN", (0, 1), (1, 1)),
        ("SPAN", (0, 2), (1, 2)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    shell = Table([[title]], colWidths=[176 * mm])
    shell.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), WHITE),
        ("BOX", (0, 0), (-1, -1), 0.7, LINE),
        ("LINEBEFORE", (0, 0), (0, -1), 4.0, ORANGE),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return shell

def _section_title(number: str, title: str, subtitle: str | None, styles) -> Table:
    headline = Paragraph(
        f'<font color="#FF762B"><b>{number}</b></font>&nbsp;&nbsp;{title}',
        styles["section_line"],
    )
    rows = [[headline]]
    if subtitle:
        rows.append([Paragraph(subtitle, styles["section_subtitle"])])
    table = Table(rows, colWidths=[CONTENT_WIDTH])
    commands = [
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 3.2),
        ("LINEBELOW", (0, 0), (-1, 0), 0.65, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    if subtitle:
        commands.extend([
            ("TOPPADDING", (0, 1), (-1, 1), 2),
            ("BOTTOMPADDING", (0, 1), (-1, 1), 0),
        ])
    table.setStyle(TableStyle(commands))
    return table


def _metric_card(label: str, value: str, accent, styles, note: str = "") -> Table:
    rows = [
        [Paragraph(label.upper(), styles["label"])],
        [Paragraph(value, styles["card_value"])],
    ]
    if note:
        rows.append([Paragraph(note, styles["small"])])
    card = Table(rows, colWidths=[54 * mm], rowHeights=[6 * mm, 11 * mm] + ([6 * mm] if note else []))
    card.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), WHITE),
        ("BOX", (0, 0), (-1, -1), 0.7, LINE),
        ("LINEBEFORE", (0, 0), (0, -1), 3.2, accent),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return card


def _callout(title: str, text: str, styles, accent=TEAL, background=SOFT_BLUE) -> Table:
    content = Paragraph(f"<b>{title}</b><br/>{text}", styles["callout"])
    table = Table([[content]], colWidths=[CONTENT_WIDTH])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), background),
        ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("LINEBEFORE", (0, 0), (0, -1), 3.4, accent),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return table


def _identification_tables(identity: ReportIdentity, settings: ContractSettings, modality_name: str, styles) -> Table | None:
    # LGPD / modo web: dados pessoais e responsáveis não são renderizados em PDFs.
    return None


def _amortization_system_label(settings: ContractSettings) -> str:
    return "Tabela Price" if settings.modality_code in PRICE_MODALITIES else "SAC"


def _summary_cards(settings: ContractSettings, result: EvolutionResult, modality_name: str, styles) -> Table:
    first = result.installments[0]
    index_value = settings.index_code if settings.index_code else "Sem correção monetária"
    values = [
        ("Modalidade", modality_name),
        ("Sistema de amortização", _amortization_system_label(settings)),
        ("Taxa de juros (a.a.)", _percent_br(settings.annual_interest_rate, settings.rounding.percentage_display_places)),
        ("Prazo total", f"{settings.term} prestações"),
        ("Data da 1ª prestação", _date_br(first.reference_due_date)),
        ("Índice de correção", index_value),
    ]
    label_style = ParagraphStyle(
        "summary_card_label", parent=styles["small"], fontName="Helvetica-Bold",
        fontSize=5.8, leading=7.1, textColor=MUTED, alignment=TA_CENTER,
    )
    value_style = ParagraphStyle(
        "summary_card_value", parent=styles["body"], fontName="Helvetica-Bold",
        fontSize=8.4, leading=10.2, textColor=PETROLEUM_DARK, alignment=TA_CENTER,
    )
    cards = []
    for label, value in values:
        card = Table(
            [[Paragraph(label, label_style)], [Paragraph(value, value_style)]],
            colWidths=[28.2 * mm], rowHeights=[10 * mm, 16 * mm],
        )
        card.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), WHITE),
            ("BOX", (0, 0), (-1, -1), 0.55, LINE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        cards.append(card)
    grid = Table([cards], colWidths=[29.2 * mm] * 6, hAlign="CENTER")
    grid.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0.5),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return grid


def _signed_money(value: Decimal, places: int = 2) -> str:
    number = _decimal(value)
    if number < 0:
        return f"- {_money_br(abs(number), places)}"
    return f"+ {_money_br(number, places)}"


def _parameters_table(identity: ReportIdentity, settings: ContractSettings, result: EvolutionResult, styles) -> Table:
    # Os dados que já aparecem nos cards não são repetidos nesta tabela.
    values = [
        ("Data da solicitação", _date_br(identity.request_date)),
        ("Valor da 1ª prestação", _money_br(result.installments[0].installment_amount)),
        ("Data do crédito", _date_br(settings.credit_date)),
        ("Valor contratado", _money_br(settings.initial_balance, settings.rounding.money_places)),
        ("Valor líquido", _money_br(identity.net_amount)),
        ("Margem consignável", _money_br(identity.margin_amount)),
        ("FGQC - concessão (operação)", _money_br(identity.fgqc_concession)),
        ("IOF", _money_br(identity.iof_amount)),
        ("Taxa administrativa", _money_br(identity.administrative_fee)),
        ("Saldo quitado na operação", _money_br(identity.settled_loan_balance)),
    ]
    data = []
    for idx in range(0, len(values), 2):
        left = values[idx]
        right = values[idx + 1] if idx + 1 < len(values) else ("", "")
        data.append([
            Paragraph(left[0], styles["label"]), Paragraph(left[1], styles["body_compact"]),
            Paragraph(right[0], styles["label"]), Paragraph(right[1], styles["body_compact"]),
        ])
    table = Table(data, colWidths=[42 * mm, 46 * mm, 42 * mm, 46 * mm])
    commands = [
        ("GRID", (0, 0), (-1, -1), 0.42, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (0, -1), SOFT_BLUE),
        ("BACKGROUND", (2, 0), (2, -1), SOFT_BLUE),
        ("BACKGROUND", (1, 0), (1, -1), WHITE),
        ("BACKGROUND", (3, 0), (3, -1), WHITE),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    table.setStyle(TableStyle(commands))
    return table

def _fraction(numerator: str, denominator: str, styles, width: float = 58 * mm) -> Table:
    table = Table([
        [Paragraph(numerator, styles["formula_small"])],
        [Paragraph(denominator, styles["formula_small"])],
    ], colWidths=[width])
    table.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (0, 0), 0.8, PETROLEUM_DARK),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return table



def _equation_shell(content, width=CONTENT_WIDTH, *, accent=PETROLEUM, background=SOFT_BLUE) -> Table:
    table = Table([[content]], colWidths=[width])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), background),
        ("BOX", (0, 0), (-1, -1), 0.55, LINE),
        ("LINEBEFORE", (0, 0), (0, -1), 3.3, accent),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return table

def _price_equation(symbolic: bool, settings: ContractSettings, result: EvolutionResult, first: InstallmentRow, styles) -> Table:
    """Equação Price dimensionada para caber dentro do card sem ultrapassar margens."""
    if symbolic:
        numerator = "VF × i"
        denominator = "1 - (1 + i)<super>-n</super>"
        timing_num, timing_den = "1", "1 + i"
        result_text = ""
    else:
        pct = _percent_br(result.monthly_interest_rate, settings.rounding.percentage_display_places)
        numerator = f"{_money_br(first.balance_before_installment)} × {pct}"
        denominator = f"1 - (1 + {pct})<super>-{settings.term}</super>"
        timing_num, timing_den = "1", f"1 + {pct}"
        result_text = f"= {_money_br(first.installment_amount)}"

    cells = [
        Paragraph("PMT =", styles["formula_big"]),
        _fraction(numerator, denominator, styles, 66 * mm),
    ]
    widths = [18 * mm, 70 * mm]
    if settings.payment_timing == 1:
        cells.extend([
            Paragraph("×", styles["formula_big"]),
            _fraction(timing_num, timing_den, styles, 28 * mm),
        ])
        widths.extend([7 * mm, 32 * mm])
    if result_text:
        cells.append(Paragraph(result_text, styles["formula_big"]))
        widths.append(30 * mm)

    eq = Table([cells], colWidths=widths, hAlign="CENTER")
    eq.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    # O card externo já fornece fundo, borda e recuos.
    return eq


def _vf_formula(settings: ContractSettings, result: EvolutionResult, first: InstallmentRow, styles) -> list:
    first_daily = [row for row in result.daily_rows if row.installment_number == 1]
    segments: OrderedDict[str, list] = OrderedDict()
    for row in first_daily:
        segments.setdefault(row.competence, []).append(row)
    symbolic = Paragraph("VF = Valor principal × (1 + i)<super>d/DC</super>", styles["formula_big"])
    factors = []
    for group in segments.values():
        days = len(group)
        divisor = calendar_days_in_day_competence(group[0].day, settings.competence_start_day, settings.due_day)
        pct = _percent_br(result.monthly_interest_rate, settings.rounding.percentage_display_places)
        factors.append(f"(1 + {pct})<super>{days}/{divisor}</super>")
    numeric = f"VF = {_money_br(settings.initial_balance)}"
    if factors:
        numeric += " × " + " × ".join(factors)
    numeric += f" = {_money_br(first.balance_before_installment)}"
    return [_equation_shell(symbolic), Spacer(1, 2 * mm), _equation_shell(Paragraph(numeric, styles["formula_big"]))]


def _sac_formula_block(settings: ContractSettings, result: EvolutionResult, first: InstallmentRow, styles) -> list:
    symbolic_rows = Table([
        [Paragraph("A =", styles["formula_big"]), _fraction("PV", "n", styles, 34 * mm)],
        [Paragraph("J = SD × i", styles["formula_big"]), ""],
        [Paragraph("P = A + J", styles["formula_big"]), ""],
    ], colWidths=[60 * mm, 45 * mm], hAlign="CENTER")
    symbolic_rows.setStyle(TableStyle([
        ("SPAN", (0, 1), (1, 1)),
        ("SPAN", (0, 2), (1, 2)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    numeric = Paragraph(
        f"A<sub>1</sub> = ({_money_br(first.balance_before_installment)} - {_money_br(first.interest_amount)}) "
        f"÷ {settings.term} = {_money_br(first.regular_amortization)}<br/>"
        f"P<sub>1</sub> = {_money_br(first.regular_amortization)} + {_money_br(first.interest_amount)} "
        f"= {_money_br(first.installment_amount)}",
        styles["formula_big"],
    )
    return [_equation_shell(symbolic_rows), Spacer(1, 2 * mm), _equation_shell(numeric)]



def _variables_legend(definitions: Sequence[tuple[str, str]], styles) -> Table:
    """Renderiza uma legenda compacta e legível para todas as variáveis da fórmula."""
    rows = [[
        Paragraph("VARIÁVEL", styles["formula_title"]),
        Paragraph("SIGNIFICADO", styles["formula_title"]),
    ]]
    for symbol, meaning in definitions:
        rows.append([
            Paragraph(f"<b>{symbol}</b>", styles["formula_note"]),
            Paragraph(meaning, styles["formula_note"]),
        ])
    table = Table(rows, colWidths=[27 * mm, FORMULA_LEGEND_WIDTH - (27 * mm)], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#E7F1F3")),
        ("TEXTCOLOR", (0, 0), (-1, 0), PETROLEUM_DARK),
        ("GRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return table


def _formula_card(
    title: str,
    equation,
    explanation: str,
    styles,
    *,
    definitions: Sequence[tuple[str, str]] = (),
    accent=PETROLEUM,
) -> Table:
    if isinstance(equation, str):
        equation = Paragraph(equation, styles["formula_big"])
    rows = [
        [Paragraph(title.upper(), styles["formula_title"])],
        [equation],
    ]
    if definitions:
        rows.extend([
            [Paragraph("Onde:", styles["formula_title"])],
            [_variables_legend(definitions, styles)],
        ])
    rows.append([Paragraph(explanation, styles["formula_note"])])
    body = Table(rows, colWidths=[FORMULA_INNER_WIDTH])
    body.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return _equation_shell(body, accent=accent, background=HexColor("#F4F8F9"))


def _step_card(number: int, title: str, text: str, styles, width: float = 85 * mm) -> Table:
    badge_style = ParagraphStyle(
        f"step_badge_{number}", parent=styles["small"], fontName="Helvetica-Bold",
        fontSize=5.8, leading=6.8, textColor=WHITE, alignment=TA_CENTER,
    )
    badge = Table(
        [[Paragraph(f"PASSO {number:02d}", badge_style)]],
        colWidths=[19 * mm], rowHeights=[6.5 * mm],
    )
    badge.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PETROLEUM),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
    ]))
    header = Table(
        [[badge, Paragraph(title, styles["step_title"])]],
        colWidths=[21 * mm, width - 31 * mm],
    )
    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    body = Paragraph(text, styles["step_body"])
    card = Table([[header], [body]], colWidths=[width - 10 * mm], rowHeights=[8 * mm, 18 * mm])
    card.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), WHITE),
        ("BOX", (0, 0), (-1, -1), 0.55, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return card

def _steps_grid(steps: Sequence[tuple[str, str]], styles) -> Table:
    cards = [_step_card(i + 1, title, body, styles) for i, (title, body) in enumerate(steps)]
    rows = []
    for i in range(0, len(cards), 2):
        rows.append([cards[i], cards[i + 1] if i + 1 < len(cards) else ""])
    grid = Table(rows, colWidths=[87 * mm, 87 * mm])
    grid.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return grid


def _daily_sequence(settings: ContractSettings, styles) -> Table:
    labels = [
        ("SALDO INICIAL", PETROLEUM),
        ("JUROS DO DIA", TEAL),
    ]
    if settings.index_code:
        labels.append(("CORREÇÃO EM DIA ÚTIL", FUNCEF_BLUE))
    labels.extend([
        ("EVENTO / AMORTIZAÇÃO", ORANGE),
        ("SALDO FINAL", SUCCESS),
    ])
    cells = []
    widths = []
    for idx, (label, color) in enumerate(labels):
        box = Table([[Paragraph(label, ParagraphStyle(f"seq_{idx}", fontName="Helvetica-Bold", fontSize=6.3, leading=7.5, textColor=WHITE, alignment=TA_CENTER))]], colWidths=[29 * mm], rowHeights=[11 * mm])
        box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), color), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        cells.append(box)
        widths.append(29 * mm)
        if idx < len(labels) - 1:
            cells.append(Paragraph("→", ParagraphStyle(f"arr_{idx}", fontName="Helvetica-Bold", fontSize=10, textColor=MUTED, alignment=TA_CENTER)))
            widths.append(7 * mm)
    table = Table([cells], colWidths=widths, hAlign="CENTER")
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    return table


def _segment_summary(settings: ContractSettings, result: EvolutionResult, styles) -> Table:
    first_rows = [r for r in result.daily_rows if r.installment_number == 1]
    grouped: OrderedDict[str, list] = OrderedDict()
    for row in first_rows:
        grouped.setdefault(row.competence, []).append(row)
    headers = ["Competência", "Dias processados", "Taxa diária de juros", "Índice de referência", "Taxa diária de correção"]
    data = [[Paragraph(h, styles["table_head"]) for h in headers]]
    for competence, rows in grouped.items():
        interest_rate = rows[0].daily_interest_rate if rows else Decimal("0")
        corr_rates = [r.daily_correction_rate for r in rows if r.daily_correction_rate != 0]
        corr_rate = corr_rates[0] if corr_rates else Decimal("0")
        refs = list(dict.fromkeys(r.index_reference for r in rows if r.index_reference))
        ref = " a ".join(refs) if refs else "Não se aplica"
        data.append([
            Paragraph(competence, styles["table"]),
            Paragraph(str(len(rows)), styles["table"]),
            Paragraph(_percent_br(interest_rate, settings.rounding.percentage_display_places), styles["table_right"]),
            Paragraph(ref, styles["table"]),
            Paragraph(_percent_br(corr_rate, settings.rounding.percentage_display_places), styles["table_right"]),
        ])
    table = Table(data, colWidths=[29 * mm, 25 * mm, 39 * mm, 43 * mm, 40 * mm], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), PETROLEUM_DARK),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for idx in range(1, len(data)):
        if idx % 2 == 0:
            style.append(("BACKGROUND", (0, idx), (-1, idx), SOFT))
    table.setStyle(TableStyle(style))
    return table


def _technical_foundation(settings: ContractSettings, styles) -> list:
    code = settings.modality_code
    if code == "MOD_001":
        return [
            Paragraph("Natureza da metodologia de amortização", styles["subsection"]),
            Paragraph(
                "A modalidade variável utiliza uma lógica de amortização vinculada ao saldo devedor atualizado e ao prazo remanescente. "
                "A prestação é formada por juros e amortização, mas a amortização não deve ser descrita como uma quantia numericamente imutável "
                "em todas as parcelas: ela é recalculada em cada ciclo a partir do saldo atualizado, deduzidos os juros do período, dividido pela "
                "quantidade de prestações remanescentes.",
                styles["body_justified"],
            ),
            Paragraph(
                "Os juros remuneram o capital utilizado no período e não reduzem o principal. A redução do saldo decorre da amortização. "
                "A correção monetária integra a evolução diária do saldo, segundo o índice e a defasagem definidos, e pode alterar a base sobre "
                "a qual a amortização do ciclo será apurada.",
                styles["body_justified"],
            ),
            _callout(
                "Ausência de incorporação de juros vencidos ao principal",
                "Na memória examinada, os juros são apropriados diariamente sobre o saldo vigente e liquidados como componente da prestação. "
                "Não se verifica a transferência de juros vencidos para o principal com o objetivo de formar uma nova base de incidência. "
                "Para a correta leitura da evolução, devem ser distinguidos a atualização monetária do saldo, os juros remuneratórios do período "
                "e a amortização do principal. Somente a amortização reduz o saldo devedor principal.",
                styles,
                accent=PETROLEUM,
                background=HexColor("#F7FAFA"),
            ),
        ]
    if code == "MOD_002":
        return [
            Paragraph(
                "O Credplan Fixo utiliza a <b>Tabela Price</b>, com prestação fixa conforme os parâmetros contratuais. Os juros são apropriados "
                "diariamente sobre o saldo vigente e a amortização corresponde à diferença entre a prestação e os juros acumulados no período. "
                "A modalidade não utiliza correção monetária.", styles["body_justified"]),
            Paragraph(
                "Em cada ciclo, a parcela de amortização reduz o principal. A parcela de juros remunera o capital utilizado durante o período e "
                "é demonstrada separadamente na composição da prestação.", styles["body_justified"]),
        ]
    if code == "MOD_004":
        return [
            Paragraph(
                "O Novo Credinâmico Fixo utiliza a <b>Tabela Price</b>, sem correção monetária. O saldo é evoluído diariamente pela taxa de juros "
                "equivalente da competência e, no vencimento, a prestação é recalculada pela fórmula Price tipo 1 sobre o saldo e o prazo remanescentes.",
                styles["body_justified"]),
            Paragraph(
                "A amortização corresponde à prestação apurada menos os juros acumulados no período, reduzindo o principal no fechamento do ciclo.",
                styles["body_justified"]),
        ]
    if code == "MOD_005":
        return [
            Paragraph(
                "O Credinâmico Fixo utiliza a <b>Tabela Price</b>, sem correção monetária e sem apropriação diária de juros. O encargo contratual "
                "é incorporado ao saldo na data de aniversário anual prevista pela metodologia e, em cada vencimento mensal, a prestação é recalculada "
                "sobre o saldo então existente e o prazo remanescente.", styles["body_justified"]),
            Paragraph(
                "Na memória da modalidade, a prestação mensal reduz o saldo já atualizado pelo eventual encargo anual do período.",
                styles["body_justified"]),
        ]
    if code == "MOD_006":
        return [
            Paragraph(
                "O Novo Credinâmico Variável utiliza a <b>Tabela Price</b> com recálculo por ciclo. O saldo é evoluído diariamente por juros equivalentes "
                "e por correção monetária pelo INPC nos dias úteis, observada a defasagem parametrizada; no vencimento, a prestação é recalculada sobre "
                "o saldo atualizado e o prazo remanescente.", styles["body_justified"]),
            Paragraph(
                "A amortização corresponde à parcela da prestação destinada à redução do principal, depois da apropriação dos encargos do ciclo.",
                styles["body_justified"]),
        ]
    return [Paragraph("Metodologia não identificada.", styles["body_justified"])]

def _method_steps(settings: ContractSettings, result: EvolutionResult, styles) -> list:
    code = settings.modality_code
    if code == "MOD_001":
        steps = [
            ("Definir a competência", "A evolução começa no dia seguinte ao crédito e fecha na referência da prestação. A competência operacional segue o ciclo 21-20."),
            ("Converter a taxa", "A taxa anual é convertida em mensal equivalente e depois em taxa diária equivalente conforme os dias corridos da competência."),
            ("Apropriar juros diários", "Os juros remuneratórios são calculados diariamente sobre o saldo de abertura de cada dia."),
            ("Aplicar a correção monetária", "O INPC, com a defasagem parametrizada, é diarizado e aplicado nos dias úteis após os juros do dia."),
            ("Processar eventos", "Eventuais amortizações extraordinárias são processadas após os encargos do dia e reduzem o saldo da etapa."),
            ("Apurar saldo e base", "No vencimento, apura-se o saldo antes da prestação e segregam-se os juros acumulados para formar a base principal de amortização."),
            ("Calcular amortização e prestação", "A base é dividida pelo número de prestações remanescentes. A prestação é formada pela amortização e pelos juros do período."),
            ("Fechar o saldo principal", "O saldo final do principal corresponde à base de amortização menos a amortização do ciclo e inicia a etapa seguinte."),
        ]
    elif code == "MOD_002":
        timing = "início do período (tipo 1)" if settings.payment_timing == 1 else "fim do período (tipo 0)"
        steps = [
            ("Definir a carência", "A data do crédito é excluída e a primeira prestação ocorre no dia 20 do mês seguinte."),
            ("Converter a taxa", "A taxa anual é convertida em mensal equivalente e depois em diária equivalente."),
            ("Evoluir o saldo", "Os juros são apropriados diariamente, sem correção monetária."),
            ("Calcular a prestação", f"A prestação fixa é apurada pela Tabela Price com pagamento no {timing}."),
            ("Apurar os juros", "Os juros do ciclo correspondem ao somatório das apropriações diárias."),
            ("Apurar a amortização", "A amortização é a diferença entre a prestação fixa e os juros do período."),
            ("Reduzir o principal", "A prestação é abatida do saldo após a evolução diária."),
            ("Encerrar o fluxo", "Na última prestação, eventual saldo residual é absorvido no fechamento."),
        ]
    elif code == "MOD_004":
        steps = [
            ("Definir o período", "A evolução diária inicia no dia seguinte ao crédito e fecha no dia 20 de cada mês."),
            ("Converter a taxa", "A taxa anual é transformada em mensal equivalente e depois em taxa diária da competência."),
            ("Apropriar os juros", "Os juros são aplicados diariamente sobre o saldo de abertura; não há correção monetária."),
            ("Processar eventos", "Eventuais amortizações extraordinárias são abatidas após os juros do dia."),
            ("Apurar o saldo", "No vencimento, identifica-se o saldo antes da prestação e os juros acumulados do ciclo."),
            ("Recalcular a Price", "A prestação é apurada com tipo 1 sobre o saldo e o prazo remanescentes."),
            ("Separar a amortização", "A amortização corresponde à prestação menos os juros do período."),
            ("Fechar o ciclo", "A prestação é abatida do saldo para formar o valor inicial do ciclo seguinte."),
        ]
    elif code == "MOD_005":
        steps = [
            ("Definir o calendário", "A primeira prestação ocorre no dia 20 do mês seguinte ao crédito."),
            ("Converter a taxa", "A taxa anual é convertida em taxa mensal equivalente para a fórmula Price."),
            ("Identificar o aniversário", "A data anual de referência é o aniversário da primeira prestação."),
            ("Aplicar o encargo anual", "Somente no aniversário anual a taxa anual contratual é incorporada ao saldo; nos demais dias não há juros."),
            ("Processar eventos", "Eventuais amortizações extraordinárias reduzem o saldo após o encargo do dia, quando houver."),
            ("Recalcular a Price", "Em cada vencimento mensal, a prestação é recalculada com tipo 0 sobre o saldo e o prazo remanescentes."),
            ("Amortizar o saldo", "A prestação reduz integralmente o saldo já atualizado pelo eventual encargo anual."),
            ("Fechar o ciclo", "O saldo remanescente passa a ser a base do mês seguinte."),
        ]
    elif code == "MOD_006":
        steps = [
            ("Definir o período", "A evolução começa no dia seguinte ao crédito e fecha no dia 20 de cada mês."),
            ("Converter a taxa", "A taxa anual é transformada em mensal equivalente e depois em diária equivalente da competência."),
            ("Aplicar os juros", "Os juros são apropriados todos os dias sobre o saldo vigente."),
            ("Aplicar o INPC", "Nos dias úteis, a correção monetária é aplicada após os juros, observada a defasagem de dois meses."),
            ("Processar eventos", "Eventuais amortizações extraordinárias são abatidas depois dos encargos do dia."),
            ("Recalcular a Price", "No vencimento, a prestação é apurada pela Price tipo 1 sobre o saldo atualizado e o prazo remanescente."),
            ("Separar a amortização", "A amortização é a prestação recalculada menos os juros acumulados no ciclo."),
            ("Fechar o saldo", "A prestação é abatida do saldo atualizado, formando a base do ciclo seguinte."),
        ]
    else:
        steps = [("Metodologia", "Modalidade sem roteiro cadastrado.")]
    # Os cards de Passo são deliberadamente distintos dos selos numerados das seções.
    return [_steps_grid(steps, styles)]

def _compact_formula_box(title: str, formula: str, note: str, styles, *, width: float, accent=PETROLEUM) -> Table:
    title_style = ParagraphStyle(
        f"compact_formula_title_{title}", parent=styles["formula_title"], fontName="Helvetica-Bold",
        fontSize=6.8, leading=8.1, textColor=INK,
    )
    formula_style = ParagraphStyle(
        f"compact_formula_{title}", parent=styles["formula"], fontName="Helvetica",
        fontSize=9.2, leading=12.2, textColor=PETROLEUM_DARK, alignment=TA_CENTER,
    )
    note_style = ParagraphStyle(
        f"compact_formula_note_{title}", parent=styles["formula_note"], fontName="Helvetica",
        fontSize=6.4, leading=8.0, textColor=MUTED,
    )
    body = Table([
        [Paragraph(title.upper(), title_style)],
        [Paragraph(formula, formula_style)],
        [Paragraph(note, note_style)],
    ], colWidths=[width - 8 * mm])
    body.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    shell = Table([[body]], colWidths=[width])
    shell.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor("#F4F8F9")),
        ("BOX", (0, 0), (-1, -1), 0.5, LINE),
        ("LINEBEFORE", (0, 0), (0, -1), 2.2, accent),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return shell


def _variable_formula_overview(settings: ContractSettings, result: EvolutionResult, styles) -> list:
    first = result.installments[0]
    third = 57.2 * mm
    pct_a = _percent_br(settings.annual_interest_rate, settings.rounding.percentage_display_places)
    pct_m = _percent_br(result.monthly_interest_rate, settings.rounding.percentage_display_places)

    f1 = _compact_formula_box(
        "Conversão da taxa anual em mensal equivalente",
        "i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1",
        f"Aplicação: i_m = (1 + {pct_a})^(1/12) - 1 = {pct_m}.",
        styles, width=third, accent=ORANGE,
    )
    f2 = _compact_formula_box(
        "Taxa diária de juros",
        "i<sub>d</sub> = (1 + i<sub>m</sub>)<super>1/DC</super> - 1",
        "DC é a quantidade de dias corridos da competência; a taxa diária é recalculada a cada competência.",
        styles, width=third, accent=PETROLEUM,
    )
    f3 = _compact_formula_box(
        "Correção monetária diária",
        "c<sub>d</sub> = (1 + c<sub>m</sub>)<super>1/DU</super> - 1",
        f"O {settings.index_code or 'índice selecionado'} é diarizado pelos dias úteis e observa defasagem de {settings.index_lag_months} meses.",
        styles, width=third, accent=FUNCEF_BLUE,
    )
    f4 = _compact_formula_box(
        "Juros do dia e juros do período",
        "J<sub>d</sub> = SD<sub>0</sub> × i<sub>d</sub><br/>J<sub>período</sub> = Σ J<sub>d</sub>",
        f"Na primeira prestação, o somatório dos juros diários resultou em {_money_br(first.interest_amount)}.",
        styles, width=third, accent=TEAL,
    )
    f5 = _compact_formula_box(
        "Evolução diária do saldo",
        "SD<sub>J</sub> = SD<sub>0</sub> × (1 + i<sub>d</sub>)<br/>SD<sub>J+CM</sub> = SD<sub>J</sub> × (1 + c<sub>d</sub>)",
        "A primeira equação é aplicada todos os dias; a segunda, apenas em dia útil, após os juros do dia.",
        styles, width=third, accent=PETROLEUM,
    )
    f6 = _compact_formula_box(
        "Amortização, prestação e saldo final",
        "Base<sub>A</sub> = SD<sub>antes</sub> - J<sub>período</sub><br/>A = Base<sub>A</sub> / n<br/>P = A + J<sub>período</sub><br/>SD<sub>final</sub> = Base<sub>A</sub> - A",
        f"1ª prestação: base = {_money_br(first.amortization_base)}; amortização = {_money_br(first.regular_amortization)}; prestação = {_money_br(first.installment_amount)}; saldo final = {_money_br(first.closing_balance)}.",
        styles, width=third, accent=ORANGE,
    )

    grid = Table([[f1, f2, f3], [f4, f5, f6]], colWidths=[58 * mm, 58 * mm, 58 * mm])
    grid.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
    ]))
    legend = _callout(
        "Legenda das fórmulas",
        "i_a = taxa anual contratual; i_m = taxa mensal equivalente; i_d = taxa diária de juros; c_m = INPC mensal; "
        "c_d = taxa diária de correção; DC = dias corridos; DU = dias úteis; SD = saldo devedor; J_d = juros do dia; "
        "J_período = somatório dos juros diários; n = prestações remanescentes; A = amortização; P = prestação.",
        styles,
        accent=PETROLEUM,
        background=WHITE,
    )
    carencia_note = _callout(
        "Competências e taxas efetivamente utilizadas",
        "A tabela abaixo apresenta todas as competências efetivamente processadas no ciclo da primeira prestação. "
        "Quando a carência atravessar dois ou mais ciclos, cada competência será exibida individualmente com seus dias processados, "
        "taxa diária de juros, índice de referência e taxa diária de correção.",
        styles,
        accent=FUNCEF_BLUE,
        background=SOFT_BLUE,
    )
    return [
        Paragraph(
            "As equações abaixo reproduzem a lógica utilizada pelo Motor de Cálculos. As fórmulas gerais são apresentadas em conjunto com "
            "a aplicação operacional e, quando pertinente, com valores da primeira prestação para permitir a conferência da memória.",
            styles["body_justified"],
        ),
        Spacer(1, 1.5 * mm),
        grid,
        Spacer(1, 2 * mm),
        legend,
        Spacer(1, 2 * mm),
        Paragraph("Competências e taxas efetivamente utilizadas na primeira prestação", styles["subsection"]),
        _segment_summary(settings, result, styles),
        Spacer(1, 1.5 * mm),
        carencia_note,
    ]


def _formula_application(settings: ContractSettings, result: EvolutionResult, styles) -> list:
    if settings.modality_code == "MOD_001":
        return _variable_formula_overview(settings, result, styles)

    first = result.installments[0]
    pct_a = _percent_br(settings.annual_interest_rate, settings.rounding.percentage_display_places)
    pct_m = _percent_br(result.monthly_interest_rate, settings.rounding.percentage_display_places)

    if settings.modality_code in {"MOD_004", "MOD_005", "MOD_006"}:
        flow = [
            _formula_card(
                "Conversão da taxa anual em mensal equivalente",
                "i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1",
                "Conversão financeira da taxa anual contratual para a taxa mensal equivalente utilizada na evolução.",
                styles,
                definitions=(("i<sub>m</sub>", "taxa mensal equivalente"), ("i<sub>a</sub>", "taxa anual contratual")),
                accent=ORANGE,
            )
        ]
        if settings.modality_code == "MOD_004":
            flow.extend([
                Spacer(1, 2 * mm),
                _formula_card(
                    "Taxa diária de juros",
                    "i<sub>d</sub> = (1 + i<sub>m</sub>)<super>1/DC</super> - 1",
                    "A taxa diária é recalculada conforme a quantidade de dias corridos da competência. Não há correção monetária.",
                    styles,
                    definitions=(("i<sub>d</sub>", "taxa diária equivalente"), ("DC", "dias corridos da competência")),
                ),
                Spacer(1, 2 * mm),
                _formula_card(
                    "Evolução diária do saldo",
                    "SD<sub>J</sub> = SD<sub>0</sub> × (1 + i<sub>d</sub>)",
                    "Os juros são apropriados diariamente sobre o saldo de abertura. Eventuais amortizações extraordinárias são processadas após os juros do dia.",
                    styles,
                    definitions=(("SD<sub>0</sub>", "saldo no início do dia"), ("SD<sub>J</sub>", "saldo após os juros")),
                ),
                Spacer(1, 2 * mm),
                _formula_card(
                    "Prestação Price recalculada - tipo 1",
                    _price_equation(True, settings, result, first, styles),
                    f"Na primeira prestação, o valor apurado foi {_money_br(first.installment_amount)}. Nos ciclos seguintes, a fórmula é reaplicada ao saldo e ao prazo remanescentes.",
                    styles,
                    definitions=(("PMT", "prestação apurada no ciclo"), ("VF", "saldo antes da prestação"), ("n", "prestações remanescentes")),
                    accent=PETROLEUM,
                ),
                Spacer(1, 2 * mm),
                _formula_card(
                    "Decomposição da prestação",
                    "A = PMT - J<sub>período</sub>  &nbsp;&nbsp; | &nbsp;&nbsp; SD<sub>final</sub> = SD<sub>antes</sub> - PMT",
                    f"Primeiro ciclo: J = {_money_br(first.interest_amount)}; A = {_money_br(first.regular_amortization)}; PMT = {_money_br(first.installment_amount)}.",
                    styles,
                    definitions=(("A", "amortização regular"), ("J<sub>período</sub>", "juros acumulados do ciclo")),
                    accent=ORANGE,
                ),
            ])
            return flow
        if settings.modality_code == "MOD_005":
            flow.extend([
                Spacer(1, 2 * mm),
                _formula_card(
                    "Encargo no aniversário anual",
                    "SD<sub>aniv</sub> = SD<sub>0</sub> × (1 + i<sub>a</sub>)",
                    "A taxa anual é aplicada somente no aniversário anual da primeira prestação. Fora dessa data, não há apropriação de juros nem correção monetária.",
                    styles,
                    definitions=(("SD<sub>aniv</sub>", "saldo após o encargo anual"), ("i<sub>a</sub>", "taxa anual contratual")),
                ),
                Spacer(1, 2 * mm),
                _formula_card(
                    "Prestação mensal recalculada - Price tipo 0",
                    "PMT = PV × i<sub>m</sub> / [1 - (1 + i<sub>m</sub>)<super>-n</super>]",
                    f"Na primeira prestação, o valor apurado foi {_money_br(first.installment_amount)}. A fórmula é reaplicada mensalmente ao saldo existente e ao prazo remanescente.",
                    styles,
                    definitions=(("PMT", "prestação mensal do ciclo"), ("PV", "saldo antes da prestação"), ("n", "prestações remanescentes")),
                    accent=PETROLEUM,
                ),
                Spacer(1, 2 * mm),
                _formula_card(
                    "Redução do saldo",
                    "A = PMT  &nbsp;&nbsp; | &nbsp;&nbsp; SD<sub>final</sub> = SD<sub>antes</sub> - PMT",
                    "Na memória da modalidade, a prestação mensal é tratada integralmente como amortização do saldo já acrescido de eventual encargo anual.",
                    styles,
                    definitions=(("A", "amortização do ciclo"), ("PMT", "prestação mensal recalculada")),
                    accent=ORANGE,
                ),
            ])
            return flow
        flow.extend([
            Spacer(1, 2 * mm),
            _formula_card(
                "Taxa diária de juros",
                "i<sub>d</sub> = (1 + i<sub>m</sub>)<super>1/DC</super> - 1",
                "Os juros são apropriados em todos os dias corridos da competência.",
                styles,
                definitions=(("i<sub>d</sub>", "taxa diária equivalente"), ("DC", "dias corridos da competência")),
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Correção monetária diária",
                "c<sub>d</sub> = (1 + c<sub>m</sub>)<super>1/DU</super> - 1",
                f"O INPC é aplicado em dias úteis, com defasagem de {settings.index_lag_months} meses, depois dos juros do dia.",
                styles,
                definitions=(("c<sub>d</sub>", "taxa diária de correção"), ("DU", "dias úteis da competência")),
                accent=FUNCEF_BLUE,
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Evolução diária do saldo",
                "SD<sub>J</sub> = SD<sub>0</sub> × (1 + i<sub>d</sub>)  &nbsp;&nbsp; | &nbsp;&nbsp; SD<sub>J+CM</sub> = SD<sub>J</sub> × (1 + c<sub>d</sub>)",
                "A correção é aplicada apenas quando o dia for útil. Eventuais eventos financeiros são processados depois dos encargos.",
                styles,
                definitions=(("SD<sub>J</sub>", "saldo após juros"), ("SD<sub>J+CM</sub>", "saldo após juros e correção")),
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Prestação Price recalculada - tipo 1",
                _price_equation(True, settings, result, first, styles),
                f"Primeira prestação: {_money_br(first.installment_amount)}. A Price é reaplicada a cada ciclo ao saldo atualizado e ao prazo remanescente.",
                styles,
                definitions=(("PMT", "prestação recalculada"), ("VF", "saldo atualizado antes da prestação"), ("n", "prestações remanescentes")),
                accent=PETROLEUM,
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Decomposição e fechamento",
                "A = PMT - J<sub>período</sub>  &nbsp;&nbsp; | &nbsp;&nbsp; SD<sub>final</sub> = SD<sub>antes</sub> - PMT",
                f"Primeiro ciclo: correção = {_money_br(first.correction_amount)}; juros = {_money_br(first.interest_amount)}; amortização = {_money_br(first.regular_amortization)}.",
                styles,
                definitions=(("A", "amortização regular"), ("J<sub>período</sub>", "juros acumulados do ciclo")),
                accent=ORANGE,
            ),
        ])
        return flow

    flow = [
        _formula_card(
            "Conversão da taxa anual em mensal equivalente",
            "i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1",
            "Conversão financeira da taxa anual contratual para a taxa mensal equivalente utilizada na evolução.",
            styles,
            definitions=(
                ("i<sub>m</sub>", "taxa de juros mensal equivalente"),
                ("i<sub>a</sub>", "taxa de juros anual informada no contrato"),
                ("12", "quantidade de meses considerada em um ano"),
            ),
            accent=ORANGE,
        ),
    ]
    if settings.modality_code == "MOD_001":
        flow.extend([
            Spacer(1, 2 * mm),
            _formula_card(
                "Taxa diária de juros",
                "i<sub>d</sub> = (1 + i<sub>m</sub>)<super>1/DC</super> - 1",
                "A conversão é refeita em cada competência, porque DC pode corresponder a 28, 29, 30 ou 31 dias.",
                styles,
                definitions=(
                    ("i<sub>d</sub>", "taxa de juros diária equivalente aplicada ao saldo"),
                    ("i<sub>m</sub>", "taxa de juros mensal equivalente"),
                    ("DC", "quantidade de dias corridos da competência de aplicação"),
                ),
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Correção monetária diária",
                "c<sub>d</sub> = (1 + c<sub>m</sub>)<super>1/DU</super> - 1",
                f"O índice utilizado é {settings.index_code or 'o índice selecionado'}, com defasagem de {settings.index_lag_months} meses. A taxa diária é aplicada somente nos dias úteis.",
                styles,
                definitions=(
                    ("c<sub>d</sub>", "taxa diária equivalente de correção monetária"),
                    ("c<sub>m</sub>", "índice mensal correspondente à competência de referência, observada a defasagem"),
                    ("DU", "quantidade de dias úteis da competência em que o índice será aplicado"),
                ),
                accent=FUNCEF_BLUE,
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Evolução diária do saldo",
                "SD<sub>J</sub> = SD<sub>0</sub> × (1 + i<sub>d</sub>)  &nbsp;&nbsp; | &nbsp;&nbsp; SD<sub>J+CM</sub> = SD<sub>J</sub> × (1 + c<sub>d</sub>)",
                "A primeira equação é aplicada todos os dias. A segunda é aplicada apenas em dia útil. Eventuais amortizações ou eventos financeiros são processados depois dos encargos do dia.",
                styles,
                definitions=(
                    ("SD<sub>0</sub>", "saldo devedor no início do dia"),
                    ("SD<sub>J</sub>", "saldo devedor após a incidência dos juros diários"),
                    ("SD<sub>J+CM</sub>", "saldo após juros e correção monetária diária nos dias úteis"),
                    ("i<sub>d</sub>", "taxa diária equivalente de juros"),
                    ("c<sub>d</sub>", "taxa diária equivalente de correção monetária"),
                ),
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Amortização e prestação do ciclo",
                "Base<sub>A</sub> = SD<sub>antes</sub> - J<sub>período</sub>  &nbsp;&nbsp; | &nbsp;&nbsp; A = Base<sub>A</sub> / n  &nbsp;&nbsp; | &nbsp;&nbsp; P = A + J",
                "No vencimento, os juros acumulados são segregados para formar a base principal; a base é dividida pelo prazo remanescente e a prestação resulta da soma entre amortização e juros.",
                styles,
                definitions=(
                    ("Base<sub>A</sub>", "base utilizada para calcular a amortização, sem os juros acumulados do período"),
                    ("SD<sub>antes</sub>", "saldo devedor imediatamente antes do processamento da prestação"),
                    ("J<sub>período</sub>", "somatório dos juros diários apurados no ciclo"),
                    ("A", "valor da amortização que reduz o saldo principal"),
                    ("n", "quantidade de prestações remanescentes considerada no ciclo"),
                    ("P", "valor da prestação, formado pela amortização e pelos juros"),
                    ("J", "juros acumulados correspondentes ao período da prestação"),
                ),
                accent=ORANGE,
            ),
        ])
    else:
        flow.extend([
            Spacer(1, 2 * mm),
            _formula_card(
                "Atualização da carência até a primeira prestação",
                "VF = PV × ∏ (1 + i<sub>m</sub>)<super>d/DC</super>",
                f"Aplicação ao contrato: PV = {_money_br(settings.initial_balance)}; o produto considera os segmentos de competência atravessados entre {_date_br(settings.credit_date)} e {_date_br(first.reference_due_date)}; VF = {_money_br(first.balance_before_installment)}.",
                styles,
                definitions=(
                    ("VF", "valor futuro, correspondente ao saldo atualizado na data da primeira prestação"),
                    ("PV", "valor principal ou saldo original na data do crédito"),
                    ("∏", "produto dos fatores de atualização dos segmentos de competência atravessados"),
                    ("i<sub>m</sub>", "taxa de juros mensal equivalente"),
                    ("d", "quantidade de dias efetivamente evoluídos no segmento"),
                    ("DC", "quantidade de dias corridos da competência à qual o segmento pertence"),
                ),
            ),
            Spacer(1, 2 * mm),
        ])
        symbolic = _price_equation(True, settings, result, first, styles)
        numeric = _price_equation(False, settings, result, first, styles)
        price_defs = [
            ("PMT", "valor da prestação fixa calculada pela Tabela Price"),
            ("VF", "saldo atualizado na data da primeira prestação, após a evolução da carência"),
            ("i", "taxa de juros mensal equivalente utilizada no cálculo"),
            ("n", "prazo total do contrato, expresso em número de prestações"),
        ]
        if settings.payment_timing == 1:
            price_defs.append(("1 / (1 + i)", "fator de antecipação utilizado quando o pagamento ocorre no início do período - tipo 1"))
        flow.extend([
            _formula_card(
                "Prestação fixa pela Tabela Price",
                symbolic,
                "A equação apresentada reproduz o tipo de pagamento parametrizado no contrato.",
                styles,
                definitions=tuple(price_defs),
                accent=PETROLEUM,
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Substituição numérica da prestação",
                numeric,
                f"Prestação apurada para o contrato: {_money_br(first.installment_amount)}.",
                styles,
                definitions=(
                    ("PMT", "resultado monetário da prestação fixa"),
                    ("VF", f"saldo devedor de {_money_br(first.balance_before_installment)} na primeira prestação"),
                    ("i", f"taxa mensal equivalente de {pct_m}"),
                    ("n", f"prazo contratual de {settings.term} prestações"),
                ),
                accent=ORANGE,
            ),
            Spacer(1, 2 * mm),
            _formula_card(
                "Decomposição da primeira prestação",
                f"J<sub>1</sub> = Σ J<sub>d</sub> = {_money_br(first.interest_amount)}  &nbsp;&nbsp; | &nbsp;&nbsp; A<sub>1</sub> = PMT - J<sub>1</sub> = {_money_br(first.regular_amortization)}",
                f"O saldo principal após a amortização é {_money_br(first.closing_balance)}. O valor dos juros não é abatido do principal.",
                styles,
                definitions=(
                    ("J<sub>1</sub>", "juros totais da primeira prestação"),
                    ("Σ", "somatório dos juros apurados diariamente durante a carência"),
                    ("J<sub>d</sub>", "valor dos juros de cada dia da carência"),
                    ("A<sub>1</sub>", "amortização da primeira prestação, responsável pela redução do principal"),
                    ("PMT", "valor total da prestação fixa"),
                ),
            ),
            Spacer(1, 3 * mm),
            Paragraph("Competências efetivamente utilizadas na carência", styles["subsection"]),
            _segment_summary(settings, result, styles),
        ])
    return flow

def _methodology_argument(settings: ContractSettings, result: EvolutionResult, styles) -> list:
    return _technical_foundation(settings, styles)

def _installments_table(rows: Sequence[InstallmentRow], overrides, places: int, styles) -> Table:
    headers = ["Nº", "Referência", "Saldo antes", "Correção", "Juros", "Amortização", "FGQC", "Prestação", "Saldo final"]
    data = [[Paragraph(h, styles["table_head"]) for h in headers]]
    for row in rows:
        values = [
            str(row.installment_number),
            _date_br(row.reference_due_date),
            _money_br(_value(row, "balance_before_installment", overrides), places),
            _money_br(_value(row, "correction_amount", overrides), places),
            _money_br(_value(row, "interest_amount", overrides), places),
            _money_br(_value(row, "regular_amortization", overrides), places),
            _money_br(_value(row, "fgqc_amount", overrides), places),
            _money_br(_value(row, "installment_amount", overrides), places),
            _money_br(_value(row, "closing_balance", overrides), places),
        ]
        data.append([
            Paragraph(values[0], styles["table"]),
            Paragraph(values[1], styles["table"]),
            *[Paragraph(v, styles["table_right"]) for v in values[2:]],
        ])
    widths = [7 * mm, 20 * mm, 24 * mm, 18 * mm, 18 * mm, 22 * mm, 17 * mm, 20 * mm, 22 * mm]
    table = Table(data, colWidths=widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), PETROLEUM_DARK),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
    ]
    for idx in range(1, len(data)):
        if idx % 2 == 0:
            style.append(("BACKGROUND", (0, idx), (-1, idx), SOFT))
    table.setStyle(TableStyle(style))
    return table



def _detail_flowables(
    rows: Sequence[InstallmentRow],
    overrides: Mapping[int, Mapping[str, Decimal]],
    settings: ContractSettings,
    styles,
):
    flowables = []
    places = settings.rounding.money_places

    for idx, row in enumerate(rows):
        interest = _value(row, "interest_amount", overrides)
        correction = _value(row, "correction_amount", overrides)
        amortization = _value(row, "regular_amortization", overrides)
        installment = _value(row, "installment_amount", overrides)
        fgqc = _value(row, "fgqc_amount", overrides)
        closing = _value(row, "closing_balance", overrides)
        balance_before = _value(row, "balance_before_installment", overrides)
        amortization_base = _decimal(row.amortization_base)
        opening = _decimal(row.opening_balance)
        extraordinary = _decimal(row.extraordinary_amortization)
        n_cycle = row.remaining_installments + 1
        code = settings.modality_code

        title_style = ParagraphStyle(
            f"detail_title_{idx}", parent=styles["body"], fontName="Helvetica-Bold",
            fontSize=funcef.SUBSECTION_SIZE, leading=11.6, textColor=PETROLEUM_DARK, spaceAfter=4,
        )
        calc_desc = ParagraphStyle(
            f"calc_desc_{idx}", parent=styles["table"], fontName="Helvetica", fontSize=funcef.TABLE_SIZE,
            leading=funcef.TABLE_LEADING, textColor=INK,
        )
        calc_result = ParagraphStyle(
            f"calc_result_{idx}", parent=styles["table_right"], fontName="Helvetica-Bold", fontSize=funcef.TABLE_SIZE,
            leading=funcef.TABLE_LEADING, textColor=PETROLEUM_DARK, alignment=TA_RIGHT,
        )

        title = Paragraph(f"Prestação {row.installment_number} - referência {_date_br(row.reference_due_date)}", title_style)

        if code == "MOD_001":
            extras = f" {_signed_money(-extraordinary, places)}" if extraordinary != 0 else ""
            balance_expr = (
                f"{_money_br(opening, places)} + {_money_br(interest, places)} {_signed_money(correction, places)}{extras}"
            )
            narrative = (
                f"O saldo inicial desta etapa é <b>{_money_br(opening, places)}</b>. Os juros do período, de <b>{_money_br(interest, places)}</b>, "
                f"correspondem ao somatório das apropriações diárias; a correção monetária, de <b>{_money_br(correction, places)}</b>, "
                f"corresponde ao somatório das correções aplicadas nos dias úteis da competência. "
                + (f"Também houve amortização extraordinária de <b>{_money_br(extraordinary, places)}</b>. " if extraordinary != 0 else "")
                + f"Esses movimentos formam o saldo antes da prestação: <b>{balance_expr} = {_money_br(balance_before, places)}</b>. "
                f"A tabela abaixo demonstra, em sequência, a segregação dos juros, a base de amortização, a amortização, a prestação e o saldo principal remanescente."
            )
            rows_calc = [
                ("Saldo antes da prestação", "Saldo inicial + Juros + Correção" + (" - Amortização extraordinária" if extraordinary != 0 else ""), balance_expr, balance_before),
                ("Base de amortização", "Saldo antes - Juros", f"{_money_br(balance_before, places)} - {_money_br(interest, places)}", amortization_base),
                ("Amortização do período", f"Base de amortização / {n_cycle}", f"{_money_br(amortization_base, places)} / {n_cycle}", amortization),
                ("Prestação total", "Amortização + Juros", f"{_money_br(amortization, places)} + {_money_br(interest, places)}", installment),
                ("Saldo final", "Base de amortização - Amortização", f"{_money_br(amortization_base, places)} - {_money_br(amortization, places)}", closing),
            ]
        elif code == "MOD_005":
            narrative = (
                f"O saldo antes da prestação foi de <b>{_money_br(balance_before, places)}</b>. "
                f"O encargo anual do ciclo, quando existente, está refletido nos juros de <b>{_money_br(interest, places)}</b>. "
                f"A prestação Price tipo 0 foi de <b>{_money_br(installment, places)}</b> e reduziu o saldo conforme a metodologia da modalidade, "
                f"encerrando a etapa em <b>{_money_br(closing, places)}</b>."
            )
            rows_calc = [
                ("Saldo antes da prestação", "Saldo disponível no vencimento", _money_br(balance_before, places), balance_before),
                ("Prestação recalculada", "Tabela Price tipo 0", _money_br(installment, places), installment),
                ("Amortização do período", "Amortização = Prestação", _money_br(installment, places), amortization),
                ("Saldo final", "Saldo antes - Prestação", f"{_money_br(balance_before, places)} - {_money_br(installment, places)}", closing),
            ]
        else:
            correction_text = f" e correção de <b>{_money_br(correction, places)}</b>" if correction != 0 else ""
            narrative = (
                f"O saldo antes da prestação foi de <b>{_money_br(balance_before, places)}</b>, com juros do período de "
                f"<b>{_money_br(interest, places)}</b>{correction_text}. A prestação totalizou <b>{_money_br(installment, places)}</b>, "
                f"sendo <b>{_money_br(amortization, places)}</b> destinados à amortização. O saldo final da etapa foi de "
                f"<b>{_money_br(closing, places)}</b>."
            )
            rows_calc = [
                ("Saldo antes da prestação", "Saldo evoluído no ciclo", _money_br(balance_before, places), balance_before),
                ("Juros do período", "Encargo remuneratório do ciclo", _money_br(interest, places), interest),
                ("Amortização do período", "Prestação - Juros", f"{_money_br(installment, places)} - {_money_br(interest, places)}", amortization),
                ("Prestação total", "Valor da prestação do ciclo", _money_br(installment, places), installment),
                ("Saldo final", "Saldo antes - Prestação", f"{_money_br(balance_before, places)} - {_money_br(installment, places)}", closing),
            ]

        info = _callout("Como os valores desta prestação foram formados", narrative, styles, accent=PETROLEUM, background=HexColor("#F7FAFA"))
        flowables.extend([KeepTogether([title, info]), Spacer(1, 2.5 * mm)])

        data = [[
            Paragraph("Etapa", styles["table_head"]),
            Paragraph("Descrição", styles["table_head"]),
            Paragraph("Substituição com valores", styles["table_head"]),
            Paragraph("Resultado", styles["table_head"]),
        ]]
        for number, (label, rule, substitution, result_value) in enumerate(rows_calc, start=1):
            data.append([
                Paragraph(str(number), ParagraphStyle(f"calc_num_{idx}_{number}", parent=styles["table"], fontName="Helvetica-Bold", alignment=TA_CENTER, textColor=PETROLEUM_DARK)),
                Paragraph(f"<b>{label}</b><br/><font color='#64777B'>{rule}</font>", calc_desc),
                Paragraph(substitution, calc_desc),
                Paragraph(_money_br(result_value, places), calc_result),
            ])
        calc_table = Table(data, colWidths=[12 * mm, 48 * mm, 76 * mm, 40 * mm], repeatRows=1)
        calc_style = [
            ("BACKGROUND", (0, 0), (-1, 0), PETROLEUM_DARK),
            ("GRID", (0, 0), (-1, -1), 0.42, LINE),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]
        for r in range(1, len(data)):
            if r % 2 == 0:
                calc_style.append(("BACKGROUND", (0, r), (-1, r), SOFT))
        calc_table.setStyle(TableStyle(calc_style))
        flowables.append(calc_table)

        if fgqc != 0:
            flowables.extend([
                Spacer(1, 1.5 * mm),
                Paragraph(
                    f"<b>FGQC:</b> {_money_br(fgqc, places)} demonstrado separadamente nesta prestação; o valor não compõe a prestação calculada acima.",
                    styles["small"],
                ),
            ])

        flowables.append(KeepTogether([
            Spacer(1, 2.2 * mm),
            Paragraph(f"Resumo da Prestação {row.installment_number}", styles["subsection"]),
            _installments_table([row], overrides, places, styles),
        ]))

        if idx != len(rows) - 1:
            flowables.append(Spacer(1, 5 * mm))
    return flowables

def _object_text(identity: ReportIdentity, settings: ContractSettings, modality_name: str, styles) -> Paragraph:
    contract = f" nº <b>{identity.contract_number.strip()}</b>" if identity.contract_number.strip() else ""
    return Paragraph(
        f"O presente parecer tem por objeto demonstrar, de forma técnica e padronizada, a evolução do contrato de "
        f"empréstimo <b>{modality_name}</b>{contract}. A análise considera os parâmetros contratuais informados, a metodologia "
        f"financeira aplicável à modalidade e a memória de cálculo produzida pelo Motor de Cálculos, permitindo a reprodução e a "
        f"conferência das etapas que formam as prestações e o saldo devedor.",
        styles["body_justified"],
    )

def _conclusion(settings: ContractSettings, first: InstallmentRow, overrides, styles) -> list:
    interest = _value(first, "interest_amount", overrides)
    amort = _value(first, "regular_amortization", overrides)
    payment = _value(first, "installment_amount", overrides)
    closing = _value(first, "closing_balance", overrides)
    fgqc = _value(first, "fgqc_amount", overrides)
    code = settings.modality_code

    if code == "MOD_001":
        paragraphs = [
            "Sob o ponto de vista estritamente técnico-financeiro e considerando a metodologia demonstrada, não foi identificada a incorporação de juros vencidos ao saldo principal para a geração de novos encargos. Os juros foram apropriados diariamente sobre o saldo vigente e apresentados como componente da prestação.",
            "A correção monetária foi aplicada segundo o índice, a defasagem e os dias úteis da competência de aplicação. A redução do principal ocorreu por meio da amortização recalculada para o saldo atualizado e para o prazo remanescente, razão pela qual seu valor pode variar entre os ciclos.",
        ]
    elif code == "MOD_005":
        paragraphs = [
            "A evolução apresentada foi processada conforme os parâmetros informados para o Credinâmico Fixo, sem correção monetária e com aplicação do encargo contratual na periodicidade anual definida pela metodologia.",
            "As prestações mensais foram apuradas pela Tabela Price sobre o saldo existente e o prazo remanescente, com demonstração segregada dos componentes utilizados no fechamento de cada ciclo.",
        ]
    elif code == "MOD_004":
        paragraphs = [
            "A evolução apresentada foi processada conforme os parâmetros informados para o Novo Credinâmico Fixo, sem correção monetária, com juros diários e recálculo da prestação pela Tabela Price sobre o saldo e o prazo remanescentes.",
            "A memória demonstra a separação entre juros, amortização, prestação e saldo final em cada ciclo de referência.",
        ]
    elif code == "MOD_006":
        paragraphs = [
            "A evolução apresentada foi processada conforme os parâmetros informados para o Novo Credinâmico Variável, com juros diários, correção monetária pelo INPC em dias úteis e recálculo da prestação pela Tabela Price a cada ciclo.",
            "A memória demonstra os encargos, a amortização e o saldo remanescente de forma rastreável em cada prestação de referência.",
        ]
    else:
        paragraphs = [
            "A evolução apresentada foi processada conforme os parâmetros informados para a modalidade, com prestação apurada pela Tabela Price e demonstração segregada dos juros, da amortização e do saldo remanescente.",
            "A memória permite a conferência da formação das prestações e da evolução do principal ao longo do contrato.",
        ]

    result_line = (
        f"Na primeira prestação de referência, a prestação de <b>{_money_br(payment)}</b> compreende juros de "
        f"<b>{_money_br(interest)}</b> e amortização de <b>{_money_br(amort)}</b>, com saldo final de <b>{_money_br(closing)}</b>."
    )
    if fgqc != 0:
        result_line += f" O FGQC de <b>{_money_br(fgqc)}</b> é demonstrado separadamente."
    contents = [Paragraph(p, styles["body_justified"]) for p in paragraphs]
    contents.append(Paragraph(result_line, styles["body_justified"]))
    return contents

def selected_installments(result: EvolutionResult, selected_installment_numbers: Sequence[int]) -> list[InstallmentRow]:
    if not result.installments:
        raise ValueError("Não há prestação concluída para compor o parecer.")
    by_number = {row.installment_number: row for row in result.installments}
    selected_rows = [by_number[number] for number in selected_installment_numbers if number in by_number]
    if not selected_rows:
        raise ValueError("Selecione pelo menos uma prestação concluída para o parecer.")
    return selected_rows


def legacy_body_story(
    *,
    settings: ContractSettings,
    result: EvolutionResult,
    modality_name: str,
    identity: ReportIdentity,
    selected_rows: Sequence[InstallmentRow],
    overrides: Mapping[int, Mapping[str, Decimal]],
) -> list:
    """Conteúdo técnico das modalidades 004 a 006 (sem abertura e sem fechamento)."""
    styles = _styles()
    first = result.installments[0]
    story = []

    story.append(KeepTogether([
        _section_title("04", "Objeto e parâmetros contratuais", None, styles),
        Spacer(1, 2.0 * mm),
        _object_text(identity, settings, modality_name, styles),
    ]))
    story.append(Spacer(1, 2.2 * mm))
    story.append(_summary_cards(settings, result, modality_name, styles))
    story.append(Spacer(1, 3.6 * mm))
    story.append(KeepTogether([_parameters_table(identity, settings, result, styles)]))

    foundation = _technical_foundation(settings, styles)
    story.append(Spacer(1, 3.5 * mm))
    story.append(CondPageBreak(38 * mm))
    story.append(KeepTogether([
        _section_title("05", "Fundamentação técnico-financeira", None, styles),
        Spacer(1, 2.0 * mm),
        foundation[0],
    ]))
    story.extend(foundation[1:])

    # Paginação fluida: a metodologia começa na mesma página quando houver área útil.
    story.append(Spacer(1, 3.5 * mm))
    story.append(CondPageBreak(42 * mm))
    story.append(_section_title("06", "Metodologia aplicada - passo a passo", None, styles))
    story.append(Spacer(1, 2.0 * mm))
    story.extend(_method_steps(settings, result, styles))
    if settings.modality_code == "MOD_001":
        story.append(Spacer(1, 2.2 * mm))
        story.append(KeepTogether([_callout(
            "Nota sobre competência 21-20",
            _competence_2120_example(settings),
            styles,
            accent=FUNCEF_BLUE,
            background=SOFT_BLUE,
        )]))
    if settings.modality_code in {"MOD_001", "MOD_002"}:
        story.append(Spacer(1, 2.6 * mm))
        story.append(CondPageBreak(31 * mm))
        story.append(_daily_sequence(settings, styles))

    # Fórmulas seguem o fluxo; a grade pode continuar na página seguinte por linha.
    story.append(Spacer(1, 3.5 * mm))
    story.append(CondPageBreak(42 * mm))
    story.append(_section_title("07", "Fórmulas e aplicação numérica", None, styles))
    story.append(Spacer(1, 2.0 * mm))
    story.extend(_formula_application(settings, result, styles))

    story.append(Spacer(1, 3.2 * mm))
    story.append(CondPageBreak(36 * mm))
    story.append(KeepTogether([
        _section_title("08", "Demonstração e detalhamento das prestações", None, styles),
        Spacer(1, 2.0 * mm),
        Paragraph("Visão consolidada das prestações de referência", styles["subsection"]),
    ]))
    story.append(_installments_table(selected_rows, overrides, settings.rounding.money_places, styles))


    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph("Detalhamento padronizado das prestações de referência", styles["subsection"]))
    story.extend(_detail_flowables(selected_rows[:2], overrides, settings, styles))

    conclusion_parts = _conclusion(settings, first, overrides, styles)
    story.append(Spacer(1, 3.5 * mm))
    story.append(CondPageBreak(40 * mm))
    story.append(KeepTogether([
        _section_title("09", "Conclusão técnica", None, styles),
        Spacer(1, 2.0 * mm),
        conclusion_parts[0],
    ]))
    story.extend(conclusion_parts[1:])
    if identity.additional_note.strip():
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph("Observação da elaboração", styles["subsection"]))
        story.append(Paragraph(identity.additional_note.strip().replace("\n", "<br/>"), styles["body_justified"]))

    return story


def build_opinion_pdf(
    *,
    settings: ContractSettings,
    result: EvolutionResult,
    modality_name: str,
    identity: ReportIdentity,
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

    output = BytesIO()
    doc = BaseDocTemplate(
        output,
        pagesize=A4,
        leftMargin=PAGE_MARGIN,
        rightMargin=PAGE_MARGIN,
        topMargin=funcef.TOP_MARGIN,
        bottomMargin=funcef.BOTTOM_MARGIN,
        title=f"Parecer técnico - {modality_name}",
        author="FUNCEF - Motor de Cálculos",
        subject="Evolução contratual",
    )
    frame = Frame(
        PAGE_MARGIN, funcef.BOTTOM_MARGIN, CONTENT_WIDTH,
        A4[1] - funcef.TOP_MARGIN - funcef.BOTTOM_MARGIN,
        id="body", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )
    doc.addPageTemplates([PageTemplate(id="all", frames=[frame])])

    story = funcef.opening_story(
        identity.header,
        width=CONTENT_WIDTH,
        logo_path=logo_path,
        contract_number=identity.contract_number,
        modality_name=modality_name,
    )
    story += legacy_body_story(
        settings=settings,
        result=result,
        modality_name=modality_name,
        identity=identity,
        selected_rows=selected_rows,
        overrides=overrides,
    )
    story += funcef.closing_story(identity.header, width=CONTENT_WIDTH)

    doc.build(funcef.prepare_pdf_story(story, width=CONTENT_WIDTH), canvasmaker=funcef.page_canvas(PAGE_MARGIN))
    return output.getvalue()



# === PARECER DESIGN SYSTEM APROVADO v0.10.5 ===
# Camada visual validada para os pareceres. As funções abaixo sobrescrevem
# somente a apresentação; motores de cálculo e estruturas numéricas permanecem intactos.
from reportlab.graphics.shapes import Drawing, Circle, Rect, Line, String, Ellipse, Path as GraphicsPath

PDF_BLUE = HexColor("#00366B")
PDF_BLUE_DARK = HexColor("#002D5B")
PDF_BLUE_LIGHT = HexColor("#EAF3FB")
PDF_BLUE_SOFT = HexColor("#F5F9FD")
PDF_ORANGE = HexColor("#F6791C")
PDF_TEXT = HexColor("#18324A")
PDF_MUTED = HexColor("#667A8D")
PDF_LINE = HexColor("#BDD0E2")
PDF_GREEN = HexColor("#18855B")

# Reaplica a paleta aos helpers legados que continuam sendo utilizados.
PETROLEUM = PDF_BLUE
PETROLEUM_DARK = PDF_BLUE
PETROLEUM_DEEP = PDF_BLUE_DARK
TEAL = HexColor("#087A96")
FUNCEF_BLUE = HexColor("#0964B4")
ORANGE = PDF_ORANGE
INK = PDF_TEXT
MUTED = PDF_MUTED
LINE = PDF_LINE
SOFT = PDF_BLUE_SOFT
SUCCESS = PDF_GREEN


def _approved_icon(kind: str, size: float = 12 * mm, badge: bool = True) -> Drawing:
    d = Drawing(size, size)
    cx = cy = size / 2
    c = PDF_BLUE
    sw = max(0.8, size * 0.034)
    k = (kind or '').lower()
    if badge:
        d.add(Circle(cx, cy, size * 0.43, fillColor=PDF_BLUE_LIGHT, strokeColor=HexColor("#BCD6EE"), strokeWidth=0.65))

    def line(x1, y1, x2, y2, w=sw):
        d.add(Line(x1, y1, x2, y2, strokeColor=c, strokeWidth=w, strokeLineCap=1))

    if k in {'percent', 'interest'}:
        d.add(String(cx, cy - size*.14, '%', fontName='Helvetica-Bold', fontSize=size*.54, fillColor=c, textAnchor='middle'))

    elif k in {'calendar', 'date', 'term'}:
        x, y, w, h = size*.27, size*.28, size*.46, size*.43
        d.add(Rect(x, y, w, h, rx=size*.025, ry=size*.025, fillColor=None, strokeColor=c, strokeWidth=sw))
        line(x, y+h*.72, x+w, y+h*.72)
        line(x+w*.27, y+h*.86, x+w*.27, y+h*1.04)
        line(x+w*.73, y+h*.86, x+w*.73, y+h*1.04)
        if k == 'date':
            line(x+w*.30, y+h*.35, x+w*.43, y+h*.21, sw*1.05)
            line(x+w*.43, y+h*.21, x+w*.72, y+h*.50, sw*1.05)
        elif k == 'term':
            line(x+w*.50, y+h*.29, x+w*.50, y+h*.56)
            line(x+w*.365, y+h*.425, x+w*.635, y+h*.425)
        else:
            for xx in (.31, .50, .69):
                d.add(Circle(x+w*xx, y+h*.40, size*.018, fillColor=c, strokeColor=c))

    elif k in {'index', 'chart'}:
        base_y = size*.30
        for i, hh in enumerate((.17,.28,.41)):
            d.add(Rect(size*(.29+i*.14), base_y, size*.075, size*hh, fillColor=c, strokeColor=c, strokeWidth=0))
        p = GraphicsPath(); p.moveTo(size*.28,size*.57); p.lineTo(size*.42,size*.64); p.lineTo(size*.55,size*.61); p.lineTo(size*.72,size*.76)
        p.strokeColor=c; p.strokeWidth=sw*1.1; p.fillColor=None; d.add(p)
        line(size*.64,size*.75,size*.72,size*.76,sw)
        line(size*.72,size*.76,size*.71,size*.68,sw)

    elif k in {'system', 'gear'}:
        d.add(Circle(cx, cy, size*.215, fillColor=None, strokeColor=c, strokeWidth=sw))
        d.add(Circle(cx, cy, size*.075, fillColor=None, strokeColor=c, strokeWidth=sw))
        for dx,dy in ((0,1),(0,-1),(1,0),(-1,0),(.71,.71),(-.71,.71),(.71,-.71),(-.71,-.71)):
            line(cx+dx*size*.22, cy+dy*size*.22, cx+dx*size*.30, cy+dy*size*.30, sw*1.25)

    elif k in {'modality', 'people'}:
        d.add(Circle(cx, size*.61, size*.075, fillColor=None, strokeColor=c, strokeWidth=sw))
        d.add(Circle(size*.35, size*.57, size*.055, fillColor=None, strokeColor=c, strokeWidth=sw*.9))
        d.add(Circle(size*.65, size*.57, size*.055, fillColor=None, strokeColor=c, strokeWidth=sw*.9))
        p=GraphicsPath(); p.moveTo(size*.34,size*.39); p.curveTo(size*.39,size*.27,size*.61,size*.27,size*.66,size*.39)
        p.strokeColor=c; p.strokeWidth=sw; p.fillColor=None; d.add(p)
        p=GraphicsPath(); p.moveTo(size*.25,size*.42); p.curveTo(size*.28,size*.34,size*.35,size*.32,size*.40,size*.35)
        p.strokeColor=c; p.strokeWidth=sw*.85; p.fillColor=None; d.add(p)
        p=GraphicsPath(); p.moveTo(size*.60,size*.35); p.curveTo(size*.65,size*.32,size*.72,size*.34,size*.75,size*.42)
        p.strokeColor=c; p.strokeWidth=sw*.85; p.fillColor=None; d.add(p)
    elif k in {'money', 'coins'}:
        for xx, yy in ((.40,.39),(.40,.46),(.40,.53),(.59,.36),(.59,.43),(.59,.50)):
            d.add(Ellipse(size*xx, size*yy, size*.10, size*.035, fillColor=None, strokeColor=c, strokeWidth=sw*.9))
        line(size*.30,size*.39,size*.30,size*.53,sw*.8); line(size*.50,size*.39,size*.50,size*.53,sw*.8)
        line(size*.49,size*.36,size*.49,size*.50,sw*.8); line(size*.69,size*.36,size*.69,size*.50,sw*.8)

    elif k == 'banknote':
        x,y,w,h=size*.25,size*.34,size*.50,size*.31
        d.add(Rect(x,y,w,h,rx=size*.025,ry=size*.025,fillColor=None,strokeColor=c,strokeWidth=sw))
        d.add(Circle(cx,cy,size*.075,fillColor=None,strokeColor=c,strokeWidth=sw*.8))
        line(x+size*.05,y+size*.05,x+size*.11,y+size*.05,sw*.7); line(x+w-size*.11,y+h-size*.05,x+w-size*.05,y+h-size*.05,sw*.7)

    elif k in {'document', 'file'}:
        x,y,w,h=size*.32,size*.25,size*.36,size*.50
        d.add(Rect(x,y,w,h,fillColor=None,strokeColor=c,strokeWidth=sw))
        line(x+w*.68,y+h,x+w,y+h*.72,sw*.8); line(x+w*.68,y+h,x+w*.68,y+h*.72,sw*.8); line(x+w*.68,y+h*.72,x+w,y+h*.72,sw*.8)
        for yy in (.58,.47,.36): line(x+w*.18,y+h*yy,x+w*.78,y+h*yy,sw*.75)

    elif k == 'calculator':
        x,y,w,h=size*.31,size*.25,size*.38,size*.51
        d.add(Rect(x,y,w,h,rx=size*.02,ry=size*.02,fillColor=None,strokeColor=c,strokeWidth=sw))
        d.add(Rect(x+w*.14,y+h*.67,w*.72,h*.16,fillColor=None,strokeColor=c,strokeWidth=sw*.8))
        for r in range(2):
            for col in range(3):
                d.add(Rect(x+w*(.16+col*.24),y+h*(.18+r*.22),w*.11,h*.10,fillColor=None,strokeColor=c,strokeWidth=sw*.65))

    elif k == 'shield':
        p=GraphicsPath(); p.moveTo(cx,size*.73); p.lineTo(size*.70,size*.64); p.lineTo(size*.67,size*.39); p.curveTo(size*.62,size*.30,size*.55,size*.25,cx,size*.22); p.curveTo(size*.45,size*.25,size*.38,size*.30,size*.33,size*.39); p.lineTo(size*.30,size*.64); p.closePath()
        p.strokeColor=c; p.strokeWidth=sw; p.fillColor=None; d.add(p)
        line(size*.39,size*.47,size*.47,size*.38,sw); line(size*.47,size*.38,size*.62,size*.54,sw)

    elif k == 'info':
        d.add(String(cx, cy-size*.13, 'i', fontName='Helvetica-Bold', fontSize=size*.48, fillColor=c, textAnchor='middle'))

    elif k == 'block':
        d.add(Circle(cx,cy,size*.22,fillColor=None,strokeColor=c,strokeWidth=sw))
        line(size*.35,size*.35,size*.65,size*.65,sw)

    elif k == 'balance':
        line(cx,size*.30,cx,size*.72,sw)
        line(size*.29,size*.61,size*.71,size*.61,sw)
        line(size*.37,size*.61,size*.30,size*.44,sw*.8); line(size*.63,size*.61,size*.70,size*.44,sw*.8)
        p=GraphicsPath(); p.moveTo(size*.24,size*.44); p.curveTo(size*.29,size*.38,size*.36,size*.38,size*.41,size*.44)
        p.strokeColor=c; p.strokeWidth=sw*.8; p.fillColor=None; d.add(p)
        p=GraphicsPath(); p.moveTo(size*.59,size*.44); p.curveTo(size*.64,size*.38,size*.71,size*.38,size*.76,size*.44)
        p.strokeColor=c; p.strokeWidth=sw*.8; p.fillColor=None; d.add(p)
        line(size*.40,size*.29,size*.60,size*.29,sw)

    elif k == 'search':
        d.add(Circle(size*.46,size*.56,size*.19,fillColor=None,strokeColor=c,strokeWidth=sw))
        line(size*.59,size*.43,size*.74,size*.28,sw*1.1)

    elif k == 'cycle':
        p=GraphicsPath(); p.moveTo(size*.31,size*.61); p.curveTo(size*.33,size*.76,size*.57,size*.79,size*.69,size*.64)
        p.strokeColor=c; p.strokeWidth=sw; p.fillColor=None; d.add(p)
        line(size*.63,size*.67,size*.69,size*.64,sw); line(size*.69,size*.64,size*.68,size*.72,sw)
        p=GraphicsPath(); p.moveTo(size*.69,size*.39); p.curveTo(size*.67,size*.24,size*.43,size*.21,size*.31,size*.36)
        p.strokeColor=c; p.strokeWidth=sw; p.fillColor=None; d.add(p)
        line(size*.37,size*.33,size*.31,size*.36,sw); line(size*.31,size*.36,size*.32,size*.28,sw)

    elif k == 'lock':
        d.add(Rect(size*.34,size*.31,size*.32,size*.30,rx=size*.03,ry=size*.03,fillColor=None,strokeColor=c,strokeWidth=sw))
        p=GraphicsPath(); p.moveTo(size*.40,size*.61); p.curveTo(size*.40,size*.78,size*.60,size*.78,size*.60,size*.61)
        p.strokeColor=c; p.strokeWidth=sw; p.fillColor=None; d.add(p)
        d.add(Circle(cx,size*.45,size*.025,fillColor=c,strokeColor=c))
        line(cx,size*.43,cx,size*.37,sw*.75)

    elif k == 'target':
        d.add(Circle(cx,cy,size*.22,fillColor=None,strokeColor=c,strokeWidth=sw*.8)); d.add(Circle(cx,cy,size*.13,fillColor=None,strokeColor=c,strokeWidth=sw*.8)); d.add(Circle(cx,cy,size*.04,fillColor=c,strokeColor=c))
        line(size*.55,size*.55,size*.73,size*.73,sw); line(size*.68,size*.73,size*.73,size*.73,sw); line(size*.73,size*.68,size*.73,size*.73,sw)

    else:
        d.add(Circle(cx,cy,size*.16,fillColor=None,strokeColor=c,strokeWidth=sw))
    return d


def _styles():
    sample = getSampleStyleSheet()
    body = funcef.BODY_SIZE
    table = funcef.TABLE_SIZE
    note = funcef.NOTE_SIZE
    return {
        'body': ParagraphStyle('body_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=body, leading=funcef.BODY_LEADING, textColor=PDF_TEXT, spaceAfter=5.0),
        'body_justified': ParagraphStyle('bodyj_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=body, leading=funcef.BODY_LEADING, textColor=PDF_TEXT, spaceAfter=5.0, alignment=4),
        'body_compact': ParagraphStyle('bodyc_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=note, leading=10.2, textColor=PDF_TEXT, spaceAfter=2.0),
        'small': ParagraphStyle('small_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=note, leading=10.2, textColor=PDF_MUTED),
        'micro': ParagraphStyle('micro_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=funcef.MIN_SIZE, leading=8.6, textColor=PDF_MUTED),
        'label': ParagraphStyle('label_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=table, leading=funcef.TABLE_LEADING, textColor=PDF_TEXT),
        'card_value': ParagraphStyle('cv_v105', parent=sample['BodyText'], fontName='Helvetica-Bold', fontSize=10.0, leading=11.7, textColor=PDF_BLUE, alignment=TA_CENTER),
        'hero_title': ParagraphStyle('hero_v105', parent=sample['Title'], fontName='Helvetica-Bold', fontSize=17.0, leading=19.0, textColor=PDF_BLUE),
        'hero_subtitle': ParagraphStyle('heros_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=note, leading=10.2, textColor=PDF_MUTED),
        'section': ParagraphStyle('sec_v105', parent=sample['Heading2'], fontName='Helvetica-Bold', fontSize=10.0, leading=11.5, textColor=PDF_BLUE, spaceBefore=0, spaceAfter=0),
        'section_line': ParagraphStyle('secl_v105', parent=sample['BodyText'], fontName='Helvetica-Bold', fontSize=10.0, leading=11.5, textColor=PDF_BLUE, spaceBefore=0, spaceAfter=0),
        'section_subtitle': ParagraphStyle('secs_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=note, leading=10.2, textColor=PDF_MUTED),
        'subsection': ParagraphStyle('sub_v105', parent=sample['Heading3'], fontName='Helvetica-Bold', fontSize=funcef.SUBSECTION_SIZE, leading=11.6, textColor=PDF_BLUE, spaceBefore=4, spaceAfter=3, keepWithNext=1),
        'step_number': ParagraphStyle('sn_v105', parent=sample['BodyText'], fontName='Helvetica-Bold', fontSize=table, leading=8.6, textColor=WHITE, alignment=TA_CENTER),
        'step_title': ParagraphStyle('st_v105', parent=sample['BodyText'], fontName='Helvetica-Bold', fontSize=8.5, leading=10.0, textColor=PDF_BLUE),
        'step_body': ParagraphStyle('sb_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=table, leading=funcef.TABLE_LEADING, textColor=PDF_TEXT),
        'formula': ParagraphStyle('f_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=10.0, leading=13.0, textColor=PDF_BLUE, alignment=TA_CENTER),
        'formula_big': ParagraphStyle('fb_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=10.5, leading=13.5, textColor=PDF_BLUE, alignment=TA_CENTER),
        'formula_small': ParagraphStyle('fs_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=note, leading=10.2, textColor=PDF_BLUE, alignment=TA_CENTER),
        'formula_title': ParagraphStyle('ft_v105', parent=sample['BodyText'], fontName='Helvetica-Bold', fontSize=table, leading=funcef.TABLE_LEADING, textColor=PDF_BLUE),
        'formula_note': ParagraphStyle('fn_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=table, leading=funcef.TABLE_LEADING, textColor=PDF_MUTED),
        'table_head': ParagraphStyle('th_v105', parent=sample['BodyText'], fontName='Helvetica-Bold', fontSize=table, leading=funcef.TABLE_LEADING, textColor=WHITE, alignment=TA_CENTER),
        'table': ParagraphStyle('t_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=table, leading=funcef.TABLE_LEADING, textColor=PDF_TEXT),
        'table_right': ParagraphStyle('tr_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=table, leading=funcef.TABLE_LEADING, textColor=PDF_TEXT, alignment=TA_RIGHT),
        'note': ParagraphStyle('n_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=note, leading=10.2, textColor=PDF_TEXT),
        'callout': ParagraphStyle('co_v105', parent=sample['BodyText'], fontName='Helvetica', fontSize=note, leading=10.6, textColor=PDF_TEXT),
    }


def _section_title(number: str, title: str, subtitle: str | None, styles) -> Table:
    return funcef.section_heading(number, title, CONTENT_WIDTH)


def _summary_cards(settings: ContractSettings, result: EvolutionResult, modality_name: str, styles) -> Table:
    first = result.installments[0]
    values = [
        ('modality','Modalidade', modality_name),
        ('system','Sistema de amortização', _amortization_system_label(settings)),
        ('percent','Taxa de juros (a.a.)', _percent_br(settings.annual_interest_rate, settings.rounding.percentage_display_places)),
        ('term','Prazo', f'{settings.term}\nprestações'),
        ('index','Índice de correção', settings.index_code or 'Sem correção'),
        ('date','Data da 1ª prestação', _date_br(first.reference_due_date)),
    ]
    label_style = ParagraphStyle('kpil_v105', fontName='Helvetica', fontSize=6.7, leading=8.2, textColor=PDF_TEXT, alignment=TA_CENTER)
    val_style = ParagraphStyle('kpiv_v105', fontName='Helvetica-Bold', fontSize=9.4, leading=10.7, textColor=PDF_BLUE, alignment=TA_CENTER)
    cards=[]
    for kind,label,value in values:
        value = value.replace('\n','<br/>')
        card=Table([[_approved_icon(kind, 11*mm)],[Paragraph(label,label_style)],[Paragraph(value,val_style)]], colWidths=[27.5*mm], rowHeights=[13.5*mm,10.5*mm,13.0*mm])
        card.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1),WHITE),('BOX',(0,0),(-1,-1),0.65,PDF_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),1),('BOTTOMPADDING',(0,0),(-1,-1),1),
        ])); cards.append(card)
    grid=Table([cards], colWidths=[29.2*mm]*6, hAlign='CENTER')
    grid.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0.5),('RIGHTPADDING',(0,0),(-1,-1),0.5),('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0)]))
    return grid


def _tiny_icon(kind: str):
    return _approved_icon(kind, 7.2*mm, badge=False)


def _parameters_table(identity: ReportIdentity, settings: ContractSettings, result: EvolutionResult, styles) -> Table:
    values = [
        ('calendar','Data da solicitação', _date_br(identity.request_date), 'coins','Saldo quitado na operação', _money_br(identity.settled_loan_balance)),
        ('calendar','Data do crédito', _date_br(settings.credit_date), 'banknote','Valor líquido', _money_br(identity.net_amount)),
        ('people','Margem consignável', _money_br(identity.margin_amount), 'document','IOF', _money_br(identity.iof_amount)),
        ('money','Valor contratado', _money_br(settings.initial_balance, settings.rounding.money_places), 'calculator','Taxa administrativa', _money_br(identity.administrative_fee)),
        ('shield','FGQC - concessão', _money_br(identity.fgqc_concession), 'document','Valor da parcela base', 'Não se aplica'),
    ]
    data=[]
    for k1,l1,v1,k2,l2,v2 in values:
        data.append([_tiny_icon(k1),Paragraph(l1,styles['label']),Paragraph(v1,styles['body_compact']),_tiny_icon(k2),Paragraph(l2,styles['label']),Paragraph(v2,styles['body_compact'])])
    table=Table(data,colWidths=[10*mm,38*mm,40*mm,10*mm,40*mm,38*mm])
    table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),0.40,PDF_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('BACKGROUND',(0,0),(-1,-1),WHITE),
        ('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),4.0),('BOTTOMPADDING',(0,0),(-1,-1),4.0),
        ('ALIGN',(0,0),(0,-1),'CENTER'),('ALIGN',(3,0),(3,-1),'CENTER')
    ]))
    return table


def _callout(title: str, text: str, styles, accent=TEAL, background=SOFT_BLUE) -> Table:
    icon = _approved_icon('info', 10*mm)
    content = Paragraph(f'<b>{title}</b><br/>{text}', styles['callout'])
    t=Table([[icon,content]],colWidths=[13*mm,CONTENT_WIDTH-13*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),PDF_BLUE_LIGHT),('BOX',(0,0),(-1,-1),0.55,PDF_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)
    ])); return t


_technical_foundation_v0104 = _technical_foundation

def _foundation_block(kind: str, title: str, paragraphs: list[str], styles) -> Table:
    body=[Paragraph(title,styles['subsection'])]+[Paragraph(p,styles['body_justified']) for p in paragraphs]
    inner=Table([[x] for x in body],colWidths=[CONTENT_WIDTH-18*mm])
    inner.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),1)]))
    outer=Table([[_approved_icon(kind,12*mm),inner]],colWidths=[17*mm,CONTENT_WIDTH-17*mm])
    outer.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),('TOPPADDING',(0,0),(-1,-1),1),('BOTTOMPADDING',(0,0),(-1,-1),4)]))
    return outer


def _technical_foundation(settings: ContractSettings, styles) -> list:
    if settings.modality_code != 'MOD_001':
        return _technical_foundation_v0104(settings, styles)
    return [
        _foundation_block('balance','Natureza da metodologia de amortização',[
            'A modalidade variável utiliza uma lógica de amortização vinculada ao saldo devedor atualizado e ao prazo remanescente. A prestação é formada por juros e amortização, mas a amortização não deve ser descrita como uma quantia numericamente imutável em todas as parcelas: ela é recalculada em cada ciclo a partir do saldo atualizado, deduzidos os juros do período, dividido pela quantidade de prestações remanescentes.',
            'Os juros remuneram o capital utilizado no período e não reduzem o principal. A redução do saldo decorre da amortização. A correção monetária integra a evolução diária do saldo, segundo o índice e a defasagem definidos, e pode alterar a base sobre a qual a amortização do ciclo será apurada.'
        ],styles),
        Table([['']],colWidths=[CONTENT_WIDTH],rowHeights=[1*mm],style=[('LINEBELOW',(0,0),(-1,-1),0.5,HexColor('#CAD8E6'))]),
        _foundation_block('block','Ausência de incorporação de juros vencidos ao principal',[
            'Na memória examinada, os juros são apropriados diariamente sobre o saldo vigente e liquidados como componente da prestação. Não se verifica a transferência de juros vencidos para o principal com o objetivo de formar uma nova base de incidência. A variação das prestações decorre da combinação entre saldo remanescente, correção monetária, juros do período e prazo ainda existente.',
            'Para a correta leitura da evolução, devem ser distinguidos a atualização monetária do saldo, os juros remuneratórios do período e a amortização do principal. Somente a amortização reduz o saldo devedor principal.'
        ],styles),
    ]


def _step_icon_for_number(n:int)->str:
    return {1:'calendar',2:'percent',3:'search',4:'index',5:'date',6:'cycle',7:'coins',8:'lock'}.get(n,'info')


def _step_card(number: int, title: str, text: str, styles, width: float = 85 * mm) -> Table:
    badge=Table([[Paragraph(f'PASSO {number:02d}',styles['step_number'])]],colWidths=[20*mm],rowHeights=[6.5*mm])
    badge.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),PDF_BLUE),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    head=Table([[badge,Paragraph(title,styles['step_title'])]],colWidths=[22*mm,width-32*mm])
    head.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),1)]))
    body=Table([[_approved_icon(_step_icon_for_number(number),11*mm),Paragraph(text,styles['step_body'])]],colWidths=[14*mm,width-24*mm])
    body.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),0)]))
    card=Table([[head],[body]],colWidths=[width-8*mm],rowHeights=[8*mm,24*mm])
    card.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),WHITE),('BOX',(0,0),(-1,-1),0.55,PDF_LINE),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]))
    return card


def _approved_formula_card(number:int,title:str,formula:str,defs:list[tuple[str,str]],note:str,styles,width:float=56.5*mm)->Table:
    badge=Table([[Paragraph(f'{number:02d}',ParagraphStyle(f'fnum{number}',fontName='Helvetica-Bold',fontSize=7.0,textColor=WHITE,alignment=TA_CENTER))]],colWidths=[7*mm],rowHeights=[6*mm])
    badge.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),PDF_BLUE),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    head=Table([[badge,Paragraph(title.upper(),styles['formula_title'])]],colWidths=[8*mm,width-12*mm]); head.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0)]))
    rows=[[Paragraph('<b>VARIÁVEL</b>',styles['micro']),Paragraph('<b>SIGNIFICADO</b>',styles['micro'])]]+[[Paragraph(a,styles['micro']),Paragraph(b,styles['micro'])] for a,b in defs]
    defs_t=Table(rows,colWidths=[14*mm,width-20*mm]); defs_t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),HexColor('#EDF4FA')),('GRID',(0,0),(-1,-1),0.3,PDF_LINE),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),2),('RIGHTPADDING',(0,0),(-1,-1),2),('TOPPADDING',(0,0),(-1,-1),1.5),('BOTTOMPADDING',(0,0),(-1,-1),1.5)]))
    body=Table([[head],[Paragraph(formula,styles['formula_big'])],[defs_t],[Paragraph(f'<b>Aplicação:</b> {note}',styles['formula_note'])]],colWidths=[width-6*mm])
    body.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),1),('RIGHTPADDING',(0,0),(-1,-1),1),('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2)]))
    shell=Table([[body]],colWidths=[width]); shell.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),WHITE),('BOX',(0,0),(-1,-1),0.55,PDF_LINE),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3)])); return shell


def _variable_formula_overview(settings: ContractSettings, result: EvolutionResult, styles) -> list:
    first=result.installments[0]
    cards=[
      _approved_formula_card(1,'Conversão da taxa anual em mensal equivalente','i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1',[('i<sub>a</sub>','taxa anual contratual'),('i<sub>m</sub>','taxa mensal equivalente')],f'taxa anual {_percent_br(settings.annual_interest_rate,settings.rounding.percentage_display_places)} convertida para {_percent_br(result.monthly_interest_rate,settings.rounding.percentage_display_places)}.',styles),
      _approved_formula_card(2,'Taxa diária de juros','i<sub>d</sub> = (1 + i<sub>m</sub>)<super>1/DC</super> - 1',[('i<sub>d</sub>','taxa diária equivalente'),('i<sub>m</sub>','taxa mensal equivalente'),('DC','dias corridos da competência')],'a taxa diária é recalculada conforme os dias corridos de cada competência.',styles),
      _approved_formula_card(3,'Correção monetária diária','c<sub>d</sub> = (1 + c<sub>m</sub>)<super>1/DU</super> - 1',[('c<sub>d</sub>','taxa diária de correção'),('c<sub>m</sub>','INPC mensal da referência'),('DU','dias úteis da competência')],_lag_formula_note(settings),styles),
      _approved_formula_card(4,'Juros do dia e juros do período','J<sub>d</sub> = SD<sub>0</sub> × i<sub>d</sub><br/>J<sub>período</sub> = Σ J<sub>d</sub>',[('J<sub>d</sub>','juros do dia'),('SD<sub>0</sub>','saldo de abertura do dia'),('i<sub>d</sub>','taxa diária de juros'),('J<sub>período</sub>','somatório dos juros diários')],f'o somatório da primeira prestação resultou em {_money_br(first.interest_amount)}.',styles),
      _approved_formula_card(5,'Evolução diária do saldo','SD<sub>J</sub> = SD<sub>0</sub> × (1 + i<sub>d</sub>)<br/>SD<sub>J+CM</sub> = SD<sub>J</sub> × (1 + c<sub>d</sub>)',[('SD<sub>0</sub>','saldo no início do dia'),('SD<sub>J</sub>','saldo após juros'),('SD<sub>J+CM</sub>','saldo após juros e correção'),('c<sub>d</sub>','taxa diária de correção')],'juros todos os dias; correção monetária apenas nos dias úteis.',styles),
      _approved_formula_card(6,'Amortização e prestação do ciclo','Base<sub>A</sub> = SD<sub>antes</sub> - J<sub>período</sub><br/>A = Base<sub>A</sub> / n<br/>P = A + J<sub>período</sub><br/>SD<sub>final</sub> = Base<sub>A</sub> - A',[('Base<sub>A</sub>','base principal sem juros do período'),('A','amortização'),('n','prestações remanescentes'),('P','prestação')],f'1ª prestação: base {_money_br(first.amortization_base)}, amortização {_money_br(first.regular_amortization)}, prestação {_money_br(first.installment_amount)}.',styles),
    ]
    grid=Table([cards[:3],cards[3:]],colWidths=[58*mm,58*mm,58*mm]); grid.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),1),('TOPPADDING',(0,0),(-1,-1),1),('BOTTOMPADDING',(0,0),(-1,-1),1)]))
    return [grid,Spacer(1,2*mm),Paragraph('Competências e taxas efetivamente utilizadas na primeira prestação',styles['subsection']),_segment_summary(settings,result,styles),Spacer(1,1.5*mm),_callout('Quando houver carência envolvendo mais de uma competência','A tabela acima refletirá todas as competências efetivamente processadas no ciclo da prestação analisada. Em carência de dois ciclos ou mais, será exibida uma linha por competência, permitindo conferir a taxa diária de juros, o índice de referência e a taxa diária de correção utilizados.',styles)]


def _installments_table(rows: Sequence[InstallmentRow], overrides, places: int, styles) -> Table:
    headers=['Nº','Referência','Saldo antes','Correção','Juros','Amortização','FGQC','Prestação','Saldo final']
    data=[[Paragraph(h,styles['table_head']) for h in headers]]
    for row in rows:
        vals=[str(row.installment_number),_date_br(row.reference_due_date),_money_br(_value(row,'balance_before_installment',overrides),places),_money_br(_value(row,'correction_amount',overrides),places),_money_br(_value(row,'interest_amount',overrides),places),_money_br(_value(row,'regular_amortization',overrides),places),_money_br(_value(row,'fgqc_amount',overrides),places),_money_br(_value(row,'installment_amount',overrides),places),_money_br(_value(row,'closing_balance',overrides),places)]
        data.append([Paragraph(vals[0],styles['table']),Paragraph(vals[1],styles['table']),*[Paragraph(v,styles['table_right']) for v in vals[2:]]])
    t=Table(data,colWidths=[7*mm,20*mm,24*mm,18*mm,18*mm,22*mm,17*mm,20*mm,22*mm],repeatRows=1,splitByRow=1)
    cmds=[('BACKGROUND',(0,0),(-1,0),PDF_BLUE),('GRID',(0,0),(-1,-1),0.35,PDF_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),2.5),('RIGHTPADDING',(0,0),(-1,-1),2.5),('TOPPADDING',(0,0),(-1,-1),3.2),('BOTTOMPADDING',(0,0),(-1,-1),3.2)]
    for i in range(1,len(data)):
        if i%2==0: cmds.append(('BACKGROUND',(0,i),(-1,i),HexColor('#F2F6FA')))
    t.setStyle(TableStyle(cmds)); return t


_detail_flowables_v0104 = _detail_flowables

def _detail_summary_table(row, overrides, places, styles):
    fields=[
      ('calendar','Saldo antes',_value(row,'balance_before_installment',overrides),'chart','Amortização',_value(row,'regular_amortization',overrides)),
      ('percent','Juros do período',_value(row,'interest_amount',overrides),'date','Prestação',_value(row,'installment_amount',overrides)),
      ('system','Correção monetária',_value(row,'correction_amount',overrides),'shield','FGQC',_value(row,'fgqc_amount',overrides)),
      ('target','Base de amortização',_decimal(row.amortization_base),'coins','Saldo final',_value(row,'closing_balance',overrides)),
    ]
    rows=[]
    lab=ParagraphStyle('sumlab_v105',fontName='Helvetica',fontSize=6.0,leading=7.3,textColor=PDF_TEXT)
    val=ParagraphStyle('sumval_v105',fontName='Helvetica-Bold',fontSize=6.2,leading=7.5,textColor=PDF_BLUE,alignment=TA_RIGHT)
    for k1,l1,v1,k2,l2,v2 in fields:
        rows.append([_approved_icon(k1,6.5*mm),Paragraph(l1,lab),Paragraph(_money_br(v1,places),val),_approved_icon(k2,6.5*mm),Paragraph(l2,lab),Paragraph(_money_br(v2,places),val)])
    t=Table(rows,colWidths=[8*mm,29*mm,45*mm,8*mm,29*mm,45*mm])
    t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.35,PDF_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('BACKGROUND',(0,0),(-1,-1),WHITE),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),2.5),('BOTTOMPADDING',(0,0),(-1,-1),2.5)])); return t


def _equation_line(text,styles):
    t=Table([[Paragraph(text,styles['formula_small'])]],colWidths=[CONTENT_WIDTH-10*mm])
    t.setStyle(TableStyle([('LINEBEFORE',(0,0),(0,-1),2.2,PDF_ORANGE),('BOX',(0,0),(-1,-1),0.35,PDF_LINE),('BACKGROUND',(0,0),(-1,-1),WHITE),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),2.5),('BOTTOMPADDING',(0,0),(-1,-1),2.5)])); return t


def _detail_flowables(rows: Sequence[InstallmentRow], overrides: Mapping[int, Mapping[str, Decimal]], settings: ContractSettings, styles):
    if settings.modality_code != 'MOD_001':
        return _detail_flowables_v0104(rows,overrides,settings,styles)
    out=[]; places=settings.rounding.money_places
    for idx,row in enumerate(rows):
        # Só inicia uma nova prestação quando houver espaço para título, narrativa e início da tabela.
        out.append(Spacer(1, 3.0 * mm if idx else 2.2 * mm))
        out.append(CondPageBreak(72 * mm))
        interest=_value(row,'interest_amount',overrides); correction=_value(row,'correction_amount',overrides); amort=_value(row,'regular_amortization',overrides); inst=_value(row,'installment_amount',overrides); closing=_value(row,'closing_balance',overrides); before=_value(row,'balance_before_installment',overrides); base=_decimal(row.amortization_base); opening=_decimal(row.opening_balance); extra=_decimal(row.extraordinary_amortization); n=row.remaining_installments+1
        title=Paragraph(f'Prestação {row.installment_number} - referência {_date_br(row.reference_due_date)}',ParagraphStyle(f'dtitle{idx}',fontName='Helvetica-Bold',fontSize=11.1,leading=13.0,textColor=PDF_BLUE,spaceAfter=4))
        extra_txt=(f' Também houve amortização extraordinária de <b>{_money_br(extra,places)}</b>.' if extra else '')
        narrative=Paragraph(f'O saldo acumulado antes da prestação totalizou <b>{_money_br(before,places)}</b>. Nesse montante estão refletidos os juros do período de <b>{_money_br(interest,places)}</b> e a correção monetária de <b>{_money_br(correction,places)}</b>. Para apurar a amortização, os juros do período são excluídos da base, obtendo-se <b>{_money_br(base,places)}</b>. A divisão pelo prazo remanescente resultou em amortização de <b>{_money_br(amort,places)}</b>. A prestação corresponde à soma da amortização com os juros e totalizou <b>{_money_br(inst,places)}</b>. Após o pagamento, o saldo principal passou a <b>{_money_br(closing,places)}</b>.{extra_txt}',styles['body_justified'])
        calc=[
          ('Saldo antes da prestação','saldo inicial do ciclo + juros do período + correção monetária',f'{_money_br(opening,places)} + {_money_br(interest,places)} {_signed_money(correction,places)}',before,'calendar'),
          ('Base de amortização','saldo antes - juros do período',f'{_money_br(before,places)} - {_money_br(interest,places)}',base,'percent'),
          ('Amortização','Base_A / prestações remanescentes',f'{_money_br(base,places)} / {n}',amort,'chart'),
          ('Prestação','amortização + juros do período',f'{_money_br(amort,places)} + {_money_br(interest,places)}',inst,'calendar'),
          ('Saldo final','Base_A - amortização',f'{_money_br(base,places)} - {_money_br(amort,places)}',closing,'money'),
        ]
        head=[[Paragraph('Etapa',styles['table_head']),Paragraph('Descrição',styles['table_head']),Paragraph('Substituição numérica',styles['table_head']),Paragraph('Resultado',styles['table_head'])]]
        desc=ParagraphStyle(f'desc{idx}',fontName='Helvetica',fontSize=6.2,leading=7.5,textColor=PDF_TEXT)
        res=ParagraphStyle(f'res{idx}',fontName='Helvetica-Bold',fontSize=6.4,leading=7.6,textColor=PDF_BLUE,alignment=TA_RIGHT)
        data=head
        for num,(label,rule,subst,val,kind) in enumerate(calc,1):
            badge=Table([[Paragraph(str(num),ParagraphStyle(f'bn{idx}{num}',fontName='Helvetica-Bold',fontSize=6.4,textColor=WHITE,alignment=TA_CENTER)),_approved_icon(kind,6*mm)]],colWidths=[6*mm,7*mm]); badge.setStyle(TableStyle([('BACKGROUND',(0,0),(0,0),PDF_BLUE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),1),('RIGHTPADDING',(0,0),(-1,-1),1),('TOPPADDING',(0,0),(-1,-1),1),('BOTTOMPADDING',(0,0),(-1,-1),1)]))
            data.append([badge,Paragraph(f'<b>{label}</b><br/><font color="#667A8D">{rule}</font>',desc),Paragraph(subst,desc),Paragraph(_money_br(val,places),res)])
        table=Table(data,colWidths=[22*mm,52*mm,68*mm,34*mm],repeatRows=1,splitByRow=1); cmds=[('BACKGROUND',(0,0),(-1,0),PDF_BLUE),('GRID',(0,0),(-1,-1),0.35,PDF_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]
        for rr in range(1,len(data)):
            if rr%2==0: cmds.append(('BACKGROUND',(0,rr),(-1,rr),HexColor('#F2F6FA')))
        table.setStyle(TableStyle(cmds))
        out.append(KeepTogether([title,narrative,Spacer(1,1.6*mm)]))
        out.append(table)
        out.append(Spacer(1,1.8*mm))
        out.append(KeepTogether([
            Paragraph(f'Resumo da Prestação {row.installment_number}',styles['subsection']),
            _detail_summary_table(row,overrides,places,styles),
        ]))
        out.extend([
            Spacer(1,1.6*mm),
            _equation_line(f'Base_A = {_money_br(before,places)} - {_money_br(interest,places)} = {_money_br(base,places)}',styles),
            Spacer(1,0.6*mm),
            _equation_line(f'A = {_money_br(base,places)} / {n} = {_money_br(amort,places)}',styles),
            Spacer(1,0.6*mm),
            _equation_line(f'P = {_money_br(amort,places)} + {_money_br(interest,places)} = {_money_br(inst,places)}',styles),
        ])
    return out



def _daily_sequence(settings: ContractSettings, styles) -> Table:
    items = [
        ('money', 'SALDO INICIAL', HexColor('#006E8A')),
        ('percent', 'JUROS DO DIA', HexColor('#008A9A')),
    ]
    if settings.index_code:
        items.append(('date', 'CORREÇÃO EM<br/>DIA ÚTIL', PDF_BLUE))
    items.extend([
        ('calendar', 'EVENTO /<br/>AMORTIZAÇÃO', PDF_ORANGE),
        ('chart', 'SALDO FINAL', HexColor('#23824D')),
    ])
    cells = []
    widths = []
    for i, (kind, label, color) in enumerate(items):
        icon = _approved_icon(kind, 9 * mm)
        txt = Paragraph(
            label,
            ParagraphStyle(
                f'seqtxt_v105_{i}', fontName='Helvetica-Bold', fontSize=6.2,
                leading=7.1, textColor=WHITE, alignment=TA_CENTER,
            ),
        )
        box = Table([[icon], [txt]], colWidths=[29 * mm], rowHeights=[11 * mm, 13 * mm])
        box.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), color),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ]))
        cells.append(box)
        widths.append(29 * mm)
        if i < len(items) - 1:
            cells.append(Paragraph(
                '→',
                ParagraphStyle(
                    f'arrow_v105_{i}', fontName='Helvetica-Bold', fontSize=13,
                    textColor=PDF_MUTED, alignment=TA_CENTER,
                ),
            ))
            widths.append(7 * mm)
    table = Table([cells], colWidths=widths, hAlign='CENTER')
    table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))
    return table



# === AJUSTES DE FIDELIDADE VISUAL E DEFASAGEM v0.10.6 ===
_MONTHS_PT = ('janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro')


def _lag_label(months: int) -> str:
    months = int(months or 0)
    if months == 0:
        return 'sem defasagem'
    return f'defasagem de {months} mês' if months == 1 else f'defasagem de {months} meses'


def _competence_2120_example(settings: ContractSettings) -> str:
    lag = int(settings.index_lag_months or 0)
    index_name = settings.index_code or 'índice selecionado'
    reference_month = _MONTHS_PT[(0 - lag) % 12]
    lag_text = _lag_label(lag)
    return (
        'A competência segue o ciclo de 21 a 20. Cada competência inicia no dia 21 de um mês civil e termina '
        'no dia 20 do mês civil seguinte.<br/><br/>'
        '<b>Exemplo:</b> na competência de 21/01 a 20/02, são apropriados juros correspondentes a 31 dias corridos e, '
        f'considerando {lag_text}, utiliza-se o {index_name} de {reference_month}.'
    )


def _lag_formula_note(settings: ContractSettings) -> str:
    return f'índice {settings.index_code or "selecionado"}, com {_lag_label(settings.index_lag_months)}.'
# === FIM DOS AJUSTES v0.10.6 ===

# === FIM DO DESIGN SYSTEM APROVADO v0.10.5 ===


# === UNIFICAÇÃO VISUAL CREDPLAN FIXO v0.10.8 ===
_technical_foundation_v0107 = _technical_foundation
_formula_application_v0107 = _formula_application
_detail_flowables_v0107 = _detail_flowables


def _technical_foundation(settings: ContractSettings, styles) -> list:
    if settings.modality_code != 'MOD_002':
        return _technical_foundation_v0107(settings, styles)
    return [
        _foundation_block('system', 'Natureza da metodologia de amortização - Tabela Price', [
            'O Credplan Fixo utiliza a <b>Tabela Price</b>. A prestação é determinada a partir do saldo financiado, da taxa periódica equivalente e do prazo contratual. Mantidos esses parâmetros, a prestação permanece fixa, enquanto sua composição interna se altera ao longo do contrato.',
            'Em cada ciclo, os juros remuneratórios são apropriados sobre o saldo devedor vigente. A amortização corresponde à diferença entre a prestação e os juros do período. Por essa razão, no início do contrato a participação dos juros tende a ser maior e, com a redução do saldo, a parcela destinada à amortização tende a aumentar.'
        ], styles),
        Table([['']], colWidths=[CONTENT_WIDTH], rowHeights=[1*mm], style=[('LINEBELOW',(0,0),(-1,-1),0.5,HexColor('#CAD8E6'))]),
        _foundation_block('block', 'Ausência de correção monetária e segregação dos juros', [
            'A modalidade Credplan Fixo <b>não utiliza correção monetária</b>. A evolução decorre da apropriação dos juros remuneratórios previstos contratualmente e do abatimento da prestação nos vencimentos.',
            'Para a correta leitura da memória, juros, amortização e saldo principal são apresentados de forma segregada. Os juros do período integram a composição da prestação; a redução do principal decorre da amortização. A demonstração separada desses componentes permite conferir a evolução do saldo sem atribuir à modalidade qualquer atualização por índice de preços.'
        ], styles),
    ]


def _fixed_formula_overview(settings: ContractSettings, result: EvolutionResult, styles) -> list:
    first = result.installments[0]
    price_type = 1 if settings.payment_timing == 1 else 0
    timing_note = 'pagamento no início do período (tipo 1)' if price_type == 1 else 'pagamento no fim do período (tipo 0)'
    if price_type == 1:
        price_formula = 'PMT = [PV × i<sub>m</sub> / (1 - (1 + i<sub>m</sub>)<super>-n</super>)] × [1 / (1 + i<sub>m</sub>)]'
    else:
        price_formula = 'PMT = PV × i<sub>m</sub> / [1 - (1 + i<sub>m</sub>)<super>-n</super>]'
    cards = [
        _approved_formula_card(1, 'Conversão da taxa anual em mensal equivalente', 'i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1', [('i<sub>a</sub>','taxa anual contratual'),('i<sub>m</sub>','taxa mensal equivalente')], f'taxa anual {_percent_br(settings.annual_interest_rate,settings.rounding.percentage_display_places)} convertida para {_percent_br(result.monthly_interest_rate,settings.rounding.percentage_display_places)}.', styles),
        _approved_formula_card(2, 'Taxa diária de juros', 'i<sub>d</sub> = (1 + i<sub>m</sub>)<super>1/DC</super> - 1', [('i<sub>d</sub>','taxa diária equivalente'),('i<sub>m</sub>','taxa mensal equivalente'),('DC','dias corridos da competência')], 'a taxa diária é recalculada conforme a quantidade de dias corridos da competência.', styles),
        _approved_formula_card(3, 'Prestação pela Tabela Price', price_formula, [('PMT','prestação contratual'),('PV','saldo financiado/base Price'),('i<sub>m</sub>','taxa mensal equivalente'),('n','quantidade de prestações')], f'{timing_note}; primeira prestação calculada em {_money_br(first.installment_amount)}.', styles),
        _approved_formula_card(4, 'Juros do dia e juros do período', 'J<sub>d</sub> = SD<sub>0</sub> × i<sub>d</sub><br/>J<sub>período</sub> = Σ J<sub>d</sub>', [('J<sub>d</sub>','juros do dia'),('SD<sub>0</sub>','saldo de abertura do dia'),('i<sub>d</sub>','taxa diária equivalente'),('J<sub>período</sub>','somatório dos juros diários')], f'o somatório dos juros da primeira prestação resultou em {_money_br(first.interest_amount)}.', styles),
        _approved_formula_card(5, 'Evolução diária do saldo', 'SD<sub>J</sub> = SD<sub>0</sub> × (1 + i<sub>d</sub>)', [('SD<sub>0</sub>','saldo no início do dia'),('SD<sub>J</sub>','saldo após a apropriação dos juros'),('i<sub>d</sub>','taxa diária equivalente')], 'a modalidade não possui correção monetária; a evolução diária decorre apenas dos juros e dos eventos financeiros aplicáveis.', styles),
        _approved_formula_card(6, 'Decomposição da prestação e fechamento', 'A = PMT - J<sub>período</sub><br/>SD<sub>final</sub> = SD<sub>antes</sub> - PMT', [('A','amortização do principal'),('PMT','prestação Price'),('J<sub>período</sub>','juros acumulados do período'),('SD<sub>final</sub>','saldo principal após a prestação')], f'1ª prestação: juros {_money_br(first.interest_amount)}, amortização {_money_br(first.regular_amortization)} e saldo final {_money_br(first.closing_balance)}.', styles),
    ]
    grid = Table([cards[:3], cards[3:]], colWidths=[58*mm,58*mm,58*mm])
    grid.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),1),('TOPPADDING',(0,0),(-1,-1),1),('BOTTOMPADDING',(0,0),(-1,-1),1)]))
    return [grid]


def _formula_application(settings: ContractSettings, result: EvolutionResult, styles) -> list:
    if settings.modality_code == 'MOD_002':
        return _fixed_formula_overview(settings, result, styles)
    return _formula_application_v0107(settings, result, styles)


def _fixed_detail_summary_table(row, overrides, places, styles):
    fields = [
        ('calendar','Saldo antes',_value(row,'balance_before_installment',overrides),'chart','Amortização',_value(row,'regular_amortization',overrides)),
        ('percent','Juros do período',_value(row,'interest_amount',overrides),'date','Prestação Price',_value(row,'installment_amount',overrides)),
        ('system','Sistema','Tabela Price','shield','Correção monetária','Não se aplica'),
        ('target','Saldo inicial',_decimal(row.opening_balance),'coins','Saldo final',_value(row,'closing_balance',overrides)),
    ]
    rows=[]
    lab=ParagraphStyle('sumlab_fixed_v108',fontName='Helvetica',fontSize=6.0,leading=7.3,textColor=PDF_TEXT)
    val=ParagraphStyle('sumval_fixed_v108',fontName='Helvetica-Bold',fontSize=6.2,leading=7.5,textColor=PDF_BLUE,alignment=TA_RIGHT)
    for k1,l1,v1,k2,l2,v2 in fields:
        def fmt(v):
            return _money_br(v,places) if isinstance(v, (Decimal, int, float)) else str(v)
        rows.append([_approved_icon(k1,6.5*mm),Paragraph(l1,lab),Paragraph(fmt(v1),val),_approved_icon(k2,6.5*mm),Paragraph(l2,lab),Paragraph(fmt(v2),val)])
    t=Table(rows,colWidths=[8*mm,29*mm,45*mm,8*mm,29*mm,45*mm])
    t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.35,PDF_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('BACKGROUND',(0,0),(-1,-1),WHITE),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),2.5),('BOTTOMPADDING',(0,0),(-1,-1),2.5)]))
    return t


def _fixed_detail_flowables(rows, overrides, settings, styles):
    out=[]; places=settings.rounding.money_places
    for idx,row in enumerate(rows):
        out.append(Spacer(1, 3.0*mm if idx else 2.2*mm)); out.append(CondPageBreak(72*mm))
        opening=_decimal(row.opening_balance); interest=_value(row,'interest_amount',overrides); amort=_value(row,'regular_amortization',overrides); inst=_value(row,'installment_amount',overrides); closing=_value(row,'closing_balance',overrides); before=_value(row,'balance_before_installment',overrides); fgqc=_value(row,'fgqc_amount',overrides); extra=_decimal(row.extraordinary_amortization)
        title=Paragraph(f'Prestação {row.installment_number} - referência {_date_br(row.reference_due_date)}',ParagraphStyle(f'fdtitle{idx}',fontName='Helvetica-Bold',fontSize=11.1,leading=13.0,textColor=PDF_BLUE,spaceAfter=4))
        extra_txt=(f' Também houve amortização extraordinária de <b>{_money_br(extra,places)}</b>.' if extra else '')
        narrative=Paragraph(f'O saldo inicial do ciclo era <b>{_money_br(opening,places)}</b>. Os juros apropriados diariamente totalizaram <b>{_money_br(interest,places)}</b>, formando saldo antes da prestação de <b>{_money_br(before,places)}</b>. A prestação pela Tabela Price foi de <b>{_money_br(inst,places)}</b>. Desse valor, <b>{_money_br(interest,places)}</b> correspondem aos juros do período e <b>{_money_br(amort,places)}</b> à amortização do principal. Após o pagamento, o saldo devedor principal passou a <b>{_money_br(closing,places)}</b>. Não há correção monetária nesta modalidade.{extra_txt}', styles['body_justified'])
        calc=[('Saldo antes da prestação','saldo inicial + juros do período' + (' - amortização extraordinária' if extra else ''), f'{_money_br(opening,places)} + {_money_br(interest,places)}' + (f' - {_money_br(extra,places)}' if extra else ''), before,'calendar'),('Juros do período','somatório dos juros apropriados diariamente','Σ Jd',interest,'percent'),('Amortização','prestação Price - juros do período',f'{_money_br(inst,places)} - {_money_br(interest,places)}',amort,'chart'),('Prestação Price','prestação contratual calculada pela Tabela Price',_money_br(inst,places),inst,'date'),('Saldo final','saldo antes - prestação',f'{_money_br(before,places)} - {_money_br(inst,places)}',closing,'money')]
        data=[[Paragraph('Etapa',styles['table_head']),Paragraph('Descrição',styles['table_head']),Paragraph('Substituição numérica',styles['table_head']),Paragraph('Resultado',styles['table_head'])]]
        desc=ParagraphStyle(f'fdesc{idx}',fontName='Helvetica',fontSize=6.2,leading=7.5,textColor=PDF_TEXT); res=ParagraphStyle(f'fres{idx}',fontName='Helvetica-Bold',fontSize=6.4,leading=7.6,textColor=PDF_BLUE,alignment=TA_RIGHT)
        for num,(label,rule,subst,val_,kind) in enumerate(calc,1):
            badge=Table([[Paragraph(str(num),ParagraphStyle(f'fbn{idx}{num}',fontName='Helvetica-Bold',fontSize=6.4,textColor=WHITE,alignment=TA_CENTER)),_approved_icon(kind,6*mm)]],colWidths=[6*mm,7*mm]); badge.setStyle(TableStyle([('BACKGROUND',(0,0),(0,0),PDF_BLUE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),1),('RIGHTPADDING',(0,0),(-1,-1),1),('TOPPADDING',(0,0),(-1,-1),1),('BOTTOMPADDING',(0,0),(-1,-1),1)]))
            data.append([badge,Paragraph(f'<b>{label}</b><br/><font color="#667A8D">{rule}</font>',desc),Paragraph(subst,desc),Paragraph(_money_br(val_,places),res)])
        table=Table(data,colWidths=[22*mm,52*mm,68*mm,34*mm],repeatRows=1,splitByRow=1); cmds=[('BACKGROUND',(0,0),(-1,0),PDF_BLUE),('GRID',(0,0),(-1,-1),0.35,PDF_LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]
        for rr in range(1,len(data)):
            if rr%2==0: cmds.append(('BACKGROUND',(0,rr),(-1,rr),HexColor('#F2F6FA')))
        table.setStyle(TableStyle(cmds)); out.append(KeepTogether([title,narrative,Spacer(1,1.6*mm)])); out.append(table)
        if fgqc != 0: out.extend([Spacer(1,1.2*mm),Paragraph(f'<b>FGQC:</b> {_money_br(fgqc,places)} demonstrado separadamente; não altera a decomposição financeira da prestação acima.',styles['small'])])
        out.append(Spacer(1,1.8*mm)); out.append(KeepTogether([Paragraph(f'Resumo da Prestação {row.installment_number}',styles['subsection']),_fixed_detail_summary_table(row,overrides,places,styles)]))
        out.extend([Spacer(1,1.6*mm),_equation_line(f'J_período = Σ Jd = {_money_br(interest,places)}',styles),Spacer(1,0.6*mm),_equation_line(f'A = PMT - J_período = {_money_br(inst,places)} - {_money_br(interest,places)} = {_money_br(amort,places)}',styles),Spacer(1,0.6*mm),_equation_line(f'SD_final = SD_antes - PMT = {_money_br(before,places)} - {_money_br(inst,places)} = {_money_br(closing,places)}',styles)])
    return out


def _detail_flowables(rows, overrides, settings, styles):
    if settings.modality_code == 'MOD_002':
        return _fixed_detail_flowables(rows, overrides, settings, styles)
    return _detail_flowables_v0107(rows, overrides, settings, styles)
