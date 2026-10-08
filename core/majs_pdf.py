from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from html import escape
from io import BytesIO
from pathlib import Path
import re

from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from .majs import (
    MAJSResult,
    SEM_CORRECAO,
    majs_correction_mode_label,
    majs_interest_mode_label,
    majs_yearfrac_basis_label,
)
from .pdf_theme import (
    PETROLEUM, PETROLEUM_DARK, PETROLEUM_DEEP, PETROLEUM_SOFT, ORANGE, ORANGE_SOFT,
    INK, MUTED, LINE, WHITE, SOFT, WARNING, SANS, SANS_BOLD,
)


BLUE_SOFT = HexColor("#F1F6FB")
SOFT_ALT = HexColor("#F4F8FA")

PORTRAIT = A4
LANDSCAPE = landscape(A4)
P_MARGIN = 12 * mm
L_MARGIN = 6 * mm
P_CONTENT = PORTRAIT[0] - 2 * P_MARGIN
L_CONTENT = LANDSCAPE[0] - 2 * L_MARGIN


@dataclass(frozen=True)
class MAJSReportIdentity:
    contract_number: str = ""
    participant_name: str = ""
    document_reference: str = ""
    elaborator: str = ""
    validator: str = ""
    observation: str = ""


def _money(value: Decimal | int | float) -> str:
    number = Decimal(str(value))
    prefix = "-R$ " if number < 0 else "R$ "
    text = f"{abs(number):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return prefix + text


def _pct(value: Decimal | int | float, places: int = 4) -> str:
    number = Decimal(str(value)) * Decimal("100")
    text = f"{number:,.{places}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return text + "%"


def _date(value: date | None) -> str:
    return value.strftime("%d/%m/%Y") if value else "-"


def _clean_filename(value: str) -> str:
    text = re.sub(r"[^A-Za-z0-9_-]+", "_", value.strip())
    return text.strip("_") or "calculo"


def majs_opinion_pdf_filename(result: MAJSResult, identity: MAJSReportIdentity | None = None) -> str:
    identity = identity or MAJSReportIdentity()
    suffix = _clean_filename(identity.contract_number) if identity.contract_number.strip() else result.settings.base_date.strftime("%Y%m%d")
    return f"parecer_majs_{suffix}.pdf"


def _has_payments(result: MAJSResult) -> bool:
    """True quando há pagamentos/créditos ou diferença de quitação sem novação."""
    return any(Decimal(str(r.amount_paid)) != 0 or r.origin == "Quitação antecipada" for r in result.differences)


def _styles() -> dict[str, ParagraphStyle]:
    sample = getSampleStyleSheet()
    return {
        "body": ParagraphStyle(
            "body", parent=sample["BodyText"], fontName=SANS, fontSize=8.1,
            leading=11.2, textColor=INK, spaceAfter=3,
        ),
        "body_small": ParagraphStyle(
            "body_small", parent=sample["BodyText"], fontName=SANS, fontSize=6.7,
            leading=8.8, textColor=MUTED,
        ),
        "section": ParagraphStyle(
            "section", parent=sample["Heading2"], fontName=SANS_BOLD, fontSize=10.3,
            leading=12.0, textColor=PETROLEUM_DEEP, spaceBefore=1, spaceAfter=3,
        ),
        "label": ParagraphStyle(
            "label", parent=sample["BodyText"], fontName=SANS_BOLD, fontSize=6.3,
            leading=7.6, textColor=PETROLEUM_DEEP,
        ),
        "value": ParagraphStyle(
            "value", parent=sample["BodyText"], fontName=SANS, fontSize=7.0,
            leading=8.4, textColor=INK,
        ),
        "table_head": ParagraphStyle(
            "table_head", parent=sample["BodyText"], fontName=SANS_BOLD, fontSize=5.7,
            leading=6.6, textColor=WHITE, alignment=TA_CENTER,
        ),
        "table": ParagraphStyle(
            "table", parent=sample["BodyText"], fontName=SANS, fontSize=5.9,
            leading=7.2, textColor=INK,
        ),
        "table_center": ParagraphStyle(
            "table_center", parent=sample["BodyText"], fontName=SANS, fontSize=5.9,
            leading=7.2, textColor=INK, alignment=TA_CENTER,
        ),
        "table_right": ParagraphStyle(
            "table_right", parent=sample["BodyText"], fontName=SANS, fontSize=5.9,
            leading=7.2, textColor=INK, alignment=TA_RIGHT,
        ),
        "tiny_head": ParagraphStyle(
            "tiny_head", parent=sample["BodyText"], fontName=SANS_BOLD, fontSize=4.3,
            leading=4.9, textColor=WHITE, alignment=TA_CENTER,
        ),
        "tiny": ParagraphStyle(
            "tiny", parent=sample["BodyText"], fontName=SANS, fontSize=4.4,
            leading=5.1, textColor=INK,
        ),
        "tiny_center": ParagraphStyle(
            "tiny_center", parent=sample["BodyText"], fontName=SANS, fontSize=4.4,
            leading=5.1, textColor=INK, alignment=TA_CENTER,
        ),
        "tiny_right": ParagraphStyle(
            "tiny_right", parent=sample["BodyText"], fontName=SANS, fontSize=4.4,
            leading=5.1, textColor=INK, alignment=TA_RIGHT,
        ),
        "difference_head": ParagraphStyle(
            "difference_head", parent=sample["BodyText"], fontName=SANS_BOLD, fontSize=6.5,
            leading=7.6, textColor=WHITE, alignment=TA_CENTER,
        ),
        "difference": ParagraphStyle(
            "difference", parent=sample["BodyText"], fontName=SANS, fontSize=6.6,
            leading=8.0, textColor=INK,
        ),
        "difference_center": ParagraphStyle(
            "difference_center", parent=sample["BodyText"], fontName=SANS, fontSize=6.6,
            leading=8.0, textColor=INK, alignment=TA_CENTER,
        ),
        "difference_right": ParagraphStyle(
            "difference_right", parent=sample["BodyText"], fontName=SANS, fontSize=6.6,
            leading=8.0, textColor=INK, alignment=TA_RIGHT,
        ),
        "note": ParagraphStyle(
            "note", parent=sample["BodyText"], fontName=SANS, fontSize=6.5,
            leading=8.6, textColor=PETROLEUM_DEEP,
        ),
        "conclusion": ParagraphStyle(
            "conclusion", parent=sample["BodyText"], fontName=SANS, fontSize=7.7,
            leading=10.3, textColor=INK,
        ),
    }



class _PremiumKpiCards(Flowable):
    """Cards executivos usados no resumo: mais próximos do mockup aprovado."""

    def __init__(self, items, width: float, height: float = 28 * mm, gap: float = 2.0 * mm, emphasize_index: int | None = None):
        super().__init__()
        self.items = list(items)
        self.width = width
        self.height = height
        self.gap = gap
        self.emphasize_index = emphasize_index

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    @staticmethod
    def _icon(c, cx, cy, kind: str):
        # círculo suave + ícone vetorial simples, sem depender de fonte de símbolos
        c.setFillColor(HexColor('#EEF3F7'))
        c.circle(cx, cy, 4.9 * mm, fill=1, stroke=0)
        c.setStrokeColor(PETROLEUM_DEEP)
        c.setFillColor(PETROLEUM_DEEP)
        c.setLineWidth(0.9)
        k = kind.upper()
        if k in {'MONEY', 'PAID'}:
            c.circle(cx, cy, 3.25 * mm, fill=0, stroke=1)
            c.setFont('Helvetica-Bold', 6.1)
            c.drawCentredString(cx, cy - 1.8, 'R$')
            if k == 'PAID':
                # pequeno marcador de confirmação
                c.setLineWidth(1.1)
                c.line(cx + 2.0*mm, cy - 2.3*mm, cx + 2.8*mm, cy - 3.0*mm)
                c.line(cx + 2.8*mm, cy - 3.0*mm, cx + 4.0*mm, cy - 1.3*mm)
        elif k == 'WALLET':
            c.roundRect(cx - 3.1*mm, cy - 2.2*mm, 6.2*mm, 4.4*mm, 0.8*mm, fill=0, stroke=1)
            c.line(cx + 0.8*mm, cy + 0.7*mm, cx + 3.1*mm, cy + 0.7*mm)
            c.circle(cx + 1.8*mm, cy + 0.7*mm, 0.35*mm, fill=1, stroke=0)
        elif k == 'DELTA':
            c.setFont('Helvetica-Bold', 8.2)
            c.drawCentredString(cx, cy - 2.3, 'Δ')
        elif k == 'DATE':
            c.roundRect(cx - 3.2*mm, cy - 2.7*mm, 6.4*mm, 5.3*mm, 0.8*mm, fill=0, stroke=1)
            c.line(cx - 3.2*mm, cy + 0.8*mm, cx + 3.2*mm, cy + 0.8*mm)
            c.line(cx - 1.8*mm, cy + 3.0*mm, cx - 1.8*mm, cy + 1.8*mm)
            c.line(cx + 1.8*mm, cy + 3.0*mm, cx + 1.8*mm, cy + 1.8*mm)
        elif k == 'TREND':
            c.setLineWidth(1.25)
            c.line(cx - 3.5*mm, cy - 2.0*mm, cx - 1.2*mm, cy + 0.1*mm)
            c.line(cx - 1.2*mm, cy + 0.1*mm, cx + 0.4*mm, cy - 1.0*mm)
            c.line(cx + 0.4*mm, cy - 1.0*mm, cx + 3.0*mm, cy + 2.2*mm)
            c.line(cx + 3.0*mm, cy + 2.2*mm, cx + 1.8*mm, cy + 2.0*mm)
            c.line(cx + 3.0*mm, cy + 2.2*mm, cx + 2.9*mm, cy + 0.9*mm)
        elif k == 'BALANCE':
            c.line(cx, cy + 2.7*mm, cx, cy - 2.8*mm)
            c.line(cx - 3.1*mm, cy + 1.6*mm, cx + 3.1*mm, cy + 1.6*mm)
            c.line(cx - 2.1*mm, cy + 1.6*mm, cx - 3.0*mm, cy - 0.8*mm)
            c.line(cx + 2.1*mm, cy + 1.6*mm, cx + 3.0*mm, cy - 0.8*mm)
            c.line(cx - 3.8*mm, cy - 0.8*mm, cx - 2.2*mm, cy - 0.8*mm)
            c.line(cx + 2.2*mm, cy - 0.8*mm, cx + 3.8*mm, cy - 0.8*mm)
        else:
            c.setFont('Helvetica-Bold', 5.4)
            c.drawCentredString(cx, cy - 1.7, kind[:4])

    def _draw_card(self, c, x, w, label, value, icon, emphasized=False):
        h = self.height
        # sombra muito discreta para aproximar o aspecto dos mockups, sem virar "panfleto"
        c.setFillColor(HexColor('#E9EEF1'))
        c.roundRect(x + 0.55*mm, -0.55*mm, w, h, 3.0*mm, fill=1, stroke=0)
        c.setFillColor(BLUE_SOFT if emphasized else WHITE)
        c.setStrokeColor(HexColor('#C8D8DF'))
        c.setLineWidth(0.55)
        c.roundRect(x, 0, w, h, 3.0*mm, fill=1, stroke=1)

        cx = x + w / 2
        icon_y = h - 8.0*mm
        self._icon(c, cx, icon_y, icon)

        c.setFillColor(INK)
        c.setFont('Helvetica-Bold' if emphasized else 'Helvetica', 6.3)
        c.drawCentredString(cx, 9.2*mm, label[:38])

        size = 12.0 if emphasized else 11.4
        min_size = 7.2
        max_width = w - 4.0*mm
        while size > min_size and stringWidth(value, 'Helvetica-Bold', size) > max_width:
            size -= 0.3
        c.setFillColor(PETROLEUM_DEEP)
        c.setFont('Helvetica-Bold', size)
        c.drawCentredString(cx, 4.2*mm, value)

        # detalhe laranja curto e alinhado como no mockup
        c.setStrokeColor(ORANGE)
        c.setLineWidth(1.25)
        c.line(cx - 4.0*mm, 1.7*mm, cx + 4.0*mm, 1.7*mm)

    def draw(self):
        c = self.canv
        c.saveState()
        n = len(self.items)
        card_w = (self.width - self.gap * (n - 1)) / n
        for idx, (label, value, icon) in enumerate(self.items):
            self._draw_card(c, idx * (card_w + self.gap), card_w, label, value, icon, idx == self.emphasize_index)
        c.restoreState()


class _KpiSummaryStrip(Flowable):
    """Faixa de KPIs pós-cálculo, inspirada no modelo aprovado da página de diferenças."""

    def __init__(self, items, width: float, height: float = 24 * mm, emphasize_index: int | None = None, warning_index: int | None = None):
        super().__init__()
        self.items = list(items)
        self.width = width
        self.height = height
        self.emphasize_index = emphasize_index
        self.warning_index = warning_index

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def draw(self):
        c = self.canv
        c.saveState()
        h = self.height
        n = len(self.items)
        seg_w = self.width / n

        # sombra e container único: elimina qualquer sensação de desalinhamento entre cards
        c.setFillColor(HexColor('#E9EEF1'))
        c.roundRect(0.55*mm, -0.55*mm, self.width, h, 3.0*mm, fill=1, stroke=0)
        c.setFillColor(WHITE)
        c.setStrokeColor(HexColor('#C8D8DF'))
        c.setLineWidth(0.55)
        c.roundRect(0, 0, self.width, h, 3.0*mm, fill=1, stroke=1)

        for idx, (label, value, icon) in enumerate(self.items):
            x = idx * seg_w
            if idx == self.emphasize_index:
                c.setFillColor(BLUE_SOFT)
                c.rect(x + 0.4*mm, 0.4*mm, seg_w - 0.8*mm, h - 0.8*mm, fill=1, stroke=0)
            if idx == self.warning_index:
                c.setFillColor(ORANGE_SOFT)
                c.rect(x + 0.4*mm, 0.4*mm, seg_w - 0.8*mm, h - 0.8*mm, fill=1, stroke=0)
            if idx:
                c.setStrokeColor(LINE)
                c.setLineWidth(0.55)
                c.line(x, 3.0*mm, x, h - 3.0*mm)

            icx = x + 8.0*mm
            icy = h / 2
            _PremiumKpiCards._icon(c, icx, icy, icon)

            tx = x + 15.5*mm
            maxw = seg_w - 18.0*mm
            c.setFillColor(INK)
            c.setFont('Helvetica-Bold' if idx in {self.emphasize_index, self.warning_index} else 'Helvetica', 6.4)
            c.drawString(tx, h - 8.2*mm, label[:34])

            size = 11.4 if idx in {self.emphasize_index, self.warning_index} else 10.9
            while size > 7.0 and stringWidth(value, 'Helvetica-Bold', size) > maxw:
                size -= 0.3
            c.setFillColor(PETROLEUM_DEEP if idx != self.warning_index else WARNING)
            c.setFont('Helvetica-Bold', size)
            c.drawString(tx, 6.2*mm, value)

        c.restoreState()

class _SectionTitle(Flowable):
    def __init__(self, number: int, title: str, width: float = P_CONTENT, height: float = 7.4 * mm):
        super().__init__()
        self.number = number
        self.title = title
        self.width = width
        self.height = height

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(PETROLEUM_DEEP)
        c.roundRect(0, 1.0 * mm, 6.2 * mm, 5.6 * mm, 1.3 * mm, fill=1, stroke=0)
        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", 6.6)
        c.drawCentredString(3.1 * mm, 2.7 * mm, str(self.number))
        c.setFillColor(PETROLEUM_DEEP)
        c.setFont("Helvetica-Bold", 10.0)
        c.drawString(8.5 * mm, 2.4 * mm, self.title)
        c.setStrokeColor(PETROLEUM_DEEP)
        c.setLineWidth(0.45)
        c.line(8.5 * mm, 0.8 * mm, self.width, 0.8 * mm)
        c.restoreState()


class _PageCountCanvas(pdfcanvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            width, _ = self._pagesize
            margin = P_MARGIN if width < 700 else L_MARGIN
            self.setFillColor(PETROLEUM_DEEP)
            self.setFont("Helvetica", 6.4)
            self.drawRightString(width - margin, 7.2 * mm, f"Página {self._pageNumber} de {page_count}")
            super().showPage()
        super().save()



def _draw_header_footer(canvas, doc, *, pagesize, subtitle: str, logo_path: Path | None, identity: MAJSReportIdentity):
    width, height = pagesize
    landscape_mode = width > height
    margin = L_MARGIN if landscape_mode else P_MARGIN
    header_h = 20 * mm if landscape_mode else 22 * mm
    header_y = height - (28 * mm if landscape_mode else 34 * mm)
    header_x = margin
    header_w = width - 2 * margin

    canvas.saveState()
    radius = 4.8 * mm

    # Faixa azul com arredondamento apenas à esquerda; o lado direito termina reto.
    canvas.setFillColor(PETROLEUM_DEEP)
    canvas.roundRect(header_x, header_y, header_w, header_h, radius, fill=1, stroke=0)
    canvas.rect(header_x + header_w - radius, header_y, radius, header_h, fill=1, stroke=0)

    # Detalhe laranja vertical com largura intermediária, alinhado à extrema direita.
    accent_w = 4.2 * mm if landscape_mode else 4.6 * mm
    accent_x = header_x + header_w - accent_w
    canvas.setFillColor(ORANGE)
    canvas.rect(accent_x, header_y, accent_w, header_h, fill=1, stroke=0)

    logo_w = 28 * mm if landscape_mode else 30 * mm
    text_x = header_x + 8 * mm
    if logo_path and logo_path.exists():
        try:
            canvas.drawImage(
                str(logo_path), header_x + 5 * mm, header_y + 4.0 * mm,
                width=logo_w - 6 * mm, height=header_h - 8 * mm,
                preserveAspectRatio=True, mask="auto", anchor="c",
            )
            text_x = header_x + logo_w + 4 * mm
            canvas.setStrokeColor(HexColor("#7EA5B1"))
            canvas.setLineWidth(0.45)
            canvas.line(header_x + logo_w, header_y + 4.2 * mm, header_x + logo_w, header_y + header_h - 4.2 * mm)
        except Exception:
            pass
    else:
        canvas.setFillColor(WHITE)
        canvas.setFont(SANS_BOLD, 10)
        canvas.drawString(header_x + 6 * mm, header_y + 9 * mm, "FUNCEF")
        text_x = header_x + 33 * mm

    canvas.setFillColor(WHITE)
    canvas.setFont(SANS_BOLD, 13.0 if landscape_mode else 11.7)
    canvas.drawString(text_x, header_y + header_h - 8.2 * mm, "PARECER TÉCNICO - RECÁLCULO MAJS")
    canvas.setFont(SANS_BOLD, 7.6 if landscape_mode else 7.2)
    canvas.setFillColor(HexColor("#D8E8ED"))
    canvas.drawString(text_x, header_y + 6.2 * mm, subtitle)

    right = accent_x - 6.0 * mm
    canvas.setFillColor(WHITE)
    canvas.setFont(SANS, 6.5)
    canvas.drawRightString(right, header_y + header_h - 7.5 * mm, "Data-base")
    canvas.setFont(SANS_BOLD, 8.4)
    canvas.drawRightString(right, header_y + 6.0 * mm, _date(doc._majs_base_date))

    # LGPD / modo web: elaboração e validação não são renderizadas.
    canvas.setStrokeColor(LINE)
    canvas.setLineWidth(0.45)
    canvas.line(margin, 11 * mm, width - margin, 11 * mm)
    canvas.setFillColor(PETROLEUM_DEEP)
    canvas.setFont(SANS_BOLD, 6.2)
    canvas.drawString(margin, 7.2 * mm, "FUNCEF")
    canvas.setFont(SANS, 6.2)
    canvas.setFillColor(MUTED)
    canvas.drawString(margin + 12.3 * mm, 7.2 * mm, "- Fundação dos Economiários Federais")
    canvas.restoreState()


def _outstanding_balance(result: MAJSResult) -> Decimal:
    value = Decimal(str(result.balance_to_mature))
    return Decimal("0") if abs(value) < Decimal("0.005") else value


def _has_outstanding_balance(result: MAJSResult) -> bool:
    return _outstanding_balance(result) > Decimal("0")


def _outstanding_balance_label(result: MAJSResult) -> str:
    last = result.last_evolution_date
    if last and last < result.settings.base_date:
        return f"Saldo devedor remanescente em {_date(last)}"
    return "Saldo devedor a vencer na data-base"


def _kpi_cards(result: MAJSResult) -> Flowable:
    values = [
        ("Valor original", _money(result.settings.original_amount), "MONEY"),
        ("Total devido", _money(result.due_total), "WALLET"),
        ("Total pago/crédito", _money(result.paid_total), "PAID"),
        ("Diferença atualizada", _money(result.updated_difference_total), "TREND"),
        ("Data do crédito", _date(result.settings.credit_date), "DATE"),
    ]
    return _PremiumKpiCards(values, P_CONTENT, height=28 * mm, gap=2.0 * mm, emphasize_index=3)

def _info_table(result: MAJSResult, identity: MAJSReportIdentity, styles) -> Table:
    s = result.settings
    modality = "Fixa - sem correção monetária" if s.correction_index_code == SEM_CORRECAO else "Variável"
    annual = s.monthly_interest_rate * Decimal("12")
    values = []
    if identity.contract_number.strip():
        values.append(("Contrato", identity.contract_number.strip()))
    values.extend([
        ("Modalidade", modality),
        ("Prazo", f"{s.term_months} prestações"),
        ("Taxa nominal", _pct(annual, 6) + " a.a."),
    ])
    if identity.document_reference.strip():
        values.append(("Documento / processo", identity.document_reference.strip()))
    values.extend([
        ("Índice do saldo", "Sem correção" if s.correction_index_code == SEM_CORRECAO else s.correction_index_code),
        ("Defasagem", "Não se aplica" if s.correction_index_code == SEM_CORRECAO else f"{s.correction_lag_months} mês(es)"),
        ("Início dos juros de mora", _date(result.fixed_interest_start) if result.fixed_interest_start else "Data de origem de cada diferença"),
        ("Data-base", _date(s.base_date)),
        ("Quitação antecipada", "Sim" if result.was_settled_early else "Não"),
        ("Forma da quitação", ("Novação" if result.early_settlement and result.early_settlement.novation else "Pagamento pelo participante") if result.was_settled_early else "Não se aplica"),
    ])
    rows = []
    for i in range(0, len(values), 2):
        left = values[i]
        right = values[i + 1] if i + 1 < len(values) else ("", "")
        rows.append([
            Paragraph(escape(left[0]), styles["label"]), Paragraph(escape(str(left[1])), styles["value"]),
            Paragraph(escape(right[0]), styles["label"]), Paragraph(escape(str(right[1])), styles["value"]),
        ])
    t = Table(rows, colWidths=[31 * mm, 57 * mm, 35 * mm, P_CONTENT - 123 * mm])
    t.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.35, LINE),
        ("LINEAFTER", (1, 0), (1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3.7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.7),
    ]))
    return t


def _interest_formula(result: MAJSResult) -> str:
    mode = result.difference_interest_mode
    basis = majs_yearfrac_basis_label(result.yearfrac_basis)
    if mode == "UM_PCT_YEARFRAC":
        return f"12% a.a. em juros simples × fração de ano segundo a convenção {basis}"
    if mode == "UM_PCT_YEARFRAC_ATE_TL":
        return f"Até 29/08/2024: 12% a.a. em juros simples × fração de ano ({basis}); depois: Taxa Legal"
    if mode == "UM_PCT":
        return "1% a.m. × (dias corridos / 30)"
    if mode == "UM_PCT_ATE_TL":
        return "Até 29/08/2024: 1% a.m. × (dias / 30); depois: Taxa Legal"
    if mode == "TAXA_LEGAL":
        return "Taxa Legal oficial, em regime simples, a partir de 30/08/2024"
    return "Sem incidência de juros sobre as diferenças"


def _premises_table(result: MAJSResult, styles) -> Table:
    s = result.settings
    basis = majs_yearfrac_basis_label(result.yearfrac_basis) if "YEARFRAC" in result.difference_interest_mode else "Não se aplica"
    headers = ["Metodologia", "Índice do saldo", "Correção das diferenças", "Juros de mora", "Convenção de dias", "Arredondamento"]
    values = [
        "MAJS",
        "Sem correção" if s.correction_index_code == SEM_CORRECAO else s.correction_index_code,
        majs_correction_mode_label(result.difference_correction_mode),
        majs_interest_mode_label(result.difference_interest_mode),
        basis,
        "Valores: centavos | fatores: 6 casas",
    ]
    data = [
        [Paragraph(escape(h), styles["label"]) for h in headers],
        [Paragraph(escape(str(v)), styles["value"]) for v in values],
    ]
    widths = [27 * mm, 27 * mm, 37 * mm, 37 * mm, 28 * mm, P_CONTENT - 156 * mm]
    t = Table(data, colWidths=widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), PETROLEUM_SOFT),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3.6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.6),
    ]))
    formula_data = [[
        Paragraph("Critério de juros", styles["label"]),
        Paragraph(escape(_interest_formula(result)), styles["value"]),
    ]]
    formula = Table(formula_data, colWidths=[32 * mm, P_CONTENT - 32 * mm])
    formula.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.4, LINE),
        ("BACKGROUND", (0, 0), (0, 0), PETROLEUM_SOFT),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    wrapper = Table([[t], [formula]], colWidths=[P_CONTENT])
    wrapper.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5 * mm),
    ]))
    return wrapper


def _event_rows(result: MAJSResult):
    rows = []
    for r in result.schedule:
        if r.event_type == "SUSPENSAO":
            rows.append((_date(r.due_date), "Suspensão", "Competência suspensa", "-", r.event_note or "Sem prestação, juros ou amortização contratual."))
        elif r.event_type == "AMORTIZACAO_EXTRA":
            rows.append((_date(r.due_date), "Amortização extraordinária", "Redução extraordinária do saldo", _money(r.extra_amortization), r.event_note or "Correção proporcional aplicada antes da amortização."))
        elif r.event_type == "QUITACAO_ANTECIPADA":
            paid = _money(r.settlement_payment) if r.settlement_payment else ("Novação - sem desembolso" if result.early_settlement and result.early_settlement.novation else "R$ 0,00")
            rows.append((_date(r.due_date), "Quitação antecipada", f"Saldo MAJS na quitação: {_money(r.installment_due)}", paid, r.event_note or "Fluxo encerrado na data da quitação."))
        elif r.extra_amortization > 0:
            rows.append((_date(r.due_date), "Amortização extraordinária", "Amortização associada à competência", _money(r.extra_amortization), r.event_note or ""))
    return rows


def _events_table(result: MAJSResult, styles) -> Table:
    headers = ["Data", "Tipo de evento", "Descrição", "Valor", "Observações"]
    event_rows = _event_rows(result)
    data = [[Paragraph(h, styles["table_head"]) for h in headers]]
    if event_rows:
        for row in event_rows[:3]:
            data.append([
                Paragraph(escape(row[0]), styles["table_center"]),
                Paragraph(escape(row[1]), styles["table"]),
                Paragraph(escape(row[2]), styles["table"]),
                Paragraph(escape(row[3]), styles["table_right"]),
                Paragraph(escape(row[4]), styles["table"]),
            ])
        if len(event_rows) > 3:
            remaining = len(event_rows) - 3
            data.append([
                Paragraph("-", styles["table_center"]),
                Paragraph("Outros eventos", styles["table"]),
                Paragraph(f"+ {remaining} evento(s) detalhado(s) na evolução MAJS", styles["table"]),
                Paragraph("-", styles["table_right"]),
                Paragraph("Consulte a evolução integral nas páginas seguintes.", styles["table"]),
            ])
    else:
        data.append([Paragraph("Nenhum evento extraordinário cadastrado.", styles["table"]), "", "", "", ""])
    widths = [24 * mm, 39 * mm, 48 * mm, 25 * mm, P_CONTENT - 136 * mm]
    t = Table(data, colWidths=widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), PETROLEUM_DEEP),
        ("GRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3.5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.2),
    ]
    if not event_rows:
        style.append(("SPAN", (0, 1), (-1, 1)))
    t.setStyle(TableStyle(style))
    return t



def _financial_table(result: MAJSResult, styles) -> Table:
    nominal = result.difference_total
    corrected = sum((r.corrected_difference for r in result.differences), Decimal("0"))
    correction_amount = corrected - nominal
    interest_amount = sum((r.interest_amount for r in result.differences), Decimal("0"))
    balance = _outstanding_balance(result)

    rows = [
        ("Total devido até a data-base", result.due_total, "normal"),
        ("Total pago / crédito conciliado", result.paid_total, "normal"),
        ("Diferença nominal (pago - devido)", nominal, "normal"),
        ("Correção monetária das diferenças", correction_amount, "normal"),
        ("Juros sobre as diferenças", interest_amount, "normal"),
        ("Diferença atualizada", result.updated_difference_total, "difference"),
    ]
    if balance > 0:
        rows.append((_outstanding_balance_label(result), balance, "balance"))

    data = [[Paragraph("Descrição", styles["table_head"]), Paragraph("Valor", styles["table_head"])]]
    for label, value, _kind in rows:
        data.append([Paragraph(escape(label), styles["table"]), Paragraph(_money(value), styles["table_right"])])

    t = Table(data, colWidths=[P_CONTENT - 48 * mm, 48 * mm])
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), PETROLEUM_DEEP),
        ("GRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3.3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.3),
    ]

    for idx, (_label, _value, kind) in enumerate(rows, start=1):
        if kind == "difference":
            style.extend([
                ("BACKGROUND", (0, idx), (-1, idx), PETROLEUM_SOFT),
                ("FONTNAME", (0, idx), (-1, idx), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, idx), (-1, idx), PETROLEUM_DEEP),
            ])
        elif kind == "balance":
            style.extend([
                ("BACKGROUND", (0, idx), (-1, idx), ORANGE_SOFT),
                ("FONTNAME", (0, idx), (-1, idx), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, idx), (-1, idx), PETROLEUM_DEEP),
                ("LINEBEFORE", (0, idx), (0, idx), 2.2, ORANGE),
            ])

    t.setStyle(TableStyle(style))
    return t


def _conclusion_box(result: MAJSResult, identity: MAJSReportIdentity, styles) -> Table:
    value = result.updated_difference_total
    balance = _outstanding_balance(result)

    if value > 0:
        text = (
            f"Na data-base de {_date(result.settings.base_date)}, o recálculo apresenta diferença atualizada credora de "
            f"<b>{_money(value)}</b>, considerando a convenção pago menos devido."
        )
    elif value < 0:
        text = (
            f"Na data-base de {_date(result.settings.base_date)}, o recálculo apresenta diferença atualizada devedora de "
            f"<b>{_money(abs(value))}</b>, considerando a convenção pago menos devido."
        )
    else:
        text = f"Na data-base de {_date(result.settings.base_date)}, não foi apurada diferença financeira atualizada."

    if result.was_settled_early and result.early_settlement:
        settlement = result.early_settlement
        form = "novação" if settlement.novation else "pagamento pelo participante"
        text += (
            f"<br/><br/>O fluxo contratual foi encerrado antecipadamente em <b>{_date(settlement.settlement_date)}</b> por {form}. "
            f"O saldo MAJS teórico apurado nessa data foi de <b>{_money(result.early_settlement_balance)}</b>."
        )
        if settlement.novation:
            text += (
                " Por se tratar de novação, esse valor não foi considerado como desembolso do participante e não compôs a tabela de diferenças de quitação."
            )
        else:
            text += (
                f" O valor efetivamente pago foi de <b>{_money(settlement.amount_paid)}</b>, e a diferença entre o saldo teórico e esse pagamento foi incluída na apuração das diferenças."
            )
        text += " Não há saldo contratual a vencer após a data de quitação no fluxo MAJS."
    elif balance > 0:
        last = result.last_evolution_date
        if last and last < result.settings.base_date:
            balance_context = (
                f"O fluxo MAJS ainda apresenta saldo devedor remanescente de <b>{_money(balance)}</b> "
                f"na última posição efetivamente evoluída, em {_date(last)}; a data-base solicitada é "
                f"{_date(result.settings.base_date)}."
            )
        else:
            balance_context = (
                f"Na mesma data-base, o contrato mantém <b>saldo devedor a vencer de {_money(balance)}</b>."
            )
        text += (
            f"<br/><br/>{balance_context} Esse saldo é distinto das diferenças apuradas entre valores pagos e devidos; "
            "portanto, a diferença acima não deve ser interpretada, isoladamente, como quitação do contrato."
        )
    elif not result.was_settled_early:
        text += "<br/><br/>O fluxo MAJS não apresenta saldo devedor a vencer após a posição final demonstrada."

    if identity.observation.strip():
        text += f"<br/><br/><b>Observação do analista:</b> {escape(identity.observation.strip())}"
    critical_warnings = [
        w for w in result.warnings
        if any(token in w.lower() for token in (
            "cálculo parcial", "calculo parcial", "não disponível", "nao disponivel",
            "não possui", "nao possui", "não projet", "nao projet", "foi limitada",
            "superaram o saldo"
        ))
    ]
    if critical_warnings:
        text += "<br/><br/><b>Avisos relevantes:</b> " + " | ".join(escape(w) for w in critical_warnings[:2])

    t = Table([[Paragraph(text, styles["conclusion"])]], colWidths=[P_CONTENT])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BLUE_SOFT),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#C9DCE8")),
        ("LINEBEFORE", (0, 0), (0, -1), 3.0, PETROLEUM_DEEP),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
    ]))
    return t

def _signature_table(identity: MAJSReportIdentity, styles) -> Table | None:
    if not (identity.elaborator.strip() or identity.validator.strip()):
        return None
    data = [[
        Paragraph("Elaboração", styles["label"]), Paragraph(escape(identity.elaborator or "Não informado"), styles["value"]),
        Paragraph("Validação", styles["label"]), Paragraph(escape(identity.validator or "Não informado"), styles["value"]),
    ]]
    t = Table(data, colWidths=[25 * mm, 61 * mm, 25 * mm, P_CONTENT - 111 * mm])
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("BACKGROUND", (0, 0), (0, -1), PETROLEUM_SOFT),
        ("BACKGROUND", (2, 0), (2, -1), PETROLEUM_SOFT),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t



def _summary_story(result: MAJSResult, identity: MAJSReportIdentity, styles):
    story = []
    story.append(_kpi_cards(result))
    story.append(Spacer(1, 3.2 * mm))

    story.append(_SectionTitle(1, "Identificação e parâmetros"))
    story.append(_info_table(result, identity, styles))
    story.append(Spacer(1, 2.6 * mm))

    story.append(_SectionTitle(2, "Resumo da operação"))
    operation_text = (
        "O contrato foi recalculado pela metodologia MAJS, com reconstrução do fluxo teórico de amortização, "
        "consideração das suspensões e amortizações extraordinárias cadastradas e confronto dos valores teoricamente "
        "devidos com os pagamentos/créditos informados. As diferenças foram atualizadas até a data-base efetivamente "
        "alcançada pelo motor, sem projeção de índices indisponíveis."
    )
    if result.was_settled_early and result.early_settlement:
        form = "novação" if result.early_settlement.novation else "pagamento pelo participante"
        operation_text += (
            f" O fluxo contratual foi encerrado em {_date(result.early_settlement.settlement_date)} por {form}, "
            "sem evolução de prestações posteriores à quitação."
        )
    story.append(Paragraph(operation_text, styles["body"]))
    story.append(Spacer(1, 1.0 * mm))

    story.append(_SectionTitle(3, "Premissas aplicadas"))
    story.append(_premises_table(result, styles))
    story.append(Spacer(1, 2.0 * mm))

    story.append(_SectionTitle(4, "Eventos relevantes"))
    story.append(_events_table(result, styles))
    story.append(Spacer(1, 2.4 * mm))

    story.append(_SectionTitle(5, "Síntese financeira"))
    story.append(_financial_table(result, styles))
    story.append(Spacer(1, 2.4 * mm))

    story.append(_SectionTitle(6, "Conclusão técnica"))
    story.append(_conclusion_box(result, identity, styles))

    return story


def _evolution_table(result: MAJSResult, styles) -> Table:
    headers = [
        "Evento", "Prestação", "Data", "Saldo inicial", "Posição fator", "Fator financeiro", "Soma fatores",
        "Prestação devida / saldo quitação", "Juros", "Amortização", "Ref. índice", "Correção aplicada", "Amortização extra aplicada",
        "Valor pago na quitação", "Excesso da amortização", "Saldo final", "Observação",
    ]
    data = [[Paragraph(h, styles["tiny_head"]) for h in headers]]

    data.append([
        Paragraph("Contratação", styles["tiny"]),
        Paragraph("-", styles["tiny_center"]),
        Paragraph(_date(result.settings.credit_date), styles["tiny_center"]),
        Paragraph(_money(result.settings.original_amount), styles["tiny_right"]),
        *[Paragraph("-", styles["tiny_center"]) for _ in range(11)],
        Paragraph(_money(result.settings.original_amount), styles["tiny_right"]),
        Paragraph("Saldo originalmente contratado.", styles["tiny"]),
    ])

    event_labels = {
        "PRESTACAO": "Prestação",
        "SUSPENSAO": "Suspensão",
        "AMORTIZACAO_EXTRA": "Amortização extraordinária",
        "QUITACAO_ANTECIPADA": "Quitação antecipada",
    }
    event_row_numbers = []
    for r in result.schedule:
        event = event_labels.get(r.event_type, r.event_type)
        ref = r.correction_reference or "Sem correção"
        values = [
            Paragraph(escape(event), styles["tiny"]),
            Paragraph("-" if r.installment_number is None else str(r.installment_number), styles["tiny_center"]),
            Paragraph(_date(r.due_date), styles["tiny_center"]),
            Paragraph(_money(r.opening_balance), styles["tiny_right"]),
            Paragraph("-" if not r.factor_position else str(r.factor_position), styles["tiny_center"]),
            Paragraph("-" if not r.financial_factor else f"{r.financial_factor:.6f}".replace(".", ","), styles["tiny_right"]),
            Paragraph("-" if not r.remaining_factor_sum else f"{r.remaining_factor_sum:.6f}".replace(".", ","), styles["tiny_right"]),
            Paragraph("-" if r.installment_due == 0 else _money(r.installment_due), styles["tiny_right"]),
            Paragraph("-" if r.interest_component == 0 else _money(r.interest_component), styles["tiny_right"]),
            Paragraph("-" if r.amortization_component == 0 else _money(r.amortization_component), styles["tiny_right"]),
            Paragraph(escape(ref), styles["tiny_center"]),
            Paragraph(_pct(r.correction_rate, 4), styles["tiny_right"]),
            Paragraph("-" if r.extra_amortization == 0 else _money(r.extra_amortization), styles["tiny_right"]),
            Paragraph("-" if r.settlement_payment == 0 else _money(r.settlement_payment), styles["tiny_right"]),
            Paragraph("-" if r.excess_extra_amortization == 0 else _money(r.excess_extra_amortization), styles["tiny_right"]),
            Paragraph(_money(r.closing_balance), styles["tiny_right"]),
            Paragraph(escape(r.event_note or ("Prestação regular." if r.event_type == "PRESTACAO" else "")), styles["tiny"]),
        ]
        data.append(values)
        if r.event_type in {"SUSPENSAO", "AMORTIZACAO_EXTRA", "QUITACAO_ANTECIPADA"} or r.extra_amortization > 0:
            event_row_numbers.append(len(data) - 1)

    balance = _outstanding_balance(result)
    final_date = result.last_evolution_date or result.settings.credit_date
    if balance > 0:
        final_event = "SALDO A VENCER"
        if final_date < result.settings.base_date:
            final_note = (
                f"Saldo MAJS remanescente na última data efetivamente evoluída. "
                f"Data-base solicitada: {_date(result.settings.base_date)}."
            )
        else:
            final_note = (
                "O contrato permanece com saldo devedor a vencer após as prestações exigíveis até a data-base. "
                "Este saldo é distinto das diferenças pago x devido."
            )
    else:
        if result.was_settled_early and result.early_settlement:
            final_event = "CONTRATO QUITADO"
            final_note = (
                f"Fluxo encerrado antecipadamente em {_date(result.early_settlement.settlement_date)}. "
                "Não são geradas prestações posteriores à quitação."
            )
        else:
            final_event = "Posição final"
            final_note = "Fluxo MAJS sem saldo devedor a vencer na posição final."
            if final_date < result.settings.base_date:
                final_note += f" Data-base solicitada: {_date(result.settings.base_date)}."

    data.append([
        Paragraph(final_event, styles["tiny"]),
        Paragraph("-", styles["tiny_center"]),
        Paragraph(_date(final_date), styles["tiny_center"]),
        *[Paragraph("-", styles["tiny_center"]) for _ in range(12)],
        Paragraph(_money(balance), styles["tiny_right"]),
        Paragraph(final_note, styles["tiny"]),
    ])

    widths_mm = [17, 10, 16, 19, 11, 15, 15, 20, 15, 16, 14, 16, 18, 18, 18, 18, 24]
    widths = [x * mm for x in widths_mm]
    t = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), PETROLEUM_DEEP),
        ("GRID", (0, 0), (-1, -1), 0.25, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.3),
        ("TOPPADDING", (0, 0), (-1, -1), 2.3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.3),
    ]
    for row in range(1, len(data)):
        if row % 2 == 0:
            style.append(("BACKGROUND", (0, row), (-1, row), SOFT_ALT))
    for row in event_row_numbers:
        style.append(("BACKGROUND", (0, row), (-1, row), ORANGE_SOFT))

    final_row = len(data) - 1
    style.extend([
        ("BACKGROUND", (0, final_row), (-1, final_row), ORANGE_SOFT if balance > 0 else PETROLEUM_SOFT),
        ("FONTNAME", (0, final_row), (-1, final_row), "Helvetica-Bold"),
        ("TEXTCOLOR", (0, final_row), (-1, final_row), PETROLEUM_DEEP),
    ])
    if balance > 0:
        style.extend([
            ("LINEBEFORE", (0, final_row), (0, final_row), 2.4, ORANGE),
            ("BOX", (0, final_row), (-1, final_row), 0.65, ORANGE),
        ])
    t.setStyle(TableStyle(style))
    return t


def _outstanding_callout(result: MAJSResult, styles, width: float, *, context: str) -> Table | None:
    balance = _outstanding_balance(result)
    if balance <= 0:
        return None

    last = result.last_evolution_date
    if last and last < result.settings.base_date:
        lead = (
            f"<b>Saldo devedor remanescente: {_money(balance)}</b> na última posição efetivamente evoluída "
            f"({_date(last)}). Data-base solicitada: {_date(result.settings.base_date)}."
        )
    else:
        lead = f"<b>Saldo devedor a vencer na data-base: {_money(balance)}</b>."

    if context == "differences":
        detail = (
            " As diferenças abaixo medem somente o confronto entre valores pagos/créditos e valores devidos até a data-base. "
            "Elas não substituem nem absorvem o saldo contratual ainda a vencer."
        )
    else:
        detail = (
            " O destaque indica que o fluxo contratual MAJS não foi integralmente amortizado até a posição demonstrada."
        )

    t = Table([[Paragraph(lead + detail, styles["note"])]], colWidths=[width])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), ORANGE_SOFT),
        ("BOX", (0, 0), (-1, -1), 0.55, HexColor("#F2C49C")),
        ("LINEBEFORE", (0, 0), (0, -1), 3.2, ORANGE),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def _evolution_story(result: MAJSResult, styles):
    story = [
        Paragraph(
            "A seguir, apresenta-se a evolução integral da operação reconstruída pela metodologia MAJS, com os fatores "
            "financeiros, correção monetária, prestações e eventos que alteraram o fluxo teórico.",
            styles["note"],
        ),
        Spacer(1, 2.5 * mm),
        _evolution_table(result, styles),
        Spacer(1, 2.0 * mm),
    ]
    balance_note = _outstanding_callout(result, styles, L_CONTENT, context="evolution")
    if balance_note is not None:
        story.extend([balance_note, Spacer(1, 2.0 * mm)])
    story.append(Paragraph(
        "<b>Nota:</b> os valores são demonstrados em reais. Valores monetários são apresentados em centavos; fatores "
        "financeiros são exibidos com seis casas decimais. Em amortizações extraordinárias ocorridas entre vencimentos, "
        "a correção é apropriada proporcionalmente até a data do evento, a amortização é aplicada em seguida e o saldo "
        "reduzido recebe a parcela remanescente da correção até o próximo vencimento.",
        styles["body_small"],
    ))
    return story

def _correction_cell(row, styles) -> Paragraph:
    amount = row.corrected_difference - row.difference_paid_minus_due
    rate = row.correction_factor - Decimal("1")
    return Paragraph(
        f"{escape(_money(amount))} <font color='#62777D'>({escape(_pct(rate, 4))})</font>",
        styles["difference_right"],
    )


def _interest_cell(row, styles) -> Paragraph:
    return Paragraph(
        f"{escape(_money(row.interest_amount))} <font color='#62777D'>({escape(_pct(row.accumulated_interest_rate, 4))})</font>",
        styles["difference_right"],
    )


def _differences_table(result: MAJSResult, styles) -> Table:
    headers = ["Data", "Valor pago", "Valor devido", "Diferença", "Correção", "Juros", "Total"]
    data = [[Paragraph(h, styles["difference_head"]) for h in headers]]
    for r in result.differences:
        data.append([
            Paragraph(_date(r.due_date), styles["difference_center"]),
            Paragraph(_money(r.amount_paid), styles["difference_right"]),
            Paragraph(_money(r.amount_due), styles["difference_right"]),
            Paragraph(_money(r.difference_paid_minus_due), styles["difference_right"]),
            _correction_cell(r, styles),
            _interest_cell(r, styles),
            Paragraph(_money(r.updated_difference), styles["difference_right"]),
        ])
    widths = [26 * mm, 39 * mm, 39 * mm, 39 * mm, 50 * mm, 50 * mm, 42 * mm]
    t = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), PETROLEUM_DEEP),
        ("GRID", (0, 0), (-1, -1), 0.3, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]
    for row in range(1, len(data)):
        if row % 2 == 0:
            style.append(("BACKGROUND", (0, row), (-1, row), SOFT_ALT))
    t.setStyle(TableStyle(style))
    return t



def _difference_kpis(result: MAJSResult) -> Flowable:
    balance = _outstanding_balance(result)
    items = [
        ("Total pago", _money(result.paid_total), "PAID"),
        ("Total devido", _money(result.due_total), "MONEY"),
        ("Diferença nominal", _money(result.difference_total), "DELTA"),
        ("Diferença atualizada", _money(result.updated_difference_total), "TREND"),
    ]
    warning_index = None
    if balance > 0:
        items.append(("Saldo a vencer", _money(balance), "WALLET"))
        warning_index = len(items) - 1
    return _KpiSummaryStrip(
        items,
        L_CONTENT,
        height=24.5 * mm,
        emphasize_index=3,
        warning_index=warning_index,
    )

def _differences_story(result: MAJSResult, styles):
    correction_label = majs_correction_mode_label(result.difference_correction_mode)
    interest_label = majs_interest_mode_label(result.difference_interest_mode)
    callout = Table([[
        Paragraph(
            "Esta seção detalha as diferenças entre valores pagos/créditos e valores devidos até a data-base. "
            "Ela é exibida quando houver pagamentos/créditos informados ou quando uma quitação antecipada sem novação gerar diferença entre o saldo MAJS e o valor efetivamente pago.",
            styles["note"],
        )
    ]], colWidths=[L_CONTENT])
    callout.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BLUE_SOFT),
        ("BOX", (0, 0), (-1, -1), 0.45, HexColor("#B9D2E5")),
        ("LINEBEFORE", (0, 0), (0, -1), 3.0, PETROLEUM_DEEP),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    explanation = Paragraph(
        f"<b>Correção:</b> valor monetário + percentual acumulado. "
        f"<b>Juros:</b> valor monetário + percentual acumulado. "
        f"Critérios: correção - {escape(correction_label)}; juros - {escape(interest_label)}.",
        styles["body_small"],
    )

    story = [
        callout,
        Spacer(1, 2.0 * mm),
        explanation,
        Spacer(1, 2.0 * mm),
        _differences_table(result, styles),
        Spacer(1, 3.0 * mm),
        _difference_kpis(result),
        Spacer(1, 2.2 * mm),
    ]

    # O alerta do saldo vem DEPOIS dos cálculos e dos KPIs-resumo. Assim, o leitor
    # enxerga primeiro a memória e depois a síntese, sem confundir diferença com quitação.
    balance_note = _outstanding_callout(result, styles, L_CONTENT, context="differences")
    if balance_note is not None:
        story.extend([balance_note, Spacer(1, 1.0 * mm)])
    return story

def build_majs_opinion_pdf(
    result: MAJSResult,
    *,
    identity: MAJSReportIdentity | None = None,
    logo_path: str | Path | None = None,
) -> bytes:
    """
    Gera o parecer MAJS premium.

    - Havendo pagamentos/créditos: resumo executivo em A4 retrato, evolução MAJS integral em A4 paisagem
      e diferenças detalhadas em A4 paisagem.
    - Sem pagamentos/créditos: o PDF contém somente a evolução MAJS, em A4 paisagem, sem informações
      de valores pagos ou diferenças.
    """
    identity = identity or MAJSReportIdentity()
    styles = _styles()
    logo = Path(logo_path) if logo_path else None
    has_payments = _has_payments(result)
    out = BytesIO()

    first_size = PORTRAIT if has_payments else LANDSCAPE
    doc = BaseDocTemplate(
        out,
        pagesize=first_size,
        leftMargin=P_MARGIN if has_payments else L_MARGIN,
        rightMargin=P_MARGIN if has_payments else L_MARGIN,
        topMargin=39 * mm if has_payments else 31 * mm,
        bottomMargin=16 * mm,
        title="Parecer técnico - Recálculo MAJS",
        author="FUNCEF - Motor de Cálculos",
        subject="Recálculo por MAJS e diferenças",
    )
    doc._majs_base_date = result.settings.base_date

    summary_frame = Frame(
        P_MARGIN, 15 * mm, P_CONTENT, PORTRAIT[1] - 52 * mm,
        id="summary_frame", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )
    evolution_frame = Frame(
        L_MARGIN, 16 * mm, L_CONTENT, LANDSCAPE[1] - 48 * mm,
        id="evolution_frame", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )
    differences_frame = Frame(
        L_MARGIN, 16 * mm, L_CONTENT, LANDSCAPE[1] - 48 * mm,
        id="differences_frame", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )

    summary_template = PageTemplate(
        id="summary", pagesize=PORTRAIT, frames=[summary_frame],
        onPage=lambda c, d: _draw_header_footer(
            c, d, pagesize=PORTRAIT, subtitle="Resumo executivo da operação", logo_path=logo, identity=identity
        ),
    )
    evolution_template = PageTemplate(
        id="evolution", pagesize=LANDSCAPE, frames=[evolution_frame],
        onPage=lambda c, d: _draw_header_footer(
            c, d, pagesize=LANDSCAPE, subtitle="Evolução MAJS", logo_path=logo, identity=identity
        ),
    )
    differences_template = PageTemplate(
        id="differences", pagesize=LANDSCAPE, frames=[differences_frame],
        onPage=lambda c, d: _draw_header_footer(
            c, d, pagesize=LANDSCAPE, subtitle="Diferenças apuradas", logo_path=logo, identity=identity
        ),
    )

    if has_payments:
        doc.addPageTemplates([summary_template, evolution_template, differences_template])
        story = _summary_story(result, identity, styles)
        story.extend([NextPageTemplate("evolution"), PageBreak()])
        story.extend(_evolution_story(result, styles))
        story.extend([NextPageTemplate("differences"), PageBreak()])
        story.extend(_differences_story(result, styles))
    else:
        # Evolução é a primeira e única seção do parecer quando não há pagamentos/créditos.
        doc.addPageTemplates([evolution_template, summary_template, differences_template])
        story = _evolution_story(result, styles)

    doc.build(story, canvasmaker=_PageCountCanvas)
    return out.getvalue()
