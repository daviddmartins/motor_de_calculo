from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, localcontext
from io import BytesIO
from pathlib import Path
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable,
)

from .monetary_update import (
    MonetaryUpdateResult, UpdateRule, UpdateDailyRow,
    CORRECTION_NONE, CORRECTION_DAILY_BUSINESS, CORRECTION_DAILY_CALENDAR, CORRECTION_PRORATA_CALENDAR,
    CORRECTION_MONTHLY_CLOSE, CORRECTION_FACTOR_TABLE,
    INTEREST_NONE, INTEREST_SIMPLE_30, INTEREST_SIMPLE_COMPETENCE, INTEREST_SIMPLE_360_TJRJ, INTEREST_SIMPLE_YEARFRAC, INTEREST_COMPOUND_EQUIV,
    INTEREST_APPLICATION_DAILY, INTEREST_APPLICATION_END,
    INTEREST_BASE_PRINCIPAL, INTEREST_BASE_CORRECTED,
    ORDER_CORRECTION_INTEREST, ORDER_INTEREST_CORRECTION,
    PENALTY_BASE_PRINCIPAL, PENALTY_BASE_CORRECTED, PENALTY_BASE_BALANCE,
    tjrj_commercial_days_30_360, tjrj_simple_360_rate, excel_yearfrac, yearfrac_simple_rate,
)

from .pdf_theme import (
    PETROLEUM, PETROLEUM_DARK, PETROLEUM_DEEP, TEAL, ORANGE, GOLD, GOLD_DARK,
    INK, MUTED, LINE, SOFT, WHITE, SANS, SANS_BOLD, SERIF, SERIF_BOLD,
    draw_document_chrome, first_page_top_margin,
)

# Identidade visual compartilhada com todos os documentos do Motor de Cálculos.
NAVY = PETROLEUM_DARK
NAVY_DARK = PETROLEUM_DEEP
LINE_DARK = HexColor("#C9D9DE")
SOFT_BLUE = HexColor("#EEF5F7")

PAGE_MARGIN = 10.5 * mm
CONTENT_WIDTH = A4[0] - 2 * PAGE_MARGIN

CORRECTION_LABELS = {
    CORRECTION_NONE: "sem correção monetária",
    CORRECTION_DAILY_BUSINESS: "índice mensal convertido em taxa diária equivalente, com incidência apenas em dias úteis",
    CORRECTION_DAILY_CALENDAR: "índice mensal convertido em taxa diária equivalente, com incidência nos dias corridos da competência",
    CORRECTION_PRORATA_CALENDAR: "pró-rata mensal por dias corridos: fator equivalente pela fração de dias efetivamente abrangida em cada mês",
    CORRECTION_MONTHLY_CLOSE: "índice mensal integral aplicado no fechamento de cada competência",
    CORRECTION_FACTOR_TABLE: "tabela prática oficial, pela razão entre fatores mensais",
}
INTEREST_LABELS = {
    INTEREST_NONE: "sem juros",
    INTEREST_SIMPLE_30: "juros simples diários, com a taxa mensal dividida por 30",
    INTEREST_SIMPLE_COMPETENCE: "juros simples diários, distribuídos pelos dias corridos da competência",
    INTEREST_SIMPLE_360_TJRJ: "juros simples pelo critério TJRJ, com ano comercial de 360 dias",
    INTEREST_SIMPLE_YEARFRAC: "juros simples por fração de ano, conforme convenção de contagem de dias",
    INTEREST_COMPOUND_EQUIV: "taxa mensal convertida em taxa diária equivalente",
}
BASE_LABELS = {
    INTEREST_BASE_PRINCIPAL: "principal nominal",
    INTEREST_BASE_CORRECTED: "saldo corrigido",
}
ORDER_LABELS = {
    ORDER_CORRECTION_INTEREST: "correção monetária e, em seguida, juros",
    ORDER_INTEREST_CORRECTION: "juros e, em seguida, correção monetária",
}
INTEREST_APPLICATION_LABELS = {
    INTEREST_APPLICATION_END: "percentual acumulado e aplicado no fechamento",
    INTEREST_APPLICATION_DAILY: "juros aplicados diariamente",
}
YEARFRAC_BASIS_LABELS = {
    0: "US/NASD 30/360",
    1: "Real/Real",
    2: "Real/360",
    3: "Real/365",
    4: "Europeu 30/360",
}

PENALTY_BASE_LABELS = {
    PENALTY_BASE_PRINCIPAL: "principal nominal",
    PENALTY_BASE_CORRECTED: "principal corrigido",
    PENALTY_BASE_BALANCE: "saldo corrigido e acrescido de juros",
}


def _dec(v) -> Decimal:
    return v if isinstance(v, Decimal) else Decimal(str(v))


def _num_br(v, places: int = 2) -> str:
    text = f"{_dec(v):,.{places}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def _money(v, places: int = 2) -> str:
    n = _dec(v)
    return ("-R$ " if n < 0 else "R$ ") + _num_br(abs(n), places)


def _pct(rate: Decimal, places: int = 4) -> str:
    return _num_br(_dec(rate) * Decimal("100"), places) + "%"


def _factor(v: Decimal, places: int = 6) -> str:
    return _num_br(v, places)


def _int_br(v: int) -> str:
    return f"{int(v):,}".replace(",", ".")


def _date(v: date | None) -> str:
    return v.strftime("%d/%m/%Y") if v else "-"


def _month_key(v: date) -> str:
    return f"{v.year:04d}-{v.month:02d}"


def _previous_month_key(key: str) -> str:
    y, m = map(int, key.split("-"))
    return f"{y-1:04d}-12" if m == 1 else f"{y:04d}-{m-1:02d}"


def _styles():
    ss = getSampleStyleSheet()
    return {
        "kicker": ParagraphStyle("mu_kicker", parent=ss["BodyText"], fontName=SANS_BOLD, fontSize=7.7, leading=9, textColor=GOLD, spaceAfter=2),
        "title": ParagraphStyle("mu_title", parent=ss["Title"], fontName=SERIF_BOLD, fontSize=19.0, leading=21.0, textColor=NAVY_DARK, alignment=TA_LEFT, spaceAfter=2),
        "subtitle": ParagraphStyle("mu_subtitle", parent=ss["BodyText"], fontName=SANS, fontSize=8.3, leading=10, textColor=MUTED),
        "body": ParagraphStyle("mu_body", parent=ss["BodyText"], fontName=SANS, fontSize=7.6, leading=10.3, textColor=INK, spaceAfter=4),
        "body_tight": ParagraphStyle("mu_body_t", parent=ss["BodyText"], fontName=SANS, fontSize=7.0, leading=9.4, textColor=INK, spaceAfter=2),
        "small": ParagraphStyle("mu_small", parent=ss["BodyText"], fontName=SANS, fontSize=6.2, leading=8.2, textColor=MUTED),
        "meta_label": ParagraphStyle("mu_ml", parent=ss["BodyText"], fontName=SANS, fontSize=5.8, leading=6.8, textColor=INK),
        "meta_value": ParagraphStyle("mu_mv", parent=ss["BodyText"], fontName=SANS_BOLD, fontSize=7.1, leading=8.5, textColor=NAVY_DARK),
        "stage_title": ParagraphStyle("mu_st", parent=ss["Heading3"], fontName=SERIF_BOLD, fontSize=12.1, leading=13.6, textColor=NAVY_DARK),
        "stage_text": ParagraphStyle("mu_sx", parent=ss["BodyText"], fontName=SANS, fontSize=6.9, leading=9.0, textColor=MUTED),
        "stage_value": ParagraphStyle("mu_sv", parent=ss["BodyText"], fontName=SERIF_BOLD, fontSize=14.5, leading=16.0, textColor=NAVY_DARK, alignment=TA_RIGHT),
        "stage_value_gold": ParagraphStyle("mu_svg", parent=ss["BodyText"], fontName=SERIF_BOLD, fontSize=14.5, leading=16.0, textColor=GOLD, alignment=TA_RIGHT),
        "subtotal_label": ParagraphStyle("mu_sbl", parent=ss["BodyText"], fontName=SANS_BOLD, fontSize=6.7, leading=8, textColor=NAVY_DARK),
        "subtotal_value": ParagraphStyle("mu_sbv", parent=ss["BodyText"], fontName=SERIF_BOLD, fontSize=11.7, leading=13.4, textColor=NAVY_DARK, alignment=TA_RIGHT),
        "total_label": ParagraphStyle("mu_tl", parent=ss["BodyText"], fontName=SANS_BOLD, fontSize=8.8, leading=10, textColor=WHITE),
        "total_value": ParagraphStyle("mu_tv", parent=ss["BodyText"], fontName=SERIF_BOLD, fontSize=18.8, leading=20.4, textColor=WHITE, alignment=TA_RIGHT),
        "section": ParagraphStyle("mu_sec", parent=ss["Heading2"], fontName=SERIF_BOLD, fontSize=10.7, leading=12.5, textColor=NAVY_DARK, spaceAfter=4),
        "table_head": ParagraphStyle("mu_th", parent=ss["BodyText"], fontName=SANS_BOLD, fontSize=5.6, leading=6.6, textColor=WHITE, alignment=TA_CENTER),
        "table": ParagraphStyle("mu_td", parent=ss["BodyText"], fontName=SANS, fontSize=5.7, leading=7.0, textColor=INK),
        "table_center": ParagraphStyle("mu_tdc", parent=ss["BodyText"], fontName=SANS, fontSize=5.7, leading=7.0, textColor=INK, alignment=TA_CENTER),
        "table_right": ParagraphStyle("mu_tdr", parent=ss["BodyText"], fontName=SANS, fontSize=5.7, leading=7.0, textColor=INK, alignment=TA_RIGHT),
        "appendix_title": ParagraphStyle("mu_at", parent=ss["Title"], fontName=SERIF_BOLD, fontSize=17.0, leading=20, textColor=NAVY_DARK, spaceAfter=4),
        "appendix_stage": ParagraphStyle("mu_as", parent=ss["Heading3"], fontName=SANS_BOLD, fontSize=8.4, leading=10, textColor=NAVY_DARK, spaceAfter=2),
        "calc_label": ParagraphStyle("mu_cl", parent=ss["BodyText"], fontName=SANS_BOLD, fontSize=6.3, leading=7.8, textColor=MUTED),
        "calc_value": ParagraphStyle("mu_cv", parent=ss["BodyText"], fontName=SANS_BOLD, fontSize=7.0, leading=8.4, textColor=INK, alignment=TA_RIGHT),
    }


class _MiniIcon(Flowable):
    def __init__(self, kind: str, size: float = 8 * mm, stroke_color=NAVY_DARK):
        super().__init__()
        self.kind = kind
        self.width = size
        self.height = size
        self.stroke_color = stroke_color

    def draw(self):
        c = self.canv
        w, h = self.width, self.height
        c.saveState()
        c.setStrokeColor(self.stroke_color)
        c.setFillColor(self.stroke_color)
        c.setLineWidth(0.8)
        if self.kind == "document":
            c.rect(w*.20, h*.13, w*.54, h*.70, fill=0, stroke=1)
            c.line(w*.57, h*.83, w*.74, h*.66)
            c.line(w*.57, h*.83, w*.57, h*.66)
            c.line(w*.57, h*.66, w*.74, h*.66)
            for yy in (.53, .40, .27): c.line(w*.30, h*yy, w*.64, h*yy)
        elif self.kind == "person":
            c.circle(w*.50, h*.64, w*.13, fill=0, stroke=1)
            c.arc(w*.23, h*.08, w*.77, h*.49, 0, 180)
        elif self.kind == "calendar":
            c.rect(w*.18, h*.18, w*.64, h*.58, fill=0, stroke=1)
            c.line(w*.18, h*.59, w*.82, h*.59)
            c.line(w*.34, h*.84, w*.34, h*.68); c.line(w*.66, h*.84, w*.66, h*.68)
            for xx in (.33,.50,.67): c.circle(w*xx, h*.43, w*.025, fill=1, stroke=0)
            for xx in (.33,.50,.67): c.circle(w*xx, h*.30, w*.025, fill=1, stroke=0)
        elif self.kind == "chart":
            c.line(w*.18, h*.18, w*.18, h*.80); c.line(w*.18, h*.18, w*.84, h*.18)
            c.rect(w*.29, h*.18, w*.10, h*.25, fill=0, stroke=1)
            c.rect(w*.48, h*.18, w*.10, h*.42, fill=0, stroke=1)
            c.rect(w*.67, h*.18, w*.10, h*.58, fill=0, stroke=1)
        elif self.kind == "clock":
            c.circle(w*.50, h*.50, w*.31, fill=0, stroke=1)
            c.line(w*.50, h*.50, w*.50, h*.70); c.line(w*.50, h*.50, w*.66, h*.42)
        elif self.kind == "scale":
            c.line(w*.50, h*.18, w*.50, h*.78); c.line(w*.25, h*.66, w*.75, h*.66)
            c.line(w*.33, h*.66, w*.22, h*.42); c.line(w*.67, h*.66, w*.78, h*.42)
            c.arc(w*.10,h*.25,w*.34,h*.47,180,180); c.arc(w*.66,h*.25,w*.90,h*.47,180,180)
            c.line(w*.32,h*.18,w*.68,h*.18)
        elif self.kind == "shield":
            p = c.beginPath(); p.moveTo(w*.50,h*.86); p.lineTo(w*.78,h*.74); p.lineTo(w*.72,h*.35); p.lineTo(w*.50,h*.12); p.lineTo(w*.28,h*.35); p.lineTo(w*.22,h*.74); p.close()
            c.drawPath(p, fill=0, stroke=1)
            c.line(w*.40,h*.48,w*.48,h*.39); c.line(w*.48,h*.39,w*.64,h*.58)
        c.restoreState()


class _Connector(Flowable):
    def __init__(self, text: str, width: float = CONTENT_WIDTH, height: float = 8 * mm):
        super().__init__(); self.width = width; self.height = height; self.text = text

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFont(SANS, 6.6)
        tw = pdfmetrics.stringWidth(self.text, SANS, 6.6)
        center = self.width / 2
        gap = min(tw + 14*mm, self.width * .72)
        y = self.height * .55
        c.setStrokeColor(LINE_DARK); c.setLineWidth(.45)
        c.line(0, y, center-gap/2, y); c.line(center+gap/2, y, self.width, y)
        c.setFillColor(GOLD); c.setStrokeColor(GOLD); c.setLineWidth(.7)
        ax = center-gap/2 + 4*mm
        c.line(ax, y+2.2*mm, ax, y-1.0*mm)
        c.line(ax, y-1.0*mm, ax-1.4*mm, y+0.5*mm); c.line(ax, y-1.0*mm, ax+1.4*mm, y+0.5*mm)
        c.drawCentredString(center+2.0*mm, y-2.1, self.text)
        c.restoreState()


class _BottomRule(Flowable):
    def __init__(self, width: float = CONTENT_WIDTH, height: float = 2.1*mm):
        super().__init__(); self.width=width; self.height=height
    def draw(self):
        self.canv.setStrokeColor(ORANGE); self.canv.setLineWidth(1.2); self.canv.line(0,self.height/2,self.width,self.height/2)


class _NoWrapRight(Flowable):
    """Texto monetário alinhado à direita que nunca quebra em duas linhas."""
    def __init__(self, text: str, *, font_name: str, font_size: float, color, min_font_size: float = 9.5):
        super().__init__()
        self.text = text
        self.font_name = font_name
        self.font_size = font_size
        self.min_font_size = min_font_size
        self.color = color
        self._fit_size = font_size
        self.width = 0
        self.height = font_size * 1.25

    def wrap(self, availWidth, availHeight):
        self.width = availWidth
        size = self.font_size
        while size > self.min_font_size and pdfmetrics.stringWidth(self.text, self.font_name, size) > availWidth:
            size -= 0.25
        self._fit_size = size
        self.height = max(size * 1.25, 11)
        return availWidth, self.height

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(self.color)
        c.setFont(self.font_name, self._fit_size)
        c.drawRightString(self.width, self._fit_size * 0.20, self.text)
        c.restoreState()


def _footer(canvas, doc):
    canvas.saveState()
    width, _ = A4
    y = 7.8 * mm
    canvas.setStrokeColor(LINE_DARK); canvas.setLineWidth(.45)
    canvas.line(PAGE_MARGIN + 72*mm, y+1.3*mm, width-PAGE_MARGIN-18*mm, y+1.3*mm)
    canvas.setFillColor(MUTED); canvas.setFont(SANS, 5.5)
    canvas.drawString(PAGE_MARGIN+9*mm, y, "Relatório emitido pelo MOTOR DE CÁLCULOS · FUNCEF")
    canvas.drawRightString(width-PAGE_MARGIN, y, f"Página {doc.page}")
    # pequeno escudo vetorial
    c = canvas; c.setStrokeColor(HexColor("#597B8A")); c.setLineWidth(.6)
    x = PAGE_MARGIN + 2.7*mm
    p=c.beginPath(); p.moveTo(x,y+5.3*mm); p.lineTo(x+3.1*mm,y+4.0*mm); p.lineTo(x+2.5*mm,y+0.7*mm); p.lineTo(x,y-1.0*mm); p.lineTo(x-2.5*mm,y+0.7*mm); p.lineTo(x-3.1*mm,y+4.0*mm); p.close(); c.drawPath(p,fill=0,stroke=1)
    c.line(x-1.1*mm,y+1.8*mm,x-.2*mm,y+1.0*mm); c.line(x-.2*mm,y+1.0*mm,x+1.5*mm,y+3.0*mm)
    canvas.restoreState()



class _PremiumKpiCards(Flowable):
    """KPIs alinhados ao padrão visual aprovado no parecer MAJS."""
    def __init__(self, items, width: float, height: float = 27.5 * mm, gap: float = 2.0 * mm, emphasize_index: int | None = None):
        super().__init__(); self.items=list(items); self.width=width; self.height=height; self.gap=gap; self.emphasize_index=emphasize_index
    def wrap(self, availWidth, availHeight): return self.width, self.height
    def draw(self):
        c=self.canv; c.saveState(); n=len(self.items); card_w=(self.width-self.gap*(n-1))/n
        for idx,(label,value,icon_kind) in enumerate(self.items):
            x=idx*(card_w+self.gap); emph=idx==self.emphasize_index; h=self.height
            c.setFillColor(HexColor('#E8EEF2')); c.roundRect(x+.45*mm,-.45*mm,card_w,h,3*mm,fill=1,stroke=0)
            c.setFillColor(HexColor('#F1F6FB') if emph else WHITE); c.setStrokeColor(HexColor('#C8D8DF')); c.setLineWidth(.55); c.roundRect(x,0,card_w,h,3*mm,fill=1,stroke=1)
            cx=x+card_w/2; cy=h-8.1*mm
            c.setFillColor(HexColor('#EEF3F7')); c.circle(cx,cy,4.8*mm,fill=1,stroke=0)
            c.setStrokeColor(PETROLEUM_DEEP); c.setFillColor(PETROLEUM_DEEP); c.setLineWidth(.9)
            if icon_kind=='money':
                c.circle(cx,cy,3.1*mm,fill=0,stroke=1); c.setFont(SANS_BOLD,5.9); c.drawCentredString(cx,cy-1.7,'R$')
            elif icon_kind=='trend':
                c.line(cx-3.1*mm,cy-1.7*mm,cx-1.0*mm,cy+.2*mm); c.line(cx-1.0*mm,cy+.2*mm,cx+.7*mm,cy-1.0*mm); c.line(cx+.7*mm,cy-1.0*mm,cx+3.1*mm,cy+2.0*mm)
            elif icon_kind=='clock':
                c.circle(cx,cy,3.0*mm,fill=0,stroke=1); c.line(cx,cy,cx,cy+1.8*mm); c.line(cx,cy,cx+1.6*mm,cy-.8*mm)
            elif icon_kind=='shield':
                p=c.beginPath(); p.moveTo(cx,cy+3.2*mm); p.lineTo(cx+2.8*mm,cy+2.0*mm); p.lineTo(cx+2.2*mm,cy-1.7*mm); p.lineTo(cx,cy-3.2*mm); p.lineTo(cx-2.2*mm,cy-1.7*mm); p.lineTo(cx-2.8*mm,cy+2.0*mm); p.close(); c.drawPath(p,fill=0,stroke=1)
            else:
                c.setFont(SANS_BOLD,6); c.drawCentredString(cx,cy-1.8,'Δ')
            c.setFillColor(INK); c.setFont(SANS_BOLD if emph else SANS,6.1); c.drawCentredString(cx,9.0*mm,label[:34])
            size=11.8 if emph else 11.1
            while size>6.8 and pdfmetrics.stringWidth(value,SANS_BOLD,size)>card_w-4*mm: size-=.3
            c.setFillColor(PETROLEUM_DEEP); c.setFont(SANS_BOLD,size); c.drawCentredString(cx,4.0*mm,value)
            c.setStrokeColor(ORANGE); c.setLineWidth(1.2); c.line(cx-4*mm,1.6*mm,cx+4*mm,1.6*mm)
        c.restoreState()


class _SectionTitle(Flowable):
    def __init__(self, number: int, title: str, width: float = CONTENT_WIDTH, height: float = 7.4*mm):
        super().__init__(); self.number=number; self.title=title; self.width=width; self.height=height
    def wrap(self, availWidth, availHeight): return self.width,self.height
    def draw(self):
        c=self.canv; c.saveState(); c.setFillColor(PETROLEUM_DEEP); c.roundRect(0,1.0*mm,6.2*mm,5.6*mm,1.3*mm,fill=1,stroke=0)
        c.setFillColor(WHITE); c.setFont(SANS_BOLD,6.6); c.drawCentredString(3.1*mm,2.7*mm,str(self.number))
        c.setFillColor(PETROLEUM_DEEP); c.setFont(SANS_BOLD,10.0); c.drawString(8.5*mm,2.4*mm,self.title)
        c.setStrokeColor(PETROLEUM_DEEP); c.setLineWidth(.45); c.line(8.5*mm,.8*mm,self.width,.8*mm); c.restoreState()


def _partial_correction_notice(result: MonetaryUpdateResult, styles) -> Table:
    cutoff = result.correction_through_date if hasattr(result, "correction_through_date") else ((result.stopped_at - timedelta(days=1)) if result.stopped_at else result.base_date)
    missing = f"{result.missing_index_code} / {result.missing_index_reference}" if result.missing_index_code else "índice/fator ausente"
    interest_text = (
        f"Os juros selecionados foram apurados até a data-base de {_date(result.base_date)}, sobre a base disponível."
        if result.interest_total
        else "Não foram aplicados juros, pois a opção de juros não gerou valor no período."
    )
    p = Paragraph(
        f"<b>Cálculo parcial de correção monetária.</b> A correção foi apurada somente até <b>{_date(cutoff)}</b>, "
        f"porque {missing} não estava disponível. O Motor não projetou índice futuro; o saldo corrigido permaneceu congelado após esse marco. "
        + interest_text,
        ParagraphStyle('mu_partial_notice', parent=styles['body'], fontSize=7.5, leading=10.1, textColor=INK),
    )
    t = Table([[p]], colWidths=[CONTENT_WIDTH])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),HexColor('#FFF5EC')),
        ('BOX',(0,0),(-1,-1),.55,HexColor('#F7B77B')),
        ('LINEBEFORE',(0,0),(0,-1),3.2,ORANGE),
        ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
        ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
    ]))
    return t


def _financial_synthesis(result: MonetaryUpdateResult, styles) -> Table:
    rows=[('Valor original', result.original_total)]
    if result.correction_total or _unique_correction_labels(result): rows.append(('(+) Correção monetária', result.correction_total))
    if result.interest_total: rows.append(('(+) Juros', result.interest_total))
    if result.penalty_total: rows.append(('(+) Multa / acréscimos', result.penalty_total))
    if result.abatements_total: rows.append(('(-) Abatimentos', -result.abatements_total))
    rows.append((('(=) Saldo atualizado na data-base' if result.is_complete else '(=) Saldo apurado na data-base · correção parcial'), result.updated_total))
    data=[]
    for idx,(label,value) in enumerate(rows):
        final=idx==len(rows)-1
        label_style=ParagraphStyle(f'mu_syn_l_{idx}',parent=styles['body_tight'],fontName=SANS_BOLD if final else SANS,textColor=PETROLEUM_DEEP if final else INK)
        value_style=ParagraphStyle(f'mu_syn_v_{idx}',parent=styles['body_tight'],fontName=SANS_BOLD,fontSize=8.3 if final else 7.3,alignment=TA_RIGHT,textColor=PETROLEUM_DEEP)
        data.append([Paragraph(label,label_style),Paragraph(_money(value),value_style)])
    t=Table(data,colWidths=[CONTENT_WIDTH-48*mm,48*mm])
    cmds=[('GRID',(0,0),(-1,-1),.35,LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),4.2),('BOTTOMPADDING',(0,0),(-1,-1),4.2)]
    for r in range(len(rows)-1): cmds.append(('BACKGROUND',(0,r),(-1,r),WHITE if r%2==0 else SOFT))
    cmds += [('BACKGROUND',(0,-1),(-1,-1),HexColor('#EEF5F7')),('LINEBEFORE',(0,-1),(0,-1),3.0,ORANGE)]
    t.setStyle(TableStyle(cmds)); return t


def _conclusion_box(result: MonetaryUpdateResult, styles) -> Table:
    if result.is_complete:
        parts=[f"Na data-base de <b>{_date(result.base_date)}</b>, o saldo atualizado totaliza <b>{_money(result.updated_total)}</b>."]
    else:
        cutoff = result.correction_through_date if hasattr(result, "correction_through_date") else ((result.stopped_at - timedelta(days=1)) if result.stopped_at else result.base_date)
        parts=[
            f"Na data-base de <b>{_date(result.base_date)}</b>, o saldo apurado totaliza <b>{_money(result.updated_total)}</b>, "
            f"com <b>correção monetária parcial até {_date(cutoff)}</b>."
        ]
        if result.interest_total:
            parts.append("Os juros selecionados foram mantidos até a data-base sobre a base corrigida disponível, sem projeção do índice ausente.")
        parts.append("O valor deve ser novamente atualizado quando a competência faltante for incorporada à base oficial de índices.")
    composition=[f"valor original de {_money(result.original_total)}"]
    if result.correction_total: composition.append(f"correção monetária de {_money(result.correction_total)}")
    if result.interest_total: composition.append(f"juros de {_money(result.interest_total)}")
    if result.penalty_total: composition.append(f"multa/acréscimos de {_money(result.penalty_total)}")
    if result.abatements_total: composition.append(f"abatimentos de {_money(result.abatements_total)}")
    parts.append("A composição considera " + ", ".join(composition) + ".")
    p=Paragraph(" ".join(parts),ParagraphStyle('mu_conclusion',parent=styles['body'],fontSize=7.7,leading=10.3,textColor=INK))
    t=Table([[p]],colWidths=[CONTENT_WIDTH]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),HexColor('#F1F6FB')),('BOX',(0,0),(-1,-1),.5,HexColor('#C9DCE8')),('LINEBEFORE',(0,0),(0,-1),3.0,PETROLEUM_DEEP),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)])); return t

def _meta_cell(kind: str, label: str, value: str, styles) -> Table:
    inner = Table([[
        _MiniIcon(kind, 7.2*mm),
        Table([[Paragraph(label, styles["meta_label"])],[Paragraph(value or "-", styles["meta_value"])]], colWidths=[28*mm]),
    ]], colWidths=[8.5*mm, 28*mm])
    inner.setStyle(TableStyle([
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0),
    ]))
    return inner


def _meta_strip(result: MonetaryUpdateResult, styles, contract_number: str, participant_name: str, index_label: str, interest_start: date | None) -> Table:
    entries = []
    if contract_number.strip():
        entries.append(("document", "Contrato", contract_number.strip(), Decimal("0.18")))
    entries.extend([
        ("calendar", "Data-base do cálculo", _date(result.base_date), Decimal("0.20")),
        ("chart", "Índice / critério", index_label, Decimal("0.18")),
        ("clock", "Início dos juros", _date(interest_start) if interest_start else "Não aplicável", Decimal("0.17")),
    ])
    total_weight = sum((entry[3] for entry in entries), Decimal("0"))
    cells = [_meta_cell(kind, label, value, styles) for kind, label, value, _ in entries]
    widths = [CONTENT_WIDTH * float(weight / total_weight) for _, _, _, weight in entries]
    t = Table([cells], colWidths=widths, rowHeights=[17.0 * mm])
    cmds = [("BOX", (0, 0), (-1, -1), .55, LINE_DARK), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BACKGROUND", (0, 0), (-1, -1), WHITE), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]
    for i in range(1, len(cells)):
        cmds.append(("LINEBEFORE", (i, 0), (i, 0), .45, LINE))
    t.setStyle(TableStyle(cmds))
    return t


def _stage_block(number: int, title: str, details: list[str], amount: Decimal, styles, *, gold=False) -> Table:
    num_style = ParagraphStyle("stage_num", fontName=SERIF_BOLD, fontSize=15.5, leading=17, textColor=WHITE, alignment=TA_CENTER)
    num = Paragraph(f"{number:02d}", num_style)
    detail_html = "<br/>".join(details)
    value_col = 52 * mm
    middle_col = CONTENT_WIDTH - 17 * mm - value_col
    middle = Table([[Paragraph(title, styles["stage_title"])],[Paragraph(detail_html, styles["stage_text"])]], colWidths=[middle_col-5*mm])
    middle.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    value_text = ("+ " if amount>0 and number>1 else "") + _money(amount)
    value = _NoWrapRight(
        value_text, font_name=SERIF_BOLD, font_size=14.5,
        color=GOLD if gold else NAVY_DARK, min_font_size=10.5,
    )
    # Altura automática: evita cortar detalhes quando o bloco possui mais linhas,
    # e o valor monetário é desenhado em uma única linha.
    t=Table([[num,middle,value]], colWidths=[17*mm, middle_col, value_col])
    bg=GOLD if gold else NAVY
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(0,0),bg),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(0,0),"CENTER"),
        ("LEFTPADDING",(0,0),(0,0),0),("RIGHTPADDING",(0,0),(0,0),0),
        ("LEFTPADDING",(1,0),(1,0),7),("RIGHTPADDING",(1,0),(1,0),4),
        ("TOPPADDING",(1,0),(1,0),4),("BOTTOMPADDING",(1,0),(1,0),4),
        ("LEFTPADDING",(2,0),(2,0),3),("RIGHTPADDING",(2,0),(2,0),3),
        ("LINEBEFORE",(2,0),(2,0),.55,LINE_DARK),
    ])); return t


def _subtotal(label: str, value: Decimal, styles) -> Table:
    t=Table([[Paragraph(label,styles["subtotal_label"]),Paragraph(_money(value),styles["subtotal_value"])]], colWidths=[CONTENT_WIDTH-55*mm,55*mm], rowHeights=[9.2*mm])
    t.setStyle(TableStyle([("BOX",(0,0),(-1,-1),.55,LINE_DARK),("BACKGROUND",(0,0),(-1,-1),SOFT_BLUE),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5)])); return t


def _total_band(value: Decimal, styles) -> list:
    t=Table([[_MiniIcon("scale",10*mm, WHITE),Paragraph("VALOR TOTAL DA DÍVIDA NA DATA-BASE",styles["total_label"]),Paragraph(_money(value),styles["total_value"])]], colWidths=[19*mm,CONTENT_WIDTH-19*mm-70*mm,70*mm], rowHeights=[16.2*mm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),("LINEBEFORE",(2,0),(2,0),.6,HexColor("#89A3AE"))]))
    return [t,_BottomRule()]


def _summary_table(rows, styles, total_increment: Decimal, final_total: Decimal) -> Table:
    headers=["Etapa","Período / Base","Índice / Percentual","Acréscimo","Resultado"]
    data=[[Paragraph(h,styles["table_head"]) for h in headers]]
    for row in rows:
        stage_no, stage_name, period, indexp, amount, result, gold = row
        badge_style=ParagraphStyle("badge",fontName=SANS_BOLD,fontSize=5.8,leading=6.8,textColor=WHITE,alignment=TA_CENTER)
        badge=Table([[Paragraph(f"{stage_no:02d}",badge_style)]],colWidths=[5.5*mm],rowHeights=[5.5*mm])
        badge.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),GOLD if gold else NAVY),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
        first=Table([[badge,Paragraph(stage_name,styles["table"])]],colWidths=[7*mm,31*mm]); first.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),2),("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
        data.append([first,Paragraph(period,styles["table_center"]),Paragraph(indexp,styles["table_center"]),Paragraph(amount,styles["table_right"]),Paragraph(result,styles["table_right"])])
    total_style = ParagraphStyle("sum_total", fontName=SANS_BOLD, fontSize=5.9, leading=7.0, textColor=NAVY_DARK, alignment=TA_CENTER)
    data.append([
        Paragraph("TOTAL", total_style), Paragraph("—", total_style), Paragraph("—", total_style),
        Paragraph(("+ " if total_increment >= 0 else "") + _money(total_increment), ParagraphStyle("sum_total_r", parent=total_style, alignment=TA_RIGHT)),
        Paragraph(_money(final_total), ParagraphStyle("sum_total_r2", parent=total_style, alignment=TA_RIGHT)),
    ])
    t=Table(data,colWidths=[39*mm,49*mm,42*mm,31*mm,CONTENT_WIDTH-161*mm],repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("GRID",(0,0),(-1,-1),.35,LINE_DARK),("ROWBACKGROUNDS",(0,1),(-1,-2),[WHITE,SOFT]),("BACKGROUND",(0,-1),(-1,-1),SOFT_BLUE),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),3),("RIGHTPADDING",(0,0),(-1,-1),3),("TOPPADDING",(0,0),(-1,-1),3.4),("BOTTOMPADDING",(0,0),(-1,-1),3.4)])); return t


@dataclass
class _Stage:
    item_number: int
    rule: UpdateRule
    start: date
    end: date
    opening_total: Decimal
    correction_amount: Decimal
    interest_amount: Decimal
    penalty_amount: Decimal
    abatements: Decimal
    closing_total: Decimal
    correction_rate: Decimal
    # Soma das taxas diárias efetivamente aplicadas no trecho. Para juros simples,
    # este é o percentual nominal acumulado que reconcilia o valor quando usado
    # com a base média ponderada de incidência.
    interest_rate: Decimal
    first_interest_day: date | None
    last_interest_day: date | None
    interest_days: int
    weighted_interest_base: Decimal | None
    index_refs: list[str]
    correction_refs: list[str]


def _compound_rates(rates: Iterable[Decimal]) -> Decimal:
    with localcontext() as ctx:
        ctx.prec = 40
        factor = Decimal("1")
        for rate in rates: factor *= Decimal("1") + rate
        return factor - Decimal("1")


def _stage_rows(result: MonetaryUpdateResult, item_number: int) -> list[_Stage]:
    rule_by_order={r.order:r for r in result.rules}
    rows=[r for r in result.daily_rows if r.item_number==item_number and r.rule_order is not None]
    if not rows: return []
    stages=[]; current=None; bucket=[]
    def flush(values):
        if not values: return
        rule=rule_by_order.get(values[0].rule_order)
        if rule is None: return
        corr_rates=[x.correction_rate for x in values if x.correction_rate!=0]
        int_rows=[x for x in values if x.interest_rate!=0]
        int_rates=[x.interest_rate for x in int_rows]
        corr_rate=_compound_rates(corr_rates) if corr_rates else Decimal("0")
        # Nos métodos diários, a soma das taxas diárias reconcilia o percentual.
        # No TJRJ/360, o motor registra no fechamento da regra a taxa acumulada
        # comercial (taxa mensal × dias/30), aplicada ao saldo corrigido final.
        int_rate=(_compound_rates(int_rates) if rule.interest_method == INTEREST_COMPOUND_EQUIV else sum(int_rates,Decimal("0")))
        refs=[]; correction_refs=[]
        for x in values:
            if x.index_reference and x.index_reference not in refs: refs.append(x.index_reference)
            if x.correction_rate!=0 and x.index_reference and x.index_reference not in correction_refs: correction_refs.append(x.index_reference)
        first,last=values[0],values[-1]
        item_origin = next((i.origin_date for i in result.items if i.item_number == item_number), rule.start_date)
        if rule.interest_method in {INTEREST_SIMPLE_360_TJRJ, INTEREST_SIMPLE_YEARFRAC} and rule.monthly_interest_rate != 0:
            first_int = max(rule.interest_start_date or rule.start_date, item_origin)
            last_int = last.day
            if rule.interest_method == INTEREST_SIMPLE_360_TJRJ:
                int_days = tjrj_commercial_days_30_360(first_int, last_int)
                int_rate = tjrj_simple_360_rate(rule.monthly_interest_rate, first_int, last_int)
            else:
                int_days = (last_int - first_int).days
                int_rate = yearfrac_simple_rate(rule.monthly_interest_rate, first_int, last_int, rule.yearfrac_basis)
        elif rule.interest_application == INTEREST_APPLICATION_END and rule.interest_method in {INTEREST_SIMPLE_30, INTEREST_SIMPLE_COMPETENCE} and rule.monthly_interest_rate != 0:
            first_int = max(rule.interest_start_date or rule.start_date, item_origin + timedelta(days=1))
            last_int = last.day
            int_days = max(0, (last_int - first_int).days + 1) if last_int >= first_int else 0
            # O motor registra a taxa acumulada inteira no fechamento da regra.
            int_rate = sum(int_rates, Decimal("0"))
        else:
            first_int=int_rows[0].day if int_rows else None
            last_int=int_rows[-1].day if int_rows else None
            int_days=len({x.day for x in int_rows})
        interest_amount=sum((x.interest_amount for x in values),Decimal("0"))
        weighted_base=(interest_amount/int_rate) if int_rate else None
        stages.append(_Stage(item_number,rule,first.day,last.day,first.opening_principal+first.opening_correction+first.opening_interest+first.opening_penalty,sum((x.correction_amount for x in values),Decimal("0")),interest_amount,sum((x.penalty_amount for x in values),Decimal("0")),sum((x.abatement_amount for x in values),Decimal("0")),last.closing_total,corr_rate,int_rate,first_int,last_int,int_days,weighted_base,refs,correction_refs))
    for row in rows:
        if current is None or row.rule_order==current: bucket.append(row); current=row.rule_order
        else: flush(bucket); bucket=[row]; current=row.rule_order
    flush(bucket); return stages


def _factor_details(stage: _Stage, factor_series: dict[str, object] | None):
    if stage.rule.correction_method!=CORRECTION_FACTOR_TABLE or not factor_series:
        return None
    series=factor_series.get(stage.rule.correction_index_code)
    if series is None:
        return None
    values=getattr(series,"values",{})
    # A tabela prática representa o fator do mês de origem e o fator do mês final.
    # Usar diretamente as competências de início/fim evita que a memória textual
    # mostre um mês intermediário apenas porque a primeira variação ocorreu depois.
    initial_ref=_month_key(stage.start)
    final_ref=_month_key(stage.end)
    if initial_ref not in values:
        return None
    # No TJRJ, a data-base pode estar no mês imediatamente posterior ao último
    # fator publicado. O SCJ WEB usa o último fator mensal disponível; por isso
    # a memória deve mostrar a competência efetivamente utilizada, e não falhar
    # apenas porque a data final pertence ao mês corrente.
    effective_final_ref = final_ref
    if effective_final_ref not in values and stage.rule.correction_index_code == "TJRJ_CIVEL_14905":
        candidates=[ref for ref in values if ref <= final_ref]
        if candidates:
            effective_final_ref=max(candidates)
    if effective_final_ref not in values:
        return None
    fi,ff=_dec(values[initial_ref]),_dec(values[effective_final_ref])
    return initial_ref,fi,effective_final_ref,ff


def _unique_correction_labels(result: MonetaryUpdateResult) -> list[str]:
    vals=[]
    for r in result.rules:
        if r.correction_method!=CORRECTION_NONE and r.correction_index_code not in vals: vals.append(r.correction_index_code)
    return vals


def _index_display(result: MonetaryUpdateResult, profile_label: str) -> str:
    vals=_unique_correction_labels(result)
    if len(vals)==1: return vals[0].replace("_"," ")[:28]
    if len(vals)>1: return "Múltiplos critérios"
    return "Sem correção"


def _interest_start(result: MonetaryUpdateResult) -> date | None:
    dates=[]
    for r in result.rules:
        if r.interest_method!=INTEREST_NONE and r.monthly_interest_rate!=0:
            dates.append(r.interest_start_date or r.start_date)
    return min(dates) if dates else None


def _processed_end(result: MonetaryUpdateResult) -> date:
    """Última data efetivamente processada pelo motor."""
    if result.daily_rows:
        return max(x.day for x in result.daily_rows)
    return min(result.base_date, max(x.origin_date for x in result.items))


def _interest_metrics(result: MonetaryUpdateResult):
    """Métricas para demonstrar e reconciliar os juros do cálculo."""
    active_rules=[r for r in result.rules if r.interest_method!=INTEREST_NONE and r.monthly_interest_rate!=0]
    if not active_rules:
        return None

    # Caso simples TJRJ: a taxa acumulada não é uma soma de dias civis. O
    # percentual segue a convenção comercial 30/360 e é aplicado ao saldo
    # corrigido no término do período.
    if len(active_rules)==1 and active_rules[0].interest_method==INTEREST_SIMPLE_360_TJRJ:
        r=active_rules[0]
        origin=min(x.origin_date for x in result.items)
        first=max(r.interest_start_date or r.start_date, origin)
        last=min(r.end_date,_processed_end(result))
        days=tjrj_commercial_days_30_360(first,last)
        accumulated=tjrj_simple_360_rate(r.monthly_interest_rate,first,last)
        base=(result.interest_total/accumulated) if accumulated else None
        return {
            "first_day":first,"last_day":last,"days":days,
            "accumulated_rate":accumulated,"weighted_base":base,
            "tjrj_360":True,"yearfrac":False,"months_equivalent":Decimal(days)/Decimal("30"),
        }

    if len(active_rules)==1 and active_rules[0].interest_method==INTEREST_SIMPLE_YEARFRAC:
        r=active_rules[0]
        origin=min(x.origin_date for x in result.items)
        first=max(r.interest_start_date or r.start_date, origin)
        last=min(r.end_date,_processed_end(result))
        fraction=excel_yearfrac(first,last,r.yearfrac_basis)
        accumulated=r.monthly_interest_rate*Decimal("12")*fraction
        base=(result.interest_total/accumulated) if accumulated else None
        return {
            "first_day":first,"last_day":last,"days":max(0,(last-first).days),
            "accumulated_rate":accumulated,"weighted_base":base,
            "tjrj_360":False,"yearfrac":True,"year_fraction":fraction,
            "yearfrac_basis":r.yearfrac_basis,"months_equivalent":fraction*Decimal("12"),
        }

    rows=[x for x in result.daily_rows if x.interest_rate!=0]
    if not rows:
        return None
    rate_by_day_rule={}
    for x in rows:
        rate_by_day_rule.setdefault((x.day,x.rule_order),x.interest_rate)
    if len(active_rules)==1 and active_rules[0].interest_method==INTEREST_COMPOUND_EQUIV:
        accumulated=_compound_rates(rate_by_day_rule.values())
    else:
        accumulated=sum(rate_by_day_rule.values(),Decimal("0"))
    weighted_base=(result.interest_total/accumulated) if accumulated else None
    days=sorted({x.day for x in rows})
    applied_at_end = len(active_rules)==1 and active_rules[0].interest_application==INTEREST_APPLICATION_END
    if applied_at_end and active_rules[0].interest_method in {INTEREST_SIMPLE_30, INTEREST_SIMPLE_COMPETENCE}:
        r=active_rules[0]
        origin=min(x.origin_date for x in result.items)
        first=max(r.interest_start_date or r.start_date, origin + timedelta(days=1))
        last=min(r.end_date,_processed_end(result))
        day_count=max(0,(last-first).days+1) if last>=first else 0
        return {
            "first_day":first,"last_day":last,"days":day_count,
            "accumulated_rate":accumulated,"weighted_base":weighted_base,
            "tjrj_360":False,"yearfrac":False,"months_equivalent":None,"applied_at_end":True,
        }
    return {
        "first_day":days[0],"last_day":days[-1],"days":len(days),
        "accumulated_rate":accumulated,"weighted_base":weighted_base,
        "tjrj_360":False,"yearfrac":False,"months_equivalent":None,"applied_at_end":False,
    }


def _daily_interest_description(result: MonetaryUpdateResult) -> str:
    rules=[r for r in result.rules if r.interest_method!=INTEREST_NONE and r.monthly_interest_rate!=0]
    if not rules:
        return ""
    if len({(r.interest_method,r.monthly_interest_rate) for r in rules})!=1:
        return "Taxas diárias conforme cada regra aplicada"
    r=rules[0]
    if r.interest_method==INTEREST_SIMPLE_30:
        if r.interest_application==INTEREST_APPLICATION_END:
            return f"Taxa mensal / 30 acumulada ao longo do período; aplicação monetária única no fechamento"
        return f"Taxa diária: {_pct(r.monthly_interest_rate/Decimal('30'),6)} a.d. (taxa mensal / 30)"
    if r.interest_method==INTEREST_SIMPLE_COMPETENCE:
        if r.interest_application==INTEREST_APPLICATION_END:
            return "Taxas simples distribuídas pelos dias de cada competência, acumuladas e aplicadas uma única vez no fechamento"
        return "Taxa diária simples variável conforme a quantidade de dias da competência"
    if r.interest_method==INTEREST_SIMPLE_360_TJRJ:
        return "Convenção TJRJ: juros simples, ano comercial de 360 dias e mês de 30 dias"
    if r.interest_method==INTEREST_SIMPLE_YEARFRAC:
        return f"Fração de ano - convenção {YEARFRAC_BASIS_LABELS.get(r.yearfrac_basis, r.yearfrac_basis)}"
    if r.interest_method==INTEREST_COMPOUND_EQUIV:
        return "Taxa diária equivalente variável conforme a quantidade de dias da competência"
    return ""


def _interest_method_short(result: MonetaryUpdateResult) -> str:
    rules=[r for r in result.rules if r.interest_method!=INTEREST_NONE and r.monthly_interest_rate!=0]
    if not rules: return "Sem juros"
    rates={r.monthly_interest_rate for r in rules}; methods={r.interest_method for r in rules}
    if len(rates)==1 and len(methods)==1:
        r=rules[0]
        if r.interest_method==INTEREST_SIMPLE_360_TJRJ:
            annual=r.monthly_interest_rate*Decimal("12")
            return f"{_pct(annual)} a.a. ({_pct(r.monthly_interest_rate)} a.m.), simples - 360 dias/ano"
        if r.interest_method==INTEREST_SIMPLE_YEARFRAC:
            annual=r.monthly_interest_rate*Decimal("12")
            return f"{_pct(annual)} a.a. ({_pct(r.monthly_interest_rate)} a.m.), simples - fração de ano ({YEARFRAC_BASIS_LABELS.get(r.yearfrac_basis, r.yearfrac_basis)})"
        kind="simples" if r.interest_method in {INTEREST_SIMPLE_30,INTEREST_SIMPLE_COMPETENCE} else "equivalente composta"
        timing = "fechamento" if r.interest_application==INTEREST_APPLICATION_END else "aplicação diária"
        return f"{_pct(r.monthly_interest_rate)} a.m., {kind} - {timing}"
    return "Taxas conforme regras aplicadas"


def _interest_base_short(result: MonetaryUpdateResult) -> str:
    vals={r.interest_base for r in result.rules if r.interest_method!=INTEREST_NONE and r.monthly_interest_rate!=0}
    if len(vals)==1: return BASE_LABELS.get(next(iter(vals)),next(iter(vals)))
    return "bases conforme regras"


def _correction_rate_display(result: MonetaryUpdateResult) -> str:
    if len(result.items)==1:
        rates=[x.correction_rate for x in result.daily_rows if x.item_number==result.items[0].item_number and x.correction_rate!=0]
        if rates: return _pct(_compound_rates(rates))
    if result.original_total:
        return _pct(result.correction_total/result.original_total)
    return "-"


def _correction_period(result: MonetaryUpdateResult) -> str:
    origin=min(x.origin_date for x in result.items)
    end = result.correction_through_date if hasattr(result, "correction_through_date") else (result.base_date if result.is_complete else ((result.stopped_at - timedelta(days=1)) if result.stopped_at else _processed_end(result)))
    return f"{_date(origin)} a {_date(end)}"


def _interest_period(result: MonetaryUpdateResult) -> str:
    metrics=_interest_metrics(result)
    if not metrics:
        return "Não aplicável"
    return f"{_date(metrics['first_day'])} a {_date(metrics['last_day'])}"


def _penalty_summary(result: MonetaryUpdateResult):
    if not result.penalties: return None
    # Primeiro evento para descrição; valor total continua consolidado.
    pen=result.penalties[0]
    return pen, PENALTY_BASE_LABELS.get(pen.base,pen.base)


def _appendix_calc_box(rows, styles, accent=NAVY):
    data=[[Paragraph(a,styles["calc_label"]),Paragraph(b,styles["calc_value"])] for a,b in rows]
    t=Table(data,colWidths=[78*mm,CONTENT_WIDTH-78*mm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),SOFT),("LINEBEFORE",(0,0),(0,-1),2,accent),("GRID",(0,0),(-1,-1),.25,LINE),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3)])); return t


def _appendix_stage(stage: _Stage, styles, factor_series=None):
    r=stage.rule
    story=[Paragraph(f"Regra {r.order:02d} · {_date(stage.start)} a {_date(stage.end)} · {r.description or 'Critério de atualização'}",styles["appendix_stage"])]
    rows=[("Saldo no início do período",_money(stage.opening_total))]
    if r.correction_method!=CORRECTION_NONE:
        fi=_factor_details(stage,factor_series)
        if fi:
            ri,a,rf,b=fi; rows += [(f"Fator inicial ({ri})",_factor(a)),(f"Fator final ({rf})",_factor(b)),("Fator de atualização",_factor(b/a)),("Variação acumulada",_pct(b/a-Decimal('1')))]
        else:
            rows += [("Índice / tabela",r.correction_index_code),("Variação acumulada",_pct(stage.correction_rate))]
        rows.append(("(+) Correção monetária",_money(stage.correction_amount)))
    if r.interest_method!=INTEREST_NONE and r.monthly_interest_rate!=0:
        if r.interest_method==INTEREST_SIMPLE_360_TJRJ:
            rows += [
                ("Taxa de juros",f"{_pct(r.monthly_interest_rate*Decimal('12'))} a.a. ({_pct(r.monthly_interest_rate)} a.m.)"),
                ("Convenção","Juros simples - ano de 360 dias / mês de 30 dias"),
                ("Período dos juros",f"{_date(stage.first_interest_day)} a {_date(stage.last_interest_day)}"),
                ("Dias considerados",f"{_int_br(stage.interest_days)} dias 30/360"),
                ("Meses equivalentes",_num_br(Decimal(stage.interest_days)/Decimal('30'),4)),
                ("Percentual acumulado",_pct(stage.interest_rate)),
            ]
            if stage.weighted_interest_base is not None:
                rows.append(("Base de incidência",_money(stage.weighted_interest_base)))
        elif r.interest_method==INTEREST_SIMPLE_YEARFRAC:
            fraction=excel_yearfrac(stage.first_interest_day,stage.last_interest_day,r.yearfrac_basis) if stage.first_interest_day and stage.last_interest_day else Decimal('0')
            rows += [
                ("Taxa de juros",f"{_pct(r.monthly_interest_rate*Decimal('12'))} a.a. ({_pct(r.monthly_interest_rate)} a.m.)"),
                ("Método","Juros simples por fração de ano"),
                ("Convenção de contagem de dias",YEARFRAC_BASIS_LABELS.get(r.yearfrac_basis,str(r.yearfrac_basis))),
                ("Período dos juros",f"{_date(stage.first_interest_day)} a {_date(stage.last_interest_day)}"),
                ("Fração do ano",_num_br(fraction,6)),
                ("Meses equivalentes",_num_br(fraction*Decimal('12'),4)),
                ("Percentual acumulado",_pct(stage.interest_rate)),
            ]
            if stage.weighted_interest_base is not None:
                rows.append(("Base de incidência",_money(stage.weighted_interest_base)))
        else:
            rows += [
                ("Taxa de juros",f"{_pct(r.monthly_interest_rate)} a.m."),
                ("Marco inicial informado",_date(r.interest_start_date or r.start_date)),
                ("Período efetivamente calculado",f"{_date(stage.first_interest_day)} a {_date(stage.last_interest_day)} ({stage.interest_days} dia(s))"),
                ("Forma de aplicação",INTEREST_APPLICATION_LABELS.get(r.interest_application,r.interest_application)),
                ("Percentual acumulado",_pct(stage.interest_rate)),
            ]
            if stage.weighted_interest_base is not None:
                if r.interest_application==INTEREST_APPLICATION_END:
                    rows.append(("Base de incidência no fechamento",_money(stage.weighted_interest_base)))
                else:
                    rows.append(("Base média ponderada de incidência",_money(stage.weighted_interest_base)))
        rows.append(("(+) Juros",_money(stage.interest_amount)))
    if stage.penalty_amount: rows.append(("(+) Multa",_money(stage.penalty_amount)))
    if stage.abatements: rows.append(("(-) Abatimentos",_money(stage.abatements)))
    rows.append(("Saldo ao final do período",_money(stage.closing_total)))
    story += [_appendix_calc_box(rows,styles),Spacer(1,3*mm)]
    return story


def build_monetary_update_pdf(
    result: MonetaryUpdateResult,
    *,
    logo_path: Path | None = None,
    profile_label: str = "",
    factor_series: dict[str, object] | None = None,
    contract_number: str = "",
    participant_name: str = "",
    elaborator: str = "",
) -> bytes:
    """Gera demonstrativo premium alinhado ao design system do parecer MAJS."""
    styles=_styles(); out=BytesIO()
    doc=SimpleDocTemplate(
        out,pagesize=A4,leftMargin=PAGE_MARGIN,rightMargin=PAGE_MARGIN,
        topMargin=first_page_top_margin(),bottomMargin=16*mm,
        title="Demonstrativo de Atualização do Saldo Devedor", author="FUNCEF - Motor de Cálculos",
    )
    story=[]

    components=["correção monetária"]
    if result.interest_total: components.append("juros de mora")
    if result.penalty_total: components.append("multa")
    subtitle=(", ".join(components[:-1])+" e "+components[-1] if len(components)>1 else components[0])
    interest_start=_interest_start(result)
    idx_display=_index_display(result,profile_label)

    adjustment_label='Abatimentos' if result.abatements_total else 'Multa'
    adjustment_value=(-result.abatements_total) if result.abatements_total else result.penalty_total
    kpis=[
        ('Valor original',_money(result.original_total),'money'),
        ((f'Correção até {_date(result.correction_through_date)}' if not result.is_complete and hasattr(result, 'correction_through_date') else 'Correção monetária'),_money(result.correction_total),'trend'),
        ('Juros',_money(result.interest_total),'clock'),
        (adjustment_label,_money(adjustment_value),'shield'),
        (('Saldo atualizado' if result.is_complete else 'Saldo apurado'),_money(result.updated_total),'money'),
    ]
    story.append(_PremiumKpiCards(kpis,CONTENT_WIDTH,emphasize_index=4)); story.append(Spacer(1,4.0*mm))
    if not result.is_complete:
        story.append(_partial_correction_notice(result, styles)); story.append(Spacer(1,4.2*mm))
    else:
        story.append(Spacer(1,1.2*mm))

    story.append(_SectionTitle(1,'Identificação e critérios')); story.append(Spacer(1,1.5*mm))
    story.append(_meta_strip(result,styles,contract_number,participant_name,idx_display,interest_start)); story.append(Spacer(1,4.2*mm))

    story.append(_SectionTitle(2,'Síntese financeira')); story.append(Spacer(1,1.5*mm))
    story.append(_financial_synthesis(result,styles)); story.append(Spacer(1,4.2*mm))

    # Memória resumida em sequência matemática, preservando percentuais e bases.
    origins=[x.origin_date for x in result.items]; origin=min(origins)
    stage_no=1; summary=[]; running=result.original_total
    summary.append((stage_no,'Saldo devedor original',f'Ref. débito: {_date(origin)}','—','—',_money(running),False)); stage_no+=1
    if result.correction_total != 0 or _unique_correction_labels(result):
        running += result.correction_total
        summary.append((stage_no,'Correção monetária',_correction_period(result),f'{idx_display}<br/><b>{_correction_rate_display(result)}</b>',('+' if result.correction_total>=0 else '')+' '+_money(result.correction_total),_money(running),False)); stage_no+=1
    if result.interest_total:
        metrics=_interest_metrics(result)
        if metrics:
            period=f"{_date(metrics['first_day'])} a {_date(metrics['last_day'])}"
            pct=_pct(metrics['accumulated_rate'])
        else:
            period='Período conforme regras aplicadas'; pct='—'
        running += result.interest_total
        summary.append((stage_no,'Juros de mora',period,f'{_interest_method_short(result)}<br/><b>{pct}</b>','+ '+_money(result.interest_total),_money(running),False)); stage_no+=1
    if result.penalty_total:
        running += result.penalty_total
        summary.append((stage_no,'Multa / acréscimos',f'Data-base {_date(result.base_date)}','Conforme regra aplicada','+ '+_money(result.penalty_total),_money(running),False)); stage_no+=1
    if result.abatements_total:
        running -= result.abatements_total
        summary.append((stage_no,'Abatimentos','Eventos cadastrados','Redução proporcional','- '+_money(result.abatements_total),_money(running),False)); stage_no+=1

    story.append(_SectionTitle(3,'Demonstração do cálculo')); story.append(Spacer(1,1.5*mm))
    story.append(_summary_table(summary,styles,result.updated_total-result.original_total,result.updated_total)); story.append(Spacer(1,4.2*mm))

    story.append(_SectionTitle(4,'Conclusão técnica')); story.append(Spacer(1,1.2*mm)); story.append(_conclusion_box(result,styles))
    if elaborator.strip():
        story.append(Spacer(1,2.4*mm))
        story.append(Paragraph(f"Elaboração: {elaborator.strip()}",ParagraphStyle('mu_elab',parent=styles['small'],alignment=TA_CENTER,textColor=MUTED)))

    # Em cenários complexos, a memória completa continua em páginas posteriores.
    if len(result.rules)>1 or len(result.items)>1:
        story.append(PageBreak())
        story.append(Paragraph('Memória matemática detalhada',styles['appendix_title']))
        story.append(Paragraph('Os critérios são apresentados em sequência cronológica, preservando saldo de entrada, índices, taxas, valores gerados e saldo final de cada etapa.',styles['body']))
        for item in result.items:
            if len(result.items)>1:
                story.append(Paragraph(f"Item {item.item_number} · origem {_date(item.origin_date)} · {_money(item.original_amount)} · {item.description or 'valor informado'}",styles['section']))
            for stage in _stage_rows(result,item.item_number):
                story.extend(_appendix_stage(stage,styles,factor_series=factor_series))

    header_title='Demonstrativo de Atualização do Saldo Devedor'
    header_meta_parts=[]
    if profile_label: header_meta_parts.append(profile_label)
    header_meta_parts.append(f'Data-base {_date(result.base_date)}')
    if contract_number: header_meta_parts.append(f'Contrato {contract_number}')
    header_meta=' · '.join(header_meta_parts)

    def _first_page(canvas, doc):
        draw_document_chrome(canvas,doc,logo_path=logo_path,title=header_title,subtitle=subtitle.capitalize(),meta=header_meta,reference=contract_number or _date(result.base_date),compact=False)
    def _later_page(canvas, doc):
        draw_document_chrome(canvas,doc,logo_path=logo_path,title=header_title,subtitle=profile_label or 'Atualização monetária',meta=f'Data-base {_date(result.base_date)}',reference=contract_number or _date(result.base_date),compact=True)

    doc.build(story,onFirstPage=_first_page,onLaterPages=_later_page)
    return out.getvalue()


def _copy_stage_lines(stage: _Stage, factor_series=None) -> list[str]:
    r=stage.rule
    lines=[f"Regra {r.order} - {_date(stage.start)} a {_date(stage.end)}",f"Saldo inicial do período: {_money(stage.opening_total)}"]
    if r.correction_method!=CORRECTION_NONE:
        factor_info=_factor_details(stage,factor_series)
        if factor_info:
            ri,fi,rf,ff=factor_info
            lines += [f"Correção monetária: {r.correction_index_code}.",f"Fator inicial ({ri}): {_factor(fi)}; fator final ({rf}): {_factor(ff)}; fator acumulado: {_factor(ff/fi)} ({_pct(ff/fi-Decimal('1'))}).",f"Correção monetária apurada: {_money(stage.correction_amount)}."]
        else:
            lines += [f"Correção monetária: {r.correction_index_code}, variação acumulada {_pct(stage.correction_rate)}.",f"Correção monetária apurada: {_money(stage.correction_amount)}."]
    if r.interest_method!=INTEREST_NONE and r.monthly_interest_rate!=0:
        configured_start=r.interest_start_date or r.start_date
        if r.interest_method==INTEREST_SIMPLE_360_TJRJ:
            annual=r.monthly_interest_rate*Decimal("12")
            months=Decimal(stage.interest_days)/Decimal("30")
            lines.append(
                f"Juros de mora: {_pct(annual)} a.a. ({_pct(r.monthly_interest_rate)} a.m.), simples, "
                f"pela convenção TJRJ de 360 dias no ano. Período de {_date(stage.first_interest_day)} a {_date(stage.last_interest_day)}, "
                f"correspondente a {_int_br(stage.interest_days)} dias comerciais e {_num_br(months,4)} meses equivalentes."
            )
            lines.append(f"Percentual acumulado no período: {_pct(stage.interest_rate)}.")
            if stage.weighted_interest_base is not None:
                lines.append(
                    f"Base de incidência: {_money(stage.weighted_interest_base)} ({BASE_LABELS.get(r.interest_base,r.interest_base)}). "
                    f"Cálculo: {_money(stage.weighted_interest_base)} × {_pct(stage.interest_rate)} = {_money(stage.interest_amount)} de juros."
                )
            else:
                lines.append(f"Juros apurados: {_money(stage.interest_amount)}.")
        elif r.interest_method==INTEREST_SIMPLE_YEARFRAC:
            annual=r.monthly_interest_rate*Decimal("12")
            fraction=excel_yearfrac(stage.first_interest_day,stage.last_interest_day,r.yearfrac_basis) if stage.first_interest_day and stage.last_interest_day else Decimal('0')
            lines.append(
                f"Juros de mora: {_pct(annual)} a.a. ({_pct(r.monthly_interest_rate)} a.m.), simples, pelo critério de fração de ano, segundo a convenção de contagem de dias selecionada. "
                f"Período de {_date(stage.first_interest_day)} a {_date(stage.last_interest_day)}; "
                f"base {YEARFRAC_BASIS_LABELS.get(r.yearfrac_basis,r.yearfrac_basis)}; fração do ano {_num_br(fraction,6)} "
                f"({_num_br(fraction*Decimal('12'),4)} meses equivalentes)."
            )
            lines.append(f"Percentual acumulado no período: {_pct(stage.interest_rate)}.")
            if stage.weighted_interest_base is not None:
                lines.append(
                    f"Base de incidência: {_money(stage.weighted_interest_base)} ({BASE_LABELS.get(r.interest_base,r.interest_base)}). "
                    f"Cálculo: {_money(stage.weighted_interest_base)} × {_pct(stage.interest_rate)} = {_money(stage.interest_amount)} de juros."
                )
            else:
                lines.append(f"Juros apurados: {_money(stage.interest_amount)}.")
        else:
            lines.append(
                f"Juros: {_pct(r.monthly_interest_rate)} a.m.; marco inicial informado em {_date(configured_start)}; "
                f"período efetivamente calculado de {_date(stage.first_interest_day)} a {_date(stage.last_interest_day)} "
                f"({stage.interest_days} dia(s)); método {INTEREST_LABELS.get(r.interest_method,r.interest_method)}; "
                f"forma de aplicação {INTEREST_APPLICATION_LABELS.get(r.interest_application,r.interest_application)}; "
                f"base {BASE_LABELS.get(r.interest_base,r.interest_base)}."
            )
            lines.append(f"Percentual acumulado no período: {_pct(stage.interest_rate)}.")
            if stage.weighted_interest_base is not None:
                if r.interest_application==INTEREST_APPLICATION_END:
                    lines.append(
                        f"A correção monetária é concluída primeiro. Em seguida, o percentual acumulado é aplicado uma única vez "
                        f"sobre a base de {_money(stage.weighted_interest_base)}; cálculo: {_money(stage.weighted_interest_base)} × "
                        f"{_pct(stage.interest_rate)} = {_money(stage.interest_amount)} de juros."
                    )
                else:
                    lines.append(
                        f"Como os juros são reconhecidos diariamente, a base pode variar ao longo do período. A base média ponderada "
                        f"de incidência foi {_money(stage.weighted_interest_base)}; assim, {_money(stage.weighted_interest_base)} × "
                        f"{_pct(stage.interest_rate)} ≈ {_money(stage.interest_amount)} de juros."
                    )
            else:
                lines.append(f"Juros apurados: {_money(stage.interest_amount)}.")
    if stage.penalty_amount: lines.append(f"Multa/acréscimo no período: {_money(stage.penalty_amount)}.")
    if stage.abatements: lines.append(f"Abatimentos no período: {_money(stage.abatements)}.")
    lines.append(f"Saldo ao final do período: {_money(stage.closing_total)}."); return lines


def build_update_copy_text(result: MonetaryUpdateResult, *, profile_label: str="", detailed: bool=False, factor_series: dict[str,object]|None=None, contract_number: str="", participant_name: str="") -> str:
    origins=[x.origin_date for x in result.items]
    processed_end=_processed_end(result)
    title="ATUALIZAÇÃO MONETÁRIA DO SALDO DEVEDOR" if result.is_complete else "ATUALIZAÇÃO MONETÁRIA DO SALDO DEVEDOR - CÁLCULO PARCIAL"
    lines=[title,""]
    if contract_number: lines.append(f"Contrato: {contract_number}")
    if participant_name: lines.append(f"Participante: {participant_name}")
    if contract_number or participant_name: lines.append("")
    if result.is_complete:
        lines += [
            f"Na data de origem de {_date(min(origins))}, o valor total informado era de {_money(result.original_total)}. "
            f"A atualização foi processada até {_date(result.base_date)}, pelo critério {profile_label or 'personalizado'}.",
            f"A correção monetária apurada foi de {_money(result.correction_total)} e os juros totalizaram {_money(result.interest_total)}."
            +(f" A multa correspondeu a {_money(result.penalty_total)}." if result.penalty_total else "")
            +(f" Foram deduzidos abatimentos de {_money(result.abatements_total)}." if result.abatements_total else ""),
            f"O saldo atualizado na data-base é de {_money(result.updated_total)}.",
        ]
    else:
        cutoff = result.correction_through_date if hasattr(result, "correction_through_date") else ((result.stopped_at - timedelta(days=1)) if result.stopped_at else processed_end)
        lines += [
            f"Na data de origem de {_date(min(origins))}, o valor total informado era de {_money(result.original_total)}. "
            f"A correção monetária foi processada até {_date(cutoff)}, pelo critério {profile_label or 'personalizado'}; "
            f"a data-base solicitada é {_date(result.base_date)}.",
            f"A correção monetária apurada até o último índice disponível foi de {_money(result.correction_total)}. "
            f"Os juros totalizaram {_money(result.interest_total)}"
            +(f" e foram apurados até {_date(result.base_date)}, conforme a regra selecionada." if result.interest_total else ".")
            +(f" A multa correspondeu a {_money(result.penalty_total)}." if result.penalty_total else "")
            +(f" Foram deduzidos abatimentos de {_money(result.abatements_total)}." if result.abatements_total else ""),
            f"O saldo apurado na data-base é de {_money(result.updated_total)}, mas a correção monetária permanece PARCIAL "
            +(f"por ausência de {result.missing_index_code} para {result.missing_index_reference}. " if result.missing_index_code else "por ausência de índice/fator. ")
            +"Nenhum índice futuro foi projetado; o saldo corrigido foi mantido após o último índice disponível.",
        ]
    if detailed:
        lines += ["","PARÂMETROS APLICADOS","","DEMONSTRAÇÃO MATEMÁTICA"]
        for item in result.items:
            if len(result.items)>1: lines += ["",f"Item {item.item_number} - origem {_date(item.origin_date)} - valor inicial {_money(item.original_amount)}"]
            for stage in _stage_rows(result,item.item_number): lines += [""]+_copy_stage_lines(stage,factor_series=factor_series)
        lines += ["","RECONCILIAÇÃO FINAL",f"Valor original: {_money(result.original_total)}",f"(+) Correção monetária: {_money(result.correction_total)}",f"(+) Juros: {_money(result.interest_total)}"]
        if result.penalty_total: lines.append(f"(+) Multa: {_money(result.penalty_total)}")
        if result.abatements_total: lines.append(f"(-) Abatimentos: {_money(result.abatements_total)}")
        lines.append(f"(=) Saldo atualizado: {_money(result.updated_total)}")
    return "\n".join(lines)


def monetary_pdf_filename(base_date: date) -> str:
    return f"Atualizacao_Monetaria_{base_date.strftime('%Y%m%d')}.pdf"
