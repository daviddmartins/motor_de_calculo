from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from typing import Mapping, Sequence

from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
    KeepTogether, CondPageBreak, HRFlowable, Image, Flowable
)
from reportlab.pdfbase.pdfmetrics import stringWidth

from . import pdf_reports as base
from . import pdf_funcef_template as funcef

_BASE_BUILD = base.build_opinion_pdf


# ---------------------------------------------------------------------------
# Pareceres Credplan - padrão visual aprovado V13
# ---------------------------------------------------------------------------
# - mesmos cálculos do Motor de Cálculos;
# - layout aprovado em 17/08/2026;
# - KPIs executivos compactos;
# - cabeçalho no padrão MAJS com logo centralizada;
# - fórmulas renderizadas como equações matemáticas;
# - paginação por fluxo, sem PageBreak fixo.

NAVY = HexColor('#073D6D')
NAVY_DARK = HexColor('#073D6D')
BLUE = HexColor('#0B4A82')
BLUE_SOFT = HexColor('#EEF5FB')
BLUE_SOFT2 = HexColor('#F7FAFD')
LINE = HexColor('#C9D7E6')
TEXT = HexColor('#243748')
MUTED = HexColor('#667A8D')
ORANGE = HexColor('#F58220')
WHITE = colors.white

MARGIN = 16.5 * mm
CONTENT_W = A4[0] - 2 * MARGIN
TOP_MARGIN = funcef.TOP_MARGIN
BOTTOM_MARGIN = funcef.BOTTOM_MARGIN


def _money(v, places: int = 2) -> str:
    return base._money_br(v, places)


def _pct(v, places: int = 4) -> str:
    return base._percent_br(v, places)


def _date(v) -> str:
    return base._date_br(v)


def _competence_br(value: str | None) -> str:
    text = str(value or '').strip()
    if len(text) == 7 and text[4] == '-' and text[:4].isdigit() and text[5:].isdigit():
        return f'{text[5:7]}/{text[:4]}'
    return text or '-'


class _KPIGrid(Flowable):
    def __init__(self, items, width=CONTENT_W, height=24 * mm, gap=1.35 * mm):
        super().__init__()
        self.items = list(items)
        self.width = width
        self.height = height
        self.gap = gap

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def _icon(self, c, cx, cy, kind):
        c.setFillColor(HexColor('#F1F5F9'))
        c.circle(cx, cy, 4.0 * mm, fill=1, stroke=0)
        c.setStrokeColor(NAVY_DARK); c.setFillColor(NAVY_DARK); c.setLineWidth(0.85)
        k = kind.upper()
        if k == 'MOD':
            for off in (1.6, 0.0, -1.6):
                c.line(cx-2.6*mm, cy+off*mm, cx, cy+(off+1.1)*mm); c.line(cx, cy+(off+1.1)*mm, cx+2.6*mm, cy+off*mm)
                c.line(cx-2.6*mm, cy+off*mm, cx, cy+(off-1.1)*mm); c.line(cx, cy+(off-1.1)*mm, cx+2.6*mm, cy+off*mm)
        elif k == 'SYS':
            c.circle(cx, cy, 2.25*mm, fill=0, stroke=1)
            c.arc(cx-3.2*mm, cy-3.2*mm, cx+3.2*mm, cy+3.2*mm, 25, 125); c.arc(cx-3.2*mm, cy-3.2*mm, cx+3.2*mm, cy+3.2*mm, 205, 125)
        elif k == 'RATE':
            c.setFont('Helvetica-Bold', 8.0); c.drawCentredString(cx, cy-2.2, '%')
        elif k == 'TERM':
            c.circle(cx, cy, 2.7*mm, fill=0, stroke=1); c.line(cx, cy, cx, cy+1.9*mm); c.line(cx, cy, cx+1.5*mm, cy-0.8*mm)
        elif k == 'INDEX':
            for j, hh in enumerate((1.4, 2.2, 3.2, 4.2)):
                x = cx + (-2.7 + j*1.75) * mm; c.rect(x, cy-2.2*mm, 0.7*mm, hh*mm, fill=0, stroke=1)
        elif k == 'DATE':
            c.roundRect(cx-2.8*mm, cy-2.2*mm, 5.6*mm, 4.5*mm, 0.7*mm, fill=0, stroke=1)
            c.line(cx-2.8*mm, cy+0.7*mm, cx+2.8*mm, cy+0.7*mm)
            c.line(cx-1.5*mm, cy+2.8*mm, cx-1.5*mm, cy+1.6*mm); c.line(cx+1.5*mm, cy+2.8*mm, cx+1.5*mm, cy+1.6*mm)

    @staticmethod
    def _label_lines(label):
        words = label.split()
        if len(label) <= 16 or len(words) == 1: return [label]
        best = None
        for i in range(1, len(words)):
            a, b = ' '.join(words[:i]), ' '.join(words[i:])
            score = max(len(a), len(b)) + abs(len(a)-len(b))*0.25
            if best is None or score < best[0]: best = (score, [a, b])
        return best[1] if best else [label]

    def draw(self):
        c = self.canv; h = self.height; n = len(self.items); card_w = (self.width-self.gap*(n-1))/n
        kinds = ['MOD','SYS','RATE','TERM','INDEX','DATE']
        for i, (label, value) in enumerate(self.items):
            x = i*(card_w+self.gap)
            c.setFillColor(HexColor('#E8EEF3')); c.roundRect(x+0.45*mm,-0.45*mm,card_w,h,2.6*mm,fill=1,stroke=0)
            c.setFillColor(WHITE); c.setStrokeColor(HexColor('#C9D8E6')); c.setLineWidth(0.5); c.roundRect(x,0,card_w,h,2.6*mm,fill=1,stroke=1)
            cx = x+card_w/2; self._icon(c,cx,h-6.4*mm,kinds[i] if i<len(kinds) else 'MOD')
            lines = self._label_lines(label); c.setFillColor(TEXT); c.setFont('Helvetica',7.0)
            if len(lines)==1: c.drawCentredString(cx,10.4*mm,lines[0])
            else: c.drawCentredString(cx,11.9*mm,lines[0]); c.drawCentredString(cx,9.3*mm,lines[1])
            c.setStrokeColor(ORANGE); c.setLineWidth(1.05); c.line(cx-3.2*mm,7.5*mm,cx+3.2*mm,7.5*mm)
            size=8.2; maxw=card_w-3.0*mm
            while size>5.7 and stringWidth(value,'Helvetica-Bold',size)>maxw: size-=0.25
            c.setFillColor(NAVY_DARK); c.setFont('Helvetica-Bold',size); c.drawCentredString(cx,3.25*mm,value)


def _styles():
    sample = getSampleStyleSheet()
    body = funcef.BODY_SIZE
    table = funcef.TABLE_SIZE
    return {
        'body': ParagraphStyle('v13_body', parent=sample['BodyText'], fontName='Helvetica', fontSize=body, leading=funcef.BODY_LEADING, textColor=TEXT, alignment=TA_JUSTIFY, spaceAfter=0),
        'small': ParagraphStyle('v13_small', fontName='Helvetica', fontSize=funcef.NOTE_SIZE, leading=10.4, textColor=TEXT, alignment=TA_JUSTIFY),
        'label': ParagraphStyle('v13_label', fontName='Helvetica', fontSize=table, leading=funcef.TABLE_LEADING, textColor=NAVY_DARK),
        'value': ParagraphStyle('v13_value', fontName='Helvetica-Bold', fontSize=table, leading=funcef.TABLE_LEADING, textColor=TEXT),
        'section': ParagraphStyle('v13_section', fontName='Helvetica-Bold', fontSize=10.0, leading=11.5, textColor=NAVY_DARK, spaceAfter=0),
        'subsection': ParagraphStyle('v13_subsection', fontName='Helvetica-Bold', fontSize=funcef.SUBSECTION_SIZE, leading=11.6, textColor=NAVY_DARK, spaceBefore=1.2*mm, spaceAfter=1.6*mm, keepWithNext=1),
        'table_head': ParagraphStyle('v13_table_head', fontName='Helvetica-Bold', fontSize=table, leading=funcef.TABLE_LEADING, textColor=WHITE, alignment=TA_CENTER),
        'table': ParagraphStyle('v13_table', fontName='Helvetica', fontSize=table, leading=funcef.TABLE_LEADING, textColor=TEXT),
        'table_right': ParagraphStyle('v13_table_right', fontName='Helvetica', fontSize=table, leading=funcef.TABLE_LEADING, textColor=TEXT, alignment=TA_RIGHT),
        'formula_title': ParagraphStyle('v13_formula_title', fontName='Helvetica-Bold', fontSize=funcef.SUBSECTION_SIZE, leading=11.6, textColor=NAVY_DARK),
        'formula': ParagraphStyle('v13_formula', fontName='Helvetica', fontSize=11.5, leading=15.0, textColor=HexColor('#073F73'), alignment=TA_CENTER),
        'formula_note': ParagraphStyle('v13_formula_note', fontName='Helvetica', fontSize=funcef.NOTE_SIZE, leading=10.4, textColor=TEXT, alignment=TA_JUSTIFY),
        'formula_app': ParagraphStyle('v13_formula_app', fontName='Helvetica', fontSize=funcef.NOTE_SIZE, leading=10.4, textColor=TEXT, alignment=TA_JUSTIFY),
        'detail_title': ParagraphStyle('v13_detail_title', fontName='Helvetica-Bold', fontSize=funcef.SUBSECTION_SIZE, leading=11.6, textColor=NAVY_DARK, keepWithNext=1),
        'summary_value': ParagraphStyle('v13_summary_value', fontName='Helvetica-Bold', fontSize=table, leading=funcef.TABLE_LEADING, textColor=NAVY_DARK, alignment=TA_RIGHT),
    }


SECTION_NUMBERS = {
    'Objeto e parâmetros contratuais': '04',
    'Fundamentação técnico-financeira': '05',
    'Fórmulas e aplicação numérica': '06',
    'Demonstração da evolução': '07',
    'Conclusão técnica': '08',
}


def _section_title(text, st):
    return funcef.section_heading(SECTION_NUMBERS.get(text, ''), text, CONTENT_W)


def _param_table(identity, settings, result, st):
    fixed_payment = result.fixed_payment if result.fixed_payment is not None else None
    left=[('Data da solicitação',_date(identity.request_date)),('Data do crédito',_date(settings.credit_date)),('Margem consignável',_money(identity.margin_amount)),('Valor contratado',_money(settings.initial_balance)),('FGQC - concessão',_money(identity.fgqc_concession))]
    right=[('Saldo quitado na operação',_money(identity.settled_loan_balance)),('Valor líquido',_money(identity.net_amount)),('IOF',_money(identity.iof_amount)),('Taxa administrativa',_money(identity.administrative_fee)),('Valor da parcela base',_money(fixed_payment) if fixed_payment is not None else 'Não se aplica')]
    rows=[]
    for i in range(max(len(left),len(right))):
        l=left[i] if i<len(left) else ('',''); r=right[i] if i<len(right) else ('','')
        rows.append([Paragraph(l[0],st['label']),Paragraph(l[1],st['value']),Paragraph(r[0],st['label']),Paragraph(r[1],st['value'])])
    t=Table(rows,colWidths=[41.5*mm,47*mm,41.5*mm,47*mm],hAlign='LEFT')
    t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.35,LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('BACKGROUND',(0,0),(0,-1),BLUE_SOFT2),('BACKGROUND',(2,0),(2,-1),BLUE_SOFT2),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),4.5),('BOTTOMPADDING',(0,0),(-1,-1),4.5)]))
    return t


def _foundation_blocks(settings, st):
    fixed=settings.modality_code=='MOD_002'
    if fixed:
        blocks=[('Natureza da Tabela Price',['Na Tabela Price, o valor da prestação é uniforme, enquanto a composição interna se modifica ao longo do contrato. Os juros são calculados sobre o saldo principal existente em cada período; a amortização corresponde à diferença entre a prestação e os juros e é o único componente que efetivamente reduz o principal.','A potência presente na fórmula expressa a equivalência financeira entre valores situados em datas distintas e permite distribuir o pagamento em prestações uniformes. Esse fator exponencial não deve ser confundido, por si só, com a incorporação de juros vencidos ao saldo devedor. A análise técnica depende da base efetivamente utilizada para a incidência dos juros.']),('Ausência de incorporação de juros vencidos ao principal',['No fluxo demonstrado, os juros remuneram o capital durante o período e são pagos como parcela da prestação. A amortização reduz o saldo principal. Nos ciclos seguintes, os novos juros são calculados sobre o saldo remanescente após a amortização anterior, sem que a parcela de juros paga seja incorporada ao principal como nova base de cálculo.','O valor total da prestação não é integralmente abatido do saldo devedor principal. A redução do principal ocorre exclusivamente pela amortização; os juros remuneram o capital no período e o FGQC, quando informado, é demonstrado separadamente. Não há correção monetária na modalidade Credplan Fixo.'])]
    else:
        blocks=[('Natureza da metodologia de amortização',['A modalidade variável utiliza uma lógica de amortização vinculada ao saldo devedor atualizado e ao prazo remanescente. A prestação é formada por juros e amortização, mas a amortização não deve ser descrita como uma quantia numericamente imutável em todas as parcelas: ela é recalculada em cada ciclo a partir do saldo atualizado, deduzidos os juros do período, dividido pela quantidade de prestações remanescentes.','Os juros remuneram o capital utilizado no período e não reduzem o principal. A redução do saldo decorre da amortização. A correção monetária integra a evolução diária do saldo, segundo o índice e a defasagem definidos, e pode alterar a base sobre a qual a amortização do ciclo será apurada.']),('Ausência de incorporação de juros vencidos ao principal',['Na memória examinada, os juros são apropriados diariamente sobre o saldo vigente e liquidados como componente da prestação. Não se verifica a transferência de juros vencidos para o principal com o objetivo de formar uma nova base de incidência. A variação das prestações decorre da combinação entre saldo remanescente, correção monetária, juros do período e prazo ainda existente.','Para a correta leitura da evolução, devem ser distinguidos a atualização monetária do saldo, os juros remuneratórios do período e a amortização do principal. Somente a amortização reduz o saldo devedor principal.'])]
    out=[]
    for idx,(title,paras) in enumerate(blocks):
        parts=[Paragraph(title,st['subsection'])]
        for p in paras: parts.extend([Paragraph(p,st['body']),Spacer(1,1.5*mm)])
        out.append(KeepTogether(parts))
        if idx==0: out.append(HRFlowable(width='100%',thickness=0.4,color=LINE,spaceBefore=1.0*mm,spaceAfter=1.8*mm))
    return out


def _formula_card(title, formula, application, note, st):
    eq_wrap=Table([[Paragraph(formula,st['formula'])]],colWidths=[CONTENT_W-14*mm],hAlign='CENTER')
    eq_wrap.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),BLUE_SOFT2),('BOX',(0,0),(-1,-1),0.35,HexColor('#D9E4EF')),('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
    content=[[Paragraph(title,st['formula_title'])],[eq_wrap],[Paragraph(f'<b>Aplicação:</b> {application}',st['formula_app'])]]
    if note: content.append([Paragraph(note,st['formula_note'])])
    t=Table(content,colWidths=[CONTENT_W],hAlign='LEFT')
    t.setStyle(TableStyle([('BOX',(0,0),(-1,-1),0.55,LINE),('BACKGROUND',(0,0),(-1,-1),WHITE),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),4.2),('BOTTOMPADDING',(0,0),(-1,-1),4.2),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    return KeepTogether([t])


def _segment_table(settings,result,st):
    first_rows=[r for r in result.daily_rows if r.installment_number==1]; grouped=OrderedDict()
    for r in first_rows: grouped.setdefault(r.competence,[]).append(r)
    headers=['Competência','Dias processados','Taxa diária de juros','Índice de referência','Taxa diária de correção']; data=[[Paragraph(h,st['table_head']) for h in headers]]
    for competence,rows in grouped.items():
        rate=rows[0].daily_interest_rate if rows else Decimal('0'); corr_rates=[r.daily_correction_rate for r in rows if r.daily_correction_rate!=0]; corr=corr_rates[0] if corr_rates else Decimal('0'); refs=list(dict.fromkeys(r.index_reference for r in rows if r.index_reference)); ref=' a '.join(_competence_br(x) for x in refs) if refs else 'Não se aplica'
        vals=[_competence_br(competence),str(len(rows)),_pct(rate,settings.rounding.percentage_display_places),ref,_pct(corr,settings.rounding.percentage_display_places)]; data.append([Paragraph(v,st['table_right'] if i>=2 else st['table']) for i,v in enumerate(vals)])
    t=Table(data,colWidths=[29*mm,28*mm,38*mm,44*mm,38*mm],repeatRows=1,hAlign='LEFT'); cmds=[('BACKGROUND',(0,0),(-1,0),NAVY),('GRID',(0,0),(-1,-1),0.35,LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),3.5),('BOTTOMPADDING',(0,0),(-1,-1),3.5)]
    for i in range(1,len(data)):
        if i%2==0: cmds.append(('BACKGROUND',(0,i),(-1,i),BLUE_SOFT2))
    t.setStyle(TableStyle(cmds)); return t


def _info_note(text,st):
    t=Table([[Paragraph(text,st['formula_note'])]],colWidths=[CONTENT_W]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),BLUE_SOFT),('BOX',(0,0),(-1,-1),0.45,LINE),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)])); return t


def _evolution_table(rows,overrides,places,st):
    headers=['Nº','Referência','Saldo antes','Correção','Juros','Amortização','FGQC','Prestação','Saldo final']; data=[[Paragraph(h,st['table_head']) for h in headers]]
    for row in rows:
        vals=[row.installment_number,_date(row.reference_due_date),_money(base._value(row,'balance_before_installment',overrides),places),_money(base._value(row,'correction_amount',overrides),places),_money(base._value(row,'interest_amount',overrides),places),_money(base._value(row,'regular_amortization',overrides),places),_money(base._value(row,'fgqc_amount',overrides),places),_money(base._value(row,'installment_amount',overrides),places),_money(base._value(row,'closing_balance',overrides),places)]
        data.append([Paragraph(str(v),st['table_right'] if i>=2 else st['table']) for i,v in enumerate(vals)])
    t=Table(data,colWidths=[8*mm,21*mm,23*mm,19*mm,18*mm,22*mm,17*mm,20*mm,29*mm],repeatRows=1,hAlign='LEFT'); cmds=[('BACKGROUND',(0,0),(-1,0),NAVY),('GRID',(0,0),(-1,-1),0.35,LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),3.2),('BOTTOMPADDING',(0,0),(-1,-1),3.2)]
    for i in range(1,len(data)):
        if i%2==0: cmds.append(('BACKGROUND',(0,i),(-1,i),BLUE_SOFT2))
    t.setStyle(TableStyle(cmds)); return t


def _detail_table(row,overrides,settings,st):
    places=settings.rounding.money_places; fixed=settings.modality_code=='MOD_002'; before=base._value(row,'balance_before_installment',overrides); interest=base._value(row,'interest_amount',overrides); correction=base._value(row,'correction_amount',overrides); amort=base._value(row,'regular_amortization',overrides); installment=base._value(row,'installment_amount',overrides); closing=base._value(row,'closing_balance',overrides); opening=Decimal(str(row.opening_balance))
    if fixed:
        bp=before-interest; rows=[(1,'Saldo antes da prestação','saldo principal do ciclo + juros do período',f'{_money(bp,places)} + {_money(interest,places)}',_money(before,places)),(2,'Base principal','saldo antes - juros do período',f'{_money(before,places)} - {_money(interest,places)}',_money(bp,places)),(3,'Prestação Price','prestação fixa apurada pela Tabela Price',_money(installment,places),_money(installment,places)),(4,'Amortização','prestação Price - juros do período',f'{_money(installment,places)} - {_money(interest,places)}',_money(amort,places)),(5,'Saldo principal final','base principal - amortização',f'{_money(bp,places)} - {_money(amort,places)}',_money(closing,places))]
    else:
        ab=Decimal(str(row.amortization_base)); remaining=row.remaining_installments+1; rows=[(1,'Saldo antes da prestação','saldo inicial do ciclo + juros do período + correção monetária',f'{_money(opening,places)} + {_money(interest,places)} {base._signed_money(correction,places)}',_money(before,places)),(2,'Base de amortização','saldo antes - juros do período',f'{_money(before,places)} - {_money(interest,places)}',_money(ab,places)),(3,'Amortização','Base<sub>A</sub> / prestações remanescentes',f'{_money(ab,places)} / {remaining}',_money(amort,places)),(4,'Prestação','amortização + juros do período',f'{_money(amort,places)} + {_money(interest,places)}',_money(installment,places)),(5,'Saldo final','Base<sub>A</sub> - amortização',f'{_money(ab,places)} - {_money(amort,places)}',_money(closing,places))]
    data=[[Paragraph(h,st['table_head']) for h in ['Etapa','Descrição','Substituição numérica','Resultado']]]; num_style=ParagraphStyle('v13_num',parent=st['table'],alignment=TA_CENTER,fontName='Helvetica-Bold',textColor=NAVY_DARK)
    for num,desc,rule,sub,res in rows: data.append([Paragraph(str(num),num_style),Paragraph(f'<b>{desc}</b><br/><font color="#667A8D">{rule}</font>',st['table']),Paragraph(sub,st['table']),Paragraph(res,st['summary_value'])])
    t=Table(data,colWidths=[13*mm,63*mm,67*mm,34*mm],repeatRows=1,hAlign='LEFT'); cmds=[('BACKGROUND',(0,0),(-1,0),NAVY),('GRID',(0,0),(-1,-1),0.35,LINE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),2.6),('BOTTOMPADDING',(0,0),(-1,-1),2.6)]
    for i in range(1,len(data)):
        if i%2==0: cmds.append(('BACKGROUND',(0,i),(-1,i),BLUE_SOFT2))
    t.setStyle(TableStyle(cmds)); return t


def _summary_table(row,overrides,settings,st):
    p=settings.rounding.money_places; fixed=settings.modality_code=='MOD_002'; before=base._value(row,'balance_before_installment',overrides); interest=base._value(row,'interest_amount',overrides); correction=base._value(row,'correction_amount',overrides); amort=base._value(row,'regular_amortization',overrides); installment=base._value(row,'installment_amount',overrides); fgqc=base._value(row,'fgqc_amount',overrides); closing=base._value(row,'closing_balance',overrides)
    if fixed:
        bp=before-interest; items=[('Saldo antes',_money(before,p)),('Juros do período',_money(interest,p)),('Base principal',_money(bp,p)),('Prestação Price',_money(installment,p)),('Amortização',_money(amort,p)),('FGQC',_money(fgqc,p)),('Correção monetária','Não se aplica'),('Saldo final',_money(closing,p))]
    else:
        items=[('Saldo antes',_money(before,p)),('Juros do período',_money(interest,p)),('Correção monetária',_money(correction,p)),('Base de amortização',_money(row.amortization_base,p)),('Amortização',_money(amort,p)),('Prestação',_money(installment,p)),('FGQC',_money(fgqc,p)),('Saldo final',_money(closing,p))]
    rows=[]
    for i in range(0,len(items),2):
        a,b=items[i],items[i+1]; rows.append([Paragraph(a[0],st['label']),Paragraph(a[1],st['summary_value']),Paragraph(b[0],st['label']),Paragraph(b[1],st['summary_value'])])
    t=Table(rows,colWidths=[39.5*mm,49*mm,39.5*mm,49*mm],hAlign='LEFT'); t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.35,LINE),('BACKGROUND',(0,0),(-1,-1),WHITE),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),2.7),('BOTTOMPADDING',(0,0),(-1,-1),2.7)])); return t


def _eq_strip(text,st):
    t=Table([[Paragraph(text,ParagraphStyle('v13_eqstrip',parent=st['small'],alignment=TA_CENTER,textColor=NAVY_DARK))]],colWidths=[CONTENT_W]); t.setStyle(TableStyle([('BOX',(0,0),(-1,-1),0.35,LINE),('BACKGROUND',(0,0),(-1,-1),WHITE),('TOPPADDING',(0,0),(-1,-1),2.2),('BOTTOMPADDING',(0,0),(-1,-1),2.2)])); return t


def _detail_narrative(row,overrides,settings):
    p=settings.rounding.money_places; fixed=settings.modality_code=='MOD_002'; before=base._value(row,'balance_before_installment',overrides); interest=base._value(row,'interest_amount',overrides); correction=base._value(row,'correction_amount',overrides); amort=base._value(row,'regular_amortization',overrides); installment=base._value(row,'installment_amount',overrides); closing=base._value(row,'closing_balance',overrides); fgqc=base._value(row,'fgqc_amount',overrides)
    if fixed:
        bp=before-interest; return f'O saldo acumulado antes da prestação totalizou <b>{_money(before,p)}</b>. Nesse montante estão refletidos os juros do período de <b>{_money(interest,p)}</b>. Para identificar o principal antes da amortização, os juros são segregados da base, obtendo-se <b>{_money(bp,p)}</b>. A prestação fixa pela Tabela Price totalizou <b>{_money(installment,p)}</b>; a amortização, obtida pela diferença entre a prestação e os juros, foi de <b>{_money(amort,p)}</b>. Somente essa amortização reduz o principal, que passou a <b>{_money(closing,p)}</b>. O FGQC de <b>{_money(fgqc,p)}</b> permanece demonstrado separadamente. Não há correção monetária.'
    ab=Decimal(str(row.amortization_base)); remaining=row.remaining_installments+1; return f'O saldo acumulado antes da prestação totalizou <b>{_money(before,p)}</b>. Nesse montante estão refletidos os juros do período de <b>{_money(interest,p)}</b> e a correção monetária de <b>{_money(correction,p)}</b>. Para apurar a amortização, os juros do período são excluídos da base, obtendo-se <b>{_money(ab,p)}</b>. A divisão pelo prazo remanescente ({remaining}) resultou em amortização de <b>{_money(amort,p)}</b>. A prestação corresponde à soma da amortização com os juros e totalizou <b>{_money(installment,p)}</b>. Após o pagamento, o saldo principal passou a <b>{_money(closing,p)}</b>. O FGQC de <b>{_money(fgqc,p)}</b> permanece demonstrado separadamente.'


def _equations_for_row(row,overrides,settings):
    p=settings.rounding.money_places; fixed=settings.modality_code=='MOD_002'; before=base._value(row,'balance_before_installment',overrides); interest=base._value(row,'interest_amount',overrides); amort=base._value(row,'regular_amortization',overrides); installment=base._value(row,'installment_amount',overrides); closing=base._value(row,'closing_balance',overrides)
    if fixed:
        bp=before-interest; return [f'J<sub>período</sub> = J<sub>d1</sub> + ... + J<sub>dn</sub> = {_money(interest,p)}',f'A = PMT - J<sub>período</sub> = {_money(installment,p)} - {_money(interest,p)} = {_money(amort,p)}',f'SD<sub>final</sub> = Base<sub>principal</sub> - A = {_money(bp,p)} - {_money(amort,p)} = {_money(closing,p)}']
    ab=Decimal(str(row.amortization_base)); remaining=row.remaining_installments+1; return [f'Base<sub>A</sub> = {_money(before,p)} - {_money(interest,p)} = {_money(ab,p)}',f'A = Base<sub>A</sub> / n = {_money(ab,p)} / {remaining} = {_money(amort,p)}',f'P = A + J<sub>período</sub> = {_money(amort,p)} + {_money(interest,p)} = {_money(installment,p)}']


def _formula_story(tmp_dir,settings,result,st):
    first=result.installments[0]; out=[]
    sep='&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'
    if settings.modality_code=='MOD_002':
        monthly=_pct(result.monthly_interest_rate,settings.rounding.percentage_display_places); annual=_pct(settings.annual_interest_rate,settings.rounding.percentage_display_places)
        timing_factor=' × 1 / (1 + i<sub>m</sub>)' if settings.payment_timing==1 else ''
        timing_note='O fator 1/(1 + i<sub>m</sub>) representa o pagamento no início do período (tipo 1), conforme parametrizado para o contrato.' if settings.payment_timing==1 else 'O contrato utiliza pagamento no fim do período (tipo 0).'
        formulas=[
            ('Taxa mensal equivalente','i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1',f'i<sub>m</sub> = (1 + {annual})<super>1/12</super> - 1 = {monthly}.','A taxa anual contratual é convertida em taxa mensal equivalente antes da evolução diária.'),
            ('Evolução da carência até a primeira prestação','VF = PV × (1 + i<sub>m</sub>)<super>d1/DC1</super> × ... × (1 + i<sub>m</sub>)<super>dk/DCk</super>',f'PV = {_money(settings.initial_balance)}; evolução de {_date(settings.credit_date)} até {_date(first.reference_due_date)}; VF = {_money(first.balance_before_installment)}.','Na carência, a data do crédito é excluída e a data da primeira prestação é incluída. Os juros são apropriados diariamente ao longo dos segmentos de competência efetivamente atravessados (d = dias do segmento; DC = dias corridos da competência).'),
            ('Prestação pela Tabela Price',f'PMT = VF × i<sub>m</sub> / [1 - (1 + i<sub>m</sub>)<super>-n</super>]{timing_factor}',f'VF = {_money(first.balance_before_installment)}; i<sub>m</sub> = {monthly}; n = {settings.term}; prestação apurada = {_money(first.installment_amount)}.',timing_note),
            ('Composição da prestação e saldo principal',f'J<sub>1</sub> = J<sub>d1</sub> + J<sub>d2</sub> + ... + J<sub>dn</sub>{sep}A<sub>1</sub> = PMT - J<sub>1</sub>{sep}SD<sub>1</sub> = Base<sub>principal</sub> - A<sub>1</sub>',f'J<sub>1</sub> = {_money(first.interest_amount)}; A<sub>1</sub> = {_money(first.regular_amortization)}; saldo principal final = {_money(first.closing_balance)}.','Os juros remuneram o capital do período. Somente a amortização reduz o saldo principal.'),
        ]
    else:
        formulas=[
            ('Conversão da taxa anual em mensal equivalente','i<sub>m</sub> = (1 + i<sub>a</sub>)<super>1/12</super> - 1',f'Taxa anual de {_pct(settings.annual_interest_rate,settings.rounding.percentage_display_places)} convertida para taxa mensal equivalente de {_pct(result.monthly_interest_rate,settings.rounding.percentage_display_places)}.','A taxa mensal equivalente é a base para a conversão diária de juros.'),
            ('Taxa diária de juros','i<sub>d</sub> = (1 + i<sub>m</sub>)<super>1/DC</super> - 1','A taxa diária é recalculada conforme os dias corridos da competência efetivamente processada.','Os juros remuneratórios são apropriados diariamente, inclusive em fins de semana e feriados.'),
            ('Correção monetária diária do INPC','c<sub>d</sub> = (1 + c<sub>m</sub>)<super>1/DU</super> - 1','O INPC mensal de referência é convertido em taxa diária equivalente pelos dias úteis da competência.',f'A defasagem parametrizada é de {int(settings.index_lag_months or 0)} meses. A correção monetária é aplicada somente nos dias úteis.'),
            ('Juros do dia e juros do período',f'J<sub>d</sub> = SD<sub>0</sub> × i<sub>d</sub>{sep}J<sub>período</sub> = J<sub>d1</sub> + J<sub>d2</sub> + ... + J<sub>dn</sub>',f'O somatório dos juros diários da primeira prestação resultou em {_money(first.interest_amount)}.','Os juros acumulados são segregados no vencimento para que a base principal de amortização seja identificada.'),
            ('Evolução diária do saldo',f'SD<sub>J</sub> = SD<sub>0</sub> × (1 + i<sub>d</sub>){sep}SD<sub>J+CM</sub> = SD<sub>J</sub> × (1 + c<sub>d</sub>)','Primeiro são apropriados os juros do dia; em dia útil, a correção monetária é aplicada na sequência.','Eventuais eventos financeiros posteriores incidem depois dos encargos do dia.'),
            ('Amortização e prestação do ciclo',f'Base<sub>A</sub> = SD<sub>antes</sub> - J<sub>período</sub>{sep}A = Base<sub>A</sub> / n{sep}P = A + J<sub>período</sub>',f'Na primeira prestação: base de amortização {_money(first.amortization_base)}; amortização {_money(first.regular_amortization)}; prestação {_money(first.installment_amount)}.','Somente a amortização reduz o saldo principal; os juros remuneram o capital do período.'),
        ]
    for title,formula,app,note in formulas: out.extend([_formula_card(title,formula,app,note,st),Spacer(1,1.8*mm)])
    return out


def _intro_text(identity,modality_name,settings):
    contract=identity.contract_number.strip() or '-'
    if settings.modality_code=='MOD_002': return f'O presente parecer tem por objeto demonstrar, de forma técnica e padronizada, a evolução do contrato de empréstimo <b>{modality_name} nº {contract}</b>. A análise considera os parâmetros contratuais informados, a metodologia financeira aplicável à modalidade e a memória de cálculo produzida pelo Motor de Cálculos, permitindo a reprodução e a conferência das etapas que formam as prestações e o saldo devedor.'
    return f'O presente parecer tem por objeto a demonstração técnico-financeira da evolução do contrato de número <b>{contract}</b>, na modalidade <b>{modality_name}</b>, considerando os parâmetros contratuais informados, a metodologia de amortização correspondente e a memória de cálculo produzida pelo sistema.'


def _kpis(settings,result,modality_name):
    index='Sem correção' if settings.modality_code=='MOD_002' else (settings.index_code or 'INPC'); first_date=_date(result.installments[0].reference_due_date) if result.installments else '-'; system='Tabela Price' if settings.modality_code=='MOD_002' else 'SAC'
    return [('Modalidade',modality_name),('Sistema',system),('Taxa de juros (a.a.)',_pct(settings.annual_interest_rate,settings.rounding.percentage_display_places)),('Prazo',f'{settings.term} prestações'),('Índice de correção',index),('1ª prestação',first_date)]


def _build_story(tmp_dir,settings,result,modality_name,identity,selected_rows,overrides,st):
    story=[_section_title('Objeto e parâmetros contratuais',st),Paragraph(_intro_text(identity,modality_name,settings),st['body']),Spacer(1,1.8*mm),Paragraph('INDICADORES PRINCIPAIS',ParagraphStyle('v13_kpi_eyebrow',fontName='Helvetica-Bold',fontSize=5.8,leading=7,textColor=MUTED)),Spacer(1,1.1*mm),_KPIGrid(_kpis(settings,result,modality_name)),Spacer(1,3*mm),_param_table(identity,settings,result,st),Spacer(1,3.3*mm),_section_title('Fundamentação técnico-financeira',st)]
    story+=_foundation_blocks(settings,st)
    story += [Spacer(1,4*mm),CondPageBreak(50*mm),_section_title('Fórmulas e aplicação numérica',st)] + _formula_story(tmp_dir,settings,result,st)
    story += [Paragraph('Competências e taxas efetivamente utilizadas na primeira prestação',st['subsection']),_segment_table(settings,result,st),Spacer(1,1.6*mm)]
    if settings.modality_code=='MOD_002': story += [_info_note('O Credplan Fixo não utiliza correção monetária. O quadro acima apresenta as competências efetivamente processadas até a primeira prestação e as taxas diárias de juros utilizadas na evolução do saldo.',st),Spacer(1,3.2*mm)]
    else: story += [_info_note('<b>Nota sobre a competência 21-20.</b> A competência é identificada pelo mês em que o ciclo se inicia. Assim, a competência de dezembro corresponde ao período de 21/12 a 20/01. Com defasagem de 2 meses, utiliza-se o INPC de outubro, diarizado pelos dias úteis da competência. Os juros são apropriados diariamente ao longo de todo o ciclo; a correção monetária é aplicada somente em dias úteis.',st),Spacer(1,3.2*mm)]
    story += [CondPageBreak(32*mm),_section_title('Demonstração da evolução',st),Paragraph('Prestações de referência',st['subsection']),_evolution_table(selected_rows,overrides,settings.rounding.money_places,st),Spacer(1,4*mm)]
    for row in selected_rows:
        story += [CondPageBreak(42*mm),KeepTogether([Paragraph(f'Prestação {row.installment_number} - referência {_date(row.reference_due_date)}',st['detail_title']),Spacer(1,1.5*mm),Paragraph(_detail_narrative(row,overrides,settings),st['body']),Spacer(1,2*mm)]),_detail_table(row,overrides,settings,st),Spacer(1,2.2*mm),Paragraph(f'Resumo da Prestação {row.installment_number}',st['subsection']),_summary_table(row,overrides,settings,st),Spacer(1,1.3*mm)]
        for e in _equations_for_row(row,overrides,settings): story += [_eq_strip(e,st),Spacer(1,0.7*mm)]
        story += [Spacer(1,2*mm)]
    story += [CondPageBreak(22*mm),_section_title('Conclusão técnica',st)]; first=selected_rows[0]
    if settings.modality_code=='MOD_002':
        concl=['Sob o ponto de vista estritamente técnico-financeiro e considerando a metodologia demonstrada, não foi identificada a incorporação de juros vencidos ao saldo principal para a formação de novos juros. O fator exponencial da Tabela Price representa a equivalência financeira necessária à apuração de uma prestação uniforme.','Em cada período, os juros remuneram o saldo principal existente e a amortização, obtida pela diferença entre a prestação e os juros, reduz o saldo devedor. A parcela de juros não é abatida do principal nem incorporada a ele como nova base de cálculo. Não há correção monetária na modalidade Credplan Fixo.',f'Na primeira prestação apresentada, o valor de {_money(base._value(first,"installment_amount",overrides))} é composto por juros de {_money(base._value(first,"interest_amount",overrides))} e amortização de {_money(base._value(first,"regular_amortization",overrides))}, resultando em saldo principal de {_money(base._value(first,"closing_balance",overrides))} após a amortização. O FGQC, quando informado, é demonstrado separadamente e não integra a redução do principal.']
    else:
        concl=['Sob o ponto de vista estritamente técnico-financeiro e considerando a metodologia demonstrada, não foi identificada a incorporação de juros vencidos ao saldo principal para a geração de novos encargos. Os juros foram apropriados diariamente sobre o saldo vigente e apresentados como componente da prestação.','A correção monetária foi aplicada segundo o índice, a defasagem e os dias úteis da competência de aplicação. A redução do principal ocorreu por meio da amortização recalculada para o saldo atualizado e para o prazo remanescente, razão pela qual seu valor pode variar entre os ciclos.',f'Na primeira prestação de referência, a prestação de {_money(base._value(first,"installment_amount",overrides))} compreende juros de {_money(base._value(first,"interest_amount",overrides))} e amortização de {_money(base._value(first,"regular_amortization",overrides))}, com saldo final de {_money(base._value(first,"closing_balance",overrides))}. O FGQC de {_money(base._value(first,"fgqc_amount",overrides))} é demonstrado separadamente.']
    for p in concl: story += [Paragraph(p,st['body']),Spacer(1,2.2*mm)]
    if identity.additional_note.strip(): story += [Spacer(1,2*mm),Paragraph('Observação da elaboração',st['subsection']),Paragraph(identity.additional_note.strip().replace('\n','<br/>'),st['body'])]
    return story


def opinion_body_story(*,settings,result,modality_name:str,identity,selected_installment_numbers:Sequence[int],overrides:Mapping[int,Mapping[str,Decimal]]|None=None)->list:
    """Conteúdo técnico do parecer (sem abertura e fechamento), comum ao PDF e ao Word."""
    overrides=overrides or {}
    selected_rows=base.selected_installments(result,selected_installment_numbers)
    if settings.modality_code not in {'MOD_001','MOD_002'}:
        story=base.legacy_body_story(settings=settings,result=result,modality_name=modality_name,identity=identity,selected_rows=selected_rows,overrides=overrides)
    else:
        story=_build_story(None,settings,result,modality_name,identity,selected_rows,overrides,_styles())
    return funcef.normalize_typography(story)


def build_opinion_pdf(*,settings,result,modality_name:str,identity,selected_installment_numbers:Sequence[int],overrides:Mapping[int,Mapping[str,Decimal]]|None=None,logo_path:str|Path|None=None)->bytes:
    overrides=overrides or {}
    if settings.modality_code not in {'MOD_001','MOD_002'}:
        return _BASE_BUILD(settings=settings,result=result,modality_name=modality_name,identity=identity,selected_installment_numbers=selected_installment_numbers,overrides=overrides,logo_path=logo_path)
    body=opinion_body_story(settings=settings,result=result,modality_name=modality_name,identity=identity,selected_installment_numbers=selected_installment_numbers,overrides=overrides)
    output=BytesIO(); doc=BaseDocTemplate(output,pagesize=A4,leftMargin=MARGIN,rightMargin=MARGIN,topMargin=TOP_MARGIN,bottomMargin=BOTTOM_MARGIN,title=f'Parecer técnico - {modality_name}',author='FUNCEF - Motor de Cálculos',subject='Evolução contratual')
    frame=Frame(MARGIN,BOTTOM_MARGIN,CONTENT_W,A4[1]-TOP_MARGIN-BOTTOM_MARGIN,leftPadding=0,rightPadding=0,topPadding=0,bottomPadding=0,id='body_v13')
    doc.addPageTemplates([PageTemplate(id='all_v13',frames=[frame])])
    story=funcef.opening_story(identity.header,width=CONTENT_W,logo_path=logo_path,contract_number=identity.contract_number,modality_name=modality_name)
    story+=body
    story+=funcef.closing_story(identity.header,width=CONTENT_W)
    doc.build(funcef.prepare_pdf_story(story,width=CONTENT_W),canvasmaker=funcef.page_canvas(MARGIN)); return output.getvalue()
