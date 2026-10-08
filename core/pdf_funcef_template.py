from __future__ import annotations

"""Padrão documental FUNCEF (modelo "Manifestação de Subsídios").

Reproduz no PDF a abertura do modelo Word usado nos trabalhos operacionais
da COPART: faixa institucional com logo e data, título, quadros 01 a 03
(Dados do Processo, Participante e Operação, Demanda), rótulo de
classificação no topo de cada página, rodapé "COPART · GERAT · DIBEN" e o
quadro "Responsável pela informação" ao final do documento.

Medidas e cores foram extraídas do MODELO.docx. A largura da abertura acompanha
a área útil do parecer; as proporções internas seguem a grade original.
"""

import re
from dataclasses import dataclass, replace
from datetime import date
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.graphics import shapes as _graphics_shapes
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import CondPageBreak, Flowable, KeepTogether, Paragraph, Spacer, Table, TableStyle


BANNER_NAVY = HexColor("#1C2541")
DATE_TEAL = HexColor("#31859B")
BANNER_TEXT = HexColor("#F2F2F2")
TITLE_BLUE = HexColor("#003E5B")
BADGE_BLUE = HexColor("#0B5B91")
LABEL_FILL = HexColor("#DBE5F1")
LABEL_TEXT = HexColor("#586674")
VALUE_TEXT = HexColor("#17212B")
GRID_LINE = HexColor("#D7DEE5")
RESPONSIBLE_LINE = HexColor("#D9D9D9")
FOOTER_TEXT = HexColor("#8A949E")
BLACK = HexColor("#000000")
WHITE = HexColor("#FFFFFF")

# Corpo começa logo abaixo do rótulo de classificação; o rodapé ocupa a faixa inferior.
TOP_MARGIN = 15 * mm
BOTTOM_MARGIN = 20 * mm

CLASSIFICATION_LABEL = "#10 Corporativo - FUNCEF"
FOOTER_LABEL = "COPART · GERAT · DIBEN"
DOCUMENT_TITLE = "Manifestação de Subsídios"
DOCUMENT_SUBTITLE = "Empréstimos e Financiamento Habitacional"
ORGANIZATION_LINES = (
    ("Fundação dos Economiários Federais", True),
    ("DIBEN – Diretoria de Benefícios", False),
    ("GERAT – Gerência de Relacionamento e Atendimento", False),
    ("COPART – Coordenação de Operação com Participantes", False),
)
DEFAULT_AUTHORS = "Fundação dos Economiários Federais – FUNCEF"
DEFAULT_OPERATION = "Empréstimo"
BRAZILIAN_STATES = (
    "", "AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB",
    "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO",
)
# Numeração única do CNJ (Resolução 65/2008): NNNNNNN-DD.AAAA.J.TR.OOOO
_CNJ_PROCESS = re.compile(r"^\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}$")

# Escala tipográfica corporativa: uma única família (Helvetica, equivalente
# métrica do Arial do modelo) e corpo mínimo legível em todo o documento.
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
FONT_ITALIC = "Helvetica-Oblique"
FONT_BOLD_ITALIC = "Helvetica-BoldOblique"
BODY_SIZE = 9.0
BODY_LEADING = 12.8
SUBSECTION_SIZE = 9.5
NOTE_SIZE = 8.0
TABLE_SIZE = 7.5
TABLE_LEADING = 9.2
MIN_SIZE = 7.0

# Os ícones vetoriais iniciam o estado gráfico com a fonte padrão do ReportLab
# (Times-Roman), que passaria a constar no PDF mesmo sem texto desenhado.
_graphics_shapes.STATE_DEFAULTS["fontName"] = FONT_REGULAR

_LOGO_FALLBACK = Path(__file__).resolve().parent.parent / "assets" / "logo_funcef_branca.png"

# Geometria do grupo de formas do cabeçalho no modelo (EMU).
_GROUP_W = 10663555
_GROUP_H = 1438275
_BANNER_ASPECT = 1000125 / 6715125
_LOGO_BOX = (513716, 241279, 914401, 914400)
_TEXT_BOX = (1605217, 224088, 5562601, 1089661)
_DATE_BOX = (7243446, 195262, 3053079, 1047750)
_SHAPE_RADIUS = 0.16667  # ajuste "roundRect" padrão do Word


@dataclass(frozen=True)
class ManifestationHeader:
    issue_date: date | None = None
    process_number: str = ""
    district: str = ""
    court: str = ""
    state: str = ""
    authors: str = DEFAULT_AUTHORS
    lawyer: str = ""
    # Modo web/LGPD: mutuário e matrícula não são coletados; os campos saem em branco.
    borrowers: str = ""
    registrations: str = ""
    operation: str = DEFAULT_OPERATION
    contracts: str = ""
    modality: str = ""
    destination_area: str = ""
    subject: str = ""
    reference: str = ""
    responsible: str = ""


_HEADER_LABELS = (
    ("process_number", "Processo nº"), ("district", "Comarca"), ("court", "Vara"), ("state", "UF"),
    ("authors", "Autor(es)"), ("lawyer", "Advogado responsável"), ("borrowers", "Mutuário(s)"),
    ("registrations", "Matrícula(s)"), ("operation", "Operação com participante"),
    ("contracts", "Contrato(s)"), ("modality", "Modalidade"), ("destination_area", "Área de destino"),
    ("subject", "Assunto"), ("reference", "Referência"), ("responsible", "Responsável pela informação"),
)


def header_review(header: ManifestationHeader) -> list[str]:
    """Avisos não bloqueantes sobre o cabeçalho preenchido manualmente."""
    notices = []
    blank = [label for field, label in _HEADER_LABELS if not str(getattr(header, field) or "").strip()]
    if blank:
        notices.append("Campos do cabeçalho em branco no PDF: " + ", ".join(blank) + ".")
    process = header.process_number.strip()
    if process and not _CNJ_PROCESS.match(process):
        notices.append(
            "O número do processo não segue o padrão CNJ (0000000-00.0000.0.00.0000). "
            "Confira antes de gerar o PDF; ele será impresso exatamente como digitado."
        )
    return notices


def _styles():
    return {
        "title": ParagraphStyle(
            "funcef_title", fontName="Helvetica-Bold", fontSize=17, leading=19.5, textColor=TITLE_BLUE,
        ),
        "subtitle": ParagraphStyle(
            "funcef_subtitle", fontName="Helvetica", fontSize=8.5, leading=9.8, textColor=LABEL_TEXT,
        ),
        "badge": ParagraphStyle(
            "funcef_badge", fontName="Helvetica-Bold", fontSize=10, leading=11.5, textColor=WHITE, alignment=TA_CENTER,
        ),
        "section": ParagraphStyle(
            "funcef_section", fontName="Helvetica-Bold", fontSize=10, leading=11.5, textColor=TITLE_BLUE,
        ),
        "label": ParagraphStyle(
            "funcef_label", fontName="Helvetica", fontSize=10, leading=11.5, textColor=LABEL_TEXT,
        ),
        "label_center": ParagraphStyle(
            "funcef_label_center", fontName="Helvetica", fontSize=10, leading=11.5, textColor=LABEL_TEXT,
            alignment=TA_CENTER,
        ),
        "value": ParagraphStyle(
            "funcef_value", fontName="Helvetica", fontSize=10, leading=11.5, textColor=VALUE_TEXT,
        ),
        "value_center": ParagraphStyle(
            "funcef_value_center", fontName="Helvetica", fontSize=10, leading=11.5, textColor=VALUE_TEXT,
            alignment=TA_CENTER,
        ),
        "responsible": ParagraphStyle(
            "funcef_responsible", fontName="Helvetica-Bold", fontSize=10, leading=11.5, textColor=BLACK,
            alignment=TA_JUSTIFY,
        ),
    }


def _text(value) -> str:
    return escape(str(value or "").strip())


class _Banner(Flowable):
    """Faixa institucional do modelo: logo, identificação das áreas e data."""

    def __init__(self, width: float, issue_date: date, logo_path: Path | None):
        super().__init__()
        self.width = width
        self.height = width * _BANNER_ASPECT
        self.issue_date = issue_date
        self.logo_path = logo_path

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def _box(self, gx, gy, gw, gh):
        sx = self.width / _GROUP_W
        sy = self.height / _GROUP_H
        w = gw * sx
        h = gh * sy
        return gx * sx, self.height - gy * sy - h, w, h

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(BANNER_NAVY)
        c.roundRect(0, 0, self.width, self.height, _SHAPE_RADIUS * self.height, stroke=0, fill=1)

        if self.logo_path and self.logo_path.exists():
            x, y, w, h = self._box(*_LOGO_BOX)
            c.drawImage(str(self.logo_path), x, y, width=w, height=h, mask="auto")

        x, y, w, h = self._box(*_TEXT_BOX)
        inset = 2.54 * mm
        line_h = 9 * 1.32
        block_top = y + h / 2 + (len(ORGANIZATION_LINES) * line_h) / 2
        c.setFillColor(BANNER_TEXT)
        for idx, (line, bold) in enumerate(ORGANIZATION_LINES):
            c.setFont("Helvetica-Bold" if bold else "Helvetica", 9)
            c.drawString(x + inset, block_top - idx * line_h - 0.8 * line_h, line)

        x, y, w, h = self._box(*_DATE_BOX)
        c.setFillColor(DATE_TEAL)
        c.roundRect(x, y, w, h, _SHAPE_RADIUS * min(w, h), stroke=0, fill=1)
        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(x + w / 2, y + 0.60 * h, "Data:")
        c.setFont("Helvetica", 11)
        c.drawCentredString(x + w / 2, y + 0.24 * h, self.issue_date.strftime("%d/%m/%Y"))
        c.restoreState()


# Grades dos quadros 01 a 03 (larguras em twips do modelo).
# Cada célula: (tipo, campo ou texto, colunas mescladas, centralizado).
_SECTION_GRIDS = (
    (
        "01", "Dados do Processo",
        (625, 425, 284, 992, 4820, 992, 283, 709, 1466),
        (496, 405, 398, 464),
        (
            (("L", "Processo nº", 3, True), ("V", "process_number", 2, False),
             ("L", "Comarca", 1, True), ("V", "district", 3, True)),
            (("L", "Vara", 1, False), ("V", "court", 6, False),
             ("L", "UF", 1, False), ("V", "state", 1, True)),
            (("L", "Autor(es)", 2, True), ("V", "authors", 7, False)),
            (("L", "Advogado responsável", 4, True), ("V", "lawyer", 5, False)),
        ),
    ),
    (
        "02", "Participante e Operação",
        (1192, 284, 283, 1134, 4675, 1275, 1753),
        (343, 391, 539, 689),
        (
            (("L", "Mutuário(s)", 1, False), ("V", "borrowers", 4, False),
             ("L", "Matrícula(s)", 1, False), ("V", "registrations", 1, False)),
            (("L", "Operação com participante", 4, False), ("V", "operation", 3, False)),
            (("L", "Contrato(s)", 3, False), ("V", "contracts", 4, False)),
            (("L", "Modalidade", 2, False), ("V", "modality", 5, False)),
        ),
    ),
    (
        "03", "Demanda",
        (909, 431, 283, 8973),
        (441, 391, 539),
        (
            (("L", "Área de destino", 3, False), ("V", "destination_area", 1, False)),
            (("L", "Assunto", 1, False), ("V", "subject", 3, False)),
            (("L", "Referência", 2, False), ("V", "reference", 2, False)),
        ),
    ),
)

_TWIP = 25.4 * mm / 1440
_BOX_PAD_X = 4.0
_BOX_PAD_TOP = 3.25
_BOX_PAD_BOTTOM = 3.0


def _heading_table(number, title, width, st) -> Table:
    badge_w = width * 351 / 10602
    heading = Table(
        [[Paragraph(number, st["badge"]), Paragraph(_text(title), st["section"])]],
        colWidths=[badge_w, width - badge_w],
    )
    heading.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), BADGE_BLUE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (0, 0), 1.25),
        ("RIGHTPADDING", (0, 0), (0, 0), 1.25),
        ("TOPPADDING", (0, 0), (0, 0), 0.9),
        ("BOTTOMPADDING", (0, 0), (0, 0), 0.9),
        ("LEFTPADDING", (1, 0), (1, 0), 2.25),
        ("RIGHTPADDING", (1, 0), (1, 0), 0),
        ("TOPPADDING", (1, 0), (1, 0), 0.5),
        ("BOTTOMPADDING", (1, 0), (1, 0), 0.5),
    ]))
    return heading


def _section_box(number, title, grid_twips, heights_twips, rows, header: ManifestationHeader, width, st) -> Table:
    inner_w = width - 2 * _BOX_PAD_X
    heading = _heading_table(number, title, inner_w, st)

    total = sum(grid_twips)
    col_widths = [inner_w * tw / total for tw in grid_twips]
    data = []
    commands = [
        ("GRID", (0, 0), (-1, -1), 0.875, GRID_LINE),
        ("BACKGROUND", (0, 0), (-1, -1), WHITE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.75),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.4),
    ]
    for row_idx, cells in enumerate(rows):
        row = [""] * len(grid_twips)
        col = 0
        for kind, content, span, centered in cells:
            if kind == "L":
                para = Paragraph(_text(content), st["label_center" if centered else "label"])
                commands.append(("BACKGROUND", (col, row_idx), (col + span - 1, row_idx), LABEL_FILL))
            else:
                para = Paragraph(_text(getattr(header, content)), st["value_center" if centered else "value"])
            row[col] = para
            if span > 1:
                commands.append(("SPAN", (col, row_idx), (col + span - 1, row_idx)))
            col += span
        data.append(row)
    grid = Table(
        data,
        colWidths=col_widths,
        minRowHeights=[tw * _TWIP for tw in heights_twips],
    )
    grid.setStyle(TableStyle(commands))

    box = Table(
        [[[Spacer(1, 1), heading, Spacer(1, 2), grid, Spacer(1, 1)]]],
        colWidths=[width],
    )
    box.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.875, GRID_LINE),
        ("BACKGROUND", (0, 0), (-1, -1), WHITE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), _BOX_PAD_X),
        ("RIGHTPADDING", (0, 0), (-1, -1), _BOX_PAD_X),
        ("TOPPADDING", (0, 0), (-1, -1), _BOX_PAD_TOP),
        ("BOTTOMPADDING", (0, 0), (-1, -1), _BOX_PAD_BOTTOM),
    ]))
    return box


def opening_story(
    header: ManifestationHeader | None,
    *,
    width: float,
    logo_path: str | Path | None,
    contract_number: str = "",
    modality_name: str = "",
) -> list:
    """Abertura do documento no padrão do modelo, anterior ao conteúdo técnico."""
    header = header or ManifestationHeader()
    if not header.contracts.strip() or not header.modality.strip():
        header = replace(
            header,
            contracts=header.contracts.strip() or contract_number.strip(),
            modality=header.modality.strip() or modality_name.strip(),
        )
    logo = Path(logo_path) if logo_path else None
    if not logo or not logo.exists():
        logo = _LOGO_FALLBACK
    st = _styles()

    story = [
        _Banner(width, header.issue_date or date.today(), logo),
        Spacer(1, 3.7 * mm),
        Paragraph(DOCUMENT_TITLE, st["title"]),
        Paragraph(DOCUMENT_SUBTITLE, st["subtitle"]),
        Spacer(1, 3.8 * mm),
    ]
    for idx, (number, title, grid, heights, rows) in enumerate(_SECTION_GRIDS):
        if idx:
            story.append(Spacer(1, 7 * mm))
        story.append(_section_box(number, title, grid, heights, rows, header, width, st))
    story.append(Spacer(1, 7 * mm))
    return story


def closing_story(header: ManifestationHeader | None, *, width: float) -> list:
    """Quadro "Responsável pela informação" que encerra o documento."""
    header = header or ManifestationHeader()
    st = _styles()
    name = _text(header.responsible)
    text = "Responsável pela informação: " + name if name else "Responsável pela informação: "
    table = Table(
        [[Paragraph(text, st["responsible"])]],
        colWidths=[width],
        minRowHeights=[389 * _TWIP],
    )
    table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, RESPONSIBLE_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    closing = KeepTogether([Spacer(1, 7 * mm), table])
    closing.funcef_closing = True
    return [closing]


def section_heading(number: str, title: str, width: float) -> Table:
    """Título de seção do conteúdo técnico, no mesmo desenho dos quadros 01 a 03."""
    table = _heading_table(number, title, width, _styles())
    table.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.875, GRID_LINE)]))
    table.spaceBefore = 1.5 * mm
    table.spaceAfter = 3.2 * mm
    table.keepWithNext = True
    return table


def _corporate_font(name: str | None) -> str:
    lowered = (name or "").lower()
    bold = "bold" in lowered
    italic = "italic" in lowered or "oblique" in lowered
    if bold and italic:
        return FONT_BOLD_ITALIC
    if bold:
        return FONT_BOLD
    if italic:
        return FONT_ITALIC
    return FONT_REGULAR


def _corporate_style(style: ParagraphStyle, cache: dict) -> ParagraphStyle:
    cached = cache.get(id(style))
    if cached is not None:
        return cached
    font = _corporate_font(style.fontName)
    size = max(float(style.fontSize), MIN_SIZE)
    leading = max(float(style.leading or 0), size * 1.18)
    if font == style.fontName and size == style.fontSize and leading == style.leading:
        result = style
    else:
        result = ParagraphStyle(f"{style.name}_corp", parent=style, fontName=font, fontSize=size, leading=leading)
    cache[id(style)] = result
    return result


def _normalize(item, cache: dict):
    if isinstance(item, Paragraph):
        style = _corporate_style(item.style, cache)
        if style is item.style:
            return item
        return Paragraph(item.text, style, bulletText=getattr(item, "bulletText", None))
    if isinstance(item, list):
        return [_normalize(x, cache) for x in item]
    if isinstance(item, KeepTogether):
        item._content = [_normalize(x, cache) for x in item._content]
        return item
    if isinstance(item, Table):
        for row_idx, row in enumerate(item._cellvalues):
            for col_idx, cell in enumerate(row):
                if isinstance(cell, (Paragraph, Table, KeepTogether, list)):
                    row[col_idx] = _normalize(cell, cache)
        for row in item._cellStyles:
            for cell_style in row:
                cell_style.fontname = _corporate_font(cell_style.fontname)
                cell_style.fontsize = max(float(cell_style.fontsize), MIN_SIZE)
                cell_style.leading = max(float(cell_style.leading), cell_style.fontsize * 1.18)
        return item
    return item


def normalize_typography(story: list) -> list:
    """Garante uma única família tipográfica e o corpo mínimo em todo o documento."""
    cache: dict = {}
    return [_normalize(item, cache) for item in story]


# Altura mínima do conteúdo que deve acompanhar um título na mesma página.
_ORPHAN_CONTENT = 32 * mm


def _keeps_with_next(item) -> bool:
    try:
        return bool(item.getKeepWithNext())
    except AttributeError:
        return False


def _height(item, width: float, frame_height: float) -> float:
    _, height = item.wrapOn(Canvas(BytesIO(), pagesize=A4), width, frame_height)
    if isinstance(item, KeepTogether):
        # KeepTogether informa uma altura fictícia para forçar a divisão; a real fica em _H.
        height = item._H
    return height + item.getSpaceBefore() + item.getSpaceAfter()


def keep_headings_with_content(story: list, *, width: float, frame_height: float) -> list:
    """Evita títulos isolados no pé da página.

    Antes de cada título (ou sequência de títulos) é inserida uma quebra
    condicional do tamanho do título somado ao início do conteúdo seguinte.
    Blocos indivisíveis entram por inteiro; tabelas e textos, até 32 mm.
    """
    out = []
    idx = 0
    while idx < len(story):
        item = story[idx]
        if not _keeps_with_next(item):
            out.append(item)
            idx += 1
            continue
        need = 0.0
        run_end = idx
        while run_end < len(story) and (
            _keeps_with_next(story[run_end]) or isinstance(story[run_end], Spacer)
        ):
            need += _height(story[run_end], width, frame_height)
            run_end += 1
        if run_end < len(story) and not isinstance(story[run_end], CondPageBreak):
            following = story[run_end]
            following_h = _height(following, width, frame_height)
            if isinstance(following, KeepTogether):
                need += min(following_h, frame_height - need)
            else:
                need += min(following_h, _ORPHAN_CONTENT)
        out.append(CondPageBreak(need))
        out.extend(story[idx:run_end])
        idx = run_end
    return out


def _draw_page_chrome(canvas: Canvas, margin: float, page_number: int, page_count: int) -> None:
    width, height = A4
    canvas.saveState()
    canvas.setFillColor(BLACK)
    canvas.setFont("Helvetica", 10)
    canvas.drawString(width - 49.86 * mm, height - 9.0 * mm, CLASSIFICATION_LABEL)
    canvas.setFillColor(FOOTER_TEXT)
    canvas.setFont("Helvetica", 6)
    canvas.drawRightString(width - margin, 16.8 * mm, FOOTER_LABEL)
    canvas.drawString(margin, 16.8 * mm, f"Página {page_number} de {page_count}")
    canvas.restoreState()


def page_canvas(margin: float):
    """Canvas com rótulo de classificação, rodapé do modelo e "Página X de Y"."""

    class _FuncefCanvas(Canvas):
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
                _draw_page_chrome(self, margin, self._pageNumber, page_count)
                Canvas.showPage(self)
            Canvas.save(self)

    return _FuncefCanvas


def _keep_closing_with_text(story: list) -> list:
    """O quadro do responsável nunca fica sozinho: acompanha o último parágrafo."""
    for idx, item in enumerate(story):
        if getattr(item, "funcef_closing", False):
            prev = idx - 1
            while prev >= 0 and isinstance(story[prev], Spacer):
                prev -= 1
            if prev >= 0 and isinstance(story[prev], Paragraph):
                merged = KeepTogether(story[prev:idx] + list(item._content))
                return story[:prev] + [merged] + story[idx + 1:]
    return story


def prepare_pdf_story(story: list, *, width: float) -> list:
    """Tipografia corporativa e controle de títulos órfãos para o PDF."""
    frame_height = A4[1] - TOP_MARGIN - BOTTOM_MARGIN
    story = keep_headings_with_content(normalize_typography(story), width=width, frame_height=frame_height)
    return _keep_closing_with_text(story)
