from __future__ import annotations

from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# Design System compartilhado pelos PDFs do Motor de Cálculos.
PETROLEUM = HexColor("#005A70")
PETROLEUM_DARK = HexColor("#003E4B")
PETROLEUM_DEEP = HexColor("#073D6D")
PETROLEUM_SOFT = HexColor("#EEF5F7")
TEAL = HexColor("#0B7F90")
FUNCEF_BLUE = PETROLEUM_DEEP
ORANGE = HexColor("#F47C20")
ORANGE_DARK = HexColor("#D96612")
ORANGE_SOFT = HexColor("#FFF4EA")
# aliases mantidos para compatibilidade com geradores antigos; não representam dourado.
GOLD = ORANGE
GOLD_DARK = ORANGE_DARK
GOLD_SOFT = ORANGE_SOFT
INK = HexColor("#18343B")
MUTED = HexColor("#62777D")
LINE = HexColor("#D6E2E5")
SOFT = HexColor("#F4F7F9")
WHITE = colors.white
SUCCESS = HexColor("#287A58")
WARNING = HexColor("#9A5B18")

# Tipografia única sans-serif. Helvetica é nativa e estável no ReportLab.
SANS = "Helvetica"
SANS_BOLD = "Helvetica-Bold"
SERIF = SANS
SERIF_BOLD = SANS_BOLD

_HEADER_X = 10.5 * mm
_HEADER_W = A4[0] - 21 * mm
_FIRST_HEADER_H = 25.5 * mm
_LATER_HEADER_H = 14.5 * mm
_HEADER_TOP_GAP = 6.0 * mm
_RADIUS = 4.2 * mm


def _date_br(value: date | None) -> str:
    return value.strftime("%d/%m/%Y") if value else "-"


def first_page_top_margin() -> float:
    return 36 * mm


def later_page_top_margin() -> float:
    return 24 * mm


def _draw_logo(canvas, logo_path: Path | None, x: float, y: float, size: float) -> None:
    if not logo_path or not logo_path.exists():
        return
    try:
        canvas.drawImage(str(logo_path), x, y, width=size, height=size, preserveAspectRatio=True, anchor="c", mask="auto")
    except Exception:
        pass


def draw_document_header(canvas, doc, *, logo_path: Path | None, title: str, subtitle: str = "", meta: str = "", compact: bool = False) -> None:
    """Cabeçalho comum: azul profundo/petróleo e laranja FUNCEF apenas como acento."""
    width, height = A4
    box_h = _LATER_HEADER_H if compact else _FIRST_HEADER_H
    box_y = height - _HEADER_TOP_GAP - box_h
    radius = _RADIUS

    canvas.saveState()
    # sombra muito discreta
    canvas.setFillColor(HexColor("#E3E9ED"))
    canvas.roundRect(_HEADER_X + 0.6 * mm, box_y - 0.8 * mm, _HEADER_W, box_h, radius, fill=1, stroke=0)

    # azul com arredondamento à esquerda e acabamento reto à direita
    canvas.setFillColor(PETROLEUM_DEEP)
    canvas.roundRect(_HEADER_X, box_y, _HEADER_W, box_h, radius, fill=1, stroke=0)
    canvas.rect(_HEADER_X + _HEADER_W - radius, box_y, radius, box_h, fill=1, stroke=0)
    accent_w = 4.0 * mm
    canvas.setFillColor(ORANGE)
    canvas.rect(_HEADER_X + _HEADER_W - accent_w, box_y, accent_w, box_h, fill=1, stroke=0)

    if compact:
        logo_size = 8.4 * mm
        logo_x = _HEADER_X + 5.2 * mm
        logo_y = box_y + (box_h - logo_size) / 2
        _draw_logo(canvas, logo_path, logo_x, logo_y, logo_size)
        copy_x = _HEADER_X + 17.0 * mm
        canvas.setFillColor(WHITE); canvas.setFont(SANS_BOLD, 7.2)
        canvas.drawString(copy_x, box_y + 8.1 * mm, title)
        if subtitle or meta:
            canvas.setFillColor(HexColor("#D9E7EA")); canvas.setFont(SANS, 5.7)
            sub = f"{subtitle}  ·  {meta}" if subtitle and meta else (subtitle or meta)
            canvas.drawString(copy_x, box_y + 4.0 * mm, sub[:105])
    else:
        logo_size = 17.0 * mm
        logo_x = _HEADER_X + 5.2 * mm
        logo_y = box_y + (box_h - logo_size) / 2
        _draw_logo(canvas, logo_path, logo_x, logo_y, logo_size)
        copy_x = _HEADER_X + 27.0 * mm
        canvas.setFillColor(HexColor("#BFD6DB")); canvas.setFont(SANS_BOLD, 5.8)
        canvas.drawString(copy_x, box_y + 18.7 * mm, "FUNCEF  ·  MOTOR DE CÁLCULOS")
        canvas.setFillColor(WHITE); canvas.setFont(SANS_BOLD, 9.7 if len(title) <= 45 else 8.8)
        canvas.drawString(copy_x, box_y + 12.4 * mm, title)
        canvas.setFillColor(HexColor("#D9E7EA")); canvas.setFont(SANS, 6.1)
        if subtitle: canvas.drawString(copy_x, box_y + 7.6 * mm, subtitle[:93])
        if meta:
            canvas.setFillColor(HexColor("#C4D9DD")); canvas.setFont(SANS, 5.5)
            canvas.drawString(copy_x, box_y + 3.7 * mm, meta[:110])
    canvas.restoreState()


def draw_document_footer(canvas, doc, *, reference: str = "") -> None:
    width, _ = A4
    canvas.saveState(); y = 10.5 * mm
    canvas.setStrokeColor(LINE); canvas.setLineWidth(0.45)
    canvas.line(_HEADER_X, y + 3.3 * mm, width - _HEADER_X, y + 3.3 * mm)
    canvas.setFillColor(ORANGE); canvas.rect(_HEADER_X, y + 3.0 * mm, 13 * mm, 0.6 * mm, fill=1, stroke=0)
    canvas.setFillColor(MUTED); canvas.setFont(SANS, 5.7)
    left = "Relatório emitido pelo MOTOR DE CÁLCULOS · FUNCEF" + (f"  ·  {reference}" if reference else "")
    canvas.drawString(_HEADER_X, y, left[:125]); canvas.drawRightString(width - _HEADER_X, y, f"Página {doc.page}")
    canvas.restoreState()


def draw_document_chrome(canvas, doc, *, logo_path: Path | None, title: str, subtitle: str = "", meta: str = "", reference: str = "", compact: bool = False) -> None:
    draw_document_header(canvas, doc, logo_path=logo_path, title=title, subtitle=subtitle, meta=meta, compact=compact)
    draw_document_footer(canvas, doc, reference=reference)
