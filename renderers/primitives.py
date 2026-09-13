"""Shape primitives that draw PerioBox-style elements at the standard palette."""

import io
from typing import Any, Dict, List, Optional

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

from . import constants as C


def _set_text(
    shape,
    text: str,
    *,
    font_size: int,
    bold: bool = True,
    color: Optional[RGBColor] = None,
    font_name: Optional[str] = None,
    align=PP_ALIGN.CENTER,
    anchor=MSO_ANCHOR.MIDDLE,
    line_spacing: float = 1.05,
) -> None:
    frame = shape.text_frame
    frame.margin_left = Inches(0.06)
    frame.margin_right = Inches(0.06)
    frame.margin_top = Inches(0.04)
    frame.margin_bottom = Inches(0.04)
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    text = str(text) if text is not None else ""
    lines = text.split("\n") if text else [""]
    for index, line in enumerate(lines):
        para = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        para.alignment = align
        para.line_spacing = line_spacing
        # Clear existing runs on the default first paragraph.
        if index == 0:
            for run in list(para.runs):
                run.text = ""
        run = para.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.bold = bold
        if font_name:
            run.font.name = font_name
        else:
            run.font.name = C.FONT_JP
        if color is not None:
            run.font.color.rgb = color


def _apply_fill(shape, fill: Optional[RGBColor]) -> None:
    if fill is None:
        shape.fill.background()
        return
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill


def _apply_line(shape, line_color: Optional[RGBColor], line_width_pt: float = 0.75) -> None:
    if line_color is None:
        shape.line.fill.background()
        return
    shape.line.color.rgb = line_color
    shape.line.width = Pt(line_width_pt)


def _disable_shadow(shape) -> None:
    """python-pptx does not expose shadow directly; we clear it via XML."""
    sppr = shape.shadow._element  # noqa: SLF001 (accessing internal helper)
    # No-op fallback: python-pptx applies a default outer shadow to auto shapes.
    # Removing it requires editing XML; skipping is acceptable for now.
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def pentagon_pill(
    slide,
    text: str,
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    fill: RGBColor = C.FILL_PENTAGON_PINK,
    text_color: RGBColor = C.TEXT_MAIN,
    font_size: int = C.FS_HEADING,
    bold: bool = True,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.REGULAR_PENTAGON,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    _apply_fill(shape, fill)
    _apply_line(shape, None)
    _disable_shadow(shape)
    _set_text(shape, text, font_size=font_size, bold=bold, color=text_color)
    return shape


def rounded_box(
    slide,
    text: str,
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    fill: Optional[RGBColor] = C.FILL_LIGHT_BLUE,
    line_color: Optional[RGBColor] = C.LINE_MUTED,
    line_width_pt: float = 0.75,
    text_color: RGBColor = C.TEXT_MAIN,
    font_size: int = C.FS_BODY,
    bold: bool = False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    _apply_fill(shape, fill)
    _apply_line(shape, line_color, line_width_pt)
    _disable_shadow(shape)
    _set_text(shape, text, font_size=font_size, bold=bold, color=text_color, align=align, anchor=anchor)
    return shape


def right_arrow(
    slide,
    *,
    left: float,
    top: float,
    width: float = C.ARROW_WIDTH,
    height: float = C.ARROW_HEIGHT,
    fill: RGBColor = C.FILL_LIGHT_PEACH,
    line_color: Optional[RGBColor] = C.LINE_MUTED,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_ARROW,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    _apply_fill(shape, fill)
    _apply_line(shape, line_color, 0.5)
    _disable_shadow(shape)
    return shape


def down_arrow(
    slide,
    *,
    left: float,
    top: float,
    width: float = C.ARROW_HEIGHT,
    height: float = C.ARROW_WIDTH,
    fill: RGBColor = C.FILL_LIGHT_PEACH,
    line_color: Optional[RGBColor] = C.LINE_MUTED,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.DOWN_ARROW,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    _apply_fill(shape, fill)
    _apply_line(shape, line_color, 0.5)
    _disable_shadow(shape)
    return shape


def oval_badge(
    slide,
    text: str,
    *,
    left: float,
    top: float,
    diameter: float = 0.4,
    fill: RGBColor = C.FILL_GREEN,
    text_color: RGBColor = C.FILL_WHITE,
    font_size: int = C.FS_SMALL,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left),
        Inches(top),
        Inches(diameter),
        Inches(diameter),
    )
    _apply_fill(shape, fill)
    _apply_line(shape, None)
    _disable_shadow(shape)
    _set_text(shape, text, font_size=font_size, bold=True, color=text_color)
    return shape


def insert_pdf_figure(
    slide,
    pdf_bytes: Optional[bytes],
    figure_refs: List[Dict[str, Any]],
    *,
    left: float,
    top: float,
    width: float,
    height: float,
) -> bool:
    """Render the first available figure_refs page onto the slide.

    Returns True if an image was placed, False otherwise. Silently falls
    through when PyMuPDF is not available or no valid page is found.
    """
    if not pdf_bytes or not figure_refs:
        return False
    try:
        import fitz  # type: ignore
    except Exception:
        return False
    page_no = None
    for ref in figure_refs:
        if not isinstance(ref, dict):
            continue
        page = ref.get("page")
        if isinstance(page, (int, float)) and int(page) > 0:
            page_no = int(page)
            break
        if isinstance(page, str) and page.strip().isdigit():
            page_no = int(page.strip())
            break
    if page_no is None:
        return False
    try:
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
        if len(document) == 0:
            return False
        idx = max(0, min(page_no - 1, len(document) - 1))
        pixmap = document.load_page(idx).get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
        stream = io.BytesIO(pixmap.tobytes("png"))
        stream.seek(0)
        document.close()
        slide.shapes.add_picture(
            stream, Inches(left), Inches(top), width=Inches(width), height=Inches(height)
        )
        return True
    except Exception:
        return False


def textbox(
    slide,
    text: str,
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    font_size: int = C.FS_BODY,
    bold: bool = False,
    color: RGBColor = C.TEXT_MAIN,
    font_name: Optional[str] = None,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    line_spacing: float = 1.1,
):
    tb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    _set_text(
        tb,
        text,
        font_size=font_size,
        bold=bold,
        color=color,
        font_name=font_name,
        align=align,
        anchor=anchor,
        line_spacing=line_spacing,
    )
    return tb


def heading_pentagon_with_subtitle(slide, heading: str, subtitle: str) -> None:
    """Draw slide-2/3 top: centered plain heading + top-left subtitle pill.

    Sample decks use a large centered text heading (not a shape) with the
    subtitle placed as a small blue rounded-rectangle pill in the top-left
    corner, so we mimic that instead of the earlier pink pentagon approach.
    """
    # Centered large heading text
    textbox(
        slide,
        heading,
        left=3.0,
        top=0.25,
        width=4.0,
        height=0.7,
        font_size=32,
        bold=True,
        color=C.TEXT_BLACK,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    if subtitle:
        rounded_box(
            slide,
            subtitle,
            left=0.55,
            top=0.9,
            width=4.0,
            height=0.42,
            fill=C.FILL_LIGHT_BLUE,
            line_color=C.LINE_MUTED,
            text_color=C.TEXT_MAIN,
            font_size=C.FS_SMALL,
            bold=True,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )
