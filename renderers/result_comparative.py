"""Results slide (slide 3) for comparative studies."""

from typing import Any, Dict, List, Optional

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

from . import constants as C
from .primitives import (
    heading_pentagon_with_subtitle,
    insert_pdf_figure,
    rounded_box,
    textbox,
)


def _as_list(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    text = str(value).strip()
    return [text] if text else []


def _first(*values, default: str = "") -> str:
    for v in values:
        if v is None:
            continue
        text = str(v).strip()
        if text and text != "本文から判別困難":
            return text
    return default


def _draw_matrix_bignumbers(
    slide, matrix: Dict[str, Any], top: float, max_height: float, right_bound: float
) -> Optional[float]:
    """Render a small comparison matrix (≤3 rows × ≤3 cols) as large red
    numbers in a free-positioned grid, matching case01's slide-3 style
    (metric headings across the top, group labels down the left, big red
    values in the cells)."""
    rows = _as_list(matrix.get("rows"))
    columns = _as_list(matrix.get("columns"))
    cells = matrix.get("cells")
    if not rows or not columns or not isinstance(cells, list):
        return None
    if len(rows) > 3 or len(columns) > 3:
        return None

    left_edge = 0.7
    total_width = min(right_bound - left_edge, 5.5)
    label_col_w = 1.4
    metric_area_w = total_width - label_col_w
    metric_col_w = metric_area_w / len(columns)

    header_y = top
    row_h = 1.05
    grid_top = header_y + 0.7

    # Column headers (metric names) — bold black
    for c_idx, col_label in enumerate(columns):
        x = left_edge + label_col_w + c_idx * metric_col_w
        textbox(
            slide,
            col_label,
            left=x,
            top=header_y,
            width=metric_col_w,
            height=0.5,
            font_size=C.FS_SUBHEADING,
            bold=True,
            color=C.TEXT_MAIN,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )

    # Group labels down the left column, and big red numbers per cell
    for r_idx, row_label in enumerate(rows):
        y = grid_top + r_idx * row_h
        textbox(
            slide,
            row_label,
            left=left_edge,
            top=y + 0.15,
            width=label_col_w,
            height=row_h - 0.3,
            font_size=C.FS_SUBHEADING,
            bold=True,
            color=C.TEXT_MAIN,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        row_cells = cells[r_idx] if r_idx < len(cells) and isinstance(cells[r_idx], list) else []
        # Emphasize the first (typically "better") row in red, rest in default color
        color = C.TEXT_RED if r_idx == 0 else C.TEXT_MAIN
        for c_idx in range(len(columns)):
            x = left_edge + label_col_w + c_idx * metric_col_w
            value = str(row_cells[c_idx]) if c_idx < len(row_cells) else ""
            textbox(
                slide,
                value,
                left=x,
                top=y,
                width=metric_col_w,
                height=row_h,
                font_size=30,
                bold=True,
                color=color,
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE,
            )

    return grid_top + len(rows) * row_h


def _draw_matrix(slide, matrix: Dict[str, Any], top: float, max_height: float) -> Optional[float]:
    """Draw the group x metric matrix as a native pptx table. Returns bottom y."""
    rows = _as_list(matrix.get("rows"))
    columns = _as_list(matrix.get("columns"))
    cells = matrix.get("cells")
    if not rows or not columns or not isinstance(cells, list):
        return None

    row_axis = str(matrix.get("row_axis", "")).strip()
    col_axis = str(matrix.get("column_axis", "")).strip()

    # Table dimensions
    tbl_cols = len(columns) + 1  # + row header
    tbl_rows = len(rows) + 1     # + column header
    if tbl_cols > 6:
        tbl_cols = 6
        columns = columns[:5]
    if tbl_rows > 6:
        tbl_rows = 6
        rows = rows[:5]

    width = min(C.CONTENT_WIDTH, 0.9 * len(columns) + 1.6)
    height = min(max_height, 0.4 * (len(rows) + 1) + 0.2)
    left = (C.SLIDE_WIDTH_IN - width) / 2

    tbl_shape = slide.shapes.add_table(tbl_rows, tbl_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    table = tbl_shape.table

    # Header row
    header_label = col_axis or ""
    _write_cell(table.cell(0, 0), header_label, bold=True, fill=C.FILL_PENTAGON_PINK, color=C.TEXT_MAIN)
    for c_idx, col_label in enumerate(columns, start=1):
        _write_cell(table.cell(0, c_idx), col_label, bold=True, fill=C.FILL_PENTAGON_PINK, color=C.TEXT_MAIN)

    # Body rows
    for r_idx, row_label in enumerate(rows, start=1):
        _write_cell(table.cell(r_idx, 0), row_label, bold=True, fill=C.FILL_LIGHT_BLUE, color=C.TEXT_MAIN)
        row_cells = cells[r_idx - 1] if r_idx - 1 < len(cells) and isinstance(cells[r_idx - 1], list) else []
        for c_idx in range(1, tbl_cols):
            value = row_cells[c_idx - 1] if c_idx - 1 < len(row_cells) else ""
            _write_cell(table.cell(r_idx, c_idx), str(value), bold=False, fill=C.FILL_WHITE, color=C.TEXT_MAIN)

    sig_note = str(matrix.get("significance_note", "")).strip()
    if sig_note:
        textbox(
            slide,
            sig_note,
            left=left,
            top=top + height + 0.05,
            width=width,
            height=0.35,
            font_size=C.FS_SMALL,
            color=C.TEXT_RED,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
            bold=True,
        )
        return top + height + 0.45
    return top + height + 0.1


def _write_cell(cell, text: str, *, bold: bool, fill, color) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
    cell.margin_left = Inches(0.05)
    cell.margin_right = Inches(0.05)
    cell.margin_top = Inches(0.03)
    cell.margin_bottom = Inches(0.03)
    tf = cell.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    # Clear existing default run
    for r in list(para.runs):
        r.text = ""
    run = para.add_run()
    run.text = str(text)
    run.font.size = Pt(C.FS_SMALL)
    run.font.bold = bold
    run.font.name = C.FONT_JP
    run.font.color.rgb = color


def render(slide, summary: Dict[str, Any], pdf_bytes: Optional[bytes] = None) -> None:
    slide3 = summary.get("slide3", {}) if isinstance(summary, dict) else {}
    faithful_rc = (
        summary.get("_faithful_summary", {})
        .get("faithful_summary", {})
        .get("result_common", {})
        if isinstance(summary.get("_faithful_summary"), dict)
        else {}
    )

    subtitle = _first(
        (faithful_rc.get("main_conclusion_ja") if isinstance(faithful_rc, dict) else None),
        (_as_list(slide3.get("main_results"))[:1] or [""])[0],
        default="",
    )
    heading_pentagon_with_subtitle(slide, C.RESULT_HEADING_TEXT, "")

    # Main conclusion band (below pentagon)
    main_results = _as_list(slide3.get("main_results"))[:3]
    if not main_results:
        main_results = [subtitle] if subtitle else ["本文から判別困難"]
    conclusion_text = "\n".join(f"・{item}" for item in main_results)
    rounded_box(
        slide,
        conclusion_text,
        left=C.CONTENT_MARGIN_X,
        top=1.45,
        width=C.CONTENT_WIDTH,
        height=1.4,
        fill=C.FILL_LIGHT_BLUE,
        line_color=C.LINE_MUTED,
        font_size=C.FS_BODY,
        bold=True,
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
    )

    # Comparison matrix (preferred) or fallback bullets of key_numbers/statistical_notes
    ext = slide3.get("pattern_extensions", {}) if isinstance(slide3, dict) else {}
    matrix = ext.get("comparison_matrix") if isinstance(ext, dict) else None
    matrix_bottom: Optional[float] = None
    # Reserve the right side for the PDF figure crop
    fig_refs = slide3.get("figure_refs") or []
    if isinstance(faithful_rc, dict) and not fig_refs:
        fig_refs = faithful_rc.get("figure_refs") or []
    right_bound = 6.4 if (fig_refs and pdf_bytes) else 9.5
    if isinstance(matrix, dict) and matrix.get("rows") and matrix.get("columns"):
        rows_n = len(_as_list(matrix.get("rows")))
        cols_n = len(_as_list(matrix.get("columns")))
        if rows_n <= 3 and cols_n <= 3:
            matrix_bottom = _draw_matrix_bignumbers(slide, matrix, top=3.05, max_height=2.6, right_bound=right_bound)
            sig_note = str(matrix.get("significance_note", "")).strip()
            if sig_note:
                textbox(
                    slide,
                    sig_note,
                    left=0.7,
                    top=(matrix_bottom or 5.5) + 0.05,
                    width=right_bound - 0.7,
                    height=0.35,
                    font_size=C.FS_SMALL,
                    color=C.TEXT_RED,
                    align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE,
                    bold=True,
                )
        else:
            matrix_bottom = _draw_matrix(slide, matrix, top=3.05, max_height=2.6)

    if matrix_bottom is None:
        # Fallback: show key_numbers + statistical_notes bullets
        key_numbers = _as_list(slide3.get("key_numbers"))[:3]
        stats = _as_list(slide3.get("statistical_notes"))[:3]
        bullets = [f"・{n}" for n in key_numbers] + [f"・{s}" for s in stats]
        if bullets:
            rounded_box(
                slide,
                "\n".join(bullets),
                left=C.CONTENT_MARGIN_X,
                top=3.05,
                width=C.CONTENT_WIDTH,
                height=2.4,
                fill=C.FILL_LIGHT_CREAM,
                line_color=C.LINE_MUTED,
                font_size=C.FS_BODY,
                bold=False,
                align=PP_ALIGN.LEFT,
                anchor=MSO_ANCHOR.TOP,
            )

    # "さらに詳しく" band
    rounded_box(
        slide,
        C.DETAIL_LABEL_TEXT,
        left=C.CONTENT_MARGIN_X,
        top=C.DETAIL_BAND_TOP,
        width=C.DETAIL_BAND_WIDTH,
        height=C.DETAIL_BAND_HEIGHT,
        fill=C.FILL_PENTAGON_PINK,
        line_color=None,
        font_size=C.FS_SMALL,
        bold=True,
        align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    detail_text = _first(
        slide3.get("more_detail"),
        faithful_rc.get("detail_note_ja") if isinstance(faithful_rc, dict) else None,
        default="",
    )
    textbox(
        slide,
        detail_text or "本文から判別困難",
        left=C.CONTENT_MARGIN_X + C.DETAIL_BAND_WIDTH + 0.15,
        top=C.DETAIL_BAND_TOP,
        width=C.CONTENT_WIDTH - C.DETAIL_BAND_WIDTH - 0.15,
        height=0.7,
        font_size=C.FS_BODY,
        color=C.TEXT_MAIN,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # Optional: crop a results figure from the source PDF and place on the right.
    if fig_refs and pdf_bytes:
        insert_pdf_figure(
            slide,
            pdf_bytes,
            fig_refs,
            left=6.6,
            top=3.05,
            width=3.0,
            height=2.5,
        )
