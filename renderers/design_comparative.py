"""Design slide (slide 2) for comparative studies.

Covers ``rct``, ``prospective_cohort``, ``retrospective_comparative`` and
``mechanism`` classes — anything with well-defined comparison groups.
"""

from typing import Any, Dict, List, Optional

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from . import constants as C
from .primitives import (
    heading_pentagon_with_subtitle,
    insert_pdf_figure,
    pentagon_pill,
    right_arrow,
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


def _extract_groups(summary: Dict[str, Any]) -> List[Dict[str, str]]:
    slide2 = summary.get("slide2", {}) if isinstance(summary, dict) else {}
    ext = slide2.get("pattern_extensions", {}) if isinstance(slide2, dict) else {}
    detail = ext.get("groups_detail") if isinstance(ext, dict) else None
    out: List[Dict[str, str]] = []
    if isinstance(detail, list):
        for item in detail:
            if not isinstance(item, dict):
                continue
            name = _first(item.get("name"))
            if not name:
                continue
            n = _first(item.get("n"))
            desc = _first(item.get("intervention_or_exposure"))
            out.append({"name": name, "n": n, "desc": desc})
    if out:
        return out[:4]
    # Fallback to legacy design_schema.groups
    ds = slide2.get("design_schema", {}) if isinstance(slide2, dict) else {}
    for g in (ds.get("groups") or [])[:4]:
        if isinstance(g, dict):
            out.append(
                {
                    "name": _first(g.get("name"), default="比較群"),
                    "n": "",
                    "desc": _first(g.get("description")),
                }
            )
        else:
            out.append({"name": str(g), "n": "", "desc": ""})
    return out


def render(slide, summary: Dict[str, Any], pdf_bytes: Optional[bytes] = None) -> None:
    slide2 = summary.get("slide2", {}) if isinstance(summary, dict) else {}
    faithful_dc = (
        summary.get("_faithful_summary", {})
        .get("faithful_summary", {})
        .get("design_common", {})
        if isinstance(summary.get("_faithful_summary"), dict)
        else {}
    )

    subtitle = _first(
        slide2.get("study_type_label"),
        slide2.get("design_summary"),
    )
    heading_pentagon_with_subtitle(slide, C.DESIGN_HEADING_TEXT, subtitle)

    # Top block: subjects + methods short summary
    subjects = _first(
        faithful_dc.get("subjects_short") if isinstance(faithful_dc, dict) else None,
        slide2.get("participants"),
        default="",
    )
    methods_lines = _as_list(slide2.get("methods"))[:3]
    if not methods_lines and isinstance(faithful_dc, dict):
        methods_short = _first(faithful_dc.get("methods_short"))
        if methods_short:
            methods_lines = [line.strip() for line in methods_short.split("\n") if line.strip()][:3]

    top_lines: List[str] = []
    if subjects:
        top_lines.append(f"対象：{subjects}")
    for m in methods_lines:
        top_lines.append(f"・{m}")
    if not top_lines:
        top_lines = ["・本文から判別困難"]
    rounded_box(
        slide,
        "\n".join(top_lines),
        left=C.CONTENT_MARGIN_X,
        top=C.TOP_BOX_TOP,
        width=C.CONTENT_WIDTH,
        height=C.TOP_BOX_HEIGHT,
        fill=C.FILL_LIGHT_BLUE,
        line_color=C.LINE_MUTED,
        font_size=C.FS_BODY,
        bold=False,
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
    )

    # Groups row
    groups = _extract_groups(summary)
    if groups:
        n = min(len(groups), 4)
        total_gap = C.GROUP_ROW_GAP * (n - 1)
        box_w = min(2.35, (C.CONTENT_WIDTH - total_gap) / n)
        row_w = box_w * n + total_gap
        start_x = (C.SLIDE_WIDTH_IN - row_w) / 2
        y = C.GROUP_ROW_TOP
        palette = [C.FILL_LIGHT_PEACH, C.FILL_LIGHT_GREEN, C.FILL_LIGHT_BLUE, C.FILL_LIGHT_CREAM]
        for idx, group in enumerate(groups[:4]):
            x = start_x + idx * (box_w + C.GROUP_ROW_GAP)
            fill = palette[idx % len(palette)]
            text_parts = [group["name"]]
            if group.get("n"):
                text_parts.append(f"n={group['n']}")
            if group.get("desc"):
                text_parts.append(group["desc"])
            rounded_box(
                slide,
                "\n".join(text_parts),
                left=x,
                top=y,
                width=box_w,
                height=C.GROUP_BOX_HEIGHT,
                fill=fill,
                line_color=C.LINE_MUTED,
                font_size=C.FS_MICRO,
                bold=True,
                align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE,
            )
            if idx < n - 1:
                right_arrow(
                    slide,
                    left=x + box_w + (C.GROUP_ROW_GAP - C.ARROW_WIDTH) / 2,
                    top=y + (C.GROUP_BOX_HEIGHT - C.ARROW_HEIGHT) / 2,
                )

    # Evaluation items row
    outcomes = _as_list(slide2.get("outcomes"))
    if not outcomes and isinstance(faithful_dc, dict):
        outcomes = _as_list(faithful_dc.get("evaluation_items"))
    outcomes = outcomes[:5]
    textbox(
        slide,
        "評価項目",
        left=C.CONTENT_MARGIN_X,
        top=C.EVAL_LABEL_TOP,
        width=2.0,
        height=0.35,
        font_size=C.FS_SECTION,
        bold=True,
        color=C.TEXT_MAIN,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    if outcomes:
        n = len(outcomes)
        total_gap = C.EVAL_PILL_GAP * (n - 1)
        pill_w = min(2.0, (C.CONTENT_WIDTH - total_gap) / n)
        row_w = pill_w * n + total_gap
        start_x = (C.SLIDE_WIDTH_IN - row_w) / 2
        for idx, item in enumerate(outcomes):
            x = start_x + idx * (pill_w + C.EVAL_PILL_GAP)
            pentagon_pill(
                slide,
                item,
                left=x,
                top=C.EVAL_ROW_TOP,
                width=pill_w,
                height=C.EVAL_PILL_HEIGHT,
                fill=C.FILL_PENTAGON_PINK,
                text_color=C.TEXT_MAIN,
                font_size=C.FS_MICRO,
            )
    else:
        rounded_box(
            slide,
            "本文から判別困難",
            left=C.CONTENT_MARGIN_X,
            top=C.EVAL_ROW_TOP,
            width=C.CONTENT_WIDTH,
            height=C.EVAL_PILL_HEIGHT,
            fill=C.FILL_LIGHT_BLUE,
            line_color=C.LINE_MUTED,
            font_size=C.FS_MICRO,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )

    # Optional: crop a figure from the source PDF and place it on the right.
    fig_refs = slide2.get("figure_refs") or []
    if isinstance(faithful_dc, dict) and not fig_refs:
        fig_refs = faithful_dc.get("figure_refs") or []
    if fig_refs and pdf_bytes:
        insert_pdf_figure(
            slide,
            pdf_bytes,
            fig_refs,
            left=6.55,
            top=1.55,
            width=3.1,
            height=1.85,
        )
