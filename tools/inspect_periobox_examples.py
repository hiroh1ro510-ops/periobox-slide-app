import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


EMU_PER_INCH = 914400


def emu_to_in(value):
    return round(value / EMU_PER_INCH, 3)


def rgb_to_hex(rgb):
    if rgb is None:
        return None
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def run_text_info(shape):
    infos = []
    if not getattr(shape, "has_text_frame", False):
        return infos
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            text = run.text.strip()
            if not text:
                continue
            color = None
            try:
                color = rgb_to_hex(run.font.color.rgb)
            except Exception:
                color = None
            infos.append(
                {
                    "text": text,
                    "font_size": round(run.font.size.pt, 1) if run.font.size else None,
                    "font_name": run.font.name,
                    "bold": run.font.bold,
                    "color": color,
                }
            )
    return infos


def shape_text(shape):
    if not getattr(shape, "has_text_frame", False):
        return ""
    return "\n".join(p.text for p in shape.text_frame.paragraphs).strip()


def inspect_deck(path):
    prs = Presentation(path)
    deck = {
        "file": str(path),
        "slides": len(prs.slides),
        "width_in": emu_to_in(prs.slide_width),
        "height_in": emu_to_in(prs.slide_height),
        "slide_items": [],
    }
    for slide_index, slide in enumerate(prs.slides, start=1):
        items = []
        for shape in slide.shapes:
            text = shape_text(shape)
            if not text:
                continue
            items.append(
                {
                    "shape_type": str(shape.shape_type),
                    "left": emu_to_in(shape.left),
                    "top": emu_to_in(shape.top),
                    "width": emu_to_in(shape.width),
                    "height": emu_to_in(shape.height),
                    "text": text,
                    "runs": run_text_info(shape),
                }
            )
        deck["slide_items"].append({"slide": slide_index, "items": items})
    return deck


def summarize(decks):
    font_sizes = Counter()
    font_names = Counter()
    colors = Counter()
    slide_counts = Counter()
    common_text = Counter()
    positions = defaultdict(list)

    for deck in decks:
        slide_counts[deck["slides"]] += 1
        for slide in deck["slide_items"]:
            for item in slide["items"]:
                text = item["text"].strip()
                if text:
                    common_text[text[:80]] += 1
                    positions[slide["slide"]].append(
                        {
                            "left": item["left"],
                            "top": item["top"],
                            "width": item["width"],
                            "height": item["height"],
                            "text": text[:80],
                        }
                    )
                for run in item["runs"]:
                    if run["font_size"]:
                        font_sizes[str(run["font_size"])] += 1
                    if run["font_name"]:
                        font_names[run["font_name"]] += 1
                    if run["color"]:
                        colors[run["color"]] += 1

    return {
        "slide_counts": slide_counts.most_common(),
        "font_sizes": font_sizes.most_common(20),
        "font_names": font_names.most_common(20),
        "colors": colors.most_common(20),
        "common_text": common_text.most_common(30),
        "positions_by_slide": {
            str(slide): items[:20] for slide, items in sorted(positions.items())
        },
    }


def main():
    root = Path(sys.argv[1])
    decks = [inspect_deck(path) for path in sorted(root.glob("case*_periobox.pptx"))]
    print(json.dumps({"summary": summarize(decks), "decks": decks}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
