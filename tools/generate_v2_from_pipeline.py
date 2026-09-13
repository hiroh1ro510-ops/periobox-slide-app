"""Generate a v2 .pptx for every case in samples/_pipeline_out/ using the new
renderers. Reads the JSON produced by tools/run_pipeline_on_samples.py.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

import app as pipeline  # noqa: E402


def main() -> None:
    in_dir = REPO_ROOT / "samples" / "_pipeline_out"
    out_dir = REPO_ROOT / "samples" / "_generated_v2"
    out_dir.mkdir(exist_ok=True)

    for path in sorted(in_dir.glob("case*.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        summary = data.get("slide_summary", {})
        faithful = data.get("faithful_summary", {})
        pdf_relpath = data.get("pdf_file")
        pdf_bytes = None
        if pdf_relpath:
            pdf_full = REPO_ROOT / pdf_relpath
            if pdf_full.exists():
                pdf_bytes = pdf_full.read_bytes()
        try:
            pptx_bytes = pipeline.create_pptx(
                summary,
                box_number="000",
                pdf_bytes=pdf_bytes,
                faithful_summary=faithful,
            )
        except Exception as exc:
            print(f"  error  {path.stem}: {exc!r}")
            continue
        out_path = out_dir / f"{path.stem}_v2.pptx"
        out_path.write_bytes(pptx_bytes)
        study = summary.get("slide2", {}).get("study_type_class", "?")
        print(f"  ok    {path.stem:8} ({study:30}) -> {out_path.name}")


if __name__ == "__main__":
    main()
