"""Run the extraction / summary / verify / slide-summary pipeline on
every case*_paper.pdf under samples/ and dump per-case JSON to
samples/_pipeline_out/.

Reads OPENAI_API_KEY from .env or the shell env. Uses OPENAI_MODEL if set.
Skips cases whose output already exists so the script is resumable.

Usage:
    python3 tools/run_pipeline_on_samples.py [--force] [--only case01,case02]
"""

from __future__ import annotations

import argparse
import io
import json
import sys
import time
from pathlib import Path

# Ensure repo root on sys.path so we can `import app` and `import auth`.
REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from auth import get_secret_value, load_local_env  # noqa: E402

load_local_env()

import app as pipeline  # noqa: E402  (imports openai etc.)


def process_case(pdf_path: Path, out_dir: Path, force: bool) -> dict:
    stem = pdf_path.stem.replace("_paper", "")
    out_path = out_dir / f"{stem}.json"
    if out_path.exists() and not force:
        return {"case": stem, "status": "skipped", "path": str(out_path)}

    model = get_secret_value("OPENAI_MODEL", "gpt-4o-mini")
    api_key = get_secret_value("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not set (put it in .env or export it)")
    client = pipeline.get_openai_client(api_key)

    started = time.time()
    pdf_bytes = pdf_path.read_bytes()
    pdf_text = pipeline.extract_pdf_text(io.BytesIO(pdf_bytes))

    sections = pipeline.extract_sections_with_openai(client, model, pdf_text)
    faithful = pipeline.create_faithful_summary_with_openai(client, model, sections)
    faithful = pipeline.verify_study_type_with_openai(client, model, faithful, sections)
    summary = pipeline.create_slide_summary_with_openai(client, model, faithful)
    elapsed = round(time.time() - started, 1)

    payload = {
        "case": stem,
        "pdf_file": str(pdf_path.relative_to(REPO_ROOT)),
        "model": model,
        "elapsed_seconds": elapsed,
        "sections": sections,
        "faithful_summary": faithful,
        "slide_summary": summary,
    }
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"case": stem, "status": "ok", "path": str(out_path), "seconds": elapsed}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true", help="Re-run even if output exists")
    parser.add_argument("--only", default="", help="Comma-separated case stems to run (e.g. case01,case02)")
    args = parser.parse_args()

    samples_dir = REPO_ROOT / "samples"
    out_dir = samples_dir / "_pipeline_out"
    out_dir.mkdir(exist_ok=True)

    only_filter = {s.strip() for s in args.only.split(",") if s.strip()}
    pdf_paths = sorted(samples_dir.glob("case*_paper.pdf"))
    if only_filter:
        pdf_paths = [p for p in pdf_paths if p.stem.replace("_paper", "") in only_filter]

    print(f"Processing {len(pdf_paths)} cases -> {out_dir}")
    for pdf in pdf_paths:
        try:
            result = process_case(pdf, out_dir, args.force)
            print(f"  {result['status']:8}  {result['case']}  {result.get('seconds', '-')}s")
        except Exception as exc:  # noqa: BLE001
            print(f"  error    {pdf.stem}: {exc}")


if __name__ == "__main__":
    main()
