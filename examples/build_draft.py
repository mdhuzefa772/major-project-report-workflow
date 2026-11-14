"""Build formatted chapter DOCX files from a draft JSON — no API key needed.

The draft JSON contains the chapter content (already structured) plus optional
anchor_overrides applied to the spec before building. This is the deterministic
formatting half of the workflow: run it anywhere the worker package is
installed (Codespace, locally) and download the generated files.

Usage:
  python build_draft.py <draft.json> [--spec spec.json] [--out output_dir]

Draft JSON shape (a single chapter, or {"chapters": [...], "anchor_overrides": {...}}):
  {"anchor_overrides": {"FONT.FAMILY": "Times New Roman"},
   "chapters": [{"chapter_num": 1, "title": "INTRODUCTION",
                 "intro_paragraph": "...",
                 "sections": [{"number": "1.1", "heading": "...",
                               "paragraphs": ["..."], "bullets": [],
                               "subsections": [], "tables": [],
                               "figure_captions": []}]}]}

Default spec: runs/phase0/spec.json (frozen from the real sample inputs).
"""

import argparse
import json
import sys
from pathlib import Path

try:
    from worker.workflows.docxgen import build_chapter_docx
    from worker.workflows.verification import verify_compliance
except ImportError:
    sys.path.insert(0, "apps/worker/src")
    from worker.workflows.docxgen import build_chapter_docx
    from worker.workflows.verification import verify_compliance

DEFAULT_SPEC = Path("runs/phase0/spec.json")


def apply_overrides(spec: dict, overrides: dict | None) -> dict:
    for path, value in (overrides or {}).items():
        group, key = path.split(".", 1)
        if group in spec and key in spec[group]:
            spec[group][key] = value
        else:
            print(f"[WARN] unknown override anchor ignored: {path}")
    return spec


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("draft", help="draft JSON file (chapter content)")
    parser.add_argument("--spec", default=str(DEFAULT_SPEC), help="anchor spec JSON")
    parser.add_argument("--out", default="output", help="output directory")
    args = parser.parse_args()

    data = json.loads(Path(args.draft).read_text(encoding="utf-8"))
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    spec = apply_overrides(spec, data.get("anchor_overrides"))

    chapters = data["chapters"] if isinstance(data, dict) and "chapters" in data else [data]
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    failures = 0
    for chapter in chapters:
        chapter_num = chapter["chapter_num"]
        docx = build_chapter_docx(chapter, spec)
        out_path = out_dir / f"Chapter-{chapter_num}.docx"
        out_path.write_bytes(docx)

        violations = verify_compliance(docx, spec, chapter)
        blocking = [v for v in violations
                    if v["severity"].startswith(("HIGH", "MEDIUM")) and not v["anchor"].startswith("DEPTH.")]
        info = [v for v in violations if v not in blocking]
        print(f"Chapter-{chapter_num}: wrote {out_path} "
              f"({len(docx)} bytes, {len(chapter.get('sections', []))} sections)")
        if blocking:
            failures += 1
            print(f"  BLOCKING VIOLATIONS ({len(blocking)}):")
            for v in blocking:
                print(f"    [{v['severity']}] {v['anchor']}: actual={v['actual']!r} expected={v['expected']!r}")
        else:
            print("  format checks: PASSED")
        for v in info:
            print(f"    [{v['severity']}] {v['anchor']}: {v['actual']}")

    if failures:
        sys.exit(1)
    print(f"\nAll chapters written to {out_dir}/ — download them from the Explorer "
          f"(right-click -> Download).")


if __name__ == "__main__":
    main()
