"""Verification and fix activities (spec section 13): anchor compliance on generated DOCX."""

import base64
import io
import json
import os
import re
import tempfile

import mistralai.workflows as workflows
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from . import inspect_docx
from .anchors import anchor

ALIGN_ENUM = {
    "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
    "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
    "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
    "JUSTIFIED": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _classify(text: str) -> str:
    if re.match(r"^\d+\.\d+\.\d+\s", text):
        return "subheading"
    if re.match(r"^\d+\.\d+\s", text):
        return "main_heading"
    return "body"


def verify_compliance(docx_bytes: bytes, spec: dict, draft: dict) -> list[dict]:
    """Check the generated DOCX against the frozen anchor spec. Returns violations."""
    violations: list[dict] = []

    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "output.docx")
        with open(path, "wb") as handle:
            handle.write(docx_bytes)
        out = inspect_docx.inspect(path)

    def check(anchor_path, actual, expected, severity="HIGH", differ=None):
        if expected is None:
            return
        if differ is not None:
            if differ:
                violations.append({"anchor": anchor_path, "actual": actual, "expected": expected, "severity": severity})
            return
        if actual != expected:
            violations.append({"anchor": anchor_path, "actual": actual, "expected": expected, "severity": severity})

    page = out.get("body_page_setup") or {}
    size = page.get("page_size") or {}
    check("PAGE.WIDTH_TWIPS", size.get("w_twips"), anchor(spec, "PAGE.WIDTH_TWIPS"))
    check("PAGE.HEIGHT_TWIPS", size.get("h_twips"), anchor(spec, "PAGE.HEIGHT_TWIPS"))
    margins = page.get("margins") or {}
    for side in ("left", "right", "top", "bottom"):
        expected = anchor(spec, f"PAGE.MARGIN_{side.upper()}_IN")
        actual = margins.get(f"{side}_inches")
        if expected is not None and actual is not None:
            check(f"PAGE.MARGIN_{side.upper()}_IN", actual, expected,
                  differ=abs(float(actual) - float(expected)) > 0.01)

    paragraphs = out.get("paragraphs") or []
    title_pattern = re.compile(r"^CHAPTER\s*-\s*\d+$", re.IGNORECASE)
    title_upper = str(draft.get("title", "")).strip().upper()

    def is_title_line(text: str) -> bool:
        return bool(title_pattern.match(text)) or text.strip().upper() == title_upper

    content_paras = [p for p in paragraphs if not p.get("is_empty") and p.get("runs")]

    for para in content_paras:
        text = para["runs"][0].get("text", "")
        if is_title_line(text):
            continue
        kind = _classify(text)
        first = para["runs"][0]
        if kind == "main_heading":
            check("FONT.MAIN_HEADING_SIZE_PT", first.get("size_pt"),
                  anchor(spec, "FONT.MAIN_HEADING_SIZE_PT"), "MEDIUM")
            check("FONT.MAIN_HEADING_BOLD", bool(first.get("bold")),
                  bool(anchor(spec, "FONT.MAIN_HEADING_BOLD", True)) if spec["FONT"]["MAIN_HEADING_BOLD"] is not None else None,
                  "MEDIUM")
        elif kind == "subheading":
            check("FONT.SUB_HEADING_SIZE_PT", first.get("size_pt"),
                  anchor(spec, "FONT.SUB_HEADING_SIZE_PT"), "MEDIUM")
            check("FONT.SUB_HEADING_ITALIC", bool(first.get("italic")),
                  bool(anchor(spec, "FONT.SUB_HEADING_ITALIC", True)) if spec["FONT"]["SUB_HEADING_ITALIC"] is not None else None,
                  "MEDIUM")
        else:
            check("FONT.BODY_SIZE_PT", first.get("size_pt"), anchor(spec, "FONT.BODY_SIZE_PT"), "MEDIUM")

    expected_family = anchor(spec, "FONT.FAMILY")
    if expected_family:
        fonts = {run.get("font") for para in content_paras for run in para.get("runs", []) if run.get("font")}
        unexpected = fonts - {expected_family}
        if unexpected:
            check("FONT.FAMILY", sorted(unexpected), expected_family, "MEDIUM", differ=True)

    check("SUBSEQUENT_PAGES.titlePg", page.get("has_titlePg"), True)

    header_refs = page.get("header_refs") or []
    if anchor(spec, "SUBSEQUENT_PAGES.HAS_HEADER", True) and not header_refs:
        check("SUBSEQUENT_PAGES.HAS_HEADER", False, True, differ=True)

    start = anchor(spec, "SUBSEQUENT_PAGES.PAGE_NUMBER_START", 2)
    actual_start = page.get("pgNumType_start")
    if actual_start is not None:
        check("SUBSEQUENT_PAGES.PAGE_NUMBER_START", int(actual_start) + 1, start, "MEDIUM")

    headers = (out.get("header_footer_parts") or {}).get("headers") or {}
    header_text = None
    for part_paragraphs in headers.values():
        for para in part_paragraphs:
            if (para.get("text") or "").strip():
                header_text = para
                break
    if header_text and anchor(spec, "SUBSEQUENT_PAGES.HEADER_TEXT_TEMPLATE"):
        expected_header = anchor(spec, "SUBSEQUENT_PAGES.HEADER_TEXT_TEMPLATE").replace(
            "{{CHAPTER_NUM}}", str(draft["chapter_num"]))
        check("SUBSEQUENT_PAGES.HEADER_TEXT_TEMPLATE", (header_text.get("text") or "").strip(),
              expected_header.strip(), "MEDIUM")
        if header_text.get("alignment"):
            check("SUBSEQUENT_PAGES.HEADER_ALIGNMENT", header_text.get("alignment").upper(),
                  str(anchor(spec, "SUBSEQUENT_PAGES.HEADER_ALIGNMENT", "RIGHT")).upper(), "MEDIUM")

    mode = draft.get("content_mode", "GENERATE")
    for section in draft.get("sections", []):
        words = sum(len(re.findall(r"\S+", p)) for p in section.get("paragraphs", []))
        for sub in section.get("subsections", []):
            words += sum(len(re.findall(r"\S+", p)) for p in sub.get("paragraphs", []))
        if words < 150:
            violations.append({
                "anchor": "DEPTH.150_WORDS",
                "actual": {"section": section.get("number"), "words": words},
                "expected": ">= 150 words",
                "severity": "INFO_ESCALATE" if mode == "REFORMAT_ONLY" else "HIGH",
            })

    return violations


@workflows.activity()
async def verify_chapter_docx(docx_b64: str, spec: dict, draft: dict) -> dict:
    docx_bytes = base64.b64decode(docx_b64)
    violations = verify_compliance(docx_bytes, spec, draft)
    return {
        "passed": not any(v["severity"].startswith(("HIGH", "MEDIUM")) for v in violations),
        "violations": violations,
    }


@workflows.activity()
async def apply_format_fixes(docx_b64: str, spec: dict, violations: list[dict]) -> dict:
    """Deterministic second pass: re-apply anchor formatting to paragraphs in the output."""
    fixable = {v["anchor"] for v in violations if v["anchor"].startswith("FONT.")}
    if not fixable:
        return {"docx_b64": docx_b64, "fixed": []}

    docx_bytes = base64.b64decode(docx_b64)
    doc = Document(io.BytesIO(docx_bytes))

    fixed = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        kind = _classify(text)
        run = para.runs[0] if para.runs else None
        if run is None:
            continue
        if kind == "main_heading" and "FONT.MAIN_HEADING_SIZE_PT" in fixable:
            run.font.size = Pt(float(anchor(spec, "FONT.MAIN_HEADING_SIZE_PT")))
            run.font.bold = anchor(spec, "FONT.MAIN_HEADING_BOLD", True)
            run.font.name = anchor(spec, "FONT.FAMILY")
            fixed.append(text[:40])
        elif kind == "subheading" and ("FONT.SUB_HEADING_SIZE_PT" in fixable or "FONT.SUB_HEADING_ITALIC" in fixable):
            run.font.size = Pt(float(anchor(spec, "FONT.SUB_HEADING_SIZE_PT")))
            run.font.bold = anchor(spec, "FONT.SUB_HEADING_BOLD", True)
            run.font.italic = anchor(spec, "FONT.SUB_HEADING_ITALIC", True)
            run.font.name = anchor(spec, "FONT.FAMILY")
            fixed.append(text[:40])
        elif kind == "body" and "FONT.BODY_SIZE_PT" in fixable:
            run.font.size = Pt(float(anchor(spec, "FONT.BODY_SIZE_PT")))
            run.font.name = anchor(spec, "FONT.FAMILY")
            fixed.append(text[:40])

    buffer = io.BytesIO()
    doc.save(buffer)
    return {
        "docx_b64": base64.b64encode(buffer.getvalue()).decode("ascii"),
        "fixed": fixed[:20],
    }
