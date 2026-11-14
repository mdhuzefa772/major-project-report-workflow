"""Extraction activities: guidelines PDF (Stream A), sample DOCX (Stream C), chapter plan."""

import base64
import io
import json
import os
import tempfile

import mistralai.workflows as workflows

from . import anchors as anchors_mod
from . import inspect_docx
from .llm import complete_json

GUIDELINES_SYSTEM = """You extract document-formatting rules for an academic report from an
institute's "Report Writing Guidelines" text. Map every rule you find onto this exact JSON schema
(omit keys the text does not specify — never guess):

{
  "PAGE": {"SIZE": str, "ORIENTATION": str,
           "MARGIN_LEFT_IN": number, "MARGIN_RIGHT_IN": number,
           "MARGIN_TOP_IN": number, "MARGIN_BOTTOM_IN": number},
  "FONT": {"FAMILY": str,
           "BODY_SIZE_PT": number, "BODY_BOLD": bool, "BODY_ITALIC": bool,
           "BODY_ALIGNMENT": "LEFT|RIGHT|CENTER|JUSTIFIED", "BODY_LINE_SPACING": number,
           "MAIN_HEADING_SIZE_PT": number, "MAIN_HEADING_BOLD": bool, "MAIN_HEADING_ALIGNMENT": str,
           "SUB_HEADING_SIZE_PT": number, "SUB_HEADING_BOLD": bool, "SUB_HEADING_ITALIC": bool,
           "SUB_HEADING_ALIGNMENT": str,
           "COVER_TITLE_SIZE_PT": number, "COVER_TITLE_BOLD": bool, "COVER_TITLE_ALIGNMENT": str},
  "TABLE": {"STYLE": str, "BORDER_TOP_VAL": str, "BORDER_TOP_SZ_EIGHTH": number,
            "BORDER_TOP_COLOR": str, "BORDER_BOTTOM_VAL": str, "BORDER_BOTTOM_SZ_EIGHTH": number,
            "BORDER_BOTTOM_COLOR": str, "HEADER_ROW_BOLD": bool, "CELL_ALIGNMENT": str,
            "CELL_SIZE_PT": number},
  "SUBSEQUENT_PAGES": {"HAS_HEADER": bool, "HAS_FOOTER": bool, "PAGE_NUMBER_DISPLAYED": bool,
                       "HEADER_SIZE_PT": number, "HEADER_ALIGNMENT": str,
                       "FOOTER_ALIGNMENT": str, "FOOTER_TYPE": str},
  "REFERENCES": {"FORMAT": str, "SEPARATE_FILE": bool, "INDENT_STYLE": str,
                 "SIZE_PT": number, "ALIGNMENT": str}
}

Rules: alignments must be one of LEFT/RIGHT/CENTER/JUSTIFIED. Font sizes are points.
Return ONLY the JSON object."""


@workflows.activity()
async def extract_guidelines_anchors(guidelines_pdf_b64: str) -> dict:
    """Stream A: extract formatting rules from the institute guidelines PDF."""
    from pypdf import PdfReader

    pdf_bytes = base64.b64decode(guidelines_pdf_b64)
    reader = PdfReader(io.BytesIO(pdf_bytes))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)

    extracted = await complete_json(GUIDELINES_SYSTEM, text[:20000])

    flat: dict[str, object] = {}
    for group, keys in extracted.items():
        if isinstance(keys, dict):
            for key, value in keys.items():
                if value is not None:
                    flat[f"{group}.{key}"] = value
    return {"stream_a_anchors": flat, "guidelines_text": text}


@workflows.activity()
async def inspect_sample_docx(docx_b64: str, kind: str) -> dict:
    """Stream C: run the inspection harness on a sample DOCX (kind: 'cover' | 'chapter')."""
    docx_bytes = base64.b64decode(docx_b64)
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, f"sample_{kind}.docx")
        with open(path, "wb") as handle:
            handle.write(docx_bytes)
        spec = inspect_docx.inspect(path)
    return {"kind": kind, "inspect": json.loads(json.dumps(spec, default=str))}


def _flatten_inspect_anchors(inspect_result: dict) -> dict:
    kind = inspect_result["kind"]
    return anchors_mod.anchors_from_inspect(inspect_result["inspect"], kind)


@workflows.activity()
async def reconcile_spec(
    stream_a_anchors: dict,
    cover_inspect: dict,
    chapter_inspect: dict,
    project: dict,
    content_mode: str,
    chapter_list: list[str] | None = None,
) -> dict:
    """Merge Stream A + Stream C into the frozen anchor spec with a conflict log."""
    stream_c: dict[str, object] = {}
    for inspect_result in (cover_inspect, chapter_inspect):
        stream_c.update(_flatten_inspect_anchors(inspect_result))

    spec, conflicts, missing = anchors_mod.reconcile(
        stream_a_anchors, stream_c, project, content_mode, chapter_list
    )
    return {"spec": spec, "conflicts": conflicts, "missing": missing}


PLAN_SYSTEM = """You plan the chapter structure of an academic project report. Given the
chapter list, the source material summary, and the project details, produce a plan where each
chapter has a numbered section outline. Use the source material's real content — never invent
algorithms, datasets, metrics, or figures that are not present in the inputs.

Return ONLY JSON:
{"chapters": [{"chapter_num": 1, "title": "INTRODUCTION",
  "intro_paragraph_hint": str,
  "sections": [{"number": "1.1", "heading": str,
                "subsections": [{"number": "1.1.1", "heading": str}],
                "content_hint": str,
                "needs_table": bool, "needs_figure_ref": bool}]}]}"""


@workflows.activity()
async def plan_chapters(
    spec: dict, source_content_text: str, project_details: dict
) -> dict:
    """Decide the chapter/section outline (structural precedence: user > guidelines > default)."""
    chapter_list = spec["PROJECT"]["CHAPTER_LIST"]
    plan = await complete_json(
        PLAN_SYSTEM,
        json.dumps({
            "report_type": spec["PROJECT"]["TYPE"],
            "chapter_list": chapter_list,
            "project": {k: spec["PROJECT"][k] for k in ("NAME", "INSTITUTE", "DEPARTMENT", "ACADEMIC_YEAR")},
            "project_details": project_details,
            "source_material": source_content_text[:24000],
        }, indent=1),
    )
    return plan


@workflows.activity()
async def extract_source_text(source_docx_b64: str) -> dict:
    """Stream B: read the user's source DOCX content (structure, not formatting)."""
    from docx import Document

    docx_bytes = base64.b64decode(source_docx_b64)
    document = Document(io.BytesIO(docx_bytes))
    lines = []
    for para in document.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)
    for table in document.tables:
        for row in table.rows:
            lines.append(" | ".join(cell.text.strip() for cell in row.cells))
    return {"source_text": "\n".join(lines)}
