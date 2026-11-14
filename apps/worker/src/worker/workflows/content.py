"""Content activities: chapter drafting (mode-aware), accuracy gate, word-count rule."""

import json
import re

import mistralai.workflows as workflows

from .llm import complete_json

DRAFT_SYSTEM = """You write one chapter of an academic project report (B.E./B.Tech. Major
Project / Technical Seminar style). Absolute rules:

1. FACTS: use ONLY the algorithms, datasets, metrics, figures and tables supplied in
   PROJECT_DETAILS and SOURCE_MATERIAL. Never invent or import facts from anywhere else —
   not from other projects, not from your training data. If a required fact is missing,
   put a placeholder of the form [NEEDS USER INPUT: what is missing] instead of inventing.
2. DEPTH: every section needs flowing academic prose. Paragraphs are 3-5 sentences.
3. STRUCTURE: sections are numbered N.1, N.2 ...; subsections N.1.1 ... There must be at
   least one introductory paragraph between the chapter title and the first numbered heading.
4. FIGURES/TABLES: reference figures as "Fig. N.M <title>" captions and tables as
   "Table N.M <title>" captions, numbered per chapter. Only reference figures/tables that
   exist in the inputs.
5. STYLE: formal third-person academic English, no bullets inside body paragraphs.

Return ONLY JSON:
{"chapter_num": int, "title": str,
 "intro_paragraph": str,
 "sections": [{"number": "1.1", "heading": str, "paragraphs": [str],
   "bullets": [str],
   "subsections": [{"number": "1.1.1", "heading": str, "paragraphs": [str]}],
   "tables": [{"caption": "Table 1.1 ...", "header_row": [str], "rows": [[str]]}],
   "figure_captions": ["Fig. 1.1 ..."]}]}

CONTENT_MODE semantics:
- REFORMAT_ONLY: transcribe the user's prose from SOURCE_MATERIAL essentially verbatim,
  reorganized into the section structure. Do not add new claims.
- GENERATE: write full chapters from the source material and project details.
- EXPAND: transcribe what the user supplied and expand only the flagged sections
  (sections whose content_hint starts with "EXPAND:")."""

ACCURACY_SYSTEM = """You are a content-accuracy auditor for an academic report chapter.
Cross-reference EVERY factual claim in the chapter draft (algorithm names, model names,
datasets, metric names and values, figure numbers, table numbers, references) against the
PROJECT_DETAILS and SOURCE_MATERIAL. Report every claim that is not backed by them,
every metric value that differs, and every figure/table that does not exist in the inputs.
Do not flag stylistic issues.

Return ONLY JSON:
{"violations": [{"type": "fabricated_algorithm|fabricated_metric|metric_mismatch|fabricated_figure|fabricated_table|fabricated_reference|other",
                 "claim": str, "evidence": str, "suggested_fix": str}],
 "verified_claims_count": int}"""


@workflows.activity()
async def draft_chapter(
    chapter_plan: dict, source_content_text: str, spec: dict, project_details: dict
) -> dict:
    """Write one chapter's structured content, honoring the CONTENT_MODE switch."""
    draft = await complete_json(
        DRAFT_SYSTEM,
        json.dumps({
            "content_mode": spec["CONTENT_MODE"],
            "chapter_plan": chapter_plan,
            "project": spec["PROJECT"],
            "project_details": project_details,
            "source_material": source_content_text[:24000],
        }, indent=1),
        temperature=0.3,
    )
    draft["content_mode"] = spec["CONTENT_MODE"]
    return draft


@workflows.activity()
async def accuracy_gate(draft: dict, project_details: dict, source_content_text: str) -> dict:
    """Cross-reference every factual claim in the draft against the user's actual project."""
    audit = await complete_json(
        ACCURACY_SYSTEM,
        json.dumps({
            "chapter_draft": draft,
            "project_details": project_details,
            "source_material": source_content_text[:20000],
        }, indent=1),
        temperature=0.0,
    )
    audit["passed"] = not audit.get("violations")
    return audit


def _section_word_count(section: dict) -> int:
    words = 0
    for paragraph in section.get("paragraphs", []):
        words += len(re.findall(r"\S+", paragraph))
    for sub in section.get("subsections", []):
        for paragraph in sub.get("paragraphs", []):
            words += len(re.findall(r"\S+", paragraph))
    return words


@workflows.activity()
async def check_depth_rule(draft: dict, min_words_per_section: int = 150) -> dict:
    """Spec section 8.10 Rule 3, mode-aware (P1/P7).

    GENERATE/EXPAND: sections below the minimum are listed for expansion.
    REFORMAT_ONLY: short sections are reported for escalation, never rewritten.
    """
    mode = draft.get("content_mode", "GENERATE")
    short_sections = []
    for section in draft.get("sections", []):
        count = _section_word_count(section)
        if count < min_words_per_section:
            short_sections.append({"number": section.get("number"), "heading": section.get("heading"), "words": count})

    result = {
        "mode": mode,
        "min_words_per_section": min_words_per_section,
        "short_sections": short_sections,
    }
    if mode == "REFORMAT_ONLY":
        result["action"] = "ESCALATE_TO_USER" if short_sections else "OK"
    else:
        result["action"] = "EXPAND_SECTIONS" if short_sections else "OK"
    return result
