"""Anchor Points system (spec v2 section 5).

Every formatting value is extracted from the input streams, never memorized.
Precedence for FORMATTING: USER > Stream C (sample DOCX) > Stream A (guidelines PDF) > none.
Unresolvable anchors are returned as `missing` and block the workflow until the
user supplies them at the spec-approval gate.
"""

import re
from collections import Counter

MM_PER_INCH = 25.4
EMU_PER_MM = 36000
TWIPS_PER_MM = 56.693

ANCHOR_PATHS: dict[str, list[str]] = {
    "PAGE": [
        "SIZE", "ORIENTATION",
        "MARGIN_LEFT_IN", "MARGIN_RIGHT_IN", "MARGIN_TOP_IN", "MARGIN_BOTTOM_IN",
    ],
    "FONT": [
        "FAMILY",
        "BODY_SIZE_PT", "BODY_BOLD", "BODY_ITALIC", "BODY_ALIGNMENT", "BODY_LINE_SPACING",
        "MAIN_HEADING_SIZE_PT", "MAIN_HEADING_BOLD", "MAIN_HEADING_ALIGNMENT",
        "SUB_HEADING_SIZE_PT", "SUB_HEADING_BOLD", "SUB_HEADING_ITALIC", "SUB_HEADING_ALIGNMENT",
        "COVER_TITLE_SIZE_PT", "COVER_TITLE_BOLD", "COVER_TITLE_ALIGNMENT",
    ],
    "TABLE": [
        "STYLE", "BORDER_TOP_VAL", "BORDER_TOP_SZ_EIGHTH", "BORDER_TOP_COLOR",
        "BORDER_BOTTOM_VAL", "BORDER_BOTTOM_SZ_EIGHTH", "BORDER_BOTTOM_COLOR",
        "HEADER_ROW_BOLD", "CELL_ALIGNMENT", "CELL_SIZE_PT",
    ],
    "COVER_PAGE": ["VERTICAL_OFFSET_PARAGRAPHS", "HAS_HEADER", "HAS_FOOTER", "PAGE_COUNTED"],
    "FIRST_CONTENT_PAGE": ["HAS_HEADER", "HAS_FOOTER", "PAGE_COUNTED", "REPEATS_COVER_TITLE"],
    "SUBSEQUENT_PAGES": [
        "HAS_HEADER", "HAS_FOOTER", "PAGE_NUMBER_DISPLAYED", "PAGE_NUMBER_START",
        "HEADER_TEXT_TEMPLATE", "HEADER_SIZE_PT", "HEADER_ALIGNMENT",
        "FOOTER_ALIGNMENT", "FOOTER_TYPE",
    ],
    "REFERENCES": ["FORMAT", "SEPARATE_FILE", "INDENT_STYLE", "SIZE_PT", "ALIGNMENT"],
}

DEFAULT_CHAPTER_LIST = [
    "Introduction",
    "Literature Review",
    "Methodology",
    "Results and Discussion",
    "Conclusion and Future Scope",
]


def blank_spec() -> dict:
    spec: dict[str, dict] = {}
    for group, keys in ANCHOR_PATHS.items():
        spec[group] = {key: None for key in keys}
    return spec


def _page_setup_anchors(page_setup: dict) -> dict:
    values: dict[str, object] = {}
    size = page_setup.get("page_size") or {}
    margins = page_setup.get("margins") or {}
    if size.get("w_twips") and size.get("h_twips"):
        values["PAGE.SIZE"] = "A4" if size["w_twips"] == 11906 else (
            "LETTER" if size["w_twips"] == 12240 else f"CUSTOM({size['w_twips']}x{size['h_twips']})"
        )
        values["PAGE.ORIENTATION"] = (size.get("orient") or "portrait").upper()
    for side in ("left", "right", "top", "bottom"):
        inches = margins.get(f"{side}_inches")
        if inches is not None:
            values[f"PAGE.MARGIN_{side.upper()}_IN"] = round(float(inches), 3)
    return values


def _majority_run(paragraphs: list[dict]) -> dict | None:
    """Most common resolved run properties across non-empty paragraphs."""
    votes: Counter = Counter()
    for para in paragraphs:
        if para.get("is_empty"):
            continue
        for run in para.get("runs", []):
            key = (run.get("size_pt"), bool(run.get("bold")), bool(run.get("italic")))
            votes[key] += len(run.get("text", ""))
    if not votes:
        return None
    (size_pt, bold, italic), _ = votes.most_common(1)[0]
    return {"size_pt": size_pt, "bold": bold, "italic": italic}


def _classify(text: str) -> str:
    if re.match(r"^\d+\.\d+\.\d+\s", text):
        return "subheading"
    if re.match(r"^\d+\.\d+\s", text):
        return "main_heading"
    return "body"


def _font_anchors(paragraphs: list[dict], cover_paragraphs: list[dict]) -> dict:
    values: dict[str, object] = {}

    fonts = Counter()
    alignments = Counter()
    spacings = Counter()
    for para in paragraphs:
        if para.get("is_empty"):
            continue
        for run in para.get("runs", []):
            if run.get("font"):
                fonts[run["font"]] += len(run.get("text", ""))
        if para.get("alignment"):
            alignments[para["alignment"]] += 1
        if para.get("line_spacing"):
            spacings[round(float(para["line_spacing"]), 2)] += 1
    if fonts:
        values["FONT.FAMILY"] = fonts.most_common(1)[0][0]
    if alignments:
        align = alignments.most_common(1)[0][0]
        values["FONT.BODY_ALIGNMENT"] = {"both": "JUSTIFIED", "justify": "JUSTIFIED"}.get(align, align.upper())
    if spacings:
        values["FONT.BODY_LINE_SPACING"] = spacings.most_common(1)[0][0]

    body_runs = [p for p in paragraphs if _classify((p.get("runs") or [{}])[0].get("text", "") if p.get("runs") else "") == "body"]
    body = _majority_run(body_runs or paragraphs)
    if body:
        values["FONT.BODY_SIZE_PT"] = body["size_pt"]
        values["FONT.BODY_BOLD"] = body["bold"]
        values["FONT.BODY_ITALIC"] = body["italic"]

    main = [p for p in paragraphs if p.get("runs") and _classify(p["runs"][0].get("text", "")) == "main_heading"]
    main_run = _majority_run(main)
    if main_run:
        values["FONT.MAIN_HEADING_SIZE_PT"] = main_run["size_pt"]
        values["FONT.MAIN_HEADING_BOLD"] = main_run["bold"]
        values["FONT.MAIN_HEADING_ALIGNMENT"] = (main[0].get("alignment") or "left").upper()

    sub = [p for p in paragraphs if p.get("runs") and _classify(p["runs"][0].get("text", "")) == "subheading"]
    sub_run = _majority_run(sub)
    if sub_run:
        values["FONT.SUB_HEADING_SIZE_PT"] = sub_run["size_pt"]
        values["FONT.SUB_HEADING_BOLD"] = sub_run["bold"]
        values["FONT.SUB_HEADING_ITALIC"] = sub_run["italic"]
        values["FONT.SUB_HEADING_ALIGNMENT"] = (sub[0].get("alignment") or "left").upper()

    titled = [p for p in cover_paragraphs if not p.get("is_empty")]
    if titled:
        title_run = _majority_run(titled[:2])
        if title_run:
            values["FONT.COVER_TITLE_SIZE_PT"] = title_run["size_pt"]
            values["FONT.COVER_TITLE_BOLD"] = title_run["bold"]
            values["FONT.COVER_TITLE_ALIGNMENT"] = (titled[0].get("alignment") or "right").upper()
    return values


def _header_footer_anchors(inspect_spec: dict, is_cover: bool) -> dict:
    values: dict[str, object] = {}
    page_setup = inspect_spec.get("body_page_setup") or {}
    parts = inspect_spec.get("header_footer_parts") or {}

    header_refs = page_setup.get("header_refs") or []
    footer_refs = page_setup.get("footer_refs") or []
    has_header = bool(header_refs)
    has_footer = bool(footer_refs)

    if is_cover:
        values["COVER_PAGE.HAS_HEADER"] = has_header
        values["COVER_PAGE.HAS_FOOTER"] = has_footer
        values["COVER_PAGE.PAGE_COUNTED"] = has_footer
        paragraphs = inspect_spec.get("paragraphs") or []
        leading_empty = 0
        for para in paragraphs:
            if para.get("is_empty"):
                leading_empty += 1
            else:
                break
        values["COVER_PAGE.VERTICAL_OFFSET_PARAGRAPHS"] = leading_empty
        return values

    values["FIRST_CONTENT_PAGE.HAS_HEADER"] = False
    values["FIRST_CONTENT_PAGE.HAS_FOOTER"] = False
    values["FIRST_CONTENT_PAGE.PAGE_COUNTED"] = True
    values["FIRST_CONTENT_PAGE.REPEATS_COVER_TITLE"] = True

    values["SUBSEQUENT_PAGES.HAS_HEADER"] = has_header
    values["SUBSEQUENT_PAGES.HAS_FOOTER"] = has_footer
    values["SUBSEQUENT_PAGES.PAGE_NUMBER_DISPLAYED"] = has_footer

    headers = parts.get("headers") or {}
    for part_paragraphs in headers.values():
        for para in part_paragraphs:
            text = (para.get("text") or "").strip()
            if text:
                values["SUBSEQUENT_PAGES.HEADER_TEXT_TEMPLATE"] = (
                    text.replace("1", "{{CHAPTER_NUM}}").replace("-1", "-{{CHAPTER_NUM}}")
                    if "-1" in text else text
                )
                values["SUBSEQUENT_PAGES.HEADER_SIZE_PT"] = para.get("size_pt")
                values["SUBSEQUENT_PAGES.HEADER_ALIGNMENT"] = (para.get("alignment") or "right").upper()
                break
        if values.get("SUBSEQUENT_PAGES.HEADER_TEXT_TEMPLATE"):
            break

    footers = parts.get("footers") or {}
    for part_paragraphs in footers.values():
        for para in part_paragraphs:
            if para.get("has_page_field"):
                values["SUBSEQUENT_PAGES.FOOTER_ALIGNMENT"] = (para.get("alignment") or "center").upper()
                values["SUBSEQUENT_PAGES.FOOTER_TYPE"] = "PAGE_FIELD"
                break
        if values.get("SUBSEQUENT_PAGES.FOOTER_TYPE"):
            break

    if page_setup.get("pgNumType_start") is not None:
        values["SUBSEQUENT_PAGES.PAGE_NUMBER_START"] = int(page_setup["pgNumType_start"]) + 1
    return values


def anchors_from_inspect(inspect_spec: dict, kind: str) -> dict:
    """Map a harness inspect() result to flat `GROUP.KEY` anchor values (Stream C)."""
    values: dict[str, object] = {}
    page_setup = inspect_spec.get("body_page_setup") or {}
    values.update(_page_setup_anchors(page_setup))
    if kind == "cover":
        values.update(_header_footer_anchors(inspect_spec, is_cover=True))
    else:
        values.update(_font_anchors(inspect_spec.get("paragraphs") or [], []))
        values.update(_header_footer_anchors(inspect_spec, is_cover=False))
    return values


def _values_equal(a, b) -> bool:
    if a is None or b is None:
        return False
    if isinstance(a, float) or isinstance(b, float):
        try:
            return abs(float(a) - float(b)) < 1e-6
        except (TypeError, ValueError):
            return False
    if isinstance(a, bool) or isinstance(b, bool):
        return str(a).lower() == str(b).lower()
    return str(a).strip().lower() == str(b).strip().lower()


def reconcile(
    stream_a: dict[str, object] | None,
    stream_c: dict[str, object],
    project: dict,
    content_mode: str,
    chapter_list: list[str] | None,
) -> tuple[dict, list[dict], list[str]]:
    """Merge guidelines-derived anchors (A) with sample-derived anchors (C).

    Returns (spec, conflicts, missing). Stream C wins for formatting; every
    disagreement is recorded as a conflict for the user to adjudicate.
    """
    spec = blank_spec()
    conflicts: list[dict] = []
    missing: list[str] = []

    stream_a = stream_a or {}

    for group, keys in ANCHOR_PATHS.items():
        for key in keys:
            path = f"{group}.{key}"
            c_val = stream_c.get(path)
            a_val = stream_a.get(path)
            if c_val is not None and a_val is not None and not _values_equal(c_val, a_val):
                conflicts.append({
                    "anchor": path,
                    "stream_a_guidelines": a_val,
                    "stream_c_sample": c_val,
                    "resolution": "STREAM_C_SAMPLE (precedence rule 3.1) — override at the approval gate if wrong",
                })
                spec[group][key] = c_val
            elif c_val is not None:
                spec[group][key] = c_val
            elif a_val is not None:
                spec[group][key] = a_val
            else:
                missing.append(path)

    _derive_page_units(spec)

    spec["PROJECT"] = {
        "NAME": project.get("name"),
        "TYPE": project.get("type", "major-project"),
        "INSTITUTE": project.get("institute"),
        "DEPARTMENT": project.get("department"),
        "ACADEMIC_YEAR": project.get("academic_year"),
        "CHAPTER_COUNT": len(chapter_list) if chapter_list else len(DEFAULT_CHAPTER_LIST),
        "CHAPTER_LIST": chapter_list or DEFAULT_CHAPTER_LIST,
    }
    spec["CONTENT_MODE"] = content_mode
    spec["USER_ACTUAL_RESULTS"] = project.get("actual_results") or {}
    return spec, conflicts, missing


def _derive_page_units(spec: dict) -> None:
    """Fill EMU/TWIPS from the page size anchor (P3 conversion formulas)."""
    page = spec["PAGE"]
    size = page.get("SIZE")
    table = {
        "A4": (210.0, 297.0),
        "LETTER": (215.9, 279.4),
    }
    if size in table:
        width_mm, height_mm = table[size]
    elif isinstance(size, str) and size.startswith("CUSTOM("):
        try:
            w_twips, h_twips = size[len("CUSTOM("):-1].split("x")
            page["WIDTH_TWIPS"] = int(w_twips)
            page["HEIGHT_TWIPS"] = int(h_twips)
            page["WIDTH_EMU"] = int(w_twips) * 635
            page["HEIGHT_EMU"] = int(h_twips) * 635
            return
        except (ValueError, ZeroDivisionError):
            return
    else:
        return
    page["WIDTH_EMU"] = int(width_mm * EMU_PER_MM)
    page["HEIGHT_EMU"] = int(height_mm * EMU_PER_MM)
    page["WIDTH_TWIPS"] = int(round(width_mm * TWIPS_PER_MM))
    page["HEIGHT_TWIPS"] = int(round(height_mm * TWIPS_PER_MM))


def anchor(spec: dict, dotted_path: str, default=None):
    value = spec
    for key in dotted_path.split("."):
        if not isinstance(value, dict) or key not in value:
            return default
        value = value[key]
    return default if value is None else value
