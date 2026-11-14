#!/usr/bin/env python3
"""
inspect_docx.py v2 — Dynamic document inspection harness.

CHANGES FROM v1 (based on DSH debrief):
  CRITICAL FIXES:
    1. Read body-level sectPr (line 171 bug) — not just paragraph-level.
       v1 looked for w:pPr/w:sectPr inside paragraphs; real documents put
       sectPr at body level. This made PAGE.SIZE, MARGIN_* unobtainable.
    2. Don't drop empty paragraphs (line 237 bug) — mark with is_empty=True.
       v1 silently discarded empty paras, violating §6.2's own rule and
       making paragraph indices non-authoritative.
    3. Fix w:orient crash (line 178 bug) — only int-cast w:w/w:h, treat
       w:orient and w:code as strings.

  NEW EXTRACTORS:
    4. theme1.xml reading — resolve inherited fonts (Calibri via theme,
       not "inherit"). This was the single most consequential missing
       extraction in the DSH engagement.
    5. docDefaults / style inheritance — resolve rPrDefault, pPrDefault,
       w:pStyle chain.
    6. Header/footer parts — parse word/header*.xml, word/footer*.xml
       and the sectPr references to them. Makes HEADER_*, FOOTER_*,
       titlePg, pgNumType anchors obtainable.
    7. Line spacing (w:spacing line/lineRule/before/after) — §5.3 anchor
       FONT.BODY_LINE_SPACING was unobtainable.
    8. Numbering / list format (w:numPr, numbering.xml) — distinguish
       bullets from decimals, detect shared-abstractNum trap.
    9. Table borders (tblBorders) — not just shading. Required for §20.4
       three-line format verification.
   10. Per-part sha256 fingerprint — so an agent can PROVE preservation
       instead of asserting it.
   11. --diff mode — compare two docx files for preservation checks.

  UX FIXES:
   12. Status messages to stdout (not stderr) — v1's stderr caused
       spurious exit code 1 on PowerShell.
   13. ASCII-only diagnostics (no ✓ → ─ glyphs) — avoids mojibake on
       cp1252 consoles.
   14. --preview made portable — caller-supplied output dir, optional
       renderer adapter.

USAGE:
  python3 inspect_docx.py input.docx
  python3 inspect_docx.py input.docx --json spec.json
  python3 inspect_docx.py input.docx --json spec.json --preview page1.png
  python3 inspect_docx.py --diff original.docx output.docx
"""

import sys
import os
import json
import zipfile
import argparse
import subprocess
import hashlib
from xml.etree import ElementTree as ET

# ----- XML namespace helpers -------------------------------------------------
W_NS   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS   = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
CT_NS  = "http://schemas.openxmlformats.org/package/2006/content-types"

NS = {"w": W_NS, "r": R_NS, "a": A_NS, "pic": PIC_NS}

def qn(tag):
    """Resolve a 'prefix:local' tag (or nested 'pfx:a/b/c') to '{ns}a/{ns}b/{ns}c'."""
    parts = tag.split("/")
    resolved = []
    for part in parts:
        prefix, local = part.split(":")
        resolved.append(f"{{{NS[prefix]}}}{local}")
    return "/".join(resolved)

def localname(elem):
    """Get the local name of an XML element (strip namespace)."""
    tag = elem.tag
    if "}" in tag:
        return tag.split("}")[1]
    return tag

def attr_int(elem, qname, default=None):
    """Get an attribute value as int, or default if not parseable."""
    v = elem.get(qname)
    if v is None:
        return default
    try:
        return int(v)
    except (ValueError, TypeError):
        return v  # return as string if not int (e.g., 'portrait')

# ----- readers ---------------------------------------------------------------
def read_docx(path):
    """Open the docx as a zip and return (document.xml, styles.xml, rels.xml,
    media_map, header_parts, footer_parts, theme_xml, numbering_xml, settings_xml,
    content_types_xml, all_parts_fingerprints)."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        doc_xml     = z.read("word/document.xml").decode("utf-8")
        styles_xml  = z.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in names else None
        rels_xml    = z.read("word/_rels/document.xml.rels").decode("utf-8")
        theme_xml   = z.read("word/theme/theme1.xml").decode("utf-8") if "word/theme/theme1.xml" in names else None
        numbering_xml = z.read("word/numbering.xml").decode("utf-8") if "word/numbering.xml" in names else None
        settings_xml  = z.read("word/settings.xml").decode("utf-8") if "word/settings.xml" in names else None
        ct_xml      = z.read("[Content_Types].xml").decode("utf-8")

        media = {n: z.read(n) for n in names if n.startswith("word/media/")}

        # Read all header/footer parts
        header_parts = {}
        footer_parts = {}
        for n in names:
            if n.startswith("word/header") and n.endswith(".xml"):
                header_parts[n] = z.read(n).decode("utf-8")
            elif n.startswith("word/footer") and n.endswith(".xml"):
                footer_parts[n] = z.read(n).decode("utf-8")

        # Compute sha256 fingerprint for every part in the package
        fingerprints = {}
        for n in names:
            fingerprints[n] = hashlib.sha256(z.read(n)).hexdigest()

    return (doc_xml, styles_xml, rels_xml, media,
            header_parts, footer_parts, theme_xml, numbering_xml, settings_xml,
            ct_xml, fingerprints)


def parse_rels(rels_xml):
    """Return {rId: {target, type}} from the relationships XML."""
    root = ET.fromstring(rels_xml)
    rels = {}
    for rel in root:
        rid   = rel.get("Id")
        tgt   = rel.get("Target")
        rtype = rel.get("Type", "").split("/")[-1]
        rels[rid] = {"target": tgt, "type": rtype}
    return rels


# ----- theme / style resolution ---------------------------------------------
def parse_theme_fonts(theme_xml):
    """Extract the major and minor font from theme1.xml.
    Returns {'major': 'Calibri', 'minor': 'Calibri'} or similar.
    This is CRITICAL — python-docx reports 'inherit' for runs that actually
    use the theme font, which may be Calibri even when guidelines say
    Times New Roman.
    """
    if not theme_xml:
        return {"major": None, "minor": None}
    root = ET.fromstring(theme_xml)
    # Theme namespace
    drawing_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    fonts = {}
    for kind in ["majorFont", "minorFont"]:
        elem = root.find(f".//{{{drawing_ns}}}{kind}")
        if elem is not None:
            latin = elem.find(f"{{{drawing_ns}}}latin")
            if latin is not None:
                fonts[kind.replace("Font", "").lower()] = latin.get("typeface")
    return fonts


def parse_doc_defaults(styles_xml):
    """Extract docDefaults (rPrDefault, pPrDefault) from styles.xml.
    These are the document-wide defaults that every 'inherited' run falls
    back to. Without this, the harness cannot tell you the body font size.
    """
    if not styles_xml:
        return {"rPr": {}, "pPr": {}}
    root = ET.fromstring(styles_xml)
    docDefaults = root.find(qn("w:docDefaults"))
    result = {"rPr": {}, "pPr": {}}
    if docDefaults is not None:
        rPrDefault = docDefaults.find(qn("w:rPrDefault"))
        if rPrDefault is not None:
            rPr = rPrDefault.find(qn("w:rPr"))
            if rPr is not None:
                result["rPr"] = parse_run_properties(rPr)
        pPrDefault = docDefaults.find(qn("w:pPrDefault"))
        if pPrDefault is not None:
            pPr = pPrDefault.find(qn("w:pPr"))
            if pPr is not None:
                result["pPr"] = parse_paragraph_properties(pPr)
    return result


def parse_style_chain(styles_xml, style_id):
    """Resolve a style by walking the w:basedOn chain.
    Returns the merged rPr/pPr from the style and all its parents.
    """
    if not styles_xml:
        return {"rPr": {}, "pPr": {}}
    root = ET.fromstring(styles_xml)
    # Find the style by w:styleId
    style_elem = None
    for s in root.iter(qn("w:style")):
        if s.get(qn("w:styleId")) == style_id:
            style_elem = s
            break
    if style_elem is None:
        return {"rPr": {}, "pPr": {}}
    # Walk basedOn chain
    merged_rPr = {}
    merged_pPr = {}
    chain = [style_elem]
    visited = set()
    while chain:
        current = chain.pop(0)
        sid = current.get(qn("w:styleId"))
        if sid in visited:
            break
        visited.add(sid)
        rPr = current.find(qn("w:rPr"))
        if rPr is not None:
            rprops = parse_run_properties(rPr)
            for k, v in rprops.items():
                if v is not None and k not in merged_rPr:
                    merged_rPr[k] = v
        pPr = current.find(qn("w:pPr"))
        if pPr is not None:
            pprops = parse_paragraph_properties(pPr)
            for k, v in pprops.items():
                if v is not None and k not in merged_pPr:
                    merged_pPr[k] = v
        based_on = current.find(qn("w:basedOn"))
        if based_on is not None:
            parent_id = based_on.get(qn("w:val"))
            for s in root.iter(qn("w:style")):
                if s.get(qn("w:styleId")) == parent_id:
                    chain.append(s)
                    break
    return {"rPr": merged_rPr, "pPr": merged_pPr}


# ----- paragraph / run extraction -------------------------------------------
def parse_run_properties(rpr):
    """Extract size, bold, italic, color, font from a <w:rPr> element."""
    spec = {"size_pt": None, "bold": False, "italic": False,
            "color": None, "font": None, "underline": False,
            "size_cs_pt": None}  # complex-script size
    if rpr is None:
        return spec

    sz = rpr.find(qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        try:
            spec["size_pt"] = int(sz.get(qn("w:val"))) // 2
        except ValueError:
            pass

    szCs = rpr.find(qn("w:szCs"))
    if szCs is not None and szCs.get(qn("w:val")):
        try:
            spec["size_cs_pt"] = int(szCs.get(qn("w:val"))) // 2
        except ValueError:
            pass

    if rpr.find(qn("w:b")) is not None:
        spec["bold"] = True
    if rpr.find(qn("w:i")) is not None:
        spec["italic"] = True
    if rpr.find(qn("w:u")) is not None:
        spec["underline"] = True

    color = rpr.find(qn("w:color"))
    if color is not None:
        spec["color"] = color.get(qn("w:val"))

    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        spec["font"] = (rfonts.get(qn("w:ascii"))
                        or rfonts.get(qn("w:hAnsi"))
                        or rfonts.get(qn("w:eastAsia")))
        # Also check for theme font references
        theme_major = rfonts.get(qn("w:asciiTheme"))
        theme_minor = rfonts.get(qn("w:hAnsiTheme"))
        if theme_major or theme_minor:
            spec["font_theme_ref"] = theme_major or theme_minor
    return spec


def parse_paragraph_properties(ppr):
    """Extract paragraph-level properties: alignment, spacing, indent, borders,
    style reference, numbering reference."""
    result = {
        "alignment": None,
        "line_spacing": None,
        "line_rule": None,
        "space_before_pt": None,
        "space_after_pt": None,
        "indent": None,
        "borders": [],
        "style_id": None,
        "num_pr": None,
    }
    if ppr is None:
        return result

    jc = ppr.find(qn("w:jc"))
    if jc is not None:
        result["alignment"] = jc.get(qn("w:val"))

    spacing = ppr.find(qn("w:spacing"))
    if spacing is not None:
        line = spacing.get(qn("w:line"))
        line_rule = spacing.get(qn("w:lineRule"))
        before = spacing.get(qn("w:before"))
        after = spacing.get(qn("w:after"))
        if line:
            try:
                line_val = int(line)
                # line=360 with lineRule=auto means 1.5x (240=1x)
                if line_rule == "auto" or line_rule is None:
                    result["line_spacing"] = round(line_val / 240, 2) if line_val else None
                else:
                    result["line_spacing"] = f"{line_val}twips({line_rule})"
                result["line_rule"] = line_rule or "auto"
            except ValueError:
                pass
        if before:
            try:
                result["space_before_pt"] = round(int(before) / 20, 1)
            except ValueError:
                pass
        if after:
            try:
                result["space_after_pt"] = round(int(after) / 20, 1)
            except ValueError:
                pass

    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        result["indent"] = {}
        for k, v in ind.attrib.items():
            # k is a string like '{ns}left' — strip namespace
            if "}" in k:
                key = k.split("}")[1]
            else:
                key = k
            try:
                result["indent"][key] = int(v)
            except (ValueError, TypeError):
                result["indent"][key] = v

    pStyle = ppr.find(qn("w:pStyle"))
    if pStyle is not None:
        result["style_id"] = pStyle.get(qn("w:val"))

    numPr = ppr.find(qn("w:numPr"))
    if numPr is not None:
        numId_elem = numPr.find(qn("w:numId"))
        ilvl_elem = numPr.find(qn("w:ilvl"))
        result["num_pr"] = {
            "numId": numId_elem.get(qn("w:val")) if numId_elem is not None else None,
            "ilvl": ilvl_elem.get(qn("w:val")) if ilvl_elem is not None else "0",
        }

    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is not None:
        for b in pbdr:
            edge = localname(b)
            result["borders"].append({
                "edge": edge,
                "style": b.get(qn("w:val")),
                "size_eighth_pt": int(b.get(qn("w:sz")) or 0) / 8,
                "color": b.get(qn("w:color")),
            })

    return result


def resolve_run_props(run_rpr, doc_defaults, style_chain_cache, styles_xml, theme_fonts):
    """Resolve 'inherited' values by walking: run rPr -> paragraph style ->
    docDefaults -> theme1.xml. This is what v1 was MISSING — it reported
    'inherited' for everything, concealing the actual body font/size.
    """
    resolved = dict(run_rpr)  # copy
    resolved["resolved_from"] = []

    # If font is None, check style chain, then docDefaults, then theme
    if not resolved.get("font"):
        # Try style chain
        for style_props in style_chain_cache.values():
            if style_props.get("rPr", {}).get("font"):
                resolved["font"] = style_props["rPr"]["font"]
                resolved["resolved_from"].append("style_chain")
                break
        if not resolved.get("font"):
            # Try docDefaults
            dd_font = doc_defaults.get("rPr", {}).get("font")
            if dd_font:
                resolved["font"] = dd_font
                resolved["resolved_from"].append("docDefaults_rPr")
            else:
                # Check theme ref
                theme_ref = doc_defaults.get("rPr", {}).get("font_theme_ref")
                if theme_ref == "minorBidi" or theme_ref == "minorHAnsi":
                    resolved["font"] = theme_fonts.get("minor")
                    resolved["resolved_from"].append("theme1.xml:minorFont")
                elif theme_ref == "majorBidi" or theme_ref == "majorHAnsi":
                    resolved["font"] = theme_fonts.get("major")
                    resolved["resolved_from"].append("theme1.xml:majorFont")
                elif theme_ref:
                    # Direct theme reference
                    resolved["font"] = theme_fonts.get("minor") or theme_fonts.get("major")
                    resolved["resolved_from"].append(f"theme1.xml:{theme_ref}")

    # If size_pt is None, check style chain, then docDefaults
    if resolved.get("size_pt") is None:
        for style_props in style_chain_cache.values():
            if style_props.get("rPr", {}).get("size_pt") is not None:
                resolved["size_pt"] = style_props["rPr"]["size_pt"]
                resolved["resolved_from"].append("style_chain")
                break
        if resolved.get("size_pt") is None:
            dd_size = doc_defaults.get("rPr", {}).get("size_pt")
            if dd_size is not None:
                resolved["size_pt"] = dd_size
                resolved["resolved_from"].append("docDefaults_rPr")

    # If color is None, check style chain, then docDefaults
    if not resolved.get("color"):
        for style_props in style_chain_cache.values():
            if style_props.get("rPr", {}).get("color"):
                resolved["color"] = style_props["rPr"]["color"]
                resolved["resolved_from"].append("style_chain")
                break
        if not resolved.get("color"):
            dd_color = doc_defaults.get("rPr", {}).get("color")
            if dd_color:
                resolved["color"] = dd_color
                resolved["resolved_from"].append("docDefaults_rPr")

    # If bold not set, check style chain
    if not resolved.get("bold"):
        for style_props in style_chain_cache.values():
            if style_props.get("rPr", {}).get("bold"):
                resolved["bold"] = True
                resolved["resolved_from"].append("style_chain")
                break

    return resolved


def parse_paragraph(p_elem, rels, doc_defaults, style_chain_cache,
                    styles_xml, theme_fonts):
    """Return a structured spec for one <w:p> element, with resolved properties."""
    runs_spec = []

    # First, get paragraph-level properties (including style_id for chain resolution)
    ppr = p_elem.find(qn("w:pPr"))
    para_props = parse_paragraph_properties(ppr)

    # Build style chain cache for this paragraph's style
    if para_props["style_id"] and para_props["style_id"] not in style_chain_cache:
        style_chain_cache[para_props["style_id"]] = parse_style_chain(
            styles_xml, para_props["style_id"])

    for r in p_elem.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        raw_props = parse_run_properties(rpr)

        # Resolve inherited values
        resolved_props = resolve_run_props(
            raw_props, doc_defaults, style_chain_cache, styles_xml, theme_fonts)

        # Keep raw values too for transparency
        resolved_props["_raw"] = raw_props

        # Collect text fragments inside this run
        text_parts = []
        for t in r.iter(qn("w:t")):
            if t.text:
                text_parts.append(t.text)

        # Detect embedded images (drawings/blips)
        images = []
        for blip in r.iter(qn("a:blip")):
            embed = blip.get(qn("r:embed"))
            if embed and embed in rels:
                images.append({"rId": embed, "target": rels[embed]["target"]})

        run_text = "".join(text_parts)
        if images:
            resolved_props["images"] = images
            if not run_text:
                run_text = f"[IMAGE: {', '.join(i['target'] for i in images)}]"

        resolved_props["text"] = run_text
        runs_spec.append(resolved_props)

    # Section-level page setup — CRITICAL FIX: check body-level sectPr too
    sect_pr = None
    if ppr is not None:
        sect_pr = ppr.find(qn("w:sectPr"))
    # Note: body-level sectPr is handled separately in inspect()

    # Check if this paragraph is empty (has no runs or all runs are empty)
    is_empty = True
    for run in runs_spec:
        if run.get("text", "").strip() or "images" in run:
            is_empty = False
            break
    # Also count truly empty paragraphs (no runs at all)
    if not p_elem.findall(qn("w:r")):
        is_empty = True

    return {
        "runs": runs_spec,
        "is_empty": is_empty,
        "alignment": para_props["alignment"],
        "line_spacing": para_props["line_spacing"],
        "line_rule": para_props["line_rule"],
        "space_before_pt": para_props["space_before_pt"],
        "space_after_pt": para_props["space_after_pt"],
        "indent": para_props["indent"],
        "borders": para_props["borders"],
        "style_id": para_props["style_id"],
        "num_pr": para_props["num_pr"],
        "has_sectPr": sect_pr is not None,
    }


def parse_tables(body):
    """Quick table inventory with borders (not just shading).
    Uses direct findall to avoid nested-table row collection."""
    tables = []
    for tbl in body.iter(qn("w:tbl")):
        rows = []
        # Use findall on direct children only (not iter) to avoid nested tables
        for tr in tbl.findall(qn("w:tr")):
            cells = []
            for tc in tr.findall(qn("w:tc")):
                texts = []
                for t in tc.iter(qn("w:t")):
                    if t.text:
                        texts.append(t.text)
                tcpr = tc.find(qn("w:tcPr"))
                shading = None
                borders = {}
                if tcpr is not None:
                    shd = tcpr.find(qn("w:shd"))
                    if shd is not None:
                        shading = shd.get(qn("w:fill"))
                    tcBorders = tcpr.find(qn("w:tcBorders"))
                    if tcBorders is not None:
                        for b in tcBorders:
                            edge = localname(b)
                            borders[edge] = {
                                "val": b.get(qn("w:val")),
                                "sz": b.get(qn("w:sz")),
                                "color": b.get(qn("w:color")),
                            }
                grid_span = None
                if tcpr is not None:
                    gs = tcpr.find(qn("w:gridSpan"))
                    if gs is not None:
                        grid_span = int(gs.get(qn("w:val")))
                cells.append({
                    "text_preview": " ".join(texts)[:80],
                    "shading_fill": shading,
                    "cell_borders": borders,
                    "grid_span": grid_span,
                })
            rows.append({
                "cells": cells,
                "cant_split": tr.find(qn("w:trPr/w:cantSplit")) is not None,
            })

        # Table-level borders
        tblPr = tbl.find(qn("w:tblPr"))
        tbl_borders = {}
        if tblPr is not None:
            tblBorders = tblPr.find(qn("w:tblBorders"))
            if tblBorders is not None:
                for b in tblBorders:
                    edge = localname(b)
                    tbl_borders[edge] = {
                        "val": b.get(qn("w:val")),
                        "sz": b.get(qn("w:sz")),
                        "color": b.get(qn("w:color")),
                    }

        tables.append({
            "row_count": len(rows),
            "rows_preview": rows[:5],
            "table_borders": tbl_borders,
            "is_three_line": _is_three_line_table(tbl_borders, rows),
        })
    return tables


def _is_three_line_table(tbl_borders, rows):
    """Check if a table uses three-line academic format:
    top border + bottom border + header row bottom border, no other borders.
    """
    if not tbl_borders:
        return False
    has_top = tbl_borders.get("top", {}).get("val") in ("single", "thick")
    has_bottom = tbl_borders.get("bottom", {}).get("val") in ("single", "thick")
    no_vertical = (tbl_borders.get("left", {}).get("val") in (None, "none", "nil")
                   and tbl_borders.get("right", {}).get("val") in (None, "none", "nil")
                   and tbl_borders.get("insideV", {}).get("val") in (None, "none", "nil"))
    return has_top and has_bottom and no_vertical


def parse_header_footer_parts(header_parts, footer_parts):
    """Parse header*.xml and footer*.xml to extract their content.
    This makes HEADER_TEXT_TEMPLATE, HEADER_SIZE_PT, HEADER_ALIGNMENT,
    FOOTER_TYPE, etc. anchors obtainable.
    """
    result = {"headers": {}, "footers": {}}
    for path, xml in header_parts.items():
        root = ET.fromstring(xml)
        paragraphs = []
        for p in root.iter(qn("w:p")):
            text = "".join(t.text or "" for t in p.iter(qn("w:t")))
            ppr = p.find(qn("w:pPr"))
            props = parse_paragraph_properties(ppr)
            first_run = p.find(qn("w:r"))
            rpr = first_run.find(qn("w:rPr")) if first_run is not None else None
            run_props = parse_run_properties(rpr) if rpr is not None else {}
            # Check for PAGE field
            has_page_field = any(
                instr.text and "PAGE" in instr.text
                for instr in p.iter(qn("w:instrText"))
            )
            paragraphs.append({
                "text": text,
                "alignment": props["alignment"],
                "size_pt": run_props.get("size_pt"),
                "bold": run_props.get("bold", False),
                "font": run_props.get("font"),
                "has_page_field": has_page_field,
            })
        result["headers"][path] = paragraphs

    for path, xml in footer_parts.items():
        root = ET.fromstring(xml)
        paragraphs = []
        for p in root.iter(qn("w:p")):
            text = "".join(t.text or "" for t in p.iter(qn("w:t")))
            ppr = p.find(qn("w:pPr"))
            props = parse_paragraph_properties(ppr)
            first_run = p.find(qn("w:r"))
            rpr = first_run.find(qn("w:rPr")) if first_run is not None else None
            run_props = parse_run_properties(rpr) if rpr is not None else {}
            has_page_field = any(
                instr.text and "PAGE" in instr.text
                for instr in p.iter(qn("w:instrText"))
            )
            paragraphs.append({
                "text": text,
                "alignment": props["alignment"],
                "size_pt": run_props.get("size_pt"),
                "bold": run_props.get("bold", False),
                "font": run_props.get("font"),
                "has_page_field": has_page_field,
            })
        result["footers"][path] = paragraphs

    return result


def parse_numbering(numbering_xml):
    """Parse numbering.xml to detect list formats and shared-abstractNum traps.
    Returns a list of abstractNum definitions with their numFmt, lvlText, etc.
    Also detects when two w:num entries share the same w:abstractNum (which
    causes list numbering to continue across what should be separate lists).
    """
    if not numbering_xml:
        return {"abstract_nums": [], "nums": [], "shared_abstractNum_warnings": []}
    root = ET.fromstring(numbering_xml)
    abstract_nums = []
    for an in root.iter(qn("w:abstractNum")):
        anId = an.get(qn("w:abstractNumId"))
        nsid_elem = an.find(qn("w:nsid"))
        nsid = nsid_elem.get(qn("w:val")) if nsid_elem is not None else None
        levels = []
        for lvl in an.findall(qn("w:lvl")):
            numFmt = lvl.find(qn("w:numFmt"))
            lvlText = lvl.find(qn("w:lvlText"))
            levels.append({
                "ilvl": lvl.get(qn("w:ilvl")),
                "numFmt": numFmt.get(qn("w:val")) if numFmt is not None else None,
                "lvlText": lvlText.get(qn("w:val")) if lvlText is not None else None,
            })
        abstract_nums.append({
            "abstractNumId": anId,
            "nsid": nsid,
            "levels": levels,
        })

    nums = []
    num_to_an = {}  # numId -> abstractNumId
    an_to_nums = {}  # abstractNumId -> [numIds]
    for n in root.iter(qn("w:num")):
        numId = n.get(qn("w:numId"))
        an_ref = n.find(qn("w:abstractNumId"))
        anId = an_ref.get(qn("w:val")) if an_ref is not None else None
        nums.append({"numId": numId, "abstractNumId": anId})
        num_to_an[numId] = anId
        an_to_nums.setdefault(anId, []).append(numId)

    # Warn about shared abstractNum (the trap DSH hit: list continued 11-15 instead of 1-5)
    shared_warnings = []
    for anId, numIds in an_to_nums.items():
        if len(numIds) > 1:
            # Check if nsids are duplicated
            nsids = [an.get("nsid") for an in abstract_nums if an["abstractNumId"] == anId]
            shared_warnings.append({
                "abstractNumId": anId,
                "shared_by_numIds": numIds,
                "warning": f"numIds {numIds} share abstractNumId {anId}. "
                           f"Lists will continue each other's numbering unless "
                           f"each has a unique abstractNum with a unique nsid.",
            })

    return {
        "abstract_nums": abstract_nums,
        "nums": nums,
        "shared_abstractNum_warnings": shared_warnings,
    }


# ----- top-level inspect -----------------------------------------------------
def inspect(path):
    (doc_xml, styles_xml, rels_xml, media,
     header_parts, footer_parts, theme_xml, numbering_xml, settings_xml,
     ct_xml, fingerprints) = read_docx(path)

    rels = parse_rels(rels_xml)
    theme_fonts = parse_theme_fonts(theme_xml)
    doc_defaults = parse_doc_defaults(styles_xml)

    root = ET.fromstring(doc_xml)
    body = root.find(qn("w:body"))

    # CRITICAL FIX: read body-level sectPr (not just paragraph-level)
    # In virtually every single-section document, sectPr lives at body level
    body_sect_pr = body.find(qn("w:sectPr"))
    body_page_setup = None
    body_header_footer_refs = []
    if body_sect_pr is not None:
        body_page_setup = _parse_sect_pr(body_sect_pr)
        for ref in body_sect_pr.findall(qn("w:headerReference")):
            body_header_footer_refs.append({
                "type": "header",
                "ref_type": ref.get(qn("w:type")),
                "rId": ref.get(qn("r:id")),
            })
        for ref in body_sect_pr.findall(qn("w:footerReference")):
            body_header_footer_refs.append({
                "type": "footer",
                "ref_type": ref.get(qn("w:type")),
                "rId": ref.get(qn("r:id")),
            })

    style_chain_cache = {}
    paragraphs = []
    for p in body.iter(qn("w:p")):
        # Skip paragraphs that are inside tables (they'll be counted in tables)
        if _is_inside_table(p, body):
            continue
        parsed = parse_paragraph(p, rels, doc_defaults, style_chain_cache,
                                 styles_xml, theme_fonts)
        # Check for paragraph-level sectPr
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            para_sect_pr = ppr.find(qn("w:sectPr"))
            if para_sect_pr is not None:
                parsed["sect_pr"] = _parse_sect_pr(para_sect_pr)
        paragraphs.append(parsed)

    tables = parse_tables(body)

    # Header/footer parts
    hf_parts = parse_header_footer_parts(header_parts, footer_parts)

    # Numbering
    numbering_info = parse_numbering(numbering_xml)

    # Which media files are actually referenced?
    referenced_targets = {r["target"] for r in rels.values() if r["type"] == "image"}
    media_status = []
    for mpath in sorted(media.keys()):
        basename = mpath.split("/")[-1]
        rel_path = "media/" + basename
        is_used = rel_path in referenced_targets or basename in referenced_targets
        media_status.append({
            "path": mpath,
            "referenced": is_used,
            "sha256": fingerprints.get(mpath, "N/A"),
            "warning": "" if is_used else "UNREFERENCED - do NOT auto-place in header"
        })

    return {
        "source": os.path.abspath(path),
        "theme_fonts": theme_fonts,
        "doc_defaults": doc_defaults,
        "media_files": media_status,
        "relationships": [{"rId": k, **v} for k, v in rels.items()],
        "paragraph_count": len(paragraphs),
        "empty_paragraph_count": sum(1 for p in paragraphs if p["is_empty"]),
        "table_count": len(tables),
        "paragraphs": paragraphs,
        "tables": tables,
        "header_footer_parts": hf_parts,
        "body_page_setup": body_page_setup,
        "body_header_footer_refs": body_header_footer_refs,
        "numbering": numbering_info,
        "part_fingerprints": fingerprints,
        "has_numbering": numbering_xml is not None,
        "has_theme": theme_xml is not None,
        "has_settings": settings_xml is not None,
    }


def _is_inside_table(p_elem, body):
    """Check if a paragraph is inside a table (skip those for paragraph count)."""
    parent = p_elem.getparent() if hasattr(p_elem, 'getparent') else None
    # ElementTree doesn't have getparent; use iter approach
    for tbl in body.iter(qn("w:tbl")):
        for p_in_tbl in tbl.iter(qn("w:p")):
            if p_in_tbl is p_elem:
                return True
    return False


def _parse_sect_pr(sect_pr):
    """Parse a sectPr element for page size, margins, titlePg, pgNumType,
    and header/footer references."""
    result = {
        "page_size": {},
        "margins": {},
        "has_titlePg": False,
        "pgNumType_start": None,
        "section_type": None,
        "header_refs": [],
        "footer_refs": [],
    }

    pg_sz = sect_pr.find(qn("w:pgSz"))
    if pg_sz is not None:
        # FIX: only int-cast w:w and w:h; treat w:orient and w:code as strings
        w = pg_sz.get(qn("w:w"))
        h = pg_sz.get(qn("w:h"))
        result["page_size"] = {
            "w_twips": int(w) if w and w.isdigit() else w,
            "h_twips": int(h) if h and h.isdigit() else h,
            "orient": pg_sz.get(qn("w:orient"), "portrait"),
            "code": pg_sz.get(qn("w:code")),
        }
        # Also provide EMU (1 twip = 635 EMU)
        if w and w.isdigit():
            result["page_size"]["w_emu"] = int(w) * 635
        if h and h.isdigit():
            result["page_size"]["h_emu"] = int(h) * 635

    pg_mar = sect_pr.find(qn("w:pgMar"))
    if pg_mar is not None:
        for attr in ["top", "right", "bottom", "left", "header", "footer", "gutter"]:
            val = pg_mar.get(qn(f"w:{attr}"))
            if val and val.isdigit():
                result["margins"][f"{attr}_twips"] = int(val)
                result["margins"][f"{attr}_inches"] = round(int(val) / 1440, 2)

    if sect_pr.find(qn("w:titlePg")) is not None:
        result["has_titlePg"] = True

    pgNum = sect_pr.find(qn("w:pgNumType"))
    if pgNum is not None:
        start = pgNum.get(qn("w:start"))
        result["pgNumType_start"] = int(start) if start and start.isdigit() else start

    sec_type = sect_pr.find(qn("w:type"))
    if sec_type is not None:
        result["section_type"] = sec_type.get(qn("w:val"))

    for ref in sect_pr.findall(qn("w:headerReference")):
        result["header_refs"].append({
            "type": ref.get(qn("w:type")),
            "rId": ref.get(qn("r:id")),
        })
    for ref in sect_pr.findall(qn("w:footerReference")):
        result["footer_refs"].append({
            "type": ref.get(qn("w:type")),
            "rId": ref.get(qn("r:id")),
        })

    return result


# ----- diff mode -------------------------------------------------------------
def diff(original_path, output_path):
    """Compare two docx files. Reports which parts are byte-identical,
    which differ, and what changed. Used for preservation verification."""
    print(f"[DIFF] Comparing:")
    print(f"  Original: {original_path}")
    print(f"  Output:   {output_path}")
    print()

    with zipfile.ZipFile(original_path) as z1:
        orig_parts = {n: z1.read(n) for n in z1.namelist()}
    with zipfile.ZipFile(output_path) as z2:
        out_parts = {n: z2.read(n) for n in z2.namelist()}

    all_parts = sorted(set(orig_parts.keys()) | set(out_parts.keys()))
    identical = []
    changed = []
    only_in_orig = []
    only_in_out = []

    for part in all_parts:
        if part in orig_parts and part in out_parts:
            if orig_parts[part] == out_parts[part]:
                identical.append(part)
            else:
                changed.append(part)
        elif part in orig_parts:
            only_in_orig.append(part)
        else:
            only_in_out.append(part)

    print(f"  IDENTICAL parts ({len(identical)}/{len(all_parts)}):")
    for p in identical:
        print(f"    [OK] {p}")

    print(f"\n  CHANGED parts ({len(changed)}):")
    for p in changed:
        orig_hash = hashlib.sha256(orig_parts[p]).hexdigest()[:16]
        out_hash = hashlib.sha256(out_parts[p]).hexdigest()[:16]
        print(f"    [CHG] {p}")
        print(f"          orig sha256={orig_hash}...")
        print(f"          out  sha256={out_hash}...")

    if only_in_orig:
        print(f"\n  REMOVED parts ({len(only_in_orig)}):")
        for p in only_in_orig:
            print(f"    [DEL] {p}")

    if only_in_out:
        print(f"\n  ADDED parts ({len(only_in_out)}):")
        for p in only_in_out:
            print(f"    [ADD] {p}")

    print()
    print(f"  SUMMARY: {len(identical)} identical, {len(changed)} changed, "
          f"{len(only_in_orig)} removed, {len(only_in_out)} added")
    fidelity_pct = round(100 * len(identical) / len(all_parts), 1) if all_parts else 0
    print(f"  BYTE FIDELITY: {fidelity_pct}%")

    return {
        "identical": identical,
        "changed": changed,
        "only_in_original": only_in_orig,
        "only_in_output": only_in_out,
        "fidelity_pct": fidelity_pct,
    }


# ----- human-readable summary ------------------------------------------------
def print_summary(spec):
    src = spec["source"]
    print("=" * 72)
    print("  DOCUMENT INSPECTION REPORT (v2)")
    print(f"  Source: {src}")
    print("=" * 72)

    # Theme fonts — CRITICAL for resolving 'inherited'
    if spec.get("theme_fonts"):
        tf = spec["theme_fonts"]
        print(f"\n--- Theme fonts (from theme1.xml) ---")
        print(f"  Major font (headings): {tf.get('major', 'N/A')}")
        print(f"  Minor font (body):    {tf.get('minor', 'N/A')}")
        print(f"  NOTE: Runs reporting 'inherit' use these fonts via the theme.")

    # Doc defaults
    dd = spec.get("doc_defaults", {})
    if dd.get("rPr"):
        print(f"\n--- Document defaults (from styles.xml docDefaults) ---")
        rPr = dd["rPr"]
        print(f"  Default font:      {rPr.get('font', 'N/A')}")
        print(f"  Default size:      {rPr.get('size_pt', 'N/A')}pt")
        print(f"  Default color:     #{rPr.get('color', 'N/A')}")
        print(f"  Default bold:      {rPr.get('bold', False)}")

    # Page setup — now reads body-level sectPr
    if spec.get("body_page_setup"):
        ps = spec["body_page_setup"]
        print(f"\n--- Page setup (body-level sectPr) ---")
        if ps.get("page_size"):
            pgs = ps["page_size"]
            print(f"  Page size:   {pgs.get('w_twips','?')} x {pgs.get('h_twips','?')} twips "
                  f"({pgs.get('w_emu','?')} x {pgs.get('h_emu','?')} EMU) "
                  f"orient={pgs.get('orient','portrait')}")
        if ps.get("margins"):
            m = ps["margins"]
            print(f"  Margins:     L={m.get('left_inches','?')}\"  R={m.get('right_inches','?')}\"  "
                  f"T={m.get('top_inches','?')}\"  B={m.get('bottom_inches','?')}\"")
        print(f"  titlePg:     {ps.get('has_titlePg', False)}")
        print(f"  pgNumType start: {ps.get('pgNumType_start', 'N/A')}")
        if ps.get("header_refs"):
            print(f"  Header refs: {ps['header_refs']}")
        if ps.get("footer_refs"):
            print(f"  Footer refs: {ps['footer_refs']}")

    # Header/footer parts
    hf = spec.get("header_footer_parts", {})
    if hf.get("headers"):
        print(f"\n--- Header parts ({len(hf['headers'])}) ---")
        for path, paras in hf["headers"].items():
            print(f"  {path}:")
            for i, p in enumerate(paras):
                print(f"    [{i}] align={p['alignment']}, size={p['size_pt']}pt, "
                      f"bold={p['bold']}, text='{p['text'][:60]}'"
                      f"{' [PAGE FIELD]' if p['has_page_field'] else ''}")
    if hf.get("footers"):
        print(f"\n--- Footer parts ({len(hf['footers'])}) ---")
        for path, paras in hf["footers"].items():
            print(f"  {path}:")
            for i, p in enumerate(paras):
                print(f"    [{i}] align={p['alignment']}, size={p['size_pt']}pt, "
                      f"bold={p['bold']}, text='{p['text'][:60]}'"
                      f"{' [PAGE FIELD]' if p['has_page_field'] else ''}")

    # Numbering
    if spec.get("has_numbering"):
        nb = spec["numbering"]
        print(f"\n--- Numbering ({len(nb['abstract_nums'])} abstractNums, "
              f"{len(nb['nums'])} nums) ---")
        for an in nb["abstract_nums"][:3]:  # show first 3
            lvl0 = an["levels"][0] if an["levels"] else {}
            print(f"  abstractNumId={an['abstractNumId']}, nsid={an['nsid']}, "
                  f"lvl0 numFmt={lvl0.get('numFmt')}, lvlText={lvl0.get('lvlText')}")
        if nb["shared_abstractNum_warnings"]:
            print(f"\n  [WARNING] Shared abstractNum detected:")
            for w in nb["shared_abstractNum_warnings"]:
                print(f"    {w['warning']}")

    # Media files
    print(f"\n--- Media files ({len(spec['media_files'])}) ---")
    for m in spec["media_files"]:
        flag = "REFERENCED" if m["referenced"] else "UNREFERENCED"
        print(f"  {m['path']:40s}  [{flag}]")
        if m["warning"]:
            print(f"      ! {m['warning']}")

    # Paragraphs — now includes empty paras and resolved fonts
    print(f"\n--- Paragraph-by-paragraph formatting ({spec['paragraph_count']} paragraphs, "
          f"{spec['empty_paragraph_count']} empty) ---")
    for i, p in enumerate(spec["paragraphs"]):
        if p["is_empty"]:
            # Show empty paras with a marker (v1 dropped these silently)
            sect_marker = " [has sectPr]" if p.get("has_sectPr") else ""
            print(f"  [P{i:>3}] (empty){sect_marker}")
            if p.get("sect_pr"):
                print(f"         sectPr: {p['sect_pr']}")
            continue

        if p["borders"]:
            for b in p["borders"]:
                print(f"  [P{i:>3}] BORDER {b['edge']:5s} style={b['style']:8s} "
                      f"size={b['size_eighth_pt']:.2f}pt color=#{b.get('color','?')}")
        if p["alignment"]:
            print(f"  [P{i:>3}] alignment={p['alignment']}", end="")
            if p["line_spacing"]:
                print(f", line_spacing={p['line_spacing']}", end="")
            if p["style_id"]:
                print(f", style={p['style_id']}", end="")
            if p["num_pr"]:
                print(f", numId={p['num_pr']['numId']}", end="")
            print()

        for run in p["runs"]:
            text = run.get("text", "")
            if not text.strip() and "images" not in run:
                continue
            size  = f"{run['size_pt']}pt" if run["size_pt"] else "inherit"
            bold  = "B" if run["bold"] else " "
            ital  = "I" if run["italic"] else " "
            color = run["color"] or "inherit"
            font  = (run["font"] or "inherit")[:20]
            src   = ",".join(run.get("resolved_from", [])) or "direct"
            print(f"  [P{i:>3}] [{size:>7s} {bold}{ital} #{color:>10s} {font:>20s}] "
                  f"<-{src[:25]:<25s}  {text[:60]}")
            if "images" in run:
                for img in run["images"]:
                    print(f"           -> IMAGE rId={img['rId']}  target={img['target']}")

    # Tables — now with borders and three-line detection
    print(f"\n--- Tables ({spec['table_count']}) ---")
    for ti, t in enumerate(spec["tables"]):
        three_line = " [THREE-LINE]" if t.get("is_three_line") else ""
        print(f"  Table {ti}: {t['row_count']} rows{three_line}")
        if t.get("table_borders"):
            print(f"    Borders: {t['table_borders']}")
        for ri, row in enumerate(t["rows_preview"]):
            for ci, cell in enumerate(row["cells"]):
                span = f" span={cell['grid_span']}" if cell["grid_span"] else ""
                shade = f" fill=#{cell['shading_fill']}" if cell["shading_fill"] else ""
                borders_str = ""
                if cell.get("cell_borders"):
                    borders_str = f" borders={list(cell['cell_borders'].keys())}"
                print(f"    R{ri}C{ci}:{span}{shade}{borders_str}  "
                      f"{cell['text_preview'][:50]}")

    # Part fingerprints
    fps = spec.get("part_fingerprints", {})
    if fps:
        print(f"\n--- Part sha256 fingerprints ({len(fps)} parts) ---")
        for path, hash_val in sorted(fps.items())[:10]:
            print(f"  {path:40s}  sha256={hash_val[:16]}...")
        if len(fps) > 10:
            print(f"  ... and {len(fps) - 10} more")

    print("\n" + "=" * 72)
    print("  v2 improvements: theme1.xml resolved, body-level sectPr read,")
    print("  empty paras kept, header/footer parts parsed, numbering detected,")
    print("  table borders extracted, sha256 fingerprints computed.")
    print("=" * 72)


# ----- preview renderer ------------------------------------------------------
def render_preview(docx_path, out_png):
    """Convert docx -> pdf -> png. Portable: tries libreoffice, then
    libreoffice-kit, then skips gracefully if neither available."""
    outdir = os.path.dirname(os.path.abspath(out_png)) or "."
    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf",
             docx_path, "--outdir", outdir],
            check=True, capture_output=True
        )
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"[WARN] libreoffice not available or failed: {e}", file=sys.stderr)
        print(f"       Skip preview, or use --json output only.", file=sys.stderr)
        return False
    pdf = os.path.join(outdir, os.path.basename(docx_path).replace(".docx", ".pdf"))
    if not os.path.exists(pdf):
        print(f"[WARN] PDF not found at {pdf}", file=sys.stderr)
        return False
    prefix = out_png.replace(".png", "")
    try:
        subprocess.run(
            ["pdftoppm", "-r", "110", "-png", pdf, prefix],
            check=True, capture_output=True
        )
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"[WARN] pdftoppm not available or failed: {e}", file=sys.stderr)
        return False
    print(f"[OK] Preview rendered -> {out_png}")
    return True


# ----- main ------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", nargs="?", help="Path to the .docx file to inspect")
    ap.add_argument("--json", dest="json_out", help="Write JSON spec to this file")
    ap.add_argument("--preview", help="Render page-1 PNG preview to this path")
    ap.add_argument("--diff", nargs=2, metavar=("ORIGINAL", "OUTPUT"),
                    help="Compare two docx files for preservation checks")
    args = ap.parse_args()

    # ASCII-only status messages to stdout (not stderr) — fixes PowerShell
    # spurious exit code 1 and cp1252 mojibake

    if args.diff:
        diff(args.diff[0], args.diff[1])
        return

    if not args.docx:
        ap.print_help()
        sys.exit(2)

    if not os.path.exists(args.docx):
        print(f"[ERROR] File not found: {args.docx}", file=sys.stderr)
        sys.exit(2)

    spec = inspect(args.docx)

    if args.json_out:
        with open(args.json_out, "w") as f:
            json.dump(spec, f, indent=2, default=str)
        print(f"[OK] JSON spec written -> {args.json_out}")

    print_summary(spec)

    if args.preview:
        render_preview(args.docx, args.preview)


if __name__ == "__main__":
    main()
