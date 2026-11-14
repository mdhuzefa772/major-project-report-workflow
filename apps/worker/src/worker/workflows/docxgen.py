"""DOCX generation activity: three-tier page structure built from the Anchor spec.

Tier 0: cover page (no header/footer, not counted).
Tier 1: first content page (repeats the cover title, no header/footer shown).
Tier 2: subsequent pages ("Chapter -N" header right-aligned, centered PAGE field footer).

Headers/footers are independent XML parts (never shared between sections) and every
sectPr child is inserted in CT_SectPr schema order (spec sections 10-11, P5).
"""

import base64
import io

import mistralai.workflows as workflows
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Emu, Inches, Pt
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .anchors import anchor

HEADER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FOOTER_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"

CT_SECTPR_ORDER = [
    "headerReference", "footerReference", "footnotePr", "endnotePr",
    "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
    "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
    "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
    "printerSettings", "sectPrChange",
]

JC_MAP = {
    "left": "left", "right": "right", "center": "center",
    "justify": "both", "justified": "both", "both": "both",
}

ALIGN_ENUM = {
    "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
    "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
    "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
    "JUSTIFIED": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "BOTH": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _align(value, default="LEFT"):
    if value is None:
        return ALIGN_ENUM[default]
    return ALIGN_ENUM.get(str(value).upper(), ALIGN_ENUM[default])


def _insert_sectpr_child(sectpr, child_or_tagname, attrs=None):
    if isinstance(child_or_tagname, str):
        elem = OxmlElement(child_or_tagname)
        if attrs:
            for key, value in attrs.items():
                elem.set(qn(key), value)
    else:
        elem = child_or_tagname

    localname = elem.tag.split("}")[1] if "}" in elem.tag else elem.tag
    target_idx = CT_SECTPR_ORDER.index(localname) if localname in CT_SECTPR_ORDER else len(CT_SECTPR_ORDER)

    insert_pos = 0
    for i, existing in enumerate(list(sectpr)):
        existing_ln = existing.tag.split("}")[1] if "}" in existing.tag else existing.tag
        if existing_ln in CT_SECTPR_ORDER:
            if CT_SECTPR_ORDER.index(existing_ln) < target_idx:
                insert_pos = i + 1
            else:
                break
        else:
            break
    sectpr.insert(insert_pos, elem)


def _create_xml_part(doc_part, xml_str, content_type, part_uri, rel_type):
    from lxml import etree

    element = etree.fromstring(xml_str.encode("utf-8"))
    blob = etree.tostring(element, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_part = Part(PackURI(part_uri), content_type, blob=blob, package=doc_part.package)
    return doc_part.relate_to(new_part, rel_type)


def _make_header_xml(text, font_family, font_size_pt, alignment):
    sz = str(int(font_size_pt * 2))
    jc = JC_MAP.get(str(alignment).lower(), "right")
    return (
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:jc w:val="' + jc + '"/>'
        '<w:spacing w:after="0" w:before="0"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="' + font_family + '" w:hAnsi="' + font_family + '"/>'
        '<w:sz w:val="' + sz + '"/></w:rPr>'
        '<w:t xml:space="preserve">' + text + "</w:t></w:r></w:p></w:hdr>"
    )


def _make_empty_hf_xml(tag):
    return (
        '<w:' + tag + ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:' + tag + ">"
    )


def _make_footer_page_xml():
    return (
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:jc w:val="center"/>'
        '<w:spacing w:after="0" w:before="0"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "</w:p></w:ftr>"
    )


def _make_pgnumtype(start: int):
    element = OxmlElement("w:pgNumType")
    element.set(qn("w:start"), str(max(1, int(start) - 1)))
    return element


def _strip_section0_refs(sectpr0):
    for tag in (qn("w:headerReference"), qn("w:footerReference")):
        for ref in list(sectpr0.findall(tag)):
            sectpr0.remove(ref)


def _setup_page(section, spec):
    width_emu = anchor(spec, "PAGE.WIDTH_EMU")
    height_emu = anchor(spec, "PAGE.HEIGHT_EMU")
    if width_emu is not None:
        section.page_width = Emu(int(width_emu))
    if height_emu is not None:
        section.page_height = Emu(int(height_emu))
    for side in ("left", "right", "top", "bottom"):
        inches = anchor(spec, f"PAGE.MARGIN_{side.upper()}_IN")
        if inches is not None:
            setattr(section, f"{side}_margin", Inches(float(inches)))


def _add_run(paragraph, text, spec, font_size=None, bold=None, italic=None):
    run = paragraph.add_run(text)
    run.font.name = anchor(spec, "FONT.FAMILY")
    if font_size is not None:
        run.font.size = Pt(float(font_size))
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    return run


def _format_paragraph(paragraph, spec, alignment=None, line_spacing=None,
                      space_before=None, space_after=None):
    fmt = paragraph.paragraph_format
    if alignment is not None:
        fmt.alignment = alignment
    fmt.line_spacing = line_spacing if line_spacing is not None else anchor(spec, "FONT.BODY_LINE_SPACING", 1.5)
    if space_before is not None:
        fmt.space_before = space_before
    if space_after is not None:
        fmt.space_after = space_after


def _add_three_line_table(doc, table_data, spec):
    header_row = table_data.get("header_row") or []
    rows = table_data.get("rows") or []
    if not header_row and not rows:
        return
    n_cols = len(header_row) or max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows) + (1 if header_row else 0), cols=n_cols)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    all_rows = ([header_row] if header_row else []) + rows
    for ri, row in enumerate(all_rows):
        for ci in range(n_cols):
            cell = table.cell(ri, ci)
            cell.text = ""
            para = cell.paragraphs[0]
            para.alignment = _align(anchor(spec, "TABLE.CELL_ALIGNMENT"), "CENTER")
            _format_paragraph(para, spec, line_spacing=1.0)
            text = str(row[ci]) if ci < len(row) else ""
            _add_run(para, text, spec,
                     font_size=anchor(spec, "TABLE.CELL_SIZE_PT"),
                     bold=(ri == 0 and anchor(spec, "TABLE.HEADER_ROW_BOLD", True)))

    tbl = table._tbl
    tbl_pr = tbl.tblPr
    for existing in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for name, val, sz, color in (
        ("top", anchor(spec, "TABLE.BORDER_TOP_VAL", "single"),
         anchor(spec, "TABLE.BORDER_TOP_SZ_EIGHTH", 4), anchor(spec, "TABLE.BORDER_TOP_COLOR", "000000")),
        ("bottom", anchor(spec, "TABLE.BORDER_BOTTOM_VAL", "single"),
         anchor(spec, "TABLE.BORDER_BOTTOM_SZ_EIGHTH", 4), anchor(spec, "TABLE.BORDER_BOTTOM_COLOR", "000000")),
    ):
        edge = OxmlElement(f"w:{name}")
        edge.set(qn("w:val"), str(val))
        edge.set(qn("w:sz"), str(sz))
        edge.set(qn("w:space"), "0")
        edge.set(qn("w:color"), str(color))
        borders.append(edge)
    for name in ("left", "right", "insideH", "insideV"):
        edge = OxmlElement(f"w:{name}")
        edge.set(qn("w:val"), "none")
        edge.set(qn("w:sz"), "0")
        edge.set(qn("w:space"), "0")
        edge.set(qn("w:color"), "auto")
        borders.append(edge)
    tbl_pr.append(borders)

    for cell in table.rows[0].cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        for existing in tc_pr.findall(qn("w:tcBorders")):
            tc_pr.remove(existing)
        tc_borders = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), str(anchor(spec, "TABLE.BORDER_TOP_SZ_EIGHTH", 4)))
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), str(anchor(spec, "TABLE.BORDER_TOP_COLOR", "000000")))
        tc_borders.append(bottom)
        tc_pr.append(tc_borders)


def _add_title_block(doc, spec, chapter_num, title, is_cover):
    for _ in range(anchor(spec, "COVER_PAGE.VERTICAL_OFFSET_PARAGRAPHS", 18) if is_cover else 0):
        doc.add_paragraph()
    for line in (f"CHAPTER -{chapter_num}", str(title).upper()):
        para = doc.add_paragraph()
        _format_paragraph(
            para, spec,
            alignment=_align(anchor(spec, "FONT.COVER_TITLE_ALIGNMENT"), "RIGHT"),
            line_spacing=1.5,
        )
        _add_run(para, line, spec,
                 font_size=anchor(spec, "FONT.COVER_TITLE_SIZE_PT"),
                 bold=anchor(spec, "FONT.COVER_TITLE_BOLD", True))


def _add_body_paragraph(doc, text, spec):
    para = doc.add_paragraph()
    _format_paragraph(para, spec,
                      alignment=_align(anchor(spec, "FONT.BODY_ALIGNMENT"), "JUSTIFIED"))
    _add_run(para, text, spec, font_size=anchor(spec, "FONT.BODY_SIZE_PT"))
    return para


def _add_main_heading(doc, text, spec):
    para = doc.add_paragraph()
    _format_paragraph(para, spec,
                      alignment=_align(anchor(spec, "FONT.MAIN_HEADING_ALIGNMENT"), "LEFT"))
    _add_run(para, text, spec,
             font_size=anchor(spec, "FONT.MAIN_HEADING_SIZE_PT"),
             bold=anchor(spec, "FONT.MAIN_HEADING_BOLD", True))
    return para


def _add_sub_heading(doc, text, spec):
    para = doc.add_paragraph()
    _format_paragraph(para, spec,
                      alignment=_align(anchor(spec, "FONT.SUB_HEADING_ALIGNMENT"), "LEFT"))
    _add_run(para, text, spec,
             font_size=anchor(spec, "FONT.SUB_HEADING_SIZE_PT"),
             bold=anchor(spec, "FONT.SUB_HEADING_BOLD", True),
             italic=anchor(spec, "FONT.SUB_HEADING_ITALIC", True))
    return para


def _add_caption(doc, text, spec):
    para = doc.add_paragraph()
    _format_paragraph(para, spec, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    _add_run(para, text, spec, font_size=anchor(spec, "FONT.BODY_SIZE_PT"), bold=True)
    return para


def _add_bullet(doc, text, spec):
    para = doc.add_paragraph(style="List Bullet")
    _format_paragraph(para, spec,
                      alignment=_align(anchor(spec, "FONT.BODY_ALIGNMENT"), "JUSTIFIED"))
    _add_run(para, text, spec, font_size=anchor(spec, "FONT.BODY_SIZE_PT"))
    return para


def _add_chapter_content(doc, draft, spec):
    if draft.get("intro_paragraph"):
        _add_body_paragraph(doc, draft["intro_paragraph"], spec)

    for section in draft.get("sections", []):
        _add_main_heading(doc, f"{section.get('number')} {section.get('heading', '')}".strip(), spec)
        for paragraph in section.get("paragraphs", []):
            _add_body_paragraph(doc, paragraph, spec)
        for bullet in section.get("bullets", []):
            _add_bullet(doc, bullet, spec)
        for table_data in section.get("tables", []):
            caption = table_data.get("caption")
            if caption:
                _add_caption(doc, caption, spec)
            _add_three_line_table(doc, table_data, spec)
        for caption in section.get("figure_captions", []):
            _add_caption(doc, caption, spec)
        for sub in section.get("subsections", []):
            _add_sub_heading(doc, f"{sub.get('number')} {sub.get('heading', '')}".strip(), spec)
            for paragraph in sub.get("paragraphs", []):
                _add_body_paragraph(doc, paragraph, spec)


def build_chapter_docx(draft: dict, spec: dict) -> bytes:
    chapter_num = draft["chapter_num"]
    title = draft.get("title", "")

    doc = Document()
    _setup_page(doc.sections[0], spec)
    _add_title_block(doc, spec, chapter_num, title, is_cover=True)

    section1 = doc.add_section(WD_SECTION.NEW_PAGE)
    _setup_page(section1, spec)

    sectpr1 = section1._sectPr
    doc_part = doc.part
    counter = 1

    if anchor(spec, "SUBSEQUENT_PAGES.HAS_HEADER", True):
        header_template = anchor(spec, "SUBSEQUENT_PAGES.HEADER_TEXT_TEMPLATE", "Chapter -{{CHAPTER_NUM}}")
        header_text = header_template.replace("{{CHAPTER_NUM}}", str(chapter_num))
        r_default = _create_xml_part(
            doc_part,
            _make_header_xml(header_text, anchor(spec, "FONT.FAMILY", "Times New Roman"),
                             anchor(spec, "SUBSEQUENT_PAGES.HEADER_SIZE_PT", 10),
                             anchor(spec, "SUBSEQUENT_PAGES.HEADER_ALIGNMENT", "RIGHT")),
            HEADER_CT, f"/word/header{counter}.xml", RT.HEADER)
        counter += 1
        r_first = _create_xml_part(doc_part, _make_empty_hf_xml("hdr"), HEADER_CT,
                                   f"/word/header{counter}.xml", RT.HEADER)
        counter += 1
        _insert_sectpr_child(sectpr1, "w:headerReference", {"w:type": "default", "r:id": r_default})
        _insert_sectpr_child(sectpr1, "w:headerReference", {"w:type": "first", "r:id": r_first})

    if anchor(spec, "SUBSEQUENT_PAGES.HAS_FOOTER", True) and anchor(spec, "SUBSEQUENT_PAGES.FOOTER_TYPE") == "PAGE_FIELD":
        r_footer = _create_xml_part(doc_part, _make_footer_page_xml(), FOOTER_CT,
                                    f"/word/footer{counter}.xml", RT.FOOTER)
        counter += 1
        r_first_footer = _create_xml_part(doc_part, _make_empty_hf_xml("ftr"), FOOTER_CT,
                                          f"/word/footer{counter}.xml", RT.FOOTER)
        counter += 1
        _insert_sectpr_child(sectpr1, "w:footerReference", {"w:type": "default", "r:id": r_footer})
        _insert_sectpr_child(sectpr1, "w:footerReference", {"w:type": "first", "r:id": r_first_footer})

    _insert_sectpr_child(sectpr1, _make_pgnumtype(anchor(spec, "SUBSEQUENT_PAGES.PAGE_NUMBER_START", 2)))
    _insert_sectpr_child(sectpr1, OxmlElement("w:titlePg"))

    _strip_section0_refs(doc.sections[0]._sectPr)

    if anchor(spec, "FIRST_CONTENT_PAGE.REPEATS_COVER_TITLE", True):
        _add_title_block(doc, spec, chapter_num, title, is_cover=False)
    _add_chapter_content(doc, draft, spec)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


@workflows.activity()
async def generate_chapter_docx(draft: dict, spec: dict) -> dict:
    docx_bytes = build_chapter_docx(draft, spec)
    return {
        "filename": f"Chapter-{draft['chapter_num']}.docx",
        "docx_b64": base64.b64encode(docx_bytes).decode("ascii"),
    }
