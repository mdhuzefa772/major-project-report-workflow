"""Offline end-to-end test of the deterministic core: anchors -> docx -> verification."""

import json
import sys

sys.path.insert(0, "apps/worker/src")

from worker.workflows import anchors as am
from worker.workflows import inspect_docx
from worker.workflows.docxgen import build_chapter_docx
from worker.workflows.verification import verify_compliance

SAMPLES = {
    "cover": "/tmp/wtest/cover.docx",
    "chapter": "/tmp/wtest/ch1.docx",
}

STREAM_A = {
    "PAGE.SIZE": "A4",
    "PAGE.MARGIN_LEFT_IN": 1.5, "PAGE.MARGIN_RIGHT_IN": 1.0,
    "PAGE.MARGIN_TOP_IN": 1.0, "PAGE.MARGIN_BOTTOM_IN": 1.0,
    "FONT.FAMILY": "Times New Roman",
    "FONT.BODY_SIZE_PT": 12, "FONT.BODY_ALIGNMENT": "JUSTIFIED", "FONT.BODY_LINE_SPACING": 1.5,
    "FONT.MAIN_HEADING_SIZE_PT": 14, "FONT.MAIN_HEADING_BOLD": True, "FONT.MAIN_HEADING_ALIGNMENT": "LEFT",
    "FONT.SUB_HEADING_SIZE_PT": 12, "FONT.SUB_HEADING_BOLD": True, "FONT.SUB_HEADING_ITALIC": True,
    "FONT.SUB_HEADING_ALIGNMENT": "LEFT",
    "FONT.COVER_TITLE_SIZE_PT": 16, "FONT.COVER_TITLE_BOLD": True, "FONT.COVER_TITLE_ALIGNMENT": "RIGHT",
    "TABLE.STYLE": "THREE_LINE_ACADEMIC",
    "TABLE.BORDER_TOP_VAL": "single", "TABLE.BORDER_TOP_SZ_EIGHTH": 4, "TABLE.BORDER_TOP_COLOR": "000000",
    "TABLE.BORDER_BOTTOM_VAL": "single", "TABLE.BORDER_BOTTOM_SZ_EIGHTH": 4, "TABLE.BORDER_BOTTOM_COLOR": "000000",
    "TABLE.HEADER_ROW_BOLD": True, "TABLE.CELL_ALIGNMENT": "CENTER", "TABLE.CELL_SIZE_PT": 11,
    "SUBSEQUENT_PAGES.FOOTER_ALIGNMENT": "CENTER", "SUBSEQUENT_PAGES.FOOTER_TYPE": "PAGE_FIELD",
    "REFERENCES.FORMAT": "IEEE",
}

LONG = ("The project predicts stock market trends using machine learning models trained on historical data "
        "collected from multiple listed companies over the last decade. Three models are compared: Multiple "
        "Linear Regression, Random Forest and XGBoost, each tuned with cross-validation on the same splits. "
        "The dataset covers ten years of daily prices along with derived technical indicators such as moving "
        "averages and momentum oscillators. Performance is measured with R-squared and RMSE so that accuracy "
        "and error magnitude can be compared consistently across models. The remainder of this section "
        "describes the system context, the motivation for automated forecasting, and the scope of the study.")

draft = {
    "chapter_num": 1, "title": "INTRODUCTION",
    "intro_paragraph": "This chapter introduces the project and its motivation in the domain of financial time series forecasting.",
    "sections": [
        {"number": "1.1", "heading": "Overview",
         "paragraphs": [LONG] * 3,
         "bullets": [], "subsections": [], "tables": [
             {"caption": "Table 1.1 Model Performance", "header_row": ["Model", "R2"],
              "rows": [["XGBoost", "57.46%"], ["Random Forest", "52.10%"]]}],
         "figure_captions": ["Fig. 1.1 System Architecture"]},
        {"number": "1.2", "heading": "Objectives",
         "paragraphs": [LONG, LONG],
         "bullets": ["Compare three regression models", "Evaluate with R2 and RMSE"],
         "subsections": [], "tables": [], "figure_captions": []},
    ],
    "content_mode": "GENERATE",
}


def main():
    stream_c = {}
    for kind, path in SAMPLES.items():
        stream_c.update(am.anchors_from_inspect(inspect_docx.inspect(path), kind))

    project = {"name": "Stock Market Prediction using ML", "type": "major-project",
               "institute": "Test Institute", "department": "CSE", "academic_year": "2024-25",
               "actual_results": {"algorithms_used": ["MLR", "Random Forest", "XGBoost"]}}
    spec, conflicts, missing = am.reconcile(STREAM_A, stream_c, project, "GENERATE", None)
    print("CONFLICTS:", [c["anchor"] for c in conflicts])
    print("MISSING:", missing)

    docx = build_chapter_docx(draft, spec)
    with open("/tmp/wtest/out_ch1.docx", "wb") as handle:
        handle.write(docx)

    violations = verify_compliance(docx, spec, draft)
    blocking = [v for v in violations if v["severity"].startswith(("HIGH", "MEDIUM"))]
    print("VIOLATIONS:", json.dumps(violations, indent=1))
    assert not blocking, f"blocking violations: {blocking}"

    out = inspect_docx.inspect("/tmp/wtest/out_ch1.docx")
    page = out["body_page_setup"]
    assert page["page_size"]["w_twips"] == 11906 and page["page_size"]["h_twips"] == 16838
    assert abs(page["margins"]["left_inches"] - 1.5) < 0.01
    assert abs(page["margins"]["right_inches"] - 1.0) < 0.01
    assert page["has_titlePg"] is True
    assert page["pgNumType_start"] == 1
    assert page["header_refs"], "section 1 must reference a header part"

    headers = out["header_footer_parts"]["headers"]
    default_headers = [p for part in headers.values() for p in part if (p.get("text") or "").strip()]
    assert default_headers and default_headers[0]["text"] == "Chapter -1", default_headers
    assert default_headers[0]["alignment"] == "right"

    paras = [p for p in out["paragraphs"] if not p["is_empty"] and p["runs"]]
    sizes = {p["runs"][0]["size_pt"] for p in paras}
    assert 12 in sizes and 14 in sizes and 16 in sizes, sizes
    assert out["table_count"] == 1
    print("\nALL OFFLINE CHECKS PASSED")


if __name__ == "__main__":
    main()
