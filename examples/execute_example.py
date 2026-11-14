"""Start an execution of the report workflow with the uploaded sample inputs.

Usage:
    python execute_example.py <guidelines.pdf> <cover.docx> <chapter1.docx> \
        [--source source.docx] [--mode GENERATE]

Requires MISTRAL_API_KEY and a deployed worker serving
`major-project-seminar-report`.
"""

import argparse
import base64
import os

from mistralai.client import Mistral


def file_input(path: str) -> dict:
    with open(path, "rb") as handle:
        return {"filename": os.path.basename(path),
                "content_b64": base64.b64encode(handle.read()).decode("ascii")}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("guidelines_pdf")
    parser.add_argument("sample_cover_docx")
    parser.add_argument("sample_chapter_docx")
    parser.add_argument("--source", help="user's source content DOCX (Stream B)")
    parser.add_argument("--mode", default="GENERATE",
                        choices=["REFORMAT_ONLY", "GENERATE", "EXPAND"])
    args = parser.parse_args()

    client = Mistral(api_key=os.environ["MISTRAL_API_KEY"])
    execution = client.workflows.execute_workflow(
        workflow_identifier="major-project-seminar-report",
        input={
            "project": {
                "name": "YOUR PROJECT TITLE",
                "type": "major-project",
                "institute": "YOUR INSTITUTE",
                "department": "YOUR DEPARTMENT",
                "academic_year": "2024-25",
            },
            "content_mode": args.mode,
            "guidelines_pdf": file_input(args.guidelines_pdf),
            "sample_cover_docx": file_input(args.sample_cover_docx),
            "sample_chapter_docx": file_input(args.sample_chapter_docx),
            "source_content_docx": file_input(args.source) if args.source else None,
            "project_details": {
                "algorithms_used": ["fill from your project"],
                "metrics": {},
                "figures": [],
                "tables": [],
                "dataset": "",
                "references": [],
            },
        },
    )
    print(execution.execution_id)


if __name__ == "__main__":
    main()
