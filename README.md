# Major Project Technical Seminar Report — Mistral Workflow

A durable [Mistral Workflow](https://docs.mistral.ai/studio/workflows/) that turns an
institute guidelines PDF, sample DOCX files and a student's raw project material into
institute-compliant, chapter-wise DOCX report files — implementing the
*Agentic Workflow for Academic Project Report Generation v2* specification.

## What it does

| Phase | Steps |
|-------|-------|
| **0 — Source analysis** | Extracts formatting rules from the guidelines PDF (Stream A), inspects the sample cover page and sample chapter DOCX with the vendored `inspect_docx` harness (Stream C), reads the user's source material (Stream B) |
| **Anchor reconciliation** | Merges Stream A + Stream C into a frozen `spec` of Anchor Points; every disagreement becomes a logged conflict (Stream C wins per the precedence rule); unresolvable anchors are listed as `missing` |
| **Gate 1 — spec approval** | Workflow pauses (`wait_for_input`, label `spec-approval`); the user reviews conflicts, fills missing anchors via `anchor_overrides`, approves |
| **1 — Content generation** | Plans the chapter/section outline, then drafts each chapter honoring the `CONTENT_MODE` switch (`REFORMAT_ONLY` / `GENERATE` / `EXPAND`), enforces the 150-word section rule mode-aware, and runs a content-accuracy cross-reference gate that flags fabricated algorithms/metrics/figures (auto-redraft up to 2×, escalates what remains) |
| **Gate 2 — per-chapter approval** | Optional (`require_chapter_approval`) |
| **2 — DOCX generation** | Builds each chapter DOCX with the three-tier page structure (cover / first content page / subsequent pages), independent header-footer XML parts per section, CT_SectPr-ordered children, three-line academic tables |
| **3 — Verification loop** | Re-inspects each generated file against the Anchor spec (page size, margins, fonts, header/footer, page numbering, depth); deterministic fixes run up to `max_fix_iterations` (3) |
| **Gate 3 — final approval** | Optional final human review before completion |

The result contains the generated `Chapter-N.docx` files (base64), the frozen spec,
verification and accuracy reports, and a full worklog.

## Repository layout

```
apps/worker/
  pyproject.toml
  src/worker/
    entrypoints/worker.py        # worker bootstrap (auto-discovers workflows)
    workflows/
      seminar_report.py          # the workflow: orchestration + human gates
      extraction.py              # Stream A/B/C extraction + chapter planning
      anchors.py                 # Anchor Points schema, merge, precedence
      content.py                 # drafting, accuracy gate, depth rule
      docxgen.py                 # three-tier DOCX builder
      verification.py            # compliance verification + fix pass
      inspect_docx.py            # vendored dynamic docx inspection harness
      llm.py                     # workflow-aware Mistral client helper
deploy/docker/Dockerfile.worker # managed-deployment image
examples/
  execute_example.py             # start an execution with the sample inputs
  test_offline_core.py           # offline test of the deterministic core
```

## Run locally

```bash
cd apps/worker
uv sync
PYTHONPATH=src uv run python -m worker.entrypoints.worker
```

Then start an execution (see `examples/execute_example.py`) or trigger it from the
Mistral Console. Human gates appear as pending update tasks in the execution's
streaming events; complete them from Studio or via the SDK's update handler.

## Deploy (managed)

1. Push this repository to GitHub.
2. Create a managed deployment pointing at `deploy/docker/Dockerfile.worker`
   (via Studio or `mistral_mcp_deploy_workflow_deployment`).
3. Execute `major-project-seminar-report` with a `ReportInput` payload.

## Input summary

```jsonc
{
  "project": {"name": "...", "type": "major-project", "institute": "...",
               "department": "...", "academic_year": "..."},
  "content_mode": "GENERATE",            // REFORMAT_ONLY | GENERATE | EXPAND
  "guidelines_pdf": {"filename": "...", "content_b64": "..."},      // Stream A
  "sample_cover_docx": {"filename": "...", "content_b64": "..."},   // Stream C
  "sample_chapter_docx": {"filename": "...", "content_b64": "..."}, // Stream C
  "source_content_docx": {"filename": "...", "content_b64": "..."}, // Stream B
  "source_content_text": "or pasted text",
  "chapter_list": ["Introduction", "..."],   // optional, default 5-chapter structure
  "project_details": {"algorithms_used": [...], "metrics": {...},
                      "figures": [...], "tables": [...], "dataset": "...",
                      "references": [...]},
  "require_spec_approval": true,
  "require_chapter_approval": false,
  "require_final_approval": true,
  "parallel_chapters": false
}
```

The workflow never invents factual content: anything not present in the inputs is
flagged as `[NEEDS USER INPUT: ...]` or escalated at a gate.
