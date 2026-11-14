"""Major Project Technical Seminar report — durable report-generation workflow.

Implements the Agentic Workflow v2 specification (three input streams, Anchor
Points, CONTENT_MODE switch, accuracy gate, three-tier DOCX structure,
verification loop) as a Mistral Workflow with human-approval gates.
"""

import asyncio
import copy
import json
from datetime import timedelta
from typing import Any, Literal

import mistralai.workflows as workflows
from mistralai.workflows import InteractiveWorkflow
from pydantic import BaseModel, Field

from .content import accuracy_gate, check_depth_rule, draft_chapter
from .docxgen import generate_chapter_docx
from .extraction import (
    extract_guidelines_anchors,
    extract_source_text,
    inspect_sample_docx,
    plan_chapters,
    reconcile_spec,
)
from .verification import apply_format_fixes, verify_chapter_docx


class ProjectInfo(BaseModel):
    name: str
    type: str = "major-project"
    institute: str | None = None
    department: str | None = None
    academic_year: str | None = None


class InputFile(BaseModel):
    filename: str
    content_b64: str


class ProjectDetails(BaseModel):
    model_config = {"extra": "allow"}

    algorithms_used: list[str] | None = None
    metrics: dict[str, Any] | None = None
    figures: list[str] | None = None
    tables: list[str] | None = None
    dataset: str | None = None
    references: list[str] | None = None


class ReportInput(BaseModel):
    project: ProjectInfo
    content_mode: Literal["REFORMAT_ONLY", "GENERATE", "EXPAND"] = "GENERATE"
    guidelines_pdf: InputFile | None = None
    sample_cover_docx: InputFile | None = None
    sample_chapter_docx: InputFile | None = None
    source_content_docx: InputFile | None = None
    source_content_text: str | None = None
    chapter_list: list[str] | None = None
    project_details: ProjectDetails = Field(default_factory=ProjectDetails)
    require_spec_approval: bool = True
    require_chapter_approval: bool = False
    require_final_approval: bool = True
    gate_timeout_hours: float = 72.0
    min_words_per_section: int = 150
    max_fix_iterations: int = 3
    parallel_chapters: bool = False


class SpecApproval(BaseModel):
    approved: bool
    anchor_overrides: dict[str, Any] = {}
    reviewer: str = ""
    comments: str = ""


class ChapterApproval(BaseModel):
    approved: bool
    revised_text: str | None = None
    comments: str = ""


class FinalApproval(BaseModel):
    approved: bool
    comments: str = ""


@workflows.workflow.define(
    name="major-project-seminar-report",
    workflow_display_name="Major Project Technical Seminar Report",
    workflow_description=(
        "Generates institute-compliant chapter DOCX files for a Major Project / "
        "Technical Seminar report from the guidelines PDF, sample DOCX files and the "
        "user's project material, with human-approval gates and a verification loop."
    ),
    execution_timeout=timedelta(days=7),
)
class SeminarReportWorkflow(InteractiveWorkflow):

    @workflows.workflow.entrypoint
    async def run(self, input: ReportInput) -> dict:
        worklog: list[dict] = []

        def log(step: str, detail: Any) -> None:
            worklog.append({"step": step, "detail": detail})

        project = input.project.model_dump()
        project["actual_results"] = input.project_details.model_dump(exclude_none=True)

        # ---- Phase 0: source analysis (L0) --------------------------------
        source_text = input.source_content_text
        if input.source_content_docx is not None and not source_text:
            extracted = await extract_source_text(input.source_content_docx.content_b64)
            source_text = extracted["source_text"]
        source_text = source_text or ""

        stream_a: dict = {}
        if input.guidelines_pdf is not None:
            result = await extract_guidelines_anchors(input.guidelines_pdf.content_b64)
            stream_a = result["stream_a_anchors"]
            log("stream_a_extracted", {"anchors": len(stream_a)})

        cover_inspect = {"kind": "cover", "inspect": {}}
        chapter_inspect = {"kind": "chapter", "inspect": {}}
        if input.sample_cover_docx is not None:
            cover_inspect = await inspect_sample_docx(input.sample_cover_docx.content_b64, "cover")
        if input.sample_chapter_docx is not None:
            chapter_inspect = await inspect_sample_docx(input.sample_chapter_docx.content_b64, "chapter")
        log("stream_c_inspected", {"cover": bool(input.sample_cover_docx), "chapter": bool(input.sample_chapter_docx)})

        reconciled = await reconcile_spec(
            stream_a, cover_inspect, chapter_inspect, project,
            input.content_mode, input.chapter_list,
        )
        spec = reconciled["spec"]
        log("spec_reconciled", {
            "conflicts": reconciled["conflicts"],
            "missing": reconciled["missing"],
        })

        # ---- Gate 1: spec approval (L6) -----------------------------------
        if input.require_spec_approval:
            try:
                approval = await self.wait_for_input(
                    SpecApproval,
                    label="spec-approval",
                    timeout=timedelta(hours=input.gate_timeout_hours),
                )
            except asyncio.TimeoutError:
                return self._terminate(worklog, "timed_out_at_spec_gate")
            if not approval.approved:
                return self._terminate(worklog, "rejected_at_spec_gate", approval.comments)
            for anchor_path, value in approval.anchor_overrides.items():
                group, key = anchor_path.split(".", 1)
                if group in spec and key in spec[group]:
                    spec[group][key] = value
            log("spec_approved", {"overrides": list(approval.anchor_overrides)})

        # ---- Phase 1: content generation (L1) -----------------------------
        plan = await plan_chapters(spec, source_text, project["actual_results"])
        chapters = plan.get("chapters", [])
        log("plan_created", {"chapters": [c.get("title") for c in chapters]})

        async def produce_content(chapter_plan: dict) -> dict:
            draft = await draft_chapter(chapter_plan, source_text, spec, project["actual_results"])
            depth = await check_depth_rule(draft, input.min_words_per_section)
            if depth["action"] == "EXPAND_SECTIONS":
                expanded_plan = copy.deepcopy(chapter_plan)
                expanded_plan["content_hint"] = (
                    "EXPAND these sections to at least "
                    f"{input.min_words_per_section} words each: "
                    + json.dumps(depth["short_sections"])
                )
                draft = await draft_chapter(expanded_plan, source_text, spec, project["actual_results"])
                depth = await check_depth_rule(draft, input.min_words_per_section)

            audit = await accuracy_gate(draft, project["actual_results"], source_text)
            retries = 0
            while not audit["passed"] and input.content_mode != "REFORMAT_ONLY" and retries < 2:
                retries += 1
                chapter_plan = copy.deepcopy(chapter_plan)
                chapter_plan["accuracy_violations_to_fix"] = audit["violations"]
                draft = await draft_chapter(chapter_plan, source_text, spec, project["actual_results"])
                audit = await accuracy_gate(draft, project["actual_results"], source_text)

            return {"draft": draft, "depth": depth, "audit": audit, "accuracy_retries": retries}

        if input.parallel_chapters and len(chapters) > 1:
            produced = await asyncio.gather(*(produce_content(c) for c in chapters))
        else:
            produced = [await produce_content(c) for c in chapters]

        for item, chapter_plan in zip(produced, chapters):
            log("chapter_drafted", {
                "chapter": chapter_plan.get("title"),
                "accuracy_passed": item["audit"]["passed"],
                "depth_action": item["depth"]["action"],
            })

        # ---- Gate 2 (optional): per-chapter content approval ---------------
        if input.require_chapter_approval:
            for item, chapter_plan in zip(produced, chapters):
                try:
                    approval = await self.wait_for_input(
                        ChapterApproval,
                        label=f"chapter-{chapter_plan.get('chapter_num')}-approval",
                        timeout=timedelta(hours=input.gate_timeout_hours),
                    )
                except asyncio.TimeoutError:
                    return self._terminate(worklog, f"timed_out_at_chapter_{chapter_plan.get('chapter_num')}_gate")
                if approval.revised_text:
                    item["draft"]["intro_paragraph"] = approval.revised_text
                if not approval.approved:
                    item["needs_review"] = True
                log("chapter_reviewed", {"chapter": chapter_plan.get("title"), "approved": approval.approved})

        # ---- Phase 2: formatting & file generation (L3) + verification ----
        files: list[dict] = []
        verification: list[dict] = []
        for item, chapter_plan in zip(produced, chapters):
            generated = await generate_chapter_docx(item["draft"], spec)
            report = await verify_chapter_docx(generated["docx_b64"], spec, item["draft"])

            iterations = 0
            while not report["passed"] and iterations < input.max_fix_iterations:
                iterations += 1
                fixed = await apply_format_fixes(generated["docx_b64"], spec, report["violations"])
                generated["docx_b64"] = fixed["docx_b64"]
                report = await verify_chapter_docx(generated["docx_b64"], spec, item["draft"])

            files.append({
                "filename": generated["filename"],
                "content_b64": generated["docx_b64"],
            })
            verification.append({
                "chapter": chapter_plan.get("title"),
                "passed": report["passed"],
                "remaining_violations": report["violations"],
                "fix_iterations": iterations,
            })
            log("chapter_generated", {
                "file": generated["filename"],
                "verification_passed": report["passed"],
                "fix_iterations": iterations,
            })

        # ---- Gate 3: final approval ----------------------------------------
        final_status = "completed"
        if input.require_final_approval:
            try:
                approval = await self.wait_for_input(
                    FinalApproval,
                    label="final-approval",
                    timeout=timedelta(hours=input.gate_timeout_hours),
                )
            except asyncio.TimeoutError:
                return self._terminate(worklog, "timed_out_at_final_gate", files=files)
            if not approval.approved:
                final_status = "changes_requested"
            log("final_review", {"approved": approval.approved, "comments": approval.comments})

        return {
            "status": final_status,
            "workflow_name": "major-project-seminar-report",
            "execution_id": workflows.get_execution_id(),
            "content_mode": input.content_mode,
            "files": files,
            "verification": verification,
            "accuracy_audits": [
                {"chapter": chapter_plan.get("title"),
                 "passed": item["audit"]["passed"],
                 "violations": item["audit"].get("violations", [])}
                for item, chapter_plan in zip(produced, chapters)
            ],
            "depth_reports": [item["depth"] for item in produced],
            "worklog": worklog,
        }

    @staticmethod
    def _terminate(worklog: list[dict], status: str, comments: str | None = None, files: list | None = None) -> dict:
        result = {
            "status": status,
            "workflow_name": "major-project-seminar-report",
            "worklog": worklog,
        }
        if comments:
            result["comments"] = comments
        if files:
            result["files"] = files
        return result
