"""Run the report workflow and persist every artifact to this repository.

Built for stateless environments (GitHub Codespaces): while the execution is
running, a checkpoint is committed to the persistence branch every --interval
minutes, and the generated DOCX files plus the final result are committed when
the execution finishes. Workflow state itself is durable on the Mistral
platform — if the Codespace restarts mid-run, re-run with --execution-id to
re-attach and keep waiting; nothing is lost.

Usage:
  python run_and_persist.py guidelines.pdf cover.docx chapter1.docx \
      [--source source.docx] [--mode GENERATE] [--interval 10] [--branch runs]

  # re-attach to an existing execution after a restart:
  python run_and_persist.py --execution-id <id> [--interval 10] [--branch runs]

Requires MISTRAL_API_KEY and a running worker in the same namespace
(python -m worker.entrypoints.worker). The persistence branch defaults to
`runs` so `main` stays clean; artifacts land under runs/<execution_id>/.
"""

import argparse
import base64
import json
import os
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

RUNS_DIR = Path("runs")


def file_input(path: str) -> dict:
    with open(path, "rb") as handle:
        return {"filename": os.path.basename(path),
                "content_b64": base64.b64encode(handle.read()).decode("ascii")}


def git(*args, check=True):
    return subprocess.run(["git", *args], check=check, capture_output=True, text=True)


def ensure_git_identity() -> None:
    result = git("config", "user.email", check=False)
    if result.returncode != 0 or not result.stdout.strip():
        git("config", "user.email", "run-and-persist@users.noreply.github.com", check=False)
        git("config", "user.name", "run-and-persist", check=False)


def switch_to_branch(branch: str) -> None:
    if git("rev-parse", "--verify", branch, check=False).returncode == 0:
        git("checkout", branch)
    else:
        git("checkout", "-b", branch)


def persist(paths: list[Path], message: str, branch: str) -> bool:
    """Commit the given paths on the persistence branch and push. Best-effort."""
    existing = [p for p in paths if p.exists()]
    if not existing:
        return False
    try:
        ensure_git_identity()
        switch_to_branch(branch)
        git("add", "--", *[str(p) for p in existing])
        if git("commit", "-m", message, check=False).returncode != 0:
            return True  # nothing new to commit
        if git("push", "origin", branch, check=False).returncode != 0:
            git("pull", "--rebase", "origin", branch, check=False)
            if git("push", "origin", branch, check=False).returncode != 0:
                print("[WARN] push failed; will retry at next checkpoint", file=sys.stderr)
                return False
        return True
    except subprocess.CalledProcessError as error:
        print(f"[WARN] persist failed: {(error.stderr or error).strip()}", file=sys.stderr)
        return False


def write_checkpoint(run_dir: Path, execution_id: str, started: float) -> Path:
    elapsed = (time.time() - started) / 60
    checkpoint = run_dir / "checkpoint.md"
    checkpoint.write_text(
        f"# checkpoint {execution_id}\n\n"
        f"- time: {datetime.now(timezone.utc).isoformat()}\n"
        f"- elapsed_min: {elapsed:.1f}\n"
        f"- note: workflow state is durable on the Mistral platform; "
        f"re-attach with --execution-id {execution_id} if this environment restarts\n"
    )
    return checkpoint


def extract_result(response) -> dict:
    result = getattr(response, "result", None)
    if hasattr(result, "model_dump"):
        result = result.model_dump()
    return result or {}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("guidelines_pdf", nargs="?")
    parser.add_argument("sample_cover_docx", nargs="?")
    parser.add_argument("sample_chapter_docx", nargs="?")
    parser.add_argument("--source", help="user's source content DOCX (Stream B)")
    parser.add_argument("--mode", default="GENERATE",
                        choices=["REFORMAT_ONLY", "GENERATE", "EXPAND"])
    parser.add_argument("--interval", type=float, default=10.0,
                        help="minutes between checkpoint commits (default 10)")
    parser.add_argument("--poll", type=float, default=30.0,
                        help="seconds between status polls (default 30)")
    parser.add_argument("--branch", default="runs",
                        help="branch that receives the artifacts (default runs)")
    parser.add_argument("--execution-id", help="re-attach to an existing execution")
    args = parser.parse_args()

    from mistralai.client import Mistral

    client = Mistral(api_key=os.environ["MISTRAL_API_KEY"])

    if args.execution_id:
        execution_id = args.execution_id
    else:
        if not (args.guidelines_pdf and args.sample_cover_docx and args.sample_chapter_docx):
            parser.error("guidelines_pdf, sample_cover_docx and sample_chapter_docx "
                         "are required (or pass --execution-id)")
        payload = {
            "project": {
                "name": os.environ.get("REPORT_PROJECT_NAME", "YOUR PROJECT TITLE"),
                "type": "major-project",
                "institute": os.environ.get("REPORT_INSTITUTE", "YOUR INSTITUTE"),
                "department": os.environ.get("REPORT_DEPARTMENT", "YOUR DEPARTMENT"),
                "academic_year": os.environ.get("REPORT_ACADEMIC_YEAR", "2024-25"),
            },
            "content_mode": args.mode,
            "guidelines_pdf": file_input(args.guidelines_pdf),
            "sample_cover_docx": file_input(args.sample_cover_docx),
            "sample_chapter_docx": file_input(args.sample_chapter_docx),
            "source_content_docx": file_input(args.source) if args.source else None,
        }
        execution = client.workflows.execute_workflow(
            workflow_identifier="major-project-seminar-report", input=payload)
        execution_id = execution.execution_id
        print(f"started execution {execution_id}")

    run_dir = RUNS_DIR / execution_id
    run_dir.mkdir(parents=True, exist_ok=True)

    if not args.execution_id:
        meta = {
            "execution_id": execution_id,
            "content_mode": args.mode,
            "guidelines_pdf": os.path.basename(args.guidelines_pdf),
            "sample_cover_docx": os.path.basename(args.sample_cover_docx),
            "sample_chapter_docx": os.path.basename(args.sample_chapter_docx),
            "source": os.path.basename(args.source) if args.source else None,
            "started_at": datetime.now(timezone.utc).isoformat(),
        }
        (run_dir / "input.json").write_text(json.dumps(meta, indent=2))
        persist([run_dir / "input.json"], f"run started: {execution_id}", args.branch)

    started = time.time()
    last_commit = 0.0
    while True:
        response = client.workflows.executions.get_workflow_execution(execution_id=execution_id)
        status = response.status
        if status != "RUNNING":
            break
        if time.time() - last_commit >= args.interval * 60:
            checkpoint = write_checkpoint(run_dir, execution_id, started)
            if persist([checkpoint], f"checkpoint: {execution_id} ({(time.time() - started) / 60:.0f} min)", args.branch):
                print(f"[{datetime.now(timezone.utc):%H:%M}] checkpoint pushed to {args.branch}")
            last_commit = time.time()
        time.sleep(args.poll)

    result = extract_result(response)
    (run_dir / "result.json").write_text(json.dumps(result, indent=2, default=str))
    saved = [run_dir / "result.json"]
    for artifact in result.get("files", []):
        out = run_dir / artifact["filename"]
        out.write_bytes(base64.b64decode(artifact["content_b64"]))
        saved.append(out)
    persist(saved, f"run {status}: {execution_id}", args.branch)
    print(f"status={status}; artifacts in {run_dir} (branch {args.branch})")


if __name__ == "__main__":
    main()
