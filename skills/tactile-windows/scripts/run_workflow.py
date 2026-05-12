#!/usr/bin/env python3
"""Run a WindowsUseSDK LLM app workflow with stable defaults."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_ROOT = SKILL_ROOT / "scripts"
if os.fspath(SCRIPTS_ROOT) not in sys.path:
    sys.path.insert(0, os.fspath(SCRIPTS_ROOT))

from utils import artifacts as artifact_utils

REPO_ENV = "WINDOWS_USE_SDK_ROOT"
DEFAULT_REPO = SKILL_ROOT / "vendor" / "WindowsUseSDK"
default_artifact_path = artifact_utils.default_artifact_path
session_scoped_output_path = artifact_utils.session_scoped_output_path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def sdk_root_from_candidate(candidate: Path) -> Path | None:
    if (candidate / "WindowsUseSDK.ps1").exists():
        return candidate
    for nested in (
        candidate / "vendor" / "WindowsUseSDK",
        candidate / "native" / "WindowsUseSDK",
    ):
        if (nested / "WindowsUseSDK.ps1").exists():
            return nested.resolve()
    return None


def find_repo_root(start: Path) -> Path | None:
    for parent in [start, *start.parents]:
        found = sdk_root_from_candidate(parent)
        if found:
            return found
    return None


def parse_args() -> tuple[argparse.Namespace, list[str]]:
    parser = argparse.ArgumentParser(description="Wrapper for WindowsUseSDK workflows/llm_app_workflow.py.")
    parser.add_argument("instruction", help="Natural-language instruction for the target app.")
    parser.add_argument(
        "--repo",
        default=os.environ.get(REPO_ENV),
        help="Path to the WindowsUseSDK repo. Defaults to the bundled vendor/WindowsUseSDK, then WINDOWS_USE_SDK_ROOT or auto-detection.",
    )
    parser.add_argument("--target", default=None, help="Optional app name, app id, exe name, path, or window title.")
    parser.add_argument("--execute", action="store_true", help="Execute the planned UI actions.")
    parser.add_argument("--dry-run", action="store_true", help="Plan only, even if --execute is omitted.")
    parser.add_argument("--debug-observation", action="store_true", help="Print summarized UI elements.")
    parser.add_argument("--plan-output", default=None, help="Path for the run log JSON.")
    parser.add_argument("--traversal-output", default=None, help="Path for latest raw traversal JSON.")
    parser.add_argument("--max-steps", type=int, default=None, help="Maximum observe-plan-act iterations.")
    parser.add_argument("--max-elements", type=int, default=None, help="Maximum UI elements sent to the LLM.")
    parser.add_argument("--uia-view", choices=["auto", "control", "raw", "content"], default=None, help="UI Automation view to use. Defaults to workflow auto probing.")
    parser.add_argument("--model", default=None, help="Model override passed to workflow.")
    parser.add_argument("--provider", default=None, help="Provider override passed to workflow.")
    args, extra = parser.parse_known_args()
    extra = extra[1:] if extra[:1] == ["--"] else extra
    return args, extra


def main() -> int:
    args, extra = parse_args()
    if args.repo:
        repo = sdk_root_from_candidate(Path(args.repo).expanduser().resolve())
    else:
        repo = (
            sdk_root_from_candidate(DEFAULT_REPO)
            or find_repo_root(Path.cwd().resolve())
            or find_repo_root(Path(__file__).resolve())
        )
    if not repo:
        print(
            f"error: WindowsUseSDK not found; expected bundled {DEFAULT_REPO}, or set --repo / {REPO_ENV}",
            file=sys.stderr,
        )
        return 2

    workflow = repo / "workflows" / "llm_app_workflow.py"
    if not workflow.exists():
        print(f"error: workflow not found: {workflow}", file=sys.stderr)
        print(f"set --repo or {REPO_ENV} to the WindowsUseSDK checkout", file=sys.stderr)
        return 2

    plan_output = args.plan_output
    if args.execute and not plan_output:
        plan_output = str(default_artifact_path("workflow-run", ".json", cwd=Path.cwd()))
    elif plan_output:
        plan_output = str(session_scoped_output_path(Path(plan_output)))

    traversal_output = args.traversal_output
    if traversal_output:
        traversal_output = str(session_scoped_output_path(Path(traversal_output)))

    cmd = [sys.executable, str(workflow), args.instruction]
    if args.target:
        cmd.extend(["--target", args.target])
    if args.execute and not args.dry_run:
        cmd.append("--execute")
    if args.debug_observation:
        cmd.append("--debug-observation")
    if plan_output:
        cmd.extend(["--plan-output", plan_output])
    if traversal_output:
        cmd.extend(["--traversal-output", traversal_output])
    if args.max_steps is not None:
        cmd.extend(["--max-steps", str(args.max_steps)])
    if args.max_elements is not None:
        cmd.extend(["--max-elements", str(args.max_elements)])
    if args.uia_view:
        cmd.extend(["--uia-view", args.uia_view])
    if args.model:
        cmd.extend(["--model", args.model])
    if args.provider:
        cmd.extend(["--provider", args.provider])
    cmd.extend(extra)

    print("running:", " ".join(cmd), file=sys.stderr)
    if plan_output:
        print(f"plan_output: {plan_output}", file=sys.stderr)
    return subprocess.run(cmd, cwd=repo).returncode


if __name__ == "__main__":
    raise SystemExit(main())
