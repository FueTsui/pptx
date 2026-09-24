#!/usr/bin/env python3
"""Maintenance-only skill audit: structure, ordinary generation, rendered regression.

Run after changing the skill. Ordinary presentation build/export never calls this
runner and does not produce its audit reports or PowerPoint regression renders.
"""
from __future__ import annotations

import argparse
import contextlib
from datetime import datetime, timezone
import hashlib
import html
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import uuid

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write(path: Path, value: dict) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _source_hashes(project: Path) -> dict:
    return {path.relative_to(project).as_posix(): _sha(path) for path in sorted((project / "slides").glob("*.html"))}


def _fixture_slide(slide_id: str, title: str, heading: str, body: str, detail: str) -> str:
    escaped_title = html.escape(title, quote=True)
    return f'''<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{escaped_title}</title>
<link rel="stylesheet" href="../runtime/deck.css"><link rel="stylesheet" href="../runtime/theme.css"><style>
/* PPTX-SLIDE-CSS:START */
.s-{slide_id} .audit-content{{display:flex;flex-direction:column;gap:62px;height:100%}}
.s-{slide_id} h1,.s-{slide_id} h2,.s-{slide_id} p{{margin:0}}
.s-{slide_id} .audit-body{{display:flex;flex-direction:column;gap:32px}}
.s-{slide_id} .audit-divider{{height:5px;width:200px;background:var(--accent)}}
.s-{slide_id} .audit-footer{{margin-top:auto}}
/* PPTX-SLIDE-CSS:END */
</style></head><body data-mode="preview"><div class="slide-preview-viewport"><div class="slide-preview-shell"><div class="slide-preview-stage">
<!-- PPTX-SLIDE:START -->
<section class="slide s-{slide_id}" data-slide-id="{slide_id}" data-title="{escaped_title}"><div class="slide-safe"><main class="audit-content" data-layout>
<h1 data-text-role="title">{html.escape(title)}</h1><div class="audit-divider"></div>
<div class="audit-body"><h2 data-text-role="heading">{html.escape(heading)}</h2><p data-text-role="body">{body}</p><p data-text-role="body">{detail}</p></div>
<p class="audit-footer" data-text-role="caption">中文生成器验证样稿 · 用于技能维护</p>
</main></div></section>
<!-- PPTX-SLIDE:END -->
</div></div></div><script src="../runtime/deck.js"></script></body></html>
'''


def _cli(args: list[str], run_dir: Path, sequence: int, env: dict) -> dict:
    command = [sys.executable, "-X", "utf8", str(SCRIPTS / "pptx_cli.py"), *args, "--compact"]
    completed = subprocess.run(command, cwd=run_dir, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=240)
    log = run_dir / f"command-{sequence:02d}.json"
    try:
        payload = json.loads(completed.stdout)
    except ValueError:
        payload = None
    evidence = {"command": command, "exit_code": completed.returncode, "stdout": completed.stdout, "stderr": completed.stderr, "result": payload}
    _write(log, evidence)
    evidence["log"] = str(log)
    if completed.returncode != 0 or not isinstance(payload, dict) or not payload.get("ok", True):
        raise RuntimeError(f"Ordinary CLI {' '.join(args[:2])} failed; evidence: {log}")
    return evidence


def audit_skill(output: str | Path, *, chrome: str | None = None) -> dict:
    """Audit the installed skill without changing its sources or package manifest."""
    output = Path(output).expanduser().resolve()
    if output == ROOT or output.is_relative_to(ROOT) or ROOT.is_relative_to(output):
        raise ValueError("Skill maintenance audit output must be outside the skill tree and cannot contain the skill root")
    output.mkdir(parents=True, exist_ok=True)
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ-") + uuid.uuid4().hex[:8]
    run_dir = output / "runs" / run_id; run_dir.mkdir(parents=True)
    report_path = output / "skill-audit.json"
    env = os.environ.copy(); env["PYTHONDONTWRITEBYTECODE"] = "1"
    if chrome:
        env["CHROME_BIN"] = str(Path(chrome).expanduser().resolve())
    sys.dont_write_bytecode = True
    from package_manifest import collect_files, compute_tree_sha256, verify_manifest
    from validate_skill import validation_errors
    from pptx_audit import audit_pptx
    from test_pptx_audit import run as run_regression

    initial_files = collect_files(ROOT)
    initial_tree = compute_tree_sha256(initial_files)
    generators = [record for record in initial_files if str(record["path"]).startswith(("scripts/", "assets/runtime/"))]
    report = {"schema_version": "pptx.skill-audit/v1", "purpose": "skill-maintenance", "created_at": datetime.now(timezone.utc).isoformat(),
              "ok": False, "status": "not_verified", "skill_root": str(ROOT), "report": str(report_path), "run_directory": str(run_dir),
              "skill_tree_sha256_before": initial_tree, "tested_generator_files": generators, "issues": [], "structure": {},
              "end_to_end": {"status": "not_verified"}, "regression": {"status": "not_verified"}, "artifacts": {}}
    blockers = []
    try:
        errors = validation_errors()
        package = verify_manifest(root=ROOT)
        report["structure"] = {"ok": not errors and bool(package["ok"]), "skill_errors": errors, "package_manifest": package}
        if errors or not package["ok"]:
            report["issues"].append({"code": "skill_structure_failed", "message": "Skill structure or package manifest validation failed; this runner never rewrites the manifest."})
    except Exception as error:
        report["issues"].append({"code": "structure_check_failed", "message": str(error)})

    project = run_dir / "fixture"
    try:
        steps = []
        steps.append(_cli(["init", str(project), "--title", "中文生成器维护验证"], run_dir, 1, env))
        data = [
            ("hierarchy", "中文文字层级保持一致", "标题和正文使用固定样式", "第一行：标题说明本页主题。<br>第二行：正文保持统一字号。<br>第三行：<strong>重点文字</strong>只通过粗体突出。", "相同层级在另一页继续复用字体、颜色、行距和字距。"),
            ("multiline", "多行文本保留足够间距", "混合文字导出为原生文本", "第一步：阅读输入材料与页面主题。<br>第二步：保留<strong>中文与 English</strong>混合段落。<br>第三步：检查每行可读且互不遮挡。", "维护验证覆盖换行、内联强调、跨页样式和真实 PowerPoint 排版。"),
        ]
        for index, (slide_id, title, heading, body, detail) in enumerate(data, 2):
            steps.append(_cli(["slide", "add", str(project), slide_id, "--title", title], run_dir, index, env))
            (project / "slides" / f"{slide_id}.html").write_text(_fixture_slide(slide_id, title, heading, body, detail), encoding="utf-8")
        hashes_before = _source_hashes(project)
        steps.append(_cli(["build", str(project)], run_dir, 4, env))
        steps.append(_cli(["export-pptx", str(project)], run_dir, 5, env))
        ordinary_audit_outputs = [str(path) for path in project.iterdir() if "audit" in path.name.lower() or path.name == "pptx-render"]
        hashes_after = _source_hashes(project)
        report["end_to_end"] = {"status": "generated", "project": str(project), "commands": [{"command": step["command"], "exit_code": step["exit_code"], "log": step["log"]} for step in steps],
                                "input_html_hashes_before": hashes_before, "input_html_hashes_after": hashes_after, "input_html_unchanged": hashes_before == hashes_after,
                                "ordinary_build_creates_no_audit": not ordinary_audit_outputs, "unexpected_ordinary_audit_artifacts": ordinary_audit_outputs}
        if hashes_before != hashes_after:
            report["issues"].append({"code": "input_html_modified", "message": "Ordinary generation changed the authored fixture HTML."})
        if ordinary_audit_outputs:
            report["issues"].append({"code": "ordinary_generation_audits", "message": "Ordinary build/export must not invoke maintenance audits or generate audit artifacts."})
        pptx_path = project / "演示文稿.pptx"
        fixture_report = run_dir / "fixture-pptx-audit.json"
        measured = audit_pptx(pptx_path, report_path=fixture_report, render_dir=run_dir / "fixture-render", project=project, source_report=project / "pptx-editability.json", require_render=True)
        report["end_to_end"].update({"status": measured["status"], "ok": measured["ok"], "render": measured["render"], "audit_report": str(fixture_report), "audit_issue_counts": measured["issue_counts"], "pptx_input_unchanged": measured["input_unchanged"]})
        report["artifacts"].update({"pptx": str(pptx_path), "html": str(project / "演示文稿.html"), "source_layout": str(project / "pptx-editability.json"), "fixture_audit": str(fixture_report), "rendered_pages": [page["image"] for page in measured["render"].get("slides", [])]})
        if measured["render"]["status"] == "not_verified":
            blockers.append({"code": "renderer_unavailable", "message": measured["render"].get("reason", "PowerPoint renderer unavailable")})
        elif not measured["ok"]:
            report["issues"].append({"code": "generator_render_failed", "message": "The ordinary generator fixture did not pass explicit maintenance rendering and text audits.", "report": str(fixture_report)})
    except Exception as error:
        report["issues"].append({"code": "generator_end_to_end_failed", "message": str(error)})

    try:
        regression_log = io.StringIO()
        with contextlib.redirect_stdout(regression_log):
            regression = run_regression(run_dir / "regression")
        (run_dir / "regression-console.txt").write_text(regression_log.getvalue(), encoding="utf-8")
        report["regression"] = {"ok": regression["ok"], "status": "passed" if regression["ok"] else "failed", "cases": regression["cases"], "report": str(run_dir / "regression" / "regression-results.json")}
        report["artifacts"]["regression_report"] = report["regression"]["report"]
        if not regression["ok"]:
            report["issues"].append({"code": "regression_failed", "message": "One or more positive/negative maintenance regression cases failed."})
    except Exception as error:
        report["issues"].append({"code": "regression_check_failed", "message": str(error)})

    final_files = collect_files(ROOT)
    final_tree = compute_tree_sha256(final_files)
    report["skill_tree_sha256_after"] = final_tree
    report["skill_source_unchanged"] = final_tree == initial_tree
    if final_tree != initial_tree:
        report["issues"].append({"code": "skill_changed_during_audit", "message": "Skill files changed during maintenance audit; regenerate the manifest and rerun against a stable version."})
    report["blockers"] = blockers
    report["ok"] = not report["issues"] and not blockers and report["structure"].get("ok", False) and report["end_to_end"].get("ok", False) and report["regression"].get("ok", False)
    report["status"] = "passed" if report["ok"] else "blocked" if blockers else "failed"
    report["summary"] = {"structure_passed": bool(report["structure"].get("ok")), "end_to_end_passed": bool(report["end_to_end"].get("ok")), "regression_cases": len(report["regression"].get("cases", [])), "regression_passed": bool(report["regression"].get("ok")), "skill_source_unchanged": report["skill_source_unchanged"], "ordinary_build_creates_no_audit": report["end_to_end"].get("ordinary_build_creates_no_audit", False)}
    _write(run_dir / "skill-audit.json", report)
    _write(report_path, report)
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True)
    parser.add_argument("--chrome")
    args = parser.parse_args()
    report = audit_skill(args.output, chrome=args.chrome)
    print(json.dumps({key: report[key] for key in ("ok", "status", "report", "summary")}, ensure_ascii=False))
    return 0 if report["ok"] else 2 if report["status"] in ("blocked", "not_verified") else 1


if __name__ == "__main__":
    raise SystemExit(main())
