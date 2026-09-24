"""State-machine decisions without deriving pages from an outline."""
from __future__ import annotations

from pathlib import Path
import hashlib
import json

from media_assets import scan_deck_media
from outline import outline_digest, outline_path, read_outline
from project import read_deck, slide_path
from slide_html import Slide, parse_slide, style_advice
from state import input_manifest, read_state


AUTHORING_RULE = "每次只完成一张真实页面。先按照 references/components.md 确定聚焦、比较、顺序、汇聚、关系、证据或数据中的主要关系，再选择 starter。必须替换全部示例文案、数值、来源和占位视觉，让 DOM/CSS 服务当前页的真实判断；完成并检查当前页后才能继续下一页。"


def _authoring_brief(project: Path) -> str:
    reference = read_outline(project).strip()
    return f"{reference}\n\n{AUTHORING_RULE}" if reference else AUTHORING_RULE


def _authoring_next(project: Path, command: callable) -> dict:
    return {
        "action": "author_slides",
        "brief": _authoring_brief(project),
        "slide_add_usage": (
            f"{command('slide', 'add', project, '<页面ID>', '--title', '<标题>')} "
            "[--starter <名称>] [--after <页面ID>]"
        ),
        "command_when_ready": command("preview", project),
    }


def _current_delivery(project: Path, state: dict, manifest: dict, *, slide_count: int) -> bool:
    """Completion is evidence for these exact inputs and delivered file bytes."""
    if state.get("build_manifest") != manifest or state.get("pptx_manifest") != manifest:
        return False
    files = {
        "build": state.get("build"),
        "pptx": state.get("pptx"),
        "pptx_editability": state.get("pptx_editability"),
    }
    try:
        for key, value in files.items():
            if not isinstance(value, str) or not value:
                return False
            path = Path(value)
            if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != state.get(f"{key}_sha256"):
                return False
        source = json.loads(Path(files["pptx_editability"]).read_text(encoding="utf-8"))
        if not (
            isinstance(source, dict)
            and source.get("ok")
            and source.get("pptx_sha256") == state.get("pptx_sha256")
            and source.get("canonical_html_sha256") == state.get("build_sha256")
            and isinstance(source.get("summary"), dict)
            and source["summary"].get("slide_count") == slide_count
            and slide_count > 0
        ):
            return False
    except (OSError, ValueError, TypeError):
        return False
    return True


def status(project_value: Path, command: callable, *, intent: str = "continue", slide: str | None = None) -> dict:
    project, deck = read_deck(project_value)
    state = read_state(project)
    current_manifest = input_manifest(project)
    slide_validation: tuple[tuple[dict, str] | None, list[Slide]] | None = None
    if intent == "edit":
        if slide:
            try:
                index = int(slide) - 1
                relative = deck["slides"][index] if index >= 0 else None
            except (ValueError, IndexError):
                relative = next((item for item in deck["slides"] if Path(item).stem == slide), None)
            if not relative:
                raise SystemExit(f"Unknown slide: {slide}")
            next_step = {"action": "edit_slide", "command": None, "path": str(slide_path(project, relative)), "issues": [], "rerun": command("status", project, "--json")}
            phase = "editing_slide"
        else:
            next_step = {"action": "edit_outline", "command": None, "path": str(outline_path(project)), "rerun": command("status", project, "--json")}
            phase = "editing_outline"
    elif not deck["slides"] and state.get("outline_seed_sha256") == outline_digest(project):
        next_step = {
            "action": "edit_outline",
            "path": str(outline_path(project)),
            "brief": "Write a concise creative reference from the request: audience, setting, core claim, evidence, and a possible narrative. Continue directly to authoring with reasonable assumptions for missing optional details.",
            "rerun": command("status", project, "--json"),
        }
        phase = "edit_outline"
    elif not deck["slides"]:
        next_step = _authoring_next(project, command)
        phase = "author_slides"
    elif (slide_validation := _validate_slides(project, deck, command))[0] is not None:
        next_step, phase = slide_validation[0]
    elif state.get("browser_manifest") == current_manifest and isinstance(state.get("browser_issues"), dict):
        report = state["browser_issues"]
        findings = [
            *report.get("missingSafeArea", []),
            *report.get("invalidLayouts", []),
            *report.get("invalidBleeds", []),
            *report.get("invalidText", []),
            *report.get("visualFindings", []),
        ]
        first = next((item for item in findings if isinstance(item, dict)), {"reason": "browser-validation"})
        slide_id = str(first.get("slide") or Path(deck["slides"][0]).stem)
        relative = next((item for item in deck["slides"] if Path(item).stem == slide_id), deck["slides"][0])
        next_step = {
            "action": "edit_slide",
            "path": str(slide_path(project, relative)),
            "issues": findings or [report],
            "rerun": command("preview", project),
        }
        phase = "repair_browser"
    elif not (project / "预览.html").is_file() or state.get("preview_manifest") != current_manifest:
        # Only the author can determine whether all requested material is covered.
        # Creating one page must not send an open-ended deck to final delivery.
        next_step = _authoring_next(project, command)
        phase = "author_slides"
    elif state.get("phase") == "complete" and _current_delivery(project, state, current_manifest, slide_count=len(deck["slides"])):
        next_step = {"action": "complete", "command": None}
        phase = "complete"
    else:
        next_step = {"action": "run_command", "command": command("build", project), "brief": "Build the current HTML and editable PPTX, and preserve their matching editability metadata."}
        phase = "ready_to_build"
    return {
        "schema_version": "pptx.status/v1",
        "ok": True,
        "project": str(project),
        "phase": phase,
        "slides": len(deck["slides"]),
        "style_advice": style_advice(
            project,
            deck,
            slides=slide_validation[1] if slide_validation is not None else None,
        ),
        "next": next_step,
    }


def _validate_slides(
    project: Path,
    deck: dict,
    command: callable,
) -> tuple[tuple[dict, str] | None, list[Slide]]:
    slides: list[Slide] = []
    first_issue: tuple[dict, str] | None = None
    for relative in deck["slides"]:
        path = slide_path(project, relative)
        try:
            slides.append(parse_slide(path, Path(relative).stem))
        except SystemExit as error:
            if first_issue is not None:
                continue
            message = str(error)
            action = "fix_media" if any(
                phrase in message
                for phrase in ("asset", "URL", "path escapes slide assets")
            ) else "edit_slide"
            first_issue = ({
                "action": action,
                "path": str(path),
                "issues": [message],
                "rerun": command("status", project, "--json"),
            }, "repair_media" if action == "fix_media" else "repair_slide")
    if first_issue is not None:
        return first_issue, slides
    media = scan_deck_media(project, deck)
    errors = media.get("errors") or []
    if errors:
        first_file = str(errors[0].get("file") or deck["slides"][0])
        page_errors = [item for item in errors if item.get("file") == first_file]
        return ({
            "action": "fix_media",
            "path": str((project / first_file).resolve()),
            "issues": page_errors,
            "rerun": command("status", project, "--json"),
        }, "repair_media"), slides
    return None, slides


def batch(projects: list[Path], command: callable) -> dict:
    payloads = [status(project, command) for project in projects]
    active = [item["next"] for item in payloads if item["next"]["action"] != "complete"]
    next_step = active[0] if active else {"action": "complete", "command": None}
    return {"schema_version": "pptx.batch/v1", "ok": True, "projects": payloads, "next": next_step}
