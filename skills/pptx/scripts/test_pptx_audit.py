#!/usr/bin/env python3
"""Regression fixtures: actual positive/negative PPTX audits, isolated outputs."""
from __future__ import annotations
import argparse
import hashlib
import json
import os
from pathlib import Path
import tempfile
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Pt
from pptx_audit import audit_pptx


def fixture(root: Path, mode: str) -> Path:
    directory = root / mode; directory.mkdir(parents=True, exist_ok=True)
    prs = Presentation(); prs.slide_width = Pt(960); prs.slide_height = Pt(540)
    source_slides = []
    for page in range(1, 3):
        slide = prs.slides.add_slide(prs.slide_layouts[6]); layout = []
        # An opaque background behind text must never trigger foreground occlusion.
        background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Pt(960), Pt(540))
        background.fill.solid(); background.fill.fore_color.rgb = RGBColor(255, 255, 255); background.line.fill.background()
        def text(value, x, y, w, h, size, role="body", spacing=None, bold=False):
            ordinal = len(layout) + 1; name = f"pptx-text|{role}|s{page:03}-e{ordinal:04}|l001|r001"
            shape = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h)); shape.name = name
            frame = shape.text_frame; frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
            frame.auto_size = MSO_AUTO_SIZE.NONE; frame.word_wrap = False
            frame.text = value
            spacing = spacing or size * 1.4
            for paragraph in frame.paragraphs:
                paragraph.space_before = paragraph.space_after = Pt(0); paragraph.line_spacing = Pt(spacing)
                for run in paragraph.runs:
                    run.font.name = "Microsoft YaHei"; run.font.size = Pt(size); run.font.bold = bold; run.font.italic = False
                    run.font.color.rgb = RGBColor(26, 40, 55)
            layout.append({"shape_name": name, "text": value, "role": role, "source_id": f"s{page:03}-e{ordinal:04}", "line_id": "l001", "segment_id": "r001", "native": True,
                           "style": {"fontFamily": "Microsoft YaHei", "fontSize": size*2, "fontWeight": 700 if bold else 400, "fontStyle": "normal", "color": "#1A2837", "lineHeight": spacing*2},
                           "box": {"x": x*2, "y": y*2, "width": w*2, "height": h*2}})
            return shape
        text("中文层级与排版验证", 60, 40, 830, 65, 32, "title", bold=True)
        size = 23 if mode == "style_drift" and page == 2 else 20
        body = text("同级文字保持统一样式\n多行文本保持足够的行距", 60, 170, 790, 105, size, spacing=12 if mode == "dense_lines" and page == 1 else None)
        if mode == "text_overlap" and page == 1:
            text("这一行与已有文字相交", 75, 174, 710, 50, 20)
        if mode == "foreground_occlusion" and page == 1:
            cover = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(70), Pt(176), Pt(320), Pt(29))
            cover.fill.solid(); cover.fill.fore_color.rgb = RGBColor(30, 50, 70); cover.line.fill.background()
        if mode == "partial_occlusion" and page == 1:
            text("一二三四五六七八九十一二三四五六七八九十一二三四五六七八九十一二三四五六七八九十", 60, 330, 840, 45, 20)
            cover = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(100), Pt(335), Pt(20), Pt(20))
            cover.fill.solid(); cover.fill.fore_color.rgb = RGBColor(30, 50, 70); cover.line.fill.background()
        if mode == "overflow" and page == 1:
            body.width = Pt(55)
        source_slides.append({"page": page, "slide_id": f"s{page:03}", "text_layout": layout, "unsupported_structures": []})
    output = directory / "fixture.pptx"; prs.save(str(output))
    (directory / "pptx-editability.json").write_text(json.dumps({"pptx_sha256": hashlib.sha256(output.read_bytes()).hexdigest(), "slides": source_slides}, ensure_ascii=False, indent=2), encoding="utf-8")
    return output


def run(root: Path) -> dict:
    root.mkdir(parents=True, exist_ok=True)
    results = []
    expected = {"normal": None, "style_drift": "role_style_drift", "text_overlap": "text_overlap", "dense_lines": "dense_multiline_spacing", "foreground_occlusion": "foreground_occlusion", "partial_occlusion": "foreground_occlusion", "overflow": "text_frame_overflow"}
    for mode, code in expected.items():
        path = fixture(root, mode)
        report = audit_pptx(path, render_dir=path.parent / "render", project=path.parent, require_render=True)
        codes = {i["code"] for i in report["issues"]}
        passed = report["ok"] if code is None else (not report["ok"] and code in codes)
        assert report["input_unchanged"], f"Auditor modified {mode} input"
        results.append({"case": mode, "passed": passed, "audit_status": report["status"], "expected_code": code, "actual_codes": sorted(codes), "report": report["report"]})
        print(json.dumps(results[-1], ensure_ascii=False), flush=True)
    original = os.environ.get("PPTX_AUDIT_DISABLE_RENDER")
    os.environ["PPTX_AUDIT_DISABLE_RENDER"] = "1"
    try:
        path = root / "normal" / "fixture.pptx"
        report = audit_pptx(path, report_path=root / "no-renderer.json", project=path.parent)
        passed = not report["ok"] and report["status"] == "not_verified" and report["render"]["status"] == "not_verified"
        results.append({"case": "no_renderer", "passed": passed, "audit_status": report["status"], "report": report["report"]})
        report = audit_pptx(path, report_path=root / "required-no-renderer.json", project=path.parent, require_render=True)
        results.append({"case": "require_renderer", "passed": not report["ok"] and report["status"] == "failed", "audit_status": report["status"], "report": report["report"]})
    finally:
        if original is None: os.environ.pop("PPTX_AUDIT_DISABLE_RENDER", None)
        else: os.environ["PPTX_AUDIT_DISABLE_RENDER"] = original
    result = {"ok": all(r["passed"] for r in results), "cases": results, "output": str(root.resolve())}
    (root / "regression-results.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--output")
    args = parser.parse_args(); root = Path(args.output) if args.output else Path(tempfile.mkdtemp(prefix="pptx-audit-tests-"))
    report = run(root)
    print(json.dumps({"ok": report["ok"], "output": report["output"]}, ensure_ascii=False))
    raise SystemExit(0 if report["ok"] else 1)
