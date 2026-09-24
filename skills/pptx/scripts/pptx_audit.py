#!/usr/bin/env python3
"""Read-only PPTX checks plus real PowerPoint layout measurement and rendering.

No save, repair, text shrinking, inferred renderer success, or unchecked exception
waivers are permitted. Shape labels and source layout are emitted by pptx_export.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import math
import os
from pathlib import Path
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone
from typing import Any

DEFAULT_RULES = Path(__file__).resolve().parents[1] / "references" / "pptx-audit-rules.json"
EMU_PER_PT = 12700


def _hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _issue(issues: list, code: str, message: str, *, severity: str = "error", **data: Any) -> None:
    issues.append({"code": code, "severity": severity, "message": message, **data})


def _number(value: Any, default: float = 0) -> float:
    match = re.search(r"[-+]?\d+(?:\.\d+)?", str(value))
    return float(match.group()) if match else default


def _box(shape: Any) -> dict:
    return {"x": shape.left / EMU_PER_PT, "y": shape.top / EMU_PER_PT,
            "width": shape.width / EMU_PER_PT, "height": shape.height / EMU_PER_PT}


def _intersect(a: dict, b: dict) -> tuple[float, float, float, float]:
    x, y = max(a["x"], b["x"]), max(a["y"], b["y"])
    w = max(0, min(a["x"] + a["width"], b["x"] + b["width"]) - x)
    h = max(0, min(a["y"] + a["height"], b["y"] + b["height"]) - y)
    return x, y, w, h


def _area(box: dict) -> float:
    return max(0, box["width"]) * max(0, box["height"])


def _outside(box: dict, bounds: dict, tolerance: float) -> bool:
    return (box["x"] < bounds["x"] - tolerance or box["y"] < bounds["y"] - tolerance
            or box["x"] + box["width"] > bounds["x"] + bounds["width"] + tolerance
            or box["y"] + box["height"] > bounds["y"] + bounds["height"] + tolerance)


def _tag(name: str) -> dict:
    fields = name.split("|")
    if len(fields) >= 5 and fields[0] == "pptx-text":
        return dict(zip(("role", "source_id", "line_id", "segment_id"), fields[1:5]))
    return {}


def _font_color(font: Any) -> str | None:
    try:
        if font.color.type is None:
            return None
        if font.color.rgb is not None:
            return str(font.color.rgb).upper()
    except (AttributeError, TypeError):
        pass
    return None


def _spacing(paragraph: Any) -> tuple[str, float]:
    value = paragraph.line_spacing
    if value is None:
        return ("multiple", 1.0)
    return ("points", round(value.pt, 3)) if hasattr(value, "pt") else ("multiple", round(float(value), 3))


def _rgb(raw: Any) -> str | None:
    value = str(raw or "").strip().lstrip("#")
    if re.fullmatch(r"[0-9a-fA-F]{6}", value):
        return value.upper()
    numbers = re.findall(r"\d+(?:\.\d+)?", value)
    return "".join(f"{min(255, max(0, round(float(n)))):02X}" for n in numbers[:3]) if len(numbers) >= 3 else None


def _source_style(item: dict) -> dict:
    style = item.get("style") or item
    family = str(style.get("fontFamily", style.get("font_family", ""))).split(",")[0].strip(" '\"")
    size = style.get("fontSize", style.get("font_size_px"))
    size = _number(size) * .5 if size is not None else _number(style.get("font_size_pt"))
    weight = str(style.get("fontWeight", style.get("font_weight", "400")))
    line = style.get("lineHeight", style.get("line_height_px"))
    return {"font_family": family or None, "font_size_pt": size or None,
            "bold": weight.lower() == "bold" or _number(weight) >= 600,
            "italic": style.get("fontStyle") in ("italic", "oblique"),
            "color": _rgb(style.get("color")),
            "line_spacing": ("points", _number(line) * .5) if _number(line) else ("multiple", 1.0),
            "letter_spacing_pt": _number(style.get("letterSpacing", style.get("letter_spacing_px"))) * .5}


def _same(field: str, a: Any, b: Any, rules: dict) -> bool:
    if field in ("font_size_pt", "letter_spacing_pt") and a is not None and b is not None:
        return abs(a - b) <= rules["size_tolerance_pt"]
    if field in ("line_spacing", "source_line_spacing") and a is not None and b is not None:
        return a[0] == b[0] and abs(a[1] - b[1]) <= rules["size_tolerance_pt"]
    if field == "font_family" and isinstance(a, str) and isinstance(b, str):
        return a.casefold() == b.casefold()
    return a == b


def _installed_fonts() -> set[str] | None:
    if os.name != "nt":
        return None
    # Enumerating GDI families includes both English and localized CJK aliases.
    import ctypes
    from ctypes import wintypes
    class LOGFONTW(ctypes.Structure):
        _fields_ = [(name, wintypes.LONG) for name in ("lfHeight", "lfWidth", "lfEscapement", "lfOrientation", "lfWeight")] + [
            (name, wintypes.BYTE) for name in ("lfItalic", "lfUnderline", "lfStrikeOut", "lfCharSet", "lfOutPrecision", "lfClipPrecision", "lfQuality", "lfPitchAndFamily")
        ] + [("lfFaceName", wintypes.WCHAR * 32)]
    fonts: set[str] = set()
    callback_type = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.POINTER(LOGFONTW), ctypes.c_void_p, wintypes.DWORD, wintypes.LPARAM)
    def collect(font: Any, metric: Any, font_type: Any, param: Any) -> int:
        fonts.add(font.contents.lfFaceName.lstrip("@").casefold())
        return 1
    callback = callback_type(collect)
    gdi, user = ctypes.windll.gdi32, ctypes.windll.user32
    user.GetDC.argtypes = [wintypes.HWND]; user.GetDC.restype = wintypes.HDC
    user.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    gdi.EnumFontFamiliesExW.argtypes = [wintypes.HDC, ctypes.POINTER(LOGFONTW), callback_type, wintypes.LPARAM, wintypes.DWORD]
    device = user.GetDC(None)
    try:
        query = LOGFONTW(); query.lfCharSet = 1
        gdi.EnumFontFamiliesExW(device, ctypes.byref(query), callback, 0, 0)
    finally:
        user.ReleaseDC(None, device)
    import winreg
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
                for index in range(winreg.QueryInfoKey(key)[1]):
                    name, _, _ = winreg.EnumValue(key, index)
                    name = re.sub(r"\s*\([^)]*\)\s*$", "", name)
                    for family in name.split(" & "):
                        fonts.add(family.strip().lstrip("@").casefold())
        except OSError:
            continue
    return fonts


def _load_source(path: Path, project: Path | None, issues: list, source_report: Path | None = None) -> dict | None:
    candidates = [source_report] if source_report else ([project / "pptx-editability.json"] if project else []) + [path.parent / "pptx-editability.json", path.with_suffix(".editability.json")]
    for candidate in dict.fromkeys(candidates):
        if not candidate.is_file():
            continue
        try:
            source = json.loads(candidate.read_text(encoding="utf-8-sig"))
        except (ValueError, OSError) as error:
            _issue(issues, "source_report_invalid", str(error), path=str(candidate))
            return None
        if source.get("pptx_sha256") != _hash(path):
            _issue(issues, "source_hash_mismatch", "源布局报告没有绑定当前 PPTX 的 SHA256。", path=str(candidate))
        source["_path"] = str(candidate)
        return source
    _issue(issues, "source_not_verified", "缺少 pptx-editability.json，不能核对原始文字与原生文本覆盖。", severity="error" if project or source_report else "unverified")
    return None


def _structural(path: Path, source: dict | None, rules: dict, issues: list) -> dict:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.enum.text import MSO_AUTO_SIZE
    prs = Presentation(str(path))
    width, height = prs.slide_width / EMU_PER_PT, prs.slide_height / EMU_PER_PT
    if not height or abs(width / height - rules["aspect_ratio"]) > rules["aspect_ratio_tolerance"]:
        _issue(issues, "slide_aspect_ratio", "页面比例必须为 16:9。", width_pt=width, height_pt=height)
    if not len(prs.slides):
        _issue(issues, "empty_presentation", "PPTX 没有页面。")
    source_slides = source.get("slides", []) if source else []
    if source and len(source_slides) != len(prs.slides):
        _issue(issues, "slide_count_mismatch", "源布局与 PPTX 页数不一致。", expected=len(source_slides), actual=len(prs.slides))
    installed = _installed_fonts()
    role_samples: dict[str, list] = defaultdict(list)
    pages, native_count, all_fonts = [], 0, set()
    for page, slide in enumerate(prs.slides, 1):
        layouts = source_slides[page-1].get("text_layout", []) if page <= len(source_slides) else []
        expected = {item.get("shape_name"): item for item in layouts}
        if len(expected) != len(layouts):
            _issue(issues, "source_shape_name_duplicate", "源布局中出现重复文字对象名。", page=page)
        source_groups: dict[str, list] = defaultdict(list)
        for item in layouts:
            source_groups[str(item.get("source_id"))].append(item)
        for source_id, fragments in source_groups.items():
            if len({item.get("line_id") for item in fragments}) < 2:
                continue
            for item in fragments:
                style = _source_style(item); mode, line = style["line_spacing"]
                ratio = line / style["font_size_pt"] if mode == "points" and style["font_size_pt"] else line
                if ratio < rules["minimum_multiline_spacing_ratio"]:
                    _issue(issues, "dense_source_multiline_spacing", "拆成多个原生框的源多行文字行距过密。", page=page, source_id=source_id, spacing_ratio=round(ratio, 3))
                    break
        if source and not layouts:
            _issue(issues, "source_layout_missing", "该页缺少逐段原生文字布局证据。", page=page)
        records, seen = [], set()
        for z, shape in enumerate(slide.shapes, 1):
            tag = _tag(shape.name); box = _box(shape)
            record = {"name": shape.name, "id": shape.shape_id, "z": z, "box": box, "tag": tag, "has_text": False, "kind": str(shape.shape_type)}
            if abs(shape.rotation) > .01 or shape.shape_type == MSO_SHAPE_TYPE.GROUP or shape.has_table:
                _issue(issues, "geometry_not_verified", "旋转、分组或表格对象的完整排版尚未自动验证。", severity="unverified", page=page, shape=shape.name)
            if shape.has_text_frame and shape.text.strip():
                native_count += 1; record["has_text"] = True; record["text"] = shape.text
                item = expected.get(shape.name)
                if shape.name in seen:
                    _issue(issues, "text_shape_name_duplicate", "PPTX 中出现重复文字对象名。", page=page, shape=shape.name)
                seen.add(shape.name)
                role = tag.get("role") or (item or {}).get("role")
                if not role:
                    _issue(issues, "text_role_missing", "原生文字缺少语义角色，无法验证同级样式。", page=page, shape=shape.name)
                elif role not in rules["allowed_roles"]:
                    _issue(issues, "unknown_text_role", "文字角色不在有限语义角色表中，不能为单页创建角色规避一致性检查。", page=page, shape=shape.name, role=role)
                if shape.text_frame.auto_size == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE:
                    _issue(issues, "automatic_shrink", "文字设置了自动缩小，不能保证同级字号一致。", page=page, shape=shape.name)
                if source and not item:
                    _issue(issues, "unexpected_text", "PPTX 文字对象在源布局中不存在。", page=page, shape=shape.name, text=shape.text)
                if item and shape.text.replace("\v", "\n") != str(item.get("text", "")).replace("\v", "\n"):
                    _issue(issues, "source_text_mismatch", "原生文字与源布局不一致。", page=page, shape=shape.name, expected=item.get("text"), actual=shape.text)
                variant = ((item or {}).get("style") or {}).get("variant") or ("emphasis" if (item or {}).get("emphasis") else "")
                if variant == "regular":
                    variant = ""
                if variant and variant not in rules["allowed_variants"]:
                    _issue(issues, "unknown_style_variant", "样式变体不在审计规则的明确白名单内。", page=page, shape=shape.name, variant=variant)
                styles = []
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if not run.text.strip():
                            continue
                        font = run.font
                        style = {"font_family": font.name or paragraph.font.name,
                                 "font_size_pt": (font.size or paragraph.font.size).pt if (font.size or paragraph.font.size) else None,
                                 "bold": bool(font.bold if font.bold is not None else paragraph.font.bold),
                                 "italic": bool(font.italic if font.italic is not None else paragraph.font.italic),
                                 "color": _font_color(font) or _font_color(paragraph.font), "line_spacing": _spacing(paragraph),
                                 "letter_spacing_pt": _number(run._r.get_or_add_rPr().get("spc")) / 100}
                        styles.append(style)
                        for required in ("font_family", "font_size_pt", "color"):
                            if style[required] is None:
                                _issue(issues, "implicit_text_style", "字体、字号和颜色必须显式声明，以避免继承或替换漂移。", page=page, shape=shape.name, field=required)
                        if style["font_size_pt"] is not None and style["font_size_pt"] < rules["minimum_font_size_pt"]:
                            _issue(issues, "font_too_small", "字号低于最小可读门槛。", page=page, shape=shape.name, size_pt=style["font_size_pt"])
                        if style["font_family"]:
                            all_fonts.add(style["font_family"])
                        if item:
                            desired = _source_style(item)
                            for field in rules["style_fields"]:
                                # Exported visual-line fragments have one PPT paragraph; CSS line-height is
                                # represented by their positions and audited independently across the source.
                                if field == "line_spacing" and item.get("segment_id") and "\n" not in str(item.get("text", "")) and "\v" not in str(item.get("text", "")):
                                    if style["line_spacing"] != ("multiple", 1.0) and not _same(field, style[field], desired[field], rules):
                                        _issue(issues, "fragment_line_spacing", "单视觉行原生文本行距必须对应源行距或统一单倍行距。", page=page, shape=shape.name, expected=desired[field], actual=style[field])
                                    continue
                                if desired[field] is not None and not _same(field, style[field], desired[field], rules):
                                    _issue(issues, "source_style_mismatch", "PPTX 文字样式与源布局不一致。", page=page, shape=shape.name, field=field, expected=desired[field], actual=style[field])
                            style["source_line_spacing"] = desired["line_spacing"]
                        if role:
                            role_samples[role].append({"page": page, "shape": shape.name, "style": style, "variant": variant})
                record["styles"] = styles
                # A minimum multiline setting is a structural guard; real line bounds are checked separately.
                if "\n" in shape.text or "\v" in shape.text:
                    for style in styles:
                        mode, value = style["line_spacing"]
                        size = style.get("font_size_pt") or 0
                        ratio = value / size if mode == "points" and size else value
                        if ratio < rules["minimum_multiline_spacing_ratio"]:
                            _issue(issues, "dense_multiline_spacing", "多行文字行距过密，必须调整布局而不是缩小字号。", page=page, shape=shape.name, spacing_ratio=round(ratio, 3))
                            break
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                record["picture"] = {"blob": shape.image.blob, "crop": [shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom]}
            records.append(record)
        for name, item in expected.items():
            if name not in seen:
                _issue(issues, "missing_native_text", "源文字没有以原生文本出现在 PPTX。", page=page, shape=name, text=item.get("text"))
        if page <= len(source_slides) and any(item.get("kind") == "rendered-text" for item in source_slides[page-1].get("unsupported_structures", [])):
            _issue(issues, "rasterized_text", "存在被栅格化的源文字，文字可编辑性未满足。", page=page)
        pages.append({"page": page, "shapes": records})
    for family in sorted(all_fonts):
        if installed is not None and family.lstrip("@").casefold() not in installed:
            _issue(issues, "font_unavailable", "本机缺少指定字体，PowerPoint 可能发生替换。", font=family)
    for role, samples in role_samples.items():
        base = next((sample for sample in samples if not sample["variant"]), samples[0])
        for sample in samples:
            variant = rules["allowed_variants"].get(sample["variant"], {})
            may_differ = set(variant.get("may_differ", []))
            for field, value in variant.get("require", {}).items():
                if not _same(field, sample["style"].get(field), value, rules):
                    _issue(issues, "style_variant_invalid", "样式变体不满足白名单定义。", role=role, page=sample["page"], shape=sample["shape"], field=field)
            for field in rules["style_fields"] + ["source_line_spacing"]:
                if field not in may_differ and not _same(field, sample["style"].get(field), base["style"].get(field), rules):
                    _issue(issues, "role_style_drift", "相同文本层级的样式发生漂移。", role=role, page=sample["page"], shape=sample["shape"], field=field, expected=base["style"].get(field), actual=sample["style"].get(field), reference_page=base["page"])
    for levels in rules.get("hierarchy", []):
        present = [(role, [s["style"]["font_size_pt"] for s in role_samples[role] if s["style"]["font_size_pt"]]) for role in levels if role in role_samples]
        for (higher, a), (lower, b) in zip(present, present[1:]):
            if a and b and min(a) < max(b) - rules["size_tolerance_pt"]:
                _issue(issues, "hierarchy_size_inversion", "高层级文字字号小于低层级文字。", higher=higher, lower=lower)
    if native_count == 0:
        _issue(issues, "no_native_text", "PPTX 没有可编辑的原生文字。")
    return {"slide_count": len(prs.slides), "width_pt": width, "height_pt": height, "native_text_count": native_count, "fonts": sorted(all_fonts), "font_inventory_available": installed is not None, "pages": pages}


def _range_box(value: Any) -> dict:
    return {"x": float(value.BoundLeft), "y": float(value.BoundTop), "width": float(value.BoundWidth), "height": float(value.BoundHeight)}


def _picture_opacity(record: dict, area: tuple) -> float:
    from PIL import Image
    box, picture = record["box"], record["picture"]
    if box["width"] <= 0 or box["height"] <= 0:
        return 0
    with Image.open(io.BytesIO(picture["blob"])) as image:
        alpha = image.convert("RGBA").getchannel("A")
        left, top, right, bottom = picture["crop"]
        x, y, w, h = area
        crop = ((left + (x-box["x"])/box["width"]*(1-left-right))*image.width,
                (top + (y-box["y"])/box["height"]*(1-top-bottom))*image.height,
                (left + (x+w-box["x"])/box["width"]*(1-left-right))*image.width,
                (top + (y+h-box["y"])/box["height"]*(1-top-bottom))*image.height)
        region = alpha.crop(tuple(round(v) for v in crop))
        histogram = region.histogram()
        return sum(histogram[230:]) / max(1, sum(histogram))


def _render_powerpoint(path: Path, directory: Path, structural: dict, rules: dict, issues: list) -> dict:
    if os.environ.get("PPTX_AUDIT_DISABLE_RENDER") == "1":
        return {"status": "not_verified", "engine": None, "reason": "PPTX_AUDIT_DISABLE_RENDER=1", "slides": []}
    if os.name != "nt":
        return {"status": "not_verified", "engine": None, "reason": "PowerPoint COM requires Windows and Microsoft PowerPoint.", "slides": []}
    try:
        import win32com.client
        import pythoncom
    except ImportError as error:
        return {"status": "not_verified", "engine": None, "reason": str(error), "slides": []}
    app = presentation = None
    rendered = []; initialized = False
    try:
        pythoncom.CoInitialize(); initialized = True
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
        except Exception as error:
            return {"status": "not_verified", "engine": None, "reason": str(error), "slides": []}
        directory.mkdir(parents=True, exist_ok=True)
        app.DisplayAlerts = 1
        presentation = app.Presentations.Open(str(path.resolve()), True, False, False)
        bounds = {"x": 0, "y": 0, "width": structural["width_pt"], "height": structural["height_pt"]}
        for page in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(page)
            records = {s["id"]: s for s in structural["pages"][page-1]["shapes"]}
            lines, objects = [], []
            for shape in slide.Shapes:
                record = records.get(int(shape.Id))
                if record is None:
                    _issue(issues, "render_shape_mismatch", "PowerPoint 的对象与文件结构不一致。", page=page, shape=str(shape.Name))
                    continue
                obj = {"record": record, "opacity": 0.0, "geometry": "rectangle"}
                if "picture" in record:
                    obj["opacity"] = 1.0
                else:
                    try:
                        if int(shape.Fill.Visible) == -1:
                            if int(shape.Fill.Type) == 1:
                                obj["opacity"] = 1 - float(shape.Fill.Transparency)
                            elif int(shape.Fill.Type) != 0:
                                _issue(issues, "foreground_fill_not_verified", "非纯色前景填充透明度未完整验证。", severity="unverified", page=page, shape=record["name"])
                        if int(shape.Type) == 1 and int(shape.AutoShapeType) != 1:
                            obj["geometry"] = "ellipse" if int(shape.AutoShapeType) == 9 else "complex"
                    except Exception:
                        pass  # A missing fill is transparent, e.g. a connector.
                objects.append(obj)
                if not record["has_text"]:
                    continue
                try:
                    text_range = shape.TextFrame.TextRange
                    count = int(text_range.Lines().Count)
                    for line_no in range(1, count + 1):
                        line = text_range.Lines(line_no, 1)
                        if not str(line.Text).strip():
                            continue
                        visible_text = str(line.Text).rstrip("\r\v")
                        # A whole TextRange includes an invisible paragraph terminator which PowerPoint
                        # gives width. Measure the visible characters, never loosen the overflow threshold.
                        visible_range = line.Characters(1, len(visible_text))
                        line_box = _range_box(visible_range)
                        if _outside(line_box, bounds, rules["bounds_tolerance_pt"]):
                            _issue(issues, "text_outside_slide", "实际文字行越出页面。", page=page, shape=record["name"], line=line_no, bounds=line_box)
                        if _outside(line_box, record["box"], rules["bounds_tolerance_pt"]):
                            _issue(issues, "text_frame_overflow", "实际文字行越出其文本框。", page=page, shape=record["name"], line=line_no, bounds=line_box, frame=record["box"])
                        if not all(math.isfinite(v) for v in line_box.values()):
                            raise ValueError("PowerPoint returned non-finite text bounds")
                        actual_font = str(line.Font.Name or "")
                        expected_fonts = {s["font_family"].casefold() for s in record.get("styles", []) if s["font_family"]}
                        if actual_font and expected_fonts and actual_font.casefold() not in expected_fonts:
                            _issue(issues, "render_font_substitution", "PowerPoint 返回的字体与声明字体不一致。", page=page, shape=record["name"], actual_font=actual_font, expected_fonts=sorted(expected_fonts))
                        lines.append({"shape": record["name"], "shape_id": record["id"], "line": line_no, "box": line_box, "z": record["z"], "text": visible_text, "font": actual_font, "font_size_pt": float(line.Font.Size)})
                except Exception as error:
                    _issue(issues, "text_measurement_failed", "无法测量实际文本行，排版未完成验证。", severity="unverified", page=page, shape=record["name"], detail=str(error))
            _check_measured_geometry(page, lines, objects, rules, issues)
            png = directory / f"slide-{page:03d}.png"
            slide.Export(str(png.resolve()), "PNG", rules["render_width"], rules["render_height"])
            if not png.is_file() or png.stat().st_size < 100:
                raise RuntimeError(f"PowerPoint did not create page {page} PNG")
            from PIL import Image
            with Image.open(png) as image:
                image.verify()
            rendered.append({"page": page, "image": str(png.resolve()), "sha256": _hash(png), "text_lines": lines})
        return {"status": "passed", "engine": "Microsoft PowerPoint COM", "read_only": True, "directory": str(directory.resolve()), "slides": rendered}
    except Exception as error:
        _issue(issues, "render_failed", "PowerPoint 实际渲染失败。", detail=str(error))
        return {"status": "failed", "engine": "Microsoft PowerPoint COM", "reason": str(error), "slides": rendered}
    finally:
        if presentation is not None:
            try: presentation.Close()
            except Exception: pass
        if app is not None:
            try: app.Quit()
            except Exception: pass
        if initialized:
            pythoncom.CoUninitialize()


def _check_measured_geometry(page: int, lines: list, objects: list, rules: dict, issues: list) -> None:
    tolerance = rules["intersection_tolerance_pt"]
    for index, line in enumerate(lines):
        for other in lines[index+1:]:
            x, y, width, height = _intersect(line["box"], other["box"])
            if width <= tolerance or height <= tolerance:
                continue
            same_shape = line["shape_id"] == other["shape_id"]
            if same_shape and height <= max(tolerance, max(line["font_size_pt"], other["font_size_pt"]) * rules["line_overlap_tolerance_fraction"]):
                continue
            fraction = width * height / max(1, min(_area(line["box"]), _area(other["box"])))
            if fraction >= rules["overlap_fraction"]:
                _issue(issues, "multiline_overlap" if same_shape else "text_overlap", "PowerPoint 实际文本行发生重叠。", page=page, shape=line["shape"], other_shape=other["shape"], line=line["line"], other_line=other["line"], intersection={"x": x, "y": y, "width": width, "height": height}, overlap_fraction=round(fraction, 4))
        for obj in objects:
            record = obj["record"]
            if record["z"] <= line["z"] or record["id"] == line["shape_id"] or obj["opacity"] < rules["opaque_alpha_threshold"]:
                continue
            intersection = _intersect(line["box"], record["box"])
            x, y, width, height = intersection
            fraction = width * height / max(1, _area(line["box"]))
            if width <= tolerance or height <= tolerance or fraction < rules["occlusion_fraction"]:
                continue
            if "picture" in record:
                fraction *= _picture_opacity(record, intersection)
            elif obj["geometry"] == "ellipse":
                # Sample the intersection rather than assuming an ellipse fills its bounding rectangle.
                b = record["box"]; inside = 0
                for i in range(12):
                    for j in range(12):
                        px, py = x + width * (i+.5)/12, y + height * (j+.5)/12
                        inside += ((px-b["x"]-b["width"]/2)/(b["width"]/2))**2 + ((py-b["y"]-b["height"]/2)/(b["height"]/2))**2 <= 1
                fraction *= inside / 144
            elif obj["geometry"] == "complex":
                _issue(issues, "foreground_geometry_not_verified", "前景复杂形状与文字相交，需检查实际渲染。", severity="unverified", page=page, shape=line["shape"], foreground=record["name"])
                continue
            if fraction >= rules["occlusion_fraction"]:
                _issue(issues, "foreground_occlusion", "位于文字前方的不透明对象遮挡实际文本行。", page=page, shape=line["shape"], foreground=record["name"], line=line["line"], occlusion_fraction=round(fraction, 4))


def audit_pptx(pptx_path: str | Path, report_path: str | Path | None = None,
               render_dir: str | Path | None = None, rules_path: str | Path | None = None,
               require_render: bool = False, project: str | Path | None = None,
               source_report: str | Path | None = None) -> dict:
    """Audit without modifying the deck; ok requires actual measured render success."""
    path = Path(pptx_path).resolve()
    destination = Path(report_path).resolve() if report_path else path.with_suffix(".audit.json")
    if destination == path:
        raise ValueError("Audit report must not overwrite the input PPTX")
    rules_file = Path(rules_path).resolve() if rules_path else DEFAULT_RULES
    rules = json.loads(rules_file.read_text(encoding="utf-8-sig"))
    issues: list = []
    result = {"schema_version": "pptx.audit/v1", "created_at": datetime.now(timezone.utc).isoformat(), "pptx": str(path), "ok": False, "status": "not_verified", "issues": issues,
              "rules": {"path": str(rules_file), "sha256": _hash(rules_file)}, "require_render": require_render,
              "render": {"status": "not_verified", "engine": None, "slides": []}, "limitations": rules.get("limitations", [])}
    before = _hash(path) if path.is_file() else None
    try:
        if before is None:
            raise FileNotFoundError(path)
        source = _load_source(path, Path(project).resolve() if project else None, issues, Path(source_report).resolve() if source_report else None)
        structural = _structural(path, source, rules, issues)
        result["summary"] = {k: v for k, v in structural.items() if k != "pages"}
        result["source_layout"] = source.get("_path") if source else None
        result["source_layout_sha256"] = _hash(Path(source["_path"])) if source else None
        directory = Path(render_dir).resolve() if render_dir else path.parent / f"{path.stem}-audit-render"
        result["render"] = _render_powerpoint(path, directory, structural, rules, issues)
        if result["render"]["status"] == "not_verified":
            _issue(issues, "renderer_unavailable", "缺少 PowerPoint 实际渲染，排版不能判为通过。", severity="error" if require_render else "unverified", reason=result["render"].get("reason"))
        if result["render"]["status"] == "passed" and len(result["render"]["slides"]) != structural["slide_count"]:
            _issue(issues, "render_page_count", "实际渲染未覆盖全部页面。")
    except Exception as error:
        _issue(issues, "audit_failed", "审计无法完成。", detail=str(error))
    after = _hash(path) if path.is_file() else None
    result["input_sha256"] = before
    result["input_unchanged"] = before is not None and before == after
    if before is not None and before != after:
        _issue(issues, "input_modified", "审计过程中原始 PPTX 发生变化。")
    errors = sum(i["severity"] == "error" for i in issues)
    unverified = any(i["severity"] == "unverified" for i in issues) or result["render"]["status"] != "passed"
    result["status"] = "failed" if errors else "not_verified" if unverified else "passed"
    result["ok"] = result["status"] == "passed"
    result["issue_counts"] = {level: sum(i["severity"] == level for i in issues) for level in ("error", "warning", "unverified")}
    result["report"] = str(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--report")
    parser.add_argument("--render-dir")
    parser.add_argument("--rules")
    parser.add_argument("--require-render", action="store_true")
    parser.add_argument("--project")
    parser.add_argument("--source-report")
    args = parser.parse_args()
    result = audit_pptx(args.pptx, args.report, args.render_dir, args.rules, args.require_render, args.project, args.source_report)
    print(json.dumps({k: result[k] for k in ("ok", "status", "report", "issue_counts", "input_unchanged")}, ensure_ascii=False))
    return 0 if result["ok"] else 2 if result["status"] == "not_verified" else 1


if __name__ == "__main__":
    sys.exit(main())
