#!/usr/bin/env python3
"""Hybrid PPTX export from canonical final HTML and its rendered DOM."""
from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from cdp_validate import WebSocket, _stop_browser, _wait_for_devtools_port
from theme import PALETTES, TYPOGRAPHY, validate_theme

SLIDE_WIDTH_PX = 1920
SLIDE_HEIGHT_PX = 1080
SLIDE_WIDTH_EMU = 12_192_000
SLIDE_HEIGHT_EMU = 6_858_000
COVERAGE_SCHEMA = "pptx.pptx-editability/v3"
POINTS_PER_PX = .5  # 1920 CSS px -> 960 PowerPoint pt, including font size.


def require_python_pptx() -> Any:
    """Import python-pptx only for export."""
    try:
        import pptx
    except (ImportError, OSError) as error:
        raise SystemExit("PPTX export requires python-pptx: python3 -m pip install python-pptx") from error
    return pptx


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _cdp_command(socket: WebSocket, command_id: int, method: str, params: dict | None = None) -> dict:
    socket.send_json({"id": command_id, "method": method, **({"params": params} if params else {})})
    while True:
        message = socket.recv_json()
        if message.get("id") != command_id:
            continue
        if "error" in message:
            raise RuntimeError(f"Chrome DevTools {method} failed: {message['error']}")
        return message.get("result") or {}


def _browser_session(chrome: str, html_path: Path) -> tuple[subprocess.Popen, tempfile.TemporaryDirectory, WebSocket]:
    profile_ctx = tempfile.TemporaryDirectory(prefix="pptx-pptx-cdp-")
    profile = Path(profile_ctx.name)
    command = [
        chrome, "--headless", "--no-sandbox", "--allow-file-access-from-files",
        "--disable-background-networking", "--disable-component-update", "--disable-default-apps",
        "--disable-features=PaintHolding,RenderDocument", "--disable-sync", "--force-color-profile=srgb",
        "--force-device-scale-factor=1", "--hide-scrollbars", "--metrics-recording-only", "--no-first-run",
        "--window-size=1920,1080", f"--user-data-dir={profile}", "--remote-debugging-port=0", html_path.resolve().as_uri(),
    ]
    process = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
    try:
        deadline, port_file = time.monotonic() + 15, profile / "DevToolsActivePort"
        port = _wait_for_devtools_port(port_file, deadline, "Chrome did not expose a DevTools port for PPTX export.")
        pages: list[dict] = []
        while time.monotonic() < deadline and not pages:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=2) as response:
                    pages = [item for item in json.load(response) if item.get("type") == "page"]
            except OSError:
                time.sleep(.05)
        if not pages:
            raise RuntimeError("Chrome opened no inspectable page for PPTX export.")
        target = html_path.resolve().as_uri()
        page = next((item for item in pages if item.get("url", "").startswith(target)), pages[0])
        socket = WebSocket(page["webSocketDebuggerUrl"])
        socket.sock.settimeout(20)
        return process, profile_ctx, socket
    except Exception:
        _stop_browser(process, profile)
        profile_ctx.cleanup()
        raise


def _layout_expression(slide_index: int) -> str:
    """Measure every rendered text node, retaining each visual line and inline run."""
    return r"""(async () => {
      if(document.readyState!=='complete') await new Promise(done=>addEventListener('load',done,{once:true}));
      await document.fonts?.ready;
      await Promise.all([...document.images].map(im => im.complete ? Promise.resolve() : new Promise(done => {im.onload=im.onerror=done;})));
      document.querySelector('[data-pptx-export-hide]')?.remove();
      document.documentElement.style.cssText += ';width:1920px!important;height:1080px!important';
      document.body.style.cssText += ';margin:0!important;width:1920px!important;height:1080px!important;overflow:hidden!important';
      const viewport=document.querySelector('.deck-viewport,.slide-preview-viewport');
      const shell=document.querySelector('.deck-stage-shell,.slide-preview-shell');
      let stage=document.querySelector('.deck-stage,.slide-preview-stage');
      if (!stage) throw new Error('canonical HTML has no deck stage');
      // Runtime resize listeners write stage.style.transform asynchronously. An
      // important stylesheet rule remains authoritative after those writes.
      let geometry=document.querySelector('[data-pptx-export-geometry]');
      if(!geometry){geometry=document.createElement('style');geometry.dataset.pptxExportGeometry='';document.head.append(geometry);}
      geometry.textContent='.deck-stage,.slide-preview-stage{position:absolute!important;left:0!important;top:0!important;width:1920px!important;height:1080px!important;transform:none!important;translate:none!important;scale:none!important;rotate:none!important;transform-origin:0 0!important}.deck-stage-shell,.slide-preview-shell{position:relative!important;width:1920px!important;height:1080px!important;transform:none!important;translate:none!important}.deck-viewport,.slide-preview-viewport{position:fixed!important;inset:0!important;width:1920px!important;height:1080px!important;transform:none!important;translate:none!important}';
      const overview=document.querySelector('[data-deck-overview]');
      if (overview && !overview.hidden) document.querySelector('[data-deck-overview-close]')?.click();
      if (overview && !overview.hidden) throw new Error('overview must be closed before PPTX capture');
      // Keep original slide nodes but detach the stage object held by runtime
      // listeners. CSSOM assignments can retain an inline !important priority;
      // a new stage prevents pending resize callbacks from changing the snapshot.
      if(!stage.dataset.pptxExportStage){
        const isolated=stage.cloneNode(false);
        isolated.dataset.pptxExportStage='true';
        while(stage.firstChild) isolated.append(stage.firstChild);
        stage.replaceWith(isolated);stage=isolated;
      }
      const chrome=[...document.querySelectorAll('.deck-counter,.progress-bar,.next-preview,.deck-overview-toggle,.deck-overview')];
      chrome.forEach(n => n.style.display='none');
      if (viewport) viewport.style.cssText += ';position:fixed!important;inset:0!important;width:1920px!important;height:1080px!important';
      if (shell) shell.style.cssText += ';position:relative!important;width:1920px!important;height:1080px!important';
      stage.style.cssText += ';position:absolute!important;left:0!important;top:0!important;width:1920px!important;height:1080px!important;transform:none!important';
      const slides=[...document.querySelectorAll('.slide')], slide=slides[__INDEX__];
      if (!slide) throw new Error('missing slide index __INDEX__');
      if (stage.querySelectorAll(':scope > .slide').length !== slides.length) throw new Error('all slides must be restored to stage before PPTX capture');
      slides.forEach((n,i) => {
        n.classList.toggle('active',i===__INDEX__);
        n.style.cssText += i===__INDEX__ ? ';display:block!important;visibility:visible!important;opacity:1!important;transition:none!important' : ';display:none!important;visibility:hidden!important;opacity:0!important;transition:none!important';
      });
      await new Promise(done => requestAnimationFrame(() => requestAnimationFrame(done)));
      const root=slide.getBoundingClientRect();
      if(Math.abs(root.left)>.1||Math.abs(root.top)>.1||Math.abs(root.width-1920)>.1||Math.abs(root.height-1080)>.1) throw new Error('Export slide must occupy exactly (0,0,1920,1080); received '+JSON.stringify({x:root.left,y:root.top,width:root.width,height:root.height}));
      const relative=r => ({x:r.left-root.left,y:r.top-root.top,width:r.width,height:r.height});
      const visible=n => {
        if (!n?.getClientRects().length) return false;
        for(let p=n;p;p=p.parentElement){
          const s=getComputedStyle(p);
          if(s.display==='none'||s.visibility==='hidden'||s.visibility==='collapse'||Number(s.opacity)===0) return false;
        }
        return true;
      };
      const all=[slide,...slide.querySelectorAll('*')];
      const sources=new Map(), text=[], unsupported=[];
      const sourceFor=n => {
        let b=n;
        while(b!==slide && ['inline','contents'].includes(getComputedStyle(b).display)) b=b.parentElement;
        if(!sources.has(b)) sources.set(b,{id:'s'+String(__INDEX__+1).padStart(3,'0')+'-e'+String(sources.size+1).padStart(4,'0'),lines:[]});
        return {node:b,...sources.get(b)};
      };
      const roleFor=(n,b) => {
        const explicit=n.closest('[data-text-role]');
        if(explicit) return explicit.getAttribute('data-text-role').replace(/[^a-zA-Z0-9_-]/g,'-');
        if(n.closest('code,pre')) return 'code';
        if(n.closest('figcaption,small') || /caption|footnote|source|footer/.test(b.className)) return 'caption';
        if(/label|eyebrow|kicker|badge|chip|tag/.test(b.className)) return 'label';
        if(/metric|stat-number/.test(b.className)) return 'metric';
        if(b.matches('h1')) return /cover|title|section/.test(slide.dataset.kind||slide.className) ? 'cover-title' : 'slide-title';
        if(b.matches('h2')) return 'slide-title';
        if(b.matches('h3,h4,h5,h6')) return 'subtitle';
        return 'body';
      };
      const reasonsFor=(n,s,rect) => {
        const reasons=[];
        if(s.writingMode!=='horizontal-tb'||s.direction==='rtl') reasons.push('vertical or bidirectional text requires a native layout adapter');
        if(!['none','uppercase','lowercase','capitalize'].includes(s.textTransform)) reasons.push('unsupported text transform');
        if(s.textShadow!=='none'||s.textDecorationLine!=='none') reasons.push('text shadow/decoration must be represented explicitly in source');
        if(s.webkitTextFillColor==='rgba(0, 0, 0, 0)') reasons.push('transparent/gradient text cannot be restored as ordinary native text');
        if(n.namespaceURI!=='http://www.w3.org/1999/xhtml') reasons.push('SVG text requires conversion to ordinary HTML text');
        for(let p=n;p&&p!==stage;p=p.parentElement){
          const a=getComputedStyle(p);
          if(a.transform!=='none'){
            const m=new DOMMatrixReadOnly(a.transform);
            if(!m.is2D||Math.abs(m.a-1)>.001||Math.abs(m.d-1)>.001||Math.abs(m.b)>.001||Math.abs(m.c)>.001) reasons.push('scaled/rotated text: fix source instead of changing individual font sizes');
          }
          if(a.filter!=='none'||a.mixBlendMode!=='normal'||Number(a.opacity)<.999) reasons.push('filtered/blended/translucent text requires an explicit native adapter');
          if(!['none',''].includes(a.clipPath)||!['none',''].includes(a.maskImage)) reasons.push('masked/clipped text requires an explicit native adapter');
          if(['hidden','clip','scroll','auto'].includes(a.overflowX)||['hidden','clip','scroll','auto'].includes(a.overflowY)){
            const clip=p.getBoundingClientRect();
            if(rect.left<clip.left-1||rect.right>clip.right+1||rect.top<clip.top-1||rect.bottom>clip.bottom+1) reasons.push('visible text extends beyond an ancestor clipping box');
          }
        }
        if(rect.left<root.left-1||rect.right>root.right+1||rect.top<root.top-1||rect.bottom>root.bottom+1) reasons.push('visible text extends beyond slide bounds');
        return [...new Set(reasons)];
      };
      // A TreeWalker includes parent text around inline spans; selecting leaf elements loses it.
      const walker=document.createTreeWalker(slide,NodeFilter.SHOW_TEXT);
      let tn, ordinal=0;
      while(tn=walker.nextNode()){
        const n=tn.parentElement;
        if(!visible(n)||n.closest('script,style,noscript,template')) continue;
        const s=getComputedStyle(n), source=sourceFor(n), role=roleFor(n,source.node);
        const fontSelector='f'+String(__INDEX__+1)+'-'+String(++ordinal);
        n.dataset.pptxFontSource=n.dataset.pptxFontSource||fontSelector;
        const declaredVariant=n.closest('[data-text-variant]')?.dataset.textVariant;
        const emphasis=!!n.closest('strong,b,[data-text-emphasis]')||declaredVariant==='emphasis';
        const style={fontFamily:s.fontFamily,fontSize:parseFloat(s.fontSize),fontWeight:s.fontWeight,fontStyle:s.fontStyle,color:s.color,lineHeight:s.lineHeight==='normal'?parseFloat(s.fontSize)*1.2:parseFloat(s.lineHeight),letterSpacing:parseFloat(s.letterSpacing)||0,variant:declaredVariant||(emphasis?'emphasis':'')};
        let segment=null, offset=0, previousCharacter='';
        for(const character of Array.from(tn.data)){
          const begin=offset; offset+=character.length;
          const range=document.createRange(); range.setStart(tn,begin); range.setEnd(tn,offset);
          const rect=range.getBoundingClientRect();
          if(rect.width<.01||rect.height<.01){previousCharacter=character;continue;}
          if(character==='\n' && !['pre','pre-wrap','break-spaces'].includes(s.whiteSpace)){previousCharacter=character;continue;}
          let value=character;
          if(/\s/u.test(character)&&!['pre','pre-wrap','break-spaces'].includes(s.whiteSpace)) value=' ';
          if(s.textTransform==='uppercase') value=value.toUpperCase();
          if(s.textTransform==='lowercase') value=value.toLowerCase();
          if(s.textTransform==='capitalize' && (begin===0||/\s/u.test(previousCharacter))) value=value.toUpperCase();
          previousCharacter=character;
          let line=source.lines.find(l => Math.abs(l.y-rect.top)<Math.max(2,Math.min(l.height,rect.height)*.30));
          if(!line){line={id:'l'+String(source.lines.length+1).padStart(3,'0'),y:rect.top,height:rect.height,segments:0};source.lines.push(line);}
          const reasons=reasonsFor(n,s,rect);
          if(!segment||segment.line_id!==line.id||rect.left<segment.box.x+root.left-.5){
            segment={index:text.length,selector:'[data-pptx-font-source="'+n.dataset.pptxFontSource+'"]',source_id:source.id,line_id:line.id,segment_id:'r'+String(++line.segments).padStart(3,'0'),role,emphasis,style:{...style},text:'',box:relative(rect),native:reasons.length===0,reason:reasons.join('; '),tag:n.tagName.toLowerCase()};
            segment.shape_name=['pptx-text',role,source.id,line.id,segment.segment_id].join('|');
            text.push(segment);
          }
          segment.text+=value;
          segment.box.width=Math.max(segment.box.width,rect.right-root.left-segment.box.x);
          segment.box.height=Math.max(segment.box.height,rect.bottom-root.top-segment.box.y);
          if(reasons.length){segment.native=false;segment.reason=[...new Set([segment.reason,...reasons].filter(Boolean))].join('; ');}
        }
      }
      // CSS generated words and canvas drawings cannot be silently counted as editable.
      for(const n of all.filter(visible)){
        for(const pseudo of ['::before','::after']){
          const ps=getComputedStyle(n,pseudo), content=ps.content;
          if(content&&!['none','normal','""',"''"].includes(content)&&ps.display!=='none'&&ps.visibility!=='hidden'&&Number(ps.opacity)!==0)
            unsupported.push({kind:'generated-text',selector:n.tagName.toLowerCase()+pseudo,text:content,reason:'Move generated text into an ordinary HTML text node.'});
        }
        if(n.matches('li')&&getComputedStyle(n).listStyleType!=='none') unsupported.push({kind:'generated-text',selector:'li::marker',text:'list marker',reason:'Use an explicit native text bullet and list-style:none.'});
        if(n.matches('canvas')) unsupported.push({kind:'canvas',selector:'canvas',reason:'Canvas may contain uninspectable text; provide an image with explicit raster content declaration or HTML labels.'});
        if(n.shadowRoot||n.matches('iframe,object,embed,input,textarea,select,video')) unsupported.push({kind:'uninspectable-content',selector:n.tagName.toLowerCase(),reason:'Embedded documents, shadow trees, form values and video may contain text outside the measured DOM. Replace with ordinary HTML labels or an explicitly declared source image.'});
      }
      const images=[...slide.querySelectorAll('img')].filter(visible).map((n,index)=>({index,selector:'img:nth-of-type('+(index+1)+')',src:n.currentSrc||n.src,box:relative(n.getBoundingClientRect()),native:false,reason:'Image pixels and CSS remain together in the background; text embedded in source images is not editable.'}));
      const cleanText=[];
      for(const item of text){
        if(item.text.trim()){cleanText.push(item);continue;}
        // A standalone space between inline elements is real source content.
        // Keep it on the preceding native run; do not invent an empty text box.
        const prior=cleanText[cleanText.length-1];
        if(prior&&prior.source_id===item.source_id&&prior.line_id===item.line_id){
          prior.text+=item.text;
          prior.box.width=Math.max(prior.box.width,item.box.x+item.box.width-prior.box.x);
        }
      }
      // Transparent text leaves backgrounds, borders, inline icons and pictures untouched.
      const hide=document.createElement('style');
      hide.dataset.pptxExportHide='';
      hide.textContent='.slide.active,.slide.active *{ -webkit-text-fill-color:transparent!important;text-shadow:none!important;text-decoration-color:transparent!important}';
      document.head.append(hide);
      await new Promise(done => requestAnimationFrame(() => requestAnimationFrame(done)));
      return {slideId:slide.dataset.slideId||'',title:slide.dataset.title||'',text:cleanText,images,unsupported,slideBox:{x:0,y:0,width:root.width,height:root.height},captureClip:{x:root.left+scrollX,y:root.top+scrollY,width:root.width,height:root.height,scale:1},chromeHidden:chrome.every(n=>getComputedStyle(n).display==='none')};
    })()""".replace("__INDEX__", str(slide_index))


def collect_render_layers(chrome: str, html_path: Path, deck: dict, background_dir: Path) -> list[dict]:
    """Capture one background plus final-DOM layers for each manifest slide."""
    paths = deck.get("slides") if isinstance(deck, dict) else None
    if not isinstance(paths, list) or not paths:
        raise RuntimeError("deck manifest has no slides for PPTX export")
    process, profile_ctx, socket = _browser_session(chrome, html_path)
    profile, results, command_id = Path(profile_ctx.name), [], 1
    try:
        _cdp_command(socket, command_id, "Page.enable"); command_id += 1
        _cdp_command(socket, command_id, "Runtime.enable"); command_id += 1
        _cdp_command(socket, command_id, "DOM.enable"); command_id += 1
        _cdp_command(socket, command_id, "CSS.enable"); command_id += 1
        document = _cdp_command(socket, command_id, "DOM.getDocument", {"depth": 0}); command_id += 1
        document_id = document["root"]["nodeId"]
        for index, _ in enumerate(paths):
            evaluated = _cdp_command(socket, command_id, "Runtime.evaluate", {"expression": _layout_expression(index), "awaitPromise": True, "returnByValue": True}); command_id += 1
            remote = evaluated.get("result") or {}
            if remote.get("subtype") == "error" or "exceptionDetails" in evaluated:
                details = evaluated.get("exceptionDetails") or {}
                message = remote.get("description") or (details.get("exception") or {}).get("description") or details.get("text") or json.dumps(evaluated, ensure_ascii=False)
                raise RuntimeError(f"Could not inspect rendered slide {index + 1}: {message}")
            layout = remote.get("value")
            if not isinstance(layout, dict) or not layout.get("slideId"):
                raise RuntimeError(f"Rendered slide {index + 1} is missing its data-slide-id.")
            if abs(layout["slideBox"]["width"] - SLIDE_WIDTH_PX) > .1 or abs(layout["slideBox"]["height"] - SLIDE_HEIGHT_PX) > .1:
                raise RuntimeError(f"Slide {index + 1} must render at exactly 1920 x 1080 CSS pixels.")
            fonts_by_selector: dict[str, list[dict]] = {}
            for item in layout.get("text") or []:
                selector = item["selector"]
                if selector not in fonts_by_selector:
                    selected = _cdp_command(socket, command_id, "DOM.querySelector", {"nodeId": document_id, "selector": selector}); command_id += 1
                    resolved = _cdp_command(socket, command_id, "CSS.getPlatformFontsForNode", {"nodeId": selected["nodeId"]}); command_id += 1
                    fonts_by_selector[selector] = resolved.get("fonts") or []
                fonts = fonts_by_selector[selector]
                item["style"]["cssFontFamily"] = item["style"]["fontFamily"]
                item["style"]["platformFonts"] = fonts
                if fonts:
                    # CSS stacks may begin with absent fonts. Use the actual browser font,
                    # not that uninstalled first choice, and explicitly write East Asian font.
                    dominant = max(fonts, key=lambda value: value.get("glyphCount", 0))
                    item["style"]["fontFamily"] = dominant["familyName"]
                    if any(value.get("isCustomFont") for value in fonts):
                        item["native"] = False
                        item["reason"] = "Browser web fonts require an installed PowerPoint typeface; use an installed font in source."
                else:
                    item["native"] = False
                    item["reason"] = "Could not verify the rendered platform font for visible text."
            screenshot = _cdp_command(socket, command_id, "Page.captureScreenshot", {"format": "png", "fromSurface": True, "captureBeyondViewport": True, "clip": layout["captureClip"]}); command_id += 1
            destination = background_dir / f"slide-{index + 1:03d}.png"
            destination.write_bytes(base64.b64decode(screenshot["data"]))
            layout["background"] = str(destination)
            results.append(layout)
    finally:
        socket.close(); _stop_browser(process, profile); profile_ctx.cleanup()
    return results


def classify_rendered_slide(layout: dict) -> dict:
    """Pure coverage classification for a final rendered slide."""
    native_text = [item for item in layout.get("text") or [] if item.get("native")]
    native_images = [item for item in layout.get("images") or [] if item.get("native")]
    unsupported = list(layout.get("unsupported") or []) + [
        {"kind": "rendered-text", "selector": item.get("selector"), "text": item.get("text", ""), "reason": item.get("reason") or "not native", "preserved_in": "blocked"}
        for item in layout.get("text") or [] if not item.get("native")
    ] + [
        {"kind": "rendered-image", "selector": item.get("selector"), "source": item.get("src", ""), "reason": item.get("reason") or "not native", "preserved_in": "background"}
        for item in layout.get("images") or [] if not item.get("native")
    ]
    return {"native_text": native_text, "native_images": native_images, "unsupported": unsupported}


def _px_to_emu(value: float) -> int:
    return round(float(value) * SLIDE_WIDTH_EMU / SLIDE_WIDTH_PX)


def _rgb(value: str) -> tuple[int, int, int] | None:
    values = re.findall(r"[\d.]+", value or "")
    return tuple(max(0, min(255, round(float(item)))) for item in values[:3]) if len(values) >= 3 else None  # type: ignore[return-value]


def _font_name(value: str) -> str:
    return (value or "Arial").split(",", 1)[0].strip().strip("\"'") or "Arial"


def _position_fraction(token: str) -> float:
    token = token.strip().lower()
    if token in {"left", "top"}: return 0.0
    if token in {"right", "bottom"}: return 1.0
    if token.endswith("%"):
        try: return max(0, min(1, float(token[:-1]) / 100))
        except ValueError: pass
    return .5


def _object_position(value: str) -> tuple[float, float]:
    parts = value.split() or ["50%"]
    return _position_fraction(parts[0]), _position_fraction(parts[1] if len(parts) > 1 else parts[0])


def _set_east_asian_font(run: Any, typeface: str) -> None:
    from lxml import etree
    from pptx.oxml.ns import qn
    properties = run._r.get_or_add_rPr(); east_asia = properties.find(qn("a:ea"))
    if east_asia is None: east_asia = etree.SubElement(properties, qn("a:ea"))
    east_asia.set("typeface", typeface)


def _add_text(slide: Any, item: dict) -> None:
    """Restore one measured visual-line segment without reflow or local font scaling."""
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.util import Pt
    from pptx.oxml.xmlchemy import OxmlElement

    box = item["box"]
    style = item.get("style") or item
    font_size = float(str(style.get("fontSize") or "16").removesuffix("px"))
    # Frames have a small font-metric allowance. Glyph bounds, not these frames,
    # are what the PowerPoint auditor checks for collisions and clipping.
    width = float(box["width"]) + max(4, font_size * .12)
    line_height = float(str(style.get("lineHeight") or font_size * 1.2).removesuffix("px"))
    height = max(float(box["height"]), font_size * 1.45, line_height)
    shape = slide.shapes.add_textbox(
        _px_to_emu(box["x"]), _px_to_emu(box["y"]),
        max(1, _px_to_emu(width)), max(1, _px_to_emu(height)),
    )
    shape.name = item.get("shape_name") or "pptx-text|body|source|l001|r001"
    frame = shape.text_frame
    frame.clear()
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    frame.word_wrap = False
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.vertical_anchor = MSO_ANCHOR.TOP
    frame._txBody.bodyPr.set("vertOverflow", "overflow")
    frame._txBody.bodyPr.set("horzOverflow", "overflow")
    paragraph = frame.paragraphs[0]
    paragraph.space_before = paragraph.space_after = Pt(0)
    paragraph.line_spacing = Pt(line_height * POINTS_PER_PX)
    paragraph.alignment = PP_ALIGN.LEFT  # Range x already includes CSS alignment.
    run = paragraph.add_run()
    run.text = item["text"]
    name = _font_name(style.get("fontFamily", ""))
    run.font.name = name
    _set_east_asian_font(run, name)
    run.font.size = Pt(font_size * POINTS_PER_PX)
    weight = str(style.get("fontWeight") or "400")
    run.font.bold = weight == "bold" or (weight.isdigit() and int(weight) >= 600)
    run.font.italic = style.get("fontStyle") in {"italic", "oblique"}
    properties = run._r.get_or_add_rPr()
    properties.set("lang", "zh-CN" if re.search(r"[\u3400-\u9fff]", item["text"]) else "en-US")
    spacing = float(str(style.get("letterSpacing") or "0").removesuffix("px"))
    if spacing:
        properties.set("spc", str(round(spacing * POINTS_PER_PX * 100)))
    if color := _rgb(style.get("color", "")):
        run.font.color.rgb = RGBColor(*color)
    alpha_match = re.match(r"rgba\([^,]+,[^,]+,[^,]+,\s*([\d.]+)\)", style.get("color", ""))
    if alpha_match:
        alpha = OxmlElement("a:alpha")
        alpha.set("val", str(round(float(alpha_match[1]) * 100000)))
        color_element = properties.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
        if color_element is not None:
            color_element.append(alpha)
    item["shape_box"] = {"x": box["x"], "y": box["y"], "width": width, "height": height}


def _source_to_file(source: str, project: Path, directory: Path, ordinal: int) -> tuple[Path, str | None]:
    """Resolve final-DOM data/local image sources for python-pptx."""
    if source.startswith("data:"):
        header, separator, payload = source.partition(",")
        if not separator or ";base64" not in header.lower(): return Path(), "image data URL is not base64 encoded"
        mime = header[5:].split(";", 1)[0].lower(); suffix = {"image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif", "image/webp": ".webp"}.get(mime)
        if not suffix: return Path(), f"{mime or 'unknown'} cannot be embedded reliably by python-pptx"
        target = directory / f"rendered-{ordinal:03d}{suffix}"
        try: target.write_bytes(base64.b64decode(payload, validate=True))
        except ValueError: return Path(), "image data URL is corrupt"
        return target, None
    if source.startswith("file:"):
        from urllib.parse import unquote, urlparse
        path = Path(unquote(urlparse(source).path)).resolve()
        if not path.is_relative_to(project) or not path.is_file(): return Path(), "rendered image is outside the project or missing"
        if path.suffix.lower() not in {".png", ".jpg", ".jpeg", ".gif", ".webp"}: return Path(), f"{path.suffix or 'unknown'} cannot be embedded reliably by python-pptx"
        return path, None
    return Path(), "rendered image is not a local or embedded source"


def _add_picture(slide: Any, image_path: Path, item: dict) -> None:
    box = item["box"]; left, top, width, height = _px_to_emu(box["x"]), _px_to_emu(box["y"]), _px_to_emu(box["width"]), _px_to_emu(box["height"])
    source_width, source_height = max(1, int(item.get("naturalWidth") or 1)), max(1, int(item.get("naturalHeight") or 1))
    fit, x_fraction, y_fraction = item.get("objectFit") or "fill", *_object_position(item.get("objectPosition") or "50% 50%")
    if fit == "contain":
        scale = min(width / source_width, height / source_height); picture_width, picture_height = round(source_width * scale), round(source_height * scale)
        slide.shapes.add_picture(str(image_path), left + round((width - picture_width) * x_fraction), top + round((height - picture_height) * y_fraction), picture_width, picture_height)
    else:
        picture = slide.shapes.add_picture(str(image_path), left, top, width, height)
        if fit == "cover":
            source_ratio, box_ratio = source_width / source_height, width / height
            if source_ratio > box_ratio:
                shown = box_ratio / source_ratio; picture.crop_left = (1 - shown) * x_fraction; picture.crop_right = (1 - shown) * (1 - x_fraction)
            elif source_ratio < box_ratio:
                shown = source_ratio / box_ratio; picture.crop_top = (1 - shown) * y_fraction; picture.crop_bottom = (1 - shown) * (1 - y_fraction)


def _validate_pptx(path: Path, expected_slides: int, pptx: Any) -> None:
    reopened = pptx.Presentation(str(path))
    if len(reopened.slides) != expected_slides or reopened.slide_width != SLIDE_WIDTH_EMU or reopened.slide_height != SLIDE_HEIGHT_EMU: raise RuntimeError("PPTX verification found an invalid slide set or size.")


def _install_pair_atomically(pptx_temp: Path, output: Path, report_temp: Path, report: Path) -> None:
    backups: list[tuple[Path, Path]] = []; installed: list[Path] = []
    try:
        for destination in (output, report):
            if destination.exists():
                fd, name = tempfile.mkstemp(prefix=f".{destination.name}.", suffix=".bak", dir=destination.parent); os.close(fd); backup = Path(name); backup.unlink(); os.replace(destination, backup); backups.append((destination, backup))
        for temporary, destination in ((pptx_temp, output), (report_temp, report)): os.replace(temporary, destination); installed.append(destination)
    except Exception:
        for item in installed: item.unlink(missing_ok=True)
        for destination, backup in reversed(backups): os.replace(backup, destination)
        raise
    finally:
        for _, backup in backups: backup.unlink(missing_ok=True)


def export_hybrid_pptx(*, project: Path, deck: dict, html_path: Path, chrome: str, output: Path, report_output: Path, capture: Callable[[str, Path, dict, Path], list[dict]] = collect_render_layers) -> dict:
    """Restore all visible HTML text as native segments over faithful image/CSS layers.

    Successful export proves package and source coverage, not visual acceptance.
    Skill-maintenance regression may audit this report and file separately. Normal generation ends after export.
    """
    pptx = require_python_pptx(); project, html_path, output, report_output = project.resolve(), html_path.resolve(), output.resolve(), report_output.resolve()
    if not html_path.is_file(): raise SystemExit(f"Canonical HTML is missing: {html_path}")
    if output.suffix.lower() != ".pptx" or report_output.suffix.lower() != ".json" or output == report_output: raise SystemExit("PPTX output must be .pptx and coverage output must be a distinct .json file.")
    if not output.parent.is_dir() or not report_output.parent.is_dir(): raise SystemExit("PPTX output parent directories must already exist.")
    paths = deck.get("slides") if isinstance(deck, dict) else None
    if not isinstance(paths, list) or not paths: raise SystemExit("PPTX export requires a deck manifest with slide HTML paths.")
    work = Path(tempfile.mkdtemp(prefix=".pptx-pptx-", dir=project)); fd, name = tempfile.mkstemp(prefix=f".{output.name}.", suffix=".tmp", dir=output.parent); os.close(fd); pptx_temp = Path(name); fd, name = tempfile.mkstemp(prefix=f".{report_output.name}.", suffix=".tmp", dir=report_output.parent); os.close(fd); report_temp = Path(name)
    try:
        backgrounds = work / "backgrounds"; backgrounds.mkdir(); layouts = capture(chrome, html_path, deck, backgrounds)
        if len(layouts) != len(paths): raise RuntimeError("Browser capture returned an incomplete slide set.")
        blocking = []
        for page, layout in enumerate(layouts, start=1):
            for finding in classify_rendered_slide(layout)["unsupported"]:
                if finding["kind"] != "rendered-image":
                    blocking.append({"page": page, "slide_id": layout.get("slideId"), **finding})
        if blocking:
            blocked_report = report_output.with_name(report_output.stem + ".blocked.json")
            blocked_report.write_text(json.dumps({
                "schema_version": COVERAGE_SCHEMA, "ok": False,
                "canonical_html": str(html_path), "canonical_html_sha256": _sha256(html_path),
                "blocking_findings": blocking,
                "required_action": "Repair the canonical HTML; do not shrink individual text boxes or silently rasterize visible text.",
            }, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            raise RuntimeError(f"PPTX export blocked by {len(blocking)} unsupported visible text/canvas regions. Repair source HTML. Evidence: {blocked_report}")
        presentation = pptx.Presentation(); presentation.slide_width = SLIDE_WIDTH_EMU; presentation.slide_height = SLIDE_HEIGHT_EMU; presentation.core_properties.title = str(deck.get("title") or "pptx"); presentation.core_properties.created = presentation.core_properties.modified = datetime(2000, 1, 1, tzinfo=timezone.utc)
        while presentation.slides: presentation.part.drop_rel(presentation.slides._sldIdLst[0].rId); del presentation.slides._sldIdLst[0]
        per_slide: list[dict] = []; total_text = total_media = total_unsupported = 0
        for page, layout in enumerate(layouts, start=1):
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            background = slide.shapes.add_picture(layout["background"], 0, 0, width=SLIDE_WIDTH_EMU, height=SLIDE_HEIGHT_EMU)
            background.name = f"pptx-background|{layout.get('slideId') or page}"
            classified = classify_rendered_slide(layout); unsupported = list(classified["unsupported"])
            for item in classified["native_text"]: _add_text(slide, item)
            native_media = 0
            for ordinal, item in enumerate(classified["native_images"], start=1):
                image_path, reason = _source_to_file(str(item.get("src") or ""), project, work, page * 1000 + ordinal)
                if reason: unsupported.append({"kind": "rendered-image", "selector": item.get("selector"), "source": item.get("src", ""), "reason": reason, "preserved_in": "background"}); continue
                _add_picture(slide, image_path, item); native_media += 1
            unsupported_count = len(unsupported)
            per_slide.append({
                "page": page, "slide_id": layout.get("slideId"), "title": layout.get("title"),
                "native_text_count": len(classified["native_text"]), "native_media_count": native_media,
                "text_layout": classified["native_text"],
                "capture_clip": layout.get("captureClip"),
                "source_text_character_count": sum(len(item["text"]) for item in layout.get("text") or []),
                "rasterized_regions": [{"kind": "deterministic-background", "source": str(html_path), "sha256": _sha256(Path(layout["background"])), "contains": "CSS geometry, decoration, and source image pixels; no ordinary visible HTML text"}],
                "unsupported_structure_count": unsupported_count, "unsupported_structures": unsupported,
            })
            total_text += len(classified["native_text"]); total_media += native_media; total_unsupported += unsupported_count
        presentation.save(str(pptx_temp)); _validate_pptx(pptx_temp, len(paths), pptx)
        theme = validate_theme(deck.get("theme")); palette = {"name": theme["palette"], **PALETTES[theme["palette"]]}; profile = theme["typography"]; typeface = _font_name(TYPOGRAPHY[profile]["font-zh"])
        native_total, denominator = total_text + total_media, total_text + total_media + total_unsupported
        report = {"schema_version": COVERAGE_SCHEMA, "ok": True, "project": str(project), "canonical_html": str(html_path), "canonical_html_sha256": _sha256(html_path), "pptx": str(output), "pptx_sha256": _sha256(pptx_temp), "slide_size": {"width_emu": SLIDE_WIDTH_EMU, "height_emu": SLIDE_HEIGHT_EMU, "width_px": SLIDE_WIDTH_PX, "height_px": SLIDE_HEIGHT_PX, "points_per_px": POINTS_PER_PX, "aspect_ratio": "16:9"}, "theme": {"palette": palette, "typography": profile, "typeface": typeface}, "summary": {"slide_count": len(per_slide), "native_text_count": total_text, "native_media_count": total_media, "native_object_count": native_total, "rasterized_region_count": len(per_slide), "unsupported_structure_count": total_unsupported, "coverage_denominator": denominator, "editable_coverage_percent": round(100 * native_total / denominator, 2) if denominator else 100.0, "editable_text_coverage_percent": 100.0, "source_text_character_count": sum(page["source_text_character_count"] for page in per_slide)}, "methodology": "Browser Range measures every visible HTML text node by visual line and inline style; each segment becomes a native single-line text box at exactly 0.5 pt per CSS pixel, with no wrapping or local font scaling. Actual browser platform fonts are written explicitly for Latin and East Asian text. Image pixels and CSS geometry remain in a raster background. Text within source images is outside HTML text editability coverage. The export metadata does not assert PowerPoint visual acceptance; skill-maintenance regression is a separate operation.", "slides": per_slide}
        report_temp.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"); _install_pair_atomically(pptx_temp, output, report_temp, report_output); return report
    finally:
        pptx_temp.unlink(missing_ok=True); report_temp.unlink(missing_ok=True); shutil.rmtree(work, ignore_errors=True)
