---
name: pptx
description: "根据材料创建、修改和检查中文或其他语言的 16:9 演示稿，逐页维护 HTML 源码，自动生成带原生可编辑文字的 PPTX 和离线 HTML；技能变更时执行专用审计回归。适用于 PPT、PowerPoint、演示文稿和逐页重建；已有 PPTX 可作为内容或视觉参考，不承诺无损导入。"
---

# pptx

从用户材料直接完成大纲、逐页创作、预览、PPTX 生成和交付。审计报告仅在本技能变更时生成，不在普通 PPT 生成后运行审计。中间产物用于自行检查，不设置人工批准关卡，不要求用户逐步回复。已有主题和材料时直接推进；只有缺少无法合理推断的关键内容才提问，其余工作继续执行。

## 入口与内容来源

将入口解析为相对本文件的绝对路径。Windows 使用 `scripts/pptx.cmd`；其他系统使用 `scripts/pptx`。下文以 `scripts/pptx` 表示该入口，不调用系统中同名的命令。

- `slides/<id>.html`：该页文字、DOM、局部 CSS 和素材的唯一内容源。
- `deck.json`：页面顺序与全局主题。
- `outline.md`：参考大纲，可随创作调整。
- `runtime/`：全局字体、层级、颜色和组件。

保留用户指定的原稿、文件名、事实、来源和页码对应。翻译或拆页可以删去冗余，不能擅自删掉事实、原因、影响、操作步骤或测验。材料中的内部名称只在讲解接口、代码或文件结构的页面保留，其他页面说明实际用途。不要补造数据、功能、引文或因果关系。不依赖外部文风 skill。

生成的 `预览.html`、`演示文稿.html` 和 PPTX 均由源码重建，修复时回到真实页面，不手改成品。

## 执行流程

```text
scripts/pptx init <项目>
scripts/pptx status <项目> --json
```

续做时先运行 `batch <项目或父目录>`，再读取目标项目状态。修改指定页使用 `status <项目> --json --intent edit --slide <页码或ID>`。每完成一个动作，执行返回的 `rerun` 或 `command` 并继续读取状态。

<!-- next-action-contract:start -->
- `author_slides`：按照 `next.brief` 逐页创作；完成一页重新读取状态，全部材料覆盖后执行 `next.command_when_ready`。
- `edit_outline`：编辑 `next.path` 中的大纲，完成后执行 `next.rerun`，直接继续创作。
- `edit_slide`：依据 `next.path`、`next.issues` 修复该页，执行 `next.rerun`。
- `fix_media`：依据真实问题修复本页素材及相对路径，然后继续。
- `run_command`：执行 `next.command`，读取结果并继续。
- `complete`：当前源文件对应的 PPTX、离线 HTML 和可编辑性说明已生成，核对交付文件后完成。
<!-- next-action-contract:end -->

### 逐页创作

进入 `author_slides` 时读取 [页面设计规范](references/components.md)。先说明本页的对象、主要事实或关系、证据和材料缺口，再选择构图。

```text
scripts/pptx starter list --json
scripts/pptx starter show <名称>
scripts/pptx slide add <项目> <页面ID> --title "<标题>" --starter <名称>
```

页面 ID 使用小写字母、数字和连字符。一次编辑一张真实页面；只为该页编写 `.s-<id>` 局部 CSS。删除 starter 的示例文字、虚构数值、占位图和多余说明。需要调整顺序时使用 `slide move`，拆页用 `slide add` 或 `slide duplicate`。

全局字体、层级和颜色通过主题与共享 token 管理。页面数量由内容决定，内容过多先调整结构、拓宽文字区域或拆页。每页保持一个主要判断，普通 8–10 页演示尽量覆盖至少四类构图。

### 文本一致性与多行安全

给文字设置稳定的 `data-text-role`：`cover-title`、`title`、`heading`、`subheading`、`body`、`lead`、`caption`、`label`、`code`、`metric`、`meta`。同一角色跨页共享字体、字号、字重、颜色、字距与行距；封面标题与正文页标题属于不同角色。禁止用页码或元素 ID 创造角色逃避审计。

角色默认值及用法见 [页面设计规范](references/components.md)。反白文字或局部强调使用审计规则中预先列出的有限命名变体，并在源元素上声明 `data-text-variant`；不能把全部实际样式自动登记为例外。

多行文字保留浏览器的真实换行与足够行距。标题默认至少 1.3 倍行高，正文 1.55 倍；普通生成依据源页面真实行边界保留位置；技能维护时另用 PowerPoint 实测验证导出器。禁用逐框缩字和自动压缩。溢出、行间相交、对象遮挡时修正源文件后重新生成，不能调低检查门槛或修改报告使其通过。

### 自动生成与交付

所有材料和页面就绪后直接执行：

```text
scripts/pptx build <项目>
```

该命令连续完成 HTML 检查、正式构建和 PPTX 生成。也可用 `export-pptx <项目>` 重建并导出。普通生成不启动 PowerPoint、不调用审计器、不等待审计报告。源页面或生成器改变后重新生成，不能交付陈旧结果。

PPTX 以 1920×1080 浏览器舞台为输入，对应 960×540 pt，几何与字号统一换算为 0.5 pt/px。按实际视觉行提取文字，保留混合内联内容，写入原生文本框并显式指定中文字体；非文字图形和复杂 CSS 可保留为背景图片。背景不能重复包含已导出的文字。不要将这种交付描述为所有图形均可编辑。

交付目录包含 `演示文稿.pptx`、`演示文稿.html`、`pptx-editability.json` 和真实 HTML 源码。可编辑性说明是导出元数据，记录逐页原生文字及栅格图形的边界，不触发成品审计。普通项目不要求 `pptx-audit.json` 或 `pptx-render/`。

## 仅在技能变更时执行审计

修改本技能的脚本、组件、模板、样式或审计规则后，读取 [维护审计说明](references/pptx-audit.md)，先更新包清单，再运行维护审计：

```text
python -X utf8 scripts/package_manifest.py --write
scripts/pptx audit-skill --output <技能目录之外的审计输出目录>
```

`audit-skill` 检查技能与包完整性，生成独立测试稿，在维护输出目录中用 PowerPoint 实测和渲染，再运行正常与故障样例回归，输出 `skill-audit.json`。它是技能变更的验证步骤，不属于普通 `build`、`export-pptx` 或 `status` 流程。

机器规则为 [pptx-audit-rules.json](references/pptx-audit-rules.json)，底层工具为 `scripts/pptx_audit.py`。仅在技能维护或用户明确要求检查某个成品时调用。审计检查文本层级、字体、源文字、真实行间相交、溢出和前景遮挡；检查器只读输入，不缩字、不改成品。

维护审计完成后查看测试稿的逐页 PNG，放大多行标题和混合内联文字。没有实际渲染器时维护结果标为未验证，不能宣称技能变更通过排版验证；这不阻塞日常生成 PPT。发生维护回归错误时修复技能，再运行受影响检查。审计报告绑定被测技能版本、规则和测试文件哈希。

## 其他参考资料

- 外部素材与保真要求：[media.md](references/media.md)
- 概念插画：[illustration.md](references/illustration.md)
- 流程、关系、界面及图表：[programmatic-visuals.md](references/programmatic-visuals.md)
- 构建或导出故障：[troubleshooting.md](references/troubleshooting.md)
- 修改本 skill 的程序与设计系统：[evolution.md](references/evolution.md)

只读取当前任务相关的参考资料。

## 展示样本

[在线浏览 Agents 演示稿](https://fuetsui.github.io/pptx/Agents.html)；离线样本：[references/Agents.html](references/Agents.html)。样本展示 HTML 演示效果，不能作为完整的逐页源码项目继续构建。技能源码和安装说明见 [GitHub](https://github.com/FueTsui/pptx)。
