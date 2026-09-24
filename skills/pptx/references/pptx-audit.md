# PPTX 技能维护验证器与审计规则

本审计用于修改 pptx 技能之后，验证技能、生成器和排版防护规则是否仍然有效。普通演示稿的 `build` / `export-pptx` 不自动调用审计、不生成审计报告，也不要求 PowerPoint COM。维护审计使用独立样稿和故意损坏的负例；不得直接在审计器中缩字、改行距或移动成品对象。

## 调用

技能修改完成、结构清单与 package manifest 更新后，执行一次维护审计：

```text
python scripts/pptx_cli.py audit-skill --output <技能目录之外的输出目录>
```

也可以直接运行 `python scripts/audit_skill.py --output <外部目录>`。可选 `--chrome <浏览器可执行文件>`。维护 runner 不修改技能源码或 manifest；它验证结构与包完整性，使用普通 CLI 创建两页中文 HTML 样稿，执行普通 `build` 和 `export-pptx`，证明普通生成没有审计输出，再显式审计生成的 PPTX 并运行九项正负回归。样稿包括跨页标题/正文、多行中文和内联强调。

维护报告固定为外部输出目录的 `skill-audit.json`。每次运行在 `runs/<时间与随机ID>/` 保留独立样稿、实际 PowerPoint PNG、单文件审计和回归报告。总报告记录技能树前后哈希、被测生成器文件哈希、结构/manifest 检查、真实端到端渲染、九项回归结果、输入 HTML 未变化证据和产物路径。输出目录不得位于技能树内。缺少真实渲染器时维护结果为 `blocked`，返回非零；不能判为维护通过。

仅在维护排错或用户明确要求检查某份 PPTX 时，独立调用单文件检查器：

```text
python scripts/pptx_audit.py <文件.pptx> --project <项目目录> --source-report <pptx-editability.json> --report <pptx-audit.json> --render-dir <渲染目录> --require-render
```

`--source-report` 未指定时查找项目/成品旁的 `pptx-editability.json` 或与成品同名的 `.editability.json`；显式指定的报告不存在时不会偷偷改用其他报告。项目生成流程必须提供源布局证据。

Python 接口：

```python
audit_pptx(pptx_path, report_path=None, render_dir=None, rules_path=None,
           require_render=False, project=None, source_report=None) -> dict
```

依赖为 `python-pptx`、Pillow，以及 Windows 的 `pywin32` 和 Microsoft PowerPoint。PowerPoint 使用新的 COM 实例、只读打开、逐页导出 PNG，关闭时不保存。审计前后比较 PPTX SHA256；输入字节变化属于阻塞错误。渲染目录只放实际 PowerPoint 生成的图片，不能用浏览器截图代替。

## 结果含义

| 状态 | 含义 | 退出码 |
| --- | --- | --- |
| `passed`，`ok=true` | 结构、源文字、样式、实际文字测量和所有页面渲染通过；无阻塞或未验证项 | 0 |
| `failed`，`ok=false` | 至少一个阻塞错误，或使用 `--require-render` 时没有可用 PowerPoint | 1 |
| `not_verified`，`ok=false` | 没有已知阻塞错误，但缺少实际渲染或某类对象无法完整测量 | 2 |

`render.status=passed` 仅表示渲染和测量过程运行完毕，不独自代表整份文件通过；必须同时检查顶层 `ok`。没有渲染器时即使结构完全正常，也不会输出整体通过。`PPTX_AUDIT_DISABLE_RENDER=1` 仅供测试该失败边界。

JSON 报告包括输入与规则 SHA256、输入未改变证据、源布局报告及其 SHA256、逐页 PNG 路径和 SHA256、实际文字行位置、字体、问题代码及对应页/对象、错误计数和检查局限。

## 执行的规则

所有机器阈值与例外均集中在 [pptx-audit-rules.json](pptx-audit-rules.json)。以下是规则类别与失败处理。

| 类别 | 自动检查 | 失败处理 |
| --- | --- | --- |
| 页面完整性 | 文件可打开、页数非零、16:9、源布局页数一致 | 修正导出或缺页 |
| 原生文字与保真 | 每段源文字有唯一原生文本对象；文本逐字匹配；源布局哈希绑定当前 PPTX；无文字栅格化 | 修正缺字、重字、标签或导出映射 |
| 同级样式 | 有限语义角色；同角色跨页的字体、字号、粗体、斜体、颜色、字距、PPT 行距、源 CSS 行高一致 | 修改共享样式，不逐页缩放 |
| 层级 | 高层级字号不小于低层级；字号不能低于规则下限 | 修正文本角色或全局字号 |
| 字体 | 明确字体/字号/颜色、GDI 与注册表字体可用性、PowerPoint 返回字体名 | 安装真实字体或统一改用已安装字体 |
| 多行 | 原生多行和拆分视觉行的源行高下限；PowerPoint 实际行范围碰撞 | 增大行距、扩展区域或拆页 |
| 溢出 | 每个实际文本行不得越出文本框或页面 | 扩展文本框/空间或拆页 |
| 文字遮挡 | 实际文本行两两相交；前方不透明对象覆盖文字；图片检查对应区域 alpha | 调整位置或堆叠顺序 |
| 编辑安全 | 禁止自动缩小文字；审计不保存成品；输入哈希不变 | 修复生成设置或外部并发修改 |

角色仅允许：`cover-title`、`title`、`subtitle`、`heading`、`subheading`、`body`、`lead`、`caption`、`label`、`code`、`metric`、`meta`，以及 HTML 语义别名 `h1`、`h2`、`h3`、`p`。不得创建每页独有角色来规避一致性检查。

源布局 `slides[].text_layout[]` 每条记录包含 `shape_name`、`text`、`role`、`source_id`、`line_id`、`segment_id`、`style`、`box`、`native`。`style` 使用 CSS 的 `fontFamily`、`fontSize`、`fontWeight`、`fontStyle`、`color`、`lineHeight`、`letterSpacing`；字号、行高与坐标单位均为 1920×1080 舞台的 CSS px。原生 PPT 使用 960×540 pt，因此 1 px = 0.5 pt。

对象名为 `pptx-text|role|source_id|line_id|segment_id`。一条视觉行中的内联样式可拆为多个原生段，源行高仍跨页参与审计；不得利用拆框隐藏行距漂移。单视觉行的 PPT 段落可以统一使用源行距或单倍行距，同级不能混用。

## 验证回归

```text
python scripts/test_pptx_audit.py --output <工作区外的测试输出目录>
```

测试自动生成并审计正常中文两页稿、同级样式漂移、文字相交、过密多行、前景遮挡、长行只遮一个字、文本框溢出、无渲染器和严格要求渲染器九种情况。正常稿还包含文字后方的不透明背景，证明背景不被误判为前景遮挡。每份输入检查审计前后哈希不变，测试结果写入 `regression-results.json`。

## 自动测量的边界与规则例外

PowerPoint 的完整 TextRange 会把不可见的末尾段落符计入宽度。审计按实际行去掉末尾 CR/VT，再测量其可见 Characters 范围，避免把隐形控制符当成溢出；页面/文本框公差仍为 1 pt。行碰撞保留 0.75 pt 的坐标公差，相交面积达到较小文字行的 1% 即阻塞；前景遮挡达到目标文字行的 1% 即阻塞。对很短或极细的局部相交仍须查看渲染图。

Bound 矩形不是逐像素字形轮廓。部分 PowerPoint 多行范围的高度会受段落行高截断，审计用原生/源多行最小行距规则补充，不能仅凭 BoundHeight 判断安全。字体检查不能证明每个字形都未发生局部回退。渐变填充、任意复杂轮廓、分组、旋转、表格等无法完整验证时报告 `not_verified`，不能凭未检测到错误改判通过。后台图片按堆叠顺序排除，前景图片按相交区域透明度检查；阴影、描边和复杂字形的视觉影响仍需逐页查看实际 PNG。

维护审计之后由执行技能修改的代理查看样稿的 PowerPoint PNG，检查文字可读性、行间距、遮挡、图片和内容完整性。发现生成器问题时修复技能并重新执行维护审计。这个流程不加入普通演示稿生成链，也不设置用户审批停点。

合法例外只允许明确命名、属性范围受限的变体：`emphasis` 只能把粗体改为 `true`，通过源 `emphasis=true` 或 `style.variant=emphasis` 声明；`inverse` 只能把颜色改为 `FFFFFF`，通过源 `style.variant=inverse` 声明。其余属性仍必须与同级基础样式一致。`regular` 等同默认样式。未知变体一律报错，不接受每对象的 `ignore`、任意异常原因或整页跳过。新增真实的公共变体必须先修改规则 JSON 的白名单与准确属性约束，补充正常和违规回归，再使用同一规则对整套成品重新审计；不得为让某份成品通过而扩大碰撞公差、降低遮挡门槛或关闭检查。
