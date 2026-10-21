# HTML/CSS → 编辑器 JSON：代码审计与接入标准

日期：2026-10-21。状态：已接入本地产品转换接口和编辑器双向桥。只持久化 PPTX，JSON 为转换时的内存表示；尚未部署线上。下文是初始审计和目标差距，当前实际范围见 [design-layer.md](design-layer.md)。

## 决策

保留编辑器 JSON。AI 整页设计使用固定 1280 × 720 的 HTML/CSS，允许用 flex/grid 排版，图标、插画、总结图可用 SVG 或位图资产。浏览器计算完布局后，将结果冻结为绝对几何，再导入当前编辑器和 PPTX 写入链路。

```text
AI → HTML/CSS + SVG/图片资产（每页 1280 × 720）
   → 浏览器加载字体、图片，完成布局和检查
   → 测量后的场景 JSON（坐标、实际文本行、样式、绘制顺序）
   → 原生对象导入桥（编辑器对象 + OOXML + 图片关系 + 历史记录）
   → 现有编辑器 / 导出 PPTX
```

HTML 是生成输入，当前 JSON 是编辑状态。用户拖拽、改文字后，以编辑器状态为准。不能用旧 HTML 自动覆盖已编辑的页面；重新设计需要明确的目标页面和版本号。

模板提供品牌、配色、字体、参考布局和可复用素材。除明确锁定的品牌元素外，不强制按槽位填充。没有布局或素材参考时允许 AI 自由设计。页数按需求规划，5 页只是之前的测试条件。

这条路线能减少模型手算坐标和换行的负担，不保证 HTML 天然比 SVG 美观；设计质量仍需检查。转换后不得再次套自动布局、自动缩字或自动适配。

## 当前代码实际有三层

1. **渲染模型**：`src/pptx-virtual-dom.ts` 中的 `PptVirtualDocument`，可表达较丰富的排版与图形。
2. **编辑命令**：`src/api/schema.ts` 的 `apply_edits` 命令，开放的字段少于渲染模型。
3. **文件写回**：`src/document-session.ts` 将编辑记录写回原 PPTX ZIP 包，不是任意 JSON 到 PPTX 的通用序列化器。

关键证据：`DocumentSession.openVirtual()` 会丢弃原包；`canExport()` 要求原包存在；`exportPptx()` 依赖 ZIP、文字/样式/结构编辑记录。仅构造一份看起来正确的 `PptVirtualDocument`，不能证明能够导出可编辑 PPTX。

因此，新增层对调用方可以叫“HTML → JSON”，内部必须同时接通可编辑对象和文件写入。

## 能力与缺口

| 特性 | 当前已有能力 | 新路线需要补什么 |
|---|---|---|
| 绝对位置、尺寸、旋转 | `Rect`、元素旋转、source 坐标变换 | 浏览器测量；明确 CSS px → 目标文件 EMU → 编辑器逻辑坐标的换算 |
| 富文本 | `TextParagraph` / `TextRun` 支持同段多种样式 | 新建导入接口、对应 XML writer，不能只依赖简化的 `addText` |
| 字体与中文字体 | `fontFamily`、`eastAsianFontFamily` 已在渲染模型中 | 测量字体核验，以及新建 run 的字体写入；三处字体一致 |
| 行距、段前后距、字距 | `ParagraphStyle` / `TextStyle` 已有字段 | 编辑命令和 writer 未完整开放；导入需显式写入并测试 |
| 文字内边距、垂直对齐、换行策略 | `TextBodyProperties` 已有字段 | 按浏览器内容框写入 `bodyPr`，取消默认边距和自动缩字干扰 |
| 纯色、线性渐变、透明度 | 模型及 `pptx-style-writer.ts` 有可复用能力 | CSS 色值、渐变角度/方向/停止点适配；区分填充透明度与整个子树合成透明度 |
| 阴影 | 模型含多种效果，常规接口可写单个外阴影 | 首版限定单个无 spread 外阴影；模糊值差异需对比确认 |
| 圆角 | 模型有 normalized `cornerRadius`；新建 roundRect 有默认值 | CSS px 半径到 OOXML adjustment 的映射；不等角圆角首版拒绝 |
| 图片与裁剪 | `ImageElement.crop`、图片创建和媒体关系 | `object-fit` / `object-position` 映射、资源加载检查 |
| SVG / 自定义路径 | 已能表示解析得到的 `customGeometry` | 不能等同于已能新建任意 SVG 路径；首版沿用 SVG 源资产 + PNG 显示资产，保留以后替换能力 |
| 图层与组合 | 现有顺序、source、group 数据和排列命令 | 正确还原 HTML 绘制顺序；DOM 遍历顺序或全局排序 z-index 都不充分 |
| 整页导入与导出 | 现有 PPTX 打开、编辑、导出可复用 | 新增原生页面导入桥，建立 XML、shape ID、rels、历史与渲染对象的一致性 |
| 浏览器布局服务 | 本机前端依赖可找到 Playwright 和 Chromium | 后端/渲染服务还未接入该转换能力；部署镜像需固定浏览器与字体版本 |

“已有字段”表示解析/渲染模型可承载，不表示新内容的创建、编辑、导出、重开已经全部验证。

### 文字与背景错位的具体风险

现有 `buildShapeXml()` 和 `buildShapeElements()` 会分别生成 XML 和渲染对象。文字框有默认排版属性，普通形状内文字与独立文字框的默认锚点也不同。

HTML 的 border box、padding box、content box 不相同。若把背景矩形和文字都直接套同一个 `getBoundingClientRect()`，再叠加 PPT 默认内边距，就会出现“坐标相同但看起来不居中”。此外，字体回退、行高与基线不同也会产生偏移。

导入时必须从同一个 DOM 对象派生背景和文字内容框，记录两者关系；分别写准边界、内边距、实际文本行和基线。能够共用一个 PPT shape/txBody 的简单对象可合并，否则保留关联和绘制顺序。不要靠额外挪动文字或自动缩字掩盖差异。

## 首版接入顺序

### P0：原生对象导入桥

先完成一页测量结果到真实编辑对象的导入，而不是先做只能预览的 DOM 抓取器。

- 在现有 PPTX 会话中创建/替换页面；无模板时使用项目自有的空白 PPTX 包。
- 为背景、文本、图片生成真实 OOXML、稳定 shape ID、媒体关系和对应编辑器对象。
- 补齐新建富文本的字体、字号、行距、段距、字距、内边距和垂直对齐写入。
- 一次页面导入作为一个事务，失败不留下半页；支持撤销与版本检查。
- 保留未被指定替换的页面和用户手动修改。
- 导入后立即导出并重开，确认编辑器显示与文件内容一致。

### P0：浏览器测量与 HTML 作者规范

规范详见 [html-authoring-contract.md](html-authoring-contract.md)。浏览器完成字体、资源和布局后，输出实际几何、文本行、富文本片段、资产和绘制顺序。转换器只翻译测量结果，不替 AI 重设计。

单位换算以目标文件尺寸为准：

```text
xEMU = xCss × deckWidthEMU / 1280
yEMU = yCss × deckHeightEMU / 720
```

目标宽高比必须为 16:9；其他比例必须显式处理，不能非等比拉伸。仅在目标采用 96 CSS px/英寸时，才可以直接用 9525 EMU/px、CSS 字号 px × 0.75 → pt。不能把这些常数套到所有旧模板的逻辑坐标上。

### P0：三阶段一致性检查

验证同一页的 **HTML 渲染 → 编辑器渲染 → 导出 PPTX 重开渲染**。先用一个技术样页验证中文换行、彩色背景居中、富文本、图片裁剪、渐变阴影、SVG 图标，再用一个当前标准模板接入。

- 检查文本内容、行数、坐标、样式、图层、资产引用与可编辑性。
- 对比截图；允许抗锯齿差异，不能容忍缺字、缺图、残留关键字、明显错位或遮挡。
- 结构检查和视觉检查分别报告结果。“未发现越界/重叠”不能作为“美观且内容完整”的结论。
- 视觉检查覆盖全部页面。20 页可以分批，但必须有覆盖记录，不能只抽查 1–4 页。
- 用户未要求时不重新生成业务 PPT；先完成转换链路的技术验证。

### P1：接入 Agent / AI PPT 插件

- Agent 继续管理模型与插件模型使用权限，转换服务不另建一套模型配置。
- 整页生成/重设计调用 HTML 作者流程；局部文字、颜色、位置修改继续走现有 JSON 编辑命令。
- 转换工具未上线前，不把线上 SKILL 改成调用不存在的 HTML 导入工具。
- 工具接入后刷新 Agent 缓存中的插件说明和 schema，避免模型沿用旧的填槽规则。

## Anthropic skill 核实

已用 `skill-installer` 下载两个固定版本，单独放在 `/home/goodtime/Documents/reference-skills`，来源和哈希见该目录的 `provenance.json`。

- 当前版本 `dbd4588f9e1033efb41dad4bef2f7947c8993d44`：新建 PPT 推荐 PptxGenJS，不含历史 `html2pptx` 流程。
- 历史版本 `69c0b1a0674149f27b61b2635f935524b6add202`：包含 `html2pptx.md` 与脚本；该流程直接生成 PPTX，不输出本项目 JSON。
- 历史指南把部分视觉效果作为图片资产处理，并不是完整 HTML/CSS 或任意 SVG 到原生 PPT 对象的转换器。不能用它证明本项目所有样式都已有可靠映射。

该子目录原样保留 Proprietary LICENSE。用户随后明确要求作为私有本地工具直接使用；历史 HTML 版已安装到 `/home/goodtime/.codex/skills/anthropic-pptx-html2pptx`，增加 1280 × 720 调用入口和项目约定，原始转换脚本保持不变。转换器现可作为后续适配基础；直接输出 PPTX 的能力不能替代本项目 JSON 导入桥。

官方来源：

- [当前 PPTX skill](https://github.com/anthropics/skills/blob/dbd4588f9e1033efb41dad4bef2f7947c8993d44/skills/pptx/SKILL.md)
- [历史 HTML 流程](https://github.com/anthropics/skills/blob/69c0b1a0674149f27b61b2635f935524b6add202/skills/pptx/html2pptx.md)
- [移除旧流程的更新](https://github.com/anthropics/skills/commit/4e6907a33c3c0c9ce7c1836980546aaba78a34b5)
- [PPTX 目录许可](https://github.com/anthropics/skills/blob/dbd4588f9e1033efb41dad4bef2f7947c8993d44/skills/pptx/LICENSE.txt)

## 代码定位

- `src/pptx-virtual-dom.ts:28`：source、坐标变换、shape 关系。
- `src/pptx-virtual-dom.ts:55`：文字、段落、文本框属性；`:152` 填充；`:193` 阴影；`:294` 形状；`:323` 图片。
- `src/api/schema.ts:102`：公开文字样式白名单及新建命令。
- `src/editor/text-ops.ts:15`：当前可编辑文字格式范围。
- `src/editor/new-shapes.ts:79`：新建 OOXML；`:120` 对应渲染对象。
- `src/pptx-text-writer.ts`：文字编辑写回范围。
- `src/pptx-style-writer.ts:98`：渐变；`:135` 阴影；`:185` 样式操作写回。
- `src/document-session.ts:150`：虚拟文档入口；`:170` 导出条件。
- `src/headless/session.ts:22`：无头会话当前以 PPTX 包为输入。
- `/home/goodtime/Documents/project/ai-ppt-server/service/ppt_runtime/deck.go:286`：现有 PPTX 渲染服务调用；`:349` 原生模板目录。

审计时新增文档与下载来源记录；用户随后要求实际安装，已增加活动 skill 和本地技术样页验证。未修改生产生成逻辑、未部署、未进行模型生成测试。验证产物位于 `/home/goodtime/Documents/deployments/html2pptx-skill-smoke-20261021`。
