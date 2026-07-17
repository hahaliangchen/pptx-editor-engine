# PPTX 渲染问题记录

本文档记录当前 PPTX 解析、虚拟 DOM 和 Rust 排版渲染中已经处理的问题，以及仍然需要继续完善的部分。

## 总体分工

- TypeScript 读取 PPTX Open XML，建立页面元素、几何坐标、文本段落和样式来源关系。
- TypeScript 输出虚拟 DOM。这里的重点不是只保存最终颜色或字号，还要保留继承来源和排版语义。
- Rust 接收虚拟 DOM，使用后端提供的系统字体文件，通过 cosmic-text 做 shaping/layout，再用 swash 光栅化字形。
- Canvas 只负责合成最终位图，不再负责字体选择、字体测量或 `fillText` 渲染。

## TypeScript 模块边界

首轮模块化已经完成：

- `src/pptx-virtual-dom.ts`：虚拟 DOM、文本、形状和样式类型；
- `src/pptx-xml.ts`：命名空间兼容的 XML 查询工具；
- `src/pptx-style-resolver.ts`：theme 颜色/字体、填充、线条、效果和 style reference 解析；
- `src/pptx-parser.ts`：ZIP/XML 流程、placeholder 关系和元素组装。

`PptxParser` 中的旧私有样式方法暂时保留为迁移期 dead code，当前实际 shape/text 样式路径已经通过 `PptxStyleResolver`。后续清理时只删除重复实现，不改变虚拟 DOM 数据结构。

## Rust 模块边界

Rust 首轮已经拆出：

- `rust-engine/src/shape_renderer.rs`：shape 几何路径、填充、渐变和描边；
- `rust-engine/src/effects.rs`：shape alpha mask、自定义高斯模糊、阴影缩放和 RGBA 合成；
- `rust-engine/src/text_layout.rs`：字体指标、paragraph/run 布局、CJK 换行、行距和 bullet buffer；
- `rust-engine/src/font_renderer.rs`：glyph 光栅化、像素 alpha blend、悬挂标点和文本诊断；
- `rust-engine/src/image_renderer.rs`：图片平滑、裁剪和 Canvas 合成；
- `rust-engine/src/lib.rs`：WASM 生命周期、字体注册、元素遍历和模块编排；不再保留 shape/effects 或文本布局的重复实现。

当前渲染路径已经调用上述模块。`lib.rs` 只负责 WASM 生命周期、字体注册、文本位图合成和元素遍历，具体的 shape、effect、image、layout 和 glyph 渲染分别由对应模块实现。

## 已处理的问题

### 1. 坐标空间和绝对视觉单位混用

早期实现把 group 的累计 `scaleX/scaleY` 同时用于几何坐标和字号、边框宽度，导致字号或边框被放大到异常尺寸。

现在分成两类值：

- `scaleX/scaleY`：只用于 shape、group、图片、表格和路径的几何坐标。
- `absoluteUnitScale`：用于 EMU/pt 转换，例如字号、边框宽度、固定内边距、阴影半径和阴影距离。

### 2. 样式继承关系过浅

虚拟 DOM 现在保留以下样式来源：

```text
theme -> master -> layout -> placeholder -> shape
```

shape 节点包含：

- `styleRefs`
- `computedStyle`
- `styleTrace`

支持了 theme 的填充、线条和效果引用，以及直接样式覆盖。

### 3. 填充、边框和效果丢失

已补充：

- solid fill
- linear/radial gradient
- line width、dash、cap、join
- outer shadow
- glow
- `prstClr`、`scrgbClr` 和主题色转换

同时修复了 `a:ln/a:solidFill` 被错误当成 shape 填充的问题。

### 4. Rust 字体渲染链路

当前链路为：

```text
后端系统字体文件
  -> Rust FontSystem
  -> cosmic-text shaping/layout
  -> swash glyph rasterization
  -> RGBA bitmap
  -> Canvas 合成
```

Rust 支持每个 run 的：

- 字体族
- 东亚字体族
- 字号
- 粗体
- 斜体
- 颜色
- 字符间距

使用无 hinting、灰度 alpha 和 2 倍内部光栅化，作为 mac-like 字体渲染配置。

### 5. 段落行距和富文本

虚拟 DOM 保留 paragraph/run 层级，不再把一个文本框强行压成单一 TextStyle。

已支持：

- `spcPct` 百分比行距
- `spcPts` 固定点数行距
- 段前和段后间距
- 段落对齐
- 多 run 字号、颜色、粗体、斜体和字体
- bodyPr 内边距、垂直锚点和 autofit 基础信息

百分比行距以实际后端字体的 alignment box 为基准：

```text
metricLineHeight + fontSize * (spcPct - 1)
```

当前针对 WPS 的渲染兼容层还会对字体 alignment box 使用 `0.92` 校准系数。这不是 PPTX 属性，只用于匹配 WPS 对本样例字体内部 leading 的处理。固定 `spcPts` 仍然直接按绝对点数换算。

### 6. 东亚排版细节

本次补充了以下 Open XML 段落字段：

- `eaLnBrk`：东亚字符换行
- `hangingPunct`：悬挂标点
- `fontAlgn`：字体在行框中的垂直对齐

Rust 对包含中文的段落使用 `Wrap::WordOrGlyph`，允许在合法字符边界回退换行，避免 `Wrap::Word` 导致中文长段落行数过少。

`hangingPunct` 已进入虚拟 DOM。当前换行使用同一个稳定的段落宽度，不能把半字宽全局加到每一行，否则后续行会错误多容纳一个 CJK 字符；行首开标点仍做视觉悬挂处理，行尾闭标点的精确越界规则还需要单独实现。`fontAlgn` 已进入虚拟 DOM；`auto` 使用 cosmic-text 根据实际字体 ascent/descent 计算的 leading。显式 top/bottom 的独立基线规则还需要更多 WPS/PowerPoint 样本校准。

当前行尾闭标点不再额外向右移动，避免句号与前一个 CJK 字符之间出现人为间隙；段落 `marL/indent` 使用横向布局比例，`spAutoFit` 文本会按实际布局高度生成位图。

文本框还会保留 `vertOverflow/horzOverflow`。默认 `overflow` 不再被 Rust 强制裁剪，只有明确的 `clip` 或 `ellipsis` 才限制绘制范围，避免底部内容在固定文本框中被截掉。

### 7. bullet 悬挂缩进

虚拟 DOM 会保留 `marL` 和 `indent`。Rust 将 bullet 单独排版，并把 bullet 的位置放在：

```text
body margin + paragraph margin + indent
```

正文从段落左边开始，非 bullet 段落的正缩进只作用于首行。

非 bullet 段落也必须区分首行和后续行的可用宽度。比如：

```xml
<a:pPr marL="525780" indent="-525780" />
```

表示首行相对 `marL` 向左悬挂，首行可用完整文本框宽度，后续行才从
`marL` 位置开始。不能只把首行绘制坐标左移，却把 `marL` 从所有行的
换行宽度中扣除。

样例第三页的原始 XML 验证结果：

- `打造个性化学` 文本框 `id=1048745` 的 `rect.w` 为 `169.2px`；`bodyPr` 没有
  写 `lIns/rIns`，因此使用 DrawingML 默认左右内边距各 `14.4px`，有效宽度为
  `140.4px`。文本框加宽时，这个右侧默认内边距不会消失。
- `智研：学科交叉分析` 文本框 `id=1048752` 的 `rect.w` 为 `248.1px`，段落
  `marL=525780`、`indent=-525780` 换算为 `82.8px`。首行有效宽度为 `219.3px`，
  后续行有效宽度为 `136.5px`。两处都没有 `a:br`，换行是自动排版结果。

### 8. 标题条红色光晕

WPS 中“目标客户”“智能体核心能力”“核心亮点”和“业务痛点”的标题条周围有红色柔化光晕。原始 PPTX 使用的是形状 `spPr` 内的 `a:outerShdw`，不是额外图片或渐变；样例参数包括 `blurRad=63500`、`sx/sy=105%`、居中对齐以及主题色透明度。

解析器会把该效果放入 shape 的 `computedStyle.effects`，包括缩放和对齐信息。Rust 现在生成独立的 alpha mask，按 `sx/sy` 变换阴影源，再用自定义高斯模糊、颜色和透明度合成 RGBA 位图，最后绘制正常填充和描边。Canvas 不再参与阴影算法，因此不会把缩放后的源边框错误地画成第二层可见边框。

### 9. 不同页面的内容底板误变红

第一页的内容底板使用 `solidFill(bg1)`，红色主要来自线条渐变；第二页的圆角底板则使用独立的 `gradFill`，并且源文件中的 gradient stop 没有按 offset 排序（`48000 -> 0 -> 100000`）。此外，页面样式中的 `fillRef idx="0"` 表示无填充，不能把它后面的 `phClr/accent1` 当成实心填充。

现在 TS 和 Rust 两层都会按 offset 排序 gradient stop；`fillRef idx=0` 解析为 `none`。这避免了第二页内容框被错误降级为 accent1 红色实心矩形，同时保留第一页和标题条的真实渐变。

### 10. 圆角几何过大

`roundRect` 不是固定圆角。PPTX 通过 `a:prstGeom/a:avLst/a:gd name="adj"` 保存每个形状的圆角调整值，不同内容框可以分别使用 `5598`、`2739`、`1933` 等参数。

虚拟 DOM 现在保留归一化后的 `cornerRadius=adj/100000`，Rust 的普通绘制和阴影 alpha mask 都使用该值；没有显式 `adj` 时才使用 OOXML 的默认比例。

### 11. 竖向箭头比例与文字镜像位置

第六页的红色竖向箭头是独立的 `a:prstGeom prst="upArrow"`，尺寸约为
`302260 x 4572635 EMU`，且 `a:avLst` 为空。空调整列表不表示可以按形状总高度
随意取百分比；OOXML 预设几何仍要使用短边 `ss=min(width,height)`。此前把箭头
头部高度设为总高度的固定比例，极高窄形状因此被画成长尖三角。Rust 现在以
`0.5 * ss` 计算默认箭头头部高度，以 `0.5 * width` 计算默认竖杆宽度；显式
`adj1/adj2` 仍由虚拟 DOM 传入。普通绘制和效果 alpha mask 共用相同参数。

同页文字使用 run 级 `a:reflection`，参数包含 `blurRad=6350`、`stA=55000`、
`endA=300`、`endPos=45500`、`dir=5400000`、`sy=-100000` 和 `algn=bl`。
反射作用于实际字形轮廓，不是包含 margin、leading 的整块文本行盒。Rust 现在先从
RGBA 文本位图提取非透明像素的垂直边界，再以实际 glyph 底边作为镜像轴；同时保留
`dist/dir` 位移。这样不会因为行盒底部的透明区域而把镜像整体向下推远。

## 核心亮点文本框的实际值

样例 PPTX 中“核心亮点”内容框约为：

- 几何尺寸：`545 x 288 px`
- 左右内边距：`14.4 px`
- 上下内边距：`7.2 px`
- 字号：`11.5 pt`，约 `23 px`
- 行距：`spcPct=120000`，即 `120%`，以实际字体 line metrics 为基准
- `anchor=t`
- 没有 `spAutoFit`、`normAutofit` 或 `noAutofit`

因此这个文本框不是 WPS 自动把文字拉伸填满，而是固定文本框配合 Office 的东亚换行、字体指标、标点悬挂和 bullet 缩进规则完成排版。

## Latest correction: grouped paragraph units

`marL` and `indent` are absolute EMU paragraph properties. They must use
`absoluteUnitScale` and must not use a group's cumulative geometry scale.
Using the group scale pushed grouped body text outside its bitmap while the
negative hanging indent left the bullet visible. Centered text anchoring also
keeps the natural signed remainder, so a slightly overfull line box is not
forced downward inside a header bar.

East Asian `hangingPunct` no longer contributes a half-character allowance to
the entire paragraph. Every visual line uses the XML text measure. The CJK
path now breaks at measured glyph/token boundaries and, when a closing
punctuation would be the line-ending character, keeps that punctuation on the
line without using its advance to evict the preceding glyph. `spAutoFit` still
uses the DrawingML default insets unless `lIns`/`rIns` are explicitly present.

The `spcPct` line-spacing path keeps the natural font alignment box and adds
the percentage delta based on paragraph text size. Font ascent/descent/leading
are therefore preserved without multiplying the entire metric box again. This
is important for mixed paragraphs such as a 13pt bold label followed by 12pt
body text: the previous full-metric multiplication made the fixed text box
taller than the WPS layout even when the visible glyph size looked identical.
The percentage increment is currently multiplied by a small WPS compatibility
factor (`0.8`) so the final baseline distance is slightly tighter than the
literal metric-plus-120% result while the XML `spcPct` value remains intact.

## 当前仍需完善

- `spAutoFit` 已支持按实际排版高度生成文本位图；TS 虚拟 DOM 中的 shape 几何和后续元素位置仍不会随自动增高同步重排。
- `normAutofit` 目前只做字号缩小循环，尚未完整复现 Office 的最小字号和迭代策略。
- 悬挂标点目前仍是视觉层处理，行尾闭标点保留原始 advance，尚未完全参与行宽计算。
- bullet 字体、bullet 大小、bullet 与正文基线的独立继承还需要继续补齐。
- `fontAlgn=baseline/auto` 的差异需要更多 WPS/PowerPoint 对比样本校准。
- 表格合并、四边线、单元格内边距和主题表格样式仍不完整。

## 验证方式

修改 Rust WASM 后由项目使用者运行：

```powershell
npm run build:wasm:debug
```

然后刷新页面并重新加载 PPTX，与 WPS 导出的截图对比。重点观察：

- 中文段落换行位置
- 核心亮点文本框的总高度
- bullet 与正文的首行关系
- 标点是否越过文本边界
- 标题和正文的字体基线及行间距
