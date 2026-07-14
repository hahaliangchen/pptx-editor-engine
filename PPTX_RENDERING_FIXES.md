# PPTX 渲染问题记录

本文档记录当前 PPTX 解析、虚拟 DOM 和 Rust 排版渲染中已经处理的问题，以及仍然需要继续完善的部分。

## 总体分工

- TypeScript 读取 PPTX Open XML，建立页面元素、几何坐标、文本段落和样式来源关系。
- TypeScript 输出虚拟 DOM。这里的重点不是只保存最终颜色或字号，还要保留继承来源和排版语义。
- Rust 接收虚拟 DOM，使用后端提供的系统字体文件，通过 cosmic-text 做 shaping/layout，再用 swash 光栅化字形。
- Canvas 只负责合成最终位图，不再负责字体选择、字体测量或 `fillText` 渲染。

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
(ascent - descent + leading) * spcPct
```

当前针对 WPS 的渲染兼容层还会对字体 alignment box 使用 `0.92` 校准系数。这不是 PPTX 属性，只用于匹配 WPS 对本样例字体内部 leading 的处理。固定 `spcPts` 仍然直接按绝对点数换算。

### 6. 东亚排版细节

本次补充了以下 Open XML 段落字段：

- `eaLnBrk`：东亚字符换行
- `hangingPunct`：悬挂标点
- `fontAlgn`：字体在行框中的垂直对齐

Rust 对包含中文的段落使用 `Wrap::WordOrGlyph`，允许在合法字符边界回退换行，避免 `Wrap::Word` 导致中文长段落行数过少。

悬挂标点会对行首开标点和行尾闭标点做半字宽的视觉偏移。`fontAlgn` 已进入虚拟 DOM；`auto` 使用 cosmic-text 根据实际字体 ascent/descent 计算的 leading。显式 top/bottom 的独立基线规则还需要更多 WPS/PowerPoint 样本校准。

### 7. bullet 悬挂缩进

虚拟 DOM 会保留 `marL` 和 `indent`。Rust 将 bullet 单独排版，并把 bullet 的位置放在：

```text
body margin + paragraph margin + indent
```

正文从段落左边开始，非 bullet 段落的正缩进只作用于首行。

### 8. 标题条红色光晕

WPS 中“目标客户”“智能体核心能力”“核心亮点”和“业务痛点”的标题条周围有红色柔化光晕。原始 PPTX 使用的是形状 `spPr` 内的 `a:outerShdw`，不是额外图片或渐变；样例参数包括 `blurRad=63500`、`sx/sy=105%`、居中对齐以及主题色透明度。

解析器会把该效果放入 shape 的 `computedStyle.effects`，包括缩放和对齐信息。Rust 采用独立阴影 pass，先按 `sx/sy` 绘制带效果的形状，再绘制正常填充和描边；这样不会把 Office 的外发光稀释成单次 Canvas `shadowBlur` 的淡边。

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

## 当前仍需完善

- `spAutoFit` 的 shape 模式还没有完整实现动态改变 shape 高度。
- `normAutofit` 目前只做字号缩小循环，尚未完整复现 Office 的最小字号和迭代策略。
- 悬挂标点目前是视觉偏移，尚未完全参与行宽计算。
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
