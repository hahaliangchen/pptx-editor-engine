# 面向 AI 的 JSON 编辑接口（设计稿）

目标：AI 通过一份 JSON 指令修改 PPT，用户在界面里的操作和 AI 的指令走同一条路径
（同一套校验、同一个撤销栈）。AI 先"读"到当前结构，再发指令，再"看"到结果。

```
读   engine.outline()            当前每页有哪些形状、id、名字、文字、位置、样式
写   engine.apply(commands)      一批 JSON 指令，逐条校验，整批可撤销
看   engine.renderSlide(n)       该页的 PNG（给视觉模型自查）
```

## 约定

| 项 | 约定 | 原因 |
|---|---|---|
| 页码 | 从 1 开始（`slide: 2` 是第二页），也可用页 id `"slide_7"` | 对人和模型都自然 |
| 长度 | 厘米，可带小数（`x: 3.5`） | 与界面里「位置与大小」一致，模型对厘米有概念 |
| 角度 | 度，顺时针，0–360 | |
| 字号 | 磅（`sizePt`） | PPT 里的通用单位 |
| 颜色 | `"#RRGGBB"`；去掉填充/边框用 `null` | 不接受颜色名，避免歧义 |
| 位置 | 左上角为原点；`x` 向右，`y` 向下 | |
| 字段名 | 带单位或含义（`fontSizePt`、`widthPt`），不用 `size` 这类含糊名 | |

## 怎么指向一个形状（`target`）

按优先级，只写一种：

```json
{ "slide": 2, "id": "12" }                    // 形状 id，最稳定（文件里的 cNvPr id）
{ "slide": 2, "name": "标题 3" }               // 名字，同名时报错并列出候选
{ "slide": 2, "placeholder": "title" }        // 占位符类型：title / subTitle / body
{ "slide": 2, "text": "市场分析" }              // 文字包含该内容的形状，多个匹配时报错
"$hero"                                       // 本批中前面指令用 "as":"hero" 起的别名
"@selection"                                  // 用户在界面里选中的对象
```

找不到或有歧义时**不猜**，返回错误和候选（见"错误"），让模型改正后重试。

## 指令

每条都有 `op`。带 `*` 的是必填。

### 读取与页面

| op | 参数 | 说明 |
|---|---|---|
| `addSlide` | `after?`（页码，默认当前页之后）、`layout?`（版式名，见 outline.layouts） | 新建页，按版式生成标题/正文占位符 |
| `duplicateSlide` | `slide*`、`after?` | 复制页（含已做的修改） |
| `deleteSlide` | `slide*` | 至少保留一页 |
| `moveSlide` | `slide*`、`to*`（目标页码） | |

### 新增内容

| op | 参数 | 说明 |
|---|---|---|
| `addText` | `slide*`、`rect*`、`text*` 或 `paragraphs*`、`style?`、`as?` | 文本框；`paragraphs` 可带每段对齐和多段文字 |
| `addShape` | `slide*`、`shape*`、`rect*`、`fill?`、`line?`、`text?`、`as?` | `shape`：`rect` `roundRect` `ellipse` `triangle` `rightArrow` `line` `arrow` |
| `addImage` | `slide*`、`src*`（`data:` URL 或 base64）、`rect?`（缺省时居中适配）、`as?` | PNG / JPEG / GIF |

`rect` 为 `{ "x": 3, "y": 2, "w": 10, "h": 4 }`（厘米）。`as` 给新对象起别名，后面的指令用 `"$别名"` 引用。

### 表格

| op | 参数 | 说明 |
|---|---|---|
| `addTable` | `slide*`、`rows*`、`header?`、`columnWidths?`、`fontSizePt?`、`rect?`、`name?`、`as?` | 新建表格；`rows` 是按行排列的文字，第一行默认当表头 |
| `setTable` | `target*`、`rows*` | 整体替换内容：行数、列数不够就增加，多了就删除，再逐格写入 |
| `setCell` | `target*`、`row*`、`col*`，以及 `text`、`bold`、`italic`、`color`、`fontSizePt`、`align`、`fill`（至少一个） | 改一个单元格的文字或样式；`fill: null` 去掉底色 |
| `insertRow` | `target*`、`at*`、`text?` | 在第 `at` 行的位置插入新行（后面的行下移），`text` 是各列的文字 |
| `deleteRow` | `target*`、`row*` | 至少保留一行 |
| `insertColumn` | `target*`、`at*`、`text?` | 在第 `at` 列的位置插入新列；各列宽度按比例缩小，整张表宽度不变 |
| `deleteColumn` | `target*`、`col*` | 至少保留一列；其余列放宽，整张表宽度不变 |

- 行号、列号都从 1 开始，第 1 行就是表头。单元格里换行写 `\n`。
- 新建的表格把格式直接写进单元格（表头用主题强调色加白色粗体、浅色隔行底、白色分隔线），在 PowerPoint、WPS 里看起来一样，不依赖表格样式。
  在有隔行底色的表格里增删行后，底色会重新交替。
- 文稿里**已有**的表格同样可以改：只换文字，单元格原来的字体、颜色、边框、对齐都保留；新增的行或列沿用相邻单元格的格式。
  表格样式（`tableStyleId`，隔行色、表头色）会读取并画出来。
- 文字写不下时所在行会自动变高，整张表随之变高（和 PowerPoint 一样）；新增行、列后请看一眼渲染，别让表格超出页面。
- 有**合并单元格**的表格可以改文字（合并单元格本身），但不能增删行列，会返回 `NOT_APPLICABLE` 并说明原因；被合并盖住的格子不能写。
- 表格可以移动、缩放（`setBounds`）、调层级、复制（`duplicate`、`copyTo`、`duplicateSlide`）、删除；不能旋转，没有整体填充和边框可设。
- 目前只能通过这些指令编辑表格；画布上双击单元格直接输入还没有做。

### 图表

| op | 参数 | 说明 |
|---|---|---|
| `addChart` | `slide*`、`type*`、`categories*`、`series*`、`title?`、`legend?`、`dataLabels?`、`stacked?`、`rect?`、`name?`、`as?` | 新建图表 |
| `setChart` | `target*`，以及上面任意字段（至少一个）；`title: null` 删除标题 | 改图表的数据或设置；`categories` / `series` 整体替换 |

- `type`：`column`（竖柱）、`bar`（横条）、`line`（折线）、`pie`（饼图）。
- `categories`：坐标轴上每组的标签（饼图是每块的名称）；`series`：`[{ "name", "values", "color?" }]`，
  `values` 的个数必须等于 `categories` 的个数，缺的数据写 `null`（折线断开、柱子留空）。饼图只能有一个系列，且不能为负。
- `legend`：`top` `bottom` `left` `right` `none`，缺省为饼图 `right`、其余 `bottom`；`dataLabels` 缺省为饼图 `true`、其余 `false`。
- `textColor`、`fontSizePt`（坐标轴、图例、数值的字号，默认 12）、`titleSizePt`（标题，默认 18）控制整张图的文字；
  不写 `textColor` 时跟随页面底色：浅色底用深灰，深色底用近白。已有图表也能改，只有真正改了的文字才会被重写。
- 没写 `color` 的系列用文稿主题的强调色（accent1…accent6）依次取色；颜色不够时在同一色上变深或变浅。
- 新建的图表带一个内嵌的 Excel 工作簿，在 PowerPoint 里"编辑数据"可以直接打开。
- 文稿里**已有**的图表同样可以用 `setChart` 改数据：只改数据、标题、图例、数据标签，原有的样式、字体、坐标轴格式保持不变；
  改 `type` 或 `stacked` 时整张图重新生成，自定义样式会丢失。
- 图表可以移动、缩放（`setBounds`）、调层级、删除、复制（`duplicate`、`copyTo`、`duplicateSlide`，复制出的图表有自己独立的数据）；
  不能旋转，也没有填充和边框可设（`setFill`、`setLine`、`setRotation` 返回 `NOT_APPLICABLE`）。
- 暂不支持的图表（组合图、面积图、散点图、雷达图、圆环图、百分比堆积等）读出来是 `chart.unsupported`，画面上显示占位框，
  位置和大小仍可编辑，数据不能改。

### 修改文字

| op | 参数 | 说明 |
|---|---|---|
| `setText` | `target*`、`text*` | 用新文字替换全部文字，**保留原来的字号、颜色、加粗等格式** |
| `replaceText` | `target*`、`find*`、`replace*`、`all?` | 只替换匹配到的部分，其余格式不动 |
| `formatText` | `target*`、`range?`（`{from,to}` 字符位置，缺省为全部）、`bold?` `italic?` `color?` `sizePt?` | 部分文字的格式 |
| `setParagraph` | `target*`、`align*`（`left` `center` `right`） | 段落对齐 |

### 修改形状

| op | 参数 | 说明 |
|---|---|---|
| `setBounds` | `target*`、`x?` `y?` `w?` `h?`、`keepRatio?` | 位置和大小，未写的字段不变 |
| `setRotation` | `target*`、`degrees*` | |
| `setFill` | `target*`，以及 `color`（颜色或 `null`）、`opacity`、`gradient`（至少一个，`color` 与 `gradient` 不能同时写） | 填充；`opacity` 单独写时保留现有颜色或渐变，只改透明度 |
| `setShadow` | `target*`、`shadow*`（对象或 `null`） | 投影；图片也可以；`null` 去掉阴影 |
| `setLine` | `target*`、`color?`（可为 `null`）、`widthPt?`、`dash?`（`solid` `dash` `dot`） | 只改写出的字段 |
| `setAnchor` | `target*`、`anchor*`（`top` `middle` `bottom`） | 文字垂直对齐 |
| `align` | `targets*`（数组）、`to*`（`slide`）、`edge*`（`left` `hcenter` `right` `top` `vcenter` `bottom`） | 对齐到页面 |
| `arrange` | `target*`、`mode*`（`front` `forward` `backward` `back`） | 层级 |
| `duplicate` | `target*`、`offset?`（厘米，缺省 0.4）、`as?` | 同页创建副本 |
| `copyTo` | `target*`、`slide*`、`as?` | 复制到另一页（图片等关系一并处理） |
| `delete` | `target*` | |

#### 页面背景

| op | 参数 | 说明 |
|---|---|---|
| `setBackground` | `slide*`，以及 `color` 或 `gradient`（二选一） | 给这一页设背景；`gradient` 的写法和 `setFill` 一样 |

- `outline()` 的每一页有 `background`（颜色或渐变）：页面自己的，没有就是版式的、母版的。深色底上的文字、图表、表格要用浅色。
- 新建的文本框默认用页面"通常的文字色"：浅色底上是黑色，深色底上是白色；图表的文字同理（深灰 / 浅灰）。
- 深色模板（母版把 `bg1` 映射到深色、`tx1` 映射到浅色）现在能正确画出：黑底白字，和 PowerPoint 一致。

#### 渐变、透明度、阴影

```json
{ "op": "setFill", "target": { "slide": 1, "name": "卡片" },
  "gradient": { "kind": "linear", "angle": 90,
                "stops": [{ "position": 0, "color": "#1F4E9C" }, { "position": 100, "color": "#5EC8F0" }] } }
{ "op": "setFill", "target": { "slide": 1, "name": "卡片" }, "opacity": 0.5 }
{ "op": "setShadow", "target": { "slide": 1, "name": "卡片" }, "shadow": { "blurPt": 10, "distancePt": 6, "opacity": 0.4 } }
```

- `gradient.kind`：`linear`（默认）或 `radial`（从中心向外）。`angle` 是度数，和 PowerPoint 里一致：0 从左到右，90 从上到下，缺省 90。
  `stops` 至少 2 个，最多 8 个，`position` 是 0–100 的百分比，每个停点可以单独带 `opacity`。
- `opacity`：0 完全透明，1 不透明。和 `color` 一起写就是"这个颜色、这个透明度"；单独写只改现有填充（纯色或渐变）的透明度，
  没有填充的形状会返回 `NOT_APPLICABLE`。
- `shadow`：`color`（默认 `#000000`）、`opacity`（0.35）、`blurPt` 柔和程度（6）、`distancePt` 距离（4）、`angle`（45，0 向右、90 向下），
  都可以省略。组合、表格、图表不能设阴影。
- `outline()` 读出来时：渐变是 `fill: { gradient: { type, angle, stops } }`，透明的纯色有 `opacity`，阴影是 `shadow: { color, opacity, blurPt, distancePt, angle }`。
- `addShape` 也可以一步写全：`fill` 可以是颜色、`null` 或 `{ "gradient": {...} }`（就是 `outline()` 读出来的样子，可以原样写回去），
  还可以带 `opacity`、`line`（`null` 表示无边框）、`shadow`；`addImage` 可以带 `shadow`。`gradient` 里 `kind` 和 `type` 两种写法都认。
- 文稿里已有的渐变、半透明、阴影都能读出和画出；改的时候只动被指定的那一项，形状的其他设置保持原样。
  形状的填充如果来自主题样式（文件里没有写明），改透明度时会把当前画出来的填充写成明确的颜色。

## 批处理与撤销

```json
{
  "label": "把标题改成红色并加一张图",
  "atomic": true,
  "commands": [
    { "op": "setText", "target": { "slide": 1, "placeholder": "title" }, "text": "2025 年市场分析" },
    { "op": "setFill", "target": { "slide": 1, "name": "背景条" }, "color": "#E60012" },
    { "op": "addImage", "slide": 1, "src": "data:image/png;base64,...", "as": "logo" },
    { "op": "setBounds", "target": "$logo", "x": 28, "y": 1, "w": 4 }
  ]
}
```

- `atomic: true`（默认）：任何一条失败，**整批都不生效**，文档保持原样。
- 成功后整批在历史记录里是**一步**，用户一次撤销就回到修改前；`label` 作为这一步的名字。
- `atomic: false`：失败的跳过，其余照常执行，结果里逐条说明。

## 返回

```json
{
  "ok": true,
  "results": [
    { "ok": true },
    { "ok": true },
    { "ok": true, "created": { "id": "45", "name": "图片 45", "slide": 1 } },
    { "ok": true }
  ]
}
```

## 错误

错误必须让模型**能直接据此改正**：

```json
{
  "ok": false,
  "failedAt": 2,
  "error": {
    "code": "SHAPE_NOT_FOUND",
    "message": "第 1 页没有名为 \"背景条\" 的形状",
    "path": "commands[1].target",
    "candidates": [
      { "id": "4", "name": "矩形 3", "text": "" },
      { "id": "5", "name": "标题 4", "text": "市场分析" }
    ]
  }
}
```

| code | 含义 |
|---|---|
| `INVALID_SCHEMA` | 字段缺失、类型或取值不对（附 JSON Schema 的错误位置） |
| `SLIDE_NOT_FOUND` | 页码超出范围（附总页数） |
| `SHAPE_NOT_FOUND` | 找不到形状（附同页候选） |
| `SHAPE_AMBIGUOUS` | 名字或文字匹配到多个（附全部候选） |
| `NOT_APPLICABLE` | 指令对该对象无意义（例如给图片设文字对齐） |
| `UNSUPPORTED` | 暂不支持（例如含图表的页面无法复制） |

## 读取：`outline()`

```json
{
  "slideSize": { "w": 33.87, "h": 19.05 },
  "layouts": ["标题幻灯片", "标题和内容", "仅标题", "空白"],
  "slides": [
    {
      "index": 1, "id": "slide_1", "layout": "标题和内容",
      "shapes": [
        { "id": "2", "name": "标题 1", "type": "text", "placeholder": "title",
          "rect": { "x": 2.1, "y": 1.2, "w": 29.6, "h": 3.2 }, "rotation": 0,
          "text": "市场分析",
          "style": { "fontSizePt": 32, "bold": true, "color": "#000000", "align": "center" } },
        { "id": "5", "name": "矩形 4", "type": "shape", "shape": "roundRect",
          "rect": { "x": 3, "y": 6, "w": 10, "h": 4 },
          "fill": "#4472C4", "line": { "color": "#2F528F", "widthPt": 1, "dash": "solid" },
          "text": "增长 30%" },
        { "id": "7", "name": "图片 6", "type": "image", "rect": { "x": 18, "y": 6, "w": 12, "h": 7 } },
        { "id": "10", "name": "表格 9", "type": "table", "rect": { "x": 2, "y": 8, "w": 20, "h": 6 },
          "table": { "rows": [["产品", "销量"], ["甲", "120"]], "columns": 2, "columnWidths": [10, 10], "rowHeights": [1.5, 1.5] } },
        { "id": "8", "name": "图表 7", "type": "chart", "rect": { "x": 1, "y": 8, "w": 15, "h": 9 },
          "chart": { "type": "column", "title": "季度营收", "categories": ["一季度", "二季度"],
                     "series": [{ "name": "华东", "values": [120, 150] }], "legend": "bottom",
                     "dataLabels": false, "stacked": false } },
        { "id": "9", "name": "组合 8", "type": "group", "rect": { "x": 1, "y": 14, "w": 30, "h": 4 },
          "children": [ ... ] }
      ]
    }
  ]
}
```

- 按需读取：`outline({ slide: 3 })` 只返回一页，避免大文稿一次塞满上下文。
- 文字很长时截断并标注 `"textTruncated": true`，需要完整内容时按 id 单独读。

## 看：`renderSlide(n)`

返回该页的 PNG（data URL），给视觉模型检查文字溢出、重叠、对齐。需要先有"无头渲染"，
见下方实现顺序。

## 工具定义（给模型用）

三个工具，定义写在 `src/api/tools.ts`，与任何厂商无关：名称、说明、输入的 JSON Schema，加一个执行函数。

| 工具 | 输入 | 返回 |
|---|---|---|
| `get_outline` | `slides?`（页码数组）、`maxTextLength?` | 文稿结构的 JSON（见上文 `outline()`） |
| `apply_edits` | 指令批（本文档上面的全部指令） | 执行结果 JSON；失败时 `ok: false`，带失败位置、`error.code`、`message`、`path` 和候选 |
| `render_slide` | `slide`（页码）、`width?`（默认 1024） | 一张 PNG 图片 + 一句说明 |

`callTool(host, name, args)` 执行一个工具，**从不抛异常**：参数不对或编辑失败都以 `ok: false` 的结果返回，
文字里写明哪里错了，模型可以据此重试。`host` 只要有 `outline`、`apply`，要能出图再加 `render`；
`HeadlessSession` 三样都有，可以直接传。

`TOOL_GUIDE` 是一段给系统提示用的文字：先读结构、整页一批提交、用别名引用新对象、改完渲染检查、
失败时读候选再改。把它放进系统提示，模型的用法会稳得多。

### 接入各家（`src/api/adapters.ts`）

三家的差别只在字段名和包装，适配层各有三样：**工具列表**、**工具结果的转换**、**一个"收到调用就执行并回复"的函数**。

```ts
import { HeadlessSession, TOOL_GUIDE, toAnthropic, callAnthropicTool } from "./src/headless";

const deck = await HeadlessSession.open(pptxBytes);

// Anthropic：tool_use 块进，tool_result 块出（图片会作为 image 块放进去）
const response = await client.messages.create({
  model, system: TOOL_GUIDE, tools: toAnthropic(), messages
});
const results = await Promise.all(
  response.content.filter(block => block.type === "tool_use").map(block => callAnthropicTool(deck, block))
);
messages.push({ role: "assistant", content: response.content }, { role: "user", content: results });
```

```ts
// OpenAI：参数是 JSON 字符串；工具消息只能放文字，图片会放进紧跟其后的 user 消息里
const reply = await openai.chat.completions.create({ model, tools: toOpenAI(), messages });
for (const call of reply.choices[0].message.tool_calls ?? []) {
  messages.push(...await callOpenAITool(deck, call));
}
```

```ts
// MCP：tools/list 返回 toMcp()；tools/call 的处理函数就是 callMcpTool
server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: toMcp() }));
server.setRequestHandler(CallToolRequestSchema, async req => callMcpTool(deck, req.params.name, req.params.arguments));
```

- 工具里的 `apply_edits` 用的 schema 把重复出现的片段（`target`、`rect`、`color` 等）提成 `$defs` 引用，
  体积从约 28 KB 降到约 13 KB（每次请求都要带，越小越省）；含义与校验用的完整 schema 完全相同（有测试逐条对照）。
- 适配层去掉了 `$schema` 字段（部分厂商不接受），每次返回的都是独立副本。
- 不提供 OpenAI 的 `strict` 模式：编辑 schema 用了 `oneOf`，strict 模式不支持。
- 命令行里可以看定义和直接试：`npm run headless -- tools anthropic|openai|mcp|guide`，
  `npm run headless -- call 文稿.pptx apply_edits '{"commands":[...]}'`。

## 现在能做 / 暂时做不了

| 能做（引擎已具备） | 暂时做不了 |
|---|---|
| 页面增删复制移动、新建带占位符 | 组合 / 取消组合 |
| 表格：读取（含表格样式）、渲染、新建、改单元格、增删行列、复制、导出 | 合并 / 拆分单元格、画布上直接编辑单元格 |
| 文本框、基本形状、直线箭头、图片 | 组合图、面积图、散点图等其他图表类型 |
| 柱形图、条形图、折线图、饼图：读取、渲染、新建、改数据、导出 | 图表的自定义样式编辑、坐标轴设置 |
| 改文字（保留格式）、局部格式、段落对齐 | 表格的边框与列宽单独设置 |
| 位置、大小、旋转、填充（纯色、渐变、透明度）、边框、阴影、页面背景、垂直对齐 | 发光、柔化边缘、倒影的编辑；边框的渐变与透明度 |
| 对齐到页面、层级、副本、跨页复制 | 动画、切换效果、备注页 |

## 实现顺序（建议）

1. **寻址与读取**：形状的稳定寻址（id / 名字 / 占位符 / 文字），`outline()`。
2. **调度与校验**：JSON Schema、`apply()`、批处理与整批撤销、结构化错误。
   现有编辑方法大多按"当前选中"工作，第一版用"先按 target 选中、再调用现有方法"的方式桥接，
   之后再逐步改成直接接收目标。
3. **测试页**：左边贴 JSON，点执行，右边看渲染和返回；加上用例测试（每个 op 一个，加批处理和错误）。
4. **工具定义**：由 schema 生成 Claude tool 定义。
5. **无头渲染**：命令行输出 PNG，接上 `renderSlide`，让 AI 能看到结果。
6. 之后按需扩 schema：组合、图表、表格。

## 实现状态（第一版）

已实现并有测试：寻址、`outline()`、`apply()`、JSON Schema 校验、整批撤销与回滚、结构化错误、
测试页、工具定义与三家适配。代码在 `src/api/`：

| 文件 | 作用 |
|---|---|
| `scene.ts` | 把一页还原成选择窗格那样的形状树（组合内嵌套） |
| `outline.ts` | `outline()`：厘米单位的结构描述 |
| `resolve.ts` | `id` / `name` / `placeholder` / `text` / `$别名` / `@selection` 寻址，找不到或有歧义时返回候选 |
| `schema.ts` | JSON Schema（32 个 op）与校验，错误指到具体字段 |
| `charts.ts` | `addChart` / `setChart` 的字段合并与校验 |
| `tools.ts` | 三个工具的定义、`TOOL_GUIDE`、`callTool`（与厂商无关） |
| `adapters.ts` | Anthropic / OpenAI / MCP 的工具列表、结果转换与"收到调用就执行"函数 |
| `engine.ts` | `EditApi`：`outline()` 与 `apply()`，走与鼠标键盘相同的编辑操作 |
| `playground.ts` | 页面右上角「AI 接口」抽屉：贴 JSON、执行、看结果与结构 |

### 怎么试

- **页面**：打开文稿后点右上角 **{ } AI 接口**，选示例或贴自己的 JSON，`Ctrl+Enter` 执行。
  执行后能直接在画布上看到结果，也可以用撤销或历史记录回到修改前。
- **控制台**：`await pptEngine.outline({ slides: [1] })`、`await pptEngine.apply([...])`。
- **导出后校验**：`npm run validate:pptx -- 导出的.pptx 原始.pptx`。

### 与设计稿的差异

- `setLine.dash` 的取值是 `solid` / `dash` / `dot`（内部对应 OOXML 的 `sysDot`）。
- `duplicate.offset` 与 `addSlide.layout` 已按设计实现；`addSlide` 的 `layout` 取自 `outline().layouts`。
- `align` 目前只支持对齐到页面（`to: "slide"`），多个 `targets` 逐个对齐。
- `addImage` 只接受 `data:` URL（PNG / JPEG / GIF）。
- 颜色在 `outline()` 里统一为小写。

### 已知限制

对没有文字框的形状（例如无文字的矩形）暂时不能用 `setText`；组合只能整体移动、对齐、层级，
不能旋转。还没有的能力见上面的"暂时做不了"表。

### 下一步

1. ~~工具定义~~：已完成，见上文"工具定义"。
2. ~~无头渲染与 `renderSlide(n)`~~：已完成（`HeadlessSession.render`）。
3. ~~图表、表格进 schema~~：已完成；组合 / 取消组合还没有。

## 用真实模型试过的结果

`npm run agent` 让任何 OpenAI 兼容的模型通过这三个工具做事，并逐轮记录（工具调用、渲染图、token）：

```sh
npm run agent -- --base http://主机:端口/v1 --model 模型名 --task "做一份……" [--deck 底稿.pptx] [--out 目录] [--no-think]
```

默认底稿是 `tools/templates/blank.pptx`（16:9，标题 / 标题和内容 / 仅标题 / 空白四个版式，用 `npx vite-node tools/make-template.ts` 生成）。
在 vLLM 上的 Qwen3.8-27B 上试过六种任务，没有一次要人介入：

| 任务 | 轮数 | 编辑失败 | 备注 |
|---|---|---|---|
| 4 页汇报：封面、要点、柱形图、表格 | 4 | 0 | 一次提交 11 条指令，24 秒 |
| 改 31 页模板：留 3 页，换图表数据、填议程表、改说明文字 | 7–8 | 0 | 大文稿先看概要后，输入 token 从 28.6 万降到 9 万 |
| 6 页发布会：渐变阴影卡片、对比表、折线图、饼图、数字卡片 | 10 | 0 | 设计感、版式都过关 |
| 日历页里 6 个同名色块、只给 5 个名称 | 5 | 0 | 主动发现数量对不上并询问，没有乱填 |
| 3 页深色风格产品介绍 | 10 | 0 | 17 秒 |
| 关掉思考（`--no-think`）的 6 页发布会 | 7–18 | 1–3 | 失败都是写了接口没有的字段，补上之后降到 1 次 |

试出来、并已经修掉的问题：

- 完整结构太大（31 页 104 KB）：页数多时先给概要，再按页读（`get_outline` 的 `detail`）。
- 模型会把 `outline()` 读出的格式原样写回去、会在 `addShape` 里直接写 `shadow` 和 `line: null`：接口现在都接受。
- 模型想调图表文字的颜色和大小，没有字段，于是把同一条指令重复了 8 次：加了 `textColor`、`fontSizePt`、`titleSizePt`，
  使用指南里也写明"做不到就换办法或如实说明，不要重复同样的指令"。
- 深色模板被画成白底黑字（页面背景和母版的颜色映射没有处理）：已修，并加了 `setBackground`。
- 报"不认识的字段"时只说字段名：现在会列出这条指令能用的字段。

模型的版面判断仍有小毛病：卡片偏高、下半部分留白，表格有时不居中。这些靠 `render_slide` 的自检发现不了，
属于需要在使用指南或模板里给更具体规则的地方。

## 真实 PPTX 能直接当模板吗：试过了，可以

ai-ppt-server 的页面模板是 SVG（`resource/ai_skills/ppt-master/templates/layouts/<模板名>/`：封面、目录、章节、内容、结尾 5 张 SVG，
加一份描述页面类型、占位符、配色的 `design_spec.md`）。把其中的"科技蓝商务"转成真实 PPTX，再让 Qwen 用它做 5 页：

```sh
tools/svg-template-to-pptx.sh 科技蓝商务          # 输出 tools/templates/imported/科技蓝商务/{template.pptx,design_spec.md}
npm run agent -- --base http://主机:端口/v1 --model 模型名 \
  --deck tools/templates/imported/科技蓝商务/template.pptx --spec tools/templates/imported/科技蓝商务/design_spec.md \
  --task "用这个模板做一份 5 页的……"
```

- 转换用的是 ai-ppt-server 自己的 `svg_to_pptx.py`（只读使用，那个项目里什么都没改）；输出的 PPTX 是原生形状，不是图片。
  `tools/templates/imported/` 已加入 `.gitignore`：那是别的项目的模板内容，不随本仓库发布。
- 模板描述 `design_spec.md` 原样保留，放进系统提示（`--spec`）；PPTX 里的 `{{TITLE}}` 一类占位符就是普通文字，
  模型用 `setText` / `replaceText` 或按文字找到后替换。
- 结果：两个不同主题各做了一次，都是 5 页、没有残留任何 `{{ }}`、模板的背景、波浪、六边形、版式和配色完整保留；
  模型自己复制内容页、删掉不需要的章节页和多余的目录项、按页数改页码。第一次 11 轮 204 秒，没有编辑失败；
  第二次 16 轮 385 秒，2 次编辑失败（同一条指令指向了不存在的页脚文字），之后自己改正。开着思考比较慢，关掉（`--no-think`）快得多。

这次试验顺带暴露并修好的渲染问题（对所有文稿都有效）：

- 文字颜色带透明度时被画成不透明（章节页的大水印字）：cosmic-text 取字形像素时丢了颜色的 alpha，现在在 Rust 里乘回去。
- 字体在服务器上只以别名存在（"微软雅黑"→ Noto Sans CJK，"Calibri"→ Carlito）时，按名字查字体库查不到，行高退化成一个字号：
  现在退回到合适的后备字体（中文用 CJK 字体）的度量，行距和 PowerPoint、LibreOffice 一致。
- 主题没有写明中文字体时，默认用微软雅黑（PowerPoint 简体中文环境的行为）。

## 配置流：模型只写内容，整份一次生成

边读边改的循环（`get_outline` → `apply_edits` → `render_slide`）灵活，但每一步都要模型想一遍，5 页要 6 分多钟。
有模板时可以换一种做法：模型一次输出整份内容的 JSON，引擎负责版面。

模板用 `{{TOKEN}}` 标出要填的位置（ai-ppt-server 的 25 套模板都是这个约定）。`inferTemplate`（`src/api/template.ts`）读出：

| 页面类型 | 认出它的占位符 | 字段 |
| --- | --- | --- |
| `cover` | `{{TITLE}}` | title、subtitle、author、date … |
| `toc` | `{{TOC_ITEM_n_TITLE}}` | `items`: 章节名，或 `{title, desc}` |
| `section` | `{{CHAPTER_*}}` | title、desc；章节号自动编 |
| `content` | `{{CONTENT_AREA}}` | title、`blocks`；放内容的位置取占位符外面的虚线框 |
| `end` | `{{THANK_YOU}}` | title、subtitle、message、contact … |

字段名由占位符转来（`CLOSING_MESSAGE`→`message`，`STAT_1_LABEL`→`stat1Label`）；页码和总页数自动填。
没填的字段，那个文本框会删掉；模板里的示例文字（品牌名、口号）用 `replace` 换掉或删掉。

```json
{"slides":[
  {"page":"cover","title":"2025 新能源汽车市场分析","subtitle":"…","author":"市场研究部"},
  {"page":"toc","items":["市场规模",{"title":"竞争格局","desc":"头部品牌"}]},
  {"page":"content","title":"核心指标","blocks":[
    {"type":"cards","items":[{"value":"1290万","title":"年销量","text":"同比 +35%"}]},
    {"type":"chart","kind":"column","categories":["Q1","Q2"],"series":[{"name":"销量","values":[1,2]}]}]},
  {"page":"end","title":"谢谢"}
]}
```

内容块：`text`、`bullets`、`cards`、`chart`（图是原生图表，系列色取模板主色）、`table`；
最多 3 块，2 块默认左右并排（带卡片则上下），`layout: "rows"|"columns"` 可改。

`compileDeck`（`src/api/deck-spec.ts`）把它编译成一批普通指令（复制模板页、`setText`、`addChart` ……），
所以原子性、撤销、导出都和手工编辑一致；出错时给出字段路径（`slides[1].titel`）。引擎做的版面：

- 文字比模板的文本框长：先向对齐方向加宽（最多 1.8 倍），不够再缩小字号（最多缩到 70%），再不够用到页边距，最后继续缩小；
- 卡片、图表、表格按内容区均分；字号按内容量估算后选最大的合适值；
- 目录项少于模板的项数时，删掉多余项的全部形状，剩下的在原区域内居中。

用法：工具 `get_template`（看模板有哪些页面和字段）、`build_deck`（生成；文稿里的模板页用完即删，只能调一次，之后用 `apply_edits` 精修）。
联调脚本的一次出整份模式：

```sh
npm run agent -- --base http://主机:端口/v1 --model 模型名 --mode oneshot \
  --deck tools/templates/imported/科技蓝商务/template.pptx --spec tools/templates/imported/科技蓝商务/design_spec.md --task "……"
```

同一个任务（科技蓝商务模板，5 页的《2025 年新能源汽车市场分析》，Qwen3.8-27B）：

| 方式 | 用时 | 对话次数 | 输出 token |
| --- | --- | --- | --- |
| 工具循环（开思考） | 385 秒 | 16 | 28,645 |
| 配置流（开思考） | 47 秒 | 1 | 5,008 |
| 配置流（关思考） | 8 秒 | 1 | 928 |

两种做法并存：配置流出初稿，需要时再用 `render_slide` 和 `apply_edits` 点对点修。
目前配置流的短板：内容页下方常有留白（块高度按区域均分，不按内容量）。
