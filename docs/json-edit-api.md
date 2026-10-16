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
| `setFill` | `target*`、`color*`（颜色或 `null`） | |
| `setLine` | `target*`、`color?`（可为 `null`）、`widthPt?`、`dash?`（`solid` `dash` `dot`） | 只改写出的字段 |
| `setAnchor` | `target*`、`anchor*`（`top` `middle` `bottom`） | 文字垂直对齐 |
| `align` | `targets*`（数组）、`to*`（`slide`）、`edge*`（`left` `hcenter` `right` `top` `vcenter` `bottom`） | 对齐到页面 |
| `arrange` | `target*`、`mode*`（`front` `forward` `backward` `back`） | 层级 |
| `duplicate` | `target*`、`offset?`（厘米，缺省 0.4）、`as?` | 同页创建副本 |
| `copyTo` | `target*`、`slide*`、`as?` | 复制到另一页（图片等关系一并处理） |
| `delete` | `target*` | |

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

同一份 schema 生成两样东西：

1. **JSON Schema**：校验入参，出错时把位置原样返回给模型。
2. **工具定义**：`apply_edits`、`get_outline`、`render_slide` 三个工具，直接用于 Claude 的 tool use。

## 现在能做 / 暂时做不了

| 能做（引擎已具备） | 暂时做不了 |
|---|---|
| 页面增删复制移动、新建带占位符 | 组合 / 取消组合 |
| 文本框、基本形状、直线箭头、图片 | 图表（解析与渲染都还没有） |
| 改文字（保留格式）、局部格式、段落对齐 | 表格单元格文字与行列 |
| 位置、大小、旋转、填充、边框、垂直对齐 | 渐变填充、阴影、透明度 |
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
测试页。代码在 `src/api/`：

| 文件 | 作用 |
|---|---|
| `scene.ts` | 把一页还原成选择窗格那样的形状树（组合内嵌套） |
| `outline.ts` | `outline()`：厘米单位的结构描述 |
| `resolve.ts` | `id` / `name` / `placeholder` / `text` / `$别名` / `@selection` 寻址，找不到或有歧义时返回候选 |
| `schema.ts` | JSON Schema（21 个 op）与校验，错误指到具体字段 |
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

1. 由 schema 生成 Claude 工具定义（`apply_edits`、`get_outline`、`render_slide`）。
2. 无头渲染与 `renderSlide(n)`。
3. 组合 / 取消组合、图表、表格进 schema。
