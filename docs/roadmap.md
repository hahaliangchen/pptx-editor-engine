# 路线图：从编辑器到 AI-PPT 底座

目标场景：用户在网页里与宿主 Agent 对话，Agent 通过工具生成并修改**真实 PPTX**；用户在同一份文件上
继续手动编辑、下载。宿主 Agent 在独立服务里调用 PPT 服务的工具（`ppt/call {tool, arguments}`），
所以**工具必须能在服务端执行**，不能依赖用户的浏览器。

```
宿主 Agent ──工具调用──▶ PPT 服务 ──▶ 本引擎（Node，无头）
                                        读写 PPTX、应用操作、渲染 PNG 供 Agent 自检
前端页面：同一个引擎加载同一份 PPTX，用户直接编辑、下载
```

## 阶段 A：引擎在 Node 里完整可用

| 步骤 | 内容 | 状态 |
|---|---|---|
| A1 | 无头渲染：`HeadlessRenderer.renderSlide(doc, index, {width})` → PNG。同一份 Rust 渲染代码，Skia canvas 代替浏览器 canvas | **完成** |
| A2 | 抽出 `DocumentSession`（`src/document-session.ts`）：解析、改动记录、页面管理（新建/复制/删除/移动、版式）、插图、导出、`EditApi`。显示相关的事通过 `SessionView` 回调交给视图，浏览器的 `PptViewer` 与 Node 共用同一份逻辑。`SlideEditor` 仍带 DOM 覆盖层，Node 里运行在 jsdom 上 | **完成** |
| A3 | `HeadlessSession`（`src/headless/session.ts`）：`open(buffer)` → `outline()` / `apply()` → `render(page)` → `export()`；命令行 `npm run headless` | **完成** |

A3 的使用方式（需要先 `npm run build:wasm:node`）：

```ts
const session = await HeadlessSession.open(readFileSync("deck.pptx"));
await session.apply([{ op: "addSlide", after: 1, layout: "空白" }, ...]);
writeFileSync("page2.png", await session.render(2));
writeFileSync("out.pptx", await session.export());
session.dispose();
```

```sh
npm run headless -- outline deck.pptx [页码]
npm run headless -- render  deck.pptx 2 page2.png [宽度]
npm run headless -- apply   deck.pptx commands.json out.pptx --render 输出目录
```

A1 的使用方式：

```ts
import { HeadlessRenderer } from "./src/headless/renderer";
const renderer = new HeadlessRenderer();   // 需要先 npm run build:wasm:node
renderer.prepare(doc);                     // 注册文稿用到的字体
const png = await renderer.renderSlide(doc, 0, { width: 1280 });
```

字体通过 fontconfig 查找（`fc-match`），缺失的字体由 fontconfig 回退。

## 阶段 B：补齐 PPTX 能力

每一项都必须同时做到：渲染、编辑、导出、进入 JSON schema，并通过 Open XML SDK 校验。
顺序按 AI 生成内容页的价值排：

1. ~~图表~~（**完成**：柱形、条形、折线、饼图的读取、渲染、新建、改数据、复制、导出，内嵌工作簿，已通过 Open XML 校验；
   下一批图表类型见下）
2. ~~表格~~（**完成**：读取含表格样式、渲染、新建、改单元格、增删行列、复制、导出；还缺合并/拆分单元格和画布上直接编辑单元格）
3. ~~渐变填充、阴影、透明度~~（**完成**：`setFill` 的 `gradient`、`opacity`，新增 `setShadow`；读取、渲染、编辑、导出都通过校验；发光、倒影、柔化边缘仍只读）
4. 组合 / 取消组合；多选对齐与分布
5. 自定义形状（`custGeom`）与图片裁剪
5b. 更多图表：组合图（柱 + 折线、双坐标轴）、面积图、圆环图、散点图、百分比堆积；坐标轴标题与数值格式；
   图表元素的手动布局（目前自动布局，不读取文件里的手动位置）
6. 页码、页脚等版式字段的渲染（目前缺失：LibreOffice 会显示页码，我们不显示）
7. ~~按占位符批量填充模板（`{{TITLE}}` 一类）~~（**完成**，见 json-edit-api.md 的"配置流"）
8. 动画（后期）

## 阶段 C：工具与接入

1. ~~中立的工具核心~~（**完成**）：`src/api/tools.ts`（三个工具的定义与 `callTool`）、`src/api/adapters.ts`
   （Anthropic / OpenAI / MCP 的工具列表、结果转换和"收到调用就执行"的函数）；`TOOL_GUIDE` 是给系统提示用的使用指南。
1b. 配置流（**完成**）：`build_deck` / `get_template` 与联调脚本的 `--mode oneshot`，整份一次生成。
    版面检查（**完成**）：`check_layout` 量出溢出、越界、重叠、留白、字太小、图表太小；一次出整份之后只把有问题的页交给模型修。
2. 注册到现有的远程插件协议，与 SVG 工具并存，对同一需求对比效果后再切换。
3. 版本控制：用户手改与 Agent 修改并存时，沿用插件已有的 revision 思路，避免互相覆盖。
4. 前端：用本引擎的编辑器替换只读 SVG 预览；模板制作页改为编辑真实 PPTX 版式。

## 已知风险

- **字体度量差异**：服务器上多半没有微软雅黑、Aptos，回退到 Noto / Liberation 后字宽不同，Agent 在渲染图里看不到的
  溢出，可能出现在用户的 PowerPoint 里。对策：服务器装度量兼容的字体，排版留余量，并在工具说明里提示。
- **渲染与 PowerPoint 的像素差异**始终存在，靠 Open XML 校验保证文件合法，靠渲染保证大体正确，不承诺逐像素一致。
- 每个 `HeadlessSession` 在 jsdom 全局上挂一个编辑器，用完要 `dispose()`；jsdom 的启动和内存开销在服务端需要留意，高并发时应复用进程内的会话而不是每次请求新开。
