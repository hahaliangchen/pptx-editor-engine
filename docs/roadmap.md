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
| A2 | 抽出不依赖 DOM 的文档核心：解析、改动记录、页面管理（新建/复制/删除/移动、版式）、插图、导出、`EditApi`，现在这些写在 `PptViewer` 里，与画布和 DOM 绑定。抽出后浏览器视图和 Node 共用同一份逻辑 | 下一步 |
| A3 | `HeadlessSession`：`open(buffer)` → `outline()` / `apply()` → `render(page)` → `export()`；附命令行，便于调试和回归 | A2 之后 |

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

1. 表格（读、渲染、改单元格、增删行列）
2. 图表（柱/折线/饼，数据可改）
3. 渐变填充、阴影、透明度（`setFill` / `setEffect`）
4. 组合 / 取消组合；多选对齐与分布
5. 自定义形状（`custGeom`）与图片裁剪
6. 页码、页脚等版式字段的渲染（目前缺失：LibreOffice 会显示页码，我们不显示）
7. 按占位符批量填充模板（`{{TITLE}}` 一类），面向模板功能
8. 动画（后期）

## 阶段 C：工具与接入

1. 中立的工具核心 `tools.ts`：`get_outline`、`apply_edits`、`render_slide` 的名字、说明、输入 schema，
   以及 `callTool(name, args)`；再加薄适配层输出 Anthropic / OpenAI / MCP 格式。
2. 注册到现有的远程插件协议，与 SVG 工具并存，对同一需求对比效果后再切换。
3. 版本控制：用户手改与 Agent 修改并存时，沿用插件已有的 revision 思路，避免互相覆盖。
4. 前端：用本引擎的编辑器替换只读 SVG 预览；模板制作页改为编辑真实 PPTX 版式。

## 已知风险

- **字体度量差异**：服务器上多半没有微软雅黑、Aptos，回退到 Noto / Liberation 后字宽不同，Agent 在渲染图里看不到的
  溢出，可能出现在用户的 PowerPoint 里。对策：服务器装度量兼容的字体，排版留余量，并在工具说明里提示。
- **渲染与 PowerPoint 的像素差异**始终存在，靠 Open XML 校验保证文件合法，靠渲染保证大体正确，不承诺逐像素一致。
- A2 的重构动到现有编辑器的核心，必须靠现有测试和浏览器回归兜底，分小步提交。
