# 页面工具：Agent 在用户打开的页面里改稿

用户在网页里打开一份稿子，宿主 Agent 的工具调用**直接在这个页面里执行**：改动立刻显示在画布上，用户随时可以接手编辑。
页面把稿子存回插件自己的后端，后端只负责存储、版本、模板、字体和缩略图。

```
宿主 Agent ──工具调用──▶ 宿主（通用机制：executor: "page" 的工具交给发起任务的页面）
                              │  tool_calls（长轮询领取） / tool_result（交回结果）
                              ▼
                    用户的页面：PptViewer + PageToolLink ──保存（带 expected_revision）──▶ 插件后端
```

## 工具定义

`src/page/page-tools.ts` 的 `PAGE_TOOLS` 就是 `TOOLS` 加上 `executor: "page"`：
`get_outline`、`apply_edits`、`render_slide`、`check_layout`、`get_template`、`build_deck`。
插件在握手的工具清单里原样声明它们，宿主据此把调用转给页面。

```sh
npm run page-tools -- ../ai-ppt-server/service/ppt_runtime/page_tools.json
```

把定义写成 JSON，供插件后端内嵌；工具定义有变化时要重新生成。

## 页面里用到的部件（`src/page/`）

| 部件 | 作用 |
|---|---|
| `runPageTool(viewer, name, args)` | 在当前打开的文稿上执行一个工具，不会抛异常（失败时返回 `ok: false`）。
| `runPageTool(viewer, name, args, {resolveAsset})` 里的 `resolveAsset` | `addImage.src` 可以写 `project:images/...`（项目里的图片，例如宿主导入的附件，`read_attachment` 返回的 `remote_path`）。页面在执行这批指令之前调用 `resolveAsset(path)` 取回图片、换成 PNG/JPEG/GIF 的 data URL，所以引擎的校验不变；不给 `resolveAsset` 时这种指令失败。同一路径只取一次。 |
| `PptViewer.renderPagePng(page, width)` | 在 Web Worker 里把一页渲染成 PNG（OffscreenCanvas），不阻塞编辑器，标签页切到后台也不会停。 |
| `PageToolLink` | 循环领取宿主交来的调用，按顺序执行、交回结果；修改类工具成功后先保存，回复里带上 `{saved: true, revision}`。 |
| `AutoSaver` | Agent 修改后立即保存，用户停顿 2.5 秒后保存；同一时间只有一次保存在进行；每次保存都带上它所基于的 revision，冲突时停下来等页面决定；上传失败会自动重试。 |
| `resolveOpenFamily` / `OPEN_FONT_MAP` | 只用开源字体渲染：微软雅黑、等线换成 Noto Sans CJK SC，宋体换成 Noto Serif CJK SC，Calibri 换成 Carlito，Cambria 换成 Caladea，Arial 换成 Liberation Sans。文件里保留原来的字体名。 |

页面需要引入 `src/editor/editor.css`（编辑层的样式），并给编辑器容器加上类名 `canvas-container`。

## 和宿主之间的约定（宿主通用，不针对 PPT）

- 插件在工具定义里写 `"executor": "page"`，这个工具就由页面执行，而不是插件的服务。
- 页面通过插件会话的两个方法与宿主交互：
  - `tool_calls {task_id, page_id, wait_ms}` 领取待执行的调用（最多挂起 25 秒）。
  - `tool_result {task_id, page_id, call_id, result}` 交回结果。
- 发送 `prompt` 时带上 `page_id` 的页面，负责执行这个任务的页面工具；同一任务只认最后一个发送 prompt 的页面。
- 等待执行的调用已经排队、但负责的页面 45 秒没来领取，或者页面领走后 180 秒没有交回结果，宿主就返回 `PAGE_CLOSED`，并结束这个任务。关闭页面即视为放弃任务。
- 结果格式就是 `ToolResult`（`{ok, content: [text | image]}`）。宿主把图片存成视觉素材，Agent 用 `artifact_id` 查看。

## 字体

页面从插件后端的 `/api/fonts`、`/api/font` 加载字体（`fontBackendUrl`）。后端只下发开源字体，不下发微软雅黑。
Noto CJK 一个字重约 20 MB，第一次加载会慢一些，之后由浏览器缓存。
渲染替换后的字体字宽和原字体略有差别，排版时要给文本框留一点余量。
