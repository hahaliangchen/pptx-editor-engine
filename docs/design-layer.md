# PPTX ⇄ JSON ⇄ HTML 双向转换

只持久化 PPTX。JSON 为内存中间表示，HTML 为模型设计视图，不新增设计 JSON 文件、对象存储键或数据库表。原生组、对象标识、模板名字及源包未知内容从 PPTX 中读取；无法由 PPTX 表达的临时业务元数据不会承诺跨浏览器重载保留。

## 转换函数

浏览器入口：`src/design/browser.ts`。实际产品使用 `createBrowserDesignConverter(viewer)`；同一个转换器提供 `syncDesign`、`materializeDesign`、`convertDesign`，共享底层导入/导出及按 ID 合并规则。

Node 入口 `src/design/index.ts` 保留给 CLI、测试及独立批处理，不是页面的转换后端。下面为 Node 调用示例：

```ts
import { pptxToJson, jsonToHtml, htmlToJson, jsonToPptx } from './src/design';
const json = await pptxToJson(currentPptxBytes);
const html = jsonToHtml(json, json.pages[0].id);
// 将 html 提供给 AI，并取得 editedHtml。
const changed = await htmlToJson(json, editedHtml);
const pptx = await jsonToPptx(changed);
// 只保存 pptx。JSON 在本次转换结束后即可丢弃。
```

`convertDesign` 额外提供 sync/apply/pages 操作；`syncDesign` 从编辑器新快照读取外观，保留本次内存会话中的身份；`materializeDesign` 写出 PPTX 再重读，以真实源对象编号建立后续编辑映射。

## 产品连接

- `DeckEditor` 直接调用浏览器转换器，不再发送 `/deck/design` 请求。PPTX 解析、JSON 合并、XML/ZIP 写出在用户浏览器执行；复杂素材预览使用现有 Rust/WASM 渲染 worker。
- HTML 在固定 1280×720 的隔离 iframe 内测量；禁止脚本及外部资源请求，用父编辑器已加载的字体。测量完成或失败后移除 iframe。哈希使用纯 JS 实现，兼容局域网 HTTP 没有 Web Crypto subtle 的情况。
- 转换仅产生本地结果；既有保存接口仍负责鉴权、版本冲突检测和 PPTX 持久化。字体、项目素材仍按需从服务器读取。
- 原服务端 `/design` 接口和 Node 转换入口保留供独立工具使用，产品前端不再依赖它们。
- 编辑器 `snapshotPptx()` 提供不清除脏标记的当前数据。`DesignBridge` 在内存中缓存当前文稿各页的 HTML。手动修改提交时同步使旧 HTML 失效，输入、拖动或批量编辑完成后自动刷新；撤销和重做也走同一路径。刷新期间不提供旧 HTML；转换期间发生修改则拒绝安装旧结果。
- `get_design(page)` 返回最新 HTML、版本、源哈希和资产引用。省略 page 返回页面概要。
- `apply_html(expected_revision, html)` 更新编辑器，随后通过原有 AutoSaver/DeckSave 保存 PPTX。保存失败沿用现有错误/冲突机制，不能报告已保存。
- `manage_design_pages` 支持显式新增、删除、移动页面，不自动排版，也不限制为 5 页。
- `design-asset:` 引用原稿资产；`project:images/...` 由页面解析为可导出的图片。模型无需接收大段图片 base64。

网页转换无需服务器启动 Chromium。只有 Node CLI/无头转换测试需要 Playwright Chromium，可用 PPT_DESIGN_BROWSER 指定可执行文件。字体由编辑器统一解析和加载，HTML 与 WASM 使用相同字体字节。本次仅完成代码接入和本地验证，尚未部署线上。

## 统一编辑线路与差异合并

AI 页面内容操作固定为 `get_design → 修改 HTML → apply_html`；原生 `apply_edits`、`build_deck` 不再对 AI 暴露，页面工具执行器也会拒绝这两个入口。用户手动操作仍使用编辑器原生能力；完成后自动生成最新 HTML。页面增删排序属于结构操作，保留 `manage_design_pages`。

原始 HTML 由宿主缓存，不接受 AI 自报基线。转换器分别测量原始 HTML 和修改后的 HTML，按稳定 data-id 比较实际变化，再把字段差异合并到原 JSON。未变化的原生几何、字体、边距等保持原值；仅改文字不会用浏览器测量结果整体覆盖原生样式。共享 CSS、flex/grid 引起的关联节点变化也会被记录。这里仍可能测量整页，限制的是写回范围，而非浏览器布局计算范围。

溢出校验针对受影响节点，未改节点原先已有的溢出不会阻止一次局部修改。锁定、必填、身份及组关系仍按原 JSON 校验。无实际变化时不重建 PPTX；成功修改返回节点及字段的变更摘要。该摘要不是新增的撤销或逐项确认界面。

HTML 和 JSON 均为临时投影，持久化仍只写 PPTX。重读后 HTML 是绝对定位投影，不保留原 flex/grid 源代码；复杂对象继续作为独立素材，不通过直接修改 OOXML 绕过设计层。

## 基础能力及边界

支持 16:9 PPTX / 1280×720 CSS px；基础文字、形状、图片、组合、纯色背景、统一实线边框、统一圆角、中文测量换行、富文本、对齐、文本内边距与段距。导出写入真实文字/形状/图片，不把整页截图，也不自动缩字。

复杂原对象以独立资产呈现，源包保留。未知且无预览的对象明确诊断，不静默删除。当前不承诺任意 SmartArt/动画/旋转效果编辑，也不承诺所有 CSS 支持。渐变、阴影、滤镜、复杂裁剪、伪元素等应作为独立视觉素材；新 SVG 使用 PNG 显示资产保留 SVG 原稿。

`group` 只表达组合关系，其背景/边框应由组内独立 shape 表达。文本及有色容器必须有 data-kind；原对象保留 data-id，新对象由转换器分配 ID。页面 body 保留 data-page-id、data-design-revision、data-source-hash。

既有测试覆盖中文富文本、组、源 XML 保留、锁定防绕过、过期拒绝、手改再设计、仅保存 PPTX 后重开、20 页增删排序、HTTP 转换与重新渲染。技术回归不等同于实际业务文稿的审美验收。另已对照一页中文/组合/独立图标样例的 HTML、引擎渲染和 LibreOffice 导出预览；修复了混合文字/字段/换行读取漏掉 a:br，以及渲染器未拆分段内硬换行的问题。保留浮点字号以避免导入后字号漂移。

## Skill

agent 仓库 `skills/ppt-html-design/` 与 PPT 插件 `ppt-master` 已声明所有 AI 页面内容修改必须走 HTML 设计层，包括单字修改、模板填充和整页重设计。Anthropic 审美章节保留原文和校验哈希；未引入其 html2pptx.js / PptxGenJS 完整生成流水线。

本轮浏览器迁移部署需要重建并发布前端（包含引擎及浏览器 WASM）；不依赖新增转换接口权限或服务器 Chromium。完整设计层的首次升级还需发布之前的后端工具声明及 skill；独立 Node 工具仍需其自身 WASM/Playwright 依赖。这里未修改线上权限或数据。

## 本地验证记录（2026-10-22）

转换相关回归、引擎及前端 TypeScript 检查、前端 QA 构建、后端转换接口与工具声明测试通过。引擎全量测试运行时为 468 通过、4 失败；4 个失败在恢复原解析器的对照运行中同样存在，涉及原有 API 操作数量断言及旧 build_deck 目录填充预期。本次另补的模板参考图替换/删除用例通过。全量测试并非全绿，不能将定向验证描述为整个项目验收通过。

后续修复：目录项的无标注示例说明仅在同一条目背景内、紧跟标题且对齐时关联，填充/删除条目会同步处理说明；带用途标注的参考图及独立正文不据相邻顺序推断为说明。更新操作数量和目录自由设计的旧断言，增加防误删回归。2026-10-22 00:15 的引擎全量回归为 **41 个测试文件、474 项测试全部通过**，TypeScript 检查通过。该结果取代上面的首次测试状态；仍未部署服务器。

浏览器迁移后，2026-10-22 00:35 全量回归为 **42 个文件、477 项测试全部通过**；引擎和前端 TypeScript 检查、前端 QA 构建通过（保留原有 CSS 顺序及包体积警告）。新增真实浏览器测试打包实际 PptViewer/Rust WASM，禁用 crypto.subtle 后验证本地转换、独立图形预览、中文换行、组合及粗体、只保存 PPTX 后重开、旧版本拒绝和 20 页管理；服务端没有收到转换请求。HTML 脚本/外部图片被拒绝，外部 CSS 被 CSP 阻止，失败后 iframe 清理正常。已人工查看浏览器导出的 PNG，文字完整且位置符合测量布局。仍未更新线上部署。

统一 HTML 编辑入口后，本轮全量回归为 **43 个测试文件、483 项测试全部通过**；引擎和前端 TypeScript 检查、前端 QA 构建、后端页面工具声明测试通过（构建仍有既有 CSS 顺序及包体积警告）。新增覆盖单词修改仅更新文本、颜色修改保留字体与边距、未修改节点 JSON 与源 XML 不变、flex 联动、已有溢出与受影响节点校验、语义无变化不重建 PPTX，以及真实浏览器内手动编辑/撤销/重做后自动刷新 HTML、旧版本拒绝。代码与 skill 已更新，尚未部署线上。
