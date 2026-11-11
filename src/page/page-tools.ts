import { TOOLS, ToolDefinition, ToolHost, ToolResult, callTool } from "../api/tools";
import type PptViewer from "../index";
import { AssetResolver, resolveProjectAssets } from "./assets";

/**
 * A tool the plugin declares with `executor: "page"`: the Agent host hands
 * calls to the page that started the task instead of the plugin's service.
 */
export interface PageToolDefinition extends ToolDefinition {
  executor: "page";
}

/** Every deck tool runs in the page, on the deck the user is looking at. */
export const DESIGN_TOOL_GUIDE = "AI 通过 HTML 编辑页面：处理某页时用 get_design(page) 获取基线，已有 HTML 可以直接复用，再用 edit_html(page, edits) 像编辑代码一样提交 old_string/new_string 局部替换；不要为局部修改重抄整页 HTML。排错用 validate_html，只读预检不修改或保存正式稿；完整页内容提交 apply_html/edit_html/fill_layout 时显式传 content_complete=true，普通写入只表示修改保存，不代表完成。填充方式由模板制作人按页标记（body 的 data-page-fill）：layout 是按版式填充，块和样式已定，用 get_layout(page) 读取块与字段、fill_layout 只提交内容（每组块的条目、文字、SVG 图标），宿主复制原块并重新排列，不输出整页 HTML，也不重画；free 或未标记是自由设计，参照模板的品牌与样式，根据内容和设计意图自行选择构图、样式和布局。浏览器测量排版结果后写入可编辑 PPT。全部内容保存后即可导出；仅用户明确要求时调用模型视觉检查。宿主只合并变化字段，未改节点保留原值；需要当前页面内容或用户手动编辑后按需读取。溢出、遮挡和越界只返回本地提示，不拦截保存或导出，也不要求视觉验收记录。保留固定 Logo、品牌素材及源对象元数据；用途变更仍需制作人确认。只保存 PPTX；manage_design_pages 管理页序与数量。不要调用 apply_edits 或 build_deck，也不要输出原生 JSON 编辑指令。";
const designTools: ToolDefinition[] = [
  { name: "get_design", description: "读取当前 PPTX 的 HTML 投影。不带 page 返回页列表；带 page 返回完整页面 HTML（按版式填充页默认只返回页面信息，改用 get_layout；确需整页 HTML 时传 html: true）、design_revision（设计版本，与保存稿件的 stored_revision 无关）、稳定对象 ID 与组关系。HTML 描述当前 PPT 和手动编辑后的状态；body 的 data-page-fill='layout' 时用 get_layout 和 fill_layout 按版式填充，free 或未标记时用本工具读取 HTML 自由设计。有效结果直接复用；版本过期、手动改稿或缺少 HTML 时再读。", inputSchema: { type: "object", properties: { page: { type: "integer", minimum: 1 }, html: { type: "boolean", description: "按版式填充页也返回整页 HTML" } }, additionalProperties: false } },
  { name: "validate_html", description: "只读预检一页完整 HTML/CSS：执行与 apply_html 相同的转换、格式和品牌保护检查，返回修改预览与诊断；不修改当前稿件、不保存、不计入页面完成。排查转换错误或尝试最小样例时使用本工具。校验通过后，将实际成稿内容交给 apply_html；预检通过不等于制作完成。", inputSchema: { type: "object", properties: { html: { type: "string", minLength: 1 } }, required: ["html"], additionalProperties: false } },
  { name: "edit_html", description: "像编辑源码一样局部修改一页 HTML。复用 get_design 返回的当前 HTML，只提交 edits 中需要变化的 old_string/new_string；不要重新输出整页。每项 old_string 必须非空且在当前源中唯一，按顺序替换；找不到或多处匹配时整批不应用，请补充准确上下文。new_string 可为空以删除片段；插入内容时保留唯一锚点并追加新 HTML。未匹配的源码原样保留。沿用 apply_html 的页面身份、固定对象保护、测量、三页并发准备和保存回执。", inputSchema: { type: "object", properties: { page: { type: "integer", minimum: 1 }, edits: { type: "array", minItems: 1, items: { type: "object", properties: { old_string: { type: "string", minLength: 1 }, new_string: { type: "string" } }, required: ["old_string", "new_string"], additionalProperties: false } } , content_complete: { type: "boolean", description: "显式声明这页已完成确认大纲要求的实际内容；默认 false，普通编辑和试写不代表制作完成。" } }, required: ["page", "edits"], additionalProperties: false } },
  { name: "apply_html", description: "兼容完整页面提交：提交一页完整 1280×720 HTML/CSS 并按差异更新编辑器。修改已有 HTML 优先用 edit_html 只提交源码补丁，不重抄未改部分。仅当本页已具备确认大纲要求的完整实际内容时传 content_complete=true，普通编辑默认不计完成；调试或试写应使用只读 validate_html。返回修改摘要和版面诊断；溢出、遮挡、越界是提示，不阻止应用。data-page-fill='layout' 的页优先用 fill_layout 填充，确需微调时再提交 HTML 并保留块样式；free 或未标记的页由 AI 自由设计构图、样式和布局。浏览器测量后转换成可编辑 PPT。保存后即可导出，视觉检查仅按用户明确要求执行。保留 data-id、组关系和页面 ID；不检查版本或 source hash。锁定对象改写或不支持的格式仍会拒绝。", inputSchema: { type: "object", properties: { expected_design_revision: { description: "兼容旧调用的可选字段，不再校验或影响提交" }, html: { type: "string", minLength: 1 }, content_complete: { type: "boolean", description: "显式声明这页已完成确认大纲要求的实际内容；默认 false，普通编辑和试写不代表制作完成。" } }, required: ["html"], additionalProperties: false } },
  { name: "get_layout", description: "只读：读取按版式填充页（data-page-fill='layout'）的块结构。返回 sets（每组可重复的块：set 名、块名、排列方式、现有条目的文字字段和图标位，以及块内嵌套的 sets）和 texts（块外的独立文字，如页标题、页码，含 id）。比 get_design 的整页 HTML 小得多，填充版式页时用它。", inputSchema: { type: "object", properties: { page: { type: "integer", minimum: 1 } }, required: ["page"], additionalProperties: false } },
  { name: "fill_layout", description: "按版式填充一页（仅 data-page-fill='layout' 的页）：只提交内容，宿主在原页面上改。sets 中每项按 get_layout 的 set 名给出全部条目 items；条目数少于模板时删除多余块，多于模板时复制最后一块，并在原区域内重新排列（行、列或网格，可用 columns 指定列数），块的样式原样保留。items 为空数组表示删除这组块。每个条目：fields 按字段名给文字（字符串或段落数组，第一段沿用第一段样式，其余沿用第二段样式，列表标记保留）；icons 按图标位名给一段内联 <svg viewBox=...>（单色线性图标，尺寸和颜色由图标位决定，宿主按占位圈颜色统一着色）；sets 给块内嵌套的块组，用块名或 set 名指定。texts 按 id 替换块外文字。未提到的块保持原样并在结果的 fillNotes 中列出，样例内容不要留在成稿里。结果里的溢出或截断提示说明哪块文字放不下：精简那块的文字后再调用一次 fill_layout，不必改用 apply_html 整页重写；块的位置和尺寸由宿主排布，长标题和单行标签放不下时宿主会加宽或适当缩小字号。", inputSchema: { type: "object", properties: { page: { type: "integer", minimum: 1 }, content_complete: { type: "boolean", description: "实际内容完整时显式传 true；默认不标记完成。" }, texts: { type: "object", additionalProperties: { anyOf: [{ type: "string" }, { type: "array", items: { type: "string" } }] } }, sets: { type: "array", items: { $ref: "#/$defs/set" } } }, required: ["page"], additionalProperties: false, $defs: {
    set: { type: "object", properties: { set: { type: "string", minLength: 1 }, columns: { type: "integer", minimum: 1 }, items: { type: "array", items: { $ref: "#/$defs/item" } } }, required: ["set", "items"], additionalProperties: false },
    item: { type: "object", properties: { fields: { type: "object", additionalProperties: { anyOf: [{ type: "string" }, { type: "array", items: { type: "string" } }] } }, icons: { type: "object", additionalProperties: { type: "string" } }, sets: { type: "array", items: { $ref: "#/$defs/set" } } }, additionalProperties: false }
  } } },
  { name: "manage_design_pages", description: "显式增加、复制、删除或移动页面。复制参考页继承原始素材、母版及保护规则，副本有独立 ID；不生成内容、不自动布局。页码从 1 开始，after=0 在最前添加。批量执行后用 get_design 获取页面 ID 并设计 HTML。", inputSchema: { type: "object", properties: { expected_design_revision: { description: "兼容旧调用的可选字段，不参与校验" }, commands: { type: "array", minItems: 1, maxItems: 100, items: { oneOf: [
    { type: "object", properties: { op: { const: "addSlide" }, after: { type: "integer", minimum: 0 } }, required: ["op"], additionalProperties: false },
    { type: "object", properties: { op: { const: "duplicateSlide" }, slide: { type: "integer", minimum: 1 }, after: { type: "integer", minimum: 0 } }, required: ["op", "slide"], additionalProperties: false },
    { type: "object", properties: { op: { const: "deleteSlide" }, slide: { type: "integer", minimum: 1 } }, required: ["op", "slide"], additionalProperties: false },
    { type: "object", properties: { op: { const: "moveSlide" }, slide: { type: "integer", minimum: 1 }, to: { type: "integer", minimum: 1 } }, required: ["op", "slide", "to"], additionalProperties: false }
  ] } } }, required: ["commands"], additionalProperties: false } }
];
export const DESIGN_TOOLS = new Set([...designTools.map(t => t.name), "get_template", "suggest_annotations"]);
export const PAGE_TOOLS: PageToolDefinition[] = [...designTools, ...TOOLS.filter(tool => !["build_deck", "apply_edits"].includes(tool.name))].map(tool => ({
  ...tool,
  ...(tool.name === "get_template" ? { description: "只读获取所有页面、当前用途规则、可选页面角色、文字示例及素材 ID。无角色/单页品牌模板也支持。它们是 HTML 设计的参考；使用 get_design 和 apply_html 修改内容。模板用途默认不可改；制作模式下 data-purpose/data-note 与页面用途变更须制作人确认。" } : {}),
  ...(tool.name === "check_template" ? { description: "检查模板基本可用性、16:9 画布、对象身份及用途记录。页面角色、版式名称和填槽数量均非必填；单页品牌模板和未分类页合法。结果不代表审美或最终 PPT 视觉检查通过。" } : {}),
  ...(tool.name === "suggest_annotations" ? { description: "只读列出当前页面对象、固定/可变/参考用途与说明，供模板制作人审视。AI 根据 HTML 与实际渲染提出用途变化，确认后使用 apply_html；不返回或执行原生 setName/JSON 指令。" } : {}),
  executor: "page"
}));

/** Tools that change the deck; a successful call is saved before it is answered. */
export const MUTATING_TOOLS = new Set(["apply_html", "edit_html", "fill_layout", "manage_design_pages"]);

export interface PageToolOptions {
  /** Fetches a project picture named by `addImage.src: "project:images/..."`; without it such a command fails. */
  resolveAsset?: AssetResolver;
  design?: { run(name: string, args: unknown): Promise<ToolResult> };
}

/** The viewer as `callTool` sees it: the open deck, its edit API and the worker renderer. */
export function viewerToolHost(viewer: PptViewer, options: PageToolOptions = {}): ToolHost {
  return {
    outline: outlineOptions => viewer.api.outline(outlineOptions),
    apply: async input => {
      const resolved = await resolveProjectAssets(input, options.resolveAsset);
      return "failure" in resolved ? resolved.failure : viewer.api.apply(resolved.batch);
    },
    render: (page, renderOptions) => viewer.renderPagePng(page, renderOptions?.width)
  };
}

/** Run one page tool against the deck open in `viewer`. Never throws. */
export function runPageTool(viewer: PptViewer, name: string, args: unknown, options?: PageToolOptions): Promise<ToolResult> {
  if (DESIGN_TOOLS.has(name)) return options?.design ? options.design.run(name, args) : Promise.resolve({ ok: false, content: [{ type: "text", text: "HTML design converter is not connected" }] });
  if (name === "build_deck" || name === "apply_edits") return Promise.resolve({ ok: false, content: [{type: "text", text: "All AI page edits must use the HTML design layer: get_design and edit_html for source patches, or apply_html for a complete document. Native edit commands are not available to the agent."}] });
  return callTool(viewerToolHost(viewer, options), name, args);
}
