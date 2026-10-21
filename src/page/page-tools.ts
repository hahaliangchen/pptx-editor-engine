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
export const DESIGN_TOOL_GUIDE = "AI 的所有页面内容修改统一走 HTML：改一个字、填模板、调样式和重设计都先 get_design(page)，再 apply_html(expected_revision, html)。宿主按原始 HTML 与新 HTML 的差异只合并变化字段，未改节点保留原值；用户手动编辑完成后自动刷新 HTML，旧版本失效。JSON 仅为内存中间表示，只保存 PPTX。manage_design_pages 只管理页面增加、复制、删除和顺序。固定 Logo、品牌背景和未标注素材不得改动或遮挡；只有制作人明确确认才可改变用途。不要调用 apply_edits 或 build_deck，也不要输出原生 JSON 编辑指令。";
const designTools: ToolDefinition[] = [
  { name: "get_design", description: "读取编辑器最新 PPTX 的 HTML 设计投影。不带 page 返回页列表；带 page 返回该页完整 HTML、设计版本、稳定对象 ID 和组关系。修改后需重新读取。", inputSchema: { type: "object", properties: { page: { type: "integer", minimum: 1 } }, additionalProperties: false } },
  { name: "apply_html", description: "AI 唯一的页面内容修改入口。提交改过的一页完整 1280×720 HTML/CSS，即使只改一个节点的文字。宿主对照原始 HTML 按 ID 与属性合并差异；未改节点保持原值，布局联动的变化也会测量，并返回修改摘要。保留 data-id、组关系、页面与 source hash；旧版本、锁定修改或不支持样式会拒绝。", inputSchema: { type: "object", properties: { expected_revision: { type: "integer", minimum: 0 }, html: { type: "string", minLength: 1 } }, required: ["expected_revision", "html"], additionalProperties: false } },
  { name: "manage_design_pages", description: "显式增加、复制、删除或移动页面。复制参考页继承原始素材、母版及保护规则，副本有独立 ID；不生成内容、不自动布局。先 get_design 读取版本；页码从 1 开始，after=0 在最前添加。批量执行后用 get_design 获取页面 ID 并设计 HTML。", inputSchema: { type: "object", properties: { expected_revision: { type: "integer", minimum: 0 }, commands: { type: "array", minItems: 1, maxItems: 100, items: { oneOf: [
    { type: "object", properties: { op: { const: "addSlide" }, after: { type: "integer", minimum: 0 } }, required: ["op"], additionalProperties: false },
    { type: "object", properties: { op: { const: "duplicateSlide" }, slide: { type: "integer", minimum: 1 }, after: { type: "integer", minimum: 0 } }, required: ["op", "slide"], additionalProperties: false },
    { type: "object", properties: { op: { const: "deleteSlide" }, slide: { type: "integer", minimum: 1 } }, required: ["op", "slide"], additionalProperties: false },
    { type: "object", properties: { op: { const: "moveSlide" }, slide: { type: "integer", minimum: 1 }, to: { type: "integer", minimum: 1 } }, required: ["op", "slide", "to"], additionalProperties: false }
  ] } } }, required: ["expected_revision", "commands"], additionalProperties: false } }
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
export const MUTATING_TOOLS = new Set(["apply_html", "manage_design_pages"]);

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
  if (name === "build_deck" || name === "apply_edits") return Promise.resolve({ ok: false, content: [{type: "text", text: "All AI page edits must use get_design and apply_html, including text-only changes. Native edit commands are not available to the agent."}] });
  return callTool(viewerToolHost(viewer, options), name, args);
}
