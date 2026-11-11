import { describeLayout, fillLayout, layoutIssues } from "../design/layout-fill";
import { templateSemanticPermissionError, templateSemanticProposal, type TemplateSemanticProposal, type TemplateSemanticEdit } from "../design/template-semantics";
import { effectivePurpose, effectiveLock } from "../design/template-policy";
import { nodeText } from "../design/model";
import type { DesignDocument } from "../design/model";
import { projectPage } from "../design/projection";
import type { DesignRequest } from "../design/service-core";
import type { ToolResult } from "../api/tools";
import { parseDesignHtml } from "../design/html-document";

export interface DesignViewer {
  snapshotPptx(): Promise<Blob>;
  hasActiveInteraction?(): boolean;
  loadPptx(buffer: ArrayBuffer): Promise<unknown>;
  getCurrentSlideIndex(): number;
  getSlidesCount(): number;
  selectSlide(index: number): Promise<unknown>;
}
export interface DesignBridgeOptions {
  mode?: "generation" | "template-authoring";
  convertInlineSvg?: (svg: string) => Promise<{ viewBox: [number, number, number, number]; shapesXml: string }>;
  resolveNativeAsset?: (path: string) => Promise<{ viewBox: [number, number, number, number]; shapesXml: string } | undefined>;
  /** Display these exact changes to the template author; never auto-approve model proposals. */
  confirmSemanticChanges?: (proposal: TemplateSemanticProposal) => Promise<boolean>;
}
export type DesignConverter = (request: DesignRequest) => Promise<DesignDocument>;
export function base64Bytes(value: string): Uint8Array {
  const decoded = atob(value), bytes = new Uint8Array(decoded.length);
  for (let i = 0; i < decoded.length; i++) bytes[i] = decoded.charCodeAt(i);
  return bytes;
}
export async function blobBase64(blob: Blob): Promise<string> {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  let text = "";
  for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
  return btoa(text);
}

/** Ephemeral design view. Only PPTX is saved; every manual edit invalidates old HTML. */
export class DesignBridge {
  private epoch = 0;
  private captured = -1;
  private design?: DesignDocument;
  private loading = false;
  private dirty = false;
  private running = false;
  private disposed = false;
  private refreshTimer?: ReturnType<typeof setTimeout>;
  private syncing?: Promise<DesignDocument>;
  private html = new Map<string, string>();
  constructor(private viewer: DesignViewer, private convert: DesignConverter,
    private resolveAsset?: (path: string) => Promise<string>, private options: DesignBridgeOptions = {}) {}
  edited(): void {
    if (this.loading || this.disposed) return;
    this.epoch++; this.html.clear();
    this.scheduleRefresh();
  }
  /** Called on edit completion too: typing may have flushed while still active. */
  scheduleRefresh(): void {
    if (this.disposed || this.loading || this.captured === this.epoch) return;
    clearTimeout(this.refreshTimer);
    this.refreshTimer = setTimeout(() => {
      if (this.running || this.viewer.hasActiveInteraction?.()) return;
      void this.refresh().catch(() => { /* Remain stale; get_design reports/retries the failure. */ });
    }, 0);
  }
  async refresh(): Promise<void> { await this.current(); }
  /** A stale projection is never exposed, including while an async refresh runs. */
  peekHtml(page: number): string | undefined {
    if (this.captured !== this.epoch || this.disposed || this.viewer.hasActiveInteraction?.()) return undefined;
    const id = this.design?.pages[page - 1]?.id;
    return id ? this.html.get(id) : undefined;
  }
  dispose(): void { this.disposed = true; this.reset(); }
  get pending(): boolean { return this.dirty; }
  saved(): void { this.dirty = false; }
  reset(): void { clearTimeout(this.refreshTimer); this.html.clear(); this.design = undefined; this.captured = -1; this.epoch++; this.dirty = false; }

  private project(design: DesignDocument): void {
    const projection = JSON.parse(JSON.stringify(design)) as DesignDocument;
    for (const page of projection.pages) for (const node of page.nodes) if (node.appearance.src) node.appearance.src = `design-asset:${encodeURIComponent(node.id)}`;
    this.html = new Map(projection.pages.map(page => [page.id, projectPage(projection, page.id)]));
  }

  private async current(): Promise<DesignDocument> {
    if (this.disposed) throw new Error("PAGE_CLOSED: design bridge disposed");
    if (this.viewer.hasActiveInteraction?.()) throw new Error("EDITOR_BUSY: 请完成当前输入或拖动后再读取设计");
    if (this.design && this.captured === this.epoch) return this.design;
    if (this.syncing) {
      try { await this.syncing; } catch { /* A concurrent manual edit may invalidate that snapshot. */ }
      return this.current();
    }
    const epoch = this.epoch;
    const work = (async () => {
      const pptxBase64 = await blobBase64(await this.viewer.snapshotPptx());
      const design = await this.convert({ operation: "sync", pptxBase64, design: this.design });
      if (this.disposed) throw new Error("PAGE_CLOSED: design bridge disposed");
      // Even an undo back to identical pixels invalidates a previously-issued HTML edit.
      design.revision = Math.max(design.revision, (this.design?.revision ?? -1) + 1);
      this.project(design);
      this.design = design; this.captured = epoch;
      return design;
    })();
    this.syncing = work;
    try { return await work; }
    finally {
      this.syncing = undefined;
      if (epoch !== this.epoch && !this.disposed) this.scheduleRefresh();
    }
  }

  /** Human author form operation, deliberately absent from the AI tool schema. */
  async updateTemplateSemantics(page: number, edits: TemplateSemanticEdit[]): Promise<ToolResult> {
    try {
      if (this.options.mode !== "template-authoring") throw new Error("TEMPLATE_PERMISSION: not a template authoring editor");
      if (this.running) throw new Error("Design conversion is already running");
      const design = await this.current(), selected = design.pages[page - 1];
      if (!selected) throw new Error("Invalid page number");
      const html = parseDesignHtml(this.html.get(selected.id)!);
      for (const edit of edits) {
        const element = edit.nodeId ? Array.from(html.querySelectorAll<HTMLElement>("[data-id]")).find(n=>n.dataset.id===edit.nodeId) : html.body;
        if (!element) throw new Error(`Unknown template node: ${edit.nodeId}`);
        const attrs: Record<string,string> = edit.nodeId ? {purpose:"data-purpose",note:"data-note",slot:"data-slot"} : {role:"data-page-role",layoutName:"data-layout-name",note:"data-page-note",fill:"data-page-fill"};
        for (const [key,attr] of Object.entries(attrs)) if (key in edit) {
          const value = (edit as unknown as Record<string,string>)[key];
          if (value) element.setAttribute(attr,value); else element.removeAttribute(attr);
        }
      }
      return this.execute("apply_html", {expected_revision:design.revision,html:"<!doctype html>"+html.documentElement.outerHTML}, true);
    } catch(error) { return {ok:false,content:[{type:"text",text:error instanceof Error?error.message:String(error)}]}; }
  }
  /** Human author form operation: wrap sibling objects in a new group, the block a layout-fill page repeats, and give it its slot. */
  async groupTemplateNodes(page: number, nodeIds: string[], slot?: string): Promise<ToolResult> {
    try {
      if (this.options.mode !== "template-authoring") throw new Error("TEMPLATE_PERMISSION: not a template authoring editor");
      if (this.running) throw new Error("Design conversion is already running");
      const design = await this.current(), selected = design.pages[page - 1];
      if (!selected) throw new Error("Invalid page number");
      if (nodeIds.length < 2) throw new Error("选择两个以上的元素再组成块");
      const html = parseDesignHtml(this.html.get(selected.id)!);
      const byId = new Map(Array.from(html.querySelectorAll<HTMLElement>("[data-id]")).map(n => [n.dataset.id!, n]));
      const members = nodeIds.map(id => { const el = byId.get(id); if (!el) throw new Error(`Unknown template node: ${id}`); return el; });
      const parent = members[0].parentElement!;
      if (members.some(el => el.parentElement !== parent)) throw new Error("只能把同一层的元素组成块：先选中同一个组合里的元素");
      if (members.some(el => el.dataset.locked || el.dataset.purpose === "fixed")) throw new Error("固定元素不能组成块：块里的元素生成时会被复制和移动");
      const px = (el: HTMLElement, key: "left" | "top" | "width" | "height") => parseFloat(el.style[key]) || 0;
      const left = Math.min(...members.map(el => px(el, "left"))), top = Math.min(...members.map(el => px(el, "top")));
      const right = Math.max(...members.map(el => px(el, "left") + px(el, "width"))), bottom = Math.max(...members.map(el => px(el, "top") + px(el, "height")));
      const group = html.createElement("div");
      // New objects are named after their id: keep the names the author sees, unique on the page.
      const replaced = new Set(members.flatMap(el => [el, ...Array.from(el.querySelectorAll<HTMLElement>("[data-id]"))]).map(el => el.dataset.id!));
      const taken = new Set([...byId.keys(), ...selected.nodes.filter(n => !replaced.has(n.id)).map(n => n.name)]);
      const fresh = (name: string) => { let id = `new:${name}`; for (let n = 2; taken.has(id); n++) id = `new:${name}-${n}`; taken.add(id); return id; };
      group.dataset.id = fresh(slot?.replace(/#\d+$/, "") || "块"); group.dataset.kind = "group";
      group.setAttribute("style", `position:absolute;left:${left}px;top:${top}px;width:${right - left}px;height:${bottom - top}px`);
      // Paint order: the block sits where its lowest member was, members keep their own order. Group membership of
      // existing objects never changes, so the members are rebuilt inside the block and keep their annotations below.
      const ordered = Array.from(parent.children).filter(el => members.includes(el as HTMLElement)) as HTMLElement[];
      parent.insertBefore(group, ordered[0]);
      const kept = new Map<string, Pick<TemplateSemanticEdit, "purpose" | "note" | "slot">>();
      for (const el of ordered) {
        el.style.left = `${px(el, "left") - left}px`; el.style.top = `${px(el, "top") - top}px`;
        for (const node of [el, ...Array.from(el.querySelectorAll<HTMLElement>("[data-id]"))]) {
          node.dataset.id = fresh((node.dataset.name || node.dataset.id!).replace(/^new:/, ""));
          if (node.dataset.note || node.dataset.slot || node.dataset.purpose === "reference") kept.set(node.dataset.id, { purpose: node.dataset.purpose as TemplateSemanticEdit["purpose"], note: node.dataset.note, slot: node.dataset.slot });
          for (const attr of ["data-purpose", "data-note", "data-slot"]) node.removeAttribute(attr);
        }
        group.appendChild(el);
      }
      const before = new Set(selected.nodes.map(n => n.id));
      const grouped = await this.execute("apply_html", { expected_revision: design.revision, html: "<!doctype html>" + html.documentElement.outerHTML }, true);
      if (!grouped.ok) return grouped;
      const created = (await this.current()).pages[page - 1].nodes.filter(n => !before.has(n.id));
      const named = (id: string) => created.find(n => n.name === id);
      const block = named(group.dataset.id);
      if (!block) throw new Error("组合已创建，但没有找到新块；请选中它后再设置块标记");
      const edits: TemplateSemanticEdit[] = Array.from(kept).flatMap(([id, value]) => { const node = named(id); return node ? [{ nodeId: node.id, purpose: value.purpose ?? "content", note: value.note, slot: value.slot }] : []; });
      if (slot) edits.unshift({ nodeId: block.id, purpose: "content", slot });
      return edits.length ? this.updateTemplateSemantics(page, edits) : grouped;
    } catch(error) { return {ok:false,content:[{type:"text",text:error instanceof Error?error.message:String(error)}]}; }
  }
  async run(name: string, args: any): Promise<ToolResult> { return this.execute(name,args); }
  private async execute(name: string, args: any, humanForm = false, nested = false): Promise<ToolResult> {
    if (this.running && !nested) return { ok: false, content: [{ type: "text", text: "Design conversion is already running" }] };
    this.running = true;
    try {
      const a = args ?? {};
      if (name === "get_template" || name === "suggest_annotations") {
        const design = await this.current();
        if (a.slide !== undefined && (!Number.isInteger(a.slide) || !design.pages[a.slide - 1])) throw new Error("Invalid page number");
        return this.result({ revision:design.revision, sourceHash:design.source.sha256, canvas:design.canvas,
          ...(name === "suggest_annotations" ? {advisory:"用途为当前规则，非修改授权。根据页面 HTML 和 render_slide 提出具体用途变化，制作人确认后才可 apply_html。"} : {}),
          pages:design.pages.map((page,index)=>({page:index+1,id:page.id,semantics:page.semantics ?? {},semanticsDefined:page.semantics!==undefined,
            objects:page.nodes.map(node=>({id:node.id,shapeId:node.origin?.shapeId ?? node.metadata?.sourceShapeId,name:node.name,kind:node.kind,parentId:node.parentId,purpose:effectivePurpose(page,node),locked:effectiveLock(page,node),note:node.note,rect:node.appearance.rect,text:nodeText(node),binding:node.binding,...(node.appearance.src?{asset:`design-asset:${encodeURIComponent(node.id)}`}:{})}))})).filter(page=>a.slide===undefined||page.page===a.slide),
          diagnostics:design.diagnostics,designGuide:"读取当前 PPT 的 HTML，有效结果直接复用。semantics.fill 为 layout 的页按版式填充：用 get_layout 和 fill_layout 只提交内容；free 或未标记的页根据内容和设计意图自由选择 HTML/CSS 构图、样式和布局；浏览器测量排版结果后写入可编辑 PPT。保留固定品牌素材。溢出和遮挡只作本地提示，不拦截保存或导出；视觉检查仅按用户明确要求执行。参考页可用 duplicateSlide 复制。"});
      }
      if (name === "get_design") {
        const design = await this.current();
        const identity = { revision: design.revision, sourceHash: design.source.sha256, canvas: design.canvas,
          totalPages: design.pages.length, diagnostics: design.diagnostics };
        // A page edit needs its own metadata; repeating all page semantics on every
        // read grows a twenty-page design session quadratically in context size.
        if (a.page === undefined) return this.result({ ...identity,
          pages: design.pages.map((p, i) => ({ page: i + 1, id: p.id, nodes: p.nodes.length, semantics: p.semantics ?? {}, semanticsDefined:p.semantics!==undefined })) });
        const page = design.pages[a.page - 1];
        if (!Number.isInteger(a.page) || !page) throw new Error("Invalid page number");
        // In generation a layout-fill page is filled from get_layout; its full HTML and node list only cost context.
        // Template authoring edits the page itself and always reads all of it.
        if (page.semantics?.fill === "layout" && a.html !== true && this.options.mode !== "template-authoring") return this.result({ ...identity, page: a.page, pageId: page.id, semantics: page.semantics, semanticsDefined: true,
          layoutFill: "按版式填充页：用 get_layout 读取块与字段，用 fill_layout 提交内容。确需整页 HTML 时传 html: true。" });
        return this.result({ ...identity, page: a.page, pageId: page.id, semantics: page.semantics ?? {}, semanticsDefined:page.semantics!==undefined, nodes: page.nodes.map(n => ({id:n.id, shapeId:n.origin?.shapeId ?? n.metadata?.sourceShapeId, name:n.name, kind:n.kind, parentId:n.parentId, purpose:effectivePurpose(page,n), note:n.note, slot:n.slot, locked:effectiveLock(page,n)})), ...(page.semantics?.fill === "layout" ? { layoutIssues: layoutIssues(this.html.get(page.id)!) } : {}), assets: page.nodes.filter(n => n.appearance.src).map(n => ({ id: n.id, name: n.name, src: `design-asset:${encodeURIComponent(n.id)}`, locked: effectiveLock(page,n), purpose: effectivePurpose(page,n), note:n.note })), html: this.html.get(page.id) });
      }
      if (name === "get_layout" || name === "fill_layout") {
        const design = await this.current(), page = design.pages[a.page - 1];
        if (!Number.isInteger(a.page) || !page) throw new Error("Invalid page number");
        if (page.semantics?.fill !== "layout") throw new Error(`FILL_LAYOUT: page ${a.page} is not marked layout fill; design it with get_design and apply_html`);
        const html = this.html.get(page.id)!;
        if (name === "get_layout") return this.result({ revision: design.revision, page: a.page, pageId: page.id, ...describeLayout(html) });
        const filled = fillLayout(html, { texts: a.texts, sets: a.sets });
        const applied = await this.execute("apply_html", { html: filled.html, content_complete: a.content_complete }, false, true);
        if (!applied.ok) return applied;
        // Content-only call, content-sized answer: counts and this page's diagnostics, not every changed field.
        const data = JSON.parse((applied.content[0] as { text: string }).text);
        const changes: Array<{ action: string }> = data.changes ?? [];
        const count = (action: string) => changes.filter(c => c.action === action).length;
        return { ...applied, content: [{ type: "text", text: JSON.stringify({ applied: data.applied, content_complete: data.content_complete, needsSave: data.needsSave, design_revision: data.design_revision, pages: data.pages, pageList: data.pageList,
          changed: { updated: count("update"), added: count("add"), removed: count("remove") },
          diagnostics: (data.diagnostics ?? []).filter((d: { pageId?: string }) => d.pageId === page.id).map((d: { message: string }) => d.message.split(" {")[0]).slice(0, 12),
          ...(filled.notes.length ? { fillNotes: filled.notes } : {}) }) }] };
      }
      await this.current();
      const before = await this.current();
      const request: DesignRequest = name === "apply_html" || name === "validate_html"
        ? { operation: "apply", design: before, expectedRevision: before.revision, html: a.html }
        : { operation: "pages", design: before, expectedRevision: before.revision, commands: a.commands };
      if (request.html?.toLowerCase().includes("<svg")) {
        const html = parseDesignHtml(request.html, true);
        for (const svg of Array.from(html.querySelectorAll("svg"))) {
          if (!this.options.convertInlineSvg) throw new Error("Inline SVG conversion is unavailable");
          const clone = svg.cloneNode(true) as SVGElement;
          // CSS layout and design metadata stay on the HTML image wrapper; the SVG
          // artwork itself is sent as a standalone, inert vector document.
          const paint = new Set(["color","stop-color","stop-opacity","fill","stroke","stroke-width","stroke-linecap","stroke-linejoin","opacity","fill-opacity","stroke-opacity"]);
          const layout = new Set(["width","height","position","left","right","top","bottom","display","margin","margin-left","margin-right","margin-top","margin-bottom","flex","flex-shrink","flex-grow","align-self"]);
          for (const el of [clone, ...Array.from(clone.querySelectorAll<SVGElement>("*"))]) {
            const style = el.style;
            for (const key of Array.from(style)) {
              if (paint.has(key)) el.setAttribute(key, style.getPropertyValue(key));
              else if (!(el === clone && layout.has(key))) throw new Error(`Unsupported inline SVG style: ${key}; use SVG presentation attributes`);
            }
            el.removeAttribute("style");
            if (el !== clone && el.hasAttribute("class")) throw new Error("Inline SVG child classes must be converted to presentation attributes");
          }
          for (const attr of Array.from(clone.attributes)) if (attr.name === "class" || attr.name.startsWith("data-") || attr.name === "aria-label" || attr.name === "role") clone.removeAttribute(attr.name);
          const source = new XMLSerializer().serializeToString(clone);
          const native = await this.options.convertInlineSvg(source);
          const encoded = btoa(Array.from(new TextEncoder().encode(source), b => String.fromCharCode(b)).join(""));
          const preview = `data:image/svg+xml;base64,${encoded}`;
          const img = html.createElement("img");
          for (const attr of Array.from(svg.attributes)) if (["class", "style", "width", "height", "id", "title", "aria-label", "role"].includes(attr.name) || attr.name.startsWith("data-")) img.setAttribute(attr.name, attr.value);
          if (img.dataset.kind && img.dataset.kind !== "image") throw new Error("Inline SVG must use data-kind=image");
          img.dataset.kind = "image"; img.src = preview;
          (request.nativeAssets ??= {})[preview] = native;
          svg.replaceWith(img);
        }
        request.html = "<!doctype html>" + html.documentElement.outerHTML;
      }
      if (request.html) {
        const pageId = parseDesignHtml(request.html).body.dataset.pageId;
        request.baselineHtml = pageId ? this.html.get(pageId) : undefined;
        if (!request.baselineHtml) throw new Error("Unknown design page");
        const proposal = templateSemanticProposal(before, request.baselineHtml, request.html);
        if (proposal.changes.length) {
          if (this.options.mode !== "template-authoring") throw templateSemanticPermissionError(proposal,"generation cannot change template purposes or annotation notes; restore the original values");
          if (name !== "validate_html" && !humanForm && !await this.options.confirmSemanticChanges?.(proposal)) throw new Error("TEMPLATE_CONFIRMATION_REQUIRED: author did not approve these annotation changes");
          if (this.disposed) throw new Error("PAGE_CLOSED: design bridge disposed");
          request.semanticAuthorization = { ...proposal, mode:"template-authoring" };
        }
      }
      if (request.html?.includes("project:")) {
        const html = parseDesignHtml(request.html);
        for (const image of Array.from(html.querySelectorAll("img"))) {
          const src = image.getAttribute("src") ?? "";
          if (!src.startsWith("project:")) continue;
          if (!this.resolveAsset) throw new Error("Project asset resolver is unavailable");
          const path = src.slice("project:".length);
          const preview = await this.resolveAsset(path);
          const native = await this.options.resolveNativeAsset?.(path);
          if (native) (request.nativeAssets ??= {})[preview] = native;
          image.setAttribute("src", preview);
        }
        request.html = "<!doctype html>" + html.documentElement.outerHTML;
      }
      // Validation uses the same conversion and protection checks on an isolated
      // snapshot, but never installs the result into the viewer or marks it dirty.
      if (name === "validate_html") request.design = JSON.parse(JSON.stringify(before));
      const next = await this.convert(request);
      if (this.disposed) throw new Error("PAGE_CLOSED: design bridge disposed");
      if (name === "validate_html") return this.result({ revision: before.revision, validated: true, applied: false, content_complete: false, needsSave: false, changes: next.changes ?? [], diagnostics: next.diagnostics });
      if (next.revision === before.revision && next.source.sha256 === before.source.sha256) return this.result({ revision: next.revision, applied: false, content_complete: a.content_complete === true, needsSave: this.dirty, changes: [] });
      const selected = this.viewer.getCurrentSlideIndex();
      this.loading = true;
      try {
        await this.viewer.loadPptx(base64Bytes(next.source.pptxBase64).buffer as ArrayBuffer);
        await this.viewer.selectSlide(Math.min(selected, this.viewer.getSlidesCount() - 1));
      } catch (error) {
        // A parser/render failure must not leave a half-installed deck.
        await this.viewer.loadPptx(base64Bytes(before.source.pptxBase64).buffer as ArrayBuffer);
        await this.viewer.selectSlide(Math.min(selected, this.viewer.getSlidesCount() - 1));
        throw error;
      } finally { this.loading = false; }
      this.epoch++; this.captured = this.epoch; this.design = next; this.dirty = true;
      this.project(next);
      return this.result({ revision: next.revision, pages: next.pages.length, pageList: next.pages.map((p,i)=>({page:i+1,id:p.id})), applied: true, content_complete: a.content_complete === true, needsSave: true, changes: next.changes ?? [], diagnostics: next.diagnostics });
    } catch (error) {
      return { ok: false, content: [{ type: "text", text: error instanceof Error ? error.message : String(error) }] };
    } finally { this.running = false; if (this.captured !== this.epoch) this.scheduleRefresh(); }
  }
  private result(value: any): ToolResult {
    const {revision,...data}=value;
    return {ok:true,content:[{type:"text",text:JSON.stringify({...data,...(revision!==undefined?{design_revision:revision}:{})})}]};
  }
}
