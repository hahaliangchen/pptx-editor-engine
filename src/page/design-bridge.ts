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
      if (epoch !== this.epoch || this.disposed) throw new Error("STALE_DESIGN: 文稿在读取期间发生变化，请重新读取");
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
        const attrs: Record<string,string> = edit.nodeId ? {purpose:"data-purpose",note:"data-note"} : {role:"data-page-role",layoutName:"data-layout-name",note:"data-page-note"};
        for (const [key,attr] of Object.entries(attrs)) if (key in edit) {
          const value = (edit as unknown as Record<string,string>)[key];
          if (value) element.setAttribute(attr,value); else element.removeAttribute(attr);
        }
      }
      return this.execute("apply_html", {expected_revision:design.revision,html:"<!doctype html>"+html.documentElement.outerHTML}, true);
    } catch(error) { return {ok:false,content:[{type:"text",text:error instanceof Error?error.message:String(error)}]}; }
  }
  async run(name: string, args: any): Promise<ToolResult> { return this.execute(name,args); }
  private async execute(name: string, args: any, humanForm = false): Promise<ToolResult> {
    if (this.running) return { ok: false, content: [{ type: "text", text: "Design conversion is already running" }] };
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
          diagnostics:design.diagnostics,designGuide:"先 get_design(page) 阅读 HTML；复用或重设计都通过 apply_html。页面角色可选，固定品牌不可修改或遮挡。参考页可用 duplicateSlide 复制。"});
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
        return this.result({ ...identity, page: a.page, pageId: page.id, semantics: page.semantics ?? {}, semanticsDefined:page.semantics!==undefined, nodes: page.nodes.map(n => ({id:n.id, shapeId:n.origin?.shapeId ?? n.metadata?.sourceShapeId, name:n.name, purpose:effectivePurpose(page,n), note:n.note, locked:effectiveLock(page,n)})), assets: page.nodes.filter(n => n.appearance.src).map(n => ({ id: n.id, name: n.name, src: `design-asset:${encodeURIComponent(n.id)}`, locked: effectiveLock(page,n), purpose: effectivePurpose(page,n), note:n.note })), html: this.html.get(page.id) });
      }
      if (!this.design || this.captured !== this.epoch) throw new Error("STALE_DESIGN: 请先 get_design 读取编辑器的最新页面");
      const epoch = this.epoch, before = this.design;
      if (a.expected_revision !== before.revision) throw new Error(`STALE_DESIGN: request revision is no longer current. ${JSON.stringify({reason:"requestRevision",expected:before.revision,received:a.expected_revision??null,recovery:"Call get_design, redo your changes on its latest HTML, and retry with that exact revision; do not relabel old HTML."})}`);
      const request: DesignRequest = name === "apply_html"
        ? { operation: "apply", design: before, expectedRevision: before.revision, html: a.html }
        : { operation: "pages", design: before, expectedRevision: before.revision, commands: a.commands };
      if (request.html) {
        const pageId = parseDesignHtml(request.html).body.dataset.pageId;
        request.baselineHtml = pageId ? this.html.get(pageId) : undefined;
        if (!request.baselineHtml) throw new Error("STALE_DESIGN: no current HTML baseline for this page");
        const proposal = templateSemanticProposal(before, request.baselineHtml, request.html);
        if (proposal.changes.length) {
          if (this.options.mode !== "template-authoring") throw templateSemanticPermissionError(proposal,"generation cannot change template purposes or annotation notes; restore the original values");
          if (!humanForm && !await this.options.confirmSemanticChanges?.(proposal)) throw new Error("TEMPLATE_CONFIRMATION_REQUIRED: author did not approve these annotation changes");
          if (epoch !== this.epoch || this.disposed) throw new Error("STALE_DESIGN: document changed while annotation confirmation was open");
          request.semanticAuthorization = { ...proposal, mode:"template-authoring" };
        }
      }
      if (request.html?.includes("project:")) {
        const html = parseDesignHtml(request.html);
        for (const image of Array.from(html.querySelectorAll("img"))) {
          const src = image.getAttribute("src") ?? "";
          if (!src.startsWith("project:")) continue;
          if (!this.resolveAsset) throw new Error("Project asset resolver is unavailable");
          image.setAttribute("src", await this.resolveAsset(src.slice("project:".length)));
        }
        request.html = "<!doctype html>" + html.documentElement.outerHTML;
      }
      const next = await this.convert(request);
      if (epoch !== this.epoch || this.disposed || this.viewer.hasActiveInteraction?.()) throw new Error("STALE_DESIGN: 转换期间用户修改了稿件，未应用 AI 结果");
      if (next.revision === before.revision && next.source.sha256 === before.source.sha256) return this.result({ revision: next.revision, applied: false, needsSave: this.dirty, changes: [] });
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
      return this.result({ revision: next.revision, pages: next.pages.length, pageList: next.pages.map((p,i)=>({page:i+1,id:p.id})), applied: true, needsSave: true, changes: next.changes ?? [] });
    } catch (error) {
      return { ok: false, content: [{ type: "text", text: error instanceof Error ? error.message : String(error) }] };
    } finally { this.running = false; if (this.captured !== this.epoch) this.scheduleRefresh(); }
  }
  private result(value: unknown): ToolResult { return { ok: true, content: [{ type: "text", text: JSON.stringify(value) }] }; }
}
