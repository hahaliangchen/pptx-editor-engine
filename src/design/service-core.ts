import type { TemplateSemanticAuthorization } from "./template-semantics";
import { assertDesign, cloneDesign, DesignDocument } from "./model";
import { base64ToBytes, sha256Hex } from "./bytes";
import { parseDesignHtml } from "./html-document";
export interface DesignRequest {
  operation: "sync" | "apply" | "pages";
  design?: DesignDocument;
  pptxBase64?: string;
  html?: string;
  /** Host-owned original projection. This is not exposed in the AI tool schema. */
  baselineHtml?: string;
  /** Host-issued exact, version-bound author confirmation; never an AI tool argument. */
  semanticAuthorization?: TemplateSemanticAuthorization;
  expectedRevision?: number;
  commands?: Array<{ op: "addSlide"; after?: number } | { op: "duplicateSlide"; slide: number; after?: number } | { op: "deleteSlide"; slide: number } | { op: "moveSlide"; slide: number; to: number }>;
}


export interface DesignRuntime {
  importDesign(bytes: Uint8Array): Promise<DesignDocument>;
  exportDesign(design: DesignDocument, identities?: Map<string, string>): Promise<Uint8Array>;
  mergeHtml(design: DesignDocument, html: string, baselineHtml?: string, authorization?: TemplateSemanticAuthorization): Promise<DesignDocument>;
  pages(bytes: Uint8Array, commands: NonNullable<DesignRequest["commands"]>): Promise<Uint8Array>;
}
export function createDesignConverter(runtime: DesignRuntime) {
  const { importDesign, exportDesign, mergeHtml } = runtime;
  /** Refresh geometry/content from the editor while retaining canonical identities and business metadata. */
  async function syncDesign(bytes: Uint8Array, previous?: DesignDocument, identities?: Map<string, string>, increment = true): Promise<DesignDocument> {
    const sha = sha256Hex(bytes);
    if (previous) {
      assertDesign(previous);
      if (sha === previous.source.sha256 && !identities) return cloneDesign(previous);
    }
    const next = await importDesign(bytes);
    if (!previous) return next;
    next.revision = previous.revision + (increment ? 1 : 0);
    for (const page of next.pages) {
      const oldPage = previous.pages.find(p => p.part === page.part);
      if (!oldPage) continue;
      page.id = oldPage.id;
      if (!increment) page.semantics = oldPage.semantics === undefined ? undefined : cloneDesign(oldPage.semantics);
      const oldById = new Map(oldPage.nodes.map(n => [n.id, n]));
      const oldBySource = new Map(oldPage.nodes.map(n => [n.origin ? `${n.origin.part}#${n.origin.shapeId}` : n.id, n]));
      const remap = new Map<string, string>();
      page.nodes = page.nodes.map(node => {
        const stable = identities?.get(node.id);
        const before = stable ? oldById.get(stable) : oldBySource.get(node.id);
        if (!before) return node;
        remap.set(node.id, before.id);
        // Source-backed appearance is fresh; HTML-inexpressible business fields survive both directions.
        return { ...before, ...node, id: before.id, role: node.role, binding: node.binding,
          ...(!increment ? { purpose: before.purpose, note: before.note } : {}),
          ...(!increment && !before.appearance.text?.frames ? { appearance: cloneDesign(before.appearance), origin: node.origin ? { ...node.origin, appearance: cloneDesign(before.appearance) } : undefined } : {}),
          required: before.required, constraints: before.constraints,
          lock: before.lock ?? node.lock, metadata: { ...before.metadata, ...node.metadata } };
      });
      for (const node of page.nodes) if (node.parentId) node.parentId = remap.get(node.parentId) ?? node.parentId;
    }
    assertDesign(next);
    return next;
  }

  /** Export then rebase: newly-created shapes keep the same ID on every subsequent edit/save/reload. */
  async function materializeDesign(design: DesignDocument): Promise<DesignDocument> {
    const identities = new Map<string, string>();
    const bytes = await exportDesign(design, identities);
    const result = await syncDesign(bytes, design, identities, false);
    if (design.changes) result.changes = cloneDesign(design.changes);
    result.diagnostics.push(...cloneDesign(design.diagnostics.filter(d => d.severity === "warning")));
    return result;
  }

  /** Stateless converter shared by browser and CLI; persistence stays with the host. */
  async function convertDesign(request: DesignRequest): Promise<DesignDocument> {
    if (request.operation === "sync") {
      if (!request.pptxBase64) throw new Error("sync requires pptxBase64");
      return syncDesign(base64ToBytes(request.pptxBase64), request.design);
    }
    const design = request.design;
    if (!design) throw new Error("Read the current design before changing it");
    assertDesign(design);
    if (request.operation === "apply") {
      if (!request.html) throw new Error("apply requires a complete HTML document");
      const resolveAssets = (source: string) => {
        const html = parseDesignHtml(source);
        for (const image of Array.from(html.querySelectorAll("img"))) {
          const src = image.getAttribute("src") ?? "";
          if (!src.startsWith("design-asset:")) continue;
          const id = decodeURIComponent(src.slice("design-asset:".length));
          const node = design.pages.flatMap(p => p.nodes).find(n => n.id === id);
          if (!node?.appearance.src) throw new Error(`Unknown design asset: ${id}`);
          image.setAttribute("src", node.appearance.src);
        }
        return "<!doctype html>" + html.documentElement.outerHTML;
      };
      const merged = await mergeHtml(design, resolveAssets(request.html), request.baselineHtml ? resolveAssets(request.baselineHtml) : undefined, request.semanticAuthorization);
      if (merged.revision === design.revision) return merged;
      return materializeDesign(merged);
    }
    if (request.operation === "pages") {
      const commands = request.commands;
      if (!commands?.length || commands.length > 100 || commands.some(c => !["addSlide", "duplicateSlide", "deleteSlide", "moveSlide"].includes(c.op))) throw new Error("Only explicit page add/duplicate/delete/move operations are allowed");
      return syncDesign(await runtime.pages(await exportDesign(design), commands), design);
    }
    throw new Error("Unknown design operation");
  }

  return { syncDesign, materializeDesign, convertDesign };
}
