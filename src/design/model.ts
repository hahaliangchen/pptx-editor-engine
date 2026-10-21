import { checkFixedProtection, effectiveLock } from "./template-policy";
import type { TemplatePurpose, TemplatePageSemantics } from "./template-policy";

/** The design document owns identities and relationships; HTML is a disposable view. */
export interface DesignRect { x: number; y: number; w: number; h: number }
export interface DesignRun {
  text: string;
  fontFamily: string;
  fontSize: number;
  color: string;
  bold: boolean;
  italic: boolean;
  letterSpacing: number;
}
export interface DesignParagraph {
  runs: DesignRun[];
  align: "left" | "center" | "right";
  lineHeight: number;
  before: number;
  after: number;
}
export interface DesignText {
  paragraphs: DesignParagraph[];
  padding: [number, number, number, number]; // top, right, bottom, left in CSS px
  verticalAlign: "top" | "middle" | "bottom";
}
export interface DesignAppearance {
  rect: DesignRect;
  fill?: string;
  border?: { color: string; width: number };
  radius?: number;
  text?: DesignText;
  src?: string;
}
export interface DesignNode {
  id: string;
  kind: "group" | "shape" | "text" | "image" | "asset";
  name: string;
  purpose?: TemplatePurpose;
  note?: string;
  parentId?: string;
  appearance: DesignAppearance;
  lock?: "all" | "content";
  required?: boolean;
  /** An unsupported source object without a usable preview is retained, never silently removed. */
  hidden?: boolean;
  role?: string;
  binding?: unknown;
  constraints?: { maxChars?: number };
  metadata?: Record<string, unknown>;
  origin?: {
    shapeId: string;
    part: string;
    /** EMU -> the canonical 1280x720 design canvas. */
    transform: { scaleX: number; scaleY: number; offsetX: number; offsetY: number };
    appearance: DesignAppearance;
  };
}
export interface DesignPage { semanticsChanged?: boolean; semantics?: TemplatePageSemantics; id: string; part: string; nodes: DesignNode[]; background?: string; originBackground?: string }
export interface DesignDocument {
  schemaVersion: 1;
  revision: number;
  canvas: { width: 1280; height: 720; unit: "css-px" };
  /** Opaque original package: masters, relationships, animations and unknown XML survive. */
  source: { pptxBase64: string; sha256: string; widthEMU: number; heightEMU: number };
  pages: DesignPage[];
  diagnostics: Array<{ pageId: string; nodeId?: string; message: string }>;
  /** Ephemeral edit summary; never persisted alongside PPTX. */
  changes?: Array<{ pageId: string; nodeId?: string; action: "add" | "remove" | "update" | "reorder"; fields: string[] }>;
}
export interface MeasuredPage {
  sourceHash?: string;
  background?: string;
  pageId: string;
  revision: number;
  nodes: Array<Pick<DesignNode, "id" | "kind" | "parentId" | "appearance">>;
  issues?: Array<{ nodeId?: string; message: string }>;
  /** Trusted browser glyph rectangles for transparent text; never HTML-provided or persisted. */
  textPaint?: Record<string, DesignRect[]>;
}

export const cloneDesign = <T>(value: T): T => JSON.parse(JSON.stringify(value));
export const nodeText = (node: DesignNode): string => node.appearance.text?.paragraphs.map(p => p.runs.map(r => r.text).join("")).join("\n") ?? "";

export function assertDesign(document: DesignDocument): void {
  if (document?.schemaVersion !== 1) throw new Error("Unsupported design JSON schemaVersion; expected 1");
  if (document.canvas?.width !== 1280 || document.canvas.height !== 720 || document.canvas.unit !== "css-px") throw new Error("Design canvas must be 1280x720 css-px");
  if (!Number.isInteger(document.revision) || document.revision < 0) throw new Error("Invalid design revision");
  const pageIds = new Set<string>();
  for (const page of document.pages) {
    if (pageIds.has(page.id)) throw new Error(`Duplicate page ID: ${page.id}`);
    pageIds.add(page.id);
    const byId = new Map(page.nodes.map(n => [n.id, n]));
    if (byId.size !== page.nodes.length) throw new Error(`Duplicate node ID on ${page.id}`);
    for (const node of page.nodes) {
      const { x, y, w, h } = node.appearance.rect;
      if (![x, y, w, h].every(Number.isFinite) || w < 0 || h < 0) throw new Error(`Invalid geometry: ${node.id}`);
      const visited = new Set([node.id]);
      let parent = node.parentId;
      while (parent) {
        const owner = byId.get(parent);
        if (!owner || owner.kind !== "group" || visited.has(parent)) throw new Error(`Invalid group relationship: ${node.id}`);
        visited.add(parent);
        parent = owner.parentId;
      }
    }
  }
}

/** Merge only the projected appearance. HTML cannot overwrite locks, bindings, source or metadata. */
export function mergeMeasuredPage(document: DesignDocument, measured: MeasuredPage, baseline?: MeasuredPage): DesignDocument {
  assertDesign(document);
  if (measured.sourceHash && measured.sourceHash !== document.source.sha256) throw new Error("STALE_DESIGN: source document changed");
  if (measured.revision !== document.revision) throw new Error("STALE_DESIGN: read the latest design JSON/HTML before editing");
  const result = cloneDesign(document);
  const page = result.pages.find(p => p.id === measured.pageId);
  if (!page) throw new Error("Unknown design page");
  const original = new Map(page.nodes.map(n => [n.id, n]));
  const seen = new Set<string>();
  const merged: DesignNode[] = [];
  for (const candidate of measured.nodes) {
    if (seen.has(candidate.id)) throw new Error(`Duplicate HTML data-id: ${candidate.id}`);
    seen.add(candidate.id);
    const before = original.get(candidate.id);
    if (!before && !candidate.id.startsWith("new:")) throw new Error(`Unknown HTML data-id: ${candidate.id}`);
    if (before && before.kind !== candidate.kind) throw new Error(`Cannot change node kind: ${candidate.id}`);
    if (before && before.parentId !== candidate.parentId) throw new Error(`Existing group membership must be preserved: ${candidate.id}`);
    const protection = before ? effectiveLock(page, before) : undefined;
    const next = before ? { ...before, appearance: cloneDesign(candidate.appearance) } : {
      ...cloneDesign(candidate), name: candidate.id, purpose: "content" as const
    };
    if (before && protection === "all" && !sameAppearance(before.appearance, next.appearance)) throw new Error(`Locked element changed: ${before.id}`);
    if (before && protection === "content" && !sameAppearance({ ...before.appearance, rect: next.appearance.rect }, next.appearance)) throw new Error(`Asset content is locked: ${before.id}`);
    if (before?.constraints?.maxChars !== undefined && nodeText(next).length > before.constraints.maxChars) throw new Error(`Text constraint exceeded: ${before.id}`);
    // Keep the original numbers and inherited XML when browser rounding is the only change.
    if (before && sameAppearance(before.appearance, next.appearance)) next.appearance = before.appearance;
    merged.push(next);
  }
  for (const before of original.values()) {
    if (seen.has(before.id)) continue;
    if (before.hidden) { merged.push(before); continue; }
    if (effectiveLock(page, before) || before.required) throw new Error(`Required/locked element removed: ${before.id}`);
  }
  checkFixedProtection(page, { ...page, nodes: merged }, baseline?.textPaint, measured.textPaint);
  page.nodes = merged;
  if (measured.background) page.background = measured.background;
  result.revision++;
  assertDesign(result);
  return result;
}

export function sameAppearance(a: unknown, b: unknown): boolean {
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) < 0.05;
  if (a === b) return true;
  if (!a || !b || typeof a !== "object" || typeof b !== "object") return false;
  const ak = Object.keys(a), bk = Object.keys(b);
  return ak.length === bk.length && ak.every(k => Object.prototype.hasOwnProperty.call(b, k) && sameAppearance((a as any)[k], (b as any)[k]));
}
