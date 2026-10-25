import type { DesignDocument, DesignNode, DesignPage, DesignRect } from "./model";
import { parseBinding } from "../api/template";

export type TemplatePurpose = "fixed" | "content" | "reference";
export interface TemplatePageSemantics { role?: string; layoutName?: string; note?: string }
export interface TemplateNodeSemantics { purpose: TemplatePurpose; note?: string; name?: string }
export interface TemplateRecord { version: 1; page?: TemplatePageSemantics; nodes: Record<string, TemplateNodeSemantics> }
export const TEMPLATE_EXT_URI = "{70D18A51-144B-49B9-A917-03139173DE67}";
export const TEMPLATE_NS = "urn:pptx-editor:template:1";
const purposes = ["fixed", "content", "reference"];
const validText = (value: unknown): value is string | undefined => value === undefined || typeof value === "string" && value.length <= 4096;
export function validateTemplateRecord(value: unknown): asserts value is TemplateRecord {
  const r = value as TemplateRecord;
  if (!r || r.version !== 1 || !r.nodes || Array.isArray(r.nodes) || typeof r.nodes !== "object" || Object.keys(r).some(k => !["version", "page", "nodes"].includes(k))) throw new Error("Invalid template policy record");
  if (r.page !== undefined && (!r.page || typeof r.page !== "object" || Array.isArray(r.page) || Object.keys(r.page).some(k => !["role", "layoutName", "note"].includes(k)) || !Object.values(r.page).every(validText))) throw new Error("Invalid template page semantics");
  for (const [id, node] of Object.entries(r.nodes)) if (!/^\d+$/.test(id) || !node || !purposes.includes(node.purpose) || Object.keys(node).some(k => !["purpose", "note", "name"].includes(k)) || !validText(node.note) || !validText(node.name)) throw new Error(`Invalid template node policy: ${id}`);
}
export function readTemplateRecord(xml: Document): TemplateRecord | undefined {
  const found = Array.from(xml.getElementsByTagNameNS(TEMPLATE_NS, "template"));
  if (found.length > 1) throw new Error("Duplicate template policy record");
  if (!found.length) return undefined;
  const owner = found[0].parentElement;
  if (owner?.localName !== "ext" || owner.getAttribute("uri") !== TEMPLATE_EXT_URI || owner.parentElement?.localName !== "extLst" || owner.parentElement.parentElement !== xml.documentElement) throw new Error("Invalid template policy extension location");
  const record = JSON.parse(found[0].textContent ?? ""); validateTemplateRecord(record);
  const ids = new Set(Array.from(xml.getElementsByTagNameNS("*", "cNvPr")).map(n => n.getAttribute("id")));
  for (const id of Object.keys(record.nodes)) if (!ids.has(id)) throw new Error(`Template policy refers to missing shape: ${id}`);
  return record;
}
export function writeTemplateRecord(xml: Document, record: TemplateRecord): void {
  validateTemplateRecord(record);
  const root = xml.documentElement, P = root.namespaceURI!;
  let list = Array.from(root.children).find(n => n.localName === "extLst");
  if (!list) { list = xml.createElementNS(P, "p:extLst"); root.appendChild(list); }
  let ext = Array.from(list.children).find(n => n.getAttribute("uri") === TEMPLATE_EXT_URI);
  if (!ext) { ext = xml.createElementNS(P, "p:ext"); ext.setAttribute("uri", TEMPLATE_EXT_URI); list.appendChild(ext); }
  const content = xml.createElementNS(TEMPLATE_NS, "design:template");
  content.textContent = JSON.stringify(record); ext.replaceChildren(content);
}
/** Only explicit annotations relax conservative asset/brand defaults. */
export function annotatedPurpose(name: string): TemplatePurpose | undefined {
  const purpose = name.match(/(?:^|[|｜])(?:用途|purpose)=(固定|可变内容|可变|参考素材|参考图|fixed|content|reference)(?:[|｜]|$)/)?.[1];
  const main = name.replace(/^@/, "").split(/[|｜.]/)[0];
  const v = purpose ?? (["固定", "可变内容", "参考素材", "fixed", "content", "reference"].includes(main) && name.startsWith("@") ? main : undefined);
  if (v) return ["固定", "fixed"].includes(v) ? "fixed" : ["参考素材", "参考图", "reference"].includes(v) ? "reference" : "content";
  const binding = parseBinding(name);
  if (binding?.reference) return "reference";
  if (binding && (binding.field || binding.token || binding.area || binding.group)) return "content";
}
const brandName = (name:string):boolean => /logo|品牌|标识|商标|页眉|页脚|背景|brand|watermark|background/i.test(name);
export function defaultPurpose(node: DesignNode): TemplatePurpose {
  const explicit = annotatedPurpose(node.name);
  if (explicit) return explicit;
  if (node.metadata?.backdrop || brandName(node.name)) return "fixed";
  if (node.kind === "image" || node.kind === "asset") return "fixed";
  if (node.kind === "shape" && node.appearance.rect.w >= 1200 && node.appearance.rect.h >= 650) return "fixed";
  return "content";
}
export function effectivePurpose(page: DesignPage, node: DesignNode): TemplatePurpose {
  let parent = node.parentId;
  while (parent) {
    const owner = page.nodes.find(n => n.id === parent);
    if (!owner) break;
    if (owner.purpose === "fixed") return "fixed";
    parent = owner.parentId;
  }
  return node.purpose ?? defaultPurpose(node);
}
export function effectiveLock(page: DesignPage, node: DesignNode): "all" | "content" | undefined {
  return effectivePurpose(page, node) === "fixed" ? "all" : node.lock;
}
const overlap = (a: DesignRect, b: DesignRect): number => Math.max(0, Math.min(a.x+a.w,b.x+b.w)-Math.max(a.x,b.x))*Math.max(0,Math.min(a.y+a.h,b.y+b.h)-Math.max(a.y,b.y));
const visiblePaint = (node: DesignNode): boolean => node.kind !== "group" && !!(node.appearance.src || node.appearance.text || node.appearance.border || node.appearance.fill && !["transparent", "rgba(0, 0, 0, 0)"].includes(node.appearance.fill));
const paintRects = (node: DesignNode, textPaint?: Record<string,DesignRect[]>): DesignRect[] => {
  const a = node.appearance;
  return a.text && !a.src && !a.border && (!a.fill || ["transparent","rgba(0, 0, 0, 0)"].includes(a.fill)) && textPaint?.[node.id] !== undefined ? textPaint[node.id] : [a.rect];
};
/** Use trusted glyph extents for transparent text, complete bounds for painted boxes and assets. */
export function checkFixedProtection(before: DesignPage, after: DesignPage, beforeTextPaint?: Record<string,DesignRect[]>, afterTextPaint?: Record<string,DesignRect[]>): Array<{nodeId?: string; message: string}> {
  const warnings: Array<{nodeId?: string; message: string}> = [];
  const beforeIndex = new Map(before.nodes.map((n,i) => [n.id,i]));
  const afterIndex = new Map(after.nodes.map((n,i) => [n.id,i]));
  const oldById = new Map(before.nodes.map(n => [n.id,n]));
  for (const fixed of before.nodes.filter(n => effectivePurpose(before,n) === "fixed" && !n.hidden)) {
    const next = after.nodes.find(n => n.id === fixed.id);
    if (!next) throw new Error(`Fixed element removed: ${fixed.id}`);
    // An ancestor's reordering can change paint order even if each object's attributes remain unchanged.
    for (const n of before.nodes) if (n.id !== fixed.id && afterIndex.has(n.id) && ((beforeIndex.get(n.id)! < beforeIndex.get(fixed.id)!) !== (afterIndex.get(n.id)! < afterIndex.get(fixed.id)!))) {
      if (overlap(n.appearance.rect, fixed.appearance.rect) > 0.5) warnings.push({nodeId:n.id,message:`Fixed element layer order changed: ${fixed.id}. Layout warning: inspect the rendered page and adjust the mutable foreground if it hides the brand.`});
    }
    const areas = fixed.metadata?.protectedRegions as DesignRect[] | undefined;
    const r = fixed.appearance.rect;
    const regions = areas ?? (r.w >= 1200 && r.h >= 650 ? [] : [r]);
    for (const foreground of after.nodes) {
      if (foreground.id === fixed.id || !visiblePaint(foreground) || afterIndex.get(foreground.id)! < afterIndex.get(fixed.id)!) continue;
      const old = oldById.get(foreground.id);
      for (const region of regions) {
        const area = paintRects(foreground,afterTextPaint).reduce((sum,r)=>sum+overlap(region,r),0);
        const oldArea = old && beforeIndex.get(old.id)! > beforeIndex.get(fixed.id)! ? paintRects(old,beforeTextPaint).reduce((sum,r)=>sum+overlap(region,r),0) : 0;
        const paintChanged = !!old && (old.appearance.src !== foreground.appearance.src || old.appearance.fill !== foreground.appearance.fill || JSON.stringify(old.appearance.border) !== JSON.stringify(foreground.appearance.border));
        if (area > 0.5 && (area > oldArea + 0.5 || !old || paintChanged)) warnings.push({nodeId:foreground.id,message:`FIXED_OCCLUSION: ${foreground.id} covers fixed element ${fixed.id}. ${JSON.stringify({nodeId:foreground.id,nodeBounds:foreground.appearance.rect,fixedNodeId:fixed.id,protectedBounds:region,overlapArea:Math.round(area*100)/100,canvas:{width:1280,height:720},recovery:"Layout warning only; this does not block applying the page. Inspect the actual rendering before deciding whether to move or reflow mutable foreground content. Retain the fixed object's content and metadata."})}`});
      }
    }
  }
  return warnings;
}
export function applyDefaultPolicies(document: DesignDocument): void {
  // Repeated identical top-of-page words are commonly brand headers. Do not infer permission from position.
  const repeated = new Map<string, Set<string>>();
  for (const page of document.pages) for (const n of page.nodes) {
    const text = n.appearance.text?.paragraphs.map(p=>p.runs.map(r=>r.text).join("")).join(" ").trim();
    if (text && text.length <= 80 && n.appearance.rect.y < 100 && !annotatedPurpose(n.name)) {
      if (!repeated.has(text)) repeated.set(text,new Set()); repeated.get(text)!.add(page.id);
    }
  }
  for (const page of document.pages) for (const node of page.nodes) {
    if (node.purpose) continue;
    const parent = page.nodes.find(n=>n.id===node.parentId);
    const text = node.appearance.text?.paragraphs.map(p=>p.runs.map(r=>r.text).join("")).join(" ").trim();
    node.purpose = annotatedPurpose(node.name) ?? (brandName(node.name) ? "fixed" : parent?.purpose === "reference" || parent?.purpose === "content" && annotatedPurpose(parent.name) ? parent.purpose
      : text && (repeated.get(text)?.size ?? 0) >= 2 ? "fixed" : defaultPurpose(node));
  }
}
