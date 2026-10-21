import { effectiveLock } from "./template-policy";
import { cloneDesign, mergeMeasuredPage, type DesignDocument, type MeasuredPage } from "./model";
import { assertDesignHtmlIdentity, parseDesignHtml } from "./html-document";
import { projectPage } from "./projection";

/** The baseline is supplied by the host's cached projection, never by an AI tool argument. */
export function baselineFor(design: DesignDocument, editedHtml: string, baselineHtml?: string): string {
  const edited = parseDesignHtml(editedHtml).body;
  const pageId = edited.dataset.pageId;
  assertDesignHtmlIdentity(design,edited);
  const baseline = baselineHtml ?? projectPage(design, pageId!);
  const body = parseDesignHtml(baseline).body;
  assertDesignHtmlIdentity(design,body,pageId);
  return baseline;
}

const equal = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b);
const copy = (value: any): any => value === undefined ? undefined : cloneDesign(value);

/** Compare two browser measurements, not a browser measurement against native PPT values. */
function patch(original: any, baseline: any, edited: any, path: string, fields: string[]): any {
  if (equal(baseline, edited)) return copy(original);
  if (Array.isArray(baseline) && Array.isArray(edited) && Array.isArray(original)
    && baseline.length === edited.length && original.length === baseline.length) {
    return edited.map((value, i) => patch(original[i], baseline[i], value, `${path}.${i}`, fields));
  }
  if (baseline && edited && typeof baseline === "object" && typeof edited === "object" && !Array.isArray(baseline) && !Array.isArray(edited)) {
    const out = copy(original) ?? {};
    for (const key of new Set([...Object.keys(baseline), ...Object.keys(edited)])) {
      const value = patch(original?.[key], baseline[key], edited[key], path ? `${path}.${key}` : key, fields);
      if (value === undefined) delete out[key]; else out[key] = value;
    }
    return out;
  }
  fields.push(path);
  return copy(edited);
}

/** Preserve every native field whose rendered HTML value did not change. */
export function mergeHtmlDifference(design: DesignDocument, baseline: MeasuredPage, edited: MeasuredPage): DesignDocument {
  if (baseline.pageId !== edited.pageId || baseline.sourceHash !== edited.sourceHash || baseline.revision !== edited.revision) throw new Error("STALE_DESIGN: HTML measurements use different baselines");
  const page = design.pages.find(p => p.id === edited.pageId);
  if (!page) throw new Error("Unknown design page");
  const original = new Map(page.nodes.map(node => [node.id, node]));
  const before = new Map(baseline.nodes.map(node => [node.id, node]));
  const changes: NonNullable<DesignDocument["changes"]> = [];
  const affected = new Set<string>();
  const nodes = edited.nodes.map(node => {
    const old = original.get(node.id), base = before.get(node.id);
    if (!old) {
      changes.push({ pageId: page.id, nodeId: node.id, action: "add", fields: [] });
      affected.add(node.id);
      return node;
    }
    if (!base) throw new Error(`Node missing from HTML baseline: ${node.id}`);
    const fields: string[] = [];
    const appearance = patch(old.appearance, base.appearance, node.appearance, "appearance", fields);
    if (fields.length) {
      changes.push({ pageId: page.id, nodeId: node.id, action: "update", fields });
      affected.add(node.id);
    }
    return { ...node, appearance };
  });
  const live = new Set(nodes.map(n => n.id));
  for (const node of page.nodes) if (!node.hidden && !live.has(node.id)) changes.push({ pageId: page.id, nodeId: node.id, action: "remove", fields: [] });
  if (!equal(baseline.nodes.filter(n => live.has(n.id)).map(n => n.id), nodes.filter(n => before.has(n.id)).map(n => n.id))) {
    changes.push({ pageId: page.id, action: "reorder", fields: ["nodes"] });
  }
  const backgroundChanged = baseline.background !== edited.background;
  if (backgroundChanged) changes.push({ pageId: page.id, action: "update", fields: ["background"] });
  const immutable = (id:string) => { const node=original.get(id); return node && effectiveLock(page,node)==="all"; };
  // Immutable edits are rejected by mergeMeasuredPage itself, before reporting layout repairs.
  // Existing unrelated overflow does not turn a one-word edit into a whole-page repair.
  for (const issue of edited.issues ?? []) {
    if (issue.nodeId ? affected.has(issue.nodeId) && !immutable(issue.nodeId) : !baseline.issues?.some(old => old.message === issue.message)) throw new Error(issue.message);
  }
  // CSS may expand a box to its padding minimum. A text-only diff must still fit
  // the preserved native geometry, rather than silently relying on that expansion.
  for (const node of nodes) if (affected.has(node.id) && !immutable(node.id) && node.appearance.text) {
    const glyphs=edited.textPaint?.[node.id]; if (!glyphs?.length) continue;
    const r=node.appearance.rect;
    const left=Math.min(...glyphs.map(g=>g.x)),top=Math.min(...glyphs.map(g=>g.y)),right=Math.max(...glyphs.map(g=>g.x+g.w)),bottom=Math.max(...glyphs.map(g=>g.y+g.h));
    if (left<r.x-0.5 || right>r.x+r.w+0.5 || top<r.y-1 || bottom>r.y+r.h+1) throw new Error(`Text is clipped or outside its native PPT box: ${node.id}. ${JSON.stringify({nodeId:node.id,nodeBounds:r,textBounds:{x:left,y:top,w:right-left,h:bottom-top},requiredTextHeight:bottom-top,recovery:"The native PPT box is smaller than the browser box (for example padding exceeds its height). Explicitly correct this mutable box or padding in HTML, preserve fixed assets and metadata, then retry apply_html. No content was applied."})}`);
  }
  const merged = mergeMeasuredPage(design, { ...edited, nodes, background: backgroundChanged ? edited.background : page.background }, baseline);
  if (!changes.length) { const unchanged = cloneDesign(design); delete unchanged.changes; return unchanged; }
  merged.changes = changes;
  return merged;
}
