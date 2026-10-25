import { effectiveLock } from "./template-policy";
import { cloneDesign, mergeMeasuredPage, nodeText, type DesignDocument, type DesignNode, type MeasuredPage } from "./model";
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

/** Paragraph texts of a text node that occur more than once in it. */
export function repeatedParagraphs(node: Pick<DesignNode, "appearance">): Array<{ text: string; count: number }> {
  const counts = new Map<string, number>();
  for (const paragraph of node.appearance.text?.paragraphs ?? []) {
    const text = paragraph.runs.map(run => run.text).join("").trim();
    if (text) counts.set(text, (counts.get(text) ?? 0) + 1);
  }
  return [...counts].filter(([, count]) => count > 1).map(([text, count]) => ({ text, count }));
}

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
  if (baseline.pageId !== edited.pageId) throw new Error("HTML measurements refer to different pages");
  const page = design.pages.find(p => p.id === edited.pageId);
  if (!page) throw new Error("Unknown design page");
  const original = new Map(page.nodes.map(node => [node.id, node]));
  const before = new Map(baseline.nodes.map(node => [node.id, node]));
  const changes: NonNullable<DesignDocument["changes"]> = [];
  const affected = new Set<string>();
  const nodes = edited.nodes.map(node => {
    const old = original.get(node.id), base = before.get(node.id);
    if (!old) {
      const after = nodeText(node);
      changes.push({ pageId: page.id, nodeId: node.id, action: "add", fields: [], ...(after ? { text: { after } } : {}) });
      affected.add(node.id);
      return node;
    }
    if (!base) throw new Error(`Node missing from HTML baseline: ${node.id}`);
    const fields: string[] = [];
    const appearance = patch(old.appearance, base.appearance, node.appearance, "appearance", fields);
    const merged = { ...node, appearance };
    if (fields.length) {
      const textBefore = nodeText(old), textAfter = nodeText(merged);
      changes.push({ pageId: page.id, nodeId: node.id, action: "update", fields,
        ...(textBefore !== textAfter ? { text: { before: textBefore, after: textAfter } } : {}) });
      affected.add(node.id);
    }
    return merged;
  });
  const live = new Set(nodes.map(n => n.id));
  for (const node of page.nodes) if (!node.hidden && !live.has(node.id)) {
    const textBefore = nodeText(node);
    changes.push({ pageId: page.id, nodeId: node.id, action: "remove", fields: [], ...(textBefore ? { text: { before: textBefore } } : {}) });
  }
  if (!equal(baseline.nodes.filter(n => live.has(n.id)).map(n => n.id), nodes.filter(n => before.has(n.id)).map(n => n.id))) {
    changes.push({ pageId: page.id, action: "reorder", fields: ["nodes"] });
  }
  const backgroundChanged = baseline.background !== edited.background;
  if (backgroundChanged) changes.push({ pageId: page.id, action: "update", fields: ["background"] });
  const immutable = (id:string) => { const node=original.get(id); return node && effectiveLock(page,node)==="all"; };
  // Layout findings are advisory. Identity, unsupported content and immutable edits remain errors.
  const warnings: DesignDocument["diagnostics"] = [];
  for (const issue of edited.issues ?? []) {
    if (issue.nodeId ? affected.has(issue.nodeId) && !immutable(issue.nodeId) : !baseline.issues?.some(old => old.message === issue.message)) warnings.push({...issue, pageId: page.id, severity:"warning"});
  }
  // CSS may expand a box to its padding minimum. A text-only diff must still fit
  // the preserved native geometry, rather than silently relying on that expansion.
  for (const node of nodes) if (affected.has(node.id) && !immutable(node.id) && node.appearance.text) {
    const glyphs=edited.textPaint?.[node.id]; if (!glyphs?.length) continue;
    const r=node.appearance.rect;
    const left=Math.min(...glyphs.map(g=>g.x)),top=Math.min(...glyphs.map(g=>g.y)),right=Math.max(...glyphs.map(g=>g.x+g.w)),bottom=Math.max(...glyphs.map(g=>g.y+g.h));
    if (left<r.x-1 || right>r.x+r.w+1 || top<r.y-1 || bottom>r.y+r.h+1) warnings.push({pageId:page.id,nodeId:node.id,severity:"warning",message:`Text is clipped or outside its native PPT box: ${node.id}. ${JSON.stringify({nodeId:node.id,nodeBounds:r,textBounds:{x:left,y:top,w:right-left,h:bottom-top},requiredTextHeight:bottom-top,recovery:"Layout warning only; the page can be applied. Inspect the rendered text. Overflow is relative to this text box's own edges; moving the entire box does not change it. If visibly clipped, adjust padding, width, height, line height or wording."})}`});
  }
  for (const node of nodes) if (affected.has(node.id) && !immutable(node.id)) {
    for (const { text, count } of repeatedParagraphs(node)) {
      warnings.push({pageId:page.id,nodeId:node.id,severity:"warning",message:`Repeated paragraph in ${original.get(node.id)?.name || node.id}: "${text.slice(0, 40)}" appears ${count} times in this text box. ${JSON.stringify({nodeId:node.id,text,count,recovery:"Layout warning only; the page was applied. If the repetition is not intended, keep one copy of the paragraph and submit the page again."})}`});
    }
  }
  const merged = mergeMeasuredPage(design, { ...edited, nodes, background: backgroundChanged ? edited.background : page.background }, baseline);
  if (!changes.length) { const unchanged = cloneDesign(design); delete unchanged.changes; return unchanged; }
  merged.changes = changes;
  merged.diagnostics.push(...warnings);
  return merged;
}
