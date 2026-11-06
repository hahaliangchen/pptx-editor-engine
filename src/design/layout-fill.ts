import { parseDesignHtml } from "./html-document";

/**
 * Layout fill: on a page the author marked `fill: "layout"`, the AI sends content only and the host
 * edits the page's own HTML. Blocks are groups whose template slot is "<block>#<set>"; siblings with
 * the same slot form a repeatable set, "sep#<set>" marks the separators between them, and every other
 * slot names a field inside its nearest block ("title", "body", "icon" …). Existing blocks keep their
 * identity and markup; extra items are copies of a block, so styles are copied, never redrawn.
 */
export interface FillItem { fields?: Record<string, string | string[]>; icons?: Record<string, string>; sets?: FillSet[] }
export interface FillSet { set: string; columns?: number; items: FillItem[] }
export interface FillRequest { texts?: Record<string, string | string[]>; sets?: FillSet[] }
export interface LayoutItem { fields: Record<string, string[]>; icons: string[]; sets: LayoutSet[] }
export interface LayoutSet { set: string; block: string; arrangement: Arrangement; columns: number; items: LayoutItem[] }
export interface LayoutSummary { texts: Array<{ id: string; text: string[] }>; sets: LayoutSet[] }
type Arrangement = "single" | "row" | "column" | "grid";
interface Rect { x: number; y: number; w: number; h: number }

const BLOCK = /^(.+)#(\d+)$/;
const slotOf = (el: Element) => (el as HTMLElement).dataset?.slot ?? "";
const isBlock = (el: Element) => (el as HTMLElement).dataset?.kind === "group" && BLOCK.test(slotOf(el)) && !slotOf(el).startsWith("sep#");
const blockName = (slot: string) => slot.replace(BLOCK, "$1");
const blockOf = (el: Element): Element | null => {
  for (let p = el.parentElement; p; p = p.parentElement) if (isBlock(p)) return p;
  return null;
};
const px = (el: HTMLElement, key: "left" | "top" | "width" | "height") => parseFloat(el.style[key]) || 0;
const rectOf = (el: Element): Rect => { const e = el as HTMLElement; return { x: px(e, "left"), y: px(e, "top"), w: px(e, "width"), h: px(e, "height") }; };
const round = (n: number) => `${Math.round(n * 100) / 100}px`;
const place = (el: Element, r: Rect) => { const s = (el as HTMLElement).style; s.left = round(r.x); s.top = round(r.y); s.width = round(r.w); s.height = round(r.h); };
const paragraphs = (el: Element) => Array.from(el.children).filter(c => c.localName === "p" || /^h[1-6]$/.test(c.localName));
const short = (text: string) => text.length > 40 ? `${text.slice(0, 40)}…` : text;

/** Sets whose nearest block is `scope` (null: the page), keyed by slot, in DOM order. */
function setsIn(root: ParentNode, scope: Element | null): Map<string, Element[]> {
  const sets = new Map<string, Element[]>();
  for (const el of Array.from(root.querySelectorAll("[data-slot]"))) {
    if (!isBlock(el) || blockOf(el) !== scope) continue;
    sets.set(slotOf(el), [...(sets.get(slotOf(el)) ?? []), el]);
  }
  return sets;
}
const fieldsIn = (block: Element) => Array.from(block.querySelectorAll("[data-slot]"))
  .filter(el => !BLOCK.test(slotOf(el)) && blockOf(el) === block);

function arrange(rects: Rect[]): { arrangement: Arrangement; columns: number } {
  if (rects.length < 2) return { arrangement: "single", columns: 1 };
  const distinct = (values: number[]) => values.filter((v, i) => values.findIndex(w => Math.abs(w - v) < 3) === i).length;
  const columns = distinct(rects.map(r => r.x)), rows = distinct(rects.map(r => r.y));
  if (rows === 1) return { arrangement: "row", columns };
  if (columns === 1) return { arrangement: "column", columns: 1 };
  return { arrangement: "grid", columns };
}

export function describeLayout(html: string): LayoutSummary {
  const doc = parseDesignHtml(html);
  const describe = (scope: Element | null): LayoutSet[] => Array.from(setsIn(scope ?? doc.body, scope)).map(([set, instances]) => ({
    set, block: blockName(set), ...arrange(instances.map(rectOf)),
    items: instances.map(item => {
      const fields: Record<string, string[]> = {}, icons: string[] = [];
      for (const field of fieldsIn(item)) {
        if ((field as HTMLElement).dataset.kind === "text") fields[slotOf(field)] = paragraphs(field).map(p => short(p.textContent ?? ""));
        else if (slotOf(field).startsWith("icon")) icons.push(slotOf(field));
      }
      return { fields, icons, sets: describe(item) };
    })
  }));
  const texts = Array.from(doc.body.querySelectorAll<HTMLElement>('[data-kind="text"][data-id]'))
    .filter(el => !el.dataset.slot && !blockOf(el) && el.dataset.purpose === "content")
    .map(el => ({ id: el.dataset.id!, text: paragraphs(el).map(p => short(p.textContent ?? "")) }));
  return { texts, sets: describe(null) };
}

/** Keep a leading list marker run ("■", "1.") and put the new words in the run after it. */
function setParagraph(p: Element, text: string): void {
  const runs = Array.from(p.querySelectorAll("span"));
  if (!runs.length) { p.textContent = text; return; }
  const marker = runs.length > 1 && /^([■●•▪◆»–-]|\d+[.、)])$/.test((runs[0].textContent ?? "").replace(/ /g, "").trim());
  const target = runs[marker ? 1 : 0];
  const lead = marker && /^ /.test(target.textContent ?? "") ? " " : "";
  // The template's marker run stays, so a marker typed into the content would show twice.
  target.textContent = lead + (marker ? text.replace(/^\s*([■●•▪◆»–-]|\d+[.、)])[\s ]*/, "") : text);
  for (const run of runs.slice(marker ? 2 : 1)) run.remove();
  for (const node of Array.from(p.childNodes)) if (node.nodeType === 3) node.remove();
}

function fillText(el: Element, value: string | string[]): void {
  const values = Array.isArray(value) ? value : [value];
  const protos = paragraphs(el);
  if (!protos.length) { el.textContent = values.join("\n"); return; }
  // Frame-positioned paragraphs cannot be re-flowed; fill them in place.
  const framed = protos.some(p => (p as HTMLElement).style.position === "absolute");
  if (framed) {
    protos.forEach((p, i) => i < values.length ? setParagraph(p, values[i]) : p.remove());
    return;
  }
  const filled = values.map((text, i) => {
    const p = protos[Math.min(i, protos.length - 1)].cloneNode(true) as Element;
    setParagraph(p, text);
    return p;
  });
  protos[0].before(...filled);
  for (const p of protos) p.remove();
  if (values.length === 1) fitLine(el as HTMLElement, values[0]);
}

/** Rough single-line width: CJK, full-width glyphs and the dashes, quotes and ellipsis CJK fonts draw full width take 1em, the rest about 0.55em. */
const lineWidth = (text: string, size: number) => Array.from(text).reduce((w, ch) => w + (/[\u2e80-\uffef\u2010-\u2027]/.test(ch) ? size : size * 0.55), 0);
/** Rough height of a text box's paragraphs wrapped at `width`; undefined for frame-positioned paragraphs. */
function textHeight(el: HTMLElement, width: number): number | undefined {
  const ps = paragraphs(el) as HTMLElement[];
  if (!ps.length || ps.some(p => p.style.position === "absolute")) return;
  const pad = (parseFloat(el.style.paddingLeft) || 0) + (parseFloat(el.style.paddingRight) || 0);
  const vertical = (parseFloat(el.style.paddingTop) || 0) + (parseFloat(el.style.paddingBottom) || 0);
  const room = Math.max(1, width - pad);
  return ps.reduce((h, p) => {
    const size = Math.max(parseFloat(p.style.fontSize) || 0, ...Array.from(p.querySelectorAll<HTMLElement>("span")).map(s => parseFloat(s.style.fontSize) || 0)) || 16;
    const line = parseFloat(p.style.lineHeight) || size * 1.3;
    return h + Math.max(1, Math.ceil(lineWidth((p.textContent ?? "").replace(/\u00a0/g, " "), size) / room)) * line;
  }, vertical + 2);
}
/** Page margin a loose one-line title may widen into. */
const PAGE_MARGIN = 48;
const pageWidth = (el: Element) => parseFloat(el.ownerDocument.head?.textContent?.match(/body\{[^}]*?width:([\d.]+)px/)?.[1] ?? "") || 1280;
/** How far to shrink a line to fit `width`: the estimate is rough, so aim 4% short, and never below three quarters. */
const fontFactor = (width: number, needed: number, pad: number) => Math.max(.75, (width - pad - 4) / (needed - pad - 4) * .96);
/** Scale a one-line text's font sizes and line heights; the box keeps its place. */
function shrinkFont(el: HTMLElement, factor: number): void {
  for (const node of [...paragraphs(el), ...Array.from(el.querySelectorAll("span"))] as HTMLElement[]) {
    for (const key of ["fontSize", "lineHeight"] as const) {
      const value = parseFloat(node.style[key]);
      if (value) node.style[key] = round(value * factor);
    }
  }
}
/**
 * A one-line label (a scene name, a ribbon title, a cover title) keeps the sample's box width, which wraps or clips
 * longer words. Widen it within its block, or within the page margins when it is not in a block, around its centre
 * when centred, and widen a shape that frames it by the same amount. When even that is too narrow, shrink the font
 * (to three quarters at most) so the line stays one line.
 */
function fitLine(el: HTMLElement, text: string): void {
  const parent = el.parentElement as HTMLElement | null;
  if (!parent) return;
  const inBlock = parent.dataset.kind === "group";
  const p = paragraphs(el)[0] as HTMLElement | undefined, run = p?.querySelector<HTMLElement>("span");
  const size = parseFloat(run?.style.fontSize || p?.style.fontSize || "") || 16;
  const line = parseFloat(p?.style.lineHeight || "") || size * 1.3;
  const r = rectOf(el);
  if (r.h > line * 1.6) return;
  const pad = (parseFloat(el.style.paddingLeft) || 0) + (parseFloat(el.style.paddingRight) || 0);
  const needed = lineWidth(text, size) + pad + 4;
  if (needed <= r.w) return;
  if (!inBlock) {
    // Only into free space on its own line: a page title must not run under the brand logo beside it.
    const besides = (Array.from(parent.children) as HTMLElement[]).filter(n => n !== el && n.dataset.kind)
      .map(rectOf).filter(n => n.w < pageWidth(el) * .9 && n.y < r.y + r.h && n.y + n.h > r.y);
    const lo = Math.max(PAGE_MARGIN, ...besides.filter(n => n.x + n.w <= r.x + 1).map(n => n.x + n.w + 8));
    const hi = Math.min(pageWidth(el) - PAGE_MARGIN, ...besides.filter(n => n.x >= r.x + r.w - 1).map(n => n.x - 8));
    const centred = p?.style.textAlign === "center";
    const max = Math.max(r.w, centred ? hi - lo : hi - r.x);
    const w = Math.min(needed, max);
    const x = centred ? Math.min(Math.max(lo, r.x + r.w / 2 - w / 2), hi - w) : r.x;
    place(el, { ...r, x, w });
    if (needed > w) shrinkFont(el, fontFactor(w, needed, pad));
    return;
  }
  // A block keeps its size here (a row of small blocks spreads in spreadRow instead), so the blocks of one set
  // stay alike. Elements beside the text on its line (a divider, a list) bound it; a frame behind it (a ribbon,
  // a title bar) widens with it but never past the block, so repeated bars keep their common edge.
  const group = rectOf(parent);
  const inside = (Array.from(parent.children) as HTMLElement[]).filter(n => n !== el && n.dataset.kind);
  const onLine = inside.map(n => ({ n, s: rectOf(n) })).filter(({ s }) => s.y < r.y + r.h && s.y + s.h > r.y);
  const frames = onLine.filter(({ n, s }) => n.dataset.kind === "shape" && s.x <= r.x + 1 && s.x + s.w >= r.x + r.w - 1 && s.y <= r.y + 1 && s.y + s.h >= r.y + r.h - 1);
  const hi = Math.min(group.w, ...onLine.filter(({ s }) => s.x >= r.x + r.w - 1).map(({ s }) => s.x - 6));
  const lo = Math.max(0, ...onLine.filter(({ s }) => s.x + s.w <= r.x + 1).map(({ s }) => s.x + s.w + 6));
  const centred = Math.abs(r.x + r.w / 2 - group.w / 2) <= 3;
  const frameRoom = Math.min(Infinity, ...frames.map(({ s }) => centred ? 2 * Math.min(s.x, group.w - s.x - s.w) : group.w - s.x - s.w));
  const room = centred ? Math.min(group.w, 2 * Math.min(r.x + r.w / 2 - lo, hi - r.x - r.w / 2)) : hi - r.x;
  const w = Math.min(needed, Math.max(r.w, Math.min(room, r.w + frameRoom))), delta = w - r.w;
  if (needed > Math.max(w, r.w)) shrinkFont(el, fontFactor(Math.max(w, r.w), needed, pad));
  if (delta <= 0) return;
  for (const { n, s } of frames) place(n, { ...s, x: centred ? s.x - delta / 2 : s.x, w: s.w + delta });
  place(el, { ...r, x: centred ? Math.max(0, r.x - delta / 2) : r.x, w });
}

const hex = (color: string) => {
  const rgb = color.match(/^rgba?\(\s*([\d.]+)[, ]+([\d.]+)[, ]+([\d.]+)/);
  return rgb ? `#${rgb.slice(1, 4).map(n => Math.round(Number(n)).toString(16).padStart(2, "0")).join("").toUpperCase()}` : color;
};
/** Every paint of the icon becomes `color`; "none", gradients and opacities stay, so outlines and tints keep their shape. */
function tint(art: Element, color: string): void {
  for (const el of [art, ...Array.from(art.querySelectorAll("*"))]) {
    const style = (el as SVGElement).style;
    for (const key of ["fill", "stroke", "stop-color"]) {
      const attr = el.getAttribute(key), inline = style?.getPropertyValue(key);
      if (attr && attr !== "none" && !attr.startsWith("url(")) el.setAttribute(key, color);
      if (inline && inline !== "none" && !inline.startsWith("url(")) style.setProperty(key, color);
    }
  }
  // An unpainted shape fills black by default; the root's fill makes it the icon color instead.
  if (!art.getAttribute("fill")) art.setAttribute("fill", color);
}

/**
 * The icon slot keeps its tile; the dashed placeholder ring is replaced by the SVG at the same box.
 * The ring's color is the one the author chose for this tile (red on light, white on red), so the icon takes it.
 */
function fillIcon(group: Element, svg: string): void {
  const shapes = Array.from(group.children).filter(c => (c as HTMLElement).dataset.kind === "shape") as HTMLElement[];
  const dashed = shapes.find(s => /dashed|dotted/.test(s.style.border + s.style.borderStyle));
  const ring = dashed ?? shapes[shapes.length - 1];
  const box = ring ? rectOf(ring) : { x: 0, y: 0, w: px(group as HTMLElement, "width"), h: px(group as HTMLElement, "height") };
  const holder = group.ownerDocument.createElement("div");
  holder.innerHTML = svg.trim();
  const art = holder.querySelector("svg");
  if (!art) throw new Error("FILL_LAYOUT: icons must be inline <svg> markup");
  const ink = dashed && (dashed.style.borderTopColor || dashed.style.borderColor);
  if (ink) tint(art, hex(ink));
  art.removeAttribute("width"); art.removeAttribute("height");
  art.setAttribute("data-kind", "image");
  art.setAttribute("style", `position:absolute;left:${round(box.x)};top:${round(box.y)};width:${round(box.w)};height:${round(box.h)}`);
  if (ring) ring.replaceWith(art); else group.appendChild(art);
}

/** Constraint-style resize: wide children keep both margins, centred ones stay centred, others keep their anchor edge. */
function axis(pos: number, size: number, total0: number, total1: number): [number, number] {
  if (Math.abs(total0 - total1) < 0.01) return [pos, size];
  const end0 = total0 - pos - size;
  if (size >= total0 * 0.4) return [pos, Math.max(1, total1 - pos - end0)];
  if (Math.abs(pos + size / 2 - total0 / 2) <= 2) return [(total1 - size) / 2, size];
  return end0 < pos ? [total1 - end0 - size, size] : [pos, size];
}
/** Icons keep their proportions: they move like other children and only shrink, uniformly, when the block no longer fits them. */
function iconAxis(pos: number, size: number, scaled: number, total0: number, total1: number): number {
  const end0 = total0 - pos - size;
  if (Math.abs(pos + size / 2 - total0 / 2) <= 2) return (total1 - scaled) / 2;
  return end0 < pos ? total1 - end0 - scaled : pos;
}
function scale(el: Element, factor: number): void {
  for (const child of Array.from(el.children)) {
    if (!(child as HTMLElement).dataset?.kind && child.localName !== "svg") continue;
    const r = rectOf(child);
    place(child, { x: r.x * factor, y: r.y * factor, w: r.w * factor, h: r.h * factor });
    scale(child, factor);
  }
}
function resize(el: Element, next: Rect): void {
  const old = rectOf(el);
  place(el, next);
  if ((el as HTMLElement).dataset.kind !== "group") return;
  for (const child of Array.from(el.children)) {
    if (!(child as HTMLElement).dataset?.kind) continue;
    const r = rectOf(child);
    if (slotOf(child).startsWith("icon")) {
      // An icon anchored to an edge (see iconAxis) keeps only that offset; a centred one keeps both margins while
      // it can, then lets them shrink so the icon still takes four fifths of the block.
      const room = (pos: number, size: number, t0: number, t1: number) => {
        const end0 = t0 - pos - size, centred = Math.abs(pos + size / 2 - t0 / 2) <= 2;
        return Math.max(1, centred ? Math.max(t1 - pos - end0, t1 * .8) : t1 - (end0 < pos ? end0 : pos)) / size;
      };
      const factor = Math.min(1, room(r.x, r.w, old.w, next.w), room(r.y, r.h, old.h, next.h));
      const w = r.w * factor, h = r.h * factor;
      place(child, { x: iconAxis(r.x, r.w, w, old.w, next.w), y: iconAxis(r.y, r.h, h, old.h, next.h), w, h });
      if (factor < 1) scale(child, factor);
      continue;
    }
    const [x, w] = axis(r.x, r.w, old.w, next.w);
    // Text takes the height its (already filled) content needs at its new width and keeps its vertical anchor,
    // unless the block became too short for it. The sample's box height would push short content out of place.
    const text = (child as HTMLElement).dataset.kind === "text";
    const h = text ? Math.min(textHeight(child as HTMLElement, w) ?? r.h, Math.max(1, next.h - 4)) : axis(r.y, r.h, old.h, next.h)[1];
    // A text box's sample height says little about where its words sit: one that starts in the upper half keeps
    // its top (a list under a header), one centred stays centred, others keep their distance to the bottom.
    const textY = () => {
      if (Math.abs(r.y + r.h / 2 - old.h / 2) <= 2) return (next.h - h) / 2;
      return r.y < old.h / 2 ? r.y : next.h - (old.h - r.y - r.h) - h;
    };
    const margin = Math.min(2, Math.max(0, old.h - r.y - r.h));
    const y = text ? Math.min(Math.max(2, textY()), next.h - h - margin) : axis(r.y, r.h, old.h, next.h)[0];
    resize(child, { x, y, w, h });
  }
}

function layout(rects: Rect[], count: number, columns?: number): Rect[] {
  const region = rects.reduce((u, r) => ({ x: Math.min(u.x, r.x), y: Math.min(u.y, r.y), w: Math.max(u.x + u.w, r.x + r.w) - Math.min(u.x, r.x), h: Math.max(u.y + u.h, r.y + r.h) - Math.min(u.y, r.y) }));
  const shape = arrange(rects), proto = rects[0];
  const gapAlong = (key: "x" | "y", size: "w" | "h") => {
    const sorted = [...rects].sort((a, b) => a[key] - b[key]);
    const gaps = sorted.slice(1).map((r, i) => r[key] - (sorted[i][key] + sorted[i][size])).filter(g => g >= 0);
    return gaps.length ? gaps.reduce((a, b) => a + b, 0) / gaps.length : 12;
  };
  const arrangement = shape.arrangement === "single" ? (region.w >= region.h ? "row" : "column") : shape.arrangement;
  if (arrangement === "row" && !columns) {
    const gap = gapAlong("x", "w"), w = (region.w - gap * (count - 1)) / count;
    return Array.from({ length: count }, (_, i) => ({ x: region.x + i * (w + gap), y: proto.y, w, h: proto.h }));
  }
  if (arrangement === "column" && !columns) {
    const gap = gapAlong("y", "h"), h = (region.h - gap * (count - 1)) / count;
    return Array.from({ length: count }, (_, i) => ({ x: proto.x, y: region.y + i * (h + gap), w: proto.w, h }));
  }
  const cols = Math.max(1, Math.min(columns ?? Math.min(shape.columns, Math.ceil(Math.sqrt(count))), count));
  const rows = Math.ceil(count / cols);
  const gx = shape.arrangement === "column" ? 12 : gapAlong("x", "w"), gy = shape.arrangement === "row" ? 12 : gapAlong("y", "h");
  const w = (region.w - gx * (cols - 1)) / cols, h = (region.h - gy * (rows - 1)) / rows;
  return Array.from({ length: count }, (_, i) => ({ x: region.x + (i % cols) * (w + gx), y: region.y + Math.floor(i / cols) * (h + gy), w, h }));
}

/** The width a block needs so each of its one-line labels fits on one line at its own font size. */
function blockNeed(block: Element, item: FillItem): number {
  const b = rectOf(block);
  let need = b.w;
  for (const field of fieldsIn(block) as HTMLElement[]) {
    const value = item.fields?.[slotOf(field)];
    const text = Array.isArray(value) ? (value.length === 1 ? value[0] : undefined) : value;
    if (field.dataset.kind !== "text" || text === undefined) continue;
    const p = paragraphs(field)[0] as HTMLElement | undefined, run = p?.querySelector<HTMLElement>("span");
    const size = parseFloat(run?.style.fontSize || p?.style.fontSize || "") || 16;
    const r = rectOf(field);
    if (r.h > (parseFloat(p?.style.lineHeight || "") || size * 1.3) * 1.6) continue;
    const pad = (parseFloat(field.style.paddingLeft) || 0) + (parseFloat(field.style.paddingRight) || 0);
    need = Math.max(need, b.w - r.w + lineWidth(text, size) + pad + 4);
  }
  return need;
}
/**
 * A row's lane: from the nearest element on its left to the nearest on its right in the same band (or the
 * container's edges). Blocks that hug an icon sit packed in the middle of a mostly empty band; their labels may
 * use that room. Returns the new rects, or undefined when the blocks already fit.
 */
function spreadRow(parent: Element, instances: Element[], seps: Element[], rects: Rect[], needs: number[], planned: Rect[]): Rect[] | undefined {
  const count = needs.length;
  if (!count || arrange(rects).arrangement !== "row" && !(rects.length === 1 && count > 1)) return;
  const w = Math.max(...needs);
  // Only when the ordinary re-layout (or the blocks as they are) would still be too narrow.
  if (w <= Math.min(...planned.map(r => r.w)) + 1) return;
  const y0 = Math.min(...rects.map(r => r.y)), y1 = Math.max(...rects.map(r => r.y + r.h));
  const x0 = Math.min(...rects.map(r => r.x)), x1 = Math.max(...rects.map(r => r.x + r.w));
  const own = new Set<Element>([...instances, ...seps]);
  const others = (Array.from(parent.children) as HTMLElement[]).filter(n => !own.has(n) && n.dataset.kind).map(rectOf)
    .filter(n => n.y < y1 && n.y + n.h > y0);
  const container = (parent as HTMLElement).dataset?.kind === "group" ? { lo: 0, hi: rectOf(parent).w } : { lo: 24, hi: pageWidth(parent) - 24 };
  const lo = Math.max(container.lo, ...others.filter(n => n.x + n.w <= x0 + 1).map(n => n.x + n.w + 8));
  const hi = Math.min(container.hi, ...others.filter(n => n.x >= x1 - 1).map(n => n.x - 8));
  const sorted = [...rects].sort((a, b) => a.x - b.x);
  const gaps = sorted.slice(1).map((r, i) => r.x - (sorted[i].x + sorted[i].w)).filter(g => g >= 0);
  let gap = gaps.length ? gaps.reduce((a, b) => a + b, 0) / gaps.length : 24, width = w;
  if (count * width + (count - 1) * gap > hi - lo) gap = count > 1 ? Math.max(8, (hi - lo - count * width) / (count - 1)) : 0;
  if (count * width + (count - 1) * gap > hi - lo) width = (hi - lo - (count - 1) * gap) / count;
  const total = count * width + (count - 1) * gap;
  const start = Math.min(Math.max(lo, (x0 + x1) / 2 - total / 2), hi - total);
  return Array.from({ length: count }, (_, i) => ({ x: start + i * (width + gap), y: rects[0].y, w: width, h: rects[0].h }));
}

function fillSet(scope: Element | null, root: ParentNode, fill: FillSet, notes: string[]): void {
  const sets = setsIn(root, scope);
  const slot = sets.has(fill.set) ? fill.set : Array.from(sets.keys()).find(s => blockName(s) === fill.set);
  if (!slot) throw new Error(`FILL_LAYOUT: no block set "${fill.set}" here; use a set from get_layout`);
  const instances = sets.get(slot)!, parent = instances[0].parentElement!;
  if (instances.some(el => el.parentElement !== parent)) throw new Error(`FILL_LAYOUT: set ${slot} spans several containers`);
  const seps = Array.from(parent.children).filter(el => slotOf(el) === `sep#${slot.replace(BLOCK, "$2")}`);
  const rects = instances.map(rectOf), count = fill.items.length;
  const elements = fill.items.map((_, i) => {
    if (i < instances.length) return instances[i];
    const copy = instances[instances.length - 1].cloneNode(true) as Element;
    copy.setAttribute("data-fill-new", "");
    return copy;
  });
  for (const extra of instances.slice(count)) extra.remove();
  if (!count && scope === null) {
    for (const sep of seps) sep.remove();
    const removed = [...rects, ...seps.map(rectOf)];
    // Decoration drawn for that row alone (the soft arc under a scene row) goes with it.
    const y0 = Math.min(...removed.map(r => r.y)), y1 = Math.max(...removed.map(r => r.y + r.h));
    for (const shape of Array.from(parent.children) as HTMLElement[]) {
      if (shape.dataset.kind !== "shape" || slotOf(shape) || shape.dataset.purpose === "fixed" || shape.dataset.locked) continue;
      const s = rectOf(shape);
      if (s.y >= y0 - 2 && s.y + s.h <= y1 + 12) { removed.push(s); shape.remove(); }
    }
    closeGap(parent, removed);
    return;
  }
  for (let i = instances.length; i < count; i++) elements[i - 1].after(elements[i]);
  // A row of small blocks whose one-line labels need more width than the blocks have spreads along its lane
  // first, so the labels are filled at the width they need instead of wrapping or shrinking.
  const ordinary = count !== instances.length || fill.columns !== undefined;
  const planned = ordinary ? layout(rects, count, fill.columns) : rects;
  const spread = spreadRow(parent, instances, seps, rects, elements.map((el, i) => blockNeed(el, fill.items[i])), planned);
  if (spread) spread.forEach((r, i) => resize(elements[i], r));
  // Texts first, so re-laid blocks size their texts by the new words, not the sample's; icons and nested sets
  // then fill the blocks at their new size.
  fill.items.forEach((item, i) => fillTexts(elements[i], item, notes));
  const relaid = !!spread || ordinary;
  if (ordinary && !spread) {
    planned.forEach((r, i) => resize(elements[i], r));
    // Re-laid blocks may be narrower than when their labels were fitted: fit the one-line labels again.
    fill.items.forEach((item, i) => refit(elements[i], item));
  }
  // One separator in each gap, centred, reusing the template's own separators first.
  if (seps.length) {
    const wanted = Math.max(0, count - 1);
    for (const extra of seps.slice(wanted)) extra.remove();
    const used = seps.slice(0, wanted);
    while (used.length < wanted) {
      const copy = seps[0].cloneNode(true) as Element;
      copy.setAttribute("data-fill-new", "");
      parent.appendChild(copy);
      used.push(copy);
    }
    if (relaid) used.forEach((sep, i) => {
      const a = rectOf(elements[i]), b = rectOf(elements[i + 1]), s = rectOf(sep);
      const row = Math.abs(a.y - b.y) < 3;
      place(sep, row ? { ...s, x: (a.x + a.w + b.x) / 2 - s.w / 2 } : { ...s, y: (a.y + a.h + b.y) / 2 - s.h / 2 });
    });
  }
  fill.items.forEach((item, i) => fillRest(elements[i], item, notes));
}

/**
 * A page-level set removed across the full content width leaves an empty band. The template's flow is gone
 * (everything is absolute), so the page content is re-spaced by bands: the gap that held the removed set shrinks
 * to the page's usual gap, and what it frees is shared evenly by all the gaps. The content keeps its top and
 * bottom; fixed elements and anything outside the blocks' vertical span (header, page number) stay put.
 */
function closeGap(container: Element, removed: Rect[]): void {
  const y0 = Math.min(...removed.map(r => r.y)), y1 = Math.max(...removed.map(r => r.y + r.h));
  const movable = (Array.from(container.children) as HTMLElement[]).filter(el => el.dataset.kind && el.dataset.purpose !== "fixed" && !el.dataset.locked);
  const blocks = movable.filter(el => BLOCK.test(slotOf(el)));
  if (!blocks.length) return;
  const top = Math.min(y0, ...blocks.map(el => rectOf(el).y)), bottom = Math.max(y1, ...blocks.map(el => { const r = rectOf(el); return r.y + r.h; }));
  const content = movable.filter(el => { const r = rectOf(el); return r.y >= top - 1 && r.y + r.h <= bottom + 1; });
  if (content.some(el => { const r = rectOf(el); return r.y < y1 - 1 && r.y + r.h > y0 + 1; })) return;
  const bands: Array<{ y0: number; y1: number; els: HTMLElement[] }> = [];
  for (const el of [...content].sort((a, b) => rectOf(a).y - rectOf(b).y)) {
    const r = rectOf(el), last = bands[bands.length - 1];
    if (last && r.y < last.y1 + .5) { last.y1 = Math.max(last.y1, r.y + r.h); last.els.push(el); }
    else bands.push({ y0: r.y, y1: r.y + r.h, els: [el] });
  }
  if (bands.length < 2) return;
  const gaps = bands.slice(1).map((b, i) => b.y0 - bands[i].y1);
  const hole = gaps.findIndex((_, i) => bands[i].y1 <= y0 + 1 && bands[i + 1].y0 >= y1 - 1);
  if (hole < 0) return;
  const others = gaps.filter((_, i) => i !== hole).sort((a, b) => a - b);
  const usual = others.length ? others[Math.floor(others.length / 2)] : 10;
  const extra = Math.max(0, gaps[hole] - usual) / gaps.length;
  let y = bands[0].y0;
  bands.forEach((band, i) => {
    if (i) y += (i - 1 === hole ? usual : gaps[i - 1]) + extra;
    const delta = y - band.y0;
    if (Math.abs(delta) > .01) for (const el of band.els) { const r = rectOf(el); place(el, { ...r, y: r.y + delta }); }
    y += band.y1 - band.y0;
  });
}

function refit(block: Element, item: FillItem): void {
  for (const field of fieldsIn(block) as HTMLElement[]) {
    const value = item.fields?.[slotOf(field)];
    const text = Array.isArray(value) ? (value.length === 1 ? value[0] : undefined) : value;
    if (field.dataset.kind === "text" && text !== undefined) fitLine(field, text);
  }
}

function fillTexts(block: Element, item: FillItem, notes: string[]): void {
  const fields = fieldsIn(block);
  for (const [name, value] of Object.entries(item.fields ?? {})) {
    const field = fields.find(f => slotOf(f) === name && (f as HTMLElement).dataset.kind === "text");
    if (field) { fillText(field, value); growDown(field as HTMLElement); } else notes.push(`${slotOf(block)} 没有文字字段 ${name}`);
  }
}

/**
 * The sample's text box is often shorter than the room its block leaves free under it. Words that need more
 * height take that room, down to the next element below in the block or the block's bottom, before they count
 * as overflowing.
 */
function growDown(el: HTMLElement): void {
  const parent = el.parentElement as HTMLElement | null;
  if (!parent || parent.dataset.kind !== "group") return;
  const r = rectOf(el), need = textHeight(el, r.w);
  if (need === undefined || need <= r.h + 1) return;
  const below = (Array.from(parent.children) as HTMLElement[]).filter(n => n !== el && n.dataset.kind).map(rectOf)
    .filter(n => n.y >= r.y + r.h - 1 && n.x < r.x + r.w && n.x + n.w > r.x);
  const floor = Math.min(rectOf(parent).h - 2, ...below.map(n => n.y - 4));
  if (floor > r.y + r.h) place(el, { ...r, h: Math.min(need, floor - r.y) });
}

function fillRest(block: Element, item: FillItem, notes: string[]): void {
  const fields = fieldsIn(block);
  for (const [name, svg] of Object.entries(item.icons ?? {})) {
    const field = fields.find(f => slotOf(f) === name);
    if (field) fillIcon(field, svg); else notes.push(`${slotOf(block)} 没有图标位 ${name}`);
  }
  for (const set of item.sets ?? []) fillSet(block, block, set, notes);
}

/** Copies are new objects: they carry no identity or template annotations of their own. */
function settleCopies(doc: Document): void {
  for (const copy of Array.from(doc.querySelectorAll("[data-fill-new]"))) {
    for (const el of [copy, ...Array.from(copy.querySelectorAll("*"))]) {
      for (const attr of ["data-fill-new", "data-id", "data-slot", "data-purpose", "data-note", "data-name", "data-locked", "data-required", "data-role", "data-bind", "data-max-chars"]) el.removeAttribute(attr);
    }
  }
}

export function fillLayout(html: string, request: FillRequest): { html: string; notes: string[] } {
  const doc = parseDesignHtml(html, true), notes: string[] = [];
  for (const [id, value] of Object.entries(request.texts ?? {})) {
    const el = Array.from(doc.querySelectorAll<HTMLElement>("[data-id]")).find(n => n.dataset.id === id);
    if (!el || el.dataset.kind !== "text") throw new Error(`FILL_LAYOUT: unknown text ${id}`);
    fillText(el, value);
  }
  const mentioned = new Set<string>();
  for (const set of request.sets ?? []) { fillSet(null, doc.body, set, notes); mentioned.add(set.set); }
  settleCopies(doc);
  const untouched = Array.from(setsIn(doc.body, null).keys()).filter(s => !mentioned.has(s) && !mentioned.has(blockName(s)));
  if (untouched.length) notes.push(`未填写、仍是模板示例的块：${untouched.join("、")}；不用的块传 items: [] 删除`);
  return { html: "<!doctype html>" + doc.documentElement.outerHTML, notes };
}
