import type { CmRect, Outline, OutlineShape } from "./outline";
import { PAGE_KINDS, PageKind } from "./template";
import { isImageShape } from "./units";

/**
 * A first guess at how to annotate a page, from the numbers alone: which shapes are the template's
 * own decoration (they repeat in the same place on other pages), which stand in an even row or
 * column and so form a repeating unit, which text is the page title or a number. It writes the
 * `setName` commands for what it found; the model looks at the page and corrects them.
 */

export interface SuggestedUnit {
  name: string;
  count: number;
  axis: "row" | "column";
  /** The shapes of each instance, with the part each plays. */
  instances: Array<Array<{ id: string; name?: string; field?: string }>>;
}

export interface AnnotationSuggestion {
  slide: number;
  /** The shapes recognised as the template's decoration: leave them unnamed. */
  decoration: string[];
  units: SuggestedUnit[];
  /** Text that is probably a field of the page itself. */
  fields: Array<{ id: string; suggested: string; sample: string }>;
  commands: Array<{ op: "setName"; target: { slide: number; id: string }; name: string }>;
  notes: string[];
}

const ROLE_WORDS: Record<PageKind, string> = { cover: "封面", toc: "目录", section: "章节", content: "内容", end: "结尾" };
const flat = (shapes: OutlineShape[]): OutlineShape[] => shapes.flatMap(shape => (shape.children ? [shape, ...flat(shape.children)] : [shape]));
const round1 = (value: number) => Math.round(value * 10) / 10;
const center = (rect: CmRect) => ({ x: rect.x + rect.w / 2, y: rect.y + rect.h / 2 });
const bounds = (rects: CmRect[]): CmRect => {
  const x = Math.min(...rects.map(r => r.x));
  const y = Math.min(...rects.map(r => r.y));
  return { x, y, w: Math.max(...rects.map(r => r.x + r.w)) - x, h: Math.max(...rects.map(r => r.y + r.h)) - y };
};
/** The distance between two rectangles in cm (0 when they touch or overlap). */
const gap = (a: CmRect, b: CmRect) => Math.hypot(Math.max(0, a.x - (b.x + b.w), b.x - (a.x + a.w)), Math.max(0, a.y - (b.y + b.h), b.y - (a.y + a.h)));
const isNumber = (text?: string) => !!text && /^\d{1,2}$/.test(text.trim());

/** What a page of the whole deck is likely to be, from its words and position. */
export function suggestRoles(outline: Outline): Array<{ slide: number; role: PageKind; why: string }> {
  const last = outline.slides.length;
  return outline.slides.map(slide => {
    const shapes = flat(slide.shapes);
    const texts = shapes.filter(shape => shape.text).map(shape => shape.text as string);
    const all = texts.join(" ");
    const numbered = texts.filter(isNumber).length;
    const big = shapes.filter(shape => shape.text && (shape.style?.fontSizePt ?? 0) >= 40);
    if (/目\s*录|contents?|catalog/i.test(all) && numbered >= 2) return { slide: slide.index, role: "toc" as const, why: "有“目录”字样和几个编号" };
    if (slide.index === last && /thank|谢谢|感谢|THANK/i.test(all)) return { slide: slide.index, role: "end" as const, why: "最后一页，写着谢谢" };
    if (slide.index === 1) return { slide: slide.index, role: "cover" as const, why: "第一页" };
    if (texts.some(isNumber) && big.some(shape => isNumber(shape.text)) && texts.length <= 6) return { slide: slide.index, role: "section" as const, why: "一个大数字加一个标题，别的很少" };
    if (/thank|谢谢|感谢/i.test(all) && texts.length <= 4) return { slide: slide.index, role: "end" as const, why: "写着谢谢，文字很少" };
    return { slide: slide.index, role: "content" as const, why: "其余的页面" };
  });
}

/** Whether a text is wording to replace: placeholder prompts, dummy English, bare numbers. */
export function looksLikeSample(text: string): boolean {
  const t = text.trim();
  return /单击此处|点击此处|请输入|点击输入|输入标题|输入文字|添加您的|lorem ipsum/i.test(t)
    || isNumber(t)
    || (/^[\x00-\x7f\s.,'’";:!?-]{25,}$/.test(t) && /\s/.test(t));
}

function evenSteps(values: number[]): number | undefined {
  if (values.length < 2) return undefined;
  const steps = values.slice(1).map((value, i) => value - values[i]);
  const mean = steps.reduce((sum, step) => sum + step, 0) / steps.length;
  return Math.abs(mean) > 0.5 && steps.every(step => Math.abs(step - mean) <= 0.7) ? mean : undefined;
}

interface Cluster { nodes: OutlineShape[]; axis: "row" | "column"; pitch: number }

export function suggestAnnotations(outline: Outline, slideIndex: number, options: { role?: PageKind; variant?: string; unit?: string } = {}): AnnotationSuggestion {
  const slide = outline.slides.find(item => item.index === slideIndex);
  if (!slide) throw new Error(`没有第 ${slideIndex} 页`);
  const role = options.role ?? suggestRoles(outline).find(item => item.slide === slideIndex)!.role;
  const roleWord = ROLE_WORDS[role];
  const variant = options.variant ?? (role === "content" ? `版式${slideIndex}` : undefined);
  const prefix = `@${roleWord}${variant ? `:${variant}` : ""}`;
  const page = outline.slides[0] ? outline.slideSize.w * outline.slideSize.h : 1;
  const notes: string[] = [];

  // Decoration: what stands in the same place, with the same name, on at least half the other pages.
  const signature = (shape: OutlineShape) => `${shape.name}|${round1(shape.rect.x)}|${round1(shape.rect.y)}|${round1(shape.rect.w)}|${round1(shape.rect.h)}`;
  const seen = new Map<string, number>();
  for (const other of outline.slides) for (const key of new Set(other.shapes.map(signature))) seen.set(key, (seen.get(key) ?? 0) + 1);
  const sampleInside = (shape: OutlineShape): boolean => (!!shape.text && looksLikeSample(shape.text)) || (shape.children ?? []).some(sampleInside);
  // The same shape in the same place on many pages is the template's own; unless it holds sample wording ("单击此处输入标题"), which is a field that merely reads the same on every page.
  const repeated = (shape: OutlineShape) => outline.slides.length >= 3 && (seen.get(signature(shape)) ?? 0) >= Math.max(3, outline.slides.length * 0.4) && !sampleInside(shape);
  const top = slide.shapes;
  const decoration = new Set<string>();
  const live: OutlineShape[] = [];
  for (const node of top) {
    const big = node.rect.w * node.rect.h > page * 0.45;
    if (repeated(node) || big || (node.type === "image" && node.rect.w * node.rect.h > page * 0.3)) decoration.add(node.id);
    else live.push(node);
  }

  // Clusters: shapes of the same size in an even row or column.
  const clusters: Cluster[] = [];
  const taken = new Set<string>();
  for (const node of live) {
    if (taken.has(node.id)) continue;
    const mates = live.filter(other => !taken.has(other.id) && other.type === node.type
      && Math.abs(other.rect.w - node.rect.w) <= Math.max(0.5, node.rect.w * 0.08) && Math.abs(other.rect.h - node.rect.h) <= Math.max(0.5, node.rect.h * 0.08));
    if (mates.length < 2) continue;
    const xs = mates.map(shape => shape.rect.x);
    const ys = mates.map(shape => shape.rect.y);
    const row = Math.max(...xs) - Math.min(...xs) >= Math.max(...ys) - Math.min(...ys);
    const sorted = [...mates].sort((a, b) => (row ? a.rect.x - b.rect.x : a.rect.y - b.rect.y));
    const pitch = evenSteps(sorted.map(shape => (row ? shape.rect.x : shape.rect.y)));
    if (pitch === undefined) continue;
    mates.forEach(mate => taken.add(mate.id));
    clusters.push({ nodes: sorted, axis: row ? "row" : "column", pitch });
  }
  // Clusters with the same count and spacing are parts of the same unit.
  const unitsOfClusters: Cluster[][] = [];
  for (const cluster of clusters) {
    const mate = unitsOfClusters.find(group => group[0].axis === cluster.axis && group[0].nodes.length === cluster.nodes.length && Math.abs(group[0].pitch - cluster.pitch) <= 0.8);
    if (mate) mate.push(cluster);
    else unitsOfClusters.push([cluster]);
  }

  const units: SuggestedUnit[] = [];
  const commands: AnnotationSuggestion["commands"] = [];
  const named = new Set<string>();
  const nameOf = (id: string, name: string) => {
    named.add(id);
    commands.push({ op: "setName", target: { slide: slideIndex, id }, name });
  };
  let unitNo = 0;
  unitsOfClusters.filter(group => group[0].nodes.length >= 2).forEach(group => {
    const { axis, pitch } = group[0];
    const count = group[0].nodes.length;
    const members: OutlineShape[][] = Array.from({ length: count }, (_, i) => group.map(cluster => cluster.nodes[i]));
    // A row of plain shapes with no text or picture (dividers, connectors) is decoration, whatever stands near it.
    if (!members.flat().some(node => node.text || isImageShape(node) || node.children?.some(child => child.text || isImageShape(child)))) {
      notes.push(`第 ${members.map(nodes => nodes[0].id).join("、")} 号形状是等距排开的装饰（可能是连线或分隔线）；如果它们属于某个重复结构，把它们命名成各项的零件，否则别动`);
      members.flat().forEach(node => taken.delete(node.id));
      return;
    }
    // Everything else that sits close to an instance (within 1.5 cm of what it already holds) belongs to it.
    const all = members.flat();
    const origin = members[0].reduce((sum, shape) => sum + (axis === "row" ? center(shape.rect).x : center(shape.rect).y), 0) / members[0].length;
    for (const node of live) {
      if (taken.has(node.id)) continue;
      const c = center(node.rect);
      const along = axis === "row" ? c.x : c.y;
      const index = Math.round((along - origin) / pitch);
      if (index < 0 || index >= count || Math.abs(along - (origin + index * pitch)) > Math.abs(pitch) * 0.55) continue;
      if (gap(node.rect, bounds(members[index].map(member => member.rect))) > 1.5) continue;
      members[index].push(node);
      taken.add(node.id);
    }
    const name = options.unit ?? (unitNo === 0 ? "项" : `项${"乙丙丁"[unitNo - 1] ?? unitNo}`);
    unitNo += 1;
    const instances = members.map((nodes, i) => {
      const texts = nodes.filter(node => node.text && !isImageShape(node));
      const sizes = texts.filter(node => !isNumber(node.text)).sort((a, b) => (b.style?.fontSizePt ?? 0) - (a.style?.fontSizePt ?? 0) || a.rect.y - b.rect.y);
      const fieldFor = (node: OutlineShape): string | undefined => {
        if (isImageShape(node)) return "图片";
        if (isNumber(node.text)) return "序号";
        const rank = sizes.indexOf(node);
        return rank === 0 ? "标题" : rank === 1 ? "说明" : rank === 2 ? "补充" : undefined;
      };
      return nodes.flatMap(node => {
        const field = fieldFor(node);
        // A contents page numbers its entries (`@目录.2.标题`); the entry numbers and decoration need no name, the engine groups them by position.
        if (role === "toc") {
          if (field !== "标题" && field !== "说明") return [];
          const fullName = `${prefix}.${i + 1}.${field === "标题" ? "标题" : "说明"}`;
          nameOf(node.id, fullName);
          return [{ id: node.id, name: fullName, field }];
        }
        const fullName = `${prefix}.${name}${i + 1}${field ? `.${field}` : ""}`;
        nameOf(node.id, fullName);
        return [{ id: node.id, name: fullName, ...(field ? { field } : {}) }];
      });
    });
    units.push({ name, count, axis, instances });
    if (count < 2) notes.push(`“${name}”只有 ${count} 个，不能增减`);
  });

  // The page's own text, by role: the title first (the biggest type, then the highest), then what the role usually has.
  const fields: AnnotationSuggestion["fields"] = [];
  const leaves = top.filter(node => !decoration.has(node.id)).flatMap(node => flat([node])).filter(node => node.text && !named.has(node.id) && !taken.has(node.id) && !node.children);
  const ranked = leaves.filter(node => !isNumber(node.text)).sort((a, b) => (b.style?.fontSizePt ?? 0) - (a.style?.fontSizePt ?? 0) || a.rect.y - b.rect.y);
  const wanted: Record<PageKind, string[]> = { cover: ["标题", "副标题", "简介"], toc: [], section: ["标题", "说明"], content: ["标题"], end: ["标题", "寄语"] };
  wanted[role].forEach((field, i) => {
    const node = ranked[i];
    if (!node) return;
    fields.push({ id: node.id, suggested: field, sample: node.text as string });
    nameOf(node.id, `${prefix}.${field}`);
  });
  if (role === "section") {
    const number = leaves.find(node => isNumber(node.text));
    if (number) {
      fields.push({ id: number.id, suggested: "章节号", sample: number.text as string });
      nameOf(number.id, "@章节号");
    }
  }
  if (role === "toc") notes.push("目录页的标题文字（“目录”）和各项的编号、装饰都不用命名；各项写成 @目录.1.标题，引擎按位置把编号和装饰归到各项");
  for (const node of leaves) {
    if (named.has(node.id) || !looksLikeSample(node.text as string) || isNumber(node.text)) continue;
    fields.push({ id: node.id, suggested: "说明", sample: node.text as string });
  }
  if (units.length === 0) notes.push("没有找到等距重复的结构；如果页面上有卡片或步骤，请对照截图自己判断");
  notes.push("先看截图再用：命令只是猜测。重复结构里没有归进来的形状（连线、图标）要自己补；不需要变的装饰不要命名");
  return { slide: slideIndex, decoration: [...decoration], units, fields, commands, notes };
}

export const SUGGESTABLE_ROLES = PAGE_KINDS;
