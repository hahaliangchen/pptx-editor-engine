import type { CmRect, Outline, OutlineShape, OutlineSlide } from "./outline";

/**
 * Reading a template deck: which of its pages is the cover, the contents page,
 * a section page, a content page and the closing page, and what can be filled
 * in on each. Templates mark what to fill with `{{TOKEN}}` text, the way the
 * SVG templates of ai-ppt-server do; no other description is needed.
 */

export type PageKind = "cover" | "toc" | "section" | "content" | "end";
export const PAGE_KINDS: PageKind[] = ["cover", "toc", "section", "content", "end"];

export interface TextSlot {
  /** The shape on the template page. */
  id: string;
  /** Its text on the template, with the `{{TOKEN}}`s in it. */
  text: string;
  rect: CmRect;
  fontSizePt?: number;
  align?: string;
}

/** One entry of a contents page: every shape that belongs to it, so it can be dropped as a whole. */
export interface TocEntry {
  shapes: string[];
  /** Top and bottom edge of the entry, in cm. */
  top: number;
  bottom: number;
  title: string;
  desc?: string;
}

export interface TemplatePage {
  kind: PageKind;
  /** The page on the template, from 1. */
  slide: number;
  /** Text shapes that carry at least one token, and the fillable name of each token. */
  tokens: TextSlot[];
  /** Fillable names this page takes from the content, token by token. */
  slots: Record<string, TextSlot>;
  /** Text that is not a token: sample wording that may need replacing. */
  fixed: TextSlot[];
  /** Every text shape of the page, by id. */
  texts: Record<string, TextSlot>;
  /** Every shape's rectangle, by id. */
  rects: Record<string, CmRect>;
  /** Contents page only. */
  entries?: TocEntry[];
  /** Content page only: where the content goes, and the shapes that only mark that place. */
  area?: CmRect;
  areaShapes?: string[];
  /** The page is dark, so text on it must be light. */
  dark: boolean;
}

export interface TemplateInfo {
  slideSize: { w: number; h: number };
  slideCount: number;
  pages: Partial<Record<PageKind, TemplatePage>>;
  /** The strongest colour of the template, for what is added to its pages. */
  accent: string;
}

/** Tokens the engine fills in itself. */
export const AUTO_TOKENS = new Set(["PAGE_NUM", "TOTAL_PAGES", "CHAPTER_NUM", "SECTION_NUM"]);
const TOKEN_RE = /\{\{([A-Z0-9_]+)\}\}/g;

const SLOT_NAMES: Record<string, string> = {
  PAGE_TITLE: "title", CHAPTER_TITLE: "title", THANK_YOU: "title", TITLE: "title",
  ENDING_SUBTITLE: "subtitle", CHAPTER_SUBTITLE: "subtitle", CHAPTER_DESC: "desc",
  CLOSING_MESSAGE: "message", CONTACT_INFO: "contact"
};

export const tokensIn = (text: string): string[] => [...text.matchAll(TOKEN_RE)].map(match => match[1]);

/** TOKEN_NAME → tokenName; the few common ones have shorter names. */
export function slotName(token: string): string {
  if (SLOT_NAMES[token]) return SLOT_NAMES[token];
  const words = token.toLowerCase().split("_");
  return words[0] + words.slice(1).map(word => word[0].toUpperCase() + word.slice(1)).join("");
}

const flatten = (shapes: OutlineShape[]): OutlineShape[] =>
  shapes.flatMap(shape => (shape.children ? flatten(shape.children) : [shape]));

function kindOf(tokens: Set<string>): PageKind | undefined {
  const has = (prefix: string) => [...tokens].some(token => token.startsWith(prefix));
  if (has("TOC_ITEM")) return "toc";
  if (tokens.has("CONTENT_AREA")) return "content";
  if (has("CHAPTER_") || tokens.has("SECTION_NAME")) return "section";
  if (tokens.has("THANK_YOU")) return "end";
  if (tokens.has("TITLE")) return "cover";
  return undefined;
}

/** Most of a colour's weight goes by how much it is used and how far it is from grey. */
function accentOf(slides: OutlineSlide[]): string {
  const weight = new Map<string, number>();
  const add = (color: string | undefined, amount: number) => {
    if (!color || !/^#[0-9a-f]{6}$/i.test(color)) return;
    const [r, g, b] = [1, 3, 5].map(i => parseInt(color.slice(i, i + 2), 16));
    const spread = Math.max(r, g, b) - Math.min(r, g, b);
    if (spread < 40) return;
    weight.set(color.toLowerCase(), (weight.get(color.toLowerCase()) ?? 0) + amount * spread);
  };
  for (const slide of slides) {
    for (const shape of flatten(slide.shapes)) {
      const fill = shape.fill;
      if (typeof fill === "string") add(fill, 1);
      else if (fill && "gradient" in fill) add(fill.gradient.stops[0]?.color, 1);
      add(shape.style?.color, 0.5);
    }
  }
  return [...weight.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "#2f5597";
}

function isDark(background: OutlineSlide["background"]): boolean {
  const color = typeof background === "string" ? background : background?.gradient.stops[0]?.color;
  if (!color || !/^#[0-9a-f]{6}$/i.test(color)) return false;
  const [r, g, b] = [1, 3, 5].map(i => parseInt(color.slice(i, i + 2), 16));
  return 0.299 * r + 0.587 * g + 0.114 * b < 110;
}

const isNumber = (text: string | undefined) => !!text && /^\d{1,2}$/.test(text.trim());

function tocEntries(shapes: OutlineShape[]): TocEntry[] {
  const titles = shapes
    .map((shape, index) => ({ shape, index }))
    .filter(({ shape }) => shape.text && /\{\{TOC_ITEM_\d+_TITLE\}\}/.test(shape.text));
  const entries: TocEntry[] = [];
  let start = 0;
  titles.forEach(({ shape, index }, order) => {
    if (order === 0) {
      // The header ("目录") is the last wording before the first entry's number.
      let back = index - 1;
      while (back >= 0 && !(shapes[back].type === "text" && shapes[back].text && !isNumber(shapes[back].text))) back--;
      start = back + 1;
    }
    let end = index;
    const next = shapes[index + 1];
    if (next && next.type === "text" && next.text && !isNumber(next.text)
      && !/\{\{TOC_ITEM_\d+_TITLE\}\}/.test(next.text)) end = index + 1;
    const desc = end > index ? shapes[end] : undefined;
    const members = shapes.slice(start, end + 1);
    entries.push({
      shapes: members.map(item => item.id),
      top: Math.min(...members.map(item => item.rect.y)),
      bottom: Math.max(...members.map(item => item.rect.y + item.rect.h)),
      title: shape.id,
      ...(desc ? { desc: desc.id } : {})
    });
    start = end + 1;
  });
  return entries;
}

function areaOf(shapes: OutlineShape[], slideSize: { w: number; h: number }): { rect?: CmRect; shapes: string[] } {
  const label = shapes.find(shape => shape.text?.includes("{{CONTENT_AREA}}"));
  if (!label) return { shapes: [] };
  const centre = { x: label.rect.x + label.rect.w / 2, y: label.rect.y + label.rect.h / 2 };
  // The frame that marks the place: the largest shape around the label that is not the whole page.
  const frames = shapes
    .filter(shape => shape.type === "shape" && shape.rect.w * shape.rect.h < slideSize.w * slideSize.h * 0.9
      && shape.rect.x <= centre.x && shape.rect.x + shape.rect.w >= centre.x
      && shape.rect.y <= centre.y && shape.rect.y + shape.rect.h >= centre.y
      && shape.rect.w > label.rect.w * 1.5 && shape.rect.h > label.rect.h * 1.5)
    .sort((a, b) => b.rect.w * b.rect.h - a.rect.w * a.rect.h);
  if (frames[0]) return { rect: frames[0].rect, shapes: [frames[0].id, label.id] };

  // Otherwise: below the title, inside the page's side margins, above the footer.
  const texts = shapes.filter(shape => shape.type === "text" && shape !== label && shape.text);
  const title = texts.find(shape => shape.text?.includes("{{PAGE_TITLE}}")) ?? texts[0];
  const margin = title?.rect.x ?? 1.5;
  const top = title ? title.rect.y + title.rect.h + 0.4 : slideSize.h * 0.2;
  const footer = Math.min(slideSize.h, ...texts.filter(shape => shape.rect.y > slideSize.h * 0.8).map(shape => shape.rect.y));
  return {
    rect: { x: margin, y: top, w: slideSize.w - 2 * margin, h: footer - 0.4 - top },
    shapes: [label.id]
  };
}

/** Read a template deck's outline: its pages by kind, with what each can take. */
export function inferTemplate(outline: Outline): TemplateInfo {
  const pages: TemplateInfo["pages"] = {};
  for (const slide of outline.slides) {
    const shapes = flatten(slide.shapes);
    const texts = shapes.filter(shape => shape.type === "text" && shape.text);
    const tokens = new Set(texts.flatMap(shape => tokensIn(shape.text!)));
    const kind = kindOf(tokens);
    if (!kind || pages[kind]) continue;

    const slots: Record<string, TextSlot> = {};
    const slotOf = (shape: OutlineShape): TextSlot => ({
      id: shape.id, text: shape.text!, rect: shape.rect, fontSizePt: shape.style?.fontSizePt, align: shape.style?.align
    });
    const tokenShapes: TextSlot[] = [];
    const fixed: TextSlot[] = [];
    for (const shape of texts) {
      const found = tokensIn(shape.text!);
      if (found.length > 0) {
        tokenShapes.push(slotOf(shape));
        for (const token of found) {
          if (AUTO_TOKENS.has(token) || token === "CONTENT_AREA" || /^TOC_ITEM_/.test(token) || /^LOGO/.test(token)) continue;
          slots[slotName(token)] ??= slotOf(shape);
        }
      } else if (!isNumber(shape.text) && /[\p{L}]{3,}|[一-鿿]{2,}/u.test(shape.text!)) {
        fixed.push(slotOf(shape));
      }
    }
    const page: TemplatePage = {
      kind, slide: slide.index, tokens: tokenShapes, slots, fixed, texts: Object.fromEntries(texts.map(shape => [shape.id, slotOf(shape)])),
      rects: Object.fromEntries(shapes.map(shape => [shape.id, shape.rect])), dark: isDark(slide.background)
    };
    if (kind === "toc") {
      page.entries = tocEntries(shapes);
      const inEntries = new Set(page.entries.flatMap(entry => entry.shapes));
      page.fixed = fixed.filter(item => !inEntries.has(item.id));
    }
    if (kind === "content") {
      const area = areaOf(shapes, outline.slideSize);
      page.area = area.rect;
      page.areaShapes = area.shapes;
    }
    pages[kind] = page;
  }
  return { slideSize: outline.slideSize, slideCount: outline.slides.length, pages, accent: accentOf(outline.slides) };
}

/** What a model needs to know to fill the template: page kinds, their fields, the room for content. */
export function describeTemplate(template: TemplateInfo): unknown {
  const rounded = (value: number) => Math.round(value * 10) / 10;
  return {
    pages: Object.fromEntries(Object.values(template.pages).map(page => [page!.kind, {
      fields: Object.keys(page!.slots),
      ...(page!.entries ? { entries: page!.entries.length, entryFields: page!.entries.some(entry => entry.desc) ? ["title", "desc"] : ["title"] } : {}),
      ...(page!.area ? { contentArea: { w: rounded(page!.area.w), h: rounded(page!.area.h) } } : {}),
      ...(page!.fixed.length > 0 ? { sampleText: page!.fixed.map(item => item.text) } : {})
    }]))
  };
}
