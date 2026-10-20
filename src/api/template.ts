import type { CmRect, Outline, OutlineShape, OutlineSlide } from "./outline";
import { CM_TO_PT, INSET_CM } from "./text-metrics";
import { ImageSlot, UnitGroup, describeUnits, findUnits, imageSlotOf, isImageShape } from "./units";

/**
 * Reading a template deck: which of its pages is the cover, the contents page,
 * a section page, a content page and the closing page, and what can be filled
 * in on each. A template marks what to fill in one of two ways: a shape named
 * `@…` (a real sample text stays in it, see `bindOutline`), or `{{TOKEN}}`
 * text, the way the SVG templates of ai-ppt-server do. Both read the same.
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
  /** A shape bound by its name keeps a real sample text; this is it. */
  sample?: string;
  /** About how many characters fit the box at its size, and on how many lines. */
  fitChars?: number;
  lines?: number;
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
  /** The layout name, when the template has several pages of this role. */
  variant?: string;
  /** When to use this layout, in the designer's words. */
  when?: string;
  /** The page on the template, from 1. */
  slide: number;
  /** Text shapes that carry at least one token, and the fillable name of each token. */
  tokens: TextSlot[];
  /** Fillable names this page takes from the content, token by token. */
  slots: Record<string, TextSlot>;
  /** The field name of a token that is not one of the usual ones (a bound shape's own keyword). */
  fieldNames: Record<string, string>;
  /** Text that is not a token: sample wording that may need replacing. */
  fixed: TextSlot[];
  /** Every text shape of the page, by id. */
  texts: Record<string, TextSlot>;
  /** Every shape's rectangle, by id. */
  rects: Record<string, CmRect>;
  /** Repeating units the layout is built with (cards, steps, …), by name. */
  units: Record<string, UnitGroup>;
  /** Picture slots of the page itself (not in a unit), by field name. */
  images: Record<string, ImageSlot>;
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
  /** The page the deck uses for each role when it names no layout: the first one. */
  pages: Partial<Record<PageKind, TemplatePage>>;
  /** Every page a deck may use, in template order: the first of each role, and any other that names its layout. */
  variants: TemplatePage[];
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

// ---------------------------------------------------------------- binding by name

/**
 * A shape is bound to a field by its name: `@封面.标题`, `@内容.标题`, `@内容.区域`,
 * `@目录.1.标题`, `@目录.1.说明`, `@章节.说明`, `@结尾.联系方式`, or an automatic one with no
 * page role such as `@页码`, `@总页数`, `@章节号`. The shape keeps a real sample text, which shows the
 * designer what it will look like and the model how much fits. English words work too (`@cover.title`).
 */
const ROLE_NAMES: Record<string, PageKind> = {
  封面: "cover", cover: "cover", 目录: "toc", toc: "toc", 章节: "section", section: "section", 章节页: "section",
  内容: "content", content: "content", 结尾: "end", 结束: "end", end: "end"
};
const AUTO_NAMES: Record<string, string> = {
  页码: "PAGE_NUM", page: "PAGE_NUM", pagenum: "PAGE_NUM", 总页数: "TOTAL_PAGES", totalpages: "TOTAL_PAGES",
  章节号: "CHAPTER_NUM", chapternum: "CHAPTER_NUM", 序号: "SECTION_NUM", sectionnum: "SECTION_NUM"
};
const AREA_NAMES = new Set(["区域", "内容区", "内容区域", "area", "contentarea"]);
/** The usual fields of each page, in the words people use → the keyword the engine knows them by. */
const FIELD_TOKENS: Record<PageKind, Record<string, string>> = {
  cover: { 标题: "TITLE", title: "TITLE", 副标题: "SUBTITLE", subtitle: "SUBTITLE", 日期: "DATE", date: "DATE", 作者: "AUTHOR", author: "AUTHOR" },
  toc: {},
  section: { 标题: "CHAPTER_TITLE", title: "CHAPTER_TITLE", 说明: "CHAPTER_DESC", desc: "CHAPTER_DESC", 副标题: "CHAPTER_SUBTITLE", subtitle: "CHAPTER_SUBTITLE" },
  content: { 标题: "PAGE_TITLE", title: "PAGE_TITLE" },
  end: { 标题: "THANK_YOU", title: "THANK_YOU", 副标题: "ENDING_SUBTITLE", subtitle: "ENDING_SUBTITLE", 联系方式: "CONTACT_INFO", contact: "CONTACT_INFO", 寄语: "CLOSING_MESSAGE", message: "CLOSING_MESSAGE" }
};

/** The usual field of a page under the words people write it in (`副标题` for `subtitle`), when it is one. */
export function fieldAlias(kind: PageKind, key: string): string | undefined {
  const token = FIELD_TOKENS[kind][key.toLowerCase()];
  return token ? slotName(token) : undefined;
}

export interface Binding {
  /** The page role the name declares, if it has one. */
  role?: PageKind;
  /** The layout name the role part carries (`内容:三圆图`): one of several pages of the same role. */
  variant?: string;
  /** Options after `|`: `适用=…` says when to use the page. */
  options?: Record<string, string>;
  /** The keyword the engine fills (a usual `TITLE`-like one, or an encoded custom one); undefined for the content frame. */
  token?: string;
  /** The field name the content uses for it. */
  field?: string;
  /** The shape only marks where the content goes. */
  area?: boolean;
  /** A repeating unit: the name of the group (`卡片` in `@内容.卡片2.标题`), which instance, and whether this shape is only a part of it. */
  group?: string;
  index?: number;
  member?: boolean;
}

const encodeField = (field: string) => "U_" + [...field].map(char => char.codePointAt(0)!.toString(16).toUpperCase()).join("_");

/** What a shape's name binds it to; undefined when it is not bound (it does not start with `@`). */
export function parseBinding(name: string | undefined): Binding | undefined {
  const raw = (name ?? "").trim().replace(/^＠/, "@");
  if (!raw.startsWith("@") || raw.length < 2) return undefined;
  const [head, ...optionParts] = raw.slice(1).split("|");
  const options: Record<string, string> = {};
  for (const part of optionParts) {
    const at = part.search(/[=＝]/);
    const key = (at < 0 ? part : part.slice(0, at)).trim();
    if (key) options[key] = at < 0 ? "" : part.slice(at + 1).trim();
  }
  const withOptions = <T extends Binding>(binding: T): T => (Object.keys(options).length > 0 ? { ...binding, options } : binding);
  const parts = head.split(/[.．。]/).map(part => part.trim()).filter(Boolean);
  if (parts.length === 0) return undefined;
  const lower = (value: string) => value.toLowerCase();
  // The role part is `内容` or `内容:三圆图` (a layout name).
  const [roleWord, ...variantWords] = parts[0].split(/[:：]/);
  const role = ROLE_NAMES[lower(roleWord.trim())];
  const variant = role ? variantWords.join(":").trim() || undefined : undefined;
  const tag = <T extends Binding>(binding: T): T => withOptions(variant ? { ...binding, variant } : binding);
  const rest = role ? parts.slice(1) : parts;
  if (rest.length === 0) return role ? tag({ role }) : undefined;
  if (rest.length === 1 && AREA_NAMES.has(lower(rest[0])) && (!role || role === "content")) return tag({ role: "content", area: true });
  if (!role && rest.length === 1 && AUTO_NAMES[lower(rest[0])]) return withOptions({ token: AUTO_NAMES[lower(rest[0])] });
  if (role === "toc") {
    // @目录.1.标题 / @目录.1.说明 (the number first), or @目录.标题1
    const index = rest.find(part => /^\d+$/.test(part)) ?? /(\d+)$/.exec(rest[rest.length - 1])?.[1];
    const word = lower(rest.filter(part => !/^\d+$/.test(part)).join(".")).replace(/\d+$/, "");
    if (index && (word === "标题" || word === "title" || word === "")) return tag({ role, token: `TOC_ITEM_${index}_TITLE`, field: "title" });
    if (index && (word === "说明" || word === "desc")) return tag({ role, token: `TOC_ITEM_${index}_DESC`, field: "desc" });
    return tag({ role });
  }
  // A repeating unit: `卡片2.标题` is a field of the second instance, `卡片2` alone is a part of it (a background, an icon).
  // A field name must not end in a digit, or it reads as an instance.
  const instance = (segment: string) => /^(\D*?)(\d+)$/.exec(segment);
  if (role && rest.length === 2 && instance(rest[0]) && !instance(rest[1])) {
    const [, group, index] = instance(rest[0])!;
    return tag({ role, group: group || "项", index: Number(index), field: rest[1] });
  }
  if (role && rest.length === 1 && instance(rest[0])) {
    const [, group, index] = instance(rest[0])!;
    return tag({ role, group: group || "项", index: Number(index), member: true });
  }
  const field = rest.join(".");
  const known = role ? FIELD_TOKENS[role][lower(field)] : undefined;
  if (known) return tag({ role, token: known, field: slotName(known) });
  return tag({ role, token: encodeField(field), field });
}

export interface BoundOutline {
  outline: Outline;
  /** The real text of every bound shape, by `slide:id`. */
  samples: Map<string, string>;
}

/**
 * The outline as if every bound shape carried its keyword as `{{TOKEN}}` text, so that one reading of a
 * template serves both ways of marking it. A deck with no bound shape comes back unchanged.
 */
export function bindOutline(outline: Outline): BoundOutline {
  const samples = new Map<string, string>();
  let any = false;
  const bind = (slide: number, shapes: OutlineShape[]): OutlineShape[] => shapes.map(shape => {
    const children = shape.children ? bind(slide, shape.children) : undefined;
    const binding = parseBinding(shape.name);
    // The frame that marks where content goes carries the content-area keyword.
    if (binding?.area) { any = true; return { ...shape, ...(children ? { children } : {}), text: "{{CONTENT_AREA}}", binding: shape.name }; }
    if (!binding || !binding.token || !shape.text) return children ? { ...shape, children } : shape;
    any = true;
    samples.set(`${slide}:${shape.id}`, shape.text);
    return { ...shape, ...(children ? { children } : {}), text: `{{${binding.token}}}`, binding: shape.name };
  });
  const slides = outline.slides.map(slide => ({ ...slide, shapes: bind(slide.index, slide.shapes) }));
  return { outline: any ? { ...outline, slides } : outline, samples };
}

/** About how much text a box takes at its size: characters (a Chinese character is one em) and lines. */
export function capacityOf(rect: CmRect, fontSizePt: number | undefined): { fitChars: number; lines: number } {
  const size = fontSizePt ?? 18;
  const perLine = Math.max(1, Math.floor(((rect.w - 2 * INSET_CM) * CM_TO_PT) / size));
  const lines = Math.max(1, Math.floor((rect.h * CM_TO_PT) / (size * 1.25)));
  return { fitChars: perLine * lines, lines };
}

export const tokensIn = (text: string): string[] => [...text.matchAll(TOKEN_RE)].map(match => match[1]);

/** TOKEN_NAME → tokenName; the few common ones have shorter names. */
export function slotName(token: string): string {
  if (SLOT_NAMES[token]) return SLOT_NAMES[token];
  const words = token.toLowerCase().split("_");
  return words[0] + words.slice(1).map(word => word[0].toUpperCase() + word.slice(1)).join("");
}

export const flatten = (shapes: OutlineShape[]): OutlineShape[] =>
  shapes.flatMap(shape => (shape.children ? flatten(shape.children) : [shape]));

/** The layout name and the "when to use" the shapes of a page declare, if any (the first shape that says wins). */
export function declaredVariant(shapes: OutlineShape[]): { variant?: string; when?: string } {
  let variant: string | undefined;
  let when: string | undefined;
  for (const shape of shapes) {
    const binding = parseBinding(shape.binding ?? shape.name);
    variant ??= binding?.variant;
    when ??= binding?.options?.["适用"] ?? binding?.options?.when;
  }
  return { ...(variant ? { variant } : {}), ...(when ? { when } : {}) };
}

/** The page role the names of a page's shapes declare, when they all agree (or the first one does). */
export function declaredKind(shapes: OutlineShape[]): PageKind | undefined {
  for (const shape of shapes) {
    const role = parseBinding(shape.binding ?? shape.name)?.role;
    if (role) return role;
  }
  return undefined;
}

export function kindOf(tokens: Set<string>): PageKind | undefined {
  const has = (prefix: string) => [...tokens].some(token => token.startsWith(prefix));
  if (has("TOC_ITEM")) return "toc";
  if (tokens.has("CONTENT_AREA")) return "content";
  if (has("CHAPTER_") || tokens.has("SECTION_NAME")) return "section";
  if (tokens.has("THANK_YOU")) return "end";
  if (tokens.has("TITLE")) return "cover";
  return undefined;
}

/** Most of a colour's weight goes by how much it is used and how far it is from grey. */
export function accentOf(slides: OutlineSlide[]): string {
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

export function isDark(background: OutlineSlide["background"]): boolean {
  const color = typeof background === "string" ? background : background?.gradient.stops[0]?.color;
  if (!color || !/^#[0-9a-f]{6}$/i.test(color)) return false;
  const [r, g, b] = [1, 3, 5].map(i => parseInt(color.slice(i, i + 2), 16));
  return 0.299 * r + 0.587 * g + 0.114 * b < 110;
}

export const isNumber = (text: string | undefined) => !!text && /^\d{1,2}$/.test(text.trim());

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
    // Only what sits on the entry's own row belongs to it: a rule under the
    // heading that comes before the first entry stays where it is.
    const rowTop = Math.min(shape.rect.y, desc?.rect.y ?? Infinity);
    const rowBottom = Math.max(shape.rect.y + shape.rect.h, desc ? desc.rect.y + desc.rect.h : -Infinity);
    const members = shapes.slice(start, end + 1)
      .filter(item => item.rect.y < rowBottom && item.rect.y + item.rect.h > rowTop);
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

/** Content keeps clear of the page edges and of the header above it, whatever the template's frame says. */
const EDGE_MARGIN = 1.2;
function safeArea(rect: CmRect, slideSize: { w: number; h: number }): CmRect {
  const left = Math.max(rect.x, EDGE_MARGIN);
  const right = Math.min(rect.x + rect.w, slideSize.w - EDGE_MARGIN);
  const top = rect.y + 0.5;
  const bottom = Math.min(rect.y + rect.h, slideSize.h - 0.8);
  return { x: left, y: top, w: Math.max(1, right - left), h: Math.max(1, bottom - top) };
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
  if (frames[0]) return { rect: safeArea(frames[0].rect, slideSize), shapes: [frames[0].id, label.id] };

  // Otherwise: below the title, inside the page's side margins, above the footer.
  const texts = shapes.filter(shape => shape.type === "text" && shape !== label && shape.text);
  const title = texts.find(shape => shape.text?.includes("{{PAGE_TITLE}}")) ?? texts[0];
  const margin = title?.rect.x ?? 1.5;
  const top = title ? title.rect.y + title.rect.h + 0.4 : slideSize.h * 0.2;
  const footer = Math.min(slideSize.h, ...texts.filter(shape => shape.rect.y > slideSize.h * 0.8).map(shape => shape.rect.y));
  return {
    rect: safeArea({ x: margin, y: top, w: slideSize.w - 2 * margin, h: footer - 0.4 - top }, slideSize),
    shapes: [label.id]
  };
}

/** The picture slots of a page: top-level shapes named `@内容.图片`, that are pictures or filled with one. */
function pageImages(shapes: OutlineShape[]): Record<string, ImageSlot> {
  const images: Record<string, ImageSlot> = {};
  shapes.forEach((shape, index) => {
    const binding = parseBinding(shape.name);
    if (!binding?.role || binding.group || binding.area || !binding.field || !isImageShape(shape)) return;
    images[binding.field] ??= imageSlotOf(shape, shapes.length - 1 - index);
  });
  return images;
}

/** Read a template deck's outline: its pages by kind, with what each can take. */
export function inferTemplate(source: Outline): TemplateInfo {
  const { outline, samples } = bindOutline(source);
  const pages: TemplateInfo["pages"] = {};
  const variants: TemplatePage[] = [];
  for (const slide of outline.slides) {
    const shapes = flatten(slide.shapes);
    // A shape bound by its name may be any shape with text (a number in a badge); a keyword in plain text must sit in a text box.
    const texts = shapes.filter(shape => (shape.type === "text" || shape.binding) && shape.text);
    const tokens = new Set(texts.flatMap(shape => tokensIn(shape.text!)));
    const kind = declaredKind(shapes) ?? kindOf(tokens);
    if (!kind) continue;
    // Later pages of a role are layouts only when they name themselves; the rest stay unused, as before.
    const named = declaredVariant(shapes);
    if (pages[kind] && !named.variant) continue;

    const slots: Record<string, TextSlot> = {};
    const fieldNames: Record<string, string> = {};
    const slotOf = (shape: OutlineShape): TextSlot => {
      const sample = samples.get(`${slide.index}:${shape.id}`);
      return {
        id: shape.id, text: shape.text!, rect: shape.rect, fontSizePt: shape.style?.fontSizePt, align: shape.style?.align,
        ...(sample !== undefined ? { sample } : {}), ...capacityOf(shape.rect, shape.style?.fontSizePt)
      };
    };
    const tokenShapes: TextSlot[] = [];
    const fixed: TextSlot[] = [];
    for (const shape of texts) {
      const found = tokensIn(shape.text!);
      if (found.length > 0) {
        tokenShapes.push(slotOf(shape));
        for (const token of found) {
          if (AUTO_TOKENS.has(token) || token === "CONTENT_AREA" || /^TOC_ITEM_/.test(token) || /^LOGO/.test(token)) continue;
          const custom = token.startsWith("U_") ? parseBinding(shape.binding)?.field : undefined;
          if (custom) fieldNames[token] = custom;
          slots[custom ?? slotName(token)] ??= slotOf(shape);
        }
      } else if (!isNumber(shape.text) && /[\p{L}]{3,}|[一-鿿]{2,}/u.test(shape.text!) && !parseBinding(shape.name)?.group) {
        fixed.push(slotOf(shape));
      }
    }
    const page: TemplatePage = {
      kind, ...named, slide: slide.index, tokens: tokenShapes, slots, fieldNames, units: findUnits(slide.shapes), images: pageImages(slide.shapes), fixed, texts: Object.fromEntries(texts.map(shape => [shape.id, slotOf(shape)])),
      rects: Object.fromEntries(shapes.map(shape => [shape.id, shape.rect])), dark: isDark(slide.background)
    };
    if (kind === "toc") {
      page.entries = tocEntries(shapes);
      const inEntries = new Set(page.entries.flatMap(entry => entry.shapes));
      page.fixed = fixed.filter(item => !inEntries.has(item.id));
    }
    if (kind === "content") {
      const frame = shapes.find(shape => parseBinding(shape.name)?.area);
      const area = frame ? { rect: safeArea(frame.rect, outline.slideSize), shapes: [frame.id] } : areaOf(shapes, outline.slideSize);
      page.area = area.rect;
      page.areaShapes = area.shapes;
    }
    pages[kind] ??= page;
    variants.push(page);
  }
  return { slideSize: outline.slideSize, slideCount: outline.slides.length, pages, variants, accent: accentOf(outline.slides) };
}

/** What a model needs to know to fill the template: page kinds, their fields, the room for content. */
export function describeTemplate(template: TemplateInfo): unknown {
  const rounded = (value: number) => Math.round(value * 10) / 10;
  const describePage = (page: TemplatePage) => ({
    ...(page.variant ? { variant: page.variant } : {}),
    ...(page.when ? { when: page.when } : {}),
    fields: Object.keys(page.slots),
    // What each field looks like on the template (its sample text) and about how much fits: write text of a similar length.
    fieldInfo: Object.fromEntries(Object.entries(page.slots).map(([name, slot]) => [name, {
      ...(slot.sample !== undefined ? { sample: slot.sample } : {}),
      fitChars: slot.fitChars, lines: slot.lines,
      ...(slot.fontSizePt ? { fontSizePt: slot.fontSizePt } : {})
    }])),
    ...(page.entries ? {
      maxEntries: page.entries.length,
      entryFields: page.entries.some(entry => entry.desc) ? ["title", "desc"] : ["title"],
      entriesNote: `items 写几条由内容的章节数决定，1 到 ${page.entries.length} 条都可以，不要为了凑满而编造；多出的条目连同编号一起删掉，剩下的自动居中`
    } : {}),
    ...(page.area ? { contentArea: { w: rounded(page.area.w), h: rounded(page.area.h) } } : {}),
    ...(Object.keys(page.images).length > 0 ? {
      images: Object.fromEntries(Object.entries(page.images).map(([name, slot]) => [name, {
        aspect: Math.round(slot.aspect * 100) / 100, ...(slot.shape ? { shape: slot.shape } : {}),
        note: "写图片的路径（project:images/…，用户上传的图片），或 placeholder 先放一张占位图"
      }]))
    } : {}),
    ...(Object.keys(page.units).length > 0 ? { groups: describeUnits(page.units) } : {}),
    ...(page.fixed.length > 0 ? { sampleText: page.fixed.map(item => item.text) } : {})
  });
  return {
    pages: Object.fromEntries(Object.values(template.pages).map(page => {
      const same = template.variants.filter(item => item.kind === page!.kind);
      return [page!.kind, {
        ...describePage(page!),
        // Several layouts of one role: pick one with `variant` (without it the first is used).
        ...(same.length > 1 ? { variants: same.map(describePage), variantNote: `这种页面有 ${same.length} 个版式，按内容选一个，用 variant 指定它的名字；不写 variant 用第一个` } : {})
      }];
    }))
  };
}
