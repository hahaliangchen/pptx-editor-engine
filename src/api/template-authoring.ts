import type { CmRect, Outline, OutlineShape, OutlineSlide } from "./outline";
import {
  AUTO_TOKENS, PAGE_KINDS, PageKind, accentOf, bindOutline, declaredKind, parseBinding, flatten, inferTemplate, describeTemplate, isDark, isNumber, kindOf, tokensIn
} from "./template";

/**
 * Helps turn an ordinary deck into a template: says what is wrong with it
 * (`checkTemplate`), writes the `{{TOKEN}}`s into a page for a role
 * (`suggestPlaceholders`), makes a role page out of any page
 * (`copySlideAsRole`) and reads the facts a design description needs
 * (`templateDesignFacts`). The rules are in docs/template-placeholders.md.
 */

export type Command = Record<string, unknown> & { op: string };

export interface TemplateIssue {
  slide: number;
  /** Only "error" stops a template from being saved. */
  severity: "error" | "warning";
  kind:
    | "missing_role" | "duplicate_role" | "unassigned" | "role_order" | "missing_token" | "role_conflict"
    | "token_in_shape" | "unknown_token" | "renamed_token" | "bad_token" | "multiple_fields"
    | "toc_header" | "toc_numbering" | "toc_order" | "content_area" | "unit" | "image_slot"
    | "empty_deck" | "page_limit" | "canvas" | "identity" | "geometry";
  /** Objects concerned (ids on that slide). */
  ids: string[];
  message: string;
}

export const ROLE_NAMES: Record<PageKind, string> = { cover: "封面", toc: "目录", section: "章节页", content: "内容页", end: "结束页" };

/** The keywords a role takes (docs/template-keywords.md, "通用字段"); `n` stands for a number. */
const COMMON: Record<PageKind | "any", string[]> = {
  cover: ["TITLE", "SUBTITLE", "AUTHOR", "DATE", "INSTITUTION", "ADVISOR", "PRESENTER", "AUTHOR_EN", "DEPARTMENT", "ORGANIZATION"],
  toc: ["TOC_ITEM_n_TITLE", "TOC_ITEM_n_DESC"],
  section: ["CHAPTER_TITLE", "CHAPTER_DESC", "CHAPTER_SUBTITLE", "CHAPTER_TITLE_EN"],
  content: ["PAGE_TITLE", "CONTENT_AREA", "SOURCE", "KEY_MESSAGE", "SECTION_NAME", "ORG_SHORT"],
  end: ["THANK_YOU", "CONTACT_INFO", "ENDING_SUBTITLE", "COPYRIGHT", "THANK_YOU_EN", "CLOSING_MESSAGE", "CONTACT_EMAIL"],
  any: ["PAGE_NUM", "CHAPTER_NUM", "SECTION_NUM", "TOTAL_PAGES", "LOGO", "DATE", "AUTHOR", "INSTITUTION", "ORGANIZATION"]
};
const KNOWN = new Set(Object.values(COMMON).flat());

/** Spellings to unify, from docs/template-keywords.md. */
const RENAMES: Record<string, string> = {
  COVER_TITLE: "TITLE", COVER_SUBTITLE: "SUBTITLE", ENDING_TITLE: "THANK_YOU", END_SUBTITLE: "ENDING_SUBTITLE",
  ENDING_MESSAGE: "CLOSING_MESSAGE", EMAIL: "CONTACT_EMAIL", SPEAKER_NAME: "PRESENTER"
};

const isKnown = (token: string) => KNOWN.has(token.replace(/\d+/g, "n")) || /^LOGO(_|$)/.test(token);

const round = (value: number) => Math.round(value * 10) / 10;
const area = (rect: CmRect) => rect.w * rect.h;

/** The role of every page, from its keywords; undefined when it has none. */
export function slideRoles(source: Outline): Array<{ slide: number; role?: PageKind }> {
  const { outline } = bindOutline(source);
  return outline.slides.map(slide => {
    const shapes = flatten(slide.shapes);
    const tokens = new Set(shapes.flatMap(shape => (shape.text ? tokensIn(shape.text) : [])));
    const savedRole = slide.template?.role;
    return { slide: slide.index, role: slide.template ? (savedRole && PAGE_KINDS.includes(savedRole as PageKind) ? savedRole as PageKind : undefined) : declaredKind(shapes) ?? kindOf(tokens) };
  });
}

// ------------------------------------------------------------------ the check

/** Validate an HTML-design source, not the retired fixed-slot generator.
 * Roles, tokens, groups and spacing describe the design; none determines whether a page survives publication.
 * Package/assets and protected-object semantics are validated when the PPTX is imported and published.
 */
export function checkTemplate(source: Outline): TemplateIssue[] {
  const issues: TemplateIssue[] = [];
  const add = (slide: number, severity: TemplateIssue["severity"], kind: TemplateIssue["kind"], ids: string[], message: string) =>
    issues.push({ slide, severity, kind, ids, message });
  const size = source.slideSize;
  if (!Number.isFinite(size.w) || !Number.isFinite(size.h) || size.w <= 0 || size.h <= 0 || Math.abs(size.w / size.h - 16 / 9) > .01) {
    add(0, "error", "canvas", [], "模板需要有效的 16:9 画布，设计视图固定为 1280×720");
  }
  if (!source.slides.length) add(0, "error", "empty_deck", [], "模板至少保留一页；单页、空白品牌页均可发布");
  if (source.slides.length > 80) add(0, "error", "page_limit", [], "模板最多保留 80 页");
  const pageIDs = new Set<string>();
  const pageIndexes = new Set<number>();
  for (const slide of source.slides) {
    if (!slide.id || pageIDs.has(slide.id) || !Number.isInteger(slide.index) || slide.index < 1 || pageIndexes.has(slide.index)) {
      add(slide.index, "error", "identity", [], "页面身份或编号重复/无效，请重新导入模板");
    }
    pageIDs.add(slide.id); pageIndexes.add(slide.index);
    const ids = new Set<string>();
    for (const shape of flatten(slide.shapes)) {
      if (!shape.id || ids.has(shape.id)) add(slide.index, "error", "identity", [shape.id], "页面中有重复或空的对象 ID，请重新导入模板");
      ids.add(shape.id);
      if (!Object.values(shape.rect).every(Number.isFinite) || shape.rect.w < 0 || shape.rect.h < 0 || !Number.isFinite(shape.rotation)) {
        add(slide.index, "error", "geometry", [shape.id], "对象的坐标、尺寸或旋转值无效");
      }
      // Optional legacy token spelling hints remain useful; HTML can edit text in any supported shape.
      for (const match of (shape.text ?? "").matchAll(/\{\{([^{}]*)\}\}/g)) {
        const token = match[1];
        if (!/^[A-Z0-9_]+$/.test(token)) {
          add(slide.index, "warning", "bad_token", [shape.id], `“{{${token}}}”不是标准占位写法；可改成真实示例，并通过用途说明表达替换意图`);
        } else if (RENAMES[token]) {
          add(slide.index, "warning", "renamed_token", [shape.id], `{{${token}}} 建议统一写成 {{${RENAMES[token]}}}`);
        } else if (!isKnown(token) && !token.startsWith("U_")) {
          add(slide.index, "warning", "unknown_token", [shape.id], `{{${token}}} 是自定义占位，请补充用途说明以便 AI 理解`);
        }
      }
    }
  }
  return issues;
}

// ------------------------------------------------------- suggesting keywords

interface Candidate {
  shape: OutlineShape;
  size: number;
  /** A painted shape with text in it: the text is taken out, the shape stays as decoration. */
  onShape: boolean;
}

const DATE_RE = /(\d{4}\s*[年\-/.]\s*\d{1,2}|\d{4}\s*年|20\d\d|\d{1,2}\s*月\s*\d{1,2}\s*日)/;
const slideOf = (outline: Outline, slide: number): OutlineSlide => {
  const found = outline.slides.find(item => item.index === slide);
  if (!found) throw new Error(`没有第 ${slide} 页，当前共 ${outline.slides.length} 页`);
  return found;
};

const targetOf = (slide: number, shape: OutlineShape) => ({ slide, id: shape.id });
const sizeOf = (shape: OutlineShape) => shape.style?.fontSizePt ?? 18;

/** Largest first; for equal sizes the one higher up and further left. */
const byProminence = (a: Candidate, b: Candidate) =>
  b.size - a.size || a.shape.rect.y - b.shape.rect.y || a.shape.rect.x - b.shape.rect.x;

function textStyleOf(shape: OutlineShape, dark: boolean, sizePt?: number) {
  const style = shape.style;
  return {
    fontSizePt: sizePt ?? style?.fontSizePt ?? 18,
    ...(style?.bold ? { bold: true } : {}),
    color: style?.color ?? (dark ? "#ffffff" : "#333333"),
    ...(style?.align ? { align: style.align } : {})
  };
}

/**
 * Commands that write the keywords of `role` into a page: the text boxes that
 * match get the keyword (position and look unchanged), sample text goes, and
 * the keywords still missing are added as new boxes. Nothing is removed from
 * the background and decoration.
 */
export function suggestPlaceholders(outline: Outline, slideIndex: number, role: PageKind): Command[] {
  const slide = slideOf(outline, slideIndex);
  const { w: pageW, h: pageH } = outline.slideSize;
  const dark = isDark(slide.background);
  const commands: Command[] = [];
  const shapes = flatten(slide.shapes);
  const existing = new Set(shapes.flatMap(shape => (shape.text ? tokensIn(shape.text) : [])));
  const handled = new Set<string>();

  const candidates: Candidate[] = shapes
    .filter(shape => shape.text && tokensIn(shape.text).length === 0 && (shape.type === "text" || shape.type === "shape"))
    .map(shape => ({ shape, size: sizeOf(shape), onShape: shape.type === "shape" }));

  /** Put `text` (with its keyword) where a candidate stands. */
  const assign = (candidate: Candidate, text: string) => {
    handled.add(candidate.shape.id);
    existing.add(tokensIn(text)[0] ?? "");
    if (candidate.onShape) {
      commands.push({ op: "setText", target: targetOf(slideIndex, candidate.shape), text: "" });
      commands.push({
        op: "addText", slide: slideIndex, rect: candidate.shape.rect, text,
        style: textStyleOf(candidate.shape, dark), name: "占位符"
      });
    } else {
      commands.push({ op: "setText", target: targetOf(slideIndex, candidate.shape), text });
    }
  };
  const free = () => candidates.filter(candidate => !handled.has(candidate.shape.id));
  const take = (token: string, pick: (list: Candidate[]) => Candidate | undefined, text = `{{${token}}}`) => {
    if (existing.has(token)) return undefined;
    const found = pick(free());
    if (found) assign(found, text);
    return found;
  };
  /** The next biggest text under `anchor`, smaller than it. */
  const below = (anchor: Candidate | undefined) => (list: Candidate[]) => anchor
    ? list.filter(c => c.shape.rect.y >= anchor.shape.rect.y && c.size <= anchor.size && c.shape.rect.y - (anchor.shape.rect.y + anchor.shape.rect.h) < 4)
      .sort(byProminence)[0]
    : undefined;
  const addBox = (token: string, rect: CmRect, sizePt: number, align: "left" | "center" | "right" = "left", text = `{{${token}}}`) => {
    if (existing.has(token)) return;
    existing.add(token);
    commands.push({
      op: "addText", slide: slideIndex, rect, text, name: "占位符",
      style: { fontSizePt: sizePt, color: dark ? "#ffffff" : "#333333", align }
    });
  };
  // Sizes for new boxes follow the deck's own: its biggest and its usual text.
  const allSizes = outline.slides.flatMap(item => flatten(item.shapes).filter(shape => shape.text).map(sizeOf)).sort((a, b) => a - b);
  const titleSize = Math.min(44, Math.max(24, allSizes[allSizes.length - 1] ?? 32));
  const bodySize = Math.max(12, Math.min(20, allSizes[Math.floor(allSizes.length / 2)] ?? 16));
  const margin = Math.max(1, round(pageW * 0.06));
  const dateOf = (list: Candidate[]) => list.find(c => DATE_RE.test(c.shape.text!) && c.shape.text!.length < 24);

  switch (role) {
    case "cover": {
      const title = take("TITLE", list => [...list].sort(byProminence)[0]);
      take("SUBTITLE", below(title));
      take("DATE", dateOf);
      addBox("TITLE", { x: margin, y: pageH * 0.35, w: pageW - 2 * margin, h: 2.4 }, titleSize);
      addBox("SUBTITLE", { x: margin, y: pageH * 0.35 + 2.8, w: pageW - 2 * margin, h: 1.4 }, bodySize + 4);
      break;
    }
    case "end": {
      const title = take("THANK_YOU", list => [...list].sort(byProminence)[0]);
      take("ENDING_SUBTITLE", below(title));
      take("DATE", dateOf);
      addBox("THANK_YOU", { x: margin, y: pageH * 0.35, w: pageW - 2 * margin, h: 2.4 }, titleSize);
      break;
    }
    case "section": {
      const title = take("CHAPTER_TITLE", list => [...list].sort(byProminence)[0]);
      // "01", "PART 02" and the like become the chapter number.
      const numberOf = (list: Candidate[]) => list.find(c => /^(part\s*|第\s*)?\d{1,2}\s*(章|部分)?$/i.test(c.shape.text!.trim()));
      const number = numberOf(free());
      if (number && !existing.has("CHAPTER_NUM")) {
        const prefix = number.shape.text!.trim().match(/^(part\s*|第\s*)/i)?.[0] ?? "";
        const suffix = number.shape.text!.trim().match(/\s*(章|部分)$/)?.[0] ?? "";
        assign(number, `${prefix}{{CHAPTER_NUM}}${suffix}`);
      }
      take("CHAPTER_DESC", below(title));
      addBox("CHAPTER_TITLE", { x: margin, y: pageH * 0.38, w: pageW - 2 * margin, h: 2.2 }, titleSize);
      break;
    }
    case "content": {
      const top = free().filter(c => c.shape.rect.y < pageH * 0.25).sort(byProminence)[0] ?? [...free()].sort(byProminence)[0];
      const title = take("PAGE_TITLE", () => top);
      take("PAGE_NUM", list => list.find(c => isNumber(c.shape.text) && c.shape.rect.y > pageH * 0.8), "{{PAGE_NUM}}");
      const titleRect = title?.shape.rect;
      const left = titleRect?.x ?? margin;
      const frameTop = titleRect ? titleRect.y + titleRect.h + 0.4 : pageH * 0.2;
      const footer = Math.min(pageH * 0.94, ...free().filter(c => c.shape.rect.y > pageH * 0.85).map(c => c.shape.rect.y));
      const frame: CmRect = { x: left, y: frameTop, w: pageW - 2 * left, h: Math.max(4, footer - 0.4 - frameTop) };
      if (!existing.has("CONTENT_AREA")) {
        // Whatever sample content sits in the area goes: text, tables, charts and big pictures.
        const inside = (rect: CmRect) => {
          const w = Math.min(rect.x + rect.w, frame.x + frame.w) - Math.max(rect.x, frame.x);
          const h = Math.min(rect.y + rect.h, frame.y + frame.h) - Math.max(rect.y, frame.y);
          return w > 0 && h > 0 && w * h > 0.5 * area(rect);
        };
        for (const shape of slide.shapes) {
          if (handled.has(shape.id)) continue;
          if (shape.type === "table" || shape.type === "chart" || (shape.type === "image" && isContentImage(shape.rect, outline.slideSize))) {
            if (inside(shape.rect)) commands.push({ op: "delete", target: targetOf(slideIndex, shape) });
          }
        }
        for (const candidate of free()) {
          if (inside(candidate.shape.rect)) {
            handled.add(candidate.shape.id);
            commands.push(candidate.onShape
              ? { op: "setText", target: targetOf(slideIndex, candidate.shape), text: "" }
              : { op: "delete", target: targetOf(slideIndex, candidate.shape) });
          }
        }
        commands.push({
          op: "addShape", slide: slideIndex, shape: "rect", rect: frame, fill: null, name: "内容区",
          line: { color: dark ? "#cccccc" : "#999999", widthPt: 1, dash: "dash" }
        });
        commands.push({
          op: "addText", slide: slideIndex, name: "占位符", text: "{{CONTENT_AREA}}",
          rect: { x: frame.x + frame.w / 2 - 3.5, y: frame.y + frame.h / 2 - 0.6, w: 7, h: 1.2 },
          style: { fontSizePt: bodySize, color: dark ? "#cccccc" : "#999999", align: "center" }
        });
        existing.add("CONTENT_AREA");
      }
      addBox("PAGE_TITLE", { x: left, y: Math.max(0.6, pageH * 0.05), w: pageW - 2 * left, h: 1.6 }, Math.min(32, titleSize));
      break;
    }
    case "toc":
      suggestToc(slide, candidates, { assign, free, handled, existing, addBox, commands, dark, margin, pageW, pageH, titleSize, bodySize, slideIndex });
      break;
  }

  // The other sample wording goes.
  for (const candidate of candidates) {
    if (handled.has(candidate.shape.id)) continue;
    if (role === "toc" && (isNumber(candidate.shape.text) || candidate.shape.text === (tocHeaderOf(candidates)?.shape.text))) continue;
    commands.push(candidate.onShape
      ? { op: "setText", target: targetOf(slideIndex, candidate.shape), text: "" }
      : { op: "delete", target: targetOf(slideIndex, candidate.shape) });
  }
  return commands;
}

/** Pictures in the middle of a page that are not its frame or its logo. */
function isContentImage(rect: CmRect, size: { w: number; h: number }): boolean {
  const edge = 0.6;
  const touches = rect.x < edge || rect.y < edge || rect.x + rect.w > size.w - edge || rect.y + rect.h > size.h - edge;
  const share = area(rect) / (size.w * size.h);
  return !touches && share >= 0.08;
}

function tocHeaderOf(candidates: Candidate[]): Candidate | undefined {
  // The header is the top-most text that is not a number.
  return [...candidates].filter(c => !isNumber(c.shape.text)).sort((a, b) => a.shape.rect.y - b.shape.rect.y)[0];
}

interface TocContext {
  assign(candidate: Candidate, text: string): void;
  free(): Candidate[];
  handled: Set<string>;
  existing: Set<string>;
  addBox(token: string, rect: CmRect, sizePt: number, align?: "left" | "center" | "right", text?: string): void;
  commands: Command[];
  dark: boolean;
  margin: number;
  pageW: number;
  pageH: number;
  titleSize: number;
  bodySize: number;
  slideIndex: number;
}

/** Entries are the repeated texts: the biggest size is the item titles, the text under or beside each is its description. */
function suggestToc(slide: OutlineSlide, candidates: Candidate[], ctx: TocContext): void {
  const header = tocHeaderOf(candidates);
  const rest = candidates.filter(c => c !== header && !isNumber(c.shape.text));
  const top = Math.max(0, ...rest.map(c => c.size));
  const titles = rest.filter(c => c.size >= top - 0.5 && c.shape.text!.length <= 40)
    .sort((a, b) => a.shape.rect.y - b.shape.rect.y || a.shape.rect.x - b.shape.rect.x);
  // A single big header is not an entry.
  const entries = header && titles.length === 1 && header.size > titles[0].size ? [] : titles;
  if (header) ctx.handled.add(header.shape.id);
  entries.forEach((entry, i) => {
    ctx.assign(entry, `{{TOC_ITEM_${i + 1}_TITLE}}`);
    const { rect } = entry.shape;
    const desc = ctx.free()
      .filter(c => !isNumber(c.shape.text) && c.size < entry.size
        && ((c.shape.rect.y >= rect.y + rect.h * 0.5 && c.shape.rect.y - (rect.y + rect.h) < Math.max(1.5, rect.h * 1.5)
            && c.shape.rect.x < rect.x + rect.w && c.shape.rect.x + c.shape.rect.w > rect.x)
          || (Math.abs(c.shape.rect.y - rect.y) < rect.h && c.shape.rect.x >= rect.x + rect.w * 0.5)))
      .sort((a, b) => a.shape.rect.y - b.shape.rect.y)[0];
    if (desc) ctx.assign(desc, `{{TOC_ITEM_${i + 1}_DESC}}`);
  });
  if (entries.length === 0) {
    // Nothing to build on: five lines under the header.
    const left = ctx.margin * 1.5;
    for (let i = 1; i <= 5; i++) {
      ctx.addBox(`TOC_ITEM_${i}_TITLE`, { x: left, y: ctx.pageH * 0.25 + (i - 1) * 2.2, w: ctx.pageW - 2 * left, h: 1.4 }, ctx.bodySize + 4, "left");
    }
  }
  if (!header) {
    ctx.commands.push({
      op: "addText", slide: ctx.slideIndex, rect: { x: ctx.margin, y: ctx.pageH * 0.07, w: 10, h: 1.8 }, text: "目录", name: "目录标题",
      style: { fontSizePt: Math.min(36, ctx.titleSize), bold: true, color: ctx.dark ? "#ffffff" : "#333333", align: "left" }
    });
    // Below everything, so it comes before the first entry.
    ctx.commands.push({ op: "arrange", target: { slide: ctx.slideIndex, name: "目录标题" }, mode: "back" });
  }
}

// ------------------------------------------------------- a role from any page

/**
 * Make a page of `role` out of page `slide`: a copy right after it, with
 * everything that was content (text, tables, charts, pictures in the middle)
 * taken off and the keywords of the role written in. Backgrounds, shapes
 * without text, logos and edge pictures stay.
 */
export function copySlideAsRole(outline: Outline, slideIndex: number, role: PageKind, options: { moveTo?: number } = {}): Command[] {
  const slide = slideOf(outline, slideIndex);
  const copy = slideIndex + 1;
  const commands: Command[] = [{ op: "duplicateSlide", slide: slideIndex }];

  // The copy's shapes keep the ids of the original, so what is cleared is named by id.
  const clear = (shapes: OutlineShape[]) => {
    for (const shape of shapes) {
      if (shape.name === "内容区" || shape.type === "table" || shape.type === "chart" || (shape.type === "text" && shape.text !== undefined)
        || (shape.type === "image" && isContentImage(shape.rect, outline.slideSize))) {
        commands.push({ op: "delete", target: targetOf(copy, shape) });
      } else if (shape.type === "shape" && shape.text) {
        commands.push({ op: "setText", target: targetOf(copy, shape), text: "" });
      } else if (shape.type === "group" && shape.children) {
        clear(shape.children);
      }
    }
  };
  clear(slide.shapes);

  // What `suggestPlaceholders` needs is a page without content: describe the copy as it will be.
  const cleared = JSON.parse(JSON.stringify(outline)) as Outline;
  const target = cleared.slides.find(item => item.index === slideIndex)!;
  const strip = (shapes: OutlineShape[]): OutlineShape[] => shapes.flatMap(shape => {
    if (shape.name === "内容区" || shape.type === "table" || shape.type === "chart" || (shape.type === "text" && shape.text !== undefined)
      || (shape.type === "image" && isContentImage(shape.rect, outline.slideSize))) return [];
    if (shape.type === "shape" && shape.text) delete shape.text;
    if (shape.children) shape.children = strip(shape.children);
    return [shape];
  });
  target.shapes = strip(target.shapes);
  cleared.slides = cleared.slides.flatMap(item => (item.index === slideIndex ? [item, { ...target, index: copy }] : [item]));
  cleared.slides.forEach((item, i) => { item.index = i + 1; });

  commands.push(...suggestPlaceholders(cleared, copy, role));
  if (options.moveTo !== undefined) commands.push({ op: "moveSlide", slide: copy, to: options.moveTo });
  return commands;
}

// --------------------------------------------------------- design description

export interface TemplateDesignFacts {
  /** The same purpose definitions generation reads, including real samples and author options. */
  templateDefinition: unknown;
  /** Page size in cm. */
  pageSize: { w: number; h: number };
  accent: string;
  /** Colours used by fills, strongest first. */
  colors: string[];
  fonts: Array<{ role: PageKind; sizes: { title?: number; body?: number } }>;
  /** Content area in cm, and the margins around it. */
  contentArea?: { w: number; h: number; margin: number };
  tocItems?: number;
  /** Per role, whether the page is dark. */
  dark: Partial<Record<PageKind, boolean>>;
  /** Every page of the template, with the layout name and when to use it when the role has several. */
  pages: Array<{ role?: string; slide: number; id: string; variant?: string; when?: string; objects: Array<{ id: string; name: string; purpose?: string; note?: string }> }>;
}

/** The numbers a design description is written from; a model adds only the words. */
export function templateDesignFacts(outline: Outline): TemplateDesignFacts {
  const info = inferTemplate(outline);
  const weight = new Map<string, number>();
  for (const slide of outline.slides) {
    for (const shape of flatten(slide.shapes)) {
      if (typeof shape.fill === "string") {
        weight.set(shape.fill.toLowerCase(), (weight.get(shape.fill.toLowerCase()) ?? 0) + Math.max(0.01, area(shape.rect)));
      }
    }
  }
  const pages = Object.values(info.pages).map(page => page!);
  return {
    templateDefinition: describeTemplate(info),
    pageSize: { w: round(outline.slideSize.w), h: round(outline.slideSize.h) },
    accent: accentOf(outline.slides),
    colors: [...weight.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6).map(([color]) => color),
    fonts: pages.map(page => {
      const sizes = Object.values(page.texts).map(slot => slot.fontSizePt ?? 18).sort((a, b) => b - a);
      return { role: page.kind, sizes: { title: sizes[0], body: sizes[Math.floor(sizes.length / 2)] } };
    }),
    ...(info.pages.content?.area
      ? { contentArea: { w: round(info.pages.content.area.w), h: round(info.pages.content.area.h), margin: round(info.pages.content.area.x) } }
      : {}),
    ...(info.pages.toc?.entries ? { tocItems: info.pages.toc.entries.length } : {}),
    dark: Object.fromEntries(pages.map(page => [page.kind, page.dark])),
    pages: outline.slides.map(slide => {
      const legacy = info.variants.find(page => page.slide === slide.index);
      return { id: slide.id, slide: slide.index, role: slide.template ? slide.template.role : legacy?.kind,
        variant: slide.template ? slide.template.layoutName : legacy?.variant, when: slide.template ? slide.template.note : legacy?.when,
        objects: flatten(slide.shapes).map(shape => ({id: shape.id, name: shape.name, purpose: shape.template?.purpose, note: shape.template?.note})) };
    })
  };
}

/** Issues as lines a model can act on. */
export function describeTemplateIssues(issues: TemplateIssue[]): string {
  return issues.map(issue => `${issue.slide > 0 ? `第 ${issue.slide} 页` : "整份"}［${issue.severity === "error" ? "必须改" : "建议"}］${issue.message}`).join("\n");
}
