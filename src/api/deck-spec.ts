import { ApiError } from "./errors";
import type { CmRect } from "./outline";
import { AUTO_TOKENS, PAGE_KINDS, PageKind, TemplateInfo, TemplatePage, TextSlot, slotName, tokensIn } from "./template";

/**
 * A whole deck as one piece of data: which template page each page is, and the
 * words and figures to put on it. `compileDeck` turns it into the commands that
 * build it from the template's pages, so a model gives content and the engine
 * does the layout.
 */

export type Block =
  | { type: "text"; text: string }
  | { type: "bullets"; items: string[] }
  | { type: "cards"; items: Array<{ title?: string; value?: string; text?: string }> }
  | { type: "chart"; kind: "column" | "bar" | "line" | "pie"; categories: string[]; series: Array<{ name: string; values: Array<number | null> }>; title?: string }
  | { type: "table"; rows: string[][] };

export interface PageSpec {
  page: PageKind;
  /** The contents page: one entry per line. */
  items?: Array<string | { title: string; desc?: string }>;
  /** The content page. */
  blocks?: Block[];
  layout?: "rows" | "columns";
  /** Sample wording on the template page → what to write instead ("" removes it). */
  replace?: Record<string, string>;
  /** Everything else is a field of the page: title, subtitle, author, date, ... */
  [field: string]: unknown;
}

export interface DeckSpec {
  slides: PageSpec[];
  /** Applied to every page. */
  replace?: Record<string, string>;
}

const BLOCK_FIELDS = "text、bullets、cards、chart、table";
const RESERVED = new Set(["page", "items", "blocks", "layout", "replace"]);

const fail = (message: string, path: string): never => {
  throw new ApiError("INVALID_SCHEMA", message, path);
};

// ---------------------------------------------------------------- text fitting

const CM_TO_PT = 28.3465;
const INSET_CM = 0.25;

/** How wide a text is, in em: Chinese characters are one em wide, Latin ones about half. */
const widthEm = (text: string) => [...text].reduce((sum, char) => sum + (/[⺀-鿿＀-￯]/.test(char) ? 1 : 0.55), 0);

function linesOf(text: string, sizePt: number, widthCm: number): number {
  const perLine = Math.max(1, ((widthCm - 2 * INSET_CM) * CM_TO_PT) / sizePt);
  return Math.max(1, text.split("\n").reduce((sum, line) => sum + Math.max(1, Math.ceil(widthEm(line) / perLine)), 0));
}

/** The largest size, in points, at which these paragraphs fit the box. */
export function fitFontSize(paragraphs: string[], box: { w: number; h: number }, maxPt: number, minPt = 9, lineGap = 1.3, paragraphGapPt = 0): number {
  for (let size = maxPt; size > minPt; size -= 0.5) {
    const lines = paragraphs.reduce((sum, text) => sum + linesOf(text, size, box.w), 0);
    const height = lines * size * lineGap + paragraphGapPt * paragraphs.length + 2 * INSET_CM * CM_TO_PT * 0.4;
    if (height <= box.h * CM_TO_PT) return size;
  }
  return minPt;
}

/**
 * Commands that make `text` fit the one-line box of a template shape. The box
 * grows first, in the direction its alignment leaves free but at most to 1.8
 * times its width (so it does not run over the template's decoration), with the
 * size shrinking down to 70% to make that enough; if it still does not fit, the
 * box may use all the room to the page margin, and the size shrinks as needed.
 */
function fitInBox(slot: TextSlot, text: string, id: string, slide: number, pageWidth: number): Command[] {
  const original = slot.fontSizePt ?? 18;
  const widest = Math.max(...text.split("\n").map(widthEm));
  const need = (size: number) => (widest * size) / CM_TO_PT + 2 * INSET_CM;
  const margin = 0.5;
  const align = slot.align ?? "left";
  const { x: x0, w: w0 } = slot.rect;
  const middle = x0 + w0 / 2;
  const room = align === "right" ? x0 + w0 - margin
    : align === "center" ? 2 * Math.min(middle - margin, pageWidth - margin - middle)
      : pageWidth - margin - x0;

  let size = original;
  const cap = Math.max(w0, Math.min(room, w0 * 1.8));
  while (need(size) > cap && size > original * 0.7) size -= 0.5;
  if (need(size) > cap) {
    while (need(size) > room && size > 9) size -= 0.5;
  }
  const w = Math.max(w0, Math.min(need(size), room));
  const x = align === "right" ? x0 + w0 - w : align === "center" ? middle - w / 2 : x0;

  const target = { slide, id };
  const commands: Command[] = [];
  if (Math.abs(w - w0) > 0.01) commands.push({ op: "setBounds", target, x: round(x), w: round(w) });
  if (size !== original) commands.push({ op: "formatText", target, fontSizePt: size });
  return commands;
}

// ---------------------------------------------------------------- validation

function textOf(value: unknown, path: string): string {
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  return fail("这里需要一段文字", path);
}

function validateBlock(block: unknown, path: string): Block {
  if (!block || typeof block !== "object") return fail("每个内容块都是一个对象", path);
  const raw = block as Record<string, unknown>;
  const list = (field: string) => {
    if (!Array.isArray(raw[field]) || (raw[field] as unknown[]).length === 0) fail(`${field} 需要至少一项`, `${path}.${field}`);
    return raw[field] as unknown[];
  };
  switch (raw.type) {
    case "text": return { type: "text", text: textOf(raw.text, `${path}.text`) };
    case "bullets": return { type: "bullets", items: list("items").map((item, i) => textOf(item, `${path}.items[${i}]`)) };
    case "cards": {
      const items = list("items").map((item, i) => {
        const card = (item ?? {}) as Record<string, unknown>;
        const pick = (field: string) => (card[field] === undefined ? undefined : textOf(card[field], `${path}.items[${i}].${field}`));
        const parsed = { title: pick("title"), value: pick("value"), text: pick("text") };
        if (!parsed.title && !parsed.value && !parsed.text) fail("卡片至少要有 title、value、text 之一", `${path}.items[${i}]`);
        return parsed;
      });
      if (items.length > 6) fail("卡片最多 6 张", `${path}.items`);
      return { type: "cards", items };
    }
    case "chart": {
      if (!["column", "bar", "line", "pie"].includes(raw.kind as string)) fail('kind 必须是 "column"、"bar"、"line" 或 "pie"', `${path}.kind`);
      const categories = list("categories").map((item, i) => textOf(item, `${path}.categories[${i}]`));
      const series = list("series").map((item, i) => {
        const entry = (item ?? {}) as Record<string, unknown>;
        if (!Array.isArray(entry.values) || entry.values.length !== categories.length) {
          fail(`每个系列的 values 要和 categories 一样多（${categories.length} 个）`, `${path}.series[${i}].values`);
        }
        return { name: textOf(entry.name, `${path}.series[${i}].name`), values: entry.values as Array<number | null> };
      });
      return { type: "chart", kind: raw.kind as "column", categories, series, ...(typeof raw.title === "string" ? { title: raw.title } : {}) };
    }
    case "table": {
      const rows = list("rows").map((row, i) => {
        if (!Array.isArray(row) || row.length === 0) fail("每一行都是文字组成的数组", `${path}.rows[${i}]`);
        return (row as unknown[]).map((cell, j) => textOf(cell, `${path}.rows[${i}][${j}]`));
      });
      return { type: "table", rows };
    }
    default:
      return fail(`未知的内容块 type，可用的有：${BLOCK_FIELDS}`, `${path}.type`);
  }
}

// ---------------------------------------------------------------- blocks

interface Palette { text: string; muted: string; accent: string; card: string }

const paletteOf = (template: TemplateInfo, dark: boolean): Palette => ({
  text: dark ? "#ffffff" : "#333333",
  muted: dark ? "#d0d7e2" : "#666666",
  accent: template.accent,
  card: dark ? "#ffffff" : "#f5f7fa"
});

type Command = Record<string, unknown>;

const mix = (hex: string, other: number, amount: number) =>
  "#" + [1, 3, 5].map(i => Math.round(parseInt(hex.slice(i, i + 2), 16) * (1 - amount) + other * amount).toString(16).padStart(2, "0")).join("");

/** Series colours: the template's accent, then darker and lighter versions of it, then greys. */
const seriesColors = (accent: string): string[] => [accent, mix(accent, 0, 0.4), mix(accent, 255, 0.45), mix(accent, 0, 0.7), mix(accent, 255, 0.7), "#8c9aa8"];

const round = (value: number) => Math.round(value * 100) / 100;
const rectOf = (rect: CmRect) => ({ x: round(rect.x), y: round(rect.y), w: round(rect.w), h: round(rect.h) });

/** Splits a rectangle into `weights.length` parts along one direction with a gap between them. */
function split(rect: CmRect, weights: number[], direction: "rows" | "columns", gap: number): CmRect[] {
  const total = weights.reduce((sum, weight) => sum + weight, 0);
  const room = (direction === "rows" ? rect.h : rect.w) - gap * (weights.length - 1);
  let offset = 0;
  return weights.map(weight => {
    const size = (room * weight) / total;
    const part = direction === "rows"
      ? { x: rect.x, y: rect.y + offset, w: rect.w, h: size }
      : { x: rect.x + offset, y: rect.y, w: size, h: rect.h };
    offset += size + gap;
    return part;
  });
}

function blockCommands(block: Block, rect: CmRect, slide: number, palette: Palette): Command[] {
  switch (block.type) {
    case "text": {
      const size = fitFontSize([block.text], rect, 24, 12);
      return [{ op: "addText", slide, rect: rectOf(rect), text: block.text, style: { fontSizePt: size, color: palette.text } }];
    }
    case "bullets": {
      const items = block.items.map(item => `• ${item}`);
      const size = fitFontSize(items, rect, 24, 12, 1.3, 8);
      return [{
        op: "addText", slide, rect: rectOf(rect), paragraphs: items.map(text => ({ text })),
        style: { fontSizePt: size, color: palette.text }
      }];
    }
    case "cards": {
      const count = block.items.length;
      const columns = count <= 4 ? count : Math.ceil(count / 2);
      const rows = Math.ceil(count / columns);
      const gap = 0.5;
      const cells = split(rect, Array(rows).fill(1), "rows", gap)
        .flatMap(row => split(row, Array(columns).fill(1), "columns", gap));
      // Cards are no taller than their contents need: the row is not stretched over the whole area.
      const commands: Command[] = [];
      const cardHeight = Math.min(cells[0].h, 7.5);
      block.items.forEach((card, index) => {
        const cell = cells[index];
        const body = { ...cell, h: cardHeight };
        commands.push({
          op: "addShape", slide, shape: "roundRect", rect: rectOf(body),
          fill: palette.card, opacity: palette.card === "#ffffff" ? 0.12 : 1,
          line: { color: palette.accent, widthPt: 1 }
        });
        // value / title / text stacked inside the card.
        const parts: Array<{ text: string; share: number; max: number; bold: boolean; color: string }> = [];
        if (card.value) parts.push({ text: card.value, share: 0.38, max: 36, bold: true, color: palette.accent });
        if (card.title) parts.push({ text: card.title, share: 0.22, max: 18, bold: true, color: palette.text });
        if (card.text) parts.push({ text: card.text, share: 0.4, max: 14, bold: false, color: palette.muted });
        const shares = parts.reduce((sum, part) => sum + part.share, 0);
        let y = body.y + 0.3;
        const usable = body.h - 0.6;
        for (const part of parts) {
          const h = (usable * part.share) / shares;
          const box = { x: body.x + 0.2, y, w: body.w - 0.4, h };
          const size = fitFontSize([part.text], box, part.max, 9);
          commands.push({
            op: "addText", slide, rect: rectOf(box), text: part.text,
            style: { fontSizePt: size, bold: part.bold, color: part.color, align: "center" }
          });
          y += h;
        }
      });
      return commands;
    }
    case "chart":
      return [{
        op: "addChart", slide, type: block.kind, categories: block.categories,
        series: block.series.map((series, i) => ({
          name: series.name, values: series.values,
          // A pie takes its slice colours from the chart's own palette.
          ...(block.kind === "pie" ? {} : { color: seriesColors(palette.accent)[i % 6] })
        })),
        ...(block.title ? { title: block.title } : {}),
        legend: block.kind === "pie" || block.series.length > 1 ? "bottom" : "none",
        dataLabels: block.kind === "pie" || block.series.length === 1,
        textColor: palette.text, rect: rectOf(rect)
      }];
    case "table": {
      const widest = Math.max(...block.rows.map(row => row.length));
      const cellBox = { w: rect.w / widest, h: rect.h / block.rows.length };
      const size = fitFontSize(block.rows.flat().map(cell => cell), cellBox, 16, 10, 1.3);
      return [{ op: "addTable", slide, rows: block.rows.map(row => [...row, ...Array(widest - row.length).fill("")]), fontSizePt: size, rect: rectOf(rect) }];
    }
  }
}

function contentCommands(page: PageSpec, area: CmRect, slide: number, palette: Palette, path: string): Command[] {
  const blocks = (page.blocks ?? []).map((block, i) => validateBlock(block, `${path}.blocks[${i}]`));
  if (blocks.length === 0) return [];
  if (blocks.length > 3) fail("一页最多 3 个内容块；内容多就拆成两页", `${path}.blocks`);
  const direction = page.layout ?? (blocks.length === 2 && !blocks.some(block => block.type === "cards") ? "columns" : "rows");
  // A row of cards needs less height than text or a chart.
  const weights = blocks.map(block => (direction === "rows" && block.type === "cards" ? 0.7 : 1));
  const parts = split(area, weights, direction, 0.6);
  return blocks.flatMap((block, i) => blockCommands(block, parts[i], slide, palette));
}

// ---------------------------------------------------------------- the deck

export interface CompiledDeck {
  label: string;
  atomic: true;
  commands: Command[];
}

/**
 * The commands that build `spec` from `template`. The template's own pages stay
 * where they are while the new ones are copied after them, then are removed, so
 * a page's number never depends on how many were made before it.
 */
export function compileDeck(spec: unknown, template: TemplateInfo): CompiledDeck {
  const deck = spec as DeckSpec;
  if (!deck || typeof deck !== "object" || !Array.isArray(deck.slides) || deck.slides.length === 0) {
    return fail("需要 slides：每页一个对象的非空数组", "slides");
  }
  const shared = deck.replace ?? {};
  const commands: Command[] = [];
  const last = template.slideCount;
  let total = last;
  let sections = 0;

  deck.slides.forEach((page, index) => {
    const path = `slides[${index}]`;
    if (!page || typeof page !== "object") fail("每页都是一个对象", path);
    const source = template.pages[page.page];
    if (!source) {
      const available = PAGE_KINDS.filter(kind => template.pages[kind]);
      fail(`模板没有 "${String(page.page)}" 这种页面，可用的有：${available.join("、")}`, `${path}.page`);
    }
    const kind = page.page;
    const tpl = source as TemplatePage;
    const number = deck.slides.length;

    for (const field of Object.keys(page)) {
      if (RESERVED.has(field) || field in tpl.slots) continue;
      fail(`"${kind}" 页没有 "${field}" 这个字段，它有：${Object.keys(tpl.slots).join("、") || "（无）"}`, `${path}.${field}`);
    }

    // 1. A copy of the template page at the end.
    commands.push({ op: "duplicateSlide", slide: tpl.slide });
    commands.push({ op: "moveSlide", slide: tpl.slide + 1, to: total + 1 });
    total += 1;
    const at = total;
    const target = (id: string) => ({ slide: at, id });
    const write = (id: string, text: string) => {
      commands.push({ op: "setText", target: target(id), text });
      const slot = tpl.texts[id];
      if (slot) commands.push(...fitInBox(slot, text, id, at, template.slideSize.w));
    };

    if (kind === "section") sections += 1;
    const auto: Record<string, string> = {
      PAGE_NUM: String(index + 1),
      TOTAL_PAGES: String(number),
      CHAPTER_NUM: String(typeof page.number === "string" || typeof page.number === "number" ? page.number : String(sections).padStart(2, "0")),
      SECTION_NUM: String(sections).padStart(2, "0")
    };
    const values: Record<string, string> = {};
    for (const name of Object.keys(tpl.slots)) if (page[name] !== undefined) values[name] = textOf(page[name], `${path}.${name}`);
    const replace = { ...shared, ...(page.replace ?? {}) };
    const done = new Set<string>();

    // 2. Contents entries: the given ones are filled, the rest removed with everything that belongs to them.
    if (kind === "toc") {
      const items = (page.items ?? []).map((item, i) => {
        if (typeof item === "string" || typeof item === "number") return { title: String(item) } as { title: string; desc?: string };
        const entry = item as { title?: unknown; desc?: unknown };
        return { title: textOf(entry.title, `${path}.items[${i}].title`), ...(entry.desc ? { desc: textOf(entry.desc, `${path}.items[${i}].desc`) } : {}) };
      });
      const entries = tpl.entries ?? [];
      if (items.length === 0) fail("目录页需要 items", `${path}.items`);
      if (items.length > entries.length) fail(`目录页最多 ${entries.length} 项，现在有 ${items.length} 项`, `${path}.items`);
      // Fewer entries than the template has: the ones kept sit in the middle of the room, not at the top.
      const regionBottom = Math.max(...entries.map(entry => entry.bottom));
      const shift = Math.max(0, (regionBottom - entries[items.length - 1].bottom) / 2);
      entries.forEach((entry, i) => {
        if (i < items.length) {
          if (shift > 0.01) {
            for (const id of entry.shapes) {
              const rect = tpl.rects[id];
              if (rect) commands.push({ op: "setBounds", target: target(id), y: round(rect.y + shift) });
            }
          }
          write(entry.title, items[i].title);
          if (entry.desc) {
            if (items[i].desc) write(entry.desc, items[i].desc as string);
            else commands.push({ op: "delete", target: target(entry.desc) });
          }
        } else {
          for (const id of entry.shapes) commands.push({ op: "delete", target: target(id) });
        }
        for (const id of entry.shapes) done.add(id);
      });
    }

    // 3. Text with tokens: fill from the fields; a shape with nothing to show is removed.
    for (const shape of tpl.tokens) {
      if (done.has(shape.id)) continue;
      let hasValue = false;
      const text = shape.text.replace(/\{\{([A-Z0-9_]+)\}\}/g, (_match, token: string) => {
        const value = AUTO_TOKENS.has(token) ? auto[token] : values[slotName(token)];
        if (value !== undefined && value !== "") hasValue = true;
        return value ?? "";
      });
      done.add(shape.id);
      if (tokensIn(shape.text).every(token => token === "CONTENT_AREA")) continue;
      if (!hasValue || text.trim() === "") commands.push({ op: "delete", target: target(shape.id) });
      else if (text !== shape.text) write(shape.id, text);
    }

    // 4. Sample wording the template carries.
    for (const shape of tpl.fixed) {
      if (done.has(shape.id) || !(shape.text in replace)) continue;
      if (replace[shape.text] === "") commands.push({ op: "delete", target: target(shape.id) });
      else write(shape.id, replace[shape.text]);
    }

    // 5. The content page: take the place-marker away and lay the blocks into its room.
    if (kind === "content") {
      for (const id of tpl.areaShapes ?? []) commands.push({ op: "delete", target: target(id) });
      if (tpl.area) commands.push(...contentCommands(page, tpl.area, at, paletteOf(template, tpl.dark), path));
      else if (page.blocks) fail("模板的内容页没有标出放内容的位置", `${path}.blocks`);
    } else if (page.blocks) {
      fail(`只有 content 页可以放 blocks，"${kind}" 页不行`, `${path}.blocks`);
    }
  });

  // 6. The template's own pages.
  for (let i = 0; i < last; i++) commands.push({ op: "deleteSlide", slide: 1 });
  return { label: `按模板生成 ${deck.slides.length} 页`, atomic: true, commands };
}

/** For a model that answers with the whole deck as one JSON object instead of calling tools. */
export const SPEC_GUIDE = [
  "你的任务是按模板写出一份演示文稿的内容。只输出一个 JSON 对象，不要任何解释，不要代码围栏。",
  '格式：{"slides":[{"page":"cover","title":"…"}, …], "replace":{"模板里的示例文字":"换成的文字"}}',
  "每页 page 取 cover（封面）、toc（目录）、section（章节页）、content（内容页）、end（结尾）之一；其余字段见下面的模板说明里每种页面的 fields。",
  'toc 页写 items: ["章节名" 或 {"title":"…","desc":"一句话说明"}]；不要写编号，不要超过模板给的 entries 数。',
  "content 页写 title 和 blocks（1 到 3 块，一页放不下就拆成两页）。块有五种：",
  '  {"type":"text","text":"一段话"}',
  '  {"type":"bullets","items":["要点","要点"]}（每条一句话，3 到 5 条）',
  '  {"type":"cards","items":[{"value":"1290万","title":"年销量","text":"同比增长 35%"}]}（2 到 4 张；value 是突出的数字）',
  '  {"type":"chart","kind":"column|bar|line|pie","categories":["…"],"series":[{"name":"…","values":[1,2]}],"title":"…"}（values 个数和 categories 一样多）',
  '  {"type":"table","rows":[["表头","表头"],["…","…"]]}（第一行是表头）',
  '两个块默认左右并排，三个块上下排列；也可以写 "layout":"rows" 或 "columns"。',
  "不要写坐标和字号，排版由引擎完成。模板里带“示例”性质的文字（品牌名、口号、机构名、保密提示等）如果和主题无关，用 replace 换成合适的文字，写空字符串就删掉。",
  "要点和说明要具体、有数据，不要空话；每个文字块一般不超过 40 个汉字。"
].join("\n");
