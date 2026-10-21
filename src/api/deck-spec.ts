import { ApiError } from "./errors";
import { CM_TO_PT, INSET_CM, widthEm } from "./text-metrics";
import type { CmRect } from "./outline";
import { AUTO_TOKENS, PAGE_KINDS, PageKind, TemplateInfo, TemplatePage, TextSlot, fieldAlias, slotName, tokensIn } from "./template";
import { ImageSlot, MIN_UNIT_GAP, UnitGroup, isAutoUnitField } from "./units";

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
const RESERVED = new Set(["page", "variant", "items", "blocks", "layout", "replace"]);

const fail = (message: string, path: string): never => {
  throw new ApiError("INVALID_SCHEMA", message, path);
};

// ---------------------------------------------------------------- text fitting

export { CM_TO_PT, INSET_CM, widthEm };

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
  // Letter spacing and glyph widths make real text wider than the estimate: keep a 12% margin.
  const need = (size: number) => (widest * size * 1.12) / CM_TO_PT + 2 * INSET_CM;
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

interface Palette { text: string; muted: string; accent: string; card: string; tint: string; onAccent: string }

const brightness = (hex: string) => (0.299 * parseInt(hex.slice(1, 3), 16) + 0.587 * parseInt(hex.slice(3, 5), 16) + 0.114 * parseInt(hex.slice(5, 7), 16));

const paletteOf = (template: TemplateInfo, dark: boolean): Palette => ({
  text: dark ? "#ffffff" : "#333333",
  muted: dark ? "#d0d7e2" : "#555555",
  accent: template.accent,
  card: dark ? "#ffffff" : "#f5f7fa",
  // A very light wash of the accent: table bands.
  tint: dark ? "#ffffff" : mix(template.accent, 255, 0.93),
  // Text on a block of the accent colour.
  onAccent: brightness(template.accent) > 160 ? "#222222" : "#ffffff"
});

type Command = Record<string, unknown>;

const mix = (hex: string, other: number, amount: number) =>
  "#" + [1, 3, 5].map(i => Math.round(parseInt(hex.slice(i, i + 2), 16) * (1 - amount) + other * amount).toString(16).padStart(2, "0")).join("");

/** Series colours: the template's accent, then darker and lighter versions of it, then greys. */
const seriesColors = (accent: string): string[] => [accent, mix(accent, 0, 0.4), mix(accent, 255, 0.45), mix(accent, 0, 0.7), mix(accent, 255, 0.7), "#8c9aa8"];

const round = (value: number) => Math.round(value * 100) / 100;
const rectOf = (rect: CmRect) => ({ x: round(rect.x), y: round(rect.y), w: round(rect.w), h: round(rect.h) });

/**
 * The type scale, in points. A block takes the largest size in its range that
 * fits its room, and what sits beside a table or cards never gets much bigger
 * than their text, so the page keeps one reading order: heading, then body.
 */
export const TYPE_SCALE = {
  text: { max: 24, min: 16 },
  bullets: { max: 22, min: 15 },
  cardValue: { max: 44, min: 26 },
  cardTitle: { max: 22, min: 15 },
  cardText: { max: 17, min: 12 },
  table: { max: 20, min: 12 }
} as const;

const LINE = 1.25;
const BLOCK_GAP = 0.6;
const ptToCm = (pt: number) => pt / CM_TO_PT;
const linesHeight = (lines: number, size: number) => ptToCm(lines * size * LINE);

/** The largest size in [min, max] at which `text` wraps to at most `maxLines` lines in `width` cm. */
function fitLines(text: string, width: number, max: number, min: number, maxLines: number): number {
  for (let size = max; size > min; size -= 0.5) if (linesOf(text, size, width) <= maxLines) return size;
  return min;
}

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

// -- cards

interface CardPart { text: string; size: number; bold: boolean; color: string; height: number }

/**
 * What a card holds, stacked and centred in its height: value, title, text, each
 * as large as its role allows. All the sizes shrink together until the stack fits.
 */
function cardStack(card: { title?: string; value?: string; text?: string }, width: number, height: number, palette: Palette): CardPart[] {
  const roles: Array<{ text?: string; scale: { max: number; min: number }; lines: number; bold: boolean; color: string }> = [
    { text: card.value, scale: TYPE_SCALE.cardValue, lines: 1, bold: true, color: palette.accent },
    { text: card.title, scale: TYPE_SCALE.cardTitle, lines: 2, bold: true, color: palette.text },
    { text: card.text, scale: TYPE_SCALE.cardText, lines: 5, bold: false, color: palette.muted }
  ];
  const inner = width - 0.5;
  let parts: CardPart[] = [];
  for (let factor = 1; factor >= 0.55; factor -= 0.05) {
    parts = roles.filter(role => role.text).map(role => {
      const max = Math.max(role.scale.min, role.scale.max * factor);
      const size = fitLines(role.text as string, inner, max, role.scale.min, role.lines);
      return { text: role.text as string, size, bold: role.bold, color: role.color, height: linesHeight(linesOf(role.text as string, size, inner), size) + 0.15 };
    });
    const total = parts.reduce((sum, part) => sum + part.height, 0) + 0.25 * (parts.length - 1);
    if (total <= height - 0.8) break;
  }
  return parts;
}

const stackHeight = (parts: CardPart[]) => parts.reduce((sum, part) => sum + part.height, 0) + 0.25 * Math.max(0, parts.length - 1);

function cardGrid(count: number, rect: CmRect): { cells: CmRect[]; cardHeight: number } {
  // A card is at most 11 cm tall; with other blocks beside it the room it is given is smaller anyway.
  const columns = count <= 4 ? count : Math.ceil(count / 2);
  const rows = Math.ceil(count / columns);
  const gap = 0.5;
  const cells = split(rect, Array(rows).fill(1), "rows", gap).flatMap(row => split(row, Array(columns).fill(1), "columns", gap));
  return { cells, cardHeight: Math.min(cells[0].h, 11) };
}

// -- tables

/** Column widths follow what the columns hold; none gets less than half the equal share. */
function columnWidths(rows: string[][], width: number): number[] {
  const columns = Math.max(...rows.map(row => row.length));
  const widest = Array.from({ length: columns }, (_, c) => Math.max(2, ...rows.map(row => widthEm(row[c] ?? ""))));
  const weights = widest.map(value => Math.sqrt(value));
  const floor = (1 / columns) * 0.5;
  const total = weights.reduce((sum, value) => sum + value, 0);
  const shares = weights.map(value => Math.max(floor, value / total));
  const sum = shares.reduce((a, b) => a + b, 0);
  return shares.map(share => round((share / sum) * width));
}

/** The largest size at which every row of the table fits its row height. */
const cellLines = (text: string, size: number, width: number) => linesOf(text, size * 1.12, width - 0.3);

function tableSize(rows: string[][], widths: number[], rowHeight: number): number {
  for (let size = TYPE_SCALE.table.max; size > TYPE_SCALE.table.min; size -= 0.5) {
    const fits = rows.every(row => {
      const lines = Math.max(...widths.map((width, c) => cellLines(row[c] ?? "", size, width)));
      return linesHeight(lines, size) + 0.5 <= rowHeight;
    });
    if (fits) return size;
  }
  return TYPE_SCALE.table.min;
}

interface TablePlan { rows: string[][]; rect: CmRect; widths: number[]; size: number }

function tablePlan(block: Extract<Block, { type: "table" }>, rect: CmRect): TablePlan {
  const columns = Math.max(...block.rows.map(row => row.length));
  const rows = block.rows.map(row => [...row, ...Array(columns - row.length).fill("")]);
  const widths = columnWidths(rows, rect.w);
  // Rows are as tall as their text asks (at most 2.4 cm), and the table sits in the middle of its room.
  const roomRow = Math.min(rect.h / rows.length, 2.8);
  const size = tableSize(rows, widths, roomRow);
  const need = Math.max(...rows.map(row => linesHeight(Math.max(...widths.map((width, c) => cellLines(row[c], size, width))), size))) + 0.7;
  const rowHeight = Math.min(roomRow, Math.max(need, 2.4));
  const height = Math.min(rect.h, rowHeight * rows.length);
  return { rows, rect: { x: rect.x, y: rect.y + (rect.h - height) / 2, w: rect.w, h: height }, widths, size };
}

// -- blocks

interface Room { min: number; need: number; max: number }

/** What a block needs in height, in cm, when it is `width` cm wide: its least, its comfortable size and its most. */
function roomFor(block: Block, width: number, palette: Palette, alone = false): Room {
  switch (block.type) {
    case "cards": {
      const count = block.items.length;
      const columns = count <= 4 ? count : Math.ceil(count / 2);
      const rows = Math.ceil(count / columns);
      const cardWidth = (width - 0.5 * (columns - 1)) / columns;
      const tallest = Math.max(...block.items.map(card => stackHeight(cardStack(card, cardWidth, 6.5, palette))));
      const one = Math.min(9.5, Math.max(4.5, tallest + 1.4));
      return { min: 3.8 * rows, need: one * rows + 0.5 * (rows - 1), max: (alone ? 11 : 9.5) * rows + 0.5 * (rows - 1) };
    }
    case "chart": return { min: 5.5, need: 8, max: 13 };
    case "table": return { min: Math.max(3.5, block.rows.length * 1.1), need: Math.max(4, block.rows.length * 1.5 + 0.3), max: block.rows.length * 2.8 };
    case "text": {
      const need = linesHeight(linesOf(block.text, 20, width), 20) + 0.8;
      return { min: Math.min(need, 2.2), need: Math.max(2.2, need), max: Math.max(2.2, need) + 2 };
    }
    case "bullets": {
      const lines = block.items.reduce((sum, item) => sum + linesOf(`• ${item}`, 20, width), 0);
      const need = linesHeight(lines, 20) + block.items.length * 0.35 + 0.8;
      return { min: Math.min(need, 3), need: Math.max(3, need), max: Math.max(3, need) + 2.5 };
    }
  }
}

/** The smallest text a block sets in a page: what text and bullets beside it are measured against. */
function referenceSize(block: Block, rect: CmRect, palette: Palette): number | undefined {
  if (block.type === "table") return tablePlan(block, rect).size;
  if (block.type === "cards") {
    const { cells, cardHeight } = cardGrid(block.items.length, rect);
    const sizes = block.items.map((card, i) => cardStack(card, cells[i].w, cardHeight, palette)).flatMap(parts => parts.filter(part => !part.bold).map(part => part.size));
    return sizes.length ? Math.min(...sizes) : undefined;
  }
  return undefined;
}

/** Up to six points on a page of their own: numbered rows that use the room, not a short list in the middle. */
const LIST_LIMIT = 6;

function listCommands(items: string[], rect: CmRect, slide: number, palette: Palette): Command[] {
  const gap = 0.4;
  const rowHeight = Math.min(2.8, Math.max(1.5, (rect.h * 0.9 - gap * (items.length - 1)) / items.length));
  const total = rowHeight * items.length + gap * (items.length - 1);
  const badge = Math.min(1.5, rowHeight * 0.62);
  const textWidth = rect.w - badge - 1.6;
  const size = Math.min(...items.map(item => fitFontSize([item], { w: textWidth, h: rowHeight }, 26, 16)));
  const commands: Command[] = [];
  let y = rect.y + Math.max(0, rect.h - total) * 0.42;
  items.forEach((item, index) => {
    commands.push({ op: "addShape", slide, shape: "roundRect", rect: rectOf({ x: rect.x, y, w: rect.w, h: rowHeight }), fill: palette.tint, line: null });
    commands.push({
      op: "addShape", slide, shape: "ellipse", rect: rectOf({ x: rect.x + 0.6, y: y + (rowHeight - badge) / 2, w: badge, h: badge }),
      fill: palette.accent, line: null, text: String(index + 1)
    });
    commands.push({
      op: "addText", slide, rect: rectOf({ x: rect.x + badge + 1.2, y, w: textWidth, h: rowHeight }), text: item,
      style: { fontSizePt: size, color: palette.text }, as: `row${slide}_${index}`
    });
    commands.push({ op: "setAnchor", target: `$row${slide}_${index}`, anchor: "middle" });
    y += rowHeight + gap;
  });
  return commands;
}

function blockCommands(block: Block, rect: CmRect, slide: number, palette: Palette, id: string, cap?: number, alone = false): Command[] {
  switch (block.type) {
    case "text": {
      const max = cap ? Math.min(TYPE_SCALE.text.max, Math.max(TYPE_SCALE.text.min, cap + 4)) : TYPE_SCALE.text.max;
      const size = fitFontSize([block.text], rect, max, TYPE_SCALE.text.min);
      return [
        { op: "addText", slide, rect: rectOf(rect), text: block.text, style: { fontSizePt: size, color: palette.text }, as: id },
        { op: "setAnchor", target: `$${id}`, anchor: "middle" }
      ];
    }
    case "bullets": {
      if (alone && block.items.length <= LIST_LIMIT) return listCommands(block.items, rect, slide, palette);
      const items = block.items.map(item => `• ${item}`);
      const max = cap ? Math.min(TYPE_SCALE.bullets.max, Math.max(TYPE_SCALE.bullets.min, cap + 4)) : TYPE_SCALE.bullets.max;
      const size = fitFontSize(items, rect, max, TYPE_SCALE.bullets.min, 1.3, 8);
      return [
        { op: "addText", slide, rect: rectOf(rect), paragraphs: items.map(text => ({ text })), style: { fontSizePt: size, color: palette.text }, as: id },
        { op: "setAnchor", target: `$${id}`, anchor: "middle" }
      ];
    }
    case "cards": {
      const { cells, cardHeight } = cardGrid(block.items.length, rect);
      const commands: Command[] = [];
      block.items.forEach((card, index) => {
        const cell = cells[index];
        const body = { ...cell, h: cardHeight, y: cell.y + (cell.h - cardHeight) / 2 };
        commands.push({
          op: "addShape", slide, shape: "roundRect", rect: rectOf(body),
          fill: palette.card, opacity: palette.card === "#ffffff" ? 0.12 : 1,
          line: { color: palette.accent, widthPt: 1 }
        });
        const parts = cardStack(card, body.w, body.h, palette);
        let y = body.y + (body.h - stackHeight(parts)) / 2;
        for (const part of parts) {
          commands.push({
            op: "addText", slide, rect: rectOf({ x: body.x + 0.25, y, w: body.w - 0.5, h: part.height }), text: part.text,
            style: { fontSizePt: part.size, bold: part.bold, color: part.color, align: "center" }
          });
          y += part.height + 0.25;
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
        fontSizePt: 14, titleSizePt: 18,
        textColor: palette.text, rect: rectOf(rect)
      }];
    case "table": {
      const plan = tablePlan(block, rect);
      const commands: Command[] = [{
        op: "addTable", slide, rows: plan.rows, fontSizePt: plan.size, columnWidths: plan.widths, rect: rectOf(plan.rect), as: id
      }];
      // Header in the accent colour, body rows banded with a wash of it, short figures centred.
      plan.rows.forEach((row, r) => row.forEach((cell, c) => {
        const target = `$${id}`;
        if (r === 0) commands.push({ op: "setCell", target, row: 1, col: c + 1, fill: palette.accent, color: palette.onAccent, bold: true, align: "center" });
        else {
          const figure = /^[\d.,%+\-¥$￥万亿元 ]{1,10}$/.test(cell.trim());
          commands.push({ op: "setCell", target, row: r + 1, col: c + 1, fill: r % 2 === 0 ? palette.tint : "#ffffff", color: palette.text, ...(figure ? { align: "center" } : { align: "left" }) });
        }
      }));
      return commands;
    }
  }
}

function contentCommands(page: PageSpec, area: CmRect, slide: number, palette: Palette, path: string): Command[] {
  const blocks = (page.blocks ?? []).map((block, i) => validateBlock(block, `${path}.blocks[${i}]`));
  if (blocks.length === 0) return [];
  if (blocks.length > 3) fail("一页最多 3 个内容块；内容多就拆成两页", `${path}.blocks`);
  const direction = page.layout ?? (blocks.length === 2 && !blocks.some(block => block.type === "cards") ? "columns" : "rows");
  const gap = BLOCK_GAP;
  const weights = blocks.map(block => (direction === "columns" && (block.type === "text" || block.type === "bullets") ? 1 : direction === "columns" ? 1.35 : 1));
  const widths = split(area, weights, "columns", gap).map(part => part.w);
  const rooms = blocks.map((block, i) => roomFor(block, direction === "columns" ? widths[i] : area.w, palette, blocks.length === 1));

  let rects: CmRect[];
  if (direction === "columns") {
    // Side by side: each block is as tall as it can usefully be and centred in the height; text fills it and centres itself.
    rects = split(area, weights, "columns", gap).map((part, i) => {
      const block = blocks[i];
      if (block.type === "text" || block.type === "bullets") return part;
      const height = Math.min(part.h, rooms[i].max);
      return { ...part, y: part.y + (part.h - height) / 2, h: height };
    });
  } else {
    // Stacked: every block gets its comfortable height; spare room goes to the blocks that can grow, and
    // what is still left is split around the stack (a little more below) so the page is balanced.
    const total = (heights: number[]) => heights.reduce((sum, h) => sum + h, 0) + gap * (blocks.length - 1);
    let heights = rooms.map(room => room.need);
    if (total(heights) > area.h) {
      const scale = (area.h - gap * (blocks.length - 1)) / heights.reduce((sum, h) => sum + h, 0);
      heights = heights.map((h, i) => Math.max(Math.min(rooms[i].min, h), h * scale));
    } else {
      const growth = rooms.map(room => room.max - room.need);
      const room = area.h - total(heights);
      const growable = growth.reduce((sum, g) => sum + g, 0);
      const used = Math.min(room, growable);
      if (growable > 0) heights = heights.map((h, i) => h + (used * growth[i]) / growable);
    }
    const left = Math.max(0, area.h - total(heights));
    let y = area.y + left * 0.42;
    rects = heights.map(h => {
      const rect = { x: area.x, y, w: area.w, h };
      y += h + gap;
      return rect;
    });
  }

  // What sits beside a table or cards is measured against their text.
  const references = blocks.map((block, i) => referenceSize(block, rects[i], palette)).filter((size): size is number => size !== undefined);
  const cap = references.length ? Math.min(...references) : undefined;
  const alone = blocks.length === 1 && page.layout === undefined;
  return blocks.flatMap((block, i) => blockCommands(block, alone && block.type === "bullets" && block.items.length <= LIST_LIMIT ? area : rects[i], slide, palette, `b${slide}_${i}`, cap, alone));
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
  // Pages are worked on as copies: the usual fields may be written in the words a template uses.
  const slides = deck.slides.map(page => (page && typeof page === "object" ? { ...page } : page));
  const commands: Command[] = [];
  const last = template.slideCount;
  let total = last;
  let sections = 0;
  let contents = 0;

  slides.forEach((page, index) => {
    const path = `slides[${index}]`;
    if (!page || typeof page !== "object") fail("每页都是一个对象", path);
    let source = template.pages[page.page];
    if (!source) {
      const available = PAGE_KINDS.filter(kind => template.pages[kind]);
      fail(`模板没有 "${String(page.page)}" 这种页面，可用的有：${available.join("、")}`, `${path}.page`);
    }
    if (page.variant !== undefined) {
      const wanted = String(page.variant).trim().toLowerCase();
      const same = template.variants.filter(item => item.kind === page.page);
      source = same.find(item => item.variant?.toLowerCase() === wanted);
      if (!source) fail(`"${String(page.page)}" 页没有叫 "${String(page.variant)}" 的版式，可用的有：${same.map(item => item.variant ?? "（默认）").join("、")}`, `${path}.variant`);
    }
    const kind = page.page;
    const tpl = source as TemplatePage;
    const number = deck.slides.length;

    for (const key of Object.keys(page)) {
      const canonical = !RESERVED.has(key) && !(key in tpl.slots) && !(key in tpl.units) ? fieldAlias(kind, key) : undefined;
      if (canonical && canonical in tpl.slots && !(canonical in page)) {
        page[canonical] = page[key];
        delete page[key];
      }
    }
    for (const field of Object.keys(page)) {
      if (RESERVED.has(field) || field in tpl.slots || field in tpl.units || field in tpl.images) continue;
      fail(`"${kind}" 页没有 "${field}" 这个字段，它有：${[...Object.keys(tpl.slots), ...Object.keys(tpl.images), ...Object.keys(tpl.units)].join("、") || "（无）"}`, `${path}.${field}`);
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
    if (kind === "content") contents += 1;
    const auto: Record<string, string> = {
      PAGE_NUM: String(index + 1),
      TOTAL_PAGES: String(number),
      // Chapters are the section pages; a deck without any numbers its content pages instead of showing "00".
      CHAPTER_NUM: String(typeof page.number === "string" || typeof page.number === "number" ? page.number : String(sections || contents).padStart(2, "0")),
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
        // A number keeps the look of the sample it replaces: "1" stays unpadded, "01" two digits.
        const value = AUTO_TOKENS.has(token) ? likeSample(auto[token], shape.sample) : values[tpl.fieldNames[token] ?? slotName(token)];
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

    // 6. Repeating units: as many instances as the content has items.
    Object.values(tpl.units).forEach((group, groupNo) => {
      if (page[group.name] === undefined) {
        fail(`这个版式需要“${group.name}”：写成数组，每项是 {${group.fields.join("、")}}（${group.fixed ? `固定 ${group.instances.length} 个` : `${group.min} 到 ${group.max} 个`}）`, `${path}.${group.name}`);
      }
      commands.push(...unitCommands(group, page[group.name], at, groupNo, template.slideSize.w, `${path}.${group.name}`));
    });

    // 7. Picture slots of the page: a deck's own picture (or a stand-in) takes the sample's place.
    Object.entries(tpl.images).forEach(([field, slot], n) => {
      if (page[field] === undefined) fail(`这个版式需要图片“${field}”：写图片的路径（project:images/…），或写 placeholder 先放一张占位图`, `${path}.${field}`);
      commands.push({ op: "delete", target: { slide: at, id: slot.id } });
      commands.push(...imageCommands(slot, textOf(page[field], `${path}.${field}`), slot.rect, at, `img${at}_${n}`, `${path}.${field}`));
    });
  });

  // 6. The template's own pages.
  for (let i = 0; i < last; i++) commands.push({ op: "deleteSlide", slide: 1 });
  return { label: `按模板生成 ${deck.slides.length} 页`, atomic: true, commands };
}

/** What a picture field may say: a project picture, a data URL, or "placeholder" for a stand-in. */
function imageSource(value: string, path: string): string {
  const text = value.trim();
  if (/^(placeholder:?|占位图?)$/i.test(text)) return "placeholder:";
  if (/^data:image\/(png|jpeg|gif);base64,/.test(text) || text.startsWith("project:images/")) return text;
  if (text.startsWith("images/")) return `project:${text}`;
  return fail("图片要写成 project:images/… 的项目图片路径（用户上传的图片），或 placeholder", path);
}

/** A picture in a slot's place and at its depth, cut to the slot's proportions and outline. */
function imageCommands(slot: ImageSlot, value: string, rect: CmRect, slide: number, alias: string, path: string): Command[] {
  return [
    { op: "addImage", slide, src: imageSource(value, path), rect: rectOf(rect), fit: { shape: slot.shape ?? "rect" }, as: alias },
    // The new picture lands on top; take it back down to where the sample was.
    ...Array.from({ length: Math.min(slot.above ?? 0, 60) }, () => ({ op: "arrange", target: `$${alias}`, mode: "backward" }))
  ];
}

/** The commands that make `raw` (one object per item) out of a repeating unit: clone, move or remove instances, fill their text. */
function unitCommands(group: UnitGroup, raw: unknown, slide: number, groupNo: number, pageWidth: number, path: string): Command[] {
  if (!Array.isArray(raw) || raw.length === 0) fail(`“${group.name}”要写成非空数组，每项是 {${group.fields.join("、")}}`, path);
  const items = (raw as unknown[]).map((item, i) => {
    if (!item || typeof item !== "object" || Array.isArray(item)) fail(`每一项都是一个对象，字段有：${group.fields.join("、")}`, `${path}[${i}]`);
    const object = item as Record<string, unknown>;
    for (const key of Object.keys(object)) {
      if (!group.fields.includes(key)) fail(`“${group.name}”的项没有 "${key}" 这个字段，它有：${group.fields.join("、")}`, `${path}[${i}].${key}`);
    }
    return Object.fromEntries(Object.entries(object).map(([key, value]) => [key, textOf(value, `${path}[${i}].${key}`)]));
  });
  const k = items.length;
  const n = group.instances.length;
  if (k < group.min || k > group.max) {
    fail(group.fixed
      ? `“${group.name}”在这个版式里固定是 ${n} 个，现在给了 ${k} 个；给 ${n} 个，或换一个版式`
      : `“${group.name}”可以给 ${group.min} 到 ${group.max} 个，现在给了 ${k} 个${k > group.max ? "；内容多就拆成两页" : ""}`, path);
  }

  const base = group.instances[0];
  const pitch = group.pitch;
  // Where each instance's top-left goes, and how much the instance shrinks to make room.
  let scale = 1;
  let origins: Array<{ x: number; y: number }>;
  if (!pitch || k === n) {
    origins = group.instances.slice(0, k).map(instance => ({ x: instance.rect.x, y: instance.rect.y }));
  } else if (k < n) {
    // Fewer than the template has: the same spacing, the block centred where the template's row was.
    origins = Array.from({ length: k }, (_, i) => ({ x: base.rect.x + pitch.dx * (i + (n - k) / 2), y: base.rect.y + pitch.dy * (i + (n - k) / 2) }));
  } else {
    const horizontal = Math.abs(pitch.dx) >= Math.abs(pitch.dy);
    const last = group.instances[n - 1].rect;
    const span = horizontal ? last.x + last.w - base.rect.x : last.y + last.h - base.rect.y;
    const size = horizontal ? base.rect.w : base.rect.h;
    const step = Math.abs(horizontal ? pitch.dx : pitch.dy);
    // More than the template has: first the gaps close up (down to a small gap), then the instances shrink.
    let newStep: number;
    if (k * size + (k - 1) * MIN_UNIT_GAP <= span) {
      newStep = (span - size) / (k - 1);
    } else {
      scale = (span - (k - 1) * MIN_UNIT_GAP) / (k * size);
      newStep = size * scale + MIN_UNIT_GAP;
    }
    const along = newStep / step;
    origins = Array.from({ length: k }, (_, i) => ({
      x: base.rect.x + pitch.dx * along * i + (horizontal ? 0 : (base.rect.w * (1 - scale)) / 2),
      y: base.rect.y + pitch.dy * along * i + (horizontal ? (base.rect.h * (1 - scale)) / 2 : 0)
    }));
  }

  const commands: Command[] = [];
  for (let i = 0; i < k; i++) {
    base.members.forEach((pattern, m) => {
      const original = group.instances[i]?.members[m];
      const alias = `u${slide}_${groupNo}_${i}_${m}`;
      const target = original ? { slide, id: original.id } : `$${alias}`;
      const field = pattern.field;
      const value = field ? (isAutoUnitField(field) ? likeSample(String(i + 1), pattern.text) : items[i][field]) : undefined;
      if (pattern.image) {
        if (value === undefined || value === "") fail(`“${group.name}”第 ${i + 1} 项缺少图片“${field}”：写图片的路径，或写 placeholder`, `${path}[${i}].${field}`);
        const at = {
          x: origins[i].x + (pattern.rect.x - base.rect.x) * scale, y: origins[i].y + (pattern.rect.y - base.rect.y) * scale,
          w: pattern.rect.w * scale, h: pattern.rect.h * scale
        };
        if (original) commands.push({ op: "delete", target });
        commands.push(...imageCommands(pattern.image, value as string, at, slide, `${alias}_img`, `${path}[${i}].${field}`));
        return;
      }
      // A field the item leaves out takes its shape away (an optional caption).
      if (field && (value === undefined || value === "")) {
        if (original) commands.push({ op: "delete", target });
        return;
      }
      if (!original) commands.push({ op: "duplicate", target: { slide, id: pattern.id }, as: alias });
      const rect = {
        x: origins[i].x + (pattern.rect.x - base.rect.x) * scale, y: origins[i].y + (pattern.rect.y - base.rect.y) * scale,
        w: pattern.rect.w * scale, h: pattern.rect.h * scale
      };
      const moved = !original || scale !== 1 || Math.abs(rect.x - pattern.rect.x) > 0.01 || Math.abs(rect.y - pattern.rect.y) > 0.01;
      if (moved) commands.push({ op: "setBounds", target, ...(scale !== 1 || !original ? rectOf(rect) : { x: round(rect.x), y: round(rect.y) }) });
      if (field && value !== undefined) {
        commands.push({ op: "setText", target, text: value });
        const design = pattern.fontSizePt ?? 18;
        const base0 = design * scale;
        const size = fitFontSize([value], { w: rect.w, h: rect.h }, base0, Math.max(8, base0 * 0.6));
        if (Math.abs(size - design) > 0.01) commands.push({ op: "formatText", target, fontSizePt: size });
      }
    });
  }
  // Instances the content does not use go, with all their shapes.
  for (let i = k; i < n; i++) for (const member of group.instances[i].members) commands.push({ op: "delete", target: { slide, id: member.id } });
  return commands;
}

const likeSample = (value: string, sample: string | undefined): string =>
  sample !== undefined && /^\d+$/.test(sample.trim()) && /^\d+$/.test(value) ? String(Number(value)).padStart(sample.trim().length, "0") : value;

/** For a model that answers with the whole deck as one JSON object instead of calling tools. */
export const SPEC_GUIDE = [
  "你的任务是按模板写出一份演示文稿的内容。只输出一个 JSON 对象，不要任何解释，不要代码围栏。",
  '格式：{"slides":[{"page":"cover","title":"…"}, …], "replace":{"模板里的示例文字":"换成的文字"}}',
  "每页 page 取 cover（封面）、toc（目录）、section（章节页）、content（内容页）、end（结尾）之一；其余字段见下面的模板说明里每种页面的 fields。",
  'toc 页写 items: ["章节名" 或 {"title":"…","desc":"一句话说明"}]；不要写编号。条数按实际章节数定，1 条到模板的 maxEntries 条都可以，不要凑满。',
  "content 页先按模板的 fields 填写；正文文本框直接填写对应字段，重复结构填写对应数组。只有模板标出正文区域（area）时才写 blocks（1 到 3 块，一页放不下就拆成两页）。块有五种：",
  '  {"type":"text","text":"一段话"}',
  '  {"type":"bullets","items":["要点","要点"]}（每条一句话，3 到 5 条）',
  '  {"type":"cards","items":[{"value":"1290万","title":"年销量","text":"同比增长 35%"}]}（2 到 4 张；value 是突出的数字）',
  '  {"type":"chart","kind":"column|bar|line|pie","categories":["…"],"series":[{"name":"…","values":[1,2]}],"title":"…"}（values 个数和 categories 一样多）',
  '  {"type":"table","rows":[["表头","表头"],["…","…"]]}（第一行是表头）',
  '两个块默认左右并排，三个块上下排列；也可以写 "layout":"rows" 或 "columns"。',
  '同一种页面有几个版式时，用 "variant":"版式名" 选一个（名字和适用场景见模板说明）；版式里的重复单元（groups）写成数组，例如 "项":[{"标题":"…","说明":"…"}, …]，条数在说明给的范围内，内容有几条就给几条。',
  "模板说明里每个字段的 fieldInfo 给出它在模板里的示例文字和大约能放多少字（fitChars），写得和示例长度相近。",
  "不要写坐标和字号，排版由引擎完成。模板里带“示例”性质的文字（品牌名、口号、机构名、保密提示等）如果和主题无关，用 replace 换成合适的文字，写空字符串就删掉。",
  "要点和说明要具体、有数据，不要空话；每个文字块一般不超过 40 个汉字。"
].join("\n");
