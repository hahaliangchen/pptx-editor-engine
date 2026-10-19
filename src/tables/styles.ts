/** What one part of a table style (the header row, the banded rows ...) says. */
export interface StyleBorder {
  color: string;
  /** EMU: a style belongs to no slide, so it has no pixel size yet. */
  width: number;
}

export interface TableStylePart {
  textColor?: string;
  bold?: boolean;
  italic?: boolean;
  /** A colour, or null for "no fill". */
  fill?: string | null;
  borders?: Partial<Record<"left" | "right" | "top" | "bottom" | "insideH" | "insideV", StyleBorder | null>>;
}

export type TableStylePartName =
  | "wholeTbl" | "band1H" | "band2H" | "band1V" | "band2V" | "firstCol" | "lastCol"
  | "firstRow" | "lastRow" | "nwCell" | "neCell" | "swCell" | "seCell";

export type TableStyleDef = Partial<Record<TableStylePartName, TableStylePart>>;

export interface TableFlags {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  bandRow: boolean;
  bandColumn: boolean;
}

export interface ResolvedCellStyle {
  fill: string | null;
  textColor?: string;
  bold?: boolean;
  italic?: boolean;
  borders: { left: StyleBorder | null; right: StyleBorder | null; top: StyleBorder | null; bottom: StyleBorder | null };
}

/** The style PowerPoint applies to a new table; files that use it often leave its definition out. */
export const DEFAULT_TABLE_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

/** Medium Style 2 - Accent 1, written out for a deck whose tableStyles.xml does not carry it. */
export const builtinDefaultStyle = (accent: string, dark: string, light: string, tint: (c: string, f: number) => string): TableStyleDef => {
  const line = (width: number): StyleBorder => ({ color: light, width: width * 12700 });
  const all = { left: line(1), right: line(1), top: line(1), bottom: line(1), insideH: line(1), insideV: line(1) };
  return {
    wholeTbl: { textColor: dark, fill: tint(accent, 0.2), borders: all },
    band1H: { fill: tint(accent, 0.4) },
    band1V: { fill: tint(accent, 0.4) },
    lastCol: { textColor: light, bold: true, fill: accent },
    firstCol: { textColor: light, bold: true, fill: accent },
    lastRow: { textColor: light, bold: true, fill: accent, borders: { top: line(3) } },
    firstRow: { textColor: light, bold: true, fill: accent, borders: { bottom: line(3) } }
  };
};

/** Which parts of a style apply to the cell at (`row`, `col`), lowest priority first (the order they are added in). */
function partsFor(flags: TableFlags, row: number, col: number, rows: number, cols: number): TableStylePartName[] {
  const parts: TableStylePartName[] = ["wholeTbl"];
  const dataRow = row - (flags.firstRow ? 1 : 0);
  const dataCol = col - (flags.firstColumn ? 1 : 0);
  const isHeader = flags.firstRow && row === 0;
  const isFooter = flags.lastRow && row === rows - 1;
  if (flags.bandColumn && !(flags.firstColumn && col === 0) && !(flags.lastColumn && col === cols - 1) && dataCol >= 0) {
    parts.push(dataCol % 2 === 0 ? "band1V" : "band2V");
  }
  if (flags.bandRow && !isHeader && !isFooter && dataRow >= 0) parts.push(dataRow % 2 === 0 ? "band1H" : "band2H");
  if (flags.firstColumn && col === 0) parts.push("firstCol");
  if (flags.lastColumn && col === cols - 1) parts.push("lastCol");
  if (isFooter) parts.push("lastRow");
  if (isHeader) parts.push("firstRow");
  if (isHeader && flags.firstColumn && col === 0) parts.push("nwCell");
  if (isHeader && flags.lastColumn && col === cols - 1) parts.push("neCell");
  if (isFooter && flags.firstColumn && col === 0) parts.push("swCell");
  if (isFooter && flags.lastColumn && col === cols - 1) parts.push("seCell");
  return parts;
}

/** The look the style gives one cell: later parts override earlier ones. */
export function resolveCellStyle(
  style: TableStyleDef | undefined,
  flags: TableFlags,
  cell: { row: number; col: number; rowSpan: number; colSpan: number },
  rows: number,
  cols: number
): ResolvedCellStyle {
  const out: ResolvedCellStyle = { fill: null, borders: { left: null, right: null, top: null, bottom: null } };
  if (!style) return out;
  const lastRow = cell.row + cell.rowSpan >= rows;
  const lastCol = cell.col + cell.colSpan >= cols;
  for (const name of partsFor(flags, cell.row, cell.col, rows, cols)) {
    const part = style[name];
    if (!part) continue;
    if (part.fill !== undefined) out.fill = part.fill;
    if (part.textColor !== undefined) out.textColor = part.textColor;
    if (part.bold !== undefined) out.bold = part.bold;
    if (part.italic !== undefined) out.italic = part.italic;
    const b = part.borders;
    if (!b) continue;
    const whole = name === "wholeTbl";
    const side = (own: keyof NonNullable<typeof b>, inside: keyof NonNullable<typeof b>, outer: boolean): StyleBorder | null | undefined =>
      whole ? (outer ? b[own] : b[inside]) : b[own];
    const left = side("left", "insideV", cell.col === 0);
    const right = side("right", "insideV", lastCol);
    const top = side("top", "insideH", cell.row === 0);
    const bottom = side("bottom", "insideH", lastRow);
    if (left !== undefined) out.borders.left = left;
    if (right !== undefined) out.borders.right = right;
    if (top !== undefined) out.borders.top = top;
    if (bottom !== undefined) out.borders.bottom = bottom;
  }
  return out;
}
