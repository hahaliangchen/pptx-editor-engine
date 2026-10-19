import { getDirectChild, getDirectChildren } from "../pptx-xml";
import { textWidth } from "../charts/layout";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

const PT = 12700;
const DEFAULT_PT = 18;

/** A table edit that cannot be made, with the reason a person can act on. */
export class TableEditError extends Error {}

export const escapeXml = (text: string): string =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

interface Parsed {
  doc: Document;
  frame: Element;
  tbl: Element;
}

function parse(xml: string): Parsed {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const frame = doc.documentElement;
  const tbl = Array.from(frame.getElementsByTagNameNS("*", "tbl"))[0];
  if (!tbl) throw new TableEditError("这不是一个表格");
  return { doc, frame, tbl };
}

const serialize = (doc: Document): string => new XMLSerializer().serializeToString(doc.documentElement);

const rowsOf = (tbl: Element) => getDirectChildren(tbl, "tr");
const cellsOf = (tr: Element) => getDirectChildren(tr, "tc");
const gridOf = (tbl: Element) => {
  const grid = getDirectChild(tbl, "tblGrid");
  return grid ? getDirectChildren(grid, "gridCol") : [];
};

const hasMerge = (tbl: Element): boolean =>
  rowsOf(tbl).some(tr => cellsOf(tr).some(tc =>
    parseInt(tc.getAttribute("gridSpan") || "1", 10) > 1 || parseInt(tc.getAttribute("rowSpan") || "1", 10) > 1
    || ["1", "true"].includes(tc.getAttribute("hMerge") || "") || ["1", "true"].includes(tc.getAttribute("vMerge") || "")));

function requireCell(tbl: Element, row: number, col: number): { tr: Element; tc: Element } {
  const rows = rowsOf(tbl);
  if (row < 0 || row >= rows.length) throw new TableEditError(`表格只有 ${rows.length} 行，没有第 ${row + 1} 行`);
  const cells = cellsOf(rows[row]);
  if (col < 0 || col >= cells.length) throw new TableEditError(`表格只有 ${cells.length} 列，没有第 ${col + 1} 列`);
  const tc = cells[col];
  if (["1", "true"].includes(tc.getAttribute("hMerge") || "") || ["1", "true"].includes(tc.getAttribute("vMerge") || "")) {
    throw new TableEditError(`第 ${row + 1} 行第 ${col + 1} 列被合并单元格盖住了，请改合并单元格本身`);
  }
  return { tr: rows[row], tc };
}

/** Keep the frame's size equal to what its grid adds up to. */
function syncFrame(frame: Element, tbl: Element): void {
  const ext = Array.from(frame.getElementsByTagNameNS("*", "ext")).find(node => node.parentNode && (node.parentNode as Element).localName === "xfrm");
  if (!ext) return;
  const width = gridOf(tbl).reduce((sum, col) => sum + (parseInt(col.getAttribute("w") || "0", 10) || 0), 0);
  const height = rowsOf(tbl).reduce((sum, row) => sum + (parseInt(row.getAttribute("h") || "0", 10) || 0), 0);
  if (width > 0) ext.setAttribute("cx", String(width));
  if (height > 0) ext.setAttribute("cy", String(height));
}

/** Identifiers PowerPoint keeps unique per row and column; a copy must not carry them over. */
function dropIds(node: Element): void {
  for (const list of getDirectChildren(node, "extLst")) node.removeChild(list);
}

const el = (doc: Document, name: string) => doc.createElementNS(A_NS, `a:${name}`);

// ------------------------------------------------------------------- text

/** The run properties that text added to a cell should wear: the cell's first run, else its end-of-paragraph mark. */
function runTemplate(doc: Document, txBody: Element): Element | null {
  const run = Array.from(txBody.getElementsByTagNameNS(A_NS, "r"))[0];
  const props = run && getDirectChild(run, "rPr");
  if (props) return props.cloneNode(true) as Element;
  const end = Array.from(txBody.getElementsByTagNameNS(A_NS, "endParaRPr"))[0];
  if (!end) return null;
  const copy = el(doc, "rPr");
  for (const attribute of Array.from(end.attributes)) copy.setAttribute(attribute.name, attribute.value);
  for (const child of Array.from(end.childNodes)) copy.appendChild(child.cloneNode(true));
  return copy;
}

function ensureTxBody(doc: Document, tc: Element): Element {
  let txBody = getDirectChild(tc, "txBody");
  if (!txBody) {
    txBody = doc.createElementNS(A_NS, "a:txBody");
    txBody.appendChild(el(doc, "bodyPr"));
    txBody.appendChild(el(doc, "lstStyle"));
    tc.insertBefore(txBody, tc.firstChild);
  }
  return txBody;
}

function writeText(doc: Document, tc: Element, text: string): void {
  const txBody = ensureTxBody(doc, tc);
  const template = runTemplate(doc, txBody);
  const firstParagraph = getDirectChild(txBody, "p");
  const pPr = firstParagraph && getDirectChild(firstParagraph, "pPr");
  for (const paragraph of getDirectChildren(txBody, "p")) txBody.removeChild(paragraph);
  for (const line of text.split("\n")) {
    const paragraph = el(doc, "p");
    if (pPr) paragraph.appendChild(pPr.cloneNode(true));
    if (line !== "") {
      const run = el(doc, "r");
      const props = template ? (template.cloneNode(true) as Element) : el(doc, "rPr");
      if (!props.getAttribute("lang")) props.setAttribute("lang", "zh-CN");
      run.appendChild(props);
      const content = el(doc, "t");
      content.textContent = line;
      run.appendChild(content);
      paragraph.appendChild(run);
    }
    const end = el(doc, "endParaRPr");
    if (template) {
      for (const attribute of Array.from(template.attributes)) end.setAttribute(attribute.name, attribute.value);
      for (const child of Array.from(template.childNodes)) end.appendChild(child.cloneNode(true));
    } else {
      end.setAttribute("lang", "zh-CN");
    }
    paragraph.appendChild(end);
    txBody.appendChild(paragraph);
  }
}

const sizeOf = (tc: Element): number => {
  const run = Array.from(tc.getElementsByTagNameNS(A_NS, "rPr"))[0] ?? Array.from(tc.getElementsByTagNameNS(A_NS, "endParaRPr"))[0];
  const size = parseInt(run?.getAttribute("sz") || "", 10);
  return Number.isFinite(size) ? size / 100 : DEFAULT_PT;
};

/** Make the row at least as tall as its text needs, as PowerPoint would when the text was typed. */
function fitRow(tbl: Element, row: number): void {
  const tr = rowsOf(tbl)[row];
  const columns = gridOf(tbl).map(col => (parseInt(col.getAttribute("w") || "0", 10) || 0) / PT);
  let needed = 0;
  cellsOf(tr).forEach((tc, c) => {
    if (["1", "true"].includes(tc.getAttribute("hMerge") || "") || ["1", "true"].includes(tc.getAttribute("vMerge") || "")) return;
    const span = parseInt(tc.getAttribute("gridSpan") || "1", 10) || 1;
    const props = getDirectChild(tc, "tcPr");
    const edge = (name: string, fallback: number) => (parseInt(props?.getAttribute(name) || String(fallback), 10) || 0) / PT;
    const width = Math.max(1, columns.slice(c, c + span).reduce((sum, w) => sum + w, 0) - edge("marL", 91440) - edge("marR", 91440));
    const size = sizeOf(tc);
    const txBody = getDirectChild(tc, "txBody");
    const lines = (txBody ? getDirectChildren(txBody, "p") : []).reduce((sum, paragraph) => {
      const text = Array.from(paragraph.getElementsByTagNameNS(A_NS, "t")).map(node => node.textContent || "").join("");
      return sum + Math.max(1, Math.ceil(textWidth(text, size) / width));
    }, 0);
    needed = Math.max(needed, lines * size * 1.2 + edge("marT", 45720) + edge("marB", 45720));
  });
  const current = parseInt(tr.getAttribute("h") || "0", 10) || 0;
  tr.setAttribute("h", String(Math.max(current, Math.round(needed * PT))));
}

export function setCellText(xml: string, row: number, col: number, text: string): string {
  const { doc, frame, tbl } = parse(xml);
  const { tc } = requireCell(tbl, row, col);
  writeText(doc, tc, text);
  fitRow(tbl, row);
  syncFrame(frame, tbl);
  return serialize(doc);
}

/** Rows, columns and whether any cell is merged. */
export function tableShape(xml: string): { rows: number; columns: number; merged: boolean } {
  const { tbl } = parse(xml);
  return { rows: rowsOf(tbl).length, columns: gridOf(tbl).length, merged: hasMerge(tbl) };
}

// ------------------------------------------------------------------ style

export interface CellStyleChange {
  fill?: string | null;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fontSizePt?: number;
  align?: "left" | "center" | "right";
}

const FILL_TAGS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
const hex = (color: string) => color.replace("#", "").toUpperCase();

function solid(doc: Document, color: string): Element {
  const fill = el(doc, "solidFill");
  const rgb = el(doc, "srgbClr");
  rgb.setAttribute("val", hex(color));
  fill.appendChild(rgb);
  return fill;
}

function styleRunProps(doc: Document, props: Element, change: CellStyleChange): void {
  if (change.bold !== undefined) props.setAttribute("b", change.bold ? "1" : "0");
  if (change.italic !== undefined) props.setAttribute("i", change.italic ? "1" : "0");
  if (change.fontSizePt !== undefined) props.setAttribute("sz", String(Math.round(change.fontSizePt * 100)));
  if (change.color) {
    for (const old of getDirectChildren(props, ...FILL_TAGS)) props.removeChild(old);
    const after = getDirectChild(props, "ln");
    props.insertBefore(solid(doc, change.color), after ? after.nextSibling : props.firstChild);
  }
}

export function setCellStyle(xml: string, row: number, col: number, change: CellStyleChange): string {
  const { doc, frame, tbl } = parse(xml);
  const { tc } = requireCell(tbl, row, col);

  if (change.fill !== undefined) {
    let props = getDirectChild(tc, "tcPr");
    if (!props) {
      props = el(doc, "tcPr");
      tc.appendChild(props);
    }
    for (const old of getDirectChildren(props, ...FILL_TAGS)) props.removeChild(old);
    const fill = change.fill === null ? el(doc, "noFill") : solid(doc, change.fill);
    // The schema puts the fill after the borders and before headers and extensions.
    props.insertBefore(fill, getDirectChildren(props, "headers", "extLst")[0] ?? null);
  }

  const txBody = ensureTxBody(doc, tc);
  const textual = change.bold !== undefined || change.italic !== undefined || change.fontSizePt !== undefined || change.color !== undefined;
  for (const paragraph of getDirectChildren(txBody, "p")) {
    if (change.align) {
      let pPr = getDirectChild(paragraph, "pPr");
      if (!pPr) {
        pPr = el(doc, "pPr");
        paragraph.insertBefore(pPr, paragraph.firstChild);
      }
      pPr.setAttribute("algn", { left: "l", center: "ctr", right: "r" }[change.align]);
    }
    if (!textual) continue;
    for (const run of getDirectChildren(paragraph, "r")) {
      let props = getDirectChild(run, "rPr");
      if (!props) {
        props = el(doc, "rPr");
        run.insertBefore(props, run.firstChild);
      }
      styleRunProps(doc, props, change);
    }
    let end = getDirectChild(paragraph, "endParaRPr");
    if (!end) {
      end = el(doc, "endParaRPr");
      paragraph.appendChild(end);
    }
    styleRunProps(doc, end, change);
  }
  if (change.fontSizePt !== undefined) fitRow(tbl, row);
  syncFrame(frame, tbl);
  return serialize(doc);
}

export function setTableFlags(xml: string, flags: { firstRow?: boolean; bandRow?: boolean; firstColumn?: boolean }): string {
  const { doc, tbl } = parse(xml);
  let props = getDirectChild(tbl, "tblPr");
  if (!props) {
    props = el(doc, "tblPr");
    tbl.insertBefore(props, tbl.firstChild);
  }
  const set = (name: string, value: boolean | undefined) => {
    if (value === undefined) return;
    if (value) props!.setAttribute(name, "1");
    else props!.removeAttribute(name);
  };
  set("firstRow", flags.firstRow);
  set("bandRow", flags.bandRow);
  set("firstCol", flags.firstColumn);
  return serialize(doc);
}

// ------------------------------------------------------------------ bands

const fillOf = (tc: Element): string | null => {
  const props = getDirectChild(tc, "tcPr");
  const solidFill = props && getDirectChild(props, "solidFill");
  const color = solidFill && getDirectChild(solidFill, "srgbClr");
  return color?.getAttribute("val")?.toUpperCase() ?? null;
};

/**
 * If the body rows alternate between exactly two plain colours (as a table made
 * here does, or any hand-banded one), those colours; otherwise null.
 */
function bandsOf(tbl: Element): { first: number; colors: [string, string] } | null {
  const props = getDirectChild(tbl, "tblPr");
  const first = ["1", "true"].includes(props?.getAttribute("firstRow") || "") ? 1 : 0;
  const body = rowsOf(tbl).slice(first);
  if (body.length < 2) return null;
  const fills: string[] = [];
  for (const tr of body) {
    const row = cellsOf(tr).map(fillOf);
    if (row.length === 0 || row[0] === null || row.some(fill => fill !== row[0])) return null;
    fills.push(row[0]!);
  }
  const distinct = [...new Set(fills)];
  if (distinct.length !== 2 || fills.some((fill, i) => i > 0 && fill === fills[i - 1])) return null;
  return { first, colors: [fills[0], fills[1]] };
}

function applyBands(tbl: Element, bands: { first: number; colors: [string, string] }): void {
  rowsOf(tbl).slice(bands.first).forEach((tr, i) => {
    for (const tc of cellsOf(tr)) {
      const color = tc.getElementsByTagNameNS(A_NS, "srgbClr");
      const fill = Array.from(color).find(node => (node.parentNode as Element).localName === "solidFill" && (node.parentNode!.parentNode as Element).localName === "tcPr");
      fill?.setAttribute("val", bands.colors[i % 2]);
    }
  });
}

// --------------------------------------------------------- rows and columns

function refuseMerged(tbl: Element, what: string): void {
  if (hasMerge(tbl)) throw new TableEditError(`这个表格有合并单元格，暂不支持${what}`);
}

function blankCell(doc: Document, source: Element, text: string): Element {
  const copy = source.cloneNode(true) as Element;
  for (const attribute of ["gridSpan", "rowSpan", "hMerge", "vMerge"]) copy.removeAttribute(attribute);
  dropIds(copy);
  writeText(doc, copy, text);
  return copy;
}

export function insertRow(xml: string, at: number, texts: string[] = []): string {
  const { doc, frame, tbl } = parse(xml);
  refuseMerged(tbl, "增加行");
  const rows = rowsOf(tbl);
  if (at < 0 || at > rows.length) throw new TableEditError(`只能在第 1 到 ${rows.length + 1} 行的位置插入，不能是 ${at + 1}`);
  const bands = bandsOf(tbl);
  const source = rows[Math.max(0, Math.min(at, rows.length) - (at >= rows.length ? 1 : 0))];
  const tr = source.cloneNode(true) as Element;
  dropIds(tr);
  cellsOf(tr).forEach((tc, c) => tr.replaceChild(blankCell(doc, tc, texts[c] ?? ""), tc));
  if (at >= rows.length) rows[rows.length - 1].after(tr);
  else rows[at].before(tr);
  if (bands) applyBands(tbl, bands);
  fitRow(tbl, rowsOf(tbl).indexOf(tr));
  syncFrame(frame, tbl);
  return serialize(doc);
}

export function deleteRow(xml: string, at: number): string {
  const { doc, frame, tbl } = parse(xml);
  refuseMerged(tbl, "删除行");
  const rows = rowsOf(tbl);
  if (rows.length <= 1) throw new TableEditError("表格至少要保留一行");
  if (at < 0 || at >= rows.length) throw new TableEditError(`表格只有 ${rows.length} 行，没有第 ${at + 1} 行`);
  const bands = bandsOf(tbl);
  tbl.removeChild(rows[at]);
  if (bands) applyBands(tbl, bands);
  syncFrame(frame, tbl);
  return serialize(doc);
}

/** Scale the columns so they add up to `total` again. */
function restoreWidth(tbl: Element, total: number): void {
  const grid = gridOf(tbl);
  const now = grid.reduce((sum, col) => sum + (parseInt(col.getAttribute("w") || "0", 10) || 0), 0);
  if (now <= 0 || total <= 0) return;
  let used = 0;
  grid.forEach((col, index) => {
    const width = index === grid.length - 1 ? total - used : Math.max(1, Math.round(((parseInt(col.getAttribute("w") || "0", 10) || 0) * total) / now));
    col.setAttribute("w", String(width));
    used += width;
  });
}

export function insertColumn(xml: string, at: number, texts: string[] = []): string {
  const { doc, frame, tbl } = parse(xml);
  refuseMerged(tbl, "增加列");
  const grid = gridOf(tbl);
  if (at < 0 || at > grid.length) throw new TableEditError(`只能在第 1 到 ${grid.length + 1} 列的位置插入，不能是 ${at + 1}`);
  const total = grid.reduce((sum, col) => sum + (parseInt(col.getAttribute("w") || "0", 10) || 0), 0);
  const sourceIndex = Math.max(0, Math.min(at, grid.length) - (at >= grid.length ? 1 : 0));
  const col = grid[sourceIndex].cloneNode(true) as Element;
  dropIds(col);
  if (at >= grid.length) grid[grid.length - 1].after(col);
  else grid[at].before(col);
  rowsOf(tbl).forEach((tr, r) => {
    const cells = cellsOf(tr);
    const fresh = blankCell(doc, cells[sourceIndex], texts[r] ?? "");
    if (at >= cells.length) cells[cells.length - 1].after(fresh);
    else cells[at].before(fresh);
  });
  restoreWidth(tbl, total);
  rowsOf(tbl).forEach((_, r) => fitRow(tbl, r));
  syncFrame(frame, tbl);
  return serialize(doc);
}

export function deleteColumn(xml: string, at: number): string {
  const { doc, frame, tbl } = parse(xml);
  refuseMerged(tbl, "删除列");
  const grid = gridOf(tbl);
  if (grid.length <= 1) throw new TableEditError("表格至少要保留一列");
  if (at < 0 || at >= grid.length) throw new TableEditError(`表格只有 ${grid.length} 列，没有第 ${at + 1} 列`);
  const total = grid.reduce((sum, col) => sum + (parseInt(col.getAttribute("w") || "0", 10) || 0), 0);
  grid[at].parentNode!.removeChild(grid[at]);
  for (const tr of rowsOf(tbl)) tr.removeChild(cellsOf(tr)[at]);
  restoreWidth(tbl, total);
  syncFrame(frame, tbl);
  return serialize(doc);
}

// ------------------------------------------------------------- a new table

export interface NewTable {
  rows: string[][];
  /** The first row is a header: accent colour, bold white text. */
  header: boolean;
  /** EMU; the frame is `width` wide and as tall as its rows. */
  width: number;
  /** Total height asked for (EMU); rows share the surplus when it is more than they need. */
  height?: number;
  /** Column widths in EMU; equal when absent. */
  columnWidths?: number[];
  fontPt: number;
  accent: string;
}

const tint = (color: string, factor: number): string => {
  const channel = (offset: number) => Math.round(parseInt(color.slice(offset, offset + 2), 16) * factor + 255 * (1 - factor));
  return `#${[1, 3, 5].map(offset => channel(offset).toString(16).padStart(2, "0")).join("")}`;
};

/**
 * A table frame with its formatting written into the cells (a header, light
 * bands, white rules), so it looks the same in every program and never leans on
 * a table style the file might lack.
 */
export function buildTableXml(
  id: string,
  name: string,
  box: { x: number; y: number },
  table: NewTable
): string {
  const columns = table.rows.reduce((most, row) => Math.max(most, row.length), 1);
  const widths = table.columnWidths && table.columnWidths.length === columns
    ? table.columnWidths
    : Array.from({ length: columns }, (_, c) => Math.floor(table.width / columns) + (c === columns - 1 ? table.width - Math.floor(table.width / columns) * columns : 0));
  const margin = { l: 91440, r: 91440, t: 45720, b: 45720 };
  const needed = table.rows.map(row => {
    let lines = 1;
    row.forEach((text, c) => {
      const available = Math.max(1, widths[c] / PT - (margin.l + margin.r) / PT);
      lines = Math.max(lines, text.split("\n").reduce((sum, line) => sum + Math.max(1, Math.ceil(textWidth(line, table.fontPt) / available)), 0));
    });
    return Math.round((lines * table.fontPt * 1.2 + (margin.t + margin.b) / PT) * PT);
  });
  const surplus = table.height ? Math.max(0, table.height - needed.reduce((sum, h) => sum + h, 0)) : 0;
  const heights = needed.map(h => h + Math.floor(surplus / needed.length));

  const rule = (tag: string, width = 12700) =>
    `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:prstDash val="solid"/></a:${tag}>`;
  const band = [tint(table.accent, 0.4), tint(table.accent, 0.2)];

  const rowXml = table.rows.map((row, r) => {
    const isHeader = table.header && r === 0;
    const fill = isHeader ? table.accent : band[(r - (table.header ? 1 : 0)) % 2];
    const color = isHeader ? "FFFFFF" : "1F1F1F";
    const cells = Array.from({ length: columns }, (_, c) => {
      const lines = (row[c] ?? "").split("\n");
      const paragraphs = lines.map(line =>
        `<a:p>${line === "" ? "" : `<a:r><a:rPr lang="zh-CN" sz="${Math.round(table.fontPt * 100)}" b="${isHeader ? 1 : 0}" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${escapeXml(line)}</a:t></a:r>`}`
        + `<a:endParaRPr lang="zh-CN" sz="${Math.round(table.fontPt * 100)}" b="${isHeader ? 1 : 0}" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:endParaRPr></a:p>`).join("");
      return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</a:txBody>`
        + `<a:tcPr marL="${margin.l}" marR="${margin.r}" marT="${margin.t}" marB="${margin.b}" anchor="ctr">`
        + `${rule("lnL")}${rule("lnR")}${rule("lnT")}${rule("lnB", isHeader ? 38100 : 12700)}<a:solidFill><a:srgbClr val="${hex(fill)}"/></a:solidFill></a:tcPr></a:tc>`;
    }).join("");
    return `<a:tr h="${heights[r]}">${cells}</a:tr>`;
  }).join("");

  const width = widths.reduce((sum, w) => sum + w, 0);
  const height = heights.reduce((sum, h) => sum + h, 0);
  return `<p:graphicFrame xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`
    + `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
    + `<p:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${width}" cy="${height}"/></p:xfrm>`
    + `<a:graphic><a:graphicData uri="${TABLE_URI}"><a:tbl><a:tblPr${table.header ? ' firstRow="1"' : ""} bandRow="1"/>`
    + `<a:tblGrid>${widths.map(w => `<a:gridCol w="${w}"/>`).join("")}</a:tblGrid>${rowXml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}
