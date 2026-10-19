import { RenderElement } from "../pptx-virtual-dom";
import { TableBorder, TableElement } from "./model";

type Edge = TableBorder | null | undefined;

/** Of two borders that meet on one edge, the one that is there, then the heavier. */
function pick(a: Edge, b: Edge): Edge {
  if (a === undefined) return b;
  if (b === undefined) return a;
  if (a && b) return b.width >= a.width ? b : a;
  return a ?? b;
}

const sameBorder = (a: Edge, b: Edge) => (a?.color ?? null) === (b?.color ?? null) && (a?.width ?? 0) === (b?.width ?? 0);

/**
 * Draw a table as ordinary shapes and text: a fill per cell, one line per run of
 * equal border, and the cell's text. A merged cell is drawn once, over all the
 * grid cells it covers.
 */
export function expandTable(el: TableElement): RenderElement[] {
  const { table, rect } = el;
  const out: RenderElement[] = [];
  let serial = 0;
  const id = (kind: string) => `table_${el.id}_${kind}${serial++}`;

  // The frame may have been resized since the grid was read; the grid follows the frame.
  const gridWidth = table.columns.reduce((sum, width) => sum + width, 0);
  const gridHeight = table.rows.reduce((sum, row) => sum + row.height, 0);
  const scaleX = gridWidth > 0 ? rect.w / gridWidth : 1;
  const scaleY = gridHeight > 0 ? rect.h / gridHeight : 1;
  const xs = [rect.x];
  for (const width of table.columns) xs.push(xs[xs.length - 1] + width * scaleX);
  const ys = [rect.y];
  for (const row of table.rows) ys.push(ys[ys.length - 1] + row.height * scaleY);
  const cols = table.columns.length;
  const rows = table.rows.length;

  // Edges between grid cells: horizontal[r][c] is the line above row r over column c.
  const horizontal: Edge[][] = Array.from({ length: rows + 1 }, () => new Array<Edge>(cols).fill(undefined));
  const vertical: Edge[][] = Array.from({ length: cols + 1 }, () => new Array<Edge>(rows).fill(undefined));

  table.rows.forEach((row, r) => row.cells.forEach((cell, c) => {
    if (cell.covered) return;
    const endCol = Math.min(cols, c + cell.colSpan);
    const endRow = Math.min(rows, r + cell.rowSpan);
    const x = xs[c];
    const y = ys[r];
    const w = xs[endCol] - x;
    const h = ys[endRow] - y;
    if (cell.fill) {
      out.push({
        type: "shape", id: id("fill"), rect: { x, y, w, h }, shapeType: "rect", fill: cell.fill,
        computedStyle: { fill: { type: "solid", color: cell.fill } }
      });
    }
    for (let k = c; k < endCol; k++) {
      horizontal[r][k] = pick(horizontal[r][k], cell.borders.top);
      horizontal[endRow][k] = pick(horizontal[endRow][k], cell.borders.bottom);
    }
    for (let k = r; k < endRow; k++) {
      vertical[c][k] = pick(vertical[c][k], cell.borders.left);
      vertical[endCol][k] = pick(vertical[endCol][k], cell.borders.right);
    }
  }));

  const line = (x1: number, y1: number, x2: number, y2: number, border: TableBorder) => {
    out.push({
      type: "shape", id: id("line"),
      rect: { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) },
      shapeType: "line", flipH: false, flipV: false, fill: "transparent",
      border: { color: border.color, width: border.width },
      computedStyle: { fill: { type: "none" }, line: { fill: { type: "solid", color: border.color }, width: border.width } }
    });
  };
  // Runs of equal border along an edge become one line.
  const run = (edges: Edge[], draw: (from: number, to: number, border: TableBorder) => void) => {
    let start = 0;
    for (let i = 1; i <= edges.length; i++) {
      if (i < edges.length && sameBorder(edges[i], edges[start])) continue;
      const border = edges[start];
      if (border) draw(start, i, border);
      start = i;
    }
  };
  horizontal.forEach((edges, r) => run(edges, (from, to, border) => line(xs[from], ys[r], xs[to], ys[r], border)));
  vertical.forEach((edges, c) => run(edges, (from, to, border) => line(xs[c], ys[from], xs[c], ys[to], border)));

  table.rows.forEach((row, r) => row.cells.forEach((cell, c) => {
    if (cell.covered || cell.text === "") return;
    const endCol = Math.min(cols, c + cell.colSpan);
    const endRow = Math.min(rows, r + cell.rowSpan);
    out.push({
      type: "text", id: id("text"),
      rect: { x: xs[c], y: ys[r], w: xs[endCol] - xs[c], h: ys[endRow] - ys[r] },
      content: cell.text, style: cell.style, paragraphs: cell.paragraphs, body: cell.body
    });
  }));
  return out;
}
