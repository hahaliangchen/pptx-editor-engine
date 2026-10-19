import { ElementSource, Rect, TextBodyProperties, TextParagraph, TextStyle } from "../pptx-virtual-dom";

export interface TableBorder {
  color: string;
  /** px */
  width: number;
}

export interface TableCell {
  colSpan: number;
  rowSpan: number;
  /** Hidden under the merged cell to its left or above. */
  covered: boolean;
  /** Resolved background, from the cell itself or the table style; null for none. */
  fill: string | null;
  borders: { left: TableBorder | null; right: TableBorder | null; top: TableBorder | null; bottom: TableBorder | null };
  paragraphs: TextParagraph[];
  /** The cell's first run, for how an empty cell would look. */
  style: TextStyle;
  body: TextBodyProperties;
  /** Plain text, paragraphs joined by a line break. */
  text: string;
}

export interface TableModel {
  /** Column widths and row heights, px. */
  columns: number[];
  rows: Array<{ height: number; cells: TableCell[] }>;
  firstRow: boolean;
  bandRow: boolean;
  firstColumn: boolean;
  /** Any merged cell: row and column edits are not offered then. */
  merged: boolean;
}

export interface TableElement {
  type: "table";
  id: string;
  rect: Rect;
  source?: ElementSource;
  table: TableModel;
  rotation?: number;
  pxPerPt: number;
}

export const cloneTable = (table: TableModel): TableModel => JSON.parse(JSON.stringify(table));

/** The text of every cell, row by row. */
export const tableText = (table: TableModel): string[][] =>
  table.rows.map(row => row.cells.map(cell => (cell.covered ? "" : cell.text)));
