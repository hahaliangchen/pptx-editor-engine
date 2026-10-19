import { describe, expect, it } from "vitest";
import { expandTable } from "../src/tables/layout";
import { TableCell, TableElement, TableModel } from "../src/tables/model";
import { builtinDefaultStyle, resolveCellStyle } from "../src/tables/styles";
import {
  TableEditError, buildTableXml, deleteColumn, deleteRow, insertColumn, insertRow, setCellStyle, setCellText, setTableFlags, tableShape
} from "../src/tables/xml";

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const frame = (rows: string[][], header = true) => buildTableXml("5", "表", { x: 914400, y: 914400 }, {
  rows, header, width: 6400800, fontPt: 14, accent: "#4472c4"
});

const doc = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");
const texts = (xml: string) => Array.from(doc(xml).getElementsByTagNameNS("*", "tr")).map(tr =>
  Array.from(tr.getElementsByTagNameNS("*", "tc")).map(tc => Array.from(tc.getElementsByTagNameNS("*", "t")).map(t => t.textContent).join("|")));
const ext = (xml: string) => {
  const node = Array.from(doc(xml).getElementsByTagNameNS("*", "ext")).find(e => (e.parentNode as Element).localName === "xfrm")!;
  return { cx: parseInt(node.getAttribute("cx")!, 10), cy: parseInt(node.getAttribute("cy")!, 10) };
};
const columnWidths = (xml: string) => Array.from(doc(xml).getElementsByTagNameNS("*", "gridCol")).map(c => parseInt(c.getAttribute("w")!, 10));
const rowHeights = (xml: string) => Array.from(doc(xml).getElementsByTagNameNS("*", "tr")).map(r => parseInt(r.getAttribute("h")!, 10));
/** The background colour of the first cell of each row. */
const fills = (xml: string) => Array.from(doc(xml).getElementsByTagNameNS("*", "tr")).map(tr => {
  const props = Array.from(tr.getElementsByTagNameNS("*", "tcPr"))[0];
  const solid = Array.from(props.children).find(child => child.localName === "solidFill");
  return [solid?.firstElementChild as Element];
});

describe("table styles", () => {
  const style = builtinDefaultStyle("#4472c4", "#000000", "#ffffff", (c, f) => c + String(f));
  const flags = { firstRow: true, lastRow: false, firstColumn: false, lastColumn: false, bandRow: true, bandColumn: false };
  const at = (row: number, col: number) => resolveCellStyle(style, flags, { row, col, rowSpan: 1, colSpan: 1 }, 4, 3);

  it("gives the header row its own fill, text colour and weight", () => {
    expect(at(0, 1)).toMatchObject({ fill: "#4472c4", textColor: "#ffffff", bold: true });
  });

  it("alternates the body rows", () => {
    expect(at(1, 0).fill).toBe("#4472c40.4");
    expect(at(2, 0).fill).toBe("#4472c40.2");
    expect(at(3, 0).fill).toBe("#4472c40.4");
  });

  it("draws the outer edge from the table's outside rule and inner edges from the inside rule", () => {
    expect(at(1, 0).borders.left?.width).toBe(12700);
    expect(at(0, 1).borders.bottom?.width).toBe(38100);
    expect(at(2, 1).borders.top?.color).toBe("#ffffff");
  });

  it("leaves a cell plain when no style is named", () => {
    expect(resolveCellStyle(undefined, flags, { row: 1, col: 0, rowSpan: 1, colSpan: 1 }, 4, 3).fill).toBeNull();
  });
});

describe("a new table", () => {
  it("has the rows and columns asked for, in a frame as big as the grid", () => {
    const xml = frame([["a", "b", "c"], ["1", "2", "3"]]);
    expect(texts(xml)).toEqual([["a", "b", "c"], ["1", "2", "3"]]);
    expect(columnWidths(xml).reduce((a, b) => a + b, 0)).toBe(ext(xml).cx);
    expect(rowHeights(xml).reduce((a, b) => a + b, 0)).toBe(ext(xml).cy);
    expect(tableShape(xml)).toEqual({ rows: 2, columns: 3, merged: false });
  });

  it("pads short rows, splits lines into paragraphs and escapes markup", () => {
    const xml = frame([["a", "b"], ["x\ny & <z>"]]);
    expect(texts(xml)[1]).toEqual(["x|y & <z>", ""]);
  });

  it("writes the formatting into the cells: header colour, white text, bands", () => {
    const xml = frame([["h"], ["1"], ["2"], ["3"]]);
    const colors = fills(xml).map(list => list[0].getAttribute("val"));
    expect(colors[0]).toBe("4472C4");
    expect(colors[1]).not.toBe(colors[2]);
    expect(colors[1]).toBe(colors[3]);
  });

  it("makes rows tall enough for long text and shares extra height", () => {
    const tall = frame([["a very long text ".repeat(20)]]);
    const short = frame([["a"]]);
    expect(rowHeights(tall)[0]).toBeGreaterThan(rowHeights(short)[0]);
    const spaced = buildTableXml("5", "t", { x: 0, y: 0 }, { rows: [["a"], ["b"]], header: true, width: 3000000, height: 3000000, fontPt: 14, accent: "#4472c4" });
    expect(ext(spaced).cy).toBeGreaterThanOrEqual(2900000);
  });

  it("honours column widths", () => {
    const xml = buildTableXml("5", "t", { x: 0, y: 0 }, { rows: [["a", "b"]], header: false, width: 0, columnWidths: [1000000, 3000000], fontPt: 14, accent: "#4472c4" });
    expect(columnWidths(xml)).toEqual([1000000, 3000000]);
  });
});

describe("editing cells", () => {
  it("replaces the text and keeps the cell's formatting", () => {
    const next = setCellText(frame([["h", "i"], ["1", "2"]]), 0, 1, "新标题");
    expect(texts(next)[0]).toEqual(["h", "新标题"]);
    const run = Array.from(doc(next).getElementsByTagNameNS("*", "rPr")).find(r => r.parentNode!.textContent === "新标题")!;
    expect(run.getAttribute("b")).toBe("1");
    expect(run.getAttribute("sz")).toBe("1400");
  });

  it("writes several lines as several paragraphs, and an empty text as an empty cell", () => {
    const next = setCellText(frame([["h"], ["1"]]), 1, 0, "甲\n乙");
    expect(texts(next)[1]).toEqual(["甲|乙"]);
    expect(texts(setCellText(next, 1, 0, ""))[1]).toEqual([""]);
  });

  it("grows the row when the text no longer fits, and the frame with it", () => {
    const xml = frame([["h"], ["1"]]);
    const next = setCellText(xml, 1, 0, "很长很长的文字".repeat(20));
    expect(rowHeights(next)[1]).toBeGreaterThan(rowHeights(xml)[1]);
    expect(ext(next).cy).toBeGreaterThan(ext(xml).cy);
  });

  it("names the cell it cannot reach", () => {
    expect(() => setCellText(frame([["a"]]), 5, 0, "x")).toThrow(TableEditError);
    expect(() => setCellText(frame([["a"]]), 0, 3, "x")).toThrow(/没有第 4 列/);
  });

  it("sets a fill, text style and alignment", () => {
    const next = setCellStyle(frame([["h"], ["1"]]), 1, 0, { fill: "#ff0000", bold: true, italic: true, color: "#00ff00", fontSizePt: 20, align: "center" });
    const cell = Array.from(doc(next).getElementsByTagNameNS("*", "tc"))[1];
    expect(cell.getElementsByTagNameNS("*", "tcPr")[0].lastElementChild!.getElementsByTagNameNS("*", "srgbClr")[0].getAttribute("val")).toBe("FF0000");
    const run = cell.getElementsByTagNameNS("*", "rPr")[0];
    expect([run.getAttribute("b"), run.getAttribute("i"), run.getAttribute("sz")]).toEqual(["1", "1", "2000"]);
    expect(run.getElementsByTagNameNS("*", "srgbClr")[0].getAttribute("val")).toBe("00FF00");
    expect(cell.getElementsByTagNameNS("*", "pPr")[0].getAttribute("algn")).toBe("ctr");
  });

  it("removes a fill with null, and keeps the schema's order inside tcPr", () => {
    const next = setCellStyle(frame([["h"], ["1"]]), 1, 0, { fill: null });
    const props = Array.from(doc(next).getElementsByTagNameNS("*", "tcPr"))[1];
    expect(Array.from(props.children).map(c => c.localName)).toEqual(["lnL", "lnR", "lnT", "lnB", "noFill"]);
  });

  it("turns the header and band flags on and off", () => {
    const next = setTableFlags(frame([["h"], ["1"]]), { firstRow: false, bandRow: false, firstColumn: true });
    const props = doc(next).getElementsByTagNameNS("*", "tblPr")[0];
    expect([props.getAttribute("firstRow"), props.getAttribute("bandRow"), props.getAttribute("firstCol")]).toEqual([null, null, "1"]);
  });
});

describe("rows and columns", () => {
  const base = frame([["h1", "h2"], ["a", "b"], ["c", "d"], ["e", "f"]]);

  it("inserts a row where asked, with the neighbour's look and the given text", () => {
    const next = insertRow(base, 2, ["x", "y"]);
    expect(texts(next).map(row => row.join(","))).toEqual(["h1,h2", "a,b", "x,y", "c,d", "e,f"]);
    expect(ext(next).cy).toBeGreaterThan(ext(base).cy);
    expect(ext(next).cy).toBe(rowHeights(next).reduce((a, b) => a + b, 0));
  });

  it("appends a row at the end and starts one with a blank cell for missing text", () => {
    expect(texts(insertRow(base, 4, ["z"]))[4]).toEqual(["z", ""]);
  });

  it("keeps the bands alternating after a row is added or removed", () => {
    const colors = (xml: string) => fills(xml).map(list => list[0].getAttribute("val"));
    const added = colors(insertRow(base, 2));
    expect(added.slice(1).every((c, i) => i === 0 || c !== added[i])).toBe(true);
    const removed = colors(deleteRow(base, 1));
    expect(removed.slice(1).every((c, i) => i === 0 || c !== removed[i])).toBe(true);
  });

  it("deletes a row and refuses to delete the last one", () => {
    expect(texts(deleteRow(base, 1)).length).toBe(3);
    expect(() => deleteRow(frame([["a"]]), 0)).toThrow(/至少要保留一行/);
    expect(() => deleteRow(base, 9)).toThrow(TableEditError);
  });

  it("inserts a column and keeps the table as wide as it was", () => {
    const next = insertColumn(base, 1, ["n1", "n2", "n3", "n4"]);
    expect(texts(next).map(row => row.join(","))).toEqual(["h1,n1,h2", "a,n2,b", "c,n3,d", "e,n4,f"]);
    expect(ext(next).cx).toBe(ext(base).cx);
    expect(columnWidths(next).reduce((a, b) => a + b, 0)).toBe(ext(base).cx);
    expect(columnWidths(next)).toHaveLength(3);
  });

  it("deletes a column and widens the rest to fill the table", () => {
    const next = deleteColumn(base, 0);
    expect(texts(next)[0]).toEqual(["h2"]);
    expect(columnWidths(next)[0]).toBe(ext(base).cx);
    expect(() => deleteColumn(frame([["a"]]), 0)).toThrow(/至少要保留一列/);
  });

  it("does not repeat the identifiers PowerPoint keeps unique", () => {
    const withIds = base
      .replace("<a:gridCol w", '<a:gridCol w').replace(/<a:tr h="(\d+)">/g, (_m, h) => `<a:tr h="${h}"><a:extLst><a:ext uri="{X}"><x:rowId xmlns:x="urn:x" val="1"/></a:ext></a:extLst>`.replace("<a:extLst>", "<a:extLst>"));
    // extLst belongs after the cells; the point is only that copies drop it.
    const next = insertRow(withIds.replace(/<a:extLst>.*?<\/a:extLst>/g, ""), 1);
    expect(next).not.toContain("rowId");
  });

  it("refuses row and column edits on a table with merged cells, saying why", () => {
    const merged = base.replace("<a:tc>", '<a:tc gridSpan="2">');
    expect(tableShape(merged).merged).toBe(true);
    for (const edit of [() => insertRow(merged, 1), () => deleteRow(merged, 1), () => insertColumn(merged, 1), () => deleteColumn(merged, 1)]) {
      expect(edit).toThrow(/合并单元格/);
    }
  });

  it("will not write into a cell a merge covers", () => {
    const covered = base.replace(/<a:tc>/, '<a:tc gridSpan="2">').replace(/<a:tc>/, '<a:tc hMerge="1">');
    expect(() => setCellText(covered, 0, 1, "x")).toThrow(/盖住/);
  });
});

describe("drawing a table", () => {
  const cell = (text: string, over: Partial<TableCell> = {}): TableCell => ({
    colSpan: 1, rowSpan: 1, covered: false, fill: null,
    borders: { left: null, right: null, top: null, bottom: null },
    paragraphs: [], style: { fontSize: 20, color: "#000000", bold: false, align: "left" },
    body: { marginLeft: 0, marginRight: 0, marginTop: 0, marginBottom: 0, verticalAnchor: "top", autoFit: "none" },
    text, ...over
  });
  const edge = { color: "#000000", width: 2 };
  const element = (table: Partial<TableModel>, w = 200, h = 100): TableElement => ({
    type: "table", id: "t", pxPerPt: 1, rect: { x: 10, y: 20, w, h },
    table: { columns: [100, 100], rows: [{ height: 50, cells: [cell("a"), cell("b")] }, { height: 50, cells: [cell("c"), cell("d")] }], firstRow: false, bandRow: false, firstColumn: false, merged: false, ...table }
  });

  it("draws a fill per filled cell and a text per non-empty cell", () => {
    const out = expandTable(element({
      rows: [{ height: 50, cells: [cell("a", { fill: "#ff0000" }), cell("")] }, { height: 50, cells: [cell("c"), cell("d", { fill: "#00ff00" })] }]
    }));
    expect(out.filter(e => e.type === "shape")).toHaveLength(2);
    expect(out.filter(e => e.type === "text").map(e => (e as { content: string }).content)).toEqual(["a", "c", "d"]);
  });

  it("puts cells where the grid says, offset by the frame", () => {
    const out = expandTable(element({}));
    const text = out.filter(e => e.type === "text").map(e => e.rect);
    expect(text[0]).toMatchObject({ x: 10, y: 20, w: 100, h: 50 });
    expect(text[3]).toMatchObject({ x: 110, y: 70, w: 100, h: 50 });
  });

  it("follows the frame when it was resized", () => {
    const out = expandTable(element({}, 400, 300));
    const text = out.filter(e => e.type === "text").map(e => e.rect);
    expect(text[3]).toMatchObject({ x: 210, y: 170, w: 200, h: 150 });
  });

  it("draws a merged cell once, over what it covers", () => {
    const out = expandTable(element({
      rows: [{ height: 50, cells: [cell("wide", { colSpan: 2, fill: "#ff0000" }), cell("", { covered: true })] }, { height: 50, cells: [cell("c"), cell("d")] }]
    }));
    const fill = out.find(e => e.type === "shape")!;
    expect(fill.rect).toMatchObject({ x: 10, y: 20, w: 200, h: 50 });
    expect(out.filter(e => e.type === "text")).toHaveLength(3);
  });

  it("draws a border shared by two cells once, as one line along the run", () => {
    const bordered = (left: boolean, top: boolean) => cell("x", { borders: { left: left ? edge : null, right: null, top: top ? edge : null, bottom: edge } });
    const out = expandTable(element({
      rows: [{ height: 50, cells: [bordered(false, true), bordered(false, true)] }, { height: 50, cells: [bordered(false, false), bordered(false, false)] }]
    }));
    const lines = out.filter(e => e.type === "shape");
    // top edge: one line over both columns; bottom edge of row 1 = top of row 2: one; bottom of row 2: one
    expect(lines).toHaveLength(3);
    expect(lines.find(l => l.rect.y === 20)!.rect.w).toBe(200);
  });

  it("lets a border that exists win over a missing one on a shared edge", () => {
    const out = expandTable(element({
      rows: [{ height: 50, cells: [cell("a", { borders: { left: null, right: null, top: null, bottom: edge } }), cell("b")] }, { height: 50, cells: [cell("c", { borders: { left: null, right: null, top: null, bottom: null } }), cell("d")] }]
    }));
    expect(out.filter(e => e.type === "shape" && (e as { rect: { y: number } }).rect.y === 70)).toHaveLength(1);
  });
});
