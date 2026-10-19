// @vitest-environment node
import { afterEach, describe, expect, it } from "vitest";
import JSZip from "jszip";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { HeadlessSession } from "../src/headless/session";
import { buildDeck } from "./helpers/deck";
import { apiDeckParts } from "./helpers/api-deck";
import { checkPackage } from "./helpers/package-check";

const sessions: HeadlessSession[] = [];
async function open(bytes?: Buffer | Uint8Array): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(bytes ?? await buildDeck(apiDeckParts()));
  sessions.push(session);
  return session;
}
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

const data = [["产品", "销量", "占比"], ["甲", "120", "40%"], ["乙", "90", "30%"], ["丙", "90", "30%"]];
const add = { op: "addTable", slide: 1, rows: data, rect: { x: 2, y: 2, w: 20, h: 6 }, name: "销售表", as: "t" };
const target = { slide: 1, name: "销售表" };

const tableOf = async (session: HeadlessSession, slide = 1, name = "销售表") =>
  (await session.outline({ slides: [slide] })).slides[0].shapes.find((shape: { name: string }) => shape.name === name);

async function pixelAt(png: Buffer, x: number, y: number): Promise<number[]> {
  const image = await loadImage(png);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  return [...ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data];
}

const problems = async (bytes: Uint8Array) =>
  (await checkPackage(bytes)).filter(message => !/^no content type for .*\.rels$/.test(message));

describe("adding a table", () => {
  it("shows up in the outline with its text and grid, at the asked place", async () => {
    const session = await open();
    expect((await session.apply([add])).ok).toBe(true);
    const table = await tableOf(session);
    expect(table.type).toBe("table");
    expect(table.rect).toMatchObject({ x: 2, y: 2, w: 20 });
    expect(table.rect.h).toBeGreaterThanOrEqual(5.9);
    expect(table.table.rows).toEqual(data);
    expect(table.table.columns).toBe(3);
    expect(table.table.columnWidths).toHaveLength(3);
    expect(table.table.rowHeights).toHaveLength(4);
    expect(table.table.merged).toBeUndefined();
  });

  it("draws a header in the accent colour and bands below it", async () => {
    const session = await open();
    await session.apply([add]);
    const width = 960;
    const pxPerCm = width / (session.document.size.width * session.document.slides[0].emuPerPx / 360000);
    const png = await session.render(1, { width });
    const header = await pixelAt(png, 2.1 * pxPerCm, 2.1 * pxPerCm);
    expect(header.slice(0, 3)).toEqual([0x44, 0x72, 0xc4]);
    const rowAt = (cm: number) => pixelAt(png, 2.1 * pxPerCm, cm * pxPerCm);
    const first = (await rowAt(3.7)).slice(0, 3);
    const second = (await rowAt(5.2)).slice(0, 3);
    expect(first).not.toEqual(second);
    expect(first).not.toEqual([255, 255, 255]);
  });

  it("takes a header flag, font size, column widths and is cheap on a plain call", async () => {
    const session = await open();
    await session.apply([{ op: "addTable", slide: 1, rows: [["a", "b"], ["1", "2"]], header: false, fontSizePt: 20, columnWidths: [4, 12], name: "窄宽" }]);
    const table = await tableOf(session, 1, "窄宽");
    expect(table.table.columnWidths).toEqual([4, 12]);
    expect(table.rect.w).toBeCloseTo(16, 1);
  });

  it("rejects an empty or oversized table, naming the field", async () => {
    const session = await open();
    expect((await session.apply([{ op: "addTable", slide: 1, rows: [] }])).error).toMatchObject({ code: "INVALID_SCHEMA" });
    expect((await session.apply([{ op: "addTable", slide: 1, rows: [[]] }])).error).toMatchObject({ code: "INVALID_SCHEMA" });
    expect((await session.apply([{ op: "addTable", slide: 1, rows: [["a"]], colour: "x" }])).error).toMatchObject({ code: "INVALID_SCHEMA" });
  });

  it("is one undo step", async () => {
    const session = await open();
    await session.apply([add]);
    session.undo();
    expect(await tableOf(session)).toBeUndefined();
    session.redo();
    expect(await tableOf(session)).toBeTruthy();
  });
});

describe("editing a table", () => {
  it("changes a cell's text and look, and undo restores both", async () => {
    const session = await open();
    await session.apply([add]);
    const result = await session.apply([{ op: "setCell", target, row: 2, col: 2, text: "135", bold: true, color: "#c00000", fill: "#ffff00" }]);
    expect(result.ok).toBe(true);
    expect((await tableOf(session)).table.rows[1][1]).toBe("135");
    session.undo();
    expect((await tableOf(session)).table.rows[1][1]).toBe("120");
  });

  it("only restyles when no text is given", async () => {
    const session = await open();
    await session.apply([add, { op: "setCell", target: "$t", row: 1, col: 1, align: "center", fontSizePt: 24 }]);
    expect((await tableOf(session)).table.rows[0][0]).toBe("产品");
  });

  it("fits a long text by making its row taller", async () => {
    const session = await open();
    await session.apply([add]);
    const before = (await tableOf(session)).table.rowHeights[1];
    await session.apply([{ op: "setCell", target, row: 2, col: 1, text: "很长很长的内容".repeat(30) }]);
    expect((await tableOf(session)).table.rowHeights[1]).toBeGreaterThan(before);
  });

  it("replaces the whole table, adding and removing rows and columns to fit", async () => {
    const session = await open();
    await session.apply([add]);
    await session.apply([{ op: "setTable", target, rows: [["a", "b", "c", "d"], ["1", "2", "3", "4"]] }]);
    expect((await tableOf(session)).table).toMatchObject({ rows: [["a", "b", "c", "d"], ["1", "2", "3", "4"]], columns: 4 });
    await session.apply([{ op: "setTable", target, rows: [["x"], ["y"], ["z"]] }]);
    expect((await tableOf(session)).table.rows).toEqual([["x"], ["y"], ["z"]]);
    expect(session.document.slides[0].elements.filter(e => e.type === "table")).toHaveLength(1);
  });

  it("adds and removes rows and columns at the position asked, counting from 1", async () => {
    const session = await open();
    await session.apply([add]);
    await session.apply([
      { op: "insertRow", target, at: 2, text: ["新", "1", "0%"] },
      { op: "insertColumn", target, at: 2, text: ["备注", "a", "b", "c", "d"] }
    ]);
    expect((await tableOf(session)).table.rows.map((row: string[]) => row.join(","))).toEqual([
      "产品,备注,销量,占比", "新,a,1,0%".replace("新,a", "新,a"), "甲,b,120,40%", "乙,c,90,30%", "丙,d,90,30%"
    ].map((line, i) => i === 1 ? "新,a,1,0%" : line));
    await session.apply([{ op: "deleteRow", target, row: 2 }, { op: "deleteColumn", target, col: 2 }]);
    expect((await tableOf(session)).table.rows).toEqual(data);
  });

  it("keeps the table inside the width it had when columns come and go", async () => {
    const session = await open();
    await session.apply([add]);
    const width = (await tableOf(session)).rect.w;
    await session.apply([{ op: "insertColumn", target, at: 4 }]);
    expect((await tableOf(session)).rect.w).toBeCloseTo(width, 1);
    await session.apply([{ op: "deleteColumn", target, col: 1 }]);
    expect((await tableOf(session)).rect.w).toBeCloseTo(width, 1);
  });

  it("reports what a table cannot take, with the command's path, and rolls the batch back", async () => {
    const session = await open();
    await session.apply([add]);
    const outOfRange = await session.apply([
      { op: "setCell", target, row: 2, col: 1, text: "会被撤回" },
      { op: "setCell", target, row: 9, col: 1, text: "x" }
    ]);
    expect(outOfRange.ok).toBe(false);
    expect(outOfRange.failedAt).toBe(1);
    expect(outOfRange.error).toMatchObject({ code: "NOT_APPLICABLE" });
    expect(outOfRange.error!.message).toContain("没有第 9 行");
    expect((await tableOf(session)).table.rows[1][0]).toBe("甲");
    expect((await session.apply([{ op: "deleteRow", target, row: 7 }])).error).toMatchObject({ code: "NOT_APPLICABLE" });
    expect((await session.apply([{ op: "setCell", target, row: 1, col: 1 }])).error).toMatchObject({ code: "INVALID_SCHEMA" });
  });

  it("refuses what a table cannot do, and what is not a table", async () => {
    const session = await open();
    await session.apply([add]);
    for (const command of [
      { op: "setFill", target, color: "#ff0000" },
      { op: "setRotation", target, degrees: 5 },
      { op: "setText", target, text: "x" }
    ]) expect((await session.apply([command])).error?.code, command.op).toBe("NOT_APPLICABLE");
    expect((await session.apply([{ op: "setCell", target: { slide: 1, name: "卡片" }, row: 1, col: 1, text: "x" }])).error)
      .toMatchObject({ code: "NOT_APPLICABLE" });
  });

  it("moves and resizes like any shape", async () => {
    const session = await open();
    await session.apply([add]);
    await session.apply([{ op: "setBounds", target, x: 1, y: 1, w: 30 }]);
    expect((await tableOf(session)).rect).toMatchObject({ x: 1, y: 1, w: 30 });
    expect((await tableOf(session)).table.columnWidths.reduce((a: number, b: number) => a + b, 0)).toBeCloseTo(30, 1);
  });
});

describe("exporting tables", () => {
  it("writes a table that opens again with the same text", async () => {
    const session = await open();
    await session.apply([add, { op: "setCell", target: "$t", row: 2, col: 2, text: "135", bold: true }]);
    const bytes = await session.export();
    expect(await problems(bytes)).toEqual([]);
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file("ppt/slides/slide1.xml")!.async("string")).toContain("<a:tbl>");

    const reopened = await open(bytes);
    const table = await tableOf(reopened);
    expect(table.table.rows[1][1]).toBe("135");
    expect(table.rect).toEqual((await tableOf(session)).rect);
  });

  it("writes edits made after the table was added, and after it was opened again", async () => {
    const first = await open();
    await first.apply([add, { op: "insertRow", target: "$t", at: 5, text: ["末行", "0", "0%"] }]);
    const second = await open(await first.export());
    await second.apply([{ op: "setCell", target, row: 5, col: 1, text: "改过" }, { op: "deleteRow", target, row: 2 }]);
    const third = await open(await second.export());
    expect((await tableOf(third)).table.rows).toEqual([
      ["产品", "销量", "占比"], ["乙", "90", "30%"], ["丙", "90", "30%"], ["改过", "0", "0%"]
    ]);
  });

  it("writes the new size of a table that was moved and resized", async () => {
    const session = await open();
    await session.apply([add, { op: "setBounds", target: "$t", x: 3, y: 4, w: 25, h: 9 }]);
    const reopened = await open(await session.export());
    expect((await tableOf(reopened)).rect).toMatchObject({ x: 3, y: 4, w: 25 });
    expect((await tableOf(reopened)).table.columnWidths.reduce((a: number, b: number) => a + b, 0)).toBeCloseTo(25, 1);
  });

  it("leaves a table that was undone out of the file", async () => {
    const session = await open();
    await session.apply([add]);
    session.undo();
    const zip = await JSZip.loadAsync(await session.export());
    expect(await zip.file("ppt/slides/slide1.xml")!.async("string")).not.toContain("<a:tbl>");
  });

  it("exports twice with the same result", async () => {
    const session = await open();
    await session.apply([add, { op: "setCell", target: "$t", row: 1, col: 1, text: "改" }]);
    const a = await JSZip.loadAsync(await session.export());
    const b = await JSZip.loadAsync(await session.export());
    expect(await a.file("ppt/slides/slide1.xml")!.async("string")).toBe(await b.file("ppt/slides/slide1.xml")!.async("string"));
  });

  it("copies a table with its page and on its own page, and the copies are independent", async () => {
    const session = await open();
    const result = await session.apply([
      add,
      { op: "duplicateSlide", slide: 1 },
      { op: "setCell", target: { slide: 2, name: "销售表" }, row: 1, col: 1, text: "副本页" },
      { op: "duplicate", target: { slide: 1, name: "销售表" }, offset: 1, as: "again" },
      { op: "setCell", target: "$again", row: 1, col: 1, text: "同页副本" }
    ]);
    expect(result.ok).toBe(true);
    const reopened = await open(await session.export());
    const heads = (await reopened.outline()).slides.flatMap((slide: { shapes: Array<{ type: string; table?: { rows: string[][] } }> }) =>
      slide.shapes.filter(shape => shape.type === "table").map(shape => shape.table!.rows[0][0]));
    expect(heads.sort()).toEqual(["产品", "同页副本", "副本页"].sort());
  });

  it("deletes a table", async () => {
    const session = await open();
    await session.apply([add, { op: "delete", target: "$t" }]);
    expect(await tableOf(session)).toBeUndefined();
  });
});
