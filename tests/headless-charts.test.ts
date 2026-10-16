// @vitest-environment node
import { afterEach, describe, expect, it } from "vitest";
import JSZip from "jszip";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { HeadlessSession } from "../src/headless/session";
import { expandChart } from "../src/charts/layout";
import { ChartElement } from "../src/charts/model";
import { ShapeElement } from "../src/pptx-virtual-dom";
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

const sales = {
  op: "addChart", slide: 1, type: "column", title: "季度营收",
  categories: ["一季度", "二季度", "三季度"],
  series: [{ name: "华东", values: [10, 20, 30] }, { name: "华北", values: [5, 15, 25], color: "#c00000" }],
  rect: { x: 1, y: 1, w: 14, h: 8 }, name: "营收图", as: "sales"
};

const chartOf = async (session: HeadlessSession, slide = 1, name = "营收图") => {
  const shapes = (await session.outline({ slides: [slide] })).slides[0].shapes;
  return shapes.find((shape: { name: string }) => shape.name === name);
};

async function pixelAt(png: Buffer, x: number, y: number): Promise<number[]> {
  const image = await loadImage(png);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  return [...ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data];
}

/** The package problems that are the export's doing (the bare fixture never types its .rels parts). */
const problems = async (bytes: Uint8Array) =>
  (await checkPackage(bytes)).filter(message => !/^no content type for .*\.rels$/.test(message));

describe("adding a chart", () => {
  it("shows up in the outline with its data, at the asked place", async () => {
    const session = await open();
    const result = await session.apply([sales]);
    expect(result.ok).toBe(true);
    const chart = await chartOf(session);
    expect(chart.type).toBe("chart");
    expect(chart.rect).toEqual({ x: 1, y: 1, w: 14, h: 8 });
    expect(chart.chart).toMatchObject({
      type: "column", title: "季度营收", legend: "bottom", dataLabels: false, stacked: false,
      categories: ["一季度", "二季度", "三季度"]
    });
    expect(chart.chart.series).toEqual([
      { name: "华东", values: [10, 20, 30] },
      { name: "华北", values: [5, 15, 25], color: "#c00000" }
    ]);
  });

  it("draws, with the theme's colour for a series that names none", async () => {
    const session = await open();
    await session.apply([sales]);
    const slide = session.document.slides[0];
    const element = slide.elements.find((e): e is ChartElement => e.type === "chart")!;
    const bar = expandChart(element).filter((e): e is ShapeElement => e.type === "shape" && e.fill === "#4472c4" && Math.abs(e.rect.w - e.rect.h) > 1)[0];
    const width = 960;
    const scale = width / session.document.size.width;
    const png = await session.render(1, { width });
    const [r, g, b] = await pixelAt(png, (bar.rect.x + bar.rect.w / 2) * scale, (bar.rect.y + bar.rect.h / 2) * scale);
    expect([r, g, b]).toEqual([0x44, 0x72, 0xc4]);
  });

  it("gives every kind a sensible default look", async () => {
    const session = await open();
    await session.apply([
      { op: "addChart", slide: 1, type: "pie", categories: ["a", "b"], series: [{ name: "S", values: [1, 2] }], name: "饼" },
      { op: "addChart", slide: 1, type: "line", categories: ["a", "b"], series: [{ name: "S", values: [1, 2] }], name: "线" },
      { op: "addChart", slide: 1, type: "bar", categories: ["a", "b"], series: [{ name: "S", values: [1, 2] }], name: "条" }
    ]);
    expect((await chartOf(session, 1, "饼")).chart).toMatchObject({ legend: "right", dataLabels: true });
    expect((await chartOf(session, 1, "线")).chart).toMatchObject({ legend: "bottom", dataLabels: false });
    expect((await chartOf(session, 1, "条")).chart.type).toBe("bar");
  });

  it("rejects data that does not add up, naming the field", async () => {
    const session = await open();
    const short = await session.apply([{ ...sales, series: [{ name: "华东", values: [1, 2] }] }]);
    expect(short.ok).toBe(false);
    expect(short.error).toMatchObject({ code: "INVALID_SCHEMA", path: "commands[0].series[0].values" });
    expect(short.error!.message).toContain("两者必须一样多");

    const pie = await session.apply([{ ...sales, type: "pie" }]);
    expect(pie.error).toMatchObject({ code: "INVALID_SCHEMA", path: "commands[0].series" });

    const negative = await session.apply([{ ...sales, type: "pie", series: [{ name: "s", values: [1, -2, 3] }] }]);
    expect(negative.error!.message).toContain("不能为负");

    const unknown = await session.apply([{ ...sales, colour: "red" }]);
    expect(unknown.error).toMatchObject({ code: "INVALID_SCHEMA" });
    expect((await session.outline()).slides[0].shapes.some((shape: { type: string }) => shape.type === "chart")).toBe(false);
  });

  it("is one undo step", async () => {
    const session = await open();
    await session.apply([sales]);
    session.undo();
    expect((await session.outline()).slides[0].shapes.some((shape: { type: string }) => shape.type === "chart")).toBe(false);
    session.redo();
    expect(await chartOf(session)).toBeTruthy();
  });
});

describe("changing a chart", () => {
  it("replaces data and settings, and undo brings the old ones back", async () => {
    const session = await open();
    await session.apply([sales]);
    const result = await session.apply([{
      op: "setChart", target: { slide: 1, name: "营收图" }, type: "line", title: null, legend: "right", dataLabels: true,
      categories: ["甲", "乙"], series: [{ name: "唯一", values: [3, 4] }]
    }]);
    expect(result.ok).toBe(true);
    expect((await chartOf(session)).chart).toEqual({
      type: "line", categories: ["甲", "乙"], series: [{ name: "唯一", values: [3, 4] }],
      legend: "right", dataLabels: true, stacked: false
    });
    session.undo();
    expect((await chartOf(session)).chart).toMatchObject({ type: "column", title: "季度营收", categories: ["一季度", "二季度", "三季度"] });
  });

  it("refuses what a chart cannot take", async () => {
    const session = await open();
    await session.apply([sales]);
    const target = { slide: 1, name: "营收图" };
    for (const command of [
      { op: "setFill", target, color: "#ff0000" },
      { op: "setRotation", target, degrees: 10 },
      { op: "setText", target, text: "x" }
    ]) {
      const result = await session.apply([command]);
      expect(result.error?.code, command.op).toBe("NOT_APPLICABLE");
    }
    const notChart = await session.apply([{ op: "setChart", target: { slide: 1, name: "卡片" }, title: "x" }]);
    expect(notChart.error).toMatchObject({ code: "NOT_APPLICABLE" });
  });

  it("can move and resize, and draws afresh for the new size", async () => {
    const session = await open();
    await session.apply([sales, { op: "setBounds", target: "$sales", w: 20, h: 10 }]);
    expect((await chartOf(session)).rect).toMatchObject({ w: 20, h: 10 });
    const before = (session.document.slides[0].elements.find((e): e is ChartElement => e.type === "chart")!);
    const wide = expandChart(before).filter((e): e is ShapeElement => e.type === "shape" && e.fill === "#4472c4" && Math.abs(e.rect.w - e.rect.h) > 1)[0].rect.w;
    await session.apply([{ op: "setBounds", target: { slide: 1, name: "营收图" }, w: 10 }]);
    const narrow = expandChart(before).filter((e): e is ShapeElement => e.type === "shape" && e.fill === "#4472c4" && Math.abs(e.rect.w - e.rect.h) > 1)[0].rect.w;
    expect(narrow).toBeLessThan(wide);
  });
});

describe("exporting charts", () => {
  it("writes the chart, its workbook and every link between them", async () => {
    const session = await open();
    await session.apply([sales]);
    const bytes = await session.export();
    expect(await problems(bytes)).toEqual([]);

    const zip = await JSZip.loadAsync(bytes);
    expect(Object.keys(zip.files)).toEqual(expect.arrayContaining([
      "ppt/charts/chart1.xml", "ppt/charts/_rels/chart1.xml.rels", "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"
    ]));
    const types = await zip.file("[Content_Types].xml")!.async("string");
    expect(types).toContain('PartName="/ppt/charts/chart1.xml"');
    expect(types).toContain("drawingml.chart+xml");
    expect(types).toContain('Extension="xlsx"');
    expect(await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string")).toContain('Target="../charts/chart1.xml"');
    const slide = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect(slide).toContain("<p:graphicFrame");
    expect(slide).toContain("drawingml/2006/chart");
    expect(await zip.file("ppt/charts/_rels/chart1.xml.rels")!.async("string")).toContain("../embeddings/Microsoft_Excel_Sheet1.xlsx");
  });

  it("puts the data in the workbook, where Edit Data in PowerPoint looks for it", async () => {
    const session = await open();
    await session.apply([sales]);
    const zip = await JSZip.loadAsync(await session.export());
    const workbook = await JSZip.loadAsync(await zip.file("ppt/embeddings/Microsoft_Excel_Sheet1.xlsx")!.async("uint8array"));
    const sheet = await workbook.file("xl/worksheets/sheet1.xml")!.async("string");
    for (const text of ["华东", "华北", "一季度", "三季度"]) expect(sheet).toContain(text);
    for (const value of ["10", "30", "25"]) expect(sheet).toContain(`<v>${value}</v>`);
    expect(sheet).toContain('r="C4"');
  });

  it("reads back as the same chart", async () => {
    const session = await open();
    await session.apply([sales]);
    const reopened = await open(await session.export());
    const { series: before, ...rest } = (await chartOf(session)).chart;
    const { series: after, ...restAfter } = (await chartOf(reopened)).chart;
    expect(restAfter).toEqual(rest);
    // The file names every series colour; the new chart only named one.
    expect(after.map((s: { name: string; values: number[] }) => [s.name, s.values])).toEqual(before.map((s: { name: string; values: number[] }) => [s.name, s.values]));
    expect(after.map((s: { color: string }) => s.color)).toEqual(["#4472c4", "#c00000"]);
    expect((await chartOf(reopened)).rect).toEqual((await chartOf(session)).rect);
  });

  it("writes edits made after the chart was added", async () => {
    const session = await open();
    await session.apply([sales]);
    await session.apply([{ op: "setChart", target: "$sales", title: "改过的", series: [{ name: "新", values: [7, 8, 9] }] }].map(c => ({ ...c, target: { slide: 1, name: "营收图" } })));
    const reopened = await open(await session.export());
    expect((await chartOf(reopened)).chart).toMatchObject({ title: "改过的", series: [{ name: "新", values: [7, 8, 9] }] });
  });

  it("leaves nothing behind for a chart that was undone", async () => {
    const session = await open();
    await session.apply([sales]);
    session.undo();
    const zip = await JSZip.loadAsync(await session.export());
    expect(Object.keys(zip.files).filter(name => name.includes("chart"))).toEqual([]);
  });

  it("exports twice with the same result", async () => {
    const session = await open();
    await session.apply([sales]);
    const a = await JSZip.loadAsync(await session.export());
    const b = await JSZip.loadAsync(await session.export());
    expect(await a.file("ppt/charts/chart1.xml")!.async("string")).toBe(await b.file("ppt/charts/chart1.xml")!.async("string"));
  });
});

describe("editing a chart of an opened deck", () => {
  async function deckWithChart() {
    const first = await open();
    await first.apply([sales]);
    return open(await first.export());
  }

  it("changes its data and keeps it a chart of its own file", async () => {
    const session = await deckWithChart();
    const result = await session.apply([{
      op: "setChart", target: { slide: 1, name: "营收图" }, categories: ["Q1", "Q2"],
      series: [{ name: "线上", values: [1, 2] }, { name: "线下", values: [3, 4] }, { name: "海外", values: [5, 6] }]
    }]);
    expect(result.ok).toBe(true);

    const bytes = await session.export();
    expect(await problems(bytes)).toEqual([]);
    const zip = await JSZip.loadAsync(bytes);
    const workbook = await JSZip.loadAsync(await zip.file("ppt/embeddings/Microsoft_Excel_Sheet1.xlsx")!.async("uint8array"));
    expect(await workbook.file("xl/worksheets/sheet1.xml")!.async("string")).toContain("海外");

    const reopened = await open(bytes);
    expect((await chartOf(reopened)).chart).toMatchObject({
      categories: ["Q1", "Q2"],
      series: [{ name: "线上", values: [1, 2] }, { name: "线下", values: [3, 4] }, { name: "海外", values: [5, 6] }]
    });
  });

  it("can change the kind, which writes the chart afresh", async () => {
    const session = await deckWithChart();
    await session.apply([{ op: "setChart", target: { slide: 1, name: "营收图" }, type: "bar", stacked: true }]);
    const reopened = await open(await session.export());
    expect((await chartOf(reopened)).chart).toMatchObject({ type: "bar", stacked: true });
  });

  it("copies the page with its chart, and the copy is independent", async () => {
    const session = await deckWithChart();
    const result = await session.apply([
      { op: "duplicateSlide", slide: 1 },
      { op: "setChart", target: { slide: 2, name: "营收图" }, title: "副本", series: [{ name: "只有一个", values: [1, 1, 1] }] }
    ]);
    expect(result.ok).toBe(true);

    const bytes = await session.export();
    expect(await problems(bytes)).toEqual([]);
    const zip = await JSZip.loadAsync(bytes);
    const charts = Object.keys(zip.files).filter(name => /^ppt\/charts\/chart\d+\.xml$/.test(name));
    expect(charts).toHaveLength(2);
    const workbooks = Object.keys(zip.files).filter(name => name.startsWith("ppt/embeddings/") && name.endsWith(".xlsx"));
    expect(workbooks).toHaveLength(2);

    const reopened = await open(bytes);
    expect((await chartOf(reopened, 1)).chart).toMatchObject({ title: "季度营收" });
    expect((await chartOf(reopened, 1)).chart.series).toHaveLength(2);
    expect((await chartOf(reopened, 2)).chart).toMatchObject({ title: "副本", series: [{ name: "只有一个" }] });
  });

  it("copies a chart on the same page, and to another page", async () => {
    const session = await deckWithChart();
    const result = await session.apply([
      { op: "duplicate", target: { slide: 1, name: "营收图" }, offset: 1, as: "copy" },
      { op: "setChart", target: "$copy", title: "同页副本" },
      { op: "addSlide", after: 1 },
      { op: "copyTo", target: { slide: 1, name: "营收图" }, slide: 2 }
    ]);
    expect(result.ok).toBe(true);

    const bytes = await session.export();
    expect(await problems(bytes)).toEqual([]);
    const reopened = await open(bytes);
    const titles = (await reopened.outline()).slides.flatMap((slide: { shapes: Array<{ type: string; chart?: { title?: string } }> }) =>
      slide.shapes.filter(shape => shape.type === "chart").map(shape => shape.chart?.title));
    expect(titles.sort()).toEqual(["同页副本", "季度营收", "季度营收"].sort());
    const zip = await JSZip.loadAsync(bytes);
    expect(Object.keys(zip.files).filter(name => /^ppt\/charts\/chart\d+\.xml$/.test(name))).toHaveLength(3);
  });

  it("deletes a chart", async () => {
    const session = await deckWithChart();
    await session.apply([{ op: "delete", target: { slide: 1, name: "营收图" } }]);
    expect(await chartOf(session)).toBeUndefined();
    const reopened = await open(await session.export());
    expect(await chartOf(reopened)).toBeUndefined();
  });
});
