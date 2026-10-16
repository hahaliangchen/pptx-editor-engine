import { describe, expect, it } from "vitest";
import { expandChart, niceAxis, textWidth } from "../src/charts/layout";
import { ChartElement, ChartModel, emptyChart, seriesColor, shiftColor } from "../src/charts/model";
import { parseChartXml } from "../src/charts/parse";
import { patchChartXml, chartShapeOf } from "../src/charts/patch";
import { chartXml, columnName } from "../src/charts/write";
import { ShapeElement, TextElement } from "../src/pptx-virtual-dom";

const context = {
  solidFill: (node: Element) => {
    const clr = node.getElementsByTagNameNS("*", "srgbClr")[0];
    return clr ? `#${clr.getAttribute("val")!.toLowerCase()}` : null;
  },
  typeface: (name: string) => name,
  palette: ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"],
  defaultFont: { family: "Arial" }
};

function model(over: Partial<ChartModel> = {}): ChartModel {
  return {
    ...emptyChart("column"),
    categories: ["一季度", "二季度", "三季度"],
    series: [{ name: "华东", values: [10, 20, 30] }, { name: "华北", values: [5, 15, 25], color: "#c00000" }],
    ...over
  };
}

const element = (chart: ChartModel, w = 800, h = 450): ChartElement => ({
  type: "chart", id: "c", rect: { x: 100, y: 50, w, h }, chart, pxPerPt: 2
});

const shapes = (elements: ReturnType<typeof expandChart>) => elements.filter((e): e is ShapeElement => e.type === "shape");
const texts = (elements: ReturnType<typeof expandChart>) => elements.filter((e): e is TextElement => e.type === "text").map(e => e.content);

describe("axis scale", () => {
  it("starts at zero and ends on a round number above the data", () => {
    const axis = niceAxis(0, 180, 6);
    expect(axis.min).toBe(0);
    expect(axis.max).toBeGreaterThanOrEqual(180);
    expect(axis.max % axis.step).toBeCloseTo(0);
    expect(axis.step).toBe(50);
  });

  it("keeps the number of steps within what was asked", () => {
    for (const high of [7, 9.5, 38, 180, 1234, 98000]) {
      const axis = niceAxis(0, high, 5);
      expect((axis.max - axis.min) / axis.step).toBeLessThanOrEqual(8);
    }
  });

  it("makes room for negative values", () => {
    const axis = niceAxis(-30, 50, 6);
    expect(axis.min).toBeLessThan(-30 + 1);
    expect(axis.min % axis.step).toBeCloseTo(0);
  });
});

describe("colours and measuring", () => {
  it("uses a series' own colour, else the theme's in order", () => {
    const chart = model();
    expect(seriesColor(chart, 0)).toBe("#4472c4");
    expect(seriesColor(chart, 1)).toBe("#c00000");
  });

  it("darkens or lightens the theme colours on later laps, never repeating one", () => {
    const chart = { ...model(), series: Array.from({ length: 12 }, (_, i) => ({ name: `s${i}`, values: [1, 2, 3] })) };
    const colors = chart.series.map((_, i) => seriesColor(chart, i));
    expect(new Set(colors).size).toBe(12);
    expect(shiftColor("#808080", -0.5)).toBe("#404040");
  });

  it("measures wide characters as full width", () => {
    expect(textWidth("季度", 10)).toBe(20);
    expect(textWidth("ab", 10)).toBeLessThan(textWidth("季度", 10));
  });

  it("names spreadsheet columns", () => {
    expect([0, 1, 25, 26, 27].map(columnName)).toEqual(["A", "B", "Z", "AA", "AB"]);
  });
});

describe("layout", () => {
  it("draws one bar per value, in the series colour, inside the chart", () => {
    const chart = model();
    const out = expandChart(element(chart));
    const colored = (color: string) => shapes(out).filter(s => s.shapeType === "rect" && s.fill === color && s.rect.w > 12 * 2 && s.rect.h > 0);
    expect(colored("#4472c4").length).toBeGreaterThanOrEqual(3);
    // legend swatches are small squares; bars are taller than wide or at least not tiny
    const bars = shapes(out).filter(s => s.shapeType === "rect" && (s.fill === "#4472c4" || s.fill === "#c00000") && s.rect.w > 20);
    expect(bars).toHaveLength(6);
    for (const bar of bars) {
      expect(bar.rect.x).toBeGreaterThanOrEqual(100);
      expect(bar.rect.x + bar.rect.w).toBeLessThanOrEqual(900);
      expect(bar.rect.y).toBeGreaterThanOrEqual(50);
      expect(bar.rect.y + bar.rect.h).toBeLessThanOrEqual(500);
    }
  });

  it("makes taller bars for bigger values", () => {
    const out = expandChart(element(model()));
    const first = shapes(out).filter(s => s.fill === "#4472c4" && s.rect.w > 20).map(s => s.rect.h);
    expect(first[0]).toBeLessThan(first[1]);
    expect(first[1]).toBeLessThan(first[2]);
    expect(first[2] / first[0]).toBeCloseTo(3, 0);
  });

  it("puts category labels, a title, and legend entries on the page", () => {
    const out = expandChart(element(model({ title: "季度营收" })));
    const labels = texts(out);
    expect(labels).toEqual(expect.arrayContaining(["季度营收", "一季度", "二季度", "三季度", "华东", "华北", "0"]));
  });

  it("shows values only when asked", () => {
    const without = texts(expandChart(element(model()))).filter(text => text === "30");
    const withLabels = texts(expandChart(element(model({ dataLabels: true })))).filter(text => text === "30");
    expect(withLabels.length).toBeGreaterThan(without.length);
  });

  it("stacks series on top of each other", () => {
    const chart = model({ stacked: true });
    const out = expandChart(element(chart));
    const bars = shapes(out).filter(s => (s.fill === "#4472c4" || s.fill === "#c00000") && s.rect.w > 20);
    const a = bars.filter(s => s.fill === "#4472c4")[2];
    const b = bars.filter(s => s.fill === "#c00000")[2];
    expect(b.rect.y + b.rect.h).toBeCloseTo(a.rect.y, 0);
  });

  it("lays bars out sideways for the bar kind, first category at the bottom", () => {
    const out = expandChart(element(model({ kind: "bar", series: [{ name: "S", values: [10, 20, 30] }] })));
    const bars = shapes(out).filter(s => s.fill === "#4472c4" && s.rect.h > 6);
    expect(bars[0].rect.w).toBeLessThan(bars[2].rect.w);
    expect(bars[0].rect.y).toBeGreaterThan(bars[2].rect.y);
  });

  it("draws a line chart as paths with a marker per point", () => {
    const out = expandChart(element(model({ kind: "line" })));
    const paths = shapes(out).filter(s => s.customGeometry);
    expect(paths).toHaveLength(2);
    const markers = shapes(out).filter(s => s.shapeType === "ellipse");
    // six data markers plus two legend markers
    expect(markers).toHaveLength(8);
  });

  it("breaks a line at a missing value", () => {
    const out = expandChart(element(model({ kind: "line", series: [{ name: "S", values: [1, null, 3, 4] }], categories: ["a", "b", "c", "d"] })));
    const path = shapes(out).find(s => s.customGeometry)!;
    const commands = path.customGeometry!.paths[0].commands.map(c => c.type);
    expect(commands.filter(type => type === "moveTo")).toHaveLength(1);
    expect(commands.filter(type => type === "lineTo")).toHaveLength(1);
  });

  it("draws one slice per value whose angles add up to a full circle", () => {
    const chart = model({ kind: "pie", categories: ["a", "b", "c"], series: [{ name: "S", values: [50, 30, 20] }], dataLabels: true });
    const out = expandChart(element(chart));
    const slices = shapes(out).filter(s => s.customGeometry);
    expect(slices).toHaveLength(3);
    expect(texts(out)).toEqual(expect.arrayContaining(["50%", "30%", "20%"]));
  });

  it("draws a single-value pie as a whole disc", () => {
    const chart = model({ kind: "pie", categories: ["a"], series: [{ name: "S", values: [7] }] });
    const slices = shapes(expandChart(element(chart))).filter(s => s.customGeometry);
    expect(slices).toHaveLength(1);
    expect(slices[0].rect.w).toBeCloseTo(slices[0].rect.h, 0);
  });

  it("moves the legend to the side it is given", () => {
    const at = (legend: ChartModel["legend"]) =>
      expandChart(element(model({ legend }))).find((e): e is TextElement => e.type === "text" && e.content === "华东")!.rect;
    expect(at("bottom").y).toBeGreaterThan(at("top").y);
    expect(at("right").x).toBeGreaterThan(at("left").x);
    expect(texts(expandChart(element(model({ legend: "none" }))))).not.toContain("华东");
  });

  it("shows a placeholder for a kind it cannot draw", () => {
    const out = expandChart(element(model({ unsupported: "散点图" })));
    expect(texts(out).join("")).toContain("散点图");
    expect(shapes(out)).toHaveLength(1);
  });

  it("copes with an empty chart and a tiny frame", () => {
    expect(() => expandChart(element(model({ categories: [], series: [] })))).not.toThrow();
    expect(() => expandChart(element(model(), 30, 20))).not.toThrow();
  });

  it("scales with the size of the frame", () => {
    const small = shapes(expandChart(element(model(), 400, 225))).filter(s => s.fill === "#4472c4" && s.rect.w > 18);
    const large = shapes(expandChart(element(model(), 1600, 900))).filter(s => s.fill === "#4472c4" && s.rect.w > 18);
    expect(large[0].rect.w).toBeGreaterThan(small[0].rect.w * 2);
  });
});

const PART = (inner: string, extra = "") =>
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${extra}<c:chart>${inner}</c:chart></c:chartSpace>`;
const ser = (index: number, name: string, values: number[], color: string) =>
  `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
  + `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`
  + `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat>`
  + `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:val></c:ser>`;

describe("reading a chart part", () => {
  const bar = PART(
    `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>销量</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val="0"/>`
    + `<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, "甲", [1, 2, 3], "FF0000")}${ser(1, "乙", [4, 5, 6], "00FF00")}`
    + `<c:dLbls><c:showVal val="1"/></c:dLbls><c:gapWidth val="80"/></c:barChart></c:plotArea><c:legend><c:legendPos val="t"/></c:legend>`,
    `<c:txPr><a:p><a:pPr><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr>`
  );

  it("reads kind, data, title, legend and labels", () => {
    const chart = parseChartXml(bar, "ppt/charts/chart1.xml", context);
    expect(chart.kind).toBe("column");
    expect(chart.unsupported).toBeUndefined();
    expect(chart.title).toBe("销量");
    expect(chart.categories).toEqual(["a", "b", "c"]);
    expect(chart.series.map(s => [s.name, s.values, s.color])).toEqual([
      ["甲", [1, 2, 3], "#ff0000"],
      ["乙", [4, 5, 6], "#00ff00"]
    ]);
    expect(chart.legend).toBe("top");
    expect(chart.dataLabels).toBe(true);
    expect(chart.gapWidth).toBe(80);
    expect(chart.part).toBe("ppt/charts/chart1.xml");
    expect(chart.fileColors).toEqual(["#ff0000", "#00ff00"]);
    expect(chart.labelPt).toBe(14);
    expect(chart.font.color).toBe("#333333");
    expect(chart.font.family).toBe("Calibri");
  });

  it("tells horizontal bars from columns, and stacking from side by side", () => {
    const sideways = parseChartXml(bar.replace('barDir val="col"', 'barDir val="bar"').replace('"clustered"', '"stacked"'), "p", context);
    expect([sideways.kind, sideways.stacked]).toEqual(["bar", true]);
  });

  it("marks kinds it cannot draw rather than guessing", () => {
    const combo = PART(`<c:plotArea><c:barChart><c:barDir val="col"/></c:barChart><c:lineChart/></c:plotArea>`);
    expect(parseChartXml(combo, "p", context).unsupported).toBe("组合图");
    const scatter = PART(`<c:plotArea><c:scatterChart/></c:plotArea>`);
    expect(parseChartXml(scatter, "p", context).unsupported).toBe("散点图");
    const percent = PART(`<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/></c:barChart></c:plotArea>`);
    expect(parseChartXml(percent, "p", context).unsupported).toBe("百分比堆积");
    expect(parseChartXml("<nonsense/>", "p", context).unsupported).toBe("无法读取");
  });

  it("reads per-slice colours of a pie", () => {
    const pie = PART(`<c:plotArea><c:pieChart>${ser(0, "份额", [1, 2, 3], "000000").replace("<c:cat>", `<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></c:spPr></c:dPt><c:cat>`)}</c:pieChart></c:plotArea>`);
    const chart = parseChartXml(pie, "p", context);
    expect(chart.kind).toBe("pie");
    expect(chart.pointColors?.[1]).toBe("#abcdef");
  });

  it("falls back to the deck's font and the 10pt default", () => {
    const chart = parseChartXml(PART(`<c:plotArea><c:lineChart>${ser(0, "s", [1, 2, 3], "000000")}</c:lineChart></c:plotArea>`), "p", context);
    expect(chart.font.family).toBe("Arial");
    expect(chart.labelPt).toBe(10);
  });
});

describe("writing a chart part", () => {
  const fields = (chart: ChartModel) => ({
    kind: chart.kind, stacked: chart.stacked, title: chart.title, categories: chart.categories,
    series: chart.series.map((s, i) => ({ name: s.name, values: s.values, color: seriesColor(chart, i) })),
    legend: chart.legend, dataLabels: chart.dataLabels
  });
  const cases: Array<[string, ChartModel]> = [
    ["column", model({ title: "T", dataLabels: true })],
    ["stacked column", model({ stacked: true, legend: "right" })],
    ["bar", model({ kind: "bar", legend: "none" })],
    ["line", model({ kind: "line", legend: "top", dataLabels: true })],
    ["stacked line", model({ kind: "line", stacked: true })],
    ["pie", model({ kind: "pie", categories: ["a", "b"], series: [{ name: "S", values: [3, 4] }], dataLabels: true, legend: "left" })],
    ["gaps", model({ series: [{ name: "S", values: [1, null, 3] }, { name: "T", values: [null, 2, 3] }] })],
    ["special characters", model({ title: "A & B <C>", categories: ["x\"y", "<z>", "&"], series: [{ name: "n&m", values: [1, 2, 3] }] })]
  ];

  for (const [name, chart] of cases) {
    it(`reads back what it wrote: ${name}`, () => {
      const xml = chartXml(chart);
      const back = parseChartXml(xml, "ppt/charts/chart9.xml", { ...context, palette: chart.palette });
      expect(back.unsupported).toBeUndefined();
      expect(fields(back)).toEqual(fields(chart));
    });
  }

  it("produces well-formed XML", () => {
    for (const [, chart] of cases) {
      const doc = new DOMParser().parseFromString(chartXml(chart), "application/xml");
      expect(doc.getElementsByTagName("parsererror")).toHaveLength(0);
    }
  });

  it("links every series to a range of the embedded workbook", () => {
    const xml = chartXml(model());
    expect(xml).toContain("<c:f>Sheet1!$A$2:$A$4</c:f>");
    expect(xml).toContain("<c:f>Sheet1!$B$1</c:f>");
    expect(xml).toContain("<c:f>Sheet1!$C$2:$C$4</c:f>");
    expect(xml).toContain('<c:externalData r:id="rId1">');
  });

  it("names the shape it wrote", () => {
    expect(chartShapeOf(chartXml(model({ kind: "bar", stacked: true })))).toEqual({ kind: "bar", stacked: true });
    expect(chartShapeOf(chartXml(model({ kind: "pie", categories: ["a"], series: [{ name: "S", values: [1] }] })))).toEqual({ kind: "pie", stacked: false });
  });
});

describe("patching a chart part", () => {
  const parse = (xml: string) => parseChartXml(xml, "p", context);

  it("replaces data and keeps what the file had", () => {
    const original = chartXml(model({ title: "旧标题", dataLabels: true })).replace('<c:gapWidth val="150"/>', '<c:gapWidth val="42"/>');
    const edited = model({ title: "新标题", dataLabels: true, gapWidth: 42, categories: ["x", "y"], series: [{ name: "唯一", values: [7, 8] }] });
    const back = parse(patchChartXml(original, edited));
    expect(back.title).toBe("新标题");
    expect(back.categories).toEqual(["x", "y"]);
    expect(back.series.map(s => [s.name, s.values])).toEqual([["唯一", [7, 8]]]);
    expect(back.gapWidth).toBe(42);
  });

  it("adds series by copying the last one, with their own colours", () => {
    const original = chartXml(model());
    const edited = model({ series: [...model().series, { name: "第三", values: [1, 2, 3] }, { name: "第四", values: [3, 2, 1], color: "#123456" }] });
    const back = parse(patchChartXml(original, edited));
    expect(back.series.map(s => s.name)).toEqual(["华东", "华北", "第三", "第四"]);
    expect(back.series[3].color).toBe("#123456");
    expect(back.series[2].color).toBe("#a5a5a5");
    expect(patchChartXml(original, edited)).toContain('<c:idx val="3"/>');
  });

  it("removes the title, the legend and the labels", () => {
    const original = chartXml(model({ title: "T", dataLabels: true }));
    const back = parse(patchChartXml(original, model({ title: undefined, legend: "none", dataLabels: false })));
    expect(back.title).toBeUndefined();
    expect(back.legend).toBe("none");
    expect(back.dataLabels).toBe(false);
  });

  it("adds a title, a legend and labels that were missing", () => {
    const original = chartXml(model({ legend: "none" }));
    const back = parse(patchChartXml(original, model({ title: "新增", legend: "left", dataLabels: true })));
    expect(back.title).toBe("新增");
    expect(back.legend).toBe("left");
    expect(back.dataLabels).toBe(true);
  });

  it("leaves series colours alone unless they were changed", () => {
    const original = chartXml(model()).replace(/C00000/gi, "112233");
    const fromFile = parse(original);
    const back = parse(patchChartXml(original, { ...fromFile, title: "x" }));
    expect(back.series[1].color).toBe("#112233");
  });
});
