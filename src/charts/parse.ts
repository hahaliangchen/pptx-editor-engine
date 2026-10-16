import { getDirectChild, getDirectChildren } from "../pptx-xml";
import { ChartKind, ChartModel, LegendPosition, emptyChart } from "./model";

export interface ChartParseContext {
  /** Colour of an element that holds an a:solidFill (resolves theme colours and their transforms). */
  solidFill(node: Element): string | null;
  /** A typeface name, with theme references like "+mn-lt" resolved. */
  typeface(name: string): string;
  palette: string[];
  /** Font to use when the chart names none: the deck's body font. */
  defaultFont: { family?: string; eastAsianFamily?: string };
}

const KNOWN: Record<string, string> = {
  barChart: "柱/条形图",
  lineChart: "折线图",
  pieChart: "饼图",
  areaChart: "面积图",
  scatterChart: "散点图",
  radarChart: "雷达图",
  doughnutChart: "圆环图",
  bubbleChart: "气泡图",
  stockChart: "股价图",
  surfaceChart: "曲面图",
  bar3DChart: "三维柱形图",
  line3DChart: "三维折线图",
  pie3DChart: "三维饼图",
  area3DChart: "三维面积图",
  ofPieChart: "复合饼图",
  surface3DChart: "三维曲面图"
};

const valOf = (node: Element | null, fallback = ""): string => node?.getAttribute("val") ?? fallback;
const isOn = (node: Element | null, fallback = false): boolean => {
  if (!node) return fallback;
  const value = node.getAttribute("val");
  return value === null ? true : value === "1" || value === "true";
};

/** Text of a c:strRef / c:numRef / c:strLit / c:numLit, one entry per point. */
function readPoints(container: Element | null): string[] {
  if (!container) return [];
  const source = getDirectChild(container, "strRef", "numRef", "multiLvlStrRef");
  const literal = getDirectChild(container, "strLit", "numLit");
  const cache = source ? getDirectChild(source, "strCache", "numCache", "multiLvlStrCache") : literal;
  if (!cache) return [];
  // A multi-level category axis lists its levels innermost first.
  const level = cache.localName === "multiLvlStrCache" ? getDirectChild(cache, "lvl") : cache;
  if (!level) return [];
  const count = parseInt(valOf(getDirectChild(cache, "ptCount"), "0"), 10) || 0;
  const points = new Array<string>(count).fill("");
  for (const pt of getDirectChildren(level, "pt")) {
    const index = parseInt(pt.getAttribute("idx") || "-1", 10);
    const text = getDirectChild(pt, "v")?.textContent ?? "";
    if (index >= 0) {
      while (points.length <= index) points.push("");
      points[index] = text;
    }
  }
  return points;
}

function textOfRich(rich: Element | null): string {
  if (!rich) return "";
  return getDirectChildren(rich, "p")
    .map(paragraph => getDirectChildren(paragraph, "r").map(run => getDirectChild(run, "t")?.textContent ?? "").join(""))
    .filter(line => line.length > 0)
    .join("\n");
}

function defaultRunProps(txPr: Element | null): Element | null {
  const paragraph = txPr ? getDirectChild(txPr, "p") : null;
  const props = paragraph ? getDirectChild(paragraph, "pPr") : null;
  return props ? getDirectChild(props, "defRPr") : null;
}

/** Read a c:chartSpace into a model. Anything the engine cannot draw yet comes back marked `unsupported`. */
export function parseChartXml(xml: string, part: string, ctx: ChartParseContext): ChartModel {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const space = doc.documentElement;
  const chart = emptyChart("column");
  chart.part = part;
  chart.palette = ctx.palette.length > 0 ? [...ctx.palette] : chart.palette;
  chart.font = { family: ctx.defaultFont.family, eastAsianFamily: ctx.defaultFont.eastAsianFamily, color: "#595959" };

  const chartNode = getDirectChild(space, "chart");
  const plotArea = chartNode ? getDirectChild(chartNode, "plotArea") : null;
  if (!chartNode || !plotArea) {
    chart.unsupported = "无法读取";
    return chart;
  }

  const types = Array.from(plotArea.childNodes).filter((node): node is Element =>
    node.nodeType === 1 && /Chart$|3DChart$/.test((node as Element).localName || ""));
  if (types.length !== 1) {
    chart.unsupported = types.length > 1 ? "组合图" : "无图表类型";
    return chart;
  }
  const typeNode = types[0];
  const typeName = typeNode.localName || "";
  const kindByType: Record<string, ChartKind | undefined> = { barChart: undefined, lineChart: "line", pieChart: "pie" };
  if (!(typeName in kindByType)) {
    chart.unsupported = KNOWN[typeName] ?? typeName;
    return chart;
  }

  const grouping = valOf(getDirectChild(typeNode, "grouping"), "standard");
  if (typeName === "barChart") chart.kind = valOf(getDirectChild(typeNode, "barDir"), "col") === "bar" ? "bar" : "column";
  else chart.kind = kindByType[typeName]!;
  if (grouping === "percentStacked") {
    chart.unsupported = "百分比堆积";
    return chart;
  }
  chart.stacked = grouping === "stacked";
  const gap = getDirectChild(typeNode, "gapWidth");
  if (gap) chart.gapWidth = parseInt(valOf(gap, "150"), 10) || 150;

  // Series.
  const seriesNodes = getDirectChildren(typeNode, "ser");
  const pointColors: Array<string | undefined> = [];
  seriesNodes.forEach((ser, index) => {
    const nameNode = getDirectChild(ser, "tx");
    const nameRef = nameNode ? readPoints(nameNode)[0] ?? getDirectChild(nameNode, "v")?.textContent ?? "" : "";
    const values = readPoints(getDirectChild(ser, "val")).map(text => {
      if (text === "") return null;
      const number = Number(text);
      return Number.isFinite(number) ? number : null;
    });
    const spPr = getDirectChild(ser, "spPr");
    const line = spPr ? getDirectChild(spPr, "ln") : null;
    const color = chart.kind === "line"
      ? (line && ctx.solidFill(line)) || (spPr && ctx.solidFill(spPr)) || undefined
      : (spPr && ctx.solidFill(spPr)) || undefined;
    chart.series.push({ name: nameRef || `系列${index + 1}`, values, ...(color ? { color } : {}) });
    if (index === 0) {
      for (const point of getDirectChildren(ser, "dPt")) {
        const at = parseInt(valOf(getDirectChild(point, "idx"), "-1"), 10);
        const pointFill = getDirectChild(point, "spPr");
        const fill = pointFill ? ctx.solidFill(pointFill) : null;
        if (at >= 0 && fill) pointColors[at] = fill;
      }
    }
  });
  chart.fileColors = chart.series.map(series => series.color);
  if (pointColors.some(Boolean)) chart.pointColors = pointColors;

  const categories = seriesNodes.length > 0 ? readPoints(getDirectChild(seriesNodes[0], "cat")) : [];
  const length = Math.max(categories.length, ...chart.series.map(series => series.values.length), 0);
  chart.categories = Array.from({ length }, (_, index) => categories[index] || String(index + 1));
  for (const series of chart.series) {
    while (series.values.length < length) series.values.push(null);
  }

  // Title: explicit text, or the series name PowerPoint shows by itself for a single-series chart.
  const titleNode = getDirectChild(chartNode, "title");
  const deleted = isOn(getDirectChild(chartNode, "autoTitleDeleted"));
  if (titleNode && !deleted) {
    const rich = (() => {
      const tx = getDirectChild(titleNode, "tx");
      return tx ? getDirectChild(tx, "rich") : null;
    })();
    const text = textOfRich(rich);
    if (text) chart.title = text;
    else if (!rich && chart.series.length === 1) chart.title = chart.series[0].name;
    const props = rich ? defaultRunProps(rich) ?? (() => {
      const run = getDirectChild(getDirectChild(rich, "p") ?? rich, "r");
      return run ? getDirectChild(run, "rPr") : null;
    })() : null;
    const size = parseInt(props?.getAttribute("sz") || "0", 10);
    if (size > 0) chart.titlePt = size / 100;
  }

  // Legend.
  const legend = getDirectChild(chartNode, "legend");
  const position = valOf(legend && getDirectChild(legend, "legendPos"), "r");
  const positions: Record<string, LegendPosition> = { t: "top", b: "bottom", l: "left", r: "right", tr: "right" };
  chart.legend = legend ? positions[position] ?? "right" : "none";

  // Data labels, on the chart type or on any series.
  const labelNodes = [getDirectChild(typeNode, "dLbls"), ...seriesNodes.map(ser => getDirectChild(ser, "dLbls"))];
  chart.dataLabels = labelNodes.some(node => node && !isOn(getDirectChild(node, "delete"))
    && (isOn(getDirectChild(node, "showVal")) || isOn(getDirectChild(node, "showPercent"))));

  // Text: size, colour and font come from the chart's default run properties.
  const axisText = ["catAx", "valAx"].map(name => getDirectChild(plotArea, name)).map(axis => axis && getDirectChild(axis, "txPr"));
  const props = [defaultRunProps(getDirectChild(space, "txPr")), ...axisText.map(defaultRunProps), defaultRunProps(legend && getDirectChild(legend, "txPr"))]
    .filter((node): node is Element => !!node);
  for (const node of props) {
    const size = parseInt(node.getAttribute("sz") || "0", 10);
    if (size > 0) {
      chart.labelPt = size / 100;
      break;
    }
  }
  if (chart.labelPt === undefined) chart.labelPt = 10;
  for (const node of props) {
    const color = ctx.solidFill(node);
    if (color) {
      chart.font.color = color;
      break;
    }
  }
  for (const node of props) {
    const latin = getDirectChild(node, "latin")?.getAttribute("typeface");
    const eastAsian = getDirectChild(node, "ea")?.getAttribute("typeface");
    if (latin) chart.font.family = ctx.typeface(latin);
    if (eastAsian) chart.font.eastAsianFamily = ctx.typeface(eastAsian);
    if (latin || eastAsian) break;
  }
  if (chart.titlePt === undefined) chart.titlePt = Math.round(chart.labelPt * 1.55 * 10) / 10;

  // A filled chart area (dark slides often have one).
  const areaFill = getDirectChild(space, "spPr");
  const background = areaFill ? ctx.solidFill(areaFill) : null;
  if (background) chart.background = background;

  return chart;
}
