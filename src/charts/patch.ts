import { getDirectChild, getDirectChildren } from "../pptx-xml";
import { ChartModel, seriesColor } from "./model";
import {
  C_NS,
  categoriesXml,
  dataLabelsXml,
  legendXml,
  seriesTextXml,
  sheetRanges,
  titleXml,
  valuesXml
} from "./write";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** Parse an XML fragment written with the chart and drawing namespaces declared. */
function fragment(doc: Document, xml: string): Element {
  const wrapper = new DOMParser().parseFromString(
    `<root xmlns:c="${C_NS}" xmlns:a="${A_NS}">${xml}</root>`, "application/xml");
  return doc.importNode(wrapper.documentElement.firstElementChild!, true) as Element;
}

const plotTypeOf = (doc: Document): Element | null => {
  const plot = getDirectChild(getDirectChild(doc.documentElement, "chart") ?? doc.documentElement, "plotArea");
  if (!plot) return null;
  return Array.from(plot.childNodes).find((node): node is Element =>
    node.nodeType === 1 && /Chart$/.test((node as Element).localName || "")) ?? null;
};

/** The kind and stacking a chart part currently has, to tell a data edit from a change of type. */
export function chartShapeOf(xml: string): { kind: ChartModel["kind"] | null; stacked: boolean } {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const type = plotTypeOf(doc);
  if (!type) return { kind: null, stacked: false };
  const grouping = getDirectChild(type, "grouping")?.getAttribute("val") ?? "standard";
  const name = type.localName;
  const kind = name === "pieChart" ? "pie" : name === "lineChart" ? "line"
    : name === "barChart" ? (getDirectChild(type, "barDir")?.getAttribute("val") === "bar" ? "bar" : "column") : null;
  return { kind, stacked: grouping === "stacked" };
}

function replaceChild(parent: Element, name: string, replacement: Element, before: string[]): void {
  const existing = getDirectChild(parent, name);
  if (existing) {
    parent.replaceChild(replacement, existing);
    return;
  }
  const anchor = getDirectChildren(parent, ...before)[0] ?? null;
  parent.insertBefore(replacement, anchor);
}

function setSeriesColor(doc: Document, ser: Element, kind: ChartModel["kind"], color: string): void {
  const hex = color.replace("#", "").toUpperCase();
  const solid = fragment(doc, `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`);
  let spPr = getDirectChild(ser, "spPr");
  if (!spPr) {
    spPr = fragment(doc, "<c:spPr/>");
    ser.insertBefore(spPr, getDirectChild(ser, "tx")?.nextSibling ?? ser.firstChild);
  }
  const holder = kind === "line" ? (getDirectChild(spPr, "ln") ?? spPr.appendChild(fragment(doc, '<a:ln w="28575" cap="rnd"/>'))) as Element : spPr;
  for (const old of getDirectChildren(holder, "solidFill", "noFill", "gradFill", "pattFill")) holder.removeChild(old);
  holder.insertBefore(solid, holder.firstChild);
  if (kind === "line") {
    const marker = getDirectChild(ser, "marker");
    const markerFill = marker && getDirectChild(marker, "spPr");
    if (markerFill) {
      for (const old of getDirectChildren(markerFill, "solidFill", "noFill", "gradFill")) markerFill.removeChild(old);
      markerFill.insertBefore(fragment(doc, `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`), markerFill.firstChild);
    }
  }
}

function setTitle(doc: Document, chartNode: Element, chart: ChartModel): void {
  const existing = getDirectChild(chartNode, "title");
  const auto = getDirectChild(chartNode, "autoTitleDeleted");
  if (!chart.title) {
    if (existing) chartNode.removeChild(existing);
    if (auto) auto.setAttribute("val", "1");
    return;
  }
  if (auto) auto.setAttribute("val", "0");
  const rich = existing && getDirectChild(getDirectChild(existing, "tx") ?? existing, "rich");
  if (!existing || !rich) {
    const created = fragment(doc, titleXml(chart));
    if (existing) chartNode.replaceChild(created, existing);
    else chartNode.insertBefore(created, chartNode.firstChild);
    return;
  }
  // Keep the title's own formatting: put the text into the first run and drop the rest.
  const paragraphs = getDirectChildren(rich, "p");
  const lines = chart.title.split("\n");
  paragraphs.forEach((paragraph, index) => {
    if (index >= lines.length) {
      rich.removeChild(paragraph);
      return;
    }
    const runs = getDirectChildren(paragraph, "r");
    runs.forEach((run, runIndex) => {
      if (runIndex > 0) paragraph.removeChild(run);
    });
    const text = runs[0] && getDirectChild(runs[0], "t");
    if (text) text.textContent = lines[index];
  });
  const last = paragraphs[0];
  for (let index = paragraphs.length; index < lines.length && last; index++) {
    const copy = last.cloneNode(true) as Element;
    const run = getDirectChild(copy, "r");
    const text = run && getDirectChild(run, "t");
    if (text) text.textContent = lines[index];
    rich.appendChild(copy);
  }
}

function setLegend(doc: Document, chartNode: Element, chart: ChartModel): void {
  const existing = getDirectChild(chartNode, "legend");
  if (chart.legend === "none") {
    if (existing) chartNode.removeChild(existing);
    return;
  }
  const position = { top: "t", bottom: "b", left: "l", right: "r" }[chart.legend];
  const pos = existing && getDirectChild(existing, "legendPos");
  if (existing && pos) {
    pos.setAttribute("val", position);
    return;
  }
  const created = fragment(doc, legendXml(chart));
  if (existing) chartNode.replaceChild(created, existing);
  else chartNode.insertBefore(created, getDirectChild(chartNode, "plotVisOnly"));
}

function setDataLabels(doc: Document, typeNode: Element, chart: ChartModel): void {
  const nodes = [getDirectChild(typeNode, "dLbls"), ...getDirectChildren(typeNode, "ser").map(ser => getDirectChild(ser, "dLbls"))]
    .filter((node): node is Element => !!node);
  const pie = chart.kind === "pie";
  if (chart.dataLabels && nodes.length === 0) {
    const sers = getDirectChildren(typeNode, "ser");
    typeNode.insertBefore(fragment(doc, dataLabelsXml(chart)), sers[sers.length - 1].nextSibling);
    return;
  }
  for (const node of nodes) {
    const flag = (name: string, on: boolean) => {
      const element = getDirectChild(node, name);
      if (element) element.setAttribute("val", on ? "1" : "0");
    };
    const deleted = getDirectChild(node, "delete");
    if (deleted) {
      if (chart.dataLabels) deleted.setAttribute("val", "0");
      else continue;
    }
    flag("showVal", chart.dataLabels && !pie);
    flag("showPercent", chart.dataLabels && pie);
  }
}

const FILL_TAGS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];

function setRunProps(doc: Document, props: Element, sizePt: number, color: string): void {
  props.setAttribute("sz", String(Math.round(sizePt * 100)));
  for (const old of getDirectChildren(props, ...FILL_TAGS)) props.removeChild(old);
  const fill = fragment(doc, `<a:solidFill><a:srgbClr val="${color.replace("#", "").toUpperCase()}"/></a:solidFill>`);
  const after = getDirectChild(props, "ln");
  props.insertBefore(fill, after ? after.nextSibling : props.firstChild);
}

/** The text colour and sizes the model now says, written over what the part had, but only where they changed. */
function patchText(doc: Document, chartNode: Element, chart: ChartModel): void {
  const was = chart.fileText;
  if (!was) return;
  const labelPt = chart.labelPt ?? was.labelPt;
  const titlePt = chart.titlePt ?? was.titlePt;
  const colorChanged = chart.font.color !== was.color;
  if (!colorChanged && labelPt === was.labelPt && titlePt === was.titlePt) return;

  const space = doc.documentElement;
  const plotArea = getDirectChild(chartNode, "plotArea");
  const holders = [
    getDirectChild(space, "txPr"),
    ...["catAx", "valAx"].map(name => plotArea && getDirectChild(plotArea, name)).map(axis => axis && getDirectChild(axis, "txPr")),
    getDirectChild(getDirectChild(chartNode, "legend") ?? chartNode, "txPr"),
    ...Array.from(doc.getElementsByTagNameNS(C_NS, "dLbls")).map(list => getDirectChild(list, "txPr"))
  ].filter((node): node is Element => !!node);
  for (const holder of holders) {
    for (const props of Array.from(holder.getElementsByTagNameNS(A_NS, "defRPr"))) setRunProps(doc, props, labelPt, chart.font.color);
  }
  const title = getDirectChild(chartNode, "title");
  if (title) {
    for (const tag of ["defRPr", "rPr"]) {
      for (const props of Array.from(title.getElementsByTagNameNS(A_NS, tag))) setRunProps(doc, props, titlePt, chart.font.color);
    }
  }
}

/**
 * Write the model's data and settings into a chart part, leaving everything else
 * about it (styles, fonts, axis formats) as the file had it.
 */
export function patchChartXml(xml: string, chart: ChartModel): string {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const chartNode = getDirectChild(doc.documentElement, "chart");
  const typeNode = plotTypeOf(doc);
  if (!chartNode || !typeNode) return xml;

  let sers = getDirectChildren(typeNode, "ser");
  while (sers.length > chart.series.length) typeNode.removeChild(sers.pop()!);
  while (sers.length < chart.series.length && sers.length > 0) {
    const copy = sers[sers.length - 1].cloneNode(true) as Element;
    sers[sers.length - 1].after(copy);
    sers.push(copy);
  }

  const ranges = sheetRanges(chart);
  sers.forEach((ser, index) => {
    getDirectChild(ser, "idx")?.setAttribute("val", String(index));
    getDirectChild(ser, "order")?.setAttribute("val", String(index));
    replaceChild(ser, "tx", fragment(doc, seriesTextXml(chart.series[index].name, ranges.name(index))), ["spPr", "invertIfNegative", "marker", "dPt", "dLbls", "cat", "val"]);
    replaceChild(ser, "cat", fragment(doc, categoriesXml(chart)), ["val", "smooth"]);
    replaceChild(ser, "val", fragment(doc, valuesXml(chart, index)), ["smooth", "shape"]);
    const known = chart.fileColors && index < chart.fileColors.length;
    const wanted = chart.series[index].color;
    if (known ? wanted && wanted !== chart.fileColors![index] : true) setSeriesColor(doc, ser, chart.kind, seriesColor(chart, index));
    if (chart.kind === "pie") {
      for (const point of getDirectChildren(ser, "dPt")) {
        if (parseInt(getDirectChild(point, "idx")?.getAttribute("val") ?? "0", 10) >= chart.categories.length) ser.removeChild(point);
      }
    }
  });

  setTitle(doc, chartNode, chart);
  patchText(doc, chartNode, chart);
  setLegend(doc, chartNode, chart);
  setDataLabels(doc, typeNode, chart);
  const gap = getDirectChild(typeNode, "gapWidth");
  if (gap && (chart.kind === "column" || chart.kind === "bar")) gap.setAttribute("val", String(chart.gapWidth));

  const out = new XMLSerializer().serializeToString(doc);
  return out.startsWith("<?xml") ? out : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
}
