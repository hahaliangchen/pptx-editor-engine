import JSZip from "jszip";
import { ChartModel, formatNumber, seriesColor, sliceColor } from "./model";

export const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const CHART_REL_TYPE = `${R_NS}/chart`;
export const PACKAGE_REL_TYPE = `${R_NS}/package`;
export const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
export const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export const escapeXml = (text: string): string =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const hex = (color: string) => color.replace("#", "").toUpperCase();
const fill = (color: string) => `<a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill>`;

/** "A", "B", ... "Z", "AA": the spreadsheet column for a 0-based index. */
export function columnName(index: number): string {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

/** Where the data sits in the embedded workbook. */
export const sheetRanges = (chart: ChartModel) => {
  const last = chart.categories.length + 1;
  return {
    categories: `Sheet1!$A$2:$A$${last}`,
    name: (series: number) => `Sheet1!$${columnName(series + 1)}$1`,
    values: (series: number) => `Sheet1!$${columnName(series + 1)}$2:$${columnName(series + 1)}$${last}`
  };
};

const runProps = (chart: ChartModel, size: number, bold = false): string => {
  const family = chart.font.family ? `<a:latin typeface="${escapeXml(chart.font.family)}"/>` : "";
  const eastAsian = chart.font.eastAsianFamily ? `<a:ea typeface="${escapeXml(chart.font.eastAsianFamily)}"/>` : "";
  return ` sz="${Math.round(size * 100)}" b="${bold ? 1 : 0}">${fill(chart.font.color)}${family}${eastAsian}`;
};

export const textProps = (chart: ChartModel, size: number): string =>
  `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${runProps(chart, size)}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;

export const titleXml = (chart: ChartModel): string => {
  const size = chart.titlePt ?? 18;
  const lines = (chart.title ?? "").split("\n");
  const paragraphs = lines.map(line =>
    `<a:p><a:pPr><a:defRPr${runProps(chart, size)}</a:defRPr></a:pPr>`
    + `<a:r><a:rPr lang="zh-CN"${runProps(chart, size)}</a:rPr><a:t>${escapeXml(line)}</a:t></a:r></a:p>`).join("");
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>${paragraphs}</c:rich></c:tx><c:overlay val="0"/></c:title>`;
};

export const legendXml = (chart: ChartModel): string => {
  const position = { top: "t", bottom: "b", left: "l", right: "r", none: "r" }[chart.legend];
  return `<c:legend><c:legendPos val="${position}"/><c:overlay val="0"/>${textProps(chart, chart.labelPt ?? 12)}</c:legend>`;
};

const strCache = (values: string[]) =>
  `<c:ptCount val="${values.length}"/>` + values.map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`).join("");

const numCache = (values: Array<number | null>) =>
  `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>`
  + values.map((value, index) => value === null ? "" : `<c:pt idx="${index}"><c:v>${formatNumber(value)}</c:v></c:pt>`).join("");

export const seriesTextXml = (name: string, ref: string): string =>
  `<c:tx><c:strRef><c:f>${ref}</c:f><c:strCache>${strCache([name])}</c:strCache></c:strRef></c:tx>`;

export const categoriesXml = (chart: ChartModel): string =>
  `<c:cat><c:strRef><c:f>${sheetRanges(chart).categories}</c:f><c:strCache>${strCache(chart.categories)}</c:strCache></c:strRef></c:cat>`;

export const valuesXml = (chart: ChartModel, index: number): string =>
  `<c:val><c:numRef><c:f>${sheetRanges(chart).values(index)}</c:f><c:numCache>${numCache(chart.series[index].values)}</c:numCache></c:numRef></c:val>`;

/** Data label settings for the chart type, as the schema orders them. */
export const dataLabelsXml = (chart: ChartModel): string => {
  const none = `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
  if (!chart.dataLabels) return none;
  const position = chart.kind === "pie" ? "outEnd" : chart.kind === "line" ? "t" : chart.stacked ? "ctr" : "outEnd";
  const pie = chart.kind === "pie";
  return `<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${textProps(chart, chart.labelPt ?? 12)}`
    + `<c:dLblPos val="${position}"/><c:showLegendKey val="0"/><c:showVal val="${pie ? 0 : 1}"/><c:showCatName val="0"/>`
    + `<c:showSerName val="0"/><c:showPercent val="${pie ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`;
};

function seriesXml(chart: ChartModel, index: number): string {
  const series = chart.series[index];
  const color = seriesColor(chart, index);
  const head = `<c:idx val="${index}"/><c:order val="${index}"/>${seriesTextXml(series.name, sheetRanges(chart).name(index))}`;
  const data = `${categoriesXml(chart)}${valuesXml(chart, index)}`;
  if (chart.kind === "line") {
    return `<c:ser>${head}<c:spPr><a:ln w="28575" cap="rnd">${fill(color)}<a:round/></a:ln></c:spPr>`
      + `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${fill(color)}<a:ln w="9525">${fill(color)}</a:ln></c:spPr></c:marker>`
      + `${data}<c:smooth val="0"/></c:ser>`;
  }
  if (chart.kind === "pie") {
    const slices = chart.categories.map((_, at) =>
      `<c:dPt><c:idx val="${at}"/><c:bubble3D val="0"/><c:spPr>${fill(sliceColor(chart, at))}<a:ln w="19050">${fill("#ffffff")}</a:ln></c:spPr></c:dPt>`).join("");
    return `<c:ser>${head}${slices}${data}</c:ser>`;
  }
  return `<c:ser>${head}<c:spPr>${fill(color)}</c:spPr><c:invertIfNegative val="0"/>${data}</c:ser>`;
}

const AXIS_IDS = [111111111, 222222222];

function axesXml(chart: ChartModel): string {
  const horizontal = chart.kind === "bar";
  const labels = textProps(chart, chart.labelPt ?? 12);
  const lineFill = (color: string) => `<c:spPr><a:noFill/><a:ln w="9525">${fill(color)}</a:ln></c:spPr>`;
  const category = `<c:catAx><c:axId val="${AXIS_IDS[0]}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>`
    + `<c:axPos val="${horizontal ? "l" : "b"}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>`
    + `<c:tickLblPos val="nextTo"/>${lineFill("#bfbfbf")}${labels}<c:crossAx val="${AXIS_IDS[1]}"/><c:crosses val="autoZero"/>`
    + `<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const value = `<c:valAx><c:axId val="${AXIS_IDS[1]}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>`
    + `<c:axPos val="${horizontal ? "b" : "l"}"/><c:majorGridlines><c:spPr><a:ln w="9525">${fill("#d9d9d9")}</a:ln></c:spPr></c:majorGridlines>`
    + `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>`
    + `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labels}<c:crossAx val="${AXIS_IDS[0]}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  return category + value;
}

function plotXml(chart: ChartModel): string {
  const series = chart.series.map((_, index) => seriesXml(chart, index)).join("");
  const axes = `<c:axId val="${AXIS_IDS[0]}"/><c:axId val="${AXIS_IDS[1]}"/>`;
  if (chart.kind === "pie") {
    return `<c:pieChart><c:varyColors val="1"/>${series}${dataLabelsXml(chart)}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  if (chart.kind === "line") {
    return `<c:lineChart><c:grouping val="${chart.stacked ? "stacked" : "standard"}"/><c:varyColors val="0"/>${series}${dataLabelsXml(chart)}`
      + `<c:marker val="1"/>${axes}</c:lineChart>${axesXml(chart)}`;
  }
  return `<c:barChart><c:barDir val="${chart.kind === "bar" ? "bar" : "col"}"/><c:grouping val="${chart.stacked ? "stacked" : "clustered"}"/>`
    + `<c:varyColors val="0"/>${series}${dataLabelsXml(chart)}<c:gapWidth val="${chart.gapWidth}"/>${chart.stacked ? '<c:overlap val="100"/>' : ""}${axes}</c:barChart>${axesXml(chart)}`;
}

/** A complete chart part for `chart`, with its data linked to the embedded workbook. */
export function chartXml(chart: ChartModel): string {
  const title = chart.title ? titleXml(chart) : "";
  const legend = chart.legend === "none" ? "" : legendXml(chart);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
    + `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"><c:date1904 val="0"/><c:roundedCorners val="0"/>`
    + `<c:chart>${title}<c:autoTitleDeleted val="${chart.title ? 0 : 1}"/><c:plotArea><c:layout/>${plotXml(chart)}<c:spPr><a:noFill/></c:spPr></c:plotArea>`
    + `${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
    + `<c:spPr>${chart.background ? fill(chart.background) : "<a:noFill/>"}<a:ln><a:noFill/></a:ln></c:spPr>`
    + `${textProps(chart, chart.labelPt ?? 12)}<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
}

export const chartRelsXml = (workbookTarget: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
  + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
  + `<Relationship Id="rId1" Type="${PACKAGE_REL_TYPE}" Target="${workbookTarget}"/></Relationships>`;

/** The slide-side shape that shows a chart part. */
export function graphicFrameXml(id: string, name: string, relId: string, box: { x: number; y: number; cx: number; cy: number }): string {
  return `<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`
    + `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>`
    + `<p:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></p:xfrm>`
    + `<a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" r:id="${relId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

/** The data as a minimal workbook: categories down column A, one series per column after it. */
export async function workbookBytes(chart: ChartModel): Promise<Uint8Array> {
  const zip = new JSZip();
  const SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const cell = (ref: string, value: string | number | null) => value === null || value === ""
    ? ""
    : typeof value === "number"
      ? `<c r="${ref}"><v>${formatNumber(value)}</v></c>`
      : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
  const rows: string[] = [];
  rows.push(`<row r="1">${chart.series.map((series, i) => cell(`${columnName(i + 1)}1`, series.name)).join("")}</row>`);
  chart.categories.forEach((category, rowIndex) => {
    const row = rowIndex + 2;
    rows.push(`<row r="${row}">${cell(`A${row}`, category)}${chart.series.map((series, i) => cell(`${columnName(i + 1)}${row}`, series.values[rowIndex] ?? null)).join("")}</row>`);
  });
  zip.file("[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
    + `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`
    + `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
  zip.file("_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file("xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${SS}" xmlns:r="${R_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rId1" Type="${R_NS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file("xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${SS}"><sheetData>${rows.join("")}</sheetData></worksheet>`);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
