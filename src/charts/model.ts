import { ElementSource, Rect } from "../pptx-virtual-dom";

export type ChartKind = "column" | "bar" | "line" | "pie";
export type LegendPosition = "top" | "bottom" | "left" | "right" | "none";

export interface ChartSeries {
  name: string;
  /** One value per category; null leaves a gap. */
  values: Array<number | null>;
  /** "#rrggbb"; without it the series takes the next colour of the deck's theme. */
  color?: string;
}

/** Everything the engine knows about a chart; layout and the chart part both derive from it. */
export interface ChartModel {
  kind: ChartKind;
  /** Column, bar and line only: series are stacked instead of side by side. */
  stacked: boolean;
  title?: string;
  categories: string[];
  series: ChartSeries[];
  legend: LegendPosition;
  /** Show each value (a pie shows percentages). */
  dataLabels: boolean;
  /** The deck's six accent colours, used for series that name no colour. */
  palette: string[];
  font: { family?: string; eastAsianFamily?: string; color: string };
  /** Space between column/bar groups, as a percentage of one bar's width. */
  gapWidth: number;
  /** Size of axis, legend and data label text in points (12 when absent). */
  labelPt?: number;
  /** Size of the title in points (18 when absent). */
  titlePt?: number;
  /** Fill behind the whole chart; transparent when absent. */
  background?: string;
  /** Colour each series had in the file, to tell later colour edits from untouched series. */
  fileColors?: Array<string | undefined>;
  /** Pie only: colour of each slice where the file names one. */
  pointColors?: Array<string | undefined>;
  /** Chart part of an opened deck, e.g. "ppt/charts/chart1.xml"; absent for charts made here. */
  part?: string;
  /** Set when the chart is of a kind the engine cannot draw yet; only a placeholder is shown. */
  unsupported?: string;
}

export interface ChartElement {
  type: "chart";
  id: string;
  rect: Rect;
  source?: ElementSource;
  chart: ChartModel;
  /** Charts do not rotate; present so every element can be asked. */
  rotation?: number;
  /** Pixels per point on this slide; chart text is sized in points. */
  pxPerPt: number;
}

export const CHART_KINDS: ChartKind[] = ["column", "bar", "line", "pie"];
export const LEGEND_POSITIONS: LegendPosition[] = ["top", "bottom", "left", "right", "none"];

/** Office's default accent colours, used when a deck has no theme to take them from. */
export const DEFAULT_PALETTE = ["#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"];

const hex2 = (value: number) => Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0");

/** Blend a "#rrggbb" colour towards black (negative) or white (positive); amount in -1..1. */
export function shiftColor(color: string, amount: number): string {
  const channel = (offset: number) => parseInt(color.slice(offset, offset + 2), 16);
  const target = amount < 0 ? 0 : 255;
  const mix = Math.abs(amount);
  return `#${[1, 3, 5].map(offset => hex2(channel(offset) + (target - channel(offset)) * mix)).join("")}`;
}

/** Colour of series `index`: its own, else the theme colour, darker or lighter on each lap as PowerPoint does. */
export function seriesColor(chart: ChartModel, index: number): string {
  const own = chart.series[index]?.color;
  if (own) return own;
  const palette = chart.palette.length > 0 ? chart.palette : DEFAULT_PALETTE;
  const base = palette[index % palette.length];
  const lap = Math.floor(index / palette.length);
  if (lap === 0) return base;
  return shiftColor(base, lap % 2 === 1 ? -0.25 * Math.ceil(lap / 2) : 0.25 * (lap / 2));
}

/** Colour of one pie slice (slices are the categories, coloured like series in the other kinds). */
export const sliceColor = (chart: ChartModel, index: number): string =>
  chart.pointColors?.[index] ?? seriesColor({ ...chart, series: [] }, index);

/** "12", "3.5", "0.25": a number as the General format shows it. */
export function formatNumber(value: number): string {
  if (!Number.isFinite(value)) return "";
  const rounded = Math.round(value * 1e6) / 1e6;
  return String(rounded);
}

export function emptyChart(kind: ChartKind): ChartModel {
  return {
    kind,
    stacked: false,
    categories: [],
    series: [],
    legend: kind === "pie" ? "right" : "bottom",
    dataLabels: false,
    palette: [...DEFAULT_PALETTE],
    font: { color: "#595959" },
    gapWidth: 150
  };
}

/** A copy that shares nothing with the original, for undo snapshots. */
export const cloneChart = (chart: ChartModel): ChartModel => JSON.parse(JSON.stringify(chart));

/** A chart this session added to the package: where its parts go and the model they are written from. */
export interface ChartAdd {
  /** e.g. "ppt/charts/chart7.xml" */
  part: string;
  /** e.g. "ppt/embeddings/Microsoft_Excel_Sheet7.xlsx" */
  workbook: string;
  /** Shared with the chart element, so later edits reach the export. */
  model: ChartModel;
}
