import { ChartKind, ChartModel, LegendPosition, cloneChart, emptyChart } from "../charts/model";
import { ApiError } from "./errors";

/** The chart fields `addChart` and `setChart` share. */
export interface ChartFields {
  type?: ChartKind;
  categories?: string[];
  series?: Array<{ name: string; values: Array<number | null>; color?: string }>;
  title?: string | null;
  legend?: LegendPosition;
  dataLabels?: boolean;
  stacked?: boolean;
}

const KIND_NAMES: Record<ChartKind, string> = { column: "柱形图", bar: "条形图", line: "折线图", pie: "饼图" };

/**
 * `base` with the given fields applied, or an error naming the field that does
 * not fit. The base is never changed.
 */
export function applyChartFields(base: ChartModel, fields: ChartFields, at: (field: string) => string): ChartModel {
  const chart = cloneChart(base);
  if (fields.type !== undefined) {
    chart.kind = fields.type;
    if (fields.type === "pie") chart.stacked = false;
  }
  if (fields.categories !== undefined) chart.categories = fields.categories.map(String);
  if (fields.series !== undefined) {
    chart.series = fields.series.map(series => ({
      name: series.name,
      values: [...series.values],
      ...(series.color ? { color: series.color.toLowerCase() } : {})
    }));
    // New data replaces the file's: its colours no longer describe these series.
    chart.fileColors = undefined;
  }
  if (fields.title !== undefined) chart.title = fields.title || undefined;
  if (fields.legend !== undefined) chart.legend = fields.legend;
  if (fields.dataLabels !== undefined) chart.dataLabels = fields.dataLabels;
  if (fields.stacked !== undefined) chart.stacked = fields.stacked;

  const kindName = KIND_NAMES[chart.kind];
  if (chart.categories.length === 0) throw new ApiError("INVALID_SCHEMA", "categories 不能为空", at("categories"));
  if (chart.series.length === 0) throw new ApiError("INVALID_SCHEMA", "series 不能为空", at("series"));
  chart.series.forEach((series, index) => {
    if (series.values.length !== chart.categories.length) {
      throw new ApiError(
        "INVALID_SCHEMA",
        `系列 "${series.name}" 有 ${series.values.length} 个数值，但有 ${chart.categories.length} 个分类，两者必须一样多`,
        at(`series[${index}].values`)
      );
    }
    series.values.forEach((value, at2) => {
      if (value !== null && !Number.isFinite(value)) {
        throw new ApiError("INVALID_SCHEMA", `数值必须是有限的数字或 null`, at(`series[${index}].values[${at2}]`));
      }
    });
  });
  if (chart.kind === "pie") {
    if (chart.series.length !== 1) {
      throw new ApiError("INVALID_SCHEMA", `${kindName}只能有一个系列，现在有 ${chart.series.length} 个`, at("series"));
    }
    if (chart.series[0].values.some(value => value !== null && value < 0)) {
      throw new ApiError("INVALID_SCHEMA", `${kindName}的数值不能为负`, at("series[0].values"));
    }
    if (chart.stacked) throw new ApiError("INVALID_SCHEMA", `${kindName}不能堆叠`, at("stacked"));
  }
  return chart;
}

/** A new chart: the kind's defaults, then the command's fields. */
export function newChart(fields: ChartFields & { type: ChartKind }, at: (field: string) => string): ChartModel {
  const base = emptyChart(fields.type);
  if (fields.type === "pie") base.dataLabels = true;
  return applyChartFields(base, fields, at);
}
