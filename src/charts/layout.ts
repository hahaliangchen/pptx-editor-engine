import { FillStyle, Rect, RenderElement, ShapeElement, SlideElement, TextElement } from "../pptx-virtual-dom";
import { ChartElement, ChartModel, formatNumber, seriesColor, sliceColor } from "./model";

const AXIS_LINE = "#bfbfbf";
const GRID_LINE = "#d9d9d9";
const WHITE = "#ffffff";

// Sizes are in points, the unit PowerPoint stores them in; `pxPerPt` converts.
const TITLE_PT = 18;
const LABEL_PT = 12;
const GRID_PT = 0.75;
const LINE_PT = 2.25;
const MARKER_PT = 6;

/** How wide `text` is at `size` px: a rough measure that only has to place legends and labels. */
export function textWidth(text: string, size: number): number {
  let width = 0;
  for (const char of text) {
    const code = char.codePointAt(0)!;
    if (code >= 0x2e80) width += size;
    else if (char === " ") width += size * 0.3;
    else if ("il.,:;'|!".includes(char)) width += size * 0.3;
    else if (/[A-Z%@mwMW]/.test(char)) width += size * 0.72;
    else width += size * 0.55;
  }
  return width;
}

const solid = (color: string): FillStyle => ({ type: "solid", color });

class Builder {
  public elements: RenderElement[] = [];
  private serial = 0;

  constructor(private owner: string) {}

  private id(kind: string): string {
    return `chart_${this.owner}_${kind}${this.serial++}`;
  }

  public rect(rect: Rect, fill: string, line?: { color: string; width: number }, shapeType = "rect"): void {
    this.elements.push({
      type: "shape",
      id: this.id("shape"),
      rect: { ...rect },
      shapeType,
      fill,
      ...(line ? { border: { color: line.color, width: line.width } } : {}),
      computedStyle: {
        fill: solid(fill),
        ...(line ? { line: { fill: solid(line.color), width: line.width } } : {})
      }
    });
  }

  /** A straight line from (x1, y1) to (x2, y2). */
  public line(x1: number, y1: number, x2: number, y2: number, color: string, width: number): void {
    this.elements.push({
      type: "shape",
      id: this.id("line"),
      rect: { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) },
      shapeType: "line",
      flipH: false,
      flipV: (x2 - x1) * (y2 - y1) < 0,
      fill: "transparent",
      border: { color, width },
      computedStyle: { fill: { type: "none" }, line: { fill: solid(color), width } }
    });
  }

  /** A closed or open path in the page's own coordinates. */
  public path(
    points: Array<Array<{ x: number; y: number }>>,
    paint: { fill?: string; line?: { color: string; width: number }; closed: boolean }
  ): void {
    const flat = points.flat();
    if (flat.length === 0) return;
    const left = Math.min(...flat.map(p => p.x));
    const top = Math.min(...flat.map(p => p.y));
    const right = Math.max(...flat.map(p => p.x));
    const bottom = Math.max(...flat.map(p => p.y));
    const width = Math.max(right - left, 0.01);
    const height = Math.max(bottom - top, 0.01);
    const commands: Array<{ type: "moveTo" | "lineTo" | "close"; x?: number; y?: number }> = [];
    for (const run of points) {
      run.forEach((point, index) => commands.push({ type: index === 0 ? "moveTo" : "lineTo", x: point.x - left, y: point.y - top }));
      if (paint.closed) commands.push({ type: "close" });
    }
    this.elements.push({
      type: "shape",
      id: this.id("path"),
      rect: { x: left, y: top, w: width, h: height },
      shapeType: "custom",
      customGeometry: {
        paths: [{
          width,
          height,
          commands: commands as never,
          ...(paint.fill ? {} : { fill: "none" as const })
        }]
      },
      fill: paint.fill ?? "transparent",
      ...(paint.line ? { border: { color: paint.line.color, width: paint.line.width } } : {}),
      computedStyle: {
        fill: paint.fill ? solid(paint.fill) : { type: "none" },
        ...(paint.line ? { line: { fill: solid(paint.line.color), width: paint.line.width, join: "round", cap: "round" } } : {})
      }
    } as ShapeElement);
  }

  public text(
    rect: Rect,
    content: string,
    style: { size: number; color: string; bold?: boolean; align: "left" | "center" | "right"; anchor?: "top" | "middle" | "bottom"; wrap?: boolean },
    font: ChartModel["font"]
  ): void {
    const runStyle = {
      fontSize: style.size,
      color: style.color,
      bold: !!style.bold,
      align: style.align,
      ...(font.family ? { fontFamily: font.family } : {}),
      ...(font.eastAsianFamily ? { eastAsianFontFamily: font.eastAsianFamily } : {})
    };
    const element: TextElement = {
      type: "text",
      id: this.id("text"),
      rect: { ...rect },
      content,
      style: runStyle,
      paragraphs: [{
        style: { align: style.align, level: 0, marginLeft: 0, indent: 0 },
        runs: [{ content, style: { ...runStyle } }]
      }],
      body: {
        marginLeft: 0,
        marginRight: 0,
        marginTop: 0,
        marginBottom: 0,
        verticalAnchor: style.anchor ?? "middle",
        autoFit: "none",
        wrap: style.wrap ? "square" : "none",
        fontScale: 1,
        textDirection: "horz"
      }
    };
    this.elements.push(element);
  }
}

/** Round axis limits and a step, in the 1-2-5 progression Excel and PowerPoint use. */
export function niceAxis(minValue: number, maxValue: number, maxTicks: number): { min: number; max: number; step: number } {
  let low = Math.min(0, minValue);
  let high = Math.max(0, maxValue);
  if (high === low) high = low + 1;
  const range = high - low;
  const target = Math.max(2, maxTicks);
  const magnitude = Math.pow(10, Math.floor(Math.log10(range)));
  let step = magnitude;
  for (const factor of [0.1, 0.2, 0.5, 1, 2, 5, 10]) {
    step = magnitude * factor;
    if (range / step <= target) break;
  }
  // Excel leaves a little headroom above the highest value.
  if (high > 0) high += range * 0.05;
  if (low < 0) low -= range * 0.05;
  return {
    min: low === 0 ? 0 : Math.floor(low / step) * step,
    max: Math.ceil(high / step - 1e-9) * step,
    step
  };
}

/** Is `color` dark enough that white text reads better on it? */
function isDark(color: string): boolean {
  const channel = (offset: number) => parseInt(color.slice(offset, offset + 2), 16);
  return 0.299 * channel(1) + 0.587 * channel(3) + 0.114 * channel(5) < 150;
}

interface Legend {
  entries: Array<{ label: string; color: string }>;
}

function legendOf(chart: ChartModel): Legend {
  return chart.kind === "pie"
    ? { entries: chart.categories.map((label, index) => ({ label, color: sliceColor(chart, index) })) }
    : { entries: chart.series.map((series, index) => ({ label: series.name, color: seriesColor(chart, index) })) };
}

/**
 * Draw `el` as ordinary shapes and text, so the renderer needs no chart code and
 * a chart always lays itself out for the size it currently has.
 */
export function expandChart(el: ChartElement): RenderElement[] {
  const { chart, rect } = el;
  const u = el.pxPerPt;
  const out = new Builder(el.id);

  if (chart.unsupported) {
    out.rect(rect, "#f2f2f2", { color: AXIS_LINE, width: u });
    out.text(
      { x: rect.x + 8 * u, y: rect.y, w: rect.w - 16 * u, h: rect.h },
      `图表（暂不支持预览：${chart.unsupported}）`,
      { size: 14 * u, color: "#7f7f7f", align: "center", wrap: true },
      chart.font
    );
    return out.elements;
  }

  if (chart.background) out.rect(rect, chart.background);
  const fontSize = (chart.labelPt ?? LABEL_PT) * u;
  const pad = 8 * u;
  let left = rect.x + pad;
  let right = rect.x + rect.w - pad;
  let top = rect.y + pad;
  let bottom = rect.y + rect.h - pad;

  if (chart.title) {
    const height = (chart.titlePt ?? TITLE_PT) * u * 1.5;
    out.text({ x: left, y: top, w: right - left, h: height }, chart.title, { size: (chart.titlePt ?? TITLE_PT) * u, color: chart.font.color, align: "center" }, chart.font);
    top += height + 4 * u;
  }

  drawLegend(out, chart, { left, right, top, bottom }, u, region => {
    left = region.left; right = region.right; top = region.top; bottom = region.bottom;
  });

  if (chart.categories.length === 0 || chart.series.length === 0) return out.elements;
  const plot = { left, right, top, bottom };
  if (chart.kind === "pie") drawPie(out, chart, plot, u);
  else drawCartesian(out, chart, plot, u, fontSize);
  return out.elements;
}

type Region = { left: number; right: number; top: number; bottom: number };

function drawLegend(out: Builder, chart: ChartModel, region: Region, u: number, shrink: (rest: Region) => void): void {
  const { entries } = legendOf(chart);
  if (chart.legend === "none" || entries.length === 0) {
    shrink(region);
    return;
  }
  const size = (chart.labelPt ?? LABEL_PT) * u;
  const swatch = chart.kind === "line" ? 14 * u : 8 * u;
  const gap = 4 * u;
  const spacing = 14 * u;
  const entryWidth = (label: string) => swatch + gap + textWidth(label, size);
  const rest = { ...region };

  const drawEntry = (x: number, y: number, entry: { label: string; color: string }) => {
    const middle = y + size * 0.8;
    if (chart.kind === "line") {
      out.line(x, middle, x + swatch, middle, entry.color, LINE_PT * u);
      out.rect({ x: x + swatch / 2 - 3 * u, y: middle - 3 * u, w: 6 * u, h: 6 * u }, entry.color, undefined, "ellipse");
    } else {
      out.rect({ x, y: middle - swatch / 2, w: swatch, h: swatch }, entry.color);
    }
    out.text({ x: x + swatch + gap, y, w: textWidth(entry.label, size) + 2 * u, h: size * 1.6 }, entry.label, { size, color: chart.font.color, align: "left" }, chart.font);
  };

  if (chart.legend === "top" || chart.legend === "bottom") {
    // Entries fill rows left to right and start a new row when one is full.
    const available = region.right - region.left;
    const rows: Array<typeof entries> = [[]];
    let used = 0;
    for (const entry of entries) {
      const width = entryWidth(entry.label);
      if (used > 0 && used + spacing + width > available) {
        rows.push([]);
        used = 0;
      }
      rows[rows.length - 1].push(entry);
      used += (used > 0 ? spacing : 0) + width;
    }
    const rowHeight = size * 1.8;
    const total = rows.length * rowHeight;
    const startY = chart.legend === "bottom" ? region.bottom - total : region.top;
    rows.forEach((row, index) => {
      const width = row.reduce((sum, entry, i) => sum + entryWidth(entry.label) + (i > 0 ? spacing : 0), 0);
      let x = region.left + (available - width) / 2;
      for (const entry of row) {
        drawEntry(x, startY + index * rowHeight, entry);
        x += entryWidth(entry.label) + spacing;
      }
    });
    if (chart.legend === "bottom") rest.bottom -= total + 4 * u;
    else rest.top += total + 4 * u;
  } else {
    const width = Math.max(...entries.map(entry => entryWidth(entry.label)));
    const rowHeight = size * 1.8;
    const total = entries.length * rowHeight;
    const x = chart.legend === "right" ? region.right - width : region.left;
    const y = region.top + (region.bottom - region.top - total) / 2;
    entries.forEach((entry, index) => drawEntry(x, y + index * rowHeight, entry));
    if (chart.legend === "right") rest.right -= width + 12 * u;
    else rest.left += width + 12 * u;
  }
  shrink(rest);
}

function drawPie(out: Builder, chart: ChartModel, region: Region, u: number): void {
  const values = chart.categories.map((_, index) => Math.max(0, chart.series[0].values[index] ?? 0));
  const total = values.reduce((sum, value) => sum + value, 0);
  const size = (chart.labelPt ?? LABEL_PT) * u;
  const centerX = (region.left + region.right) / 2;
  const centerY = (region.top + region.bottom) / 2;
  const margin = chart.dataLabels ? size * 2.6 : size * 0.6;
  const radius = Math.max(0, Math.min(region.right - region.left, region.bottom - region.top) / 2 - margin);
  if (total <= 0 || radius <= 0) return;

  let start = -Math.PI / 2;
  values.forEach((value, index) => {
    if (value <= 0) return;
    const sweep = (value / total) * Math.PI * 2;
    const steps = Math.max(2, Math.ceil(sweep / (Math.PI / 90)));
    const arc = Array.from({ length: steps + 1 }, (_, step) => ({
      x: centerX + Math.cos(start + (sweep * step) / steps) * radius,
      y: centerY + Math.sin(start + (sweep * step) / steps) * radius
    }));
    const points = sweep >= Math.PI * 2 - 1e-6 ? arc.slice(0, -1) : [{ x: centerX, y: centerY }, ...arc];
    out.path([points], { fill: sliceColor(chart, index), line: { color: WHITE, width: 1.25 * u }, closed: true });

    if (chart.dataLabels) {
      const middle = start + sweep / 2;
      const text = `${Math.round((value / total) * 100)}%`;
      const reach = radius + size * 1.1;
      const width = textWidth(text, size) + 4 * u;
      out.text(
        { x: centerX + Math.cos(middle) * reach - width / 2, y: centerY + Math.sin(middle) * reach - size * 0.7, w: width, h: size * 1.4 },
        text,
        { size, color: chart.font.color, align: "center" },
        chart.font
      );
    }
    start += sweep;
  });
}

function drawCartesian(out: Builder, chart: ChartModel, region: Region, u: number, size: number): void {
  const horizontal = chart.kind === "bar";
  const categories = chart.categories.length;
  const seriesCount = chart.series.length;
  const stacked = chart.stacked && chart.kind !== "pie";

  // Value range: stacked series add up, others are taken one by one.
  let low = 0;
  let high = 0;
  for (let i = 0; i < categories; i++) {
    let positive = 0;
    let negative = 0;
    for (const series of chart.series) {
      const value = series.values[i] ?? 0;
      if (stacked) {
        if (value >= 0) positive += value; else negative += value;
      } else {
        high = Math.max(high, value);
        low = Math.min(low, value);
      }
    }
    if (stacked) {
      high = Math.max(high, positive);
      low = Math.min(low, negative);
    }
  }

  const height = region.bottom - region.top;
  const width = region.right - region.left;
  const maxTicks = Math.max(3, Math.min(10, Math.floor((horizontal ? width : height) / (size * 3.2))));
  const axis = niceAxis(low, high, maxTicks);
  const ticks: number[] = [];
  for (let value = axis.min; value <= axis.max + axis.step * 1e-6; value += axis.step) ticks.push(Math.round(value / axis.step) * axis.step);
  const tickLabel = (value: number) => formatNumber(value);
  const widestTick = Math.max(...ticks.map(value => textWidth(tickLabel(value), size)));

  // Reserve room for the two axes' labels.
  let plot: Region;
  let slot: number;
  let categoryLines = 1;
  if (horizontal) {
    const widest = Math.max(...chart.categories.map(label => textWidth(label, size)));
    const labelWidth = Math.min(widest + 8 * u, width * 0.35);
    plot = { left: region.left + labelWidth, right: region.right - widestTick / 2 - 2 * u, top: region.top + size * 0.3, bottom: region.bottom - size * 1.25 - 6 * u };
    slot = (plot.bottom - plot.top) / categories;
  } else {
    const labelWidth = widestTick + 8 * u;
    const available = width - labelWidth;
    const eachSlot = available / categories;
    categoryLines = Math.min(3, Math.max(1, ...chart.categories.map(label => Math.ceil(textWidth(label, size) / Math.max(eachSlot - 4 * u, 1)))));
    plot = { left: region.left + labelWidth, right: region.right, top: region.top + size * 0.7, bottom: region.bottom - categoryLines * size * 1.25 - 6 * u };
    slot = (plot.right - plot.left) / categories;
  }
  const plotWidth = plot.right - plot.left;
  const plotHeight = plot.bottom - plot.top;
  if (plotWidth <= 0 || plotHeight <= 0) return;

  const along = (value: number) => horizontal
    ? plot.left + ((value - axis.min) / (axis.max - axis.min)) * plotWidth
    : plot.bottom - ((value - axis.min) / (axis.max - axis.min)) * plotHeight;
  const zero = along(Math.max(axis.min, Math.min(axis.max, 0)));

  // Gridlines and value labels.
  for (const value of ticks) {
    const position = along(value);
    const onAxis = value === Math.max(axis.min, Math.min(axis.max, 0));
    const color = onAxis ? AXIS_LINE : GRID_LINE;
    if (horizontal) {
      out.line(position, plot.top, position, plot.bottom, color, GRID_PT * u);
      const labelWidth = textWidth(tickLabel(value), size) + 4 * u;
      out.text({ x: position - labelWidth / 2, y: plot.bottom + 4 * u, w: labelWidth, h: size * 1.3 }, tickLabel(value), { size, color: chart.font.color, align: "center" }, chart.font);
    } else {
      out.line(plot.left, position, plot.right, position, color, GRID_PT * u);
      out.text({ x: region.left, y: position - size * 0.65, w: plot.left - region.left - 6 * u, h: size * 1.3 }, tickLabel(value), { size, color: chart.font.color, align: "right" }, chart.font);
    }
  }

  // Category labels.
  chart.categories.forEach((label, index) => {
    if (horizontal) {
      const center = plot.bottom - (index + 0.5) * slot;
      out.text({ x: region.left, y: center - size * 0.65, w: plot.left - region.left - 6 * u, h: size * 1.3 }, label, { size, color: chart.font.color, align: "right" }, chart.font);
    } else {
      out.text(
        { x: plot.left + index * slot + 2 * u, y: plot.bottom + 4 * u, w: slot - 4 * u, h: categoryLines * size * 1.25 },
        label,
        { size, color: chart.font.color, align: "center", anchor: "top", wrap: true },
        chart.font
      );
    }
  });

  const labelOf = (value: number) => formatNumber(value);
  const drawLabel = (centerX: number, centerY: number, value: number, color: string) => {
    const text = labelOf(value);
    const labelWidth = textWidth(text, size) + 4 * u;
    out.text({ x: centerX - labelWidth / 2, y: centerY - size * 0.7, w: labelWidth, h: size * 1.4 }, text, { size, color, align: "center" }, chart.font);
  };

  if (chart.kind === "line") {
    const above = new Array<number>(categories).fill(0);
    const below = new Array<number>(categories).fill(0);
    chart.series.forEach((series, seriesIndex) => {
      const color = seriesColor(chart, seriesIndex);
      const runs: Array<Array<{ x: number; y: number }>> = [];
      let run: Array<{ x: number; y: number }> = [];
      const points: Array<{ x: number; y: number; value: number }> = [];
      for (let i = 0; i < categories; i++) {
        const raw = series.values[i];
        if (raw === null || raw === undefined) {
          if (run.length > 0) runs.push(run);
          run = [];
          continue;
        }
        let value = raw;
        if (stacked) {
          if (value >= 0) { above[i] += value; value = above[i]; } else { below[i] += value; value = below[i]; }
        }
        const point = { x: plot.left + (i + 0.5) * slot, y: along(value) };
        run.push(point);
        points.push({ ...point, value: raw });
      }
      if (run.length > 0) runs.push(run);
      const multi = runs.filter(entry => entry.length > 1);
      if (multi.length > 0) out.path(multi, { line: { color, width: LINE_PT * u }, closed: false });
      for (const point of points) {
        out.rect({ x: point.x - (MARKER_PT * u) / 2, y: point.y - (MARKER_PT * u) / 2, w: MARKER_PT * u, h: MARKER_PT * u }, color, undefined, "ellipse");
        if (chart.dataLabels) drawLabel(point.x, point.y - size * 1.1, point.value, chart.font.color);
      }
    });
    return;
  }

  // Columns and bars.
  const gap = chart.gapWidth / 100;
  const groupSize = stacked ? slot / (1 + gap) : slot / (seriesCount + gap);
  const startOffset = stacked ? (slot - groupSize) / 2 : (slot - groupSize * seriesCount) / 2;
  for (let i = 0; i < categories; i++) {
    let positive = 0;
    let negative = 0;
    chart.series.forEach((series, seriesIndex) => {
      const value = series.values[i];
      if (value === null || value === undefined) return;
      const from = stacked ? (value >= 0 ? positive : negative) : 0;
      const to = from + value;
      if (stacked) {
        if (value >= 0) positive = to; else negative = to;
      }
      const a = along(Math.max(axis.min, Math.min(axis.max, from)));
      const b = along(Math.max(axis.min, Math.min(axis.max, to)));
      const lowEdge = Math.min(a, b);
      const extent = Math.abs(b - a);
      const offset = startOffset + (stacked ? 0 : seriesIndex * groupSize);
      const color = seriesColor(chart, seriesIndex);
      const box: Rect = horizontal
        ? { x: Math.min(a, b), y: plot.bottom - (i + 1) * slot + offset, w: extent, h: groupSize }
        : { x: plot.left + i * slot + offset, y: lowEdge, w: groupSize, h: extent };
      if (extent > 0) out.rect(box, color);
      if (chart.dataLabels) {
        if (stacked) {
          if (extent > size) drawLabel(box.x + box.w / 2, box.y + box.h / 2, value, isDark(color) ? WHITE : "#404040");
        } else if (horizontal) {
          const labelWidth = textWidth(labelOf(value), size) + 4 * u;
          const outside = value >= 0 ? box.x + box.w + labelWidth / 2 + 2 * u : box.x - labelWidth / 2 - 2 * u;
          drawLabel(outside, box.y + box.h / 2, value, chart.font.color);
        } else {
          drawLabel(box.x + box.w / 2, value >= 0 ? box.y - size * 0.8 : box.y + box.h + size * 0.8, value, chart.font.color);
        }
      }
    });
  }
  void zero;
}

/** Replace every chart in `elements` by the shapes and text that draw it. */
export function expandForRender(elements: SlideElement[]): RenderElement[] {
  const out: RenderElement[] = [];
  for (const element of elements) {
    if (element.type === "chart") out.push(...expandChart(element));
    else out.push(element);
  }
  return out;
}
