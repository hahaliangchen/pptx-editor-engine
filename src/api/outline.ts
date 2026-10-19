import { ImageElement, LineStyle, PptVirtualDocument, ShadowStyle, ShapeElement, Slide, TextElement } from "../pptx-virtual-dom";
import { ChartElement, ChartKind, LegendPosition } from "../charts/model";
import { splitColor } from "../color";
import { TableElement, tableText } from "../tables/model";
import { SceneNode, SlideScene, buildScene, textOf } from "./scene";

export interface CmRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface OutlineGradient {
  gradient: {
    type: "linear" | "radial";
    /** Degrees: 0 runs left to right, 90 top to bottom. */
    angle: number;
    /** Positions are percentages along the gradient. */
    stops: Array<{ position: number; color: string; opacity?: number }>;
  };
}

export interface OutlineShadow {
  color: string;
  opacity: number;
  blurPt: number;
  distancePt: number;
  /** Degrees: 0 casts to the right, 90 downwards. */
  angle: number;
}

export interface OutlineShape {
  id: string;
  name: string;
  type: SceneNode["kind"];
  /** Preset geometry for shapes, e.g. "roundRect". */
  shape?: string;
  placeholder?: string;
  rect: CmRect;
  rotation: number;
  text?: string;
  textTruncated?: boolean;
  style?: { fontSizePt?: number; bold?: boolean; italic?: boolean; color?: string; align?: string };
  /** A colour, null for none, a gradient, or a pattern or picture the commands cannot describe. */
  fill?: string | null | OutlineGradient | { other: "pattern" | "picture" };
  /** Opacity of the fill when it is see-through (1 is left out). */
  opacity?: number;
  /** The drop shadow, when there is one. */
  shadow?: OutlineShadow;
  line?: { color: string; widthPt: number; dash: string } | null;
  children?: OutlineShape[];
  /** The text of a table, row by row, and its grid; merged cells make `merged` true. */
  table?: { rows: string[][]; columns: number; columnWidths: number[]; rowHeights: number[]; merged?: boolean };
  /** The data and settings of a chart; `unsupported` names kinds that are only drawn as a placeholder. */
  chart?: {
    type: ChartKind;
    title?: string;
    categories: string[];
    series: Array<{ name: string; values: Array<number | null>; color?: string }>;
    legend: LegendPosition;
    dataLabels: boolean;
    stacked: boolean;
    unsupported?: string;
  };
}

export interface OutlineSlide {
  index: number;
  id: string;
  layout?: string;
  shapes: OutlineShape[];
}

export interface Outline {
  slideSize: { w: number; h: number };
  layouts: string[];
  slides: OutlineSlide[];
}

export interface OutlineOptions {
  /** Only these slides (1-based); default all. */
  slides?: number[];
  /** Longer text is cut to this many characters. */
  maxTextLength?: number;
  layouts?: string[];
}

/** Colours are reported in one spelling, whatever case the file or a command used. */
const lower = (color: string) => color.toLowerCase();

const round = (value: number, digits = 2) => {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
};

export const emuPerPxOf = (slide: Slide, widthPx: number): number => slide.emuPerPx || 12192000 / widthPx;

function describeFill(shape: ShapeElement): { fill: NonNullable<OutlineShape["fill"]> | null; opacity?: number } {
  const fill = shape.computedStyle?.fill;
  const plain = (color: string) => {
    const split = splitColor(color);
    return { fill: split.hex, ...(split.opacity < 1 ? { opacity: round(split.opacity) } : {}) };
  };
  if (!fill) return shape.fill === "transparent" ? { fill: null } : plain(shape.fill);
  if (fill.type === "none") return { fill: null };
  if (fill.type === "solid") return plain(fill.color);
  if (fill.type === "gradient") {
    return {
      fill: {
        gradient: {
          type: fill.kind,
          angle: round(fill.angle ?? 0, 1),
          stops: fill.stops.map(stop => {
            const split = splitColor(stop.color);
            return { position: round(stop.position * 100, 1), color: split.hex, ...(split.opacity < 1 ? { opacity: round(split.opacity) } : {}) };
          })
        }
      }
    };
  }
  return { fill: { other: fill.type } };
}

function describeShadow(shadow: ShadowStyle | undefined, emuPerPx: number): OutlineShadow | undefined {
  if (!shadow) return undefined;
  return {
    color: lower(shadow.color),
    opacity: round(shadow.opacity),
    blurPt: round((shadow.blur * emuPerPx) / 12700, 1),
    distancePt: round((shadow.distance * emuPerPx) / 12700, 1),
    angle: round(shadow.direction, 1)
  };
}

function describeNode(node: SceneNode, emuPerPx: number, maxText: number): OutlineShape {
  const cm = (px: number) => round((px * emuPerPx) / 360000);
  const out: OutlineShape = {
    id: node.id,
    name: node.name,
    type: node.kind,
    rect: { x: cm(node.rect.x), y: cm(node.rect.y), w: cm(node.rect.w), h: cm(node.rect.h) },
    rotation: round(node.rotation, 1)
  };
  if (node.placeholder) out.placeholder = node.placeholder;

  if (node.kind === "group") {
    out.children = node.children.map(child => describeNode(child, emuPerPx, maxText));
    return out;
  }

  if (node.kind === "chart") {
    const chart = node.elements.find((element): element is ChartElement => element.type === "chart")!.chart;
    out.chart = chart.unsupported
      ? { type: chart.kind, categories: [], series: [], legend: chart.legend, dataLabels: false, stacked: false, unsupported: chart.unsupported }
      : {
        type: chart.kind,
        ...(chart.title ? { title: chart.title } : {}),
        categories: chart.categories,
        series: chart.series.map(series => ({
          name: series.name,
          values: series.values,
          ...(series.color ? { color: lower(series.color) } : {})
        })),
        legend: chart.legend,
        dataLabels: chart.dataLabels,
        stacked: chart.stacked
      };
    return out;
  }

  if (node.kind === "table") {
    const table = node.elements.find((element): element is TableElement => element.type === "table");
    if (table) {
      const frame = table.table;
      const scaleX = frame.columns.reduce((sum, width) => sum + width, 0) > 0 ? table.rect.w / frame.columns.reduce((sum, width) => sum + width, 0) : 1;
      const scaleY = frame.rows.reduce((sum, row) => sum + row.height, 0) > 0 ? table.rect.h / frame.rows.reduce((sum, row) => sum + row.height, 0) : 1;
      out.table = {
        rows: tableText(frame),
        columns: frame.columns.length,
        columnWidths: frame.columns.map(width => cm(width * scaleX)),
        rowHeights: frame.rows.map(row => cm(row.height * scaleY)),
        ...(frame.merged ? { merged: true } : {})
      };
    }
    return out;
  }

  const picture = node.elements.find((element): element is ImageElement => element.type === "image");
  if (picture) {
    const shadow = describeShadow(picture.effects?.outerShadow, emuPerPx);
    if (shadow) out.shadow = shadow;
  }

  const shape = node.elements.find((element): element is ShapeElement => element.type === "shape");
  if (shape && node.kind !== "image") {
    out.shape = shape.shapeType;
    const painted = describeFill(shape);
    out.fill = painted.fill;
    if (painted.opacity !== undefined) out.opacity = painted.opacity;
    const shadow = describeShadow(shape.computedStyle?.effects?.outerShadow, emuPerPx);
    if (shadow) out.shadow = shadow;
    const line: LineStyle | undefined = shape.computedStyle?.line;
    if (line && line.fill.type === "solid") {
      out.line = { color: lower(line.fill.color), widthPt: round((line.width * emuPerPx) / 12700), dash: line.dash || "solid" };
    } else if (!line && shape.border) {
      out.line = { color: lower(shape.border.color), widthPt: round((shape.border.width * emuPerPx) / 12700), dash: "solid" };
    } else {
      out.line = null;
    }
  }

  const text = textOf(node);
  if (text) {
    out.text = text.length > maxText ? text.slice(0, maxText) : text;
    if (text.length > maxText) out.textTruncated = true;
    const element = node.elements.find((candidate): candidate is TextElement => candidate.type === "text" && candidate.content !== "");
    const run = element?.paragraphs?.flatMap(paragraph => paragraph.runs).find(candidate => candidate.content !== "");
    if (run) {
      out.style = {
        // Sizes are whole pixels internally; PowerPoint sizes come in half points.
        fontSizePt: Math.round((run.style.fontSize * emuPerPx) / 12700 * 2) / 2,
        bold: run.style.bold,
        italic: !!run.style.italic,
        color: lower(run.style.color),
        align: element?.paragraphs?.[0]?.style.align
      };
    }
  }
  return out;
}

/** What is on the slides, in the units and names the edit commands use. */
export function buildOutline(doc: PptVirtualDocument, options: OutlineOptions = {}, scenes?: SlideScene[]): Outline {
  const all = scenes || buildScene(doc.slides);
  const maxText = options.maxTextLength ?? 500;
  const wanted = options.slides ? new Set(options.slides) : null;
  const emuFirst = doc.slides[0] ? emuPerPxOf(doc.slides[0], doc.size.width) : 6350;
  return {
    slideSize: {
      w: round((doc.size.width * emuFirst) / 360000),
      h: round((doc.size.height * emuFirst) / 360000)
    },
    layouts: options.layouts || [],
    slides: all
      .filter(scene => !wanted || wanted.has(scene.index))
      .map(scene => {
        const emuPerPx = emuPerPxOf(scene.slide, doc.size.width);
        return {
          index: scene.index,
          id: scene.slide.id,
          ...(scene.slide.layout ? { layout: scene.slide.layout } : {}),
          shapes: scene.roots.map(node => describeNode(node, emuPerPx, maxText))
        };
      })
  };
}
