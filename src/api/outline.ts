import { annotatedPurpose, readTemplateRecord, type TemplateNodeSemantics, type TemplatePageSemantics } from "../design/template-policy";
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

/** Native paragraph/run metrics for read-only layout diagnostics (all lengths in points). */
export interface OutlineTextParagraph {
  runs: Array<{ text: string; fontSizePt: number; letterSpacingPt: number }>;
  lineSpacing?: { unit: "percent" | "points"; value: number };
  spaceBeforePt: number;
  spaceAfterPt: number;
  marginLeftPt: number;
  indentPt: number;
}

export interface OutlineShape {
  /** Persisted semantic policy; a manually changed annotation name takes precedence. */
  template?: TemplateNodeSemantics;
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
  /** Same bounded text as `text`, retaining native breaks, run sizes and paragraph spacing. */
  textParagraphs?: OutlineTextParagraph[];
  textBox?: { margins: {left: number; right: number; top: number; bottom: number}; wrap?: string; autoFit: string; verticalOverflow?: string; horizontalOverflow?: string };
  style?: { fontSizePt?: number; bold?: boolean; italic?: boolean; color?: string; opacity?: number; align?: string; /** Where the text sits in its box when it is not the top. */ anchor?: "middle" | "bottom" };
  /** A colour, null for none, a gradient, or a pattern or picture the commands cannot describe. */
  fill?: string | null | OutlineGradient | { other: "pattern" | "picture" };
  /** Opacity of the fill when it is see-through (1 is left out). */
  opacity?: number;
  /** The drop shadow, when there is one. */
  shadow?: OutlineShadow;
  line?: { color: string; widthPt: number; dash: string } | null;
  children?: OutlineShape[];
  /** Set when the shape's name binds it to a template field (see `bindOutline`): its text is then the field's keyword, the real text is the sample. */
  binding?: string;
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
    textColor?: string;
    fontSizePt?: number;
    titleSizePt?: number;
    unsupported?: string;
  };
}

export interface OutlineSlide {
  /** Optional authoring semantics persisted in the PPTX extension record. */
  template?: TemplatePageSemantics;
  index: number;
  id: string;
  layout?: string;
  /** The page's background (its own, or the layout's or master's): a colour, or a gradient. */
  background?: string | OutlineGradient;
  /** The visible full-page paint can be a shape over a different page background. */
  appearance?: { background?: string | OutlineGradient; source: "slide" | "shape"; shapeId?: string; tone: "light" | "dark" | "mixed" | "unknown" };
  shapes: OutlineShape[];
}

export function slideAppearance(slide: OutlineSlide, size: { w: number; h: number }): NonNullable<OutlineSlide["appearance"]> {
  const leaves = (nodes: OutlineShape[]): OutlineShape[] => nodes.flatMap(node => node.children ? leaves(node.children) : [node]);
  let background = slide.background;
  let shapeId: string | undefined;
  for (const shape of leaves(slide.shapes)) {
    if (shape.text?.trim()) break;
    const r = shape.rect;
    if (shape.type !== "shape" || !["rect", "roundRect"].includes(shape.shape ?? "") || (shape.opacity ?? 1) < .95
      || r.x > .05 || r.y > .05 || r.x + r.w < size.w - .05 || r.y + r.h < size.h - .05) continue;
    const fill = shape.fill;
    if (typeof fill === "string" && /^#[a-f0-9]{6}$/i.test(fill)) { background = fill; shapeId = shape.id; }
    else if (fill && typeof fill === "object" && "gradient" in fill && fill.gradient.stops.every(stop => (stop.opacity ?? 1) >= .95)) { background = fill; shapeId = shape.id; }
  }
  const colors = typeof background === "string" ? [background] : background?.gradient.stops.map(stop => stop.color) ?? [];
  const dark = colors.filter(color => /^#[a-f0-9]{6}$/i.test(color)).map(color => {
    const [r,g,b] = [1,3,5].map(index => parseInt(color.slice(index,index+2),16));
    return .299*r+.587*g+.114*b < 110;
  });
  const tone = !dark.length ? "unknown" : dark.every(Boolean) ? "dark" : dark.every(value => !value) ? "light" : "mixed";
  return { background, source: shapeId ? "shape" : "slide", ...(shapeId ? { shapeId } : {}), tone };
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

function describeBackground(slide: Slide): { background?: string | OutlineGradient } {
  const element = slide.elements.find((candidate): candidate is ShapeElement => candidate.type === "shape" && candidate.id === "slide_background");
  if (!element) return {};
  const { fill } = describeFill(element);
  return typeof fill === "string" || (fill && "gradient" in fill) ? { background: fill as string | OutlineGradient } : {};
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
        stacked: chart.stacked,
        textColor: lower(chart.font.color),
        fontSizePt: chart.labelPt ?? 12,
        titleSizePt: chart.titlePt ?? 18
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
    if (element?.body) {
      const body = element.body, cm = (px: number) => round(px * emuPerPx / 360000, 3);
      out.textBox = { margins: {left: cm(body.marginLeft), right: cm(body.marginRight), top: cm(body.marginTop), bottom: cm(body.marginBottom)}, wrap: body.wrap, autoFit: body.autoFit, verticalOverflow: body.verticalOverflow, horizontalOverflow: body.horizontalOverflow };
    }
    if (element?.paragraphs) {
      const pt = (px: number) => round(px * emuPerPx / 12700, 4);
      let remaining = maxText;
      out.textParagraphs = [];
      for (const paragraph of element.paragraphs) {
        if (remaining <= 0) break;
        const runs: OutlineTextParagraph["runs"] = [];
        for (const source of paragraph.runs) {
          if (remaining <= 0) break;
          const text = source.content.slice(0, remaining);
          remaining -= text.length;
          runs.push({ text, fontSizePt: pt(source.style.fontSize), letterSpacingPt: pt(source.style.letterSpacing ?? 0) });
        }
        const spacing = paragraph.style.lineSpacing;
        out.textParagraphs.push({
          runs,
          ...(spacing ? { lineSpacing: { unit: spacing.unit, value: spacing.unit === "points" ? pt(spacing.value) : spacing.value } } : {}),
          spaceBeforePt: pt(paragraph.style.spaceBefore ?? 0),
          spaceAfterPt: pt(paragraph.style.spaceAfter ?? 0),
          marginLeftPt: pt(paragraph.style.marginLeft),
          indentPt: pt(paragraph.style.indent)
        });
        remaining--; // The paragraph separator included in out.text.
      }
    }
    if (run) {
      out.style = {
        // Sizes are whole pixels internally; PowerPoint sizes come in half points.
        fontSizePt: Math.round((run.style.fontSize * emuPerPx) / 12700 * 2) / 2,
        bold: run.style.bold,
        italic: !!run.style.italic,
        color: splitColor(run.style.color).hex,
        ...(splitColor(run.style.color).opacity < 1 ? { opacity: round(splitColor(run.style.color).opacity) } : {}),
        align: element?.paragraphs?.[0]?.style.align,
        ...(element?.body?.verticalAnchor === "middle" || element?.body?.verticalAnchor === "bottom" ? { anchor: element.body.verticalAnchor } : {})
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
        const page: OutlineSlide = {
          index: scene.index,
          id: scene.slide.id,
          ...(scene.slide.layout ? { layout: scene.slide.layout } : {}),
          ...describeBackground(scene.slide),
          shapes: scene.roots.map(node => describeNode(node, emuPerPx, maxText))
        };
        if (scene.slide.rawXml) {
          const record = readTemplateRecord(new DOMParser().parseFromString(scene.slide.rawXml, "application/xml"));
          if (record?.page) page.template = { ...record.page };
          const assign = (nodes: OutlineShape[]) => {
            for (const node of nodes) {
              const saved = record?.nodes[node.id];
              const named = annotatedPurpose(node.name);
              if (saved && (saved.name === undefined || saved.name === node.name)) node.template = { ...saved };
              else if (named) node.template = { purpose: named, name: node.name };
              if (node.children) assign(node.children);
            }
          };
          assign(page.shapes);
        }
        page.appearance = slideAppearance(page, {w: (doc.size.width * emuPerPx) / 360000, h: (doc.size.height * emuPerPx) / 360000});
        return page;
      })
  };
}
