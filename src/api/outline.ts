import { LineStyle, PptVirtualDocument, ShapeElement, Slide, TextElement } from "../pptx-virtual-dom";
import { SceneNode, SlideScene, buildScene, textOf } from "./scene";

export interface CmRect {
  x: number;
  y: number;
  w: number;
  h: number;
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
  fill?: string | null;
  line?: { color: string; widthPt: number; dash: string } | null;
  children?: OutlineShape[];
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

  const shape = node.elements.find((element): element is ShapeElement => element.type === "shape");
  if (shape && node.kind !== "image" && node.kind !== "table") {
    out.shape = shape.shapeType;
    const fill = shape.computedStyle?.fill;
    out.fill = fill ? (fill.type === "solid" ? lower(fill.color) : null) : shape.fill === "transparent" ? null : lower(shape.fill);
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
