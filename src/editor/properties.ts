import { FillStyle, ShapeElement, TextElement } from "../pptx-virtual-dom";
import { LineDash } from "./shape-ops";

export type Paint =
  | { kind: "none" }
  | { kind: "solid"; color: string }
  | { kind: "mixed" }
  /** Gradient, picture or pattern: shown as such, replaced when a colour is picked. */
  | { kind: "other" };

export interface SelectionProperties {
  /** Shapes whose fill/border can change. */
  shapeCount: number;
  textCount: number;
  fill: Paint;
  lineColor: Paint;
  /** null when mixed or no border. */
  lineWidthPt: number | null;
  lineDash: LineDash | null;
  align: "left" | "center" | "right" | null;
  anchor: "top" | "middle" | "bottom" | null;
}

const PT_EMU = 12700;

function paintOfFill(shape: ShapeElement): Paint {
  const fill: FillStyle | undefined = shape.computedStyle?.fill;
  if (!fill) {
    return shape.fill === "transparent" ? { kind: "none" } : { kind: "solid", color: shape.fill.toLowerCase() };
  }
  if (fill.type === "none") return { kind: "none" };
  if (fill.type === "solid") return { kind: "solid", color: fill.color.toLowerCase() };
  return { kind: "other" };
}

function paintOfLine(shape: ShapeElement): Paint {
  const line = shape.computedStyle?.line;
  if (line) {
    if (line.fill.type === "solid") return { kind: "solid", color: line.fill.color.toLowerCase() };
    return line.fill.type === "none" ? { kind: "none" } : { kind: "other" };
  }
  return shape.border ? { kind: "solid", color: shape.border.color.toLowerCase() } : { kind: "none" };
}

const samePaint = (a: Paint, b: Paint): boolean =>
  a.kind === b.kind && (a.kind !== "solid" || (b.kind === "solid" && a.color === b.color));

function mergePaints(paints: Paint[]): Paint {
  if (paints.length === 0) return { kind: "none" };
  return paints.every(paint => samePaint(paint, paints[0])) ? paints[0] : { kind: "mixed" };
}

function same<T>(values: T[]): T | null {
  return values.length > 0 && values.every(value => value === values[0]) ? values[0] : null;
}

export function describeSelection(
  shapes: ShapeElement[],
  texts: TextElement[],
  emuPerPx: number
): SelectionProperties {
  const widths = shapes
    .filter(shape => paintOfLine(shape).kind !== "none")
    .map(shape => Math.round((shape.computedStyle?.line?.width ?? shape.border?.width ?? 0) * emuPerPx / PT_EMU * 100) / 100);
  const dashes = shapes
    .filter(shape => paintOfLine(shape).kind !== "none")
    .map(shape => (shape.computedStyle?.line?.dash as LineDash | undefined) || "solid");
  const aligns = texts.flatMap(text => (text.paragraphs || []).map(paragraph => paragraph.style.align));

  return {
    shapeCount: shapes.length,
    textCount: texts.length,
    fill: mergePaints(shapes.map(paintOfFill)),
    lineColor: mergePaints(shapes.map(paintOfLine)),
    lineWidthPt: same(widths),
    lineDash: same(dashes),
    align: same(aligns),
    anchor: same(texts.map(text => text.body?.verticalAnchor ?? "top"))
  };
}
