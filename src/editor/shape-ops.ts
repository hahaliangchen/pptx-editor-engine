import { withAlpha, splitColor } from "../color";
import { FillStyle, ImageElement, ShapeElement, TextElement } from "../pptx-virtual-dom";

export type LineDash = "solid" | "dash" | "sysDot";

export interface GradientSpec {
  kind: "linear" | "radial";
  /** Degrees, as PowerPoint shows them: 0 runs left to right, 90 top to bottom. */
  angle: number;
  /** Positions are 0..1 along the gradient. */
  stops: Array<{ position: number; color: string; opacity?: number }>;
}

/** What a page's background can be set to: a colour or a gradient. */
export interface BackgroundSpec {
  color?: string;
  gradient?: GradientSpec;
}

/** The fill the renderer draws for a gradient spec. */
export function gradientFill(gradient: GradientSpec): FillStyle {
  const stops = gradient.stops
    .map(stop => ({ position: stop.position, color: withAlpha(stop.color, stop.opacity ?? 1) }))
    .sort((a, b) => a.position - b.position);
  return { type: "gradient", kind: gradient.kind, stops, angle: gradient.angle, rotateWithShape: true };
}

export interface ShadowSpec {
  color: string;
  /** 0..1 */
  opacity: number;
  blurPt: number;
  distancePt: number;
  /** Degrees: 0 casts to the right, 90 downwards. */
  angle: number;
}

/** One change to a shape's look, replayed on the OOXML node at export time. */
export type ShapeStyleOp =
  /** `color` null removes the fill; `opacity` (0..1) makes it see-through. */
  | { kind: "fill"; color: string | null; opacity?: number }
  | { kind: "gradient"; gradient: GradientSpec }
  /** Only the opacity of the fill changes; `base` is the fill as drawn now, for a file that leaves it to a style. */
  | { kind: "opacity"; opacity: number; base?: FillStyle }
  /** `shadow` null removes the drop shadow. */
  | { kind: "shadow"; shadow: ShadowSpec | null }
  /** Only the given fields change; `color` null removes the border. */
  | { kind: "line"; color?: string | null; widthPt?: number; dash?: LineDash }
  | { kind: "anchor"; value: "top" | "middle" | "bottom" }
  /** The shape's name (p:cNvPr@name), which a template uses to say what a shape is for. */
  | { kind: "name"; value: string };

const PT_EMU = 12700;
const DEFAULT_LINE_PT = 1;

/** What a snapshot of a shape element holds, so undo can restore it exactly. */
export interface ShapeSnapshot {
  fill: string;
  border: ShapeElement["border"];
  computedStyle: ShapeElement["computedStyle"];
  /** Pictures keep their effects here. */
  effects?: ImageElement["effects"];
}

/** Elements whose look the style ops change. */
export type Styled = ShapeElement | ImageElement;

const clone = <T>(value: T): T => (value === undefined ? value : JSON.parse(JSON.stringify(value)));

export const snapshotShape = (shape: Styled): ShapeSnapshot =>
  shape.type === "image"
    ? { fill: "", border: undefined, computedStyle: undefined, effects: clone(shape.effects) }
    : { fill: shape.fill, border: clone(shape.border), computedStyle: clone(shape.computedStyle) };

export function restoreShape(shape: Styled, snapshot: ShapeSnapshot): void {
  if (shape.type === "image") {
    shape.effects = clone(snapshot.effects);
    return;
  }
  shape.fill = snapshot.fill;
  shape.border = clone(snapshot.border);
  shape.computedStyle = clone(snapshot.computedStyle);
}

/** Apply a fill/line op to the renderer's shape element. */
export function applyShapeStyleOp(shape: Styled, op: ShapeStyleOp, emuPerPx: number): void {
  if (op.kind === "name") return;
  if (op.kind === "shadow") {
    const holder = shape.type === "image" ? shape : (shape.computedStyle ??= { fill: { type: "none" } });
    const effects = { ...(shape.type === "image" ? shape.effects : shape.computedStyle?.effects) };
    if (op.shadow === null) delete effects.outerShadow;
    else {
      effects.outerShadow = {
        color: splitColor(op.shadow.color).hex,
        opacity: op.shadow.opacity,
        blur: op.shadow.blurPt * PT_EMU / emuPerPx,
        distance: op.shadow.distancePt * PT_EMU / emuPerPx,
        direction: op.shadow.angle
      };
    }
    const next = Object.keys(effects).length > 0 ? effects : undefined;
    if (shape.type === "image") shape.effects = next;
    else (holder as NonNullable<ShapeElement["computedStyle"]>).effects = next;
    return;
  }
  if (shape.type === "image") return;

  if (op.kind === "fill") {
    const color = op.color === null ? null : withAlpha(op.color, op.opacity ?? 1);
    shape.fill = color ?? "transparent";
    if (shape.computedStyle) shape.computedStyle.fill = color === null ? { type: "none" } : { type: "solid", color };
    return;
  }
  if (op.kind === "gradient") {
    const fill = gradientFill(op.gradient);
    shape.fill = fill.type === "gradient" ? fill.stops[0].color : "transparent";
    if (shape.computedStyle) shape.computedStyle.fill = fill;
    else shape.computedStyle = { fill };
    return;
  }
  if (op.kind === "opacity") {
    const fill = shape.computedStyle?.fill;
    if (fill?.type === "solid") {
      fill.color = withAlpha(fill.color, op.opacity);
      shape.fill = fill.color;
    } else if (fill?.type === "gradient") {
      fill.stops = fill.stops.map(stop => ({ ...stop, color: withAlpha(stop.color, op.opacity) }));
      shape.fill = fill.stops[0].color;
    } else if (!fill && shape.fill !== "transparent") {
      shape.fill = withAlpha(shape.fill, op.opacity);
    }
    return;
  }
  if (op.kind !== "line") return;

  const current = shape.computedStyle?.line;
  const currentColor = current?.fill.type === "solid" ? current.fill.color : shape.border?.color;
  const color = op.color === undefined ? currentColor : op.color;
  if (!color) {
    if (shape.computedStyle) shape.computedStyle.line = undefined;
    shape.border = undefined;
    return;
  }
  const width = op.widthPt !== undefined
    ? op.widthPt * PT_EMU / emuPerPx
    : current?.width ?? shape.border?.width ?? DEFAULT_LINE_PT * PT_EMU / emuPerPx;
  shape.border = { color, width };
  if (shape.computedStyle) {
    shape.computedStyle.line = {
      ...(current || {}),
      fill: { type: "solid", color },
      width,
      dash: op.dash !== undefined ? (op.dash === "solid" ? undefined : op.dash) : current?.dash
    };
  }
}

export function applyAnchorOp(text: TextElement, value: "top" | "middle" | "bottom"): void {
  if (!text.body) return;
  text.body = { ...text.body, verticalAnchor: value };
}
