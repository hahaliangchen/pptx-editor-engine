import { ShapeElement, TextElement } from "../pptx-virtual-dom";

export type LineDash = "solid" | "dash" | "sysDot";

/** One change to a shape's look, replayed on the OOXML node at export time. */
export type ShapeStyleOp =
  /** `color` null removes the fill. */
  | { kind: "fill"; color: string | null }
  /** Only the given fields change; `color` null removes the border. */
  | { kind: "line"; color?: string | null; widthPt?: number; dash?: LineDash }
  | { kind: "anchor"; value: "top" | "middle" | "bottom" };

const PT_EMU = 12700;
const DEFAULT_LINE_PT = 1;

/** What a snapshot of a shape element holds, so undo can restore it exactly. */
export interface ShapeSnapshot {
  fill: string;
  border: ShapeElement["border"];
  computedStyle: ShapeElement["computedStyle"];
}

const clone = <T>(value: T): T => (value === undefined ? value : JSON.parse(JSON.stringify(value)));

export const snapshotShape = (shape: ShapeElement): ShapeSnapshot => ({
  fill: shape.fill,
  border: clone(shape.border),
  computedStyle: clone(shape.computedStyle)
});

export function restoreShape(shape: ShapeElement, snapshot: ShapeSnapshot): void {
  shape.fill = snapshot.fill;
  shape.border = clone(snapshot.border);
  shape.computedStyle = clone(snapshot.computedStyle);
}

/** Apply a fill/line op to the renderer's shape element. */
export function applyShapeStyleOp(shape: ShapeElement, op: ShapeStyleOp, emuPerPx: number): void {
  if (op.kind === "fill") {
    shape.fill = op.color ?? "transparent";
    if (shape.computedStyle) {
      shape.computedStyle.fill = op.color === null ? { type: "none" } : { type: "solid", color: op.color };
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
