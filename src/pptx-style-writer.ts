import { ShapeStyleOp } from "./editor/shape-ops";
import { getDirectChild, getDirectChildren } from "./pptx-xml";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const FILL_TAGS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
const GEOMETRY_TAGS = ["prstGeom", "custGeom"];
const ANCHOR = { top: "t", middle: "ctr", bottom: "b" } as const;

function solidFill(doc: Document, color: string): Element {
  const solid = doc.createElementNS(A_NS, "a:solidFill");
  const rgb = doc.createElementNS(A_NS, "a:srgbClr");
  rgb.setAttribute("val", color.replace("#", "").toUpperCase());
  solid.appendChild(rgb);
  return solid;
}

/** Schema order inside spPr: xfrm, geometry, fill, ln, effects... */
function insertAfter(parent: Element, node: Element, ...after: string[][]): void {
  for (const tags of after) {
    const anchors = getDirectChildren(parent, ...tags);
    if (anchors.length > 0) {
      parent.insertBefore(node, anchors[anchors.length - 1].nextSibling);
      return;
    }
  }
  parent.insertBefore(node, parent.firstChild);
}

function ensureSpPr(shape: Element): Element {
  let spPr = getDirectChild(shape, "spPr");
  if (!spPr) {
    spPr = shape.ownerDocument.createElementNS(shape.namespaceURI, "p:spPr");
    const after = getDirectChild(shape, "nvSpPr", "nvCxnSpPr", "nvPicPr");
    shape.insertBefore(spPr, after ? after.nextSibling : shape.firstChild);
  }
  return spPr;
}

function writeFill(shape: Element, color: string | null): void {
  const spPr = ensureSpPr(shape);
  const doc = spPr.ownerDocument;
  for (const fill of getDirectChildren(spPr, ...FILL_TAGS)) spPr.removeChild(fill);
  const fill = color === null ? doc.createElementNS(A_NS, "a:noFill") : solidFill(doc, color);
  insertAfter(spPr, fill, GEOMETRY_TAGS, ["xfrm"]);
}

function writeLine(shape: Element, op: Extract<ShapeStyleOp, { kind: "line" }>): void {
  const spPr = ensureSpPr(shape);
  const doc = spPr.ownerDocument;
  let ln = getDirectChild(spPr, "ln");
  if (!ln) {
    ln = doc.createElementNS(A_NS, "a:ln");
    insertAfter(spPr, ln, FILL_TAGS, GEOMETRY_TAGS, ["xfrm"]);
  }

  if (op.widthPt !== undefined) ln.setAttribute("w", String(Math.round(op.widthPt * 12700)));

  if (op.color !== undefined) {
    for (const fill of getDirectChildren(ln, ...FILL_TAGS)) ln.removeChild(fill);
    const fill = op.color === null ? doc.createElementNS(A_NS, "a:noFill") : solidFill(doc, op.color);
    ln.insertBefore(fill, ln.firstChild);
  }

  if (op.dash !== undefined) {
    for (const dash of getDirectChildren(ln, "prstDash", "custDash")) ln.removeChild(dash);
    if (op.dash !== "solid") {
      const dash = doc.createElementNS(A_NS, "a:prstDash");
      dash.setAttribute("val", op.dash);
      insertAfter(ln, dash, FILL_TAGS);
    }
  }
}

/** Replay recorded look changes, in order, on a shape node. */
export function writeShapeStyleOps(shape: Element, ops: ShapeStyleOp[]): void {
  for (const op of ops) {
    if (op.kind === "fill") writeFill(shape, op.color);
    else if (op.kind === "line") writeLine(shape, op);
    else {
      const bodyPr = getDirectChild(getDirectChild(shape, "txBody") || shape, "bodyPr");
      bodyPr?.setAttribute("anchor", ANCHOR[op.value]);
    }
  }
}
