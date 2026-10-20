import { splitColor } from "./color";
import { BackgroundSpec, GradientSpec, ShadowSpec, ShapeStyleOp } from "./editor/shape-ops";
import { FillStyle } from "./pptx-virtual-dom";
import { getDirectChild, getDirectChildren } from "./pptx-xml";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const FILL_TAGS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
const GEOMETRY_TAGS = ["prstGeom", "custGeom"];
const ANCHOR = { top: "t", middle: "ctr", bottom: "b" } as const;

/** `<a:srgbClr>` for a model colour, with `<a:alpha>` when it is see-through (or `opacity` says so). */
function colorNode(doc: Document, color: string, opacity?: number): Element {
  const split = splitColor(color);
  const rgb = doc.createElementNS(A_NS, "a:srgbClr");
  rgb.setAttribute("val", split.hex.replace("#", "").toUpperCase());
  const alpha = opacity ?? split.opacity;
  if (alpha < 1) {
    const node = doc.createElementNS(A_NS, "a:alpha");
    node.setAttribute("val", String(Math.round(Math.max(0, alpha) * 100000)));
    rgb.appendChild(node);
  }
  return rgb;
}

function solidFill(doc: Document, color: string, opacity?: number): Element {
  const solid = doc.createElementNS(A_NS, "a:solidFill");
  solid.appendChild(colorNode(doc, color, opacity));
  return solid;
}

const percent = (fraction: number) => String(Math.round(Math.max(0, Math.min(1, fraction)) * 100000));

function gradientFill(
  doc: Document,
  kind: "linear" | "radial",
  angle: number,
  stops: Array<{ position: number; color: string; opacity?: number }>
): Element {
  const fill = doc.createElementNS(A_NS, "a:gradFill");
  fill.setAttribute("rotWithShape", "1");
  const list = doc.createElementNS(A_NS, "a:gsLst");
  for (const stop of stops) {
    const gs = doc.createElementNS(A_NS, "a:gs");
    gs.setAttribute("pos", percent(stop.position));
    gs.appendChild(colorNode(doc, stop.color, stop.opacity));
    list.appendChild(gs);
  }
  fill.appendChild(list);
  if (kind === "radial") {
    const path = doc.createElementNS(A_NS, "a:path");
    path.setAttribute("path", "circle");
    const rect = doc.createElementNS(A_NS, "a:fillToRect");
    for (const side of ["l", "t", "r", "b"]) rect.setAttribute(side, "50000");
    path.appendChild(rect);
    fill.appendChild(path);
  } else {
    const lin = doc.createElementNS(A_NS, "a:lin");
    lin.setAttribute("ang", String(Math.round((((angle % 360) + 360) % 360) * 60000)));
    lin.setAttribute("scaled", "0");
    fill.appendChild(lin);
  }
  return fill;
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

function setFillNode(spPr: Element, fill: Element): void {
  for (const old of getDirectChildren(spPr, ...FILL_TAGS)) spPr.removeChild(old);
  insertAfter(spPr, fill, GEOMETRY_TAGS, ["xfrm"]);
}

function writeFill(shape: Element, color: string | null, opacity?: number): void {
  const spPr = ensureSpPr(shape);
  const doc = spPr.ownerDocument;
  setFillNode(spPr, color === null ? doc.createElementNS(A_NS, "a:noFill") : solidFill(doc, color, opacity));
}

function writeGradient(shape: Element, gradient: GradientSpec): void {
  const spPr = ensureSpPr(shape);
  setFillNode(spPr, gradientFill(spPr.ownerDocument, gradient.kind, gradient.angle, gradient.stops));
}

const COLOR_TAGS = ["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"];

/** Change only the opacity of the fill the shape has; a fill left to a style is written out from `base`. */
function writeOpacity(shape: Element, opacity: number, base?: FillStyle): void {
  const spPr = ensureSpPr(shape);
  const doc = spPr.ownerDocument;
  const fill = getDirectChild(spPr, "solidFill", "gradFill");
  if (!fill) {
    if (base?.type === "solid") setFillNode(spPr, solidFill(doc, base.color, opacity));
    else if (base?.type === "gradient") {
      setFillNode(spPr, gradientFill(doc, base.kind, base.angle ?? 0, base.stops.map(stop => ({ ...stop, opacity }))));
    }
    return;
  }
  const holders = fill.localName === "gradFill"
    ? getDirectChildren(getDirectChild(fill, "gsLst") ?? fill, "gs")
    : [fill];
  for (const holder of holders) {
    for (const color of getDirectChildren(holder, ...COLOR_TAGS)) {
      for (const old of getDirectChildren(color, "alpha")) color.removeChild(old);
      if (opacity < 1) {
        const alpha = doc.createElementNS(A_NS, "a:alpha");
        alpha.setAttribute("val", String(Math.round(Math.max(0, opacity) * 100000)));
        color.appendChild(alpha);
      }
    }
  }
}

/** Effects that sort before a drop shadow inside a:effectLst. */
const BEFORE_OUTER_SHADOW = ["blur", "fillOverlay", "glow", "innerShdw"];

function writeShadow(shape: Element, shadow: ShadowSpec | null): void {
  const spPr = ensureSpPr(shape);
  const doc = spPr.ownerDocument;
  let list = getDirectChild(spPr, "effectLst");
  if (!list) {
    list = doc.createElementNS(A_NS, "a:effectLst");
    insertAfter(spPr, list, ["ln"], FILL_TAGS, GEOMETRY_TAGS, ["xfrm"]);
  }
  for (const old of getDirectChildren(list, "outerShdw")) list.removeChild(old);
  // An empty list stays: it also switches off a shadow the theme would give the shape.
  if (!shadow) return;
  const node = doc.createElementNS(A_NS, "a:outerShdw");
  node.setAttribute("blurRad", String(Math.round(shadow.blurPt * 12700)));
  node.setAttribute("dist", String(Math.round(shadow.distancePt * 12700)));
  node.setAttribute("dir", String(Math.round((((shadow.angle % 360) + 360) % 360) * 60000)));
  node.setAttribute("algn", "ctr");
  node.setAttribute("rotWithShape", "0");
  node.appendChild(colorNode(doc, shadow.color, shadow.opacity));
  const anchors = getDirectChildren(list, ...BEFORE_OUTER_SHADOW);
  list.insertBefore(node, anchors.length > 0 ? anchors[anchors.length - 1].nextSibling : list.firstChild);
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
    if (op.kind === "fill") writeFill(shape, op.color, op.opacity);
    else if (op.kind === "gradient") writeGradient(shape, op.gradient);
    else if (op.kind === "opacity") writeOpacity(shape, op.opacity, op.base);
    else if (op.kind === "shadow") writeShadow(shape, op.shadow);
    else if (op.kind === "line") writeLine(shape, op);
    else if (op.kind === "name") {
      const holder = Array.from(shape.children).find(child => /^nv.*Pr$/.test(child.localName));
      const properties = holder ? Array.from(holder.children).find(child => child.localName === "cNvPr") : undefined;
      properties?.setAttribute("name", op.value);
    } else {
      const bodyPr = getDirectChild(getDirectChild(shape, "txBody") || shape, "bodyPr");
      bodyPr?.setAttribute("anchor", ANCHOR[op.value]);
    }
  }
}

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** Give a slide its own background: `<p:bg>` is the first child of `<p:cSld>`. */
export function writeBackground(doc: Document, spec: BackgroundSpec): void {
  const cSld = doc.getElementsByTagNameNS(P_NS, "cSld")[0];
  if (!cSld) return;
  for (const old of getDirectChildren(cSld, "bg")) cSld.removeChild(old);
  const fill = spec.gradient
    ? gradientFill(doc, spec.gradient.kind, spec.gradient.angle, spec.gradient.stops)
    : solidFill(doc, spec.color ?? "#ffffff");
  const properties = doc.createElementNS(P_NS, "p:bgPr");
  properties.appendChild(fill);
  properties.appendChild(doc.createElementNS(A_NS, "a:effectLst"));
  const bg = doc.createElementNS(P_NS, "p:bg");
  bg.appendChild(properties);
  cSld.insertBefore(bg, cSld.firstChild);
}
