import JSZip from "jszip";
import { ElementSource, Rect, Slide } from "./pptx-virtual-dom";
import { getDirectChild } from "./pptx-xml";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const SHAPE_TAGS = new Set(["sp", "pic", "cxnSp", "grpSp", "graphicFrame"]);

/** Edited shapes, keyed by slide part then by p:cNvPr id. */
export type DirtyShapes = Map<string, Set<string>>;

export function markDirty(dirty: DirtyShapes, source: ElementSource | undefined): void {
  if (!source) return;
  let ids = dirty.get(source.part);
  if (!ids) dirty.set(source.part, ids = new Set());
  ids.add(source.shapeId);
}

function toEmu(rect: Rect, source: ElementSource) {
  const { scaleX, scaleY, offsetX, offsetY } = source.xform;
  return {
    x: Math.round((rect.x - offsetX) / scaleX),
    y: Math.round((rect.y - offsetY) / scaleY),
    cx: Math.max(0, Math.round(rect.w / scaleX)),
    cy: Math.max(0, Math.round(rect.h / scaleY))
  };
}

function findShapeNode(doc: Document, shapeId: string): Element | null {
  for (const node of Array.from(doc.getElementsByTagName("*"))) {
    if ((node.localName || "") !== "cNvPr" || node.getAttribute("id") !== shapeId) continue;
    const owner = node.parentNode?.parentNode as Element | null;
    if (owner && SHAPE_TAGS.has(owner.localName)) return owner;
  }
  return null;
}

function writeXfrm(doc: Document, shape: Element, emu: ReturnType<typeof toEmu>): void {
  // graphicFrame keeps its transform in p:xfrm; every other shape uses spPr/a:xfrm.
  let xfrm: Element | null;
  if (shape.localName === "graphicFrame" || shape.localName === "grpSp") {
    const holder = shape.localName === "grpSp" ? getDirectChild(shape, "grpSpPr") : shape;
    xfrm = holder ? getDirectChild(holder, "xfrm") : null;
    if (!xfrm) return;
  } else {
    let spPr = getDirectChild(shape, "spPr");
    if (!spPr) {
      spPr = doc.createElementNS(P_NS, "p:spPr");
      shape.appendChild(spPr);
    }
    xfrm = getDirectChild(spPr, "xfrm");
    if (!xfrm) {
      xfrm = doc.createElementNS(A_NS, "a:xfrm");
      spPr.insertBefore(xfrm, spPr.firstChild);
    }
  }

  const ensure = (name: "off" | "ext"): Element => {
    let node = getDirectChild(xfrm!, name);
    if (!node) {
      node = doc.createElementNS(A_NS, `a:${name}`);
      xfrm!.appendChild(node);
    }
    return node;
  };
  const off = ensure("off");
  off.setAttribute("x", String(emu.x));
  off.setAttribute("y", String(emu.y));
  const ext = ensure("ext");
  ext.setAttribute("cx", String(emu.cx));
  ext.setAttribute("cy", String(emu.cy));
}

/**
 * Write the edited geometry back into the original package. Only slide parts
 * that were touched are rewritten, and only the targeted shape nodes inside
 * them change, so everything the engine cannot edit yet (charts, animations,
 * masters, media) is carried over byte-for-byte.
 */
export async function exportEditedPptx(
  zip: JSZip,
  slides: Slide[],
  dirty: DirtyShapes
): Promise<Blob> {
  const serializer = new XMLSerializer();

  for (const [part, shapeIds] of dirty) {
    const file = zip.file(part);
    if (!file) continue;
    const doc = new DOMParser().parseFromString(await file.async("string"), "application/xml");
    if (doc.getElementsByTagName("parsererror").length > 0) {
      throw new Error(`无法解析 ${part}，已取消导出。`);
    }

    for (const shapeId of shapeIds) {
      const shape = findShapeNode(doc, shapeId);
      if (!shape) continue;

      // A shape is represented by 1-2 elements (fill + text); either carries the same box.
      const element = slides
        .flatMap(slide => slide.elements)
        .find(candidate => candidate.source?.part === part && candidate.source.shapeId === shapeId);
      if (!element || !element.source) {
        shape.parentNode?.removeChild(shape);
      } else {
        writeXfrm(doc, shape, toEmu(element.rect, element.source));
      }
    }

    let xml = serializer.serializeToString(doc);
    if (!xml.startsWith("<?xml")) {
      xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
    }
    zip.file(part, xml);
  }

  return zip.generateAsync({
    type: "blob",
    mimeType: PPTX_MIME,
    compression: "DEFLATE"
  });
}
