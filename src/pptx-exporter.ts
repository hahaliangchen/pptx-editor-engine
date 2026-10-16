import JSZip from "jszip";
import { ElementSource, GroupRef, Rect, Slide } from "./pptx-virtual-dom";
import { getDirectChild } from "./pptx-xml";
import { TextOp } from "./editor/text-ops";
import { writeTextOps } from "./pptx-text-writer";
import { ShapeStyleOp } from "./editor/shape-ops";
import { writeShapeStyleOps } from "./pptx-style-writer";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const SHAPE_TAGS = new Set(["sp", "pic", "cxnSp", "grpSp", "graphicFrame"]);

/** Edited shapes, keyed by slide part then by p:cNvPr id. */
export type DirtyShapes = Map<string, Set<string>>;

/** Edits recorded per shape, in the order they were made. */
export type OpLog<T> = Map<string, Map<string, T[]>>;

export type TextEdits = OpLog<TextOp>;
export type StyleEdits = OpLog<ShapeStyleOp>;

export function pushOp<T>(log: OpLog<T>, source: ElementSource, op: T): void {
  let shapes = log.get(source.part);
  if (!shapes) log.set(source.part, shapes = new Map());
  const ops = shapes.get(source.shapeId) || [];
  ops.push(op);
  shapes.set(source.shapeId, ops);
}

export function popOp<T>(log: OpLog<T>, source: ElementSource): void {
  const ops = log.get(source.part)?.get(source.shapeId);
  ops?.pop();
  if (ops && ops.length === 0) log.get(source.part)!.delete(source.shapeId);
  if (log.get(source.part)?.size === 0) log.delete(source.part);
}

export const pushTextOp = pushOp<TextOp>;
export const popTextOp = popOp<TextOp>;

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

/**
 * Moving or resizing a group only touches the group's own off/ext: its
 * children live in the group's child space and follow automatically. The new
 * box is derived from how the members' bounding box changed since opening.
 */
function writeGroupXfrm(shape: Element, ref: GroupRef, now: Rect): void {
  const grpSpPr = getDirectChild(shape, "grpSpPr");
  const xfrm = grpSpPr ? getDirectChild(grpSpPr, "xfrm") : null;
  const off = xfrm ? getDirectChild(xfrm, "off") : null;
  const ext = xfrm ? getDirectChild(xfrm, "ext") : null;
  if (!off || !ext) return;

  const { scaleX, scaleY } = ref.xform;
  const before = ref.origRect;
  const sx = before.w > 0 ? now.w / before.w : 1;
  const sy = before.h > 0 ? now.h / before.h : 1;
  const x = parseInt(off.getAttribute("x") || "0", 10);
  const y = parseInt(off.getAttribute("y") || "0", 10);
  const cx = parseInt(ext.getAttribute("cx") || "0", 10);
  const cy = parseInt(ext.getAttribute("cy") || "0", 10);
  off.setAttribute("x", String(Math.round(x + (now.x - before.x) / scaleX)));
  off.setAttribute("y", String(Math.round(y + (now.y - before.y) / scaleY)));
  ext.setAttribute("cx", String(Math.max(0, Math.round(cx * sx))));
  ext.setAttribute("cy", String(Math.max(0, Math.round(cy * sy))));
}

function findShapeNode(doc: Document, shapeId: string): Element | null {
  for (const node of Array.from(doc.getElementsByTagName("*"))) {
    if ((node.localName || "") !== "cNvPr" || node.getAttribute("id") !== shapeId) continue;
    const owner = node.parentNode?.parentNode as Element | null;
    if (owner && SHAPE_TAGS.has(owner.localName)) return owner;
  }
  return null;
}

/** Tables size themselves from gridCol/tr, not from the frame, so scale those too. */
function scaleTableGrid(shape: Element, oldCx: number, oldCy: number, emu: ReturnType<typeof toEmu>): void {
  const scaleAll = (tag: string, attribute: string, factor: number) => {
    for (const node of Array.from(shape.getElementsByTagNameNS("*", tag))) {
      const value = parseInt(node.getAttribute(attribute) || "", 10);
      if (Number.isFinite(value)) node.setAttribute(attribute, String(Math.max(1, Math.round(value * factor))));
    }
  };
  if (oldCx > 0 && emu.cx !== oldCx) scaleAll("gridCol", "w", emu.cx / oldCx);
  if (oldCy > 0 && emu.cy !== oldCy) scaleAll("tr", "h", emu.cy / oldCy);
}

function writeXfrm(
  doc: Document,
  shape: Element,
  emu: ReturnType<typeof toEmu>,
  rotation?: number
): void {
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

  if (rotation !== undefined && shape.localName !== "graphicFrame") {
    const turns = Math.round((((rotation % 360) + 360) % 360) * 60000);
    if (turns === 0) xfrm.removeAttribute("rot");
    else xfrm.setAttribute("rot", String(turns));
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
  if (shape.localName === "graphicFrame") {
    scaleTableGrid(
      shape,
      parseInt(ext.getAttribute("cx") || "0", 10),
      parseInt(ext.getAttribute("cy") || "0", 10),
      emu
    );
  }
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
  dirty: DirtyShapes,
  textEdits: TextEdits = new Map(),
  styleEdits: StyleEdits = new Map()
): Promise<Blob> {
  const serializer = new XMLSerializer();
  // Edited parts are put back once the archive is built, so a later export
  // (after an undo, say) starts again from the original package.
  const originals = new Map<string, string>();

  try {
    const parts = new Set([...dirty.keys(), ...textEdits.keys(), ...styleEdits.keys()]);
    for (const part of parts) {
      const shapeIds = dirty.get(part) || new Set<string>();
      const file = zip.file(part);
      if (!file) continue;
      const originalXml = await file.async("string");
      originals.set(part, originalXml);
      const doc = new DOMParser().parseFromString(originalXml, "application/xml");
      if (doc.getElementsByTagName("parsererror").length > 0) {
        throw new Error(`无法解析 ${part}，已取消导出。`);
      }

      for (const shapeId of shapeIds) {
        const shape = findShapeNode(doc, shapeId);
        if (!shape) continue;

        // A shape is 1-2 elements (fill + text) and a table is one per cell;
        // the box of the object is the union of them.
        const inGroup = shape.localName === "grpSp";
        const members = slides
          .flatMap(slide => slide.elements)
          .filter(candidate => candidate.source?.part === part && (inGroup
            ? candidate.source.groups?.some(group => group.shapeId === shapeId)
            : candidate.source.shapeId === shapeId));
        if (members.length === 0) {
          shape.parentNode?.removeChild(shape);
        } else {
          const left = Math.min(...members.map(m => m.rect.x));
          const top = Math.min(...members.map(m => m.rect.y));
          const right = Math.max(...members.map(m => m.rect.x + m.rect.w));
          const bottom = Math.max(...members.map(m => m.rect.y + m.rect.h));
          const box = { x: left, y: top, w: right - left, h: bottom - top };
          if (inGroup) {
            const ref = members[0].source!.groups!.find(group => group.shapeId === shapeId)!;
            writeGroupXfrm(shape, ref, box);
          } else {
            writeXfrm(
              doc,
              shape,
              toEmu(box, members[0].source!),
              members.map(m => m.rotation).find((r): r is number => typeof r === "number")
            );
          }
        }
      }

      for (const [shapeId, ops] of styleEdits.get(part) || []) {
        const shape = findShapeNode(doc, shapeId);
        if (shape && ops.length > 0) writeShapeStyleOps(shape, ops);
      }

      for (const [shapeId, ops] of textEdits.get(part) || []) {
        const shape = findShapeNode(doc, shapeId);
        if (shape && ops.length > 0) writeTextOps(shape, ops);
      }

      let xml = serializer.serializeToString(doc);
      if (!xml.startsWith("<?xml")) {
        xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
      }
      zip.file(part, xml);
    }

    return await zip.generateAsync({
      type: "blob",
      mimeType: PPTX_MIME,
      compression: "DEFLATE"
    });
  } finally {
    for (const [part, xml] of originals) zip.file(part, xml);
  }
}
