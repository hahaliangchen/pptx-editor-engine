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

export type OrderOp = "front" | "back" | "forward" | "backward";

/** A shape node added by paste/duplicate, already carrying its final ids. */
export interface InsertedShape {
  id: string;
  xml: string;
}

/** A relationship a pasted shape needs in its slide (an image, a hyperlink...). */
export interface RelAdd {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/** A picture file added to the package. */
export interface MediaAdd {
  path: string;
  bytes: Uint8Array;
  extension: string;
  contentType: string;
}

/** Changes to the shape tree itself, kept as per-slide lists in the order they happened. */
export interface StructureEdits {
  orders: Map<string, Array<{ shapeId: string; op: OrderOp }>>;
  inserts: Map<string, InsertedShape[]>;
  relAdds: Map<string, RelAdd[]>;
  media: MediaAdd[];
  /** Files that exist only in the edit session (slides made by "new" or "duplicate", and their relationships). */
  virtual: Map<string, string>;
}

export const emptyStructure = (): StructureEdits => ({
  orders: new Map(),
  inserts: new Map(),
  relAdds: new Map(),
  media: [],
  virtual: new Map()
});

export function pushItem<T>(log: Map<string, T[]>, part: string, item: T): void {
  const list = log.get(part) || [];
  list.push(item);
  log.set(part, list);
}

export function popItem<T>(log: Map<string, T[]>, part: string): void {
  const list = log.get(part);
  list?.pop();
  if (list && list.length === 0) log.delete(part);
}

/** Everything an export replays on top of the original package. */
export interface EditSet {
  dirty: DirtyShapes;
  textEdits: TextEdits;
  styleEdits: StyleEdits;
  structure: StructureEdits;
}

function shapeSiblings(parent: Element): Element[] {
  return Array.from(parent.children).filter(child => SHAPE_TAGS.has(child.localName));
}

function applyOrderOp(node: Element, op: OrderOp): void {
  const parent = node.parentNode as Element | null;
  if (!parent) return;
  const siblings = shapeSiblings(parent);
  const index = siblings.indexOf(node);
  if (op === "front") {
    parent.appendChild(node);
  } else if (op === "back") {
    parent.insertBefore(node, siblings[0]);
  } else if (op === "forward" && index < siblings.length - 1) {
    parent.insertBefore(node, siblings[index + 1].nextSibling);
  } else if (op === "backward" && index > 0) {
    parent.insertBefore(node, siblings[index - 1]);
  }
}

function spTreeOf(doc: Document): Element | null {
  return doc.getElementsByTagNameNS(P_NS, "spTree")[0] || null;
}

/**
 * Replay every recorded edit on one slide part. Order matters: pasted shapes
 * must exist before anything that addresses them, and moves come before
 * z-order changes.
 */
function applyPartEdits(doc: Document, part: string, slides: Slide[], edits: EditSet): void {
  const spTree = spTreeOf(doc);
  for (const inserted of edits.structure.inserts.get(part) || []) {
    const fragment = new DOMParser().parseFromString(inserted.xml, "application/xml");
    if (fragment.getElementsByTagName("parsererror").length > 0 || !spTree) continue;
    spTree.appendChild(doc.importNode(fragment.documentElement, true));
  }

  for (const shapeId of edits.dirty.get(part) || []) {
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
      continue;
    }
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

  for (const [shapeId, ops] of edits.styleEdits.get(part) || []) {
    const shape = findShapeNode(doc, shapeId);
    if (shape && ops.length > 0) writeShapeStyleOps(shape, ops);
  }

  for (const [shapeId, ops] of edits.textEdits.get(part) || []) {
    const shape = findShapeNode(doc, shapeId);
    if (shape && ops.length > 0) writeTextOps(shape, ops);
  }

  for (const { shapeId, op } of edits.structure.orders.get(part) || []) {
    const shape = findShapeNode(doc, shapeId);
    if (shape) applyOrderOp(shape, op);
  }
}

const relsPathOf = (part: string): string => {
  const slash = part.lastIndexOf("/");
  return `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`;
};

const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function readDocument(
  zip: JSZip,
  path: string,
  virtual?: Map<string, string>
): Promise<{ xml: string; doc: Document } | null> {
  const file = virtual?.has(path) ? null : zip.file(path);
  if (!file && !virtual?.has(path)) return null;
  const xml = virtual?.has(path) ? virtual.get(path)! : await file!.async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) {
    throw new Error(`无法解析 ${path}，已取消导出。`);
  }
  return { xml, doc };
}

/** Existing relationships of a slide, plus those pending for it. */
export async function readRelationships(
  zip: JSZip,
  part: string,
  pending: RelAdd[] = [],
  virtual?: Map<string, string>
): Promise<Map<string, RelAdd>> {
  const result = new Map<string, RelAdd>();
  const rels = await readDocument(zip, relsPathOf(part), virtual);
  for (const rel of Array.from(rels?.doc.getElementsByTagNameNS(REL_NS, "Relationship") || [])) {
    result.set(rel.getAttribute("Id") || "", {
      id: rel.getAttribute("Id") || "",
      type: rel.getAttribute("Type") || "",
      target: rel.getAttribute("Target") || "",
      targetMode: rel.getAttribute("TargetMode") || undefined
    });
  }
  for (const rel of pending) result.set(rel.id, rel);
  return result;
}

/** Relationship ids a shape node refers to (r:embed, r:id, r:link ...). */
export function relationshipIdsIn(node: Element): string[] {
  const ids = new Set<string>();
  for (const element of [node, ...Array.from(node.getElementsByTagName("*"))]) {
    for (const attribute of Array.from(element.attributes)) {
      if (attribute.namespaceURI === R_NS && attribute.value) ids.add(attribute.value);
    }
  }
  return [...ids];
}

/**
 * The XML a shape has right now, with every pending edit applied, for copy
 * and paste. Returns the shapes' relationships too, so a paste into another
 * slide can recreate them.
 */
export async function materializeShape(
  zip: JSZip,
  part: string,
  shapeId: string,
  slides: Slide[],
  edits: EditSet
): Promise<{ xml: string; relationships: RelAdd[] } | null> {
  const loaded = await readDocument(zip, part, edits.structure.virtual);
  if (!loaded) return null;
  applyPartEdits(loaded.doc, part, slides, edits);
  const original = findShapeNode(loaded.doc, shapeId);
  if (!original) return null;
  // Custom data (tags) belongs to one specific shape; a copy must not share it.
  const node = original.cloneNode(true) as Element;
  for (const custom of Array.from(node.getElementsByTagNameNS("*", "custDataLst"))) {
    custom.parentNode?.removeChild(custom);
  }
  const known = await readRelationships(zip, part, edits.structure.relAdds.get(part), edits.structure.virtual);
  const relationships = relationshipIdsIn(node)
    .map(id => known.get(id))
    .filter((rel): rel is RelAdd => !!rel);
  return { xml: new XMLSerializer().serializeToString(node), relationships };
}

/**
 * A whole slide as it is right now (every pending edit applied) with its
 * relationships, for duplicating it.
 */
export async function materializeSlide(
  zip: JSZip,
  part: string,
  slides: Slide[],
  edits: EditSet
): Promise<{ xml: string; relationships: RelAdd[] } | null> {
  const loaded = await readDocument(zip, part, edits.structure.virtual);
  if (!loaded) return null;
  applyPartEdits(loaded.doc, part, slides, edits);
  const known = await readRelationships(zip, part, edits.structure.relAdds.get(part), edits.structure.virtual);
  return {
    xml: new XMLSerializer().serializeToString(loaded.doc),
    relationships: [...known.values()]
  };
}

const CONTENT_TYPES = "[Content_Types].xml";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

/** Add picture files and make sure the package declares their file types. */
async function writeMedia(
  zip: JSZip,
  media: MediaAdd[],
  originals: Map<string, string>,
  addedFiles: string[]
): Promise<void> {
  for (const item of media) {
    zip.file(item.path, item.bytes);
    addedFiles.push(item.path);
  }
  const loaded = await readDocument(zip, CONTENT_TYPES);
  if (!loaded) return;
  const root = loaded.doc.documentElement;
  const declared = new Set(
    Array.from(root.getElementsByTagNameNS(CT_NS, "Default")).map(node => (node.getAttribute("Extension") || "").toLowerCase())
  );
  let changed = false;
  for (const item of media) {
    if (declared.has(item.extension.toLowerCase())) continue;
    const node = loaded.doc.createElementNS(CT_NS, "Default");
    node.setAttribute("Extension", item.extension);
    node.setAttribute("ContentType", item.contentType);
    root.insertBefore(node, root.firstChild);
    declared.add(item.extension.toLowerCase());
    changed = true;
  }
  if (changed) {
    originals.set(CONTENT_TYPES, loaded.xml);
    zip.file(CONTENT_TYPES, new XMLSerializer().serializeToString(loaded.doc));
  }
}

const PRESENTATION = "ppt/presentation.xml";
const PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels";
const SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

/** "ppt/slides/../media/x.png" -> "ppt/media/x.png", also for absolute targets. */
export function resolveTarget(baseDir: string, target: string): string {
  const parts = (target.startsWith("/") ? target.slice(1) : `${baseDir}/${target}`).split("/");
  const out: string[] = [];
  for (const part of parts) {
    if (part === "..") out.pop();
    else if (part !== ".") out.push(part);
  }
  return out.join("/");
}

interface Section {
  list: Element;
  ids: string[];
}

/**
 * Sections are ranges of slides, so after slides are added, removed or moved
 * each one is recomputed from where its first surviving slide ended up.
 * Slides without a place of their own (new copies) join the section they sit in.
 */
function updateSections(doc: Document, finalIds: string[]): void {
  const lists = Array.from(doc.getElementsByTagName("*")).filter(node => node.localName === "sectionLst");
  if (lists.length === 0) return;
  const sections: Section[] = [];
  for (const section of Array.from(lists[0].children).filter(node => node.localName === "section")) {
    const list = Array.from(section.children).find(node => node.localName === "sldIdLst");
    if (!list) continue;
    sections.push({
      list,
      ids: Array.from(list.children).filter(node => node.localName === "sldId").map(node => node.getAttribute("id") || "")
    });
  }
  if (sections.length === 0) return;

  const surviving = new Set(finalIds);
  const anchors = sections.map(section => section.ids.find(id => surviving.has(id)) ?? null);
  const members: string[][] = sections.map(() => []);
  let current = 0;
  for (const id of finalIds) {
    const anchored = anchors.findIndex((anchor, index) => anchor === id && index > current);
    if (anchored > 0) current = anchored;
    members[current].push(id);
  }
  sections.forEach((section, index) => {
    const template = Array.from(section.list.children).find(node => node.localName === "sldId");
    const namespace = template?.namespaceURI || section.list.namespaceURI;
    const prefix = template?.prefix ? `${template.prefix}:` : section.list.prefix ? `${section.list.prefix}:` : "";
    for (const child of Array.from(section.list.children)) section.list.removeChild(child);
    for (const id of members[index]) {
      const node = doc.createElementNS(namespace, `${prefix}sldId`);
      node.setAttribute("id", id);
      section.list.appendChild(node);
    }
  });
}

/**
 * Make the package's slide list match the editor's: new slides get an entry,
 * deleted ones lose theirs (and their files), and the order follows `slides`.
 * Slide parts that disappear are kept in memory so a later export can bring
 * them back after an undo.
 */
async function syncSlideList(
  zip: JSZip,
  slides: Slide[],
  originals: Map<string, string>,
  addedFiles: string[],
  removedFiles: Map<string, Uint8Array>
): Promise<void> {
  const presentation = await readDocument(zip, PRESENTATION);
  const relationships = await readDocument(zip, PRESENTATION_RELS);
  if (!presentation || !relationships) return;
  const list = presentation.doc.getElementsByTagNameNS(P_NS, "sldIdLst")[0];
  if (!list) return;

  const relNodes = new Map(
    Array.from(relationships.doc.getElementsByTagNameNS(REL_NS, "Relationship")).map(node => [node.getAttribute("Id") || "", node])
  );
  const entries = Array.from(list.children).filter(node => node.localName === "sldId").map(node => {
    const rId = node.getAttributeNS(R_NS, "id") || "";
    return { node, rId, id: node.getAttribute("id") || "", part: resolveTarget("ppt", relNodes.get(rId)?.getAttribute("Target") || "") };
  });
  const finalParts = slides.map(slide => slide.part!);
  if (finalParts.length === entries.length && finalParts.every((part, index) => entries[index].part === part)) return;

  const byPart = new Map(entries.map(entry => [entry.part, entry]));
  let nextSlideId = Math.max(255, ...entries.map(entry => parseInt(entry.id, 10) || 0));
  let nextRel = Math.max(0, ...[...relNodes.keys()].map(id => parseInt(id.replace(/\D/g, ""), 10) || 0));
  const contentTypes = await readDocument(zip, CONTENT_TYPES);
  const overrides = new Map(
    Array.from(contentTypes?.doc.getElementsByTagNameNS(CT_NS, "Override") || []).map(node => [node.getAttribute("PartName") || "", node])
  );
  let typesChanged = false;

  const finalNodes: Element[] = [];
  const finalIds: string[] = [];
  for (const part of finalParts) {
    let entry = byPart.get(part);
    if (!entry) {
      nextSlideId += 1;
      nextRel += 1;
      const rId = `rId${nextRel}`;
      const rel = relationships.doc.createElementNS(REL_NS, "Relationship");
      rel.setAttribute("Id", rId);
      rel.setAttribute("Type", SLIDE_REL);
      rel.setAttribute("Target", part.replace(/^ppt\//, ""));
      relationships.doc.documentElement.appendChild(rel);
      const node = presentation.doc.createElementNS(P_NS, "p:sldId");
      node.setAttribute("id", String(nextSlideId));
      node.setAttributeNS(R_NS, "r:id", rId);
      entry = { node, rId, id: String(nextSlideId), part };
      if (contentTypes && !overrides.has(`/${part}`)) {
        const override = contentTypes.doc.createElementNS(CT_NS, "Override");
        override.setAttribute("PartName", `/${part}`);
        override.setAttribute("ContentType", SLIDE_CONTENT_TYPE);
        contentTypes.doc.documentElement.appendChild(override);
        typesChanged = true;
      }
    }
    finalNodes.push(entry.node);
    finalIds.push(entry.id);
  }

  // Deleted slides: entry, relationship, content type, the slide's files and its notes page.
  for (const entry of entries.filter(candidate => !finalParts.includes(candidate.part))) {
    relNodes.get(entry.rId)?.parentNode?.removeChild(relNodes.get(entry.rId)!);
    const slideDir = entry.part.slice(0, entry.part.lastIndexOf("/"));
    const rels = await readDocument(zip, relsPathOf(entry.part));
    const doomed = [entry.part, relsPathOf(entry.part)];
    for (const rel of Array.from(rels?.doc.getElementsByTagNameNS(REL_NS, "Relationship") || [])) {
      if ((rel.getAttribute("Type") || "").endsWith("/notesSlide")) {
        const notes = resolveTarget(slideDir, rel.getAttribute("Target") || "");
        doomed.push(notes, relsPathOf(notes));
      }
    }
    for (const path of doomed) {
      const file = zip.file(path);
      if (!file) continue;
      removedFiles.set(path, await file.async("uint8array"));
      zip.remove(path);
      const override = overrides.get(`/${path}`);
      if (override) {
        override.parentNode?.removeChild(override);
        typesChanged = true;
      }
    }
  }

  for (const node of Array.from(list.children)) list.removeChild(node);
  for (const node of finalNodes) list.appendChild(node);
  updateSections(presentation.doc, finalIds);

  const serializer = new XMLSerializer();
  originals.set(PRESENTATION, presentation.xml);
  zip.file(PRESENTATION, serializer.serializeToString(presentation.doc));
  originals.set(PRESENTATION_RELS, relationships.xml);
  zip.file(PRESENTATION_RELS, serializer.serializeToString(relationships.doc));
  if (contentTypes && typesChanged) {
    if (!originals.has(CONTENT_TYPES)) originals.set(CONTENT_TYPES, contentTypes.xml);
    zip.file(CONTENT_TYPES, serializer.serializeToString(contentTypes.doc));
  }
}

async function writeRelationships(zip: JSZip, part: string, adds: RelAdd[], originals: Map<string, string>): Promise<void> {
  const path = relsPathOf(part);
  const loaded = await readDocument(zip, path);
  if (!loaded) return;
  originals.set(path, loaded.xml);
  const root = loaded.doc.documentElement;
  for (const add of adds) {
    const rel = loaded.doc.createElementNS(REL_NS, "Relationship");
    rel.setAttribute("Id", add.id);
    rel.setAttribute("Type", add.type);
    rel.setAttribute("Target", add.target);
    if (add.targetMode) rel.setAttribute("TargetMode", add.targetMode);
    root.appendChild(rel);
  }
  zip.file(path, new XMLSerializer().serializeToString(loaded.doc));
}

/**
 * Write the edited deck back into the original package. Only slide parts
 * that were touched are rewritten, and only the targeted shape nodes inside
 * them change, so everything the engine cannot edit yet (charts, animations,
 * masters, media) is carried over byte-for-byte.
 */
export async function exportEditedPptx(
  zip: JSZip,
  slides: Slide[],
  dirty: DirtyShapes,
  textEdits: TextEdits = new Map(),
  styleEdits: StyleEdits = new Map(),
  structure: StructureEdits = emptyStructure()
): Promise<Blob> {
  const serializer = new XMLSerializer();
  const edits: EditSet = { dirty, textEdits, styleEdits, structure };
  // Edited parts are put back once the archive is built, so a later export
  // (after an undo, say) starts again from the original package.
  const originals = new Map<string, string>();
  const addedFiles: string[] = [];
  const removedFiles = new Map<string, Uint8Array>();
  // Only slides that are still in the deck are written (and sorted out below).
  const listed = slides.length > 0 && slides.every(slide => !!slide.part);
  const livingParts = new Set(slides.map(slide => slide.part));

  try {
    for (const [path, xml] of structure.virtual) {
      zip.file(path, xml);
      addedFiles.push(path);
    }
    const parts = new Set([
      ...dirty.keys(), ...textEdits.keys(), ...styleEdits.keys(),
      ...structure.orders.keys(), ...structure.inserts.keys(), ...structure.relAdds.keys()
    ]);
    for (const part of parts) {
      if (listed && !livingParts.has(part)) continue;
      const loaded = await readDocument(zip, part);
      if (!loaded) continue;
      originals.set(part, loaded.xml);
      applyPartEdits(loaded.doc, part, slides, edits);

      let xml = serializer.serializeToString(loaded.doc);
      if (!xml.startsWith("<?xml")) {
        xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
      }
      zip.file(part, xml);

      const adds = structure.relAdds.get(part);
      if (adds && adds.length > 0) await writeRelationships(zip, part, adds, originals);
    }

    if (structure.media.length > 0) await writeMedia(zip, structure.media, originals, addedFiles);
    if (listed) await syncSlideList(zip, slides, originals, addedFiles, removedFiles);

    return await zip.generateAsync({
      type: "blob",
      mimeType: PPTX_MIME,
      compression: "DEFLATE"
    });
  } finally {
    for (const [path, xml] of originals) zip.file(path, xml);
    for (const [path, bytes] of removedFiles) zip.file(path, bytes);
    for (const path of addedFiles) zip.remove(path);
  }
}
