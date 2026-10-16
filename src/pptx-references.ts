const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** Ids (p:cNvPr@id) of a shape node and everything inside it. */
export function shapeIdsOf(shape: Element): string[] {
  return Array.from(shape.getElementsByTagNameNS(P_NS, "cNvPr")).map(node => node.getAttribute("id") || "");
}

const elementChildren = (node: Element) => Array.from(node.children);

/**
 * After shapes are removed, nothing in the slide may still point at them.
 * PowerPoint offers to "repair" a file whose animations target a missing shape,
 * so the effect is dropped (as PowerPoint itself does when the shape is
 * deleted), and connector ends that were glued to it are let go.
 */
export function pruneDeletedReferences(doc: Document, deleted: Set<string>): void {
  if (deleted.size === 0) return;

  for (const tag of ["stCxn", "endCxn"]) {
    for (const node of Array.from(doc.getElementsByTagNameNS(A_NS, tag))) {
      if (deleted.has(node.getAttribute("id") || "")) node.parentNode?.removeChild(node);
    }
  }

  const timing = doc.getElementsByTagNameNS(P_NS, "timing")[0];
  if (!timing) return;

  for (const target of Array.from(timing.getElementsByTagNameNS(P_NS, "spTgt"))) {
    if (!deleted.has(target.getAttribute("spid") || "") || !target.parentNode) continue;
    // The effect is the p:par whose p:cTn carries a preset id.
    let cursor: Element | null = target;
    while (cursor && !(cursor.localName === "cTn" && cursor.hasAttribute("presetID"))) cursor = cursor.parentElement;
    let node: Element | null = cursor?.parentElement?.localName === "par" ? cursor.parentElement : null;
    if (!node) continue;

    // Remove it, then any wrapper that ends up with nothing inside.
    while (node) {
      const list: Element | null = node.parentElement;
      node.parentNode?.removeChild(node);
      if (!list || elementChildren(list).length > 0) break;
      const owner: Element | null = list.parentElement;
      if (!owner || owner.getAttribute("nodeType") === "mainSeq") break;
      node = owner.parentElement?.localName === "par" ? owner.parentElement : null;
    }
  }

  // Targets outside plain effects (a video's play controls, an interactive
  // sequence started by clicking the shape) cannot be cut out safely: the
  // slide's animation tree goes, which is also what PowerPoint's repair does.
  const stillPointing = Array.from(timing.getElementsByTagNameNS(P_NS, "spTgt"))
    .some(node => deleted.has(node.getAttribute("spid") || ""));
  if (stillPointing) {
    timing.parentNode?.removeChild(timing);
    return;
  }

  for (const entry of Array.from(timing.getElementsByTagNameNS(P_NS, "bldP"))) {
    if (deleted.has(entry.getAttribute("spid") || "")) entry.parentNode?.removeChild(entry);
  }
  for (const list of Array.from(timing.getElementsByTagNameNS(P_NS, "bldLst"))) {
    if (elementChildren(list).length === 0) list.parentNode?.removeChild(list);
  }
  // No animation left at all: the empty timing tree goes too.
  if (timing.getElementsByTagNameNS(P_NS, "spTgt").length === 0 && timing.getElementsByTagNameNS(P_NS, "cTn").length > 0) {
    const hasEffect = Array.from(timing.getElementsByTagNameNS(P_NS, "cTn")).some(node => node.hasAttribute("presetID"));
    if (!hasEffect) timing.parentNode?.removeChild(timing);
  }
}
