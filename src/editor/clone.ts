const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Shape ids found in a serialized shape node, in document order. */
export function shapeIdsIn(xml: string): string[] {
  const root = new DOMParser().parseFromString(xml, "application/xml").documentElement;
  return Array.from(root.getElementsByTagName("*"))
    .filter(node => node.localName === "cNvPr")
    .map(node => node.getAttribute("id") || "");
}

/**
 * Prepare a copied shape node for insertion: give every shape inside it a
 * fresh id, point relationship attributes at the target slide's ids, and drop
 * connector attachments (they name shapes that are not in the copy).
 */
export function remapShapeXml(
  xml: string,
  allocateId: () => string,
  relIds: Map<string, string> = new Map()
): { xml: string; idMap: Map<string, string> } {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  const idMap = new Map<string, string>();

  const all = [root, ...Array.from(root.getElementsByTagName("*"))];
  for (const node of all) {
    if (node.localName === "cNvPr") {
      const oldId = node.getAttribute("id") || "";
      const fresh = allocateId();
      idMap.set(oldId, fresh);
      node.setAttribute("id", fresh);
    }
    for (const attribute of Array.from(node.attributes)) {
      if (attribute.namespaceURI === R_NS && relIds.has(attribute.value)) {
        attribute.value = relIds.get(attribute.value)!;
      }
    }
  }
  for (const node of Array.from(root.getElementsByTagName("*"))) {
    if (node.localName === "stCxn" || node.localName === "endCxn") node.parentNode?.removeChild(node);
  }
  return { xml: new XMLSerializer().serializeToString(root), idMap };
}
