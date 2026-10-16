import JSZip from "jszip";

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");

function resolve(baseDir: string, target: string): string {
  const parts = (target.startsWith("/") ? target.slice(1) : baseDir ? `${baseDir}/${target}` : target).split("/");
  const out: string[] = [];
  for (const part of parts) {
    if (part === "..") out.pop();
    else if (part !== ".") out.push(part);
  }
  return out.join("/");
}

/**
 * Checks a .pptx for the structural mistakes that make PowerPoint offer to
 * repair a file but that a schema validator does not see: dangling
 * relationships, undeclared content types, duplicate ids and the like.
 * Returns a list of problems; empty means the package is consistent.
 */
export async function checkPackage(bytes: ArrayBuffer | Uint8Array | Buffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes);
  const problems: string[] = [];
  const names = Object.keys(zip.files).filter(name => !zip.files[name].dir);
  const text = async (name: string) => zip.file(name)!.async("string");

  // Every XML part is well-formed.
  for (const name of names.filter(n => /\.(xml|rels)$/.test(n))) {
    if (parse(await text(name)).getElementsByTagName("parsererror").length > 0) problems.push(`not well-formed: ${name}`);
  }

  // Every part has a content type.
  const types = parse(await text("[Content_Types].xml"));
  const defaults = new Set(Array.from(types.getElementsByTagNameNS(CT_NS, "Default")).map(n => (n.getAttribute("Extension") || "").toLowerCase()));
  const overrides = new Set(Array.from(types.getElementsByTagNameNS(CT_NS, "Override")).map(n => n.getAttribute("PartName") || ""));
  for (const name of names) {
    if (name === "[Content_Types].xml") continue;
    const ext = name.split(".").pop()!.toLowerCase();
    if (!defaults.has(ext) && !overrides.has(`/${name}`)) problems.push(`no content type for ${name}`);
  }
  for (const part of overrides) if (!zip.file(part.slice(1))) problems.push(`content type override for a missing part: ${part}`);

  // Relationships: unique ids, internal targets exist, attributes in the part point at them.
  const relIds = new Map<string, Set<string>>();
  for (const rels of names.filter(n => n.endsWith(".rels"))) {
    const owner = rels.replace("_rels/", "").replace(/\.rels$/, "");
    const dir = owner.includes("/") ? owner.slice(0, owner.lastIndexOf("/")) : "";
    const ids = new Set<string>();
    for (const rel of Array.from(parse(await text(rels)).getElementsByTagNameNS(REL_NS, "Relationship"))) {
      const id = rel.getAttribute("Id") || "";
      if (ids.has(id)) problems.push(`duplicate relationship id ${id} in ${rels}`);
      ids.add(id);
      if (rel.getAttribute("TargetMode") === "External") continue;
      const target = resolve(dir, rel.getAttribute("Target") || "");
      if (!zip.file(target)) problems.push(`${rels}: ${id} points at missing ${target}`);
    }
    relIds.set(owner, ids);
  }

  const slides = names.filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  for (const slide of slides) {
    const doc = parse(await text(slide));
    const ids = Array.from(doc.getElementsByTagNameNS(P_NS, "cNvPr")).map(n => n.getAttribute("id") || "");
    for (const id of new Set(ids)) if (ids.filter(x => x === id).length > 1) problems.push(`${slide}: shape id ${id} used twice`);
    const known = relIds.get(slide) || new Set();
    for (const node of Array.from(doc.getElementsByTagName("*"))) {
      for (const attribute of Array.from(node.attributes)) {
        // An empty id is how PowerPoint writes "no target" (e.g. an empty hyperlink).
        if (attribute.namespaceURI === R_NS && attribute.value && !known.has(attribute.value)) {
          problems.push(`${slide}: ${attribute.name}="${attribute.value}" has no relationship`);
        }
      }
    }
  }

  // The slide list: unique ids, resolvable, every slide part listed once.
  const presentation = parse(await text("ppt/presentation.xml"));
  const presRels = parse(await text("ppt/_rels/presentation.xml.rels"));
  const targets = new Map(Array.from(presRels.getElementsByTagNameNS(REL_NS, "Relationship")).map(n =>
    [n.getAttribute("Id") || "", resolve("ppt", n.getAttribute("Target") || "")]));
  const listed: string[] = [];
  const slideIds = new Set<string>();
  for (const entry of Array.from(presentation.getElementsByTagNameNS(P_NS, "sldId"))) {
    const id = entry.getAttribute("id") || "";
    if (slideIds.has(id)) problems.push(`slide id ${id} listed twice`);
    slideIds.add(id);
    if (parseInt(id, 10) < 256) problems.push(`slide id ${id} below 256`);
    const part = targets.get(entry.getAttributeNS(R_NS, "id") || "");
    if (!part || !zip.file(part)) problems.push(`slide entry ${id} points at ${part ?? "nothing"}`);
    else listed.push(part);
  }
  for (const part of new Set(listed)) if (listed.filter(x => x === part).length > 1) problems.push(`${part} listed twice`);
  for (const slide of slides) if (!listed.includes(slide)) problems.push(`orphan slide part ${slide}`);

  // Notes pages belong to a slide that still exists.
  for (const notes of names.filter(n => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n))) {
    const rels = `ppt/notesSlides/_rels/${notes.split("/").pop()}.rels`;
    const owner = zip.file(rels)
      ? Array.from(parse(await text(rels)).getElementsByTagNameNS(REL_NS, "Relationship"))
        .find(n => (n.getAttribute("Type") || "").endsWith("/slide"))
      : null;
    const slide = owner ? resolve("ppt/notesSlides", owner.getAttribute("Target") || "") : null;
    if (!slide || !listed.includes(slide)) problems.push(`orphan notes page ${notes}`);
  }
  return problems;
}
