import { RelAdd } from "../pptx-exporter";

const P_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A slide with a layout and nothing on it. */
export const blankSlideXml = (): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld ${P_NS}><p:cSld><p:spTree>`
  + `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
  + `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`
  + `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

const escapeAttr = (value: string) => value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

export const relationshipsXml = (rels: RelAdd[]): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL_NS}">`
  + rels.map(rel =>
    `<Relationship Id="${escapeAttr(rel.id)}" Type="${escapeAttr(rel.type)}" Target="${escapeAttr(rel.target)}"`
    + `${rel.targetMode ? ` TargetMode="${escapeAttr(rel.targetMode)}"` : ""}/>`).join("")
  + `</Relationships>`;

/** Relationships whose targets belong to one slide only: dropped from a copy. */
const DROPPED = ["notesSlide", "tags", "comments", "slideUpdateInfo"];
/** Targets that a copy can share with the original. */
const SHARED = ["slideLayout", "image", "hyperlink", "audio", "video", "media"];

export interface SlideCopy {
  xml: string;
  relationships: RelAdd[];
}

/**
 * Turn a slide snapshot into the base of a duplicate. Parts that belong to a
 * single slide (its notes, tags) are left behind; pictures and the layout are
 * shared. Anything else (charts, diagrams, embedded objects) would end up
 * owned by two slides, so the copy is refused.
 */
export function prepareSlideCopy(snapshot: SlideCopy): SlideCopy {
  const kept: RelAdd[] = [];
  for (const rel of snapshot.relationships) {
    const kind = rel.type.replace(`${REL}/`, "").replace(/^.*\//, "");
    if (DROPPED.includes(kind)) continue;
    if (!SHARED.includes(kind)) {
      throw new Error("这一页包含图表、SmartArt 或嵌入对象，暂不支持复制。");
    }
    kept.push(rel);
  }

  // Tag data referenced the relationships that were just left behind.
  const doc = new DOMParser().parseFromString(snapshot.xml, "application/xml");
  for (const node of Array.from(doc.getElementsByTagName("*"))) {
    if (node.localName === "custDataLst") node.parentNode?.removeChild(node);
  }
  return { xml: new XMLSerializer().serializeToString(doc), relationships: kept };
}
