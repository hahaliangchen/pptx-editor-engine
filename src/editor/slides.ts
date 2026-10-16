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
/** Targets a copy must own: the caller gives each its own copy of the part. */
const CLONED = ["chart"];

export interface SlideCopy {
  xml: string;
  relationships: RelAdd[];
}

/**
 * Turn a slide snapshot into the base of a duplicate. Parts that belong to a
 * single slide (its notes, tags) are left behind; pictures and the layout are
 * shared; charts stay in the list for the caller to give each a part of its own.
 * Anything else (diagrams, embedded objects) would end up owned by two slides,
 * so the copy is refused.
 */
export function prepareSlideCopy(snapshot: SlideCopy): SlideCopy {
  const kept: RelAdd[] = [];
  for (const rel of snapshot.relationships) {
    const kind = rel.type.replace(`${REL}/`, "").replace(/^.*\//, "");
    if (DROPPED.includes(kind)) continue;
    if (!SHARED.includes(kind) && !CLONED.includes(kind)) {
      throw new Error("这一页包含 SmartArt 或嵌入对象，暂不支持复制。");
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

export interface PlaceholderSpec {
  type?: string;
  idx?: string;
  orient?: string;
  sz?: string;
}

/** Placeholders that carry text; footers, dates and slide numbers are not created on a new slide. */
const TEXT_PLACEHOLDERS = new Set(["title", "ctrTitle", "subTitle", "body", "obj"]);

/** The placeholders a layout offers, as a new slide would instantiate them. */
export function layoutPlaceholders(layoutXml: string): PlaceholderSpec[] {
  const doc = new DOMParser().parseFromString(layoutXml, "application/xml");
  const specs: PlaceholderSpec[] = [];
  for (const node of Array.from(doc.getElementsByTagName("*"))) {
    if (node.localName !== "ph") continue;
    const type = node.getAttribute("type") || undefined;
    // A placeholder without a type is the layout's content placeholder.
    if (type !== undefined && !TEXT_PLACEHOLDERS.has(type)) continue;
    specs.push({
      type,
      idx: node.getAttribute("idx") || undefined,
      orient: node.getAttribute("orient") || undefined,
      sz: node.getAttribute("sz") || undefined
    });
  }
  return specs;
}

const PLACEHOLDER_NAMES: Record<string, string> = {
  title: "标题",
  ctrTitle: "标题",
  subTitle: "副标题",
  body: "文本占位符",
  obj: "内容占位符"
};

/** A slide holding one empty placeholder per layout placeholder; positions come from the layout. */
export function placeholderSlideXml(specs: PlaceholderSpec[]): string {
  const shapes = specs.map((spec, index) => {
    const id = index + 2;
    const name = `${PLACEHOLDER_NAMES[spec.type ?? "obj"] ?? "占位符"} ${id}`;
    const attrs = [
      spec.type ? ` type="${escapeAttr(spec.type)}"` : "",
      spec.orient ? ` orient="${escapeAttr(spec.orient)}"` : "",
      spec.sz ? ` sz="${escapeAttr(spec.sz)}"` : "",
      spec.idx ? ` idx="${escapeAttr(spec.idx)}"` : ""
    ].join("");
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>`
      + `<p:nvPr><p:ph${attrs}/></p:nvPr></p:nvSpPr><p:spPr/>`
      + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN" altLang="en-US"/></a:p></p:txBody></p:sp>`;
  }).join("");
  return blankSlideXml().replace("</p:spTree>", `${shapes}</p:spTree>`);
}

/** The hint PowerPoint shows inside an empty placeholder. */
export function placeholderPrompt(type: string): string {
  if (type === "title" || type === "ctrTitle") return "单击此处添加标题";
  if (type === "subTitle") return "单击此处添加副标题";
  return "单击此处添加文本";
}
