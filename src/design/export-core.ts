import JSZip from "jszip";
import { writeTemplateRecord, readTemplateRecord, defaultPurpose, type TemplateRecord } from "./template-policy";
import { base64ToBytes, sha256Hex } from "./bytes";
import { assertDesign, DesignDocument, DesignNode, DesignText, sameAppearance } from "./model";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const direct = (node: Element, name: string) => Array.from(node.children).find(n => n.localName === name);
const all = (node: Element, name: string) => Array.from(node.getElementsByTagNameNS("*", name));
const x = (value: string) => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");
const stringify = (xml: Node) => new XMLSerializer().serializeToString(xml);
const fragment = (xml: string) => parse(`<root xmlns:a="${A}" xmlns:p="${P}" xmlns:r="${R}">${xml}</root>`).documentElement.firstElementChild!;
function ensure(parent: Element, name: string): Element {
  const existing = direct(parent, name.split(":").pop()!);
  if (existing) return existing;
  const child = parent.ownerDocument.createElementNS(name.startsWith("a:") ? A : P, name);
  parent.appendChild(child);
  return child;
}
function replaceChild(parent: Element, name: string, replacement: Element): void {
  const old = direct(parent, name);
  if (old) parent.replaceChild(parent.ownerDocument.importNode(replacement, true), old);
  else parent.appendChild(parent.ownerDocument.importNode(replacement, true));
}
function colorXml(color: string): string {
  let hex = color.replace(/^#/, ""), alpha = 1;
  if (hex.length === 3) hex = hex.split("").map(c => c + c).join("");
  const rgb = color.match(/^rgba?\(\s*([\d.]+)[, ]+([\d.]+)[, ]+([\d.]+)(?:\s*[,/]\s*([\d.]+))?\s*\)$/);
  if (rgb) { hex = rgb.slice(1, 4).map(n => Math.round(Number(n)).toString(16).padStart(2, "0")).join(""); alpha = rgb[4] === undefined ? 1 : Number(rgb[4]); }
  if (hex.length === 8) { alpha = parseInt(hex.slice(6), 16) / 255; hex = hex.slice(0, 6); }
  if (!/^[0-9a-f]{6}$/i.test(hex)) throw new Error(`Unsupported color: ${color}`);
  return `<a:srgbClr val="${hex.toUpperCase()}">${alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ""}</a:srgbClr>`;
}
function setFill(parent: Element, fill?: string): void {
  for (const child of Array.from(parent.children)) if (["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(child.localName)) child.remove();
  const paint = fragment(fill && fill !== "transparent" ? `<a:solidFill>${colorXml(fill)}</a:solidFill>` : "<a:noFill/>");
  const before = Array.from(parent.children).find(n => ["ln", "effectLst", "effectDag", "scene3d", "sp3d", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "extLst"].includes(n.localName));
  parent.insertBefore(parent.ownerDocument.importNode(paint, true), before ?? null);
}

function writeText(shape: Element, text: DesignText, emuPerCss: number): void {
  const tx = ensure(shape, "p:txBody");
  let body = direct(tx, "bodyPr");
  if (!body) { body = shape.ownerDocument.createElementNS(A, "a:bodyPr"); tx.insertBefore(body, tx.firstChild); }
  const pt = (pixels: number) => Math.round(pixels * emuPerCss / 127);
  const [top, right, bottom, left] = text.padding;
  for (const [key, value] of Object.entries({ tIns: top, rIns: right, bIns: bottom, lIns: left })) body.setAttribute(key, String(Math.round(value * emuPerCss)));
  body.setAttribute("anchor", ({ top: "t", middle: "ctr", bottom: "b" })[text.verticalAlign]);
  body.setAttribute("wrap", "none");
  body.setAttribute("vertOverflow", "overflow");
  for (const fit of Array.from(body.children).filter(n => ["spAutoFit", "normAutofit", "noAutofit"].includes(n.localName))) fit.remove();
  body.appendChild(fragment("<a:noAutofit/>"));
  ensure(tx, "a:lstStyle");
  const oldParagraphs = Array.from(tx.children).filter(n => n.localName === "p");
  for (const old of oldParagraphs) old.remove();
  text.paragraphs.forEach((paragraph, pIndex) => {
    const p = fragment("<a:p/>");
    const old = oldParagraphs[pIndex];
    const pPr = direct(old ?? p, "pPr")?.cloneNode(true) as Element | undefined ?? fragment("<a:pPr/>");
    pPr.setAttribute("algn", ({ left: "l", center: "ctr", right: "r" })[paragraph.align]);
    for (const key of ["lnSpc", "spcBef", "spcAft"]) direct(pPr, key)?.remove();
    const anchor = pPr.firstChild;
    for (const [key, value] of [["lnSpc", paragraph.lineHeight], ["spcBef", paragraph.before], ["spcAft", paragraph.after]] as const) pPr.insertBefore(fragment(`<a:${key}><a:spcPts val="${pt(value)}"/></a:${key}>`), anchor);
    p.appendChild(pPr);
    const oldRuns = old ? all(old, "r") : [];
    paragraph.runs.forEach((run, rIndex) => {
      const properties = direct(oldRuns[rIndex] ?? p, "rPr")?.cloneNode(true) as Element | undefined ?? fragment("<a:rPr/>");
      for (const [key, value] of Object.entries({ sz: pt(run.fontSize), b: run.bold ? 1 : 0, i: run.italic ? 1 : 0, spc: pt(run.letterSpacing) })) properties.setAttribute(key, String(value));
      setFill(properties, run.color);
      replaceChild(properties, "latin", fragment(`<a:latin typeface="${x(run.fontFamily)}"/>`));
      replaceChild(properties, "ea", fragment(`<a:ea typeface="${x(run.fontFamily)}"/>`));
      run.text.split(/\u2028|\n/).forEach((piece, index) => {
        if (index) { const br = fragment("<a:br/>"); br.appendChild(properties.cloneNode(true)); p.appendChild(br); }
        if (piece !== "") { const r = fragment(`<a:r><a:t xml:space="preserve">${x(piece)}</a:t></a:r>`); r.insertBefore(properties.cloneNode(true), r.firstChild); p.appendChild(r); }
      });
    });
    const end = fragment("<a:endParaRPr/>");
    if (paragraph.runs[0]) end.setAttribute("sz", String(pt(paragraph.runs[0].fontSize)));
    p.appendChild(end);
    tx.insertBefore(tx.ownerDocument.importNode(p, true), direct(tx, "extLst") ?? null);
  });
}

/** Materialize the canonical JSON onto its opaque source package; untouched XML parts pass through. */
export async function exportDesignBytes(design: DesignDocument, identities?: Map<string, string>): Promise<Uint8Array> {
  assertDesign(design);
  const source = base64ToBytes(design.source.pptxBase64);
  if (sha256Hex(source) !== design.source.sha256) throw new Error("Design source package checksum mismatch");
  const zip = await JSZip.loadAsync(source);
  const unit = design.source.widthEMU / 1280;
  let changed = false, mediaSerial = 0;
  for (const page of design.pages) {
    const file = zip.file(page.part);
    if (!file) throw new Error(`Missing original slide part: ${page.part}`);
    const xml = parse(await file.async("string"));
    const tree = xml.getElementsByTagNameNS(P, "spTree")[0];
    const owners = new Map<string, Element>();
    for (const props of all(tree, "cNvPr")) owners.set(props.getAttribute("id")!, props.parentElement!.parentElement!);
    let serial = Math.max(1, ...[...owners.keys()].map(Number).filter(Number.isFinite));
    const beforeXml = stringify(xml), hadRecord = !!readTemplateRecord(xml);
    if (page.background && page.background !== (page.originBackground ?? "#ffffff")) {
      const cSld = xml.getElementsByTagNameNS(P, "cSld")[0];
      direct(cSld, "bg")?.remove();
      cSld.insertBefore(fragment(`<p:bg><p:bgPr><a:solidFill>${colorXml(page.background)}</a:solidFill><a:effectLst/></p:bgPr></p:bg>`), cSld.firstChild);
    }
    const liveSourceIds = new Set(page.nodes.map(n => n.origin?.shapeId ?? n.metadata?.sourceShapeId).filter(Boolean));
    for (const [id, owner] of owners) if (owner !== tree && !liveSourceIds.has(id)) owner.remove();
    const nodeElements = new Map<string, Element>();
    const record: TemplateRecord = { version: 1, ...(page.semantics ? { page: page.semantics } : {}), nodes: {} };
    const transform = (parent: Element): { sx: number; sy: number; ox: number; oy: number } => {
      if (parent === tree) return { sx: 1 / unit, sy: 1 / unit, ox: 0, oy: 0 };
      const base = transform(parent.parentElement!);
      const xf = direct(direct(parent, "grpSpPr")!, "xfrm");
      const value = (tag: string, attr: string, fallback: number) => Number(xf && direct(xf, tag)?.getAttribute(attr) || fallback);
      const sx = value("ext", "cx", 1) / value("chExt", "cx", 1), sy = value("ext", "cy", 1) / value("chExt", "cy", 1);
      return { sx: base.sx * sx, sy: base.sy * sy, ox: base.ox + base.sx * (value("off", "x", 0) - value("chOff", "x", 0) * sx), oy: base.oy + base.sy * (value("off", "y", 0) - value("chOff", "y", 0) * sy) };
    };
    const image = async (node: DesignNode, owner: Element) => {
      const match = node.appearance.src?.match(/^data:image\/(png|jpeg|gif);base64,([\s\S]+)$/);
      if (!match) throw new Error(`Image must be embedded PNG/JPEG/GIF: ${node.id}`);
      let media: string;
      do { media = `ppt/media/design-${++mediaSerial}.${match[1]}`; } while (zip.file(media));
      zip.file(media, base64ToBytes(match[2]));
      const relPart = page.part.replace(/\/([^/]+)$/, "/_rels/$1.rels");
      const relFile = zip.file(relPart);
      const rels = parse(relFile ? await relFile.async("string") : `<Relationships xmlns="${REL}"/>`);
      let relSerial = 1;
      const taken = new Set(all(rels.documentElement, "Relationship").map(n => n.getAttribute("Id")));
      while (taken.has(`rIdDesign${relSerial}`)) relSerial++;
      const id = `rIdDesign${relSerial}`;
      const rel = rels.createElementNS(REL, "Relationship");
      rel.setAttribute("Id", id); rel.setAttribute("Type", `${R}/image`); rel.setAttribute("Target", `../media/${media.split("/").pop()}`);
      rels.documentElement.appendChild(rel); zip.file(relPart, stringify(rels));
      const fill = ensure(owner, "p:blipFill"), blip = ensure(fill, "a:blip");
      blip.setAttributeNS(R, "r:embed", id);
      if (!direct(fill, "stretch")) fill.appendChild(fragment("<a:stretch><a:fillRect/></a:stretch>"));
      const types = parse(await zip.file("[Content_Types].xml")!.async("string"));
      if (!all(types.documentElement, "Default").some(n => n.getAttribute("Extension") === match[1])) {
        const type = types.createElementNS(types.documentElement.namespaceURI, "Default");
        type.setAttribute("Extension", match[1]); type.setAttribute("ContentType", `image/${match[1]}`); types.documentElement.appendChild(type);
        zip.file("[Content_Types].xml", stringify(types));
      }
    };
    const emit = async (node: DesignNode, parent: Element, parentChanged = false, children?: DesignNode[]): Promise<void> => {
      if (node.metadata?.backdrop) return;
      const sourceId = node.origin?.shapeId ?? node.metadata?.sourceShapeId;
      if (node.appearance.text?.frames) {
        const a = node.appearance, text = a.text!;
        // Keep the source object's identity (including animation references) on
        // a native group; its background and independently laid-out paragraphs
        // become editable children at the browser's actual coordinates.
        if (sourceId) {
          owners.get(String(sourceId))?.remove();
          owners.set(String(sourceId), fragment(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${sourceId}" name="${x(node.name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:grpSp>`));
        }
        const parts: DesignNode[] = [];
        if (a.fill || a.border || a.radius) parts.push({id:`${node.id}:paint`,kind:"shape",name:`${node.name} background`,purpose:node.purpose,
          appearance:{rect:a.rect,fill:a.fill,border:a.border,radius:a.radius}});
        text.frames!.forEach((frame,i) => parts.push({id:`${node.id}:paragraph:${i}`,kind:"text",name:`${node.name} paragraph ${i+1}`,purpose:node.purpose,
          appearance:{rect:frame.rect,text:{paragraphs:[{...text.paragraphs[i],before:0,after:0}],padding:frame.padding,verticalAlign:"top"}}}));
        await emit({...node,kind:"group",appearance:{rect:a.rect}},parent,true,parts);
        return;
      }
      let owner = sourceId ? owners.get(String(sourceId)) : undefined;
      if (node.hidden) return;
      let fresh = !owner;
      if (owner && node.kind === "asset" && node.origin?.appearance.src !== node.appearance.src) {
        // Replacing a template reference visual replaces its source object, rather than adding a blipFill to a chart/shape.
        owner.remove();
        owner = fragment(`<p:pic><p:nvPicPr><p:cNvPr id="${sourceId}" name="${x(node.name)}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
        fresh = true;
      }
      if (!owner) {
        const id = String(++serial), nv = `<p:cNvPr id="${id}" name="${x(node.name)}"/>`;
        if (node.kind === "group") owner = fragment(`<p:grpSp><p:nvGrpSpPr>${nv}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${design.source.widthEMU}" cy="${design.source.heightEMU}"/><a:chOff x="0" y="0"/><a:chExt cx="${design.source.widthEMU}" cy="${design.source.heightEMU}"/></a:xfrm></p:grpSpPr></p:grpSp>`);
        else if (node.kind === "image" || node.kind === "asset") owner = fragment(`<p:pic><p:nvPicPr>${nv}<p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
        else owner = fragment(`<p:sp><p:nvSpPr>${nv}<p:cNvSpPr${node.kind === "text" ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr><p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`);
      }
      parent.appendChild(owner);
      nodeElements.set(node.id, owner);
      const physicalId = all(owner, "cNvPr")[0]?.getAttribute("id");
      if (physicalId) {
        identities?.set(`${page.part}#${physicalId}`, node.id);
        record.nodes[physicalId] = { purpose: node.purpose ?? defaultPurpose(node), ...(node.note ? { note: node.note } : {}), name: node.name };
      }
      if (node.kind === "group") {
        const changedGroup = fresh || parentChanged || !sameAppearance(node.origin?.appearance.rect, node.appearance.rect);
        if (changedGroup) {
          const t = transform(parent), r = node.appearance.rect;
          const w = Math.max(1, Math.round(r.w / t.sx)), h = Math.max(1, Math.round(r.h / t.sy));
          replaceChild(ensure(owner, "p:grpSpPr"), "xfrm", fragment(`<a:xfrm><a:off x="${Math.round((r.x - t.ox) / t.sx)}" y="${Math.round((r.y - t.oy) / t.sy)}"/><a:ext cx="${w}" cy="${h}"/><a:chOff x="0" y="0"/><a:chExt cx="${w}" cy="${h}"/></a:xfrm>`));
        }
        for (const child of children ?? page.nodes.filter(n => n.parentId === node.id)) await emit(child, owner, changedGroup);
        return;
      }
      const old = node.origin?.appearance, a = node.appearance;
      if (!fresh && !parentChanged && old && sameAppearance(old, a)) return;
      const props = ensure(owner, "p:spPr");
      if (fresh || parentChanged || !sameAppearance(old?.rect, a.rect)) {
        const t = transform(parent), r = a.rect;
        const xf = direct(props, "xfrm") ?? fragment("<a:xfrm/>");
        if (!xf.parentNode || xf.parentNode !== props) props.insertBefore(xf, props.firstChild);
        replaceChild(xf, "off", fragment(`<a:off x="${Math.round((r.x - t.ox) / t.sx)}" y="${Math.round((r.y - t.oy) / t.sy)}"/>`));
        replaceChild(xf, "ext", fragment(`<a:ext cx="${Math.round(r.w / t.sx)}" cy="${Math.round(r.h / t.sy)}"/>`));
      }
      if (node.kind === "image" || node.kind === "asset") { if (fresh || old?.src !== a.src) await image(node, owner); return; }
      if (fresh || old?.fill !== a.fill) setFill(props, a.fill);
      if (fresh || !sameAppearance(old?.border, a.border)) replaceChild(props, "ln", fragment(a.border ? `<a:ln w="${Math.round(a.border.width * unit)}"><a:solidFill>${colorXml(a.border.color)}</a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>"));
      if (fresh || old?.radius !== a.radius) {
        const adjustment = a.radius ? Math.min(50000, Math.round(a.radius / Math.max(1, Math.min(a.rect.w, a.rect.h)) * 100000)) : 0;
        replaceChild(props, "prstGeom", fragment(`<a:prstGeom prst="${adjustment ? "roundRect" : "rect"}"><a:avLst>${adjustment ? `<a:gd name="adj" fmla="val ${adjustment}"/>` : ""}</a:avLst></a:prstGeom>`));
      }
      if (a.text && (fresh || !sameAppearance(old?.text, a.text))) writeText(owner, a.text, unit);
      else if (!a.text && old?.text) direct(owner, "txBody")?.remove();
    };
    for (const node of page.nodes.filter(n => !n.parentId)) await emit(node, tree);
    if (page.semanticsChanged || hadRecord || stringify(xml) !== beforeXml) writeTemplateRecord(xml, record);
    if (stringify(xml) !== beforeXml) { changed = true; zip.file(page.part, stringify(xml)); }
  }
  return changed ? await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }) : source;
}
