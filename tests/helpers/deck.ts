import JSZip from "jszip";
import { PptxParser } from "../../src/pptx-parser";
import { TextElement, TextParagraph } from "../../src/pptx-virtual-dom";

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = (items: Array<[string, string, string]>) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
  + items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join("")
  + `</Relationships>`;

/** One slide px = 9144000 / 1920 EMU; so 1 pt (12700 EMU) is this many px. */
export const PX_PER_PT = (12700 * 1920) / 9144000;

export interface DeckParts {
  /** Inside p:txStyles, e.g. "<p:titleStyle>...</p:titleStyle><p:bodyStyle>...</p:bodyStyle>". */
  txStyles?: string;
  /** Extra shapes on the master / layout / slide, as p:sp XML. */
  masterShapes?: string;
  layoutShapes?: string;
  slideShapes?: string;
  /** Children of p:defaultTextStyle in presentation.xml. */
  defaultTextStyle?: string;
}

export const ph = (type: string | null, idx: number | null, lstStyle = "", extra = "") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${idx ?? 99}" name="ph"/><p:cNvSpPr/><p:nvPr><p:ph${type ? ` type="${type}"` : ""}${idx !== null ? ` idx="${idx}"` : ""}/></p:nvPr></p:nvSpPr>`
  + `<p:spPr>${extra}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle>${lstStyle}</a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;

export const box = (x: number, y: number, w: number, h: number) =>
  `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;

/** A slide placeholder with its own paragraphs. */
export const slidePh = (type: string | null, idx: number | null, paragraphs: string, lstStyle = "") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${idx ?? 90}" name="s"/><p:cNvSpPr/><p:nvPr><p:ph${type ? ` type="${type}"` : ""}${idx !== null ? ` idx="${idx}"` : ""}/></p:nvPr></p:nvSpPr>`
  + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle>${lstStyle}</a:lstStyle>${paragraphs}</p:txBody></p:sp>`;

/** A paragraph; `pPr` is a full <a:pPr .../> or <a:pPr ...>...</a:pPr> string. */
export const para = (text: string, pPr = "", rPr = "") =>
  `<a:p>${pPr}<a:r>${rPr}<a:t>${text}</a:t></a:r></a:p>`;

export async function buildDeck(parts: DeckParts): Promise<Buffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`);
  zip.file("ppt/presentation.xml",
    `<p:presentation ${NS}><p:sldMasterIdLst/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>`
    + `<p:sldSz cx="9144000" cy="5143500"/>`
    + (parts.defaultTextStyle ? `<p:defaultTextStyle>${parts.defaultTextStyle}</p:defaultTextStyle>` : "")
    + `</p:presentation>`);
  zip.file("ppt/_rels/presentation.xml.rels", RELS([["rId1", "slideMaster", "slideMasters/slideMaster1.xml"], ["rId2", "slide", "slides/slide1.xml"], ["rId3", "theme", "theme/theme1.xml"]]));
  zip.file("ppt/theme/theme1.xml",
    `<a:theme ${NS} name="t"><a:themeElements><a:clrScheme name="c">`
    + `<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>`
    + `<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>`
    + `<a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>`
    + `<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>`
    + `<a:fontScheme name="f"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>`
    + `</a:themeElements></a:theme>`);
  zip.file("ppt/slideMasters/slideMaster1.xml",
    `<p:sldMaster ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${parts.masterShapes ?? ""}</p:spTree></p:cSld>`
    + `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`
    + `<p:txStyles>${parts.txStyles ?? ""}</p:txStyles></p:sldMaster>`);
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", RELS([["rId1", "theme", "../theme/theme1.xml"]]));
  zip.file("ppt/slideLayouts/slideLayout1.xml",
    `<p:sldLayout ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${parts.layoutShapes ?? ""}</p:spTree></p:cSld></p:sldLayout>`);
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", RELS([["rId1", "slideMaster", "../slideMasters/slideMaster1.xml"]]));
  zip.file("ppt/slides/slide1.xml",
    `<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${parts.slideShapes ?? ""}</p:spTree></p:cSld></p:sld>`);
  zip.file("ppt/slides/_rels/slide1.xml.rels", RELS([["rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"]]));
  return Buffer.from(await zip.generateAsync({ type: "uint8array" }));
}

/** Text elements of the first slide, in order, with their paragraphs. */
export async function parseTexts(parts: DeckParts): Promise<TextElement[]> {
  const parser = new PptxParser();
  const doc = await parser.parse((await buildDeck(parts)) as unknown as ArrayBuffer);
  return doc.slides[0].elements.filter((element): element is TextElement => element.type === "text" && !!element.source);
}

export const firstRun = (paragraph: TextParagraph) => paragraph.runs[0].style;
