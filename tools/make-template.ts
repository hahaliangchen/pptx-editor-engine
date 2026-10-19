/**
 * Builds tools/templates/blank.pptx: a 16:9 deck with a theme, a master and the
 * four layouts people expect (title, title and content, title only, blank), and
 * one empty title slide. A starting point for generating decks, and a deck
 * for trying the tools on.
 *
 *   npx vite-node tools/make-template.ts
 */
import { mkdirSync, writeFileSync } from "node:fs";
import JSZip from "jszip";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;
const REL = (id: string, type: string, target: string) =>
  `<Relationship Id="${id}" Type="${NS_R}/${type}" Target="${target}"/>`;
const RELS = (items: string[]) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.join("")}</Relationships>`;
const HEAD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

const W = 12192000;
const H = 6858000;

const clr = (name: string, val: string) => `<a:${name}><a:srgbClr val="${val}"/></a:${name}>`;
const theme = `${HEAD}<a:theme xmlns:a="${NS_A}" name="Blank"><a:themeElements>`
  + `<a:clrScheme name="Blue">${clr("dk1", "1F1F1F")}${clr("lt1", "FFFFFF")}${clr("dk2", "1F3A5F")}${clr("lt2", "EEF2F7")}`
  + `${clr("accent1", "2F6FDB")}${clr("accent2", "F28C28")}${clr("accent3", "3BA776")}${clr("accent4", "8E5CD9")}${clr("accent5", "E5484D")}${clr("accent6", "5B6B7F")}`
  + `${clr("hlink", "2F6FDB")}${clr("folHlink", "8E5CD9")}</a:clrScheme>`
  + `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface=""/></a:majorFont>`
  + `<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface=""/></a:minorFont></a:fontScheme>`
  + `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`
  + `<a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst>`
  + `<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="28575"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>`
  + `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>`
  + `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>`
  + `</a:fmtScheme></a:themeElements></a:theme>`;

const level = (n: number, size: number, marL: number, bullet: boolean) =>
  `<a:lvl${n}pPr marL="${marL}" indent="${bullet ? -285750 : 0}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`
  + `<a:lnSpc><a:spcPct val="110000"/></a:lnSpc><a:spcBef><a:spcPts val="${bullet ? 900 : 0}"/></a:spcBef>`
  + (bullet ? `<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/>` : "<a:buNone/>")
  + `<a:defRPr sz="${size}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${n}pPr>`;

const placeholder = (id: number, name: string, ph: string, box: [number, number, number, number], extra = "", text = "") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${ph}/></p:nvPr></p:nvSpPr>`
  + `<p:spPr>${box[2] ? `<a:xfrm><a:off x="${box[0]}" y="${box[1]}"/><a:ext cx="${box[2]}" cy="${box[3]}"/></a:xfrm>` : ""}</p:spPr>`
  + `<p:txBody><a:bodyPr${extra}/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>`;
void placeholder;

const tree = (shapes: string) =>
  `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
  + `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}</p:spTree></p:cSld>`;

const sp = (id: number, name: string, ph: string, box: [number, number, number, number] | null, body = "<a:bodyPr/>", lst = "") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${ph}/></p:nvPr></p:nvSpPr>`
  + `<p:spPr>${box ? `<a:xfrm><a:off x="${box[0]}" y="${box[1]}"/><a:ext cx="${box[2]}" cy="${box[3]}"/></a:xfrm>` : ""}</p:spPr>`
  + `<p:txBody>${body}<a:lstStyle>${lst}</a:lstStyle><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>`;

const MARGIN = 762000;
const master = `${HEAD}<p:sldMaster ${NS}>${tree(
  sp(2, "标题占位符 1", 'type="title"', [MARGIN, 457200, W - 2 * MARGIN, 1066800], '<a:bodyPr vert="horz" lIns="0" tIns="45720" rIns="0" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>')
  + sp(3, "文本占位符 2", 'type="body" idx="1"', [MARGIN, 1752600, W - 2 * MARGIN, 4343400], '<a:bodyPr vert="horz" lIns="0" tIns="45720" rIns="0" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr>')
)}`
  + `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`
  + `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/><p:sldLayoutId id="2147483651" r:id="rId3"/><p:sldLayoutId id="2147483652" r:id="rId4"/></p:sldLayoutIdLst>`
  + `<p:txStyles><p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>`
  + `<a:defRPr sz="3600" b="1" kern="1200"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>`
  + `<p:bodyStyle>${level(1, 2400, 285750, true)}${level(2, 2000, 742950, true)}${level(3, 1800, 1200150, true)}</p:bodyStyle>`
  + `<p:otherStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr>${level(1, 1800, 0, false)}</p:otherStyle></p:txStyles></p:sldMaster>`;

const layout = (name: string, type: string, shapes: string) =>
  `${HEAD}<p:sldLayout ${NS} type="${type}" preserve="1">${tree(shapes).replace("<p:cSld>", `<p:cSld name="${name}">`)}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const layouts = [
  layout("标题幻灯片", "title",
    sp(2, "标题 1", 'type="ctrTitle"', [MARGIN, 2057400, W - 2 * MARGIN, 1371600],
      '<a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr>', '<a:lvl1pPr algn="l"><a:defRPr sz="4800" b="1"/></a:lvl1pPr>')
    + sp(3, "副标题 2", 'type="subTitle" idx="1"', [MARGIN, 3657600, W - 2 * MARGIN, 1066800],
      "<a:bodyPr><a:normAutofit/></a:bodyPr>", '<a:lvl1pPr marL="0" indent="0" algn="l"><a:buNone/><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></a:defRPr></a:lvl1pPr>')),
  layout("标题和内容", "obj",
    sp(2, "标题 1", 'type="title"', null) + sp(3, "内容占位符 2", "idx=\"1\"", null)),
  layout("仅标题", "titleOnly", sp(2, "标题 1", 'type="title"', null)),
  layout("空白", "blank", "")
];

const slide1 = `${HEAD}<p:sld ${NS}>${tree(
  sp(2, "标题 1", 'type="ctrTitle"', null) + sp(3, "副标题 2", 'type="subTitle" idx="1"', null)
)}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
  + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
  + `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`
  + `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`
  + layouts.map((_, i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`).join("")
  + `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  + `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`);
zip.file("_rels/.rels", RELS([REL("rId1", "officeDocument", "ppt/presentation.xml")]));
zip.file("ppt/presentation.xml", `${HEAD}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>`
  + `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="${W}" cy="${H}"/><p:notesSz cx="6858000" cy="9144000"/>`
  + `<p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr></p:defaultTextStyle></p:presentation>`);
zip.file("ppt/_rels/presentation.xml.rels", RELS([REL("rId1", "slideMaster", "slideMasters/slideMaster1.xml"), REL("rId2", "slide", "slides/slide1.xml"), REL("rId3", "theme", "theme/theme1.xml")]));
zip.file("ppt/theme/theme1.xml", theme);
zip.file("ppt/slideMasters/slideMaster1.xml", master);
zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", RELS([
  ...layouts.map((_, i) => REL(`rId${i + 1}`, "slideLayout", `../slideLayouts/slideLayout${i + 1}.xml`)),
  REL("rId5", "theme", "../theme/theme1.xml")
]));
layouts.forEach((xml, i) => {
  zip.file(`ppt/slideLayouts/slideLayout${i + 1}.xml`, xml);
  zip.file(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, RELS([REL("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")]));
});
zip.file("ppt/slides/slide1.xml", slide1);
zip.file("ppt/slides/_rels/slide1.xml.rels", RELS([REL("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")]));

mkdirSync("tools/templates", { recursive: true });
writeFileSync("tools/templates/blank.pptx", await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
console.log("tools/templates/blank.pptx");
