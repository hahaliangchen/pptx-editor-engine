import { DeckParts, box, para, ph, slidePh } from "./deck";

const lvl = (attrs: string, inner: string) => `<a:lvl1pPr ${attrs}>${inner}</a:lvl1pPr>`;
const tx1 = '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>';

const sp = (id: number, name: string, x: number, y: number, w: number, h: number, geometry: string, text = "", paint = "") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
  + `<p:spPr>${box(x, y, w, h)}<a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>${paint}</p:spPr>`
  + (text ? `<p:txBody><a:bodyPr/><a:lstStyle/>${para(text)}</p:txBody>` : "")
  + `</p:sp>`;

const card = sp(10, "卡片", 914400, 1828800, 3657600, 1371600, "roundRect", "增长 30%",
  `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="2F528F"/></a:solidFill></a:ln>`);

/** A text shape with two runs, the first bold and red. */
const mixed = `<p:sp><p:nvSpPr><p:cNvPr id="11" name="双格式"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${box(914400, 5029200, 5486400, 457200)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
  + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>重点：</a:t></a:r><a:r><a:rPr lang="zh-CN"/><a:t>营收增长三成</a:t></a:r></a:p></p:txBody></p:sp>`;

const member = (id: number, name: string, x: number) => sp(id, name, x, 3657600, 1371600, 457200, "rect", "收入");
const group =
  `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="数据组"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
  + `<p:grpSpPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="3200400" cy="457200"/><a:chOff x="914400" y="3657600"/><a:chExt cx="3200400" cy="457200"/></a:xfrm></p:grpSpPr>`
  + member(21, "指标甲", 914400) + member(22, "指标乙", 2743200) + `</p:grpSp>`;

const line = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="30" name="连接线"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${box(914400, 4572000, 3000000, 0)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>`;

/** One slide: title placeholder, a filled card, a two-run text box, a group of two, a line. */
export const apiDeckParts = (): DeckParts => ({
  txStyles: `<p:titleStyle>${lvl('algn="ctr"', `<a:defRPr sz="3200" b="1">${tx1}</a:defRPr>`)}</p:titleStyle><p:bodyStyle/>`,
  masterShapes: ph("title", null, "", box(457200, 228600, 8229600, 857250)),
  layoutShapes: ph("title", null),
  slideShapes: slidePh("title", null, para("市场分析")).replace('name="s"', 'name="标题 1"') + card + mixed + group + line
});
