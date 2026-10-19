import { describe, expect, it } from "vitest";
import { PptxParser } from "../src/pptx-parser";
import { PptVirtualDocument, ShapeElement, TextElement } from "../src/pptx-virtual-dom";
import { DeckParts, box, para } from "./helpers/deck";
import { buildDeck } from "./helpers/deck";

async function parse(parts: DeckParts): Promise<PptVirtualDocument> {
  return new PptxParser().parse((await buildDeck(parts)) as unknown as ArrayBuffer);
}
const backgroundOf = (doc: PptVirtualDocument) =>
  doc.slides[0].elements.find((element): element is ShapeElement => element.type === "shape" && element.id === "slide_background");
const SOLID = (color: string) => `<p:bgPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:effectLst/></p:bgPr>`;
const REF = (color: string) => `<p:bgRef idx="1001"><a:schemeClr val="${color}"/></p:bgRef>`;

describe("page backgrounds", () => {
  it("has none to draw when no level names one", async () => {
    expect(backgroundOf(await parse({}))).toBeUndefined();
  });

  it("covers the whole page with the master's colour", async () => {
    const doc = await parse({ masterBackground: SOLID("102030") });
    const background = backgroundOf(doc)!;
    expect(background.rect).toEqual({ x: 0, y: 0, w: doc.size.width, h: doc.size.height });
    expect(background.computedStyle!.fill).toEqual({ type: "solid", color: "#102030" });
    expect(doc.slides[0].elements[0]).toBe(background);
  });

  it("takes the slide's own over the layout's, and the layout's over the master's", async () => {
    const colorOf = async (parts: DeckParts) => (backgroundOf(await parse(parts))!.computedStyle!.fill as { color: string }).color;
    expect(await colorOf({ masterBackground: SOLID("111111"), layoutBackground: SOLID("222222") })).toBe("#222222");
    expect(await colorOf({ masterBackground: SOLID("111111"), layoutBackground: SOLID("222222"), slideBackground: SOLID("333333") })).toBe("#333333");
  });

  it("reads a gradient background", async () => {
    const doc = await parse({
      masterBackground: `<p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></p:bgPr>`
    });
    expect(backgroundOf(doc)!.computedStyle!.fill).toMatchObject({ type: "gradient", kind: "linear", angle: 90 });
  });

  it("reads a background that refers to the theme's background styles", async () => {
    // The fixture's theme: lt1 FFFFFF, dk2 44546A.
    const doc = await parse({ masterBackground: REF("tx2") });
    expect(backgroundOf(doc)!.computedStyle!.fill).toEqual({ type: "solid", color: "#44546a" });
  });
});

describe("the master's colour map", () => {
  const dark = { bg1: "dk1", tx1: "lt1", bg2: "dk2", tx2: "lt2" };
  const textColorOf = (doc: PptVirtualDocument) =>
    (doc.slides[0].elements.find((element): element is TextElement => element.type === "text")!).style.color.toLowerCase();
  const slideText = (scheme: string) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="4" name="文字"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${box(0, 0, 3000000, 500000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
    + `<p:txBody><a:bodyPr/><a:lstStyle/>${para("字", "", `<a:rPr lang="zh-CN"><a:solidFill><a:schemeClr val="${scheme}"/></a:solidFill></a:rPr>`)}</p:txBody></p:sp>`;

  it("makes bg1 the dark colour and tx1 the light one on a dark master", async () => {
    const doc = await parse({ clrMap: dark, masterBackground: REF("bg1"), slideShapes: slideText("tx1") });
    expect((backgroundOf(doc)!.computedStyle!.fill as { color: string }).color).toBe("#000000");
    expect(textColorOf(doc)).toBe("#ffffff");
  });

  it("leaves tx1 dark and bg1 light on an ordinary master", async () => {
    const doc = await parse({ masterBackground: REF("bg1"), slideShapes: slideText("tx1") });
    expect((backgroundOf(doc)!.computedStyle!.fill as { color: string }).color).toBe("#ffffff");
    expect(textColorOf(doc)).toBe("#000000");
  });

  it("lets a layout override the master's map", async () => {
    const doc = await parse({
      layoutClrMapOvr: '<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>',
      slideShapes: slideText("tx1")
    });
    expect(textColorOf(doc)).toBe("#ffffff");
  });

  it("does not leak one slide's map into the next parse", async () => {
    const parser = new PptxParser();
    const darkDoc = await parser.parse((await buildDeck({ clrMap: dark, slideShapes: slideText("tx1") })) as unknown as ArrayBuffer);
    const lightDoc = await parser.parse((await buildDeck({ slideShapes: slideText("tx1") })) as unknown as ArrayBuffer);
    expect(textColorOf(darkDoc)).toBe("#ffffff");
    expect(textColorOf(lightDoc)).toBe("#000000");
  });
});
