// @vitest-environment node
import { beforeAll, describe, expect, it } from "vitest";
import { loadImage, createCanvas } from "@napi-rs/canvas";
import { HeadlessRenderer } from "../src/headless/renderer";
import { installNodeEnv } from "../src/headless/node-env";
import { PptxParser } from "../src/pptx-parser";
import { PptVirtualDocument } from "../src/pptx-virtual-dom";
import { box, buildDeck, para } from "./helpers/deck";
import { apiDeckParts } from "./helpers/api-deck";

installNodeEnv();

/** Decode a PNG and read one pixel as [r, g, b, a]. */
async function pixelAt(png: Buffer, x: number, y: number): Promise<number[]> {
  const image = await loadImage(png);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  return [...ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data];
}

describe("headless rendering", () => {
  let doc: PptVirtualDocument;
  let renderer: HeadlessRenderer;

  beforeAll(async () => {
    doc = await new PptxParser().parse((await buildDeck(apiDeckParts())) as unknown as ArrayBuffer);
    renderer = new HeadlessRenderer();
    renderer.prepare(doc);
  });

  it("produces a PNG with the requested width and the slide's aspect ratio", async () => {
    const png = await renderer.renderSlide(doc, 0, { width: 960 });
    expect([...png.subarray(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47]);
    const image = await loadImage(png);
    expect(image.width).toBe(960);
    expect(image.height).toBe(Math.round(960 * doc.size.height / doc.size.width));
  });

  it("draws shapes where the slide puts them, in their own colour, on a white page", async () => {
    const width = 960;
    const png = await renderer.renderSlide(doc, 0, { width });
    const pxPerEmu = width / (doc.size.width * doc.slides[0].emuPerPx);
    // Inside the blue card (x 914400.., y 1828800..), clear of the rounded corner and the text.
    const [r, g, b] = await pixelAt(png, (914400 + 400000) * pxPerEmu, (1828800 + 100000) * pxPerEmu);
    expect([r, g, b]).toEqual([0x44, 0x72, 0xc4]);
    expect(await pixelAt(png, width - 5, 5)).toEqual([255, 255, 255, 255]);
  });

  it("renders the same slide at two sizes", async () => {
    const small = await loadImage(await renderer.renderSlide(doc, 0, { width: 320 }));
    const large = await loadImage(await renderer.renderSlide(doc, 0, { width: 1280 }));
    expect([small.width, large.width]).toEqual([320, 1280]);
  });

  it("rejects a page that does not exist", async () => {
    await expect(renderer.renderSlide(doc, 5)).rejects.toThrow("没有第 6 页");
  });

  it("draws see-through text see-through", async () => {
    const text = (alpha: number) => `<p:sp><p:nvSpPr><p:cNvPr id="4" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${box(0, 0, 9000000, 4000000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
      + `<p:txBody><a:bodyPr/><a:lstStyle/>${para("■■■", "", `<a:rPr lang="en" sz="12000"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="${alpha}"/></a:srgbClr></a:solidFill></a:rPr>`)}</p:txBody></p:sp>`;
    const darkest = async (alpha: number) => {
      const document = await new PptxParser().parse((await buildDeck({ slideShapes: text(alpha) })) as unknown as ArrayBuffer);
      const local = new HeadlessRenderer();
      local.prepare(document);
      const png = await local.renderSlide(document, 0, { width: 800 });
      const image = await loadImage(png);
      const canvas = createCanvas(image.width, image.height);
      const ctx = canvas.getContext("2d");
      ctx.drawImage(image, 0, 0);
      const data = ctx.getImageData(0, 0, image.width, image.height).data;
      let green = 255;
      for (let i = 0; i < data.length; i += 4) green = Math.min(green, data[i + 1]);
      return green;
    };
    expect(await darkest(100000)).toBe(0);
    // 30% red over white: green and blue come out near 255 * 0.7.
    const faded = await darkest(30000);
    expect(faded).toBeGreaterThan(165);
    expect(faded).toBeLessThan(195);
  });

  it("sets lines of text in a font that is not installed at a natural line height, not one em", async () => {
    const lines = ["中文第一行", "中文第二行", "中文第三行", "中文第四行"];
    const body = lines.map(line => para(line, "", `<a:rPr lang="zh-CN" sz="2000"><a:latin typeface="NoSuchFont"/><a:ea typeface="NoSuchEastAsianFont"/></a:rPr>`)).join("");
    const shape = `<p:sp><p:nvSpPr><p:cNvPr id="4" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${box(0, 0, 5000000, 4000000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`;
    const document = await new PptxParser().parse((await buildDeck({ slideShapes: shape })) as unknown as ArrayBuffer);
    const local = new HeadlessRenderer();
    local.prepare(document);
    const width = 960;
    const image = await loadImage(await local.renderSlide(document, 0, { width }));
    const canvas = createCanvas(image.width, image.height);
    const ctx = canvas.getContext("2d");
    ctx.drawImage(image, 0, 0);
    const data = ctx.getImageData(0, 0, image.width, Math.round(image.height * 0.8)).data;
    const inked: number[] = [];
    for (let y = 0; y < data.length / 4 / image.width; y++) {
      let dark = false;
      for (let x = 0; x < image.width && !dark; x += 2) dark = data[(y * image.width + x) * 4] < 120;
      if (dark) inked.push(y);
    }
    const tops: number[] = [];
    inked.forEach((y, i) => { if (i === 0 || y - inked[i - 1] > 3) tops.push(y); });
    expect(tops.length).toBe(4);
    const pitch = (tops[3] - tops[0]) / 3;
    const fontPx = 20 * width / (document.size.width * document.slides[0].emuPerPx / 12700);
    // One em would be a pitch of 1.00; PowerPoint and LibreOffice set Chinese text at about 1.2 to 1.3.
    expect(pitch / fontPx).toBeGreaterThan(1.12);
    expect(pitch / fontPx).toBeLessThan(1.5);
  });
});
