// @vitest-environment node
import { beforeAll, describe, expect, it } from "vitest";
import { loadImage, createCanvas } from "@napi-rs/canvas";
import { HeadlessRenderer } from "../src/headless/renderer";
import { installNodeEnv } from "../src/headless/node-env";
import { PptxParser } from "../src/pptx-parser";
import { PptVirtualDocument } from "../src/pptx-virtual-dom";
import { buildDeck } from "./helpers/deck";
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
});
