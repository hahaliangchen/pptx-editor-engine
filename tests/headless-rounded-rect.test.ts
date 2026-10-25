// @vitest-environment node
import { describe, expect, it } from "vitest";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { HeadlessRenderer } from "../src/headless/renderer";
import { installNodeEnv } from "../src/headless/node-env";
import { PptxParser } from "../src/pptx-parser";
import { box, buildDeck } from "./helpers/deck";

installNodeEnv();

async function assertOutline(width: number, height: number, adjustment: number) {
  const unit = 9525;
  const shapes = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="rounded"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${box(40 * unit, 40 * unit, width * unit, height * unit)}`
    + `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adjustment}"/></a:avLst></a:prstGeom>`
    + `<a:solidFill><a:srgbClr val="E60012"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
  const doc = await new PptxParser().parse(await buildDeck({ slideShapes: shapes }) as unknown as ArrayBuffer);
  const shape = doc.slides[0].elements.find(element => element.type === "shape");
  expect(shape?.type).toBe("shape");
  const bounds = shape!.rect;
  const renderer = new HeadlessRenderer();
  const scale = 4;
  const image = await loadImage(await renderer.renderSlide(doc, 0, { width: doc.size.width * scale }));
  const actual = createCanvas(image.width, image.height).getContext("2d");
  actual.drawImage(image, 0, 0);

  const reference = createCanvas(image.width, image.height).getContext("2d");
  reference.fillStyle = "white";
  reference.fillRect(0, 0, image.width, image.height);
  reference.scale(scale, scale);
  reference.fillStyle = "#e60012";
  const radius = Math.min(bounds.w, bounds.h) * Math.min(0.5, Math.max(0, adjustment / 100000));
  reference.beginPath();
  if (width === height && radius === bounds.w / 2) {
    // An independent circle primitive catches the old rounded-square outline.
    reference.ellipse(bounds.x + radius, bounds.y + radius, radius, radius, 0, 0, Math.PI * 2);
  } else {
    reference.roundRect(bounds.x, bounds.y, bounds.w, bounds.h, radius);
  }
  reference.fill();

  const x = Math.floor(bounds.x * scale) - 1, y = Math.floor(bounds.y * scale) - 1;
  const w = Math.ceil(bounds.w * scale) + 2, h = Math.ceil(bounds.h * scale) + 2;
  const a = actual.getImageData(x, y, w, h).data;
  const b = reference.getImageData(x, y, w, h).data;
  let error = 0;
  // Green is zero inside the red shape and 255 on the white background.
  for (let i = 1; i < a.length; i += 4) error += Math.abs(a[i] - b[i]);
  expect(error / (w * h)).toBeLessThan(1.5);
}

describe("rounded rectangle raster geometry", () => {
  it("renders the 24px dot with half-size corners as a true circle", async () => {
    await assertOutline(24, 24, 50000);
  });
  it("clamps an oversized corner radius to the same circle", async () => {
    await assertOutline(24, 24, 70000);
  });
  it("preserves circular ends on a non-square capsule", async () => {
    await assertOutline(120, 48, 50000);
  });
  it("keeps a zero-radius rectangle square", async () => {
    await assertOutline(48, 24, 0);
  });
});
