import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import JSZip from "jszip";
import { beforeAll, describe, expect, it } from "vitest";
import { exportEditedPptx, markDirty } from "../src/pptx-exporter";
import { PptVirtualDocument, PptxParser } from "../src/pptx-parser";

// The sample deck is not part of the repository; the test only runs where it exists.
const DECK = resolve(__dirname, "../元景行业智能体方案-20260713(2) - 副本.pptx");
const hasDeck = existsSync(DECK);

describe.skipIf(!hasDeck)("real deck roundtrip", () => {
  const parser = new PptxParser();
  let doc: PptVirtualDocument;
  let bytes: Buffer;

  beforeAll(async () => {
    URL.createObjectURL = () => "blob:test";
    bytes = readFileSync(DECK);
    // A Buffer avoids the jsdom/Node ArrayBuffer realm mismatch inside JSZip.
    doc = await parser.parse(bytes as unknown as ArrayBuffer);
  }, 120_000);

  it("tags slide shapes with their source and leaves template decoration untagged", () => {
    const elements = doc.slides.flatMap(slide => slide.elements);
    const tagged = elements.filter(element => element.source);
    expect(tagged.length).toBeGreaterThan(20);
    expect(tagged.every(element => /^ppt\/slides\/slide\d+\.xml$/.test(element.source!.part))).toBe(true);
    expect(elements.length).toBeGreaterThan(tagged.length);
  });

  it("round-trips an edited shape: the new position survives a re-parse", async () => {
    const slide = doc.slides[0];
    const target = slide.elements.find(element => element.source && element.type === "text")!;
    const moved = { x: target.rect.x + 100, y: target.rect.y + 50 };
    const original = { x: target.rect.x, y: target.rect.y };
    target.rect = { ...target.rect, ...moved };

    const dirty = new Map();
    markDirty(dirty, target.source);
    const blob = await exportEditedPptx(parser.getZip()!, doc.slides, dirty);
    const buffer = await new Promise<ArrayBuffer>(resolveBuffer => {
      const reader = new FileReader();
      reader.onload = () => resolveBuffer(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob);
    });

    const reparsed = await new PptxParser().parse(buffer);
    const again = reparsed.slides[0].elements.find(element =>
      element.source?.shapeId === target.source!.shapeId && element.type === "text")!;
    expect(again.rect.x).toBeCloseTo(moved.x, 0);
    expect(again.rect.y).toBeCloseTo(moved.y, 0);
    expect(again.rect.x).not.toBeCloseTo(original.x, 0);

    // Every other part is untouched.
    const before = await JSZip.loadAsync(bytes);
    const after = await JSZip.loadAsync(buffer);
    for (const name of Object.keys(before.files)) {
      if (name === target.source!.part || before.files[name].dir) continue;
      expect(await after.file(name)!.async("uint8array")).toEqual(await before.file(name)!.async("uint8array"));
    }
  }, 120_000);
});
