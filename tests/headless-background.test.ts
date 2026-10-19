// @vitest-environment node
import { afterEach, describe, expect, it } from "vitest";
import JSZip from "jszip";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { HeadlessSession } from "../src/headless/session";
import { buildDeck } from "./helpers/deck";
import { apiDeckParts } from "./helpers/api-deck";
import { checkPackage } from "./helpers/package-check";

const sessions: HeadlessSession[] = [];
async function open(bytes?: Buffer | Uint8Array, parts = apiDeckParts()): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(bytes ?? await buildDeck(parts));
  sessions.push(session);
  return session;
}
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

const backgroundOf = async (session: HeadlessSession, slide = 1) => (await session.outline({ slides: [slide] })).slides[0].background;

async function pixelAt(png: Buffer, x: number, y: number): Promise<number[]> {
  const image = await loadImage(png);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  return [...ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data];
}

const SOLID = (color: string) => `<p:bgPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:effectLst/></p:bgPr>`;

describe("reading a background", () => {
  it("says nothing when the deck has none", async () => {
    expect(await backgroundOf(await open())).toBeUndefined();
  });

  it("reports the colour a deck's master gives every page", async () => {
    const session = await open(undefined, { ...apiDeckParts(), masterBackground: SOLID("0B1026") });
    expect(await backgroundOf(session)).toBe("#0b1026");
  });

  it("draws it, so a page of a dark deck is dark", async () => {
    const session = await open(undefined, { ...apiDeckParts(), masterBackground: SOLID("0B1026") });
    const [r, g, b] = await pixelAt(await session.render(1, { width: 480 }), 470, 260);
    expect([r, g, b]).toEqual([0x0b, 0x10, 0x26]);
  });
});

describe("what is added to a dark page", () => {
  const dark = { ...apiDeckParts(), masterBackground: SOLID("0B1026") };
  const shapeOf = async (session: HeadlessSession, name: string) =>
    (await session.outline({ slides: [1] })).slides[0].shapes.find((shape: { name: string }) => shape.name === name);

  it("gives a new text box light text, as the deck's own text colour would be", async () => {
    const session = await open(undefined, dark);
    await session.apply([{ op: "addText", slide: 1, rect: { x: 1, y: 8, w: 8, h: 2 }, text: "夜间", name: "文" }]);
    expect((await shapeOf(session, "文")).style.color).toBe("#ffffff");
  });

  it("keeps black text on a light page", async () => {
    const session = await open();
    await session.apply([{ op: "addText", slide: 1, rect: { x: 1, y: 8, w: 8, h: 2 }, text: "日间", name: "文" }]);
    expect((await shapeOf(session, "文")).style.color).toBe("#000000");
  });

  it("gives a chart near-white text", async () => {
    const session = await open(undefined, dark);
    await session.apply([{ op: "addChart", slide: 1, type: "column", categories: ["a"], series: [{ name: "s", values: [1] }], name: "图" }]);
    const element = session.document.slides[0].elements.find(candidate => candidate.type === "chart");
    expect(element && element.type === "chart" && element.chart.font.color).toBe("#f2f2f2");
  });

  it("follows a background that is set on the page, too", async () => {
    const session = await open();
    await session.apply([
      { op: "setBackground", slide: 1, color: "#101820" },
      { op: "addText", slide: 1, rect: { x: 1, y: 8, w: 8, h: 2 }, text: "刚改深", name: "文" }
    ]);
    expect((await shapeOf(session, "文")).style.color).toBe("#ffffff");
  });
});

describe("setting a background", () => {
  const set = (session: HeadlessSession, command: Record<string, unknown>) =>
    session.apply([{ op: "setBackground", slide: 1, ...command }]);

  it("sets a colour, reads it back and draws it", async () => {
    const session = await open();
    expect((await set(session, { color: "#14224e" })).ok).toBe(true);
    expect(await backgroundOf(session)).toBe("#14224e");
    expect((await pixelAt(await session.render(1, { width: 480 }), 470, 260)).slice(0, 3)).toEqual([0x14, 0x22, 0x4e]);
  });

  it("sets a gradient", async () => {
    const session = await open();
    await set(session, { gradient: { angle: 90, stops: [{ position: 0, color: "#000000" }, { position: 100, color: "#ffffff" }] } });
    expect(await backgroundOf(session)).toMatchObject({ gradient: { type: "linear", angle: 90 } });
    const png = await session.render(1, { width: 480 });
    expect((await pixelAt(png, 470, 10))[0]).toBeLessThan((await pixelAt(png, 470, 260))[0]);
  });

  it("changes a background the deck already had, and puts it back on undo", async () => {
    const session = await open(undefined, { ...apiDeckParts(), masterBackground: SOLID("0B1026") });
    await set(session, { color: "#ffffff" });
    expect(await backgroundOf(session)).toBe("#ffffff");
    session.undo();
    expect(await backgroundOf(session)).toBe("#0b1026");
  });

  it("is undone, leaving no background where there was none", async () => {
    const session = await open();
    await set(session, { color: "#c00000" });
    session.undo();
    expect(await backgroundOf(session)).toBeUndefined();
    session.redo();
    expect(await backgroundOf(session)).toBe("#c00000");
  });

  it("belongs to the page it was set on", async () => {
    const session = await open();
    await session.apply([{ op: "duplicateSlide", slide: 1 }, { op: "setBackground", slide: 2, color: "#00b050" }]);
    expect(await backgroundOf(session, 1)).toBeUndefined();
    expect(await backgroundOf(session, 2)).toBe("#00b050");
  });

  it("asks for a colour or a gradient, not both, not neither", async () => {
    const session = await open();
    const neither = await session.apply([{ op: "setBackground", slide: 1 }]);
    expect(neither.error!.message).toContain("至少要写出其中一个字段：color、gradient");
    const both = await set(session, { color: "#000000", gradient: { stops: [{ position: 0, color: "#000000" }, { position: 100, color: "#ffffff" }] } });
    expect(both.error!.message).toContain("不能同时");
    expect((await set(session, { color: "black" })).error).toMatchObject({ code: "INVALID_SCHEMA" });
    expect((await session.apply([{ op: "setBackground", slide: 9, color: "#000000" }])).error).toMatchObject({ code: "SLIDE_NOT_FOUND" });
  });
});

describe("exporting a background", () => {
  it("writes p:bg first in the page, and the page reads back the same", async () => {
    const session = await open();
    await session.apply([{ op: "setBackground", slide: 1, gradient: { angle: 45, stops: [{ position: 0, color: "#102030" }, { position: 100, color: "#405060" }] } }]);
    const bytes = await session.export();
    expect((await checkPackage(bytes)).filter(message => !/\.rels$/.test(message))).toEqual([]);

    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect(xml).toMatch(/<p:cSld><p:bg><p:bgPr><a:gradFill/);

    const reopened = await open(bytes);
    expect(await backgroundOf(reopened)).toEqual(await backgroundOf(session));
  });

  it("replaces a background the page already wrote", async () => {
    const first = await open();
    await first.apply([{ op: "setBackground", slide: 1, color: "#111111" }]);
    const second = await open(await first.export());
    await second.apply([{ op: "setBackground", slide: 1, color: "#222222" }]);
    const zip = await JSZip.loadAsync(await second.export());
    const xml = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect(xml.match(/<p:bg>/g)).toHaveLength(1);
    expect(xml).toContain("222222");
    expect(await backgroundOf(await open(await second.export()))).toBe("#222222");
  });
});
