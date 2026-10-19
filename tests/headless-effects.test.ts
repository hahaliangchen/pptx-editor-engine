// @vitest-environment node
import { afterEach, describe, expect, it } from "vitest";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { HeadlessSession } from "../src/headless/session";
import { buildDeck } from "./helpers/deck";
import { apiDeckParts } from "./helpers/api-deck";
import { checkPackage } from "./helpers/package-check";

const sessions: HeadlessSession[] = [];
async function open(bytes?: Buffer | Uint8Array): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(bytes ?? await buildDeck(apiDeckParts()));
  sessions.push(session);
  return session;
}
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

const card = { slide: 1, name: "卡片" };
const shapeOf = async (session: HeadlessSession, name = "卡片") =>
  (await session.outline({ slides: [1] })).slides[0].shapes.find((shape: { name: string }) => shape.name === name);

async function pixelAt(png: Buffer, x: number, y: number): Promise<number[]> {
  const image = await loadImage(png);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  return [...ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data];
}

const WIDTH = 960;
const px = (session: HeadlessSession, emu: number) => (emu / (session.document.size.width * session.document.slides[0].emuPerPx)) * WIDTH;
// The card: x 914400, y 1828800, 3657600 x 1371600 EMU.
const cardPoint = (session: HeadlessSession, fx: number, fy: number) =>
  [px(session, 914400 + 3657600 * fx), px(session, 1828800 + 1371600 * fy)];

const problems = async (bytes: Uint8Array) =>
  (await checkPackage(bytes)).filter(message => !/^no content type for .*\.rels$/.test(message));

describe("gradients", () => {
  const gradient = { angle: 0, stops: [{ position: 0, color: "#ff0000" }, { position: 100, color: "#0000ff" }] };

  it("sets a gradient, reads it back, and draws it from one colour to the other", async () => {
    const session = await open();
    expect((await session.apply([{ op: "setFill", target: card, gradient }])).ok).toBe(true);
    expect((await shapeOf(session)).fill).toEqual({
      gradient: { type: "linear", angle: 0, stops: [{ position: 0, color: "#ff0000" }, { position: 100, color: "#0000ff" }] }
    });
    const png = await session.render(1, { width: WIDTH });
    const [lx, ly] = cardPoint(session, 0.12, 0.5);
    const [rx, ry] = cardPoint(session, 0.88, 0.5);
    const left = await pixelAt(png, lx, ly);
    const right = await pixelAt(png, rx, ry);
    expect(left[0]).toBeGreaterThan(left[2]);
    expect(right[2]).toBeGreaterThan(right[0]);
  });

  it("runs the other way when the angle says so", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: card, gradient: { ...gradient, angle: 90 } }]);
    const png = await session.render(1, { width: WIDTH });
    const [tx, ty] = cardPoint(session, 0.5, 0.12);
    const [bx, by] = cardPoint(session, 0.5, 0.88);
    expect((await pixelAt(png, tx, ty))[0]).toBeGreaterThan((await pixelAt(png, bx, by))[0]);
  });

  it("makes a radial gradient and a see-through one", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: card, gradient: { kind: "radial", stops: [{ position: 0, color: "#ffffff" }, { position: 100, color: "#000000", opacity: 0.5 }] } }]);
    expect((await shapeOf(session)).fill).toMatchObject({ gradient: { type: "radial", stops: [{ position: 0, color: "#ffffff" }, { position: 100, color: "#000000", opacity: 0.5 }] } });
  });

  it("is undone as a step", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: card, gradient }]);
    session.undo();
    expect((await shapeOf(session)).fill).toBe("#4472c4");
  });

  it("rejects what does not make a gradient", async () => {
    const session = await open();
    const one = await session.apply([{ op: "setFill", target: card, gradient: { stops: [{ position: 0, color: "#ff0000" }] } }]);
    expect(one.error).toMatchObject({ code: "INVALID_SCHEMA" });
    const both = await session.apply([{ op: "setFill", target: card, color: "#ff0000", gradient }]);
    expect(both.error).toMatchObject({ code: "INVALID_SCHEMA" });
    expect(both.error!.message).toContain("不能同时");
    const bad = await session.apply([{ op: "setFill", target: card, gradient: { stops: [{ position: 150, color: "#ff0000" }, { position: 0, color: "#000000" }] } }]);
    expect(bad.error).toMatchObject({ code: "INVALID_SCHEMA" });
  });
});

describe("opacity", () => {
  it("sets a colour with an opacity, and the outline says so", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: card, color: "#c00000", opacity: 0.5 }]);
    expect(await shapeOf(session)).toMatchObject({ fill: "#c00000", opacity: 0.5 });
    const [x, y] = cardPoint(session, 0.5, 0.15);
    const blended = await pixelAt(await session.render(1, { width: WIDTH }), x, y);
    // Half red over white is pink: red stays high, green and blue rise.
    expect(blended[0]).toBeGreaterThan(200);
    expect(blended[1]).toBeGreaterThan(100);
  });

  it("changes only the opacity when no colour is given, and keeps the colour", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: card, opacity: 0.4 }]);
    expect(await shapeOf(session)).toMatchObject({ fill: "#4472c4", opacity: 0.4 });
    await session.apply([{ op: "setFill", target: card, opacity: 1 }]);
    expect((await shapeOf(session)).opacity).toBeUndefined();
  });

  it("applies to every stop of a gradient", async () => {
    const session = await open();
    await session.apply([
      { op: "setFill", target: card, gradient: { stops: [{ position: 0, color: "#ff0000" }, { position: 100, color: "#0000ff" }] } },
      { op: "setFill", target: card, opacity: 0.5 }
    ]);
    const stops = (await shapeOf(session)).fill.gradient.stops;
    expect(stops.map((stop: { opacity: number }) => stop.opacity)).toEqual([0.5, 0.5]);
  });

  it("will not set the opacity of a shape that has no fill", async () => {
    const session = await open();
    const result = await session.apply([{ op: "setFill", target: { slide: 1, name: "双格式" }, opacity: 0.5 }]);
    expect(result.error).toMatchObject({ code: "NOT_APPLICABLE", path: "commands[0].opacity" });
  });

  it("is undone", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: card, opacity: 0.4 }]);
    session.undo();
    expect((await shapeOf(session)).opacity).toBeUndefined();
  });
});

describe("shadows", () => {
  const shadow = { color: "#000000", opacity: 0.5, blurPt: 8, distancePt: 6, angle: 90 };

  it("adds a shadow, reads it back in points and degrees, and draws it below the shape", async () => {
    const session = await open();
    const before = await session.render(1, { width: WIDTH });
    await session.apply([{ op: "setShadow", target: card, shadow }]);
    expect((await shapeOf(session)).shadow).toEqual({ color: "#000000", opacity: 0.5, blurPt: 8, distancePt: 6, angle: 90 });
    const [x, y] = [px(session, 914400 + 1800000), px(session, 1828800 + 1371600) + 8];
    const after = await session.render(1, { width: WIDTH });
    expect((await pixelAt(after, x, y))[0]).toBeLessThan((await pixelAt(before, x, y))[0]);
  });

  it("fills in what is left out", async () => {
    const session = await open();
    await session.apply([{ op: "setShadow", target: card, shadow: {} }]);
    expect((await shapeOf(session)).shadow).toEqual({ color: "#000000", opacity: 0.35, blurPt: 6, distancePt: 4, angle: 45 });
  });

  it("removes a shadow with null, and undo brings it back", async () => {
    const session = await open();
    await session.apply([{ op: "setShadow", target: card, shadow }]);
    await session.apply([{ op: "setShadow", target: card, shadow: null }]);
    expect((await shapeOf(session)).shadow).toBeUndefined();
    session.undo();
    expect((await shapeOf(session)).shadow).toBeTruthy();
  });

  it("works on a picture too", async () => {
    const session = await open();
    const canvas = createCanvas(40, 30);
    canvas.getContext("2d").fillRect(0, 0, 40, 30);
    const src = `data:image/png;base64,${canvas.toBuffer("image/png").toString("base64")}`;
    await session.apply([
      { op: "addImage", slide: 1, src, rect: { x: 20, y: 2, w: 4, h: 3 }, name: "图" },
      { op: "setShadow", target: { slide: 1, name: "图" }, shadow }
    ]);
    expect((await shapeOf(session, "图")).shadow).toMatchObject({ blurPt: 8, distancePt: 6 });
  });

  it("refuses what cannot have one", async () => {
    const session = await open();
    await session.apply([
      { op: "addTable", slide: 1, rows: [["a"]], name: "表" },
      { op: "addChart", slide: 1, type: "pie", categories: ["a"], series: [{ name: "s", values: [1] }], name: "图表" }
    ]);
    for (const name of ["表", "图表", "数据组"]) {
      expect((await session.apply([{ op: "setShadow", target: { slide: 1, name }, shadow }])).error?.code, name).toBe("NOT_APPLICABLE");
    }
  });
});

describe("what is natural to write when a shape is made", () => {
  const rect = { x: 2, y: 10, w: 6, h: 4 };

  it("takes the fill, opacity, border and shadow in one command", async () => {
    const session = await open();
    const result = await session.apply([{
      op: "addShape", slide: 1, shape: "roundRect", rect, name: "一步到位",
      fill: { gradient: { type: "linear", angle: 45, stops: [{ position: 0, color: "#112233" }, { position: 100, color: "#aabbcc" }] } },
      opacity: 0.8, line: null, shadow: { blurPt: 9 }
    }]);
    expect(result.ok).toBe(true);
    const shape = await shapeOf(session, "一步到位");
    expect(shape.fill.gradient).toMatchObject({ type: "linear", angle: 45 });
    expect(shape.fill.gradient.stops.map((stop: { opacity?: number }) => stop.opacity)).toEqual([0.8, 0.8]);
    expect(shape.line).toBeNull();
    expect(shape.shadow).toMatchObject({ blurPt: 9, distancePt: 4 });
  });

  it("takes a fill back exactly as outline() reported it", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: card, gradient: { angle: 30, stops: [{ position: 0, color: "#ff0000" }, { position: 100, color: "#0000ff" }] } }]);
    const reported = (await shapeOf(session)).fill;
    expect((await session.apply([{ op: "addShape", slide: 1, shape: "rect", rect, name: "拷贝", fill: reported }])).ok).toBe(true);
    expect((await shapeOf(session, "拷贝")).fill).toEqual(reported);
  });

  it("reads `type` and `kind` alike in a gradient", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: card, gradient: { type: "radial", stops: [{ position: 0, color: "#ffffff" }, { position: 100, color: "#000000" }] } }]);
    expect((await shapeOf(session)).fill.gradient.type).toBe("radial");
  });

  it("takes a shadow for a picture in the command that adds it", async () => {
    const session = await open();
    const canvas = createCanvas(20, 20);
    canvas.getContext("2d").fillRect(0, 0, 20, 20);
    const src = `data:image/png;base64,${canvas.toBuffer("image/png").toString("base64")}`;
    await session.apply([{ op: "addImage", slide: 1, src, rect: { x: 20, y: 2, w: 3, h: 3 }, name: "带影图", shadow: { distancePt: 7 } }]);
    expect((await shapeOf(session, "带影图")).shadow).toMatchObject({ distancePt: 7 });
  });

  it("will not set an opacity on a shape made without a fill", async () => {
    const session = await open();
    const result = await session.apply([{ op: "addShape", slide: 1, shape: "rect", rect, fill: null, opacity: 0.5 }]);
    expect(result.error).toMatchObject({ code: "INVALID_SCHEMA", path: "commands[0].opacity" });
  });

  it("lists the fields a command has when it is given one it does not know", async () => {
    const session = await open();
    const result = await session.apply([{ op: "addShape", slide: 1, shape: "rect", rect, glow: 3 }]);
    expect(result.error!.message).toContain('不认识的字段 "glow"');
    expect(result.error!.message).toContain("这里可用的字段");
    expect(result.error!.message).toContain("shadow");
  });
});

describe("exporting looks", () => {
  it("writes gradient, opacity and shadow so they read back the same", async () => {
    const session = await open();
    await session.apply([
      { op: "setFill", target: card, gradient: { angle: 30, stops: [{ position: 0, color: "#112233" }, { position: 60, color: "#445566", opacity: 0.8 }, { position: 100, color: "#778899" }] } },
      { op: "setShadow", target: card, shadow: { color: "#102030", opacity: 0.3, blurPt: 5, distancePt: 3, angle: 120 } },
      { op: "setFill", target: { slide: 1, name: "双格式" }, color: "#00b050", opacity: 0.25 }
    ]);
    const bytes = await session.export();
    expect(await problems(bytes)).toEqual([]);
    const reopened = await open(bytes);
    const once = await shapeOf(session);
    const twice = await shapeOf(reopened);
    expect(twice.fill).toEqual(once.fill);
    expect(twice.shadow).toEqual(once.shadow);
    expect(await shapeOf(reopened, "双格式")).toMatchObject({ fill: "#00b050", opacity: 0.25 });
  });

  it("edits a look that was written by an earlier export", async () => {
    const first = await open();
    await first.apply([{ op: "setFill", target: card, color: "#c00000", opacity: 0.5 }, { op: "setShadow", target: card, shadow: {} }]);
    const second = await open(await first.export());
    await second.apply([{ op: "setFill", target: card, opacity: 0.8 }, { op: "setShadow", target: card, shadow: { blurPt: 12 } }]);
    const third = await open(await second.export());
    expect(await shapeOf(third)).toMatchObject({ fill: "#c00000", opacity: 0.8, shadow: { blurPt: 12, distancePt: 4 } });
  });
});
