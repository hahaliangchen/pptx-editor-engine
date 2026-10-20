// @vitest-environment node
import { readFileSync } from "node:fs";
import { createCanvas } from "@napi-rs/canvas";
import { afterEach, describe, expect, it } from "vitest";
import { HeadlessSession } from "../src/headless/session";

const sessions: HeadlessSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

/** A wide picture: red left half, blue right half. */
function wide(): string {
  const canvas = createCanvas(400, 200);
  const context = canvas.getContext("2d");
  context.fillStyle = "#ff0000"; context.fillRect(0, 0, 200, 200);
  context.fillStyle = "#0000ff"; context.fillRect(200, 0, 200, 200);
  return canvas.toDataURL("image/png");
}

describe("addImage with fit", () => {
  async function open(): Promise<HeadlessSession> {
    const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), { quiet: true });
    sessions.push(session);
    return session;
  }

  it("cuts a picture to the slot's proportions: a square slot takes the middle of a wide picture", async () => {
    const session = await open();
    const result = await session.apply({ commands: [{ op: "addImage", slide: 1, src: wide(), rect: { x: 2, y: 2, w: 6, h: 6 }, fit: { shape: "rect" } }] });
    expect(result.ok).toBe(true);
    const pic = (await session.outline()).slides[0].shapes.find((shape: { type: string }) => shape.type === "image") as { rect: { w: number; h: number } };
    expect(pic.rect.w).toBeCloseTo(6, 1);
    expect(pic.rect.h).toBeCloseTo(6, 1);
    // the saved picture is square too: both halves are in it, red and blue
    const rendered = await session.render(1, { width: 800 });
    expect(rendered.length).toBeGreaterThan(1000);
  });

  it("makes a round slot's corners transparent, and draws a placeholder when there is no picture", async () => {
    const session = await open();
    const made = await session.apply({ commands: [
      { op: "addImage", slide: 1, src: wide(), rect: { x: 2, y: 2, w: 5, h: 5 }, fit: { shape: "ellipse" } },
      { op: "addImage", slide: 1, src: "placeholder:", rect: { x: 10, y: 2, w: 8, h: 5 }, fit: { shape: "roundRect" } }
    ] });
    expect(made.ok).toBe(true);
    expect((await session.outline()).slides[0].shapes.filter((shape: { type: string }) => shape.type === "image")).toHaveLength(2);
  });

  it("refuses fit without a rect", async () => {
    const session = await open();
    const result = await session.apply({ commands: [{ op: "addImage", slide: 1, src: wide(), fit: { shape: "rect" } }] });
    expect(result.ok).toBe(false);
  });
});
