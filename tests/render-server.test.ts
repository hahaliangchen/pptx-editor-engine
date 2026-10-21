// @vitest-environment node
import { readFileSync } from "node:fs";
import { AddressInfo } from "node:net";
import type { Server } from "node:http";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createRenderServer } from "../src/headless/render-server";

const deck = readFileSync("tools/templates/blank.pptx");
let server: Server;
let base = "";

beforeAll(async () => {
  server = createRenderServer({ maxBodyBytes: 2 * 1048576, maxQueue: 3 });
  await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
  base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});
afterAll(() => new Promise<void>(resolve => server.close(() => resolve())));

const post = (path: string, body: Uint8Array | string = deck) => fetch(`${base}${path}`, { method: "POST", body: body as BodyInit });
const PNG = [0x89, 0x50, 0x4e, 0x47];
const widthOf = (png: Buffer) => png.readUInt32BE(16);

describe("render service", () => {
  it("answers a health check", async () => {
    expect(await (await fetch(`${base}/health`)).json()).toEqual({ ok: true });
  });

  it("converts PPTX to transient JSON and HTML back through the HTTP worker", async () => {
    const imported = await post("/design", JSON.stringify({ operation: "sync", pptxBase64: deck.toString("base64") }));
    expect(imported.status).toBe(200);
    const { design } = await imported.json();
    const { projectPage } = await import("../src/design/projection");
    const html = projectPage(design, design.pages[0].id).replace("background:#ffffff", "background:#183d32").replace("</body>", '<div data-kind="text" style="position:absolute;left:80px;top:80px;width:800px;height:100px"><p style="margin:0;font:32px/48px Arial;color:#ffffff">HTTP round trip</p></div></body>');
    const applied = await post("/design", JSON.stringify({ operation: "apply", design, expectedRevision: design.revision, html }));
    expect(applied.status, await applied.clone().text()).toBe(200);
    const next = (await applied.json()).design;
    expect(next.pages[0].background).toBe("#183d32");
    const bytes = Buffer.from(next.source.pptxBase64, "base64");
    const rendered = await post("/render?page=1&width=320", bytes);
    expect(rendered.status).toBe(200);
    expect(widthOf(Buffer.from(await rendered.arrayBuffer()))).toBe(320);
    const stale = await post("/design", JSON.stringify({ operation: "apply", design: next, expectedRevision: next.revision, html }));
    expect(stale.status).toBe(422);
    expect((await stale.json()).error.message).toContain("STALE_DESIGN");
  }, 20000);

  it("returns the first page as a PNG of the default width, with the slide count", async () => {
    const response = await post("/thumbnail");
    expect(response.status).toBe(200);
    expect(response.headers.get("content-type")).toBe("image/png");
    expect(Number(response.headers.get("x-slide-count"))).toBeGreaterThanOrEqual(1);
    const png = Buffer.from(await response.arrayBuffer());
    expect([...png.subarray(0, 4)]).toEqual(PNG);
    expect(widthOf(png)).toBe(480);
  });

  it("takes the width and the page, and draws any page", async () => {
    const png = Buffer.from(await (await post("/thumbnail?width=200")).arrayBuffer());
    expect(widthOf(png)).toBe(200);
    const page = await post("/render?page=1&width=640");
    expect(widthOf(Buffer.from(await page.arrayBuffer()))).toBe(640);
  });

  it("refuses what it cannot use, saying why", async () => {
    const notADeck = await post("/thumbnail", "this is not a pptx");
    expect(notADeck.status).toBe(400);
    expect((await notADeck.json()).error.code).toBe("INVALID_PPTX");
    expect((await post("/thumbnail", "")).status).toBe(400);
    const outOfRange = await post("/render?page=999");
    expect(outOfRange.status).toBe(400);
    expect((await outOfRange.json()).error.message).toContain("没有第 999 页");
    expect((await post("/thumbnail?width=5")).status).toBe(400);
    expect((await post("/render?page=abc")).status).toBe(400);
    expect((await fetch(`${base}/nowhere`)).status).toBe(404);
    expect((await post("/thumbnail", new Uint8Array(3 * 1048576))).status).toBe(413);
  });

  it("draws requests that arrive together, one after another, and says busy when the queue is full", async () => {
    const together = await Promise.all([1, 2, 3].map(() => post("/thumbnail?width=100")));
    expect(together.map(response => response.status)).toEqual([200, 200, 200]);
    const many = await Promise.all(Array.from({ length: 12 }, () => post("/thumbnail?width=100")));
    const statuses = many.map(response => response.status);
    expect(statuses).toContain(200);
    expect(statuses.every(status => status === 200 || status === 503)).toBe(true);
    // and it is still serving afterwards
    expect((await post("/thumbnail?width=100")).status).toBe(200);
  });
});
