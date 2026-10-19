// @vitest-environment node
import { afterEach, describe, expect, it } from "vitest";
import Ajv from "ajv";
import { createCanvas } from "@napi-rs/canvas";
import { HeadlessSession } from "../src/headless/session";
import { PAGE_TOOLS, runPageTool } from "../src/page/page-tools";
import { resolveProjectAssets } from "../src/page/assets";
import type PptViewer from "../src/index";
import { buildDeck } from "./helpers/deck";
import { apiDeckParts } from "./helpers/api-deck";

const sessions: HeadlessSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

/** Enough of a viewer for the page tools: the headless session has the same edit API. */
async function viewer(): Promise<{ viewer: PptViewer; session: HeadlessSession }> {
  const session = await HeadlessSession.open(await buildDeck(apiDeckParts()));
  sessions.push(session);
  return { session, viewer: { api: session, renderPagePng: () => Promise.reject(new Error("no renderer")) } as unknown as PptViewer };
}

const pngDataUrl = () => {
  const canvas = createCanvas(40, 30);
  canvas.getContext("2d").fillRect(0, 0, 40, 30);
  return `data:image/png;base64,${canvas.toBuffer("image/png").toString("base64")}`;
};
const json = (result: { content: Array<{ type: string; text?: string }> }) =>
  JSON.parse(result.content.find(part => part.type === "text")!.text!);

describe("project pictures in page tools", () => {
  it("declares project references in apply_edits but keeps the engine's schema strict", () => {
    const schema = PAGE_TOOLS.find(tool => tool.name === "apply_edits")!.inputSchema;
    const validate = new Ajv({ strict: false }).compile(schema);
    expect(validate({ commands: [{ op: "addImage", slide: 1, src: "project:images/ab_logo.png" }] })).toBe(true);
    expect(validate({ commands: [{ op: "addImage", slide: 1, src: "https://example.com/a.png" }] })).toBe(false);
    expect(JSON.stringify(schema)).toContain("remote_path");
  });

  it("fetches the picture once and places it on the slide", async () => {
    const { viewer: page, session } = await viewer();
    const asked: string[] = [];
    const resolveAsset = async (path: string) => { asked.push(path); return pngDataUrl(); };
    const result = await runPageTool(page, "apply_edits", {
      commands: [
        { op: "addImage", slide: 1, src: "project:images/ab_logo.png", name: "标志", rect: { x: 1, y: 1, w: 4, h: 3 } },
        { op: "addImage", slide: 1, src: "project:images/ab_logo.png", name: "标志二" }
      ]
    }, { resolveAsset });
    expect(result.ok).toBe(true);
    expect(asked).toEqual(["images/ab_logo.png"]);
    const names = (await session.outline()).slides[0].shapes.map((shape: { name: string }) => shape.name);
    expect(names).toEqual(expect.arrayContaining(["标志", "标志二"]));
  });

  it("fails the batch, changing nothing, when a picture cannot be read", async () => {
    const { viewer: page, session } = await viewer();
    const before = (await session.outline()).slides[0].shapes.length;
    const commands = [
      { op: "addText", slide: 1, text: "说明" },
      { op: "addImage", slide: 1, src: "project:images/missing.png" }
    ];
    const missing = await runPageTool(page, "apply_edits", { commands }, { resolveAsset: () => Promise.reject(new Error("404")) });
    expect(missing.ok).toBe(false);
    expect(json(missing)).toMatchObject({ ok: false, failedAt: 1, error: { path: "commands[1].src", message: expect.stringContaining("404") } });

    const unsupported = await runPageTool(page, "apply_edits", { commands });
    expect(json(unsupported)).toMatchObject({ ok: false, failedAt: 1, error: { code: "UNSUPPORTED" } });
    expect((await session.outline()).slides[0].shapes.length).toBe(before);
  });

  it("only reads under images/, and leaves other batches alone", async () => {
    const resolve = () => Promise.resolve(pngDataUrl());
    for (const src of ["project:../deck.pptx", "project:images/../deck.pptx", "project:sources/brief.md"]) {
      const out = await resolveProjectAssets([{ op: "addImage", slide: 1, src }], resolve);
      expect("failure" in out && out.failure.error?.code).toBe("INVALID_SCHEMA");
    }
    const plain = { commands: [{ op: "addImage", slide: 1, src: "data:image/png;base64,AAAA" }] };
    expect(await resolveProjectAssets(plain, resolve)).toEqual({ batch: plain });
    const bare = [{ op: "addImage", slide: 1, src: "project:images/x.png" }];
    const out = await resolveProjectAssets(bare, resolve);
    expect("batch" in out && (out.batch as Array<{ src: string }>)[0].src).toMatch(/^data:image\/png;base64,/);
    expect(bare[0].src).toBe("project:images/x.png");
  });
});
