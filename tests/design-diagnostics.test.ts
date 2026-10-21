// @vitest-environment node
import { createCanvas } from "@napi-rs/canvas";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { measureHtml } from "../src/design/measure";

const pixel = createCanvas(1, 1).toDataURL("image/png");
// Preserve the real rejected markup; only resolve remote asset references for
// the isolated measurement test, just as the host embeds its local assets.
const fixture = (name: string) => readFileSync(`tests/fixtures/html-diagnostics/${name}`, "utf8")
  .replace(/src="(?:project:|design-asset:)[^"]*"/g, `src="${pixel}"`);
const html = (content: string) => `<!doctype html><html><head><style>html,body{margin:0;width:1280px;height:720px}p{margin:0}</style></head><body data-page-id="ppt/slides/slide1.xml" data-design-revision="7" data-source-hash="${"a".repeat(64)}">${content}</body></html>`;
async function diagnostic(markup: string) {
  try { await measureHtml(markup, undefined, true); }
  catch (error) {
    const match = String(error).match(/(\{"code":.*\})/);
    expect(match, String(error)).not.toBeNull();
    return JSON.parse(match![1]);
  }
  throw new Error("Invalid HTML unexpectedly passed measurement");
}

describe("actionable HTML validation diagnostics", () => {
  it.each([1064, 1086])("identifies the real project820 duplicate data-kind error (seq%s) before generic text checks", async seq => {
    const result = await diagnostic(fixture(`high-qa-apply-${seq}.html`));
    expect(result).toMatchObject({ code: "INVALID_DESIGN_KIND", pageId: "ppt/slides/slide5.xml", tag: "div", nodeId: null,
      ownerId: null, ownerKind: "new:d37666c7572958a3f5cad85445999fdd:9" });
    expect(result.textSnippet).toContain("排期要回到产能约束");
    expect(result.textSnippet.length).toBeLessThanOrEqual(160);
    expect(result.recovery).toContain("data-id, not data-kind");
    expect(result.recovery).toContain("duplicate data-kind");
    expect(result.recovery).toContain("No page content was applied");
  }, 20000);

  it("locates the real project820 image wrapper error and requests a direct img element", async () => {
    const result = await diagnostic(fixture("fourth-turn-failed.html"));
    expect(result).toMatchObject({ code: "IMAGE_TAG_REQUIRED", pageId: "ppt/slides/slide5.xml", tag: "div", ownerKind: "image" });
    expect(result.recovery).toContain('<img data-kind="image"');
    expect(result.recovery).toContain("not a div wrapper");
  }, 20000);

  it("reports the nearest owner and bounded text for a paragraph directly inside a group", async () => {
    const result = await diagnostic(html(`<div data-id="group-1" data-kind="group"><p>${"未归属文字 ".repeat(60)}</p></div>`));
    expect(result).toMatchObject({ code: "TEXT_OWNER_REQUIRED", pageId: "ppt/slides/slide1.xml", revision: 7, tag: "p", nodeId: null, ownerId: "group-1", ownerKind: "group", ownerIndex: 0 });
    expect(result.textSnippet.length).toBe(160);
    expect(result.recovery).toContain('data-kind="text"');
  }, 20000);

  it("reports an unowned text node without inventing an owner", async () => {
    const result = await diagnostic(html("<p>正文没有设计节点</p>"));
    expect(result).toMatchObject({ code: "TEXT_OWNER_REQUIRED", tag: "p", ownerId: null, ownerKind: null, ownerIndex: null, textSnippet: "正文没有设计节点" });
  }, 20000);

  it("still measures correctly marked text and images without changing identity or geometry", async () => {
    const result = await measureHtml(html(`<div data-id="text-1" data-kind="text" style="position:absolute;left:70px;top:120px;width:500px;height:100px"><p style="font:24px/32px Arial">Valid content</p></div><img data-id="image-1" data-kind="image" style="position:absolute;left:800px;top:100px;width:64px;height:64px;object-fit:fill" src="${pixel}">`));
    expect(result.nodes.map(node => node.id)).toEqual(["text-1", "image-1"]);
    expect(result.nodes[0].appearance.rect).toEqual({ x: 70, y: 120, w: 500, h: 100 });
    expect(result.nodes[0].appearance.text?.paragraphs[0].runs[0].text).toBe("Valid content");
    expect(result.nodes[1].appearance.rect).toEqual({ x: 800, y: 100, w: 64, h: 64 });
  }, 20000);
});
