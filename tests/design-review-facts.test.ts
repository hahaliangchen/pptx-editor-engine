import { describe, expect, it } from "vitest";
import { cloneDesign, type DesignDocument, type MeasuredPage } from "../src/design/model";
import { mergeHtmlDifference } from "../src/design/html-diff";
import { checkLayout } from "../src/api/layout-check";
import type { Outline, OutlineShape } from "../src/api/outline";
import { callTool, MAX_RENDER_SLIDES, type ToolHost } from "../src/api/tools";

const run = (text: string) => ({ text, fontFamily: "Noto Sans SC", fontSize: 30, color: "#ffffff", bold: true, italic: false, letterSpacing: 0 });
const paragraph = (text: string) => ({ align: "center" as const, before: 0, after: 0, lineHeight: 36, runs: [run(text)] });
const textNode = (id: string, name: string, ...texts: string[]) => ({
  id, kind: "text" as const, name,
  appearance: { rect: { x: 100, y: 100, w: 400, h: 60 }, text: { padding: [0, 0, 0, 0] as [number, number, number, number], verticalAlign: "top" as const, paragraphs: texts.map(paragraph) } }
});

/** The table of contents from project 816: a subtitle and the first number marker. */
const fixture = () => {
  const design: DesignDocument = { schemaVersion: 1, revision: 6, canvas: { width: 1280, height: 720, unit: "css-px" },
    source: { pptxBase64: "", sha256: "a".repeat(64), widthEMU: 1, heightEMU: 1 }, diagnostics: [],
    pages: [{ id: "toc", part: "ppt/slides/slide2.xml", nodes: [textNode("toc#7", "TextBox 7", "CONTENTS"), textNode("toc#10", "TextBox 10", "一"), textNode("toc#11", "TextBox 11", "旧条目")] }] };
  const baseline: MeasuredPage = { pageId: "toc", revision: 6, sourceHash: design.source.sha256, nodes: cloneDesign(design.pages[0].nodes) };
  return { design, baseline, edited: cloneDesign(baseline) };
};

describe("facts an edit reports about text", () => {
  it("shows each text node's words before and after, so an unintended change is visible", () => {
    const { design, baseline, edited } = fixture();
    edited.nodes[0].appearance.text!.paragraphs[0].runs[0].text = "DISCUSSION";
    edited.nodes[1].appearance.text!.paragraphs.push(paragraph("一"));
    edited.nodes.splice(2, 1);
    edited.nodes.push(textNode("new:1", "New", "新条目"));
    const changes = mergeHtmlDifference(design, baseline, edited).changes!;
    expect(changes.find(c => c.nodeId === "toc#7")).toMatchObject({ action: "update", text: { before: "CONTENTS", after: "DISCUSSION" } });
    expect(changes.find(c => c.nodeId === "toc#10")).toMatchObject({ action: "update", text: { before: "一", after: "一\n一" } });
    expect(changes.find(c => c.nodeId === "toc#11")).toMatchObject({ action: "remove", text: { before: "旧条目" } });
    expect(changes.find(c => c.nodeId === "new:1")).toMatchObject({ action: "add", text: { after: "新条目" } });
  });

  it("leaves text out of a change that did not touch the words", () => {
    const { design, baseline, edited } = fixture();
    edited.nodes[0].appearance.rect.w = 500;
    const change = mergeHtmlDifference(design, baseline, edited).changes!.find(c => c.nodeId === "toc#7")!;
    expect(change.fields).toEqual(["appearance.rect.w"]);
    expect(change.text).toBeUndefined();
  });

  it("warns when one text box holds the same paragraph twice, without rejecting the page", () => {
    const { design, baseline, edited } = fixture();
    edited.nodes[1].appearance.text!.paragraphs.push(paragraph("一"));
    const next = mergeHtmlDifference(design, baseline, edited);
    expect(next.revision).toBe(7);
    const warning = next.diagnostics.find(d => d.nodeId === "toc#10")!;
    expect(warning.severity).toBe("warning");
    expect(warning.message).toContain('Repeated paragraph in TextBox 10: "一" appears 2 times');
    expect(warning.message).toContain("keep one copy");
    // An untouched box with legitimately repeated words is not re-reported on every edit.
    expect(next.diagnostics.filter(d => d.nodeId !== "toc#10")).toEqual([]);
  });
});

describe("check_layout over the whole deck", () => {
  it("reports a paragraph repeated inside one box", () => {
    const shape: OutlineShape = { id: "8", name: "TextBox 8", type: "text", rotation: 0, rect: { x: 14, y: 6, w: 4.8, h: 0.8 },
      text: "讨论与行动\n讨论与行动", style: { fontSizePt: 18 } };
    const outline: Outline = { slideSize: { w: 33.87, h: 19.05 }, layouts: [], slides: [{ index: 5, id: "5", shapes: [shape] }] };
    const repeated = checkLayout(outline).filter(issue => issue.kind === "repeated_text");
    expect(repeated).toEqual([{ slide: 5, kind: "repeated_text", ids: ["8"], message: expect.stringContaining("“讨论与行动”出现了 2 次") }]);
    shape.text = "讨论与行动\n欢迎讨论";
    expect(checkLayout(outline).some(issue => issue.kind === "repeated_text")).toBe(false);
  });
});

describe("render_slide over several pages", () => {
  const drawn: number[] = [];
  const host: ToolHost = {
    outline: async () => { throw new Error("not needed"); },
    apply: async () => ({ ok: true, results: [] }),
    render: async page => { drawn.push(page); return new Uint8Array([0x89, 0x50, 0x4e, 0x47, page]); }
  };

  it("returns one image per page in the order asked, each saying which page it is", async () => {
    drawn.length = 0;
    const result = await callTool(host, "render_slide", { slides: [3, 1, 2], width: 480 });
    expect(result.ok).toBe(true);
    const images = result.content.filter(part => part.type === "image");
    expect(images.map(image => (image as { meta?: { slide: number } }).meta)).toEqual([{ slide: 3 }, { slide: 1 }, { slide: 2 }]);
    expect(drawn).toEqual([3, 1, 2]);
    expect(result.content.at(-1)).toEqual({ type: "text", text: "第 3、1、2 页，每页一张，宽 480 像素" });
    const single = await callTool(host, "render_slide", { slide: 2 });
    expect(single.content.filter(part => part.type === "image").map(image => (image as { meta?: unknown }).meta)).toEqual([{ slide: 2 }]);
  });

  it("says what is wrong with the page list", async () => {
    for (const [args, message] of [
      [{ slides: [] }, "slides"],
      [{ slides: [1, 1] }, "不能重复"],
      [{ slides: Array.from({ length: MAX_RENDER_SLIDES + 1 }, (_, i) => i + 1) }, `最多 ${MAX_RENDER_SLIDES} 页`],
      [{ slides: [0] }, "slides"],
      [{ slide: 1, slides: [2] }, "只给一个"],
      [{}, "slide"]
    ] as const) {
      const result = await callTool(host, "render_slide", args);
      expect(result.ok, JSON.stringify(args)).toBe(false);
      expect(JSON.stringify(result.content)).toContain(message);
    }
  });
});
