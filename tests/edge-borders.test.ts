// @vitest-environment node
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { importDesign } from "../src/design/import";
import { exportDesign } from "../src/design/export";
import { projectPage } from "../src/design/projection";
import { mergeHtml } from "../src/design/measure";

const blank = async () => importDesign(readFileSync("tools/templates/blank.pptx"));
// The 2026-11-12 run: an accent bar on one side and a dashed divider on top
// were refused as "Only uniform ... borders", failing the whole page.
const card = `<div data-kind="shape" style="position:absolute;left:100px;top:100px;width:400px;height:200px;background:#fff5f2;border-left:6px solid #c00000"></div>
<div data-kind="text" style="position:absolute;left:600px;top:100px;width:500px;height:80px;border-top:2px dashed #d8d8d8"><p style="margin:0;font-size:20px;font-family:'Noto Sans CJK SC'">分隔线下的说明</p></div>
<div data-kind="shape" style="position:absolute;left:100px;top:400px;width:300px;height:150px;border-top:4px solid #e84c22;border-left:4px solid #e84c22"></div>`;

describe("borders on some sides only", () => {
  it("become separate editable shapes that survive export and a second edit", async () => {
    const design = await blank();
    const merged = await mergeHtml(design, projectPage(design, design.pages[0].id).replace("</body>", `${card}</body>`));
    const nodes = merged.pages[0].nodes.filter(n => n.id.startsWith("new:"));
    const edges = nodes.filter(n => n.id.endsWith("-border"));
    expect(edges.map(n => n.id.split(":").at(-1))).toEqual(["left-border", "top-border", "top-border", "left-border"]);
    const [accent, divider, frameTop, frameLeft] = edges;
    expect(accent.appearance.fill).toBe("#c00000");
    expect(accent.appearance.rect).toEqual({ x: 100, y: 100, w: 6, h: 200 });
    expect(divider.appearance.border).toEqual({ color: "#d8d8d8", width: 2, style: "dashed" });
    expect(divider.appearance.rect).toEqual({ x: 600, y: 101, w: 500, h: 0 });
    expect(frameTop.appearance.rect).toEqual({ x: 100, y: 400, w: 300, h: 4 });
    expect(frameLeft.appearance.rect).toEqual({ x: 100, y: 400, w: 4, h: 150 });
    // The node keeps its own paint; only the per-side border moved to the edges.
    const owner = nodes.find(n => n.appearance.fill === "#fff5f2")!;
    expect(owner.appearance.border).toBeUndefined();
    // Edges paint over their node.
    expect(nodes.indexOf(accent)).toBe(nodes.indexOf(owner) + 1);

    const reopened = await importDesign(await exportDesign(merged));
    const native = reopened.pages[0].nodes;
    expect(native.some(n => n.appearance.fill === "#c00000" && Math.round(n.appearance.rect.w) === 6 && Math.round(n.appearance.rect.h) === 200)).toBe(true);
    expect(native.some(n => n.appearance.border?.style === "dashed" && Math.round(n.appearance.rect.w) === 500)).toBe(true);

    // Reading the page back and submitting it unchanged changes nothing.
    const again = await mergeHtml(merged, projectPage(merged, merged.pages[0].id));
    const shape = (n: { id: string; appearance: unknown }) => [n.id, n.appearance];
    expect(again.pages[0].nodes.map(shape)).toEqual(merged.pages[0].nodes.map(shape));
  }, 30000);
});
