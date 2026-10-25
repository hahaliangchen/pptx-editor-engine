// @vitest-environment node
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { importDesign } from "../src/design/import";
import { projectPage } from "../src/design/projection";
import { mergeHtml } from "../src/design/measure";
import { nodeText } from "../src/design/model";

// Project 816: the agent copied a page and put data-paragraph on the run as well
// as on its paragraph. The browser measurement must read the words once.
const marker = `<div data-kind="text" style="position:absolute;left:496px;top:182px;width:48px;height:54px;display:flex;flex-direction:column;justify-content:center;padding:0px">`
  + `<p data-paragraph="true" style="font-family:'Noto Sans SC';font-size:34px;font-weight:700;margin:0;text-align:center;line-height:41px">`
  + `<span data-paragraph="true" style="font-family:'Noto Sans SC';font-size:34px;color:#ffffff;font-weight:700">一</span></p></div>`;

describe("HTML paragraph structure", () => {
  it("reads a paragraph marker nested inside another paragraph once", async () => {
    const design = await importDesign(readFileSync("tools/templates/blank.pptx"));
    const html = projectPage(design, design.pages[0].id).replace("</body>", `${marker}</body>`);
    const merged = await mergeHtml(design, html);
    const node = merged.pages[0].nodes.find(n => n.id.startsWith("new:"))!;
    expect(node.appearance.text!.paragraphs).toHaveLength(1);
    expect(nodeText(node)).toBe("一");
    expect(merged.changes).toContainEqual(expect.objectContaining({ nodeId: node.id, action: "add", text: { after: "一" } }));
    expect(merged.diagnostics.some(d => d.message.includes("Repeated paragraph"))).toBe(false);
  }, 30000);
});
