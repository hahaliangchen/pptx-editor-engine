// @vitest-environment node
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { HeadlessSession } from "../src/headless/session";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { importDesign } from "../src/design/import";
import { exportDesign } from "../src/design/export";
import { projectPage } from "../src/design/projection";
import { mergeHtml, measureHtml } from "../src/design/measure";
import { nodeText } from "../src/design/model";

const content = "这是一段通过浏览器排版的中文内容，文字换行和组合关系应当保留。";
const compose = `<div data-kind="group" style="position:absolute;left:100px;top:100px;width:700px;height:250px;display:flex;gap:24px">
  <div data-kind="text" style="width:338px;height:250px;background:#eaf2ee;padding:24px;display:flex;flex-direction:column;justify-content:center">
    <p style="margin:0;line-height:32px;font-family:'Noto Sans CJK SC';font-size:22px;color:#183d32">${content}</p>
  </div>
  <div data-kind="text" style="width:338px;height:250px;background:#183d32;padding:24px;display:flex;flex-direction:column;justify-content:center">
    <p style="margin:0;line-height:32px;font-family:'Noto Sans CJK SC';font-size:22px;color:#ffffff">保留 <strong>富文本</strong> 与可编辑文字</p>
  </div>
</div>`;
const blank = async () => importDesign(readFileSync("tools/templates/blank.pptx"));

describe("HTML design measurement", () => {
  it("measures flex and Chinese wrapping, merges IDs, then exports editable grouped text", async () => {
    const design = await blank();
    const html = projectPage(design, design.pages[0].id).replace("</body>", `${compose}</body>`);
    const merged = await mergeHtml(design, html);
    const nodes = merged.pages[0].nodes.filter(n => n.id.startsWith("new:"));
    expect(nodes).toHaveLength(3);
    expect(nodes[1].parentId).toBe(nodes[0].id);
    expect(nodes[1].appearance.rect.x).toBe(100);
    expect(nodes[2].appearance.rect.x).toBe(462);
    expect(nodeText(nodes[1]).replace(/\u2028/g, "")).toBe(content);
    expect(nodeText(nodes[1])).toContain("\u2028");
    const reopened = await importDesign(await exportDesign(merged));
    const text = reopened.pages[0].nodes.find(n => nodeText(n).replace(/\u2028/g, "") === content)!;
    expect(text).toBeDefined();
    expect(text.parentId).toBeDefined();
    expect(text.appearance.rect).toEqual(nodes[1].appearance.rect);
    expect(reopened.pages[0].nodes.some(n => n.appearance.text?.paragraphs.some(p => p.runs.some(r => r.bold && r.text.includes("富文本"))))).toBe(true);
  }, 20000);

  it("renders measured hard line breaks as two lines and retains fractional font sizes", async () => {
    const design = await blank();
    const html = projectPage(design, design.pages[0].id).replace("</body>", '<div data-kind="text" style="position:absolute;left:100px;top:100px;width:500px;height:160px"><p style="margin:0;font-family:Arial;font-size:25px;line-height:50px;color:#000000">Alpha<br>Beta</p></div></body>');
    const merged = await mergeHtml(design, html);
    const bytes = await exportDesign(merged);
    const reimported = await importDesign(bytes);
    const node = reimported.pages[0].nodes.find(n => nodeText(n).includes("Alpha"))!;
    expect(nodeText(node)).toBe("Alpha\u2028Beta");
    expect(node.appearance.text!.paragraphs[0].runs[0].fontSize).toBeCloseTo(25, 4);
    const session = await HeadlessSession.open(bytes, { quiet: true });
    try {
      const png = await session.render(1, { width: 1280 });
      const canvas = createCanvas(1280, 720), ctx = canvas.getContext("2d");
      ctx.drawImage(await loadImage(png), 0, 0);
      const pixels = ctx.getImageData(100, 100, 500, 160).data;
      const rows: number[] = [];
      for (let y = 0; y < 160; y++) {
        let dark = 0;
        for (let x = 0; x < 500; x++) { const i = (y * 500 + x) * 4; if (pixels[i] < 100 && pixels[i+1] < 100 && pixels[i+2] < 100 && pixels[i+3] > 200) dark++; }
        if (dark > 2) rows.push(y);
      }
      const starts = rows.filter((y, i) => i === 0 || y - rows[i-1] > 8);
      expect(starts).toHaveLength(2);
      expect(starts[1] - starts[0]).toBeGreaterThan(40);
    } finally { session.dispose(); }
  }, 20000);

  it("renders unchanged projection without changing its JSON or source PPTX", async () => {
    const design = await blank();
    const result = await mergeHtml(design, projectPage(design, design.pages[0].id));
    expect(result).toEqual(design);
    expect(await exportDesign(result)).toEqual(Buffer.from(design.source.pptxBase64, "base64"));
  }, 20000);

  it("does not allow removing HTML lock attributes to bypass JSON locks", async () => {
    const design = await blank();
    design.pages[0].nodes.push({ id: "new:locked", kind: "shape", name: "Locked brand mark", lock: "all", appearance: { rect: { x: 50, y: 50, w: 100, h: 100 }, fill: "#183d32" } });
    const html = projectPage(design, design.pages[0].id).replace('data-locked="all"', "").replace("left:50px;top:50px", "left:60px;top:50px");
    await expect(mergeHtml(design, html)).rejects.toThrow(/Locked|Read-only HTML attribute/);
    const original = projectPage(design,design.pages[0].id);
    await expect(mergeHtml(design,original.replace("<html>",'<html style="opacity:0">'))).rejects.toThrow("Unsupported CSS");
    await expect(mergeHtml(design,original.replace("<html>",'<html style="clip-path:inset(0 0 700px 0)">'))).rejects.toThrow("Unsupported CSS");
    await expect(mergeHtml(design,original.replace("</body>",'<div data-kind="shape" data-purpose="reference" style="position:absolute;left:800px;top:400px;width:100px;height:100px;background:red"></div></body>'))).rejects.toThrow("New AI nodes use content purpose");
  }, 20000);

  it("rejects clipped text and unsupported effects instead of dropping them", async () => {
    const design = await blank(), base = projectPage(design, design.pages[0].id);
    await expect(measureHtml(base.replace("</body>", '<div data-kind="text" style="width:300px;height:10px"><p style="font-size:40px">不能裁切</p></div></body>'))).rejects.toThrow(/clipped|outside/);
    await expect(measureHtml(base.replace("</body>", '<div data-kind="shape" style="width:100px;height:100px;filter:blur(2px)"></div></body>'))).rejects.toThrow("Unsupported CSS");
  }, 20000);
});
