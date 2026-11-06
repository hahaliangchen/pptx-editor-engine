// @vitest-environment node
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { convertDesign, syncDesign } from "../src/design/service";
import { projectPage } from "../src/design/projection";
import { HeadlessSession } from "../src/headless/session";
import type { DesignDocument } from "../src/design/model";
import { readNativePaint } from "../src/design/paint";

const base = readFileSync("tools/templates/blank.pptx");
const shapes = `
  <div data-kind="shape" data-id="new:band" style="position:absolute;left:40px;top:40px;width:400px;height:60px;background-image:linear-gradient(90deg, #e60012 0%, #ec4a57 100%)"></div>
  <div data-kind="shape" data-id="new:card" style="position:absolute;left:40px;top:140px;width:300px;height:160px;background-color:#ffffff;border-radius:8px;box-shadow:0px 2px 8px rgba(0, 0, 0, 0.1)"></div>
  <div data-kind="shape" data-id="new:tab" style="position:absolute;left:400px;top:140px;width:300px;height:60px;background-color:#e60012;border-radius:12px 12px 0px 0px"></div>
  <div data-kind="shape" data-id="new:ring" style="position:absolute;left:760px;top:140px;width:48px;height:48px;border:2px dashed #e60012;border-radius:50%"></div>
  <div data-kind="shape" data-id="new:ribbon" style="position:absolute;left:400px;top:260px;width:400px;height:40px;background-color:#c4000f;clip-path:polygon(5% 0%, 95% 0%, 100% 100%, 0% 100%)"></div>
  <div data-kind="text" data-id="new:label" style="position:absolute;left:40px;top:360px;width:400px;height:44px;background-image:linear-gradient(180deg, #e60012, #b8000e);border-radius:6px;padding:8px 14px">
    <p style="font-family:Arial;font-size:20px;line-height:28px;color:#ffffff">Header</p></div>`;
const apply = (design: DesignDocument, html: string) => convertDesign({ operation: "apply", design, expectedRevision: design.revision, html });
const addTo = (design: DesignDocument, extra: string) => projectPage(design, design.pages[0].id).replace("</body>", `${extra}</body>`);
const byName = (design: DesignDocument, id: string) => design.pages[0].nodes.find(n => n.id === id || n.name === id)!;

describe("native gradient, shadow, dash and outline paint", () => {
  it("writes DrawingML, reopens as editable shapes and projects the same CSS", async () => {
    const generated = await apply(await syncDesign(base), addTo(await syncDesign(base), shapes));
    const bytes = Buffer.from(generated.source.pptxBase64, "base64");
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file(generated.pages[0].part)!.async("string");
    expect(xml).toContain('<a:lin ang="0" scaled="0"/>');
    expect(xml).toContain('<a:lin ang="5400000" scaled="0"/>');
    expect(xml).toMatch(/<a:outerShdw blurRad="76200" dist="19050" dir="5400000"[^>]*><a:srgbClr val="000000"><a:alpha val="10000"\/>/);
    expect(xml).toContain('<a:prstDash val="dash"/>');
    expect(xml).toContain('<a:prstGeom prst="round2SameRect"><a:avLst><a:gd name="adj1" fmla="val 20000"/><a:gd name="adj2" fmla="val 0"/>');
    expect(xml).toContain('<a:custGeom>');
    // spPr children keep schema order: geometry, fill, line, effects.
    expect(xml).toMatch(/<a:prstGeom prst="roundRect">[\s\S]*?<\/a:prstGeom><a:solidFill>[\s\S]*?<a:ln>[\s\S]*?<a:effectLst><a:outerShdw/);

    const reopened = await syncDesign(bytes);
    const ids = new Map(generated.pages[0].nodes.map(n => [n.id, n.name]));
    const node = (id: string) => byName(reopened, ids.get(id)!);
    for (const id of ["new:band", "new:card", "new:tab", "new:ring", "new:ribbon"]) expect(node(id).kind).toBe("shape");
    expect(node("new:band").appearance.gradient).toEqual({ angle: 90, stops: [{ color: "#e60012", position: 0 }, { color: "#ec4a57", position: 100 }] });
    expect(node("new:card").appearance.shadow).toEqual({ x: 0, y: 2, blur: 8, color: "rgba(0, 0, 0, 0.1)" });
    expect(node("new:card").appearance.radius).toBeCloseTo(8, 1);
    expect(node("new:tab").appearance.corners).toEqual({ top: 12, bottom: 0 });
    expect(node("new:ring").appearance.border).toEqual({ color: "#e60012", width: 2, style: "dashed" });
    expect(node("new:ribbon").appearance.polygon).toEqual([[5, 0], [95, 0], [100, 100], [0, 100]]);
    expect(node("new:label").kind).toBe("text");
    expect(node("new:label").appearance.gradient?.angle).toBe(180);

    const html = projectPage(reopened, reopened.pages[0].id);
    for (const css of ["linear-gradient(90deg, #e60012 0%, #ec4a57 100%)", "box-shadow:0px 2px 8px rgba(0, 0, 0, 0.1)", "border-radius:12px 12px 0px 0px", "2px dashed #e60012", "clip-path:polygon(5% 0%, 95% 0%, 100% 100%, 0% 100%)"]) expect(html).toContain(css);
    // The projected page measures back to the same appearance: nothing is rewritten.
    const again = await apply(reopened, html);
    expect(again.changes ?? []).toEqual([]);
  }, 60000);

  it("renders the top-only corners and the gradient natively", async () => {
    const generated = await apply(await syncDesign(base), addTo(await syncDesign(base), shapes));
    const session = await HeadlessSession.open(Buffer.from(generated.source.pptxBase64, "base64"), { quiet: true });
    try {
      const image = await loadImage(await session.render(1, { width: 1280 }));
      const ctx = createCanvas(image.width, image.height).getContext("2d");
      ctx.drawImage(image, 0, 0);
      const pixel = (x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data.slice(0, 3));
      // round2SameRect: top-left corner is cut, bottom-left corner stays square.
      expect(pixel(401, 141)).toEqual([255, 255, 255]);
      expect(pixel(401, 198)[1]).toBeLessThan(40);
      // The band runs from #e60012 on the left to #ec4a57 on the right.
      expect(pixel(45, 70)[1]).toBeLessThan(20);
      expect(pixel(435, 70)[1]).toBeGreaterThan(55);
      // The ribbon's slanted end leaves the bottom-left corner filled and the top-left empty.
      expect(pixel(402, 262)).toEqual([255, 255, 255]);
      expect(pixel(402, 298)[1]).toBeLessThan(40);
    } finally { session.dispose(); }
  }, 60000);

  it("keeps a stroked polygon, such as an outlined SVG icon path, out of editable paint", () => {
    const spPr = (line: string) => new DOMParser().parseFromString(`<p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="50" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="E60012"/></a:solidFill>${line}</p:spPr>`, "application/xml").documentElement;
    expect(readNativePaint(spPr(""), { w: 40, h: 40 }, 9525)?.polygon).toHaveLength(3);
    expect(readNativePaint(spPr('<a:ln w="16192"><a:solidFill><a:srgbClr val="E60012"/></a:solidFill></a:ln>'), { w: 40, h: 40 }, 9525)).toBeUndefined();
  });

  it("rejects CSS that has no native equivalent", async () => {
    const first = await syncDesign(base);
    const reject = (style: string) => expect(apply(first, addTo(first, `<div data-kind="shape" style="position:absolute;left:40px;top:40px;width:200px;height:80px;${style}"></div>`))).rejects.toThrow();
    await reject("background-image:radial-gradient(#e60012, #ffffff)");
    await reject("background-color:#ffffff;box-shadow:inset 0px 2px 8px #000000");
    await reject("background-color:#ffffff;box-shadow:0px 2px 8px #000000, 0px 4px 2px #000000");
    await reject("background-color:#ffffff;border-radius:8px 0px 8px 0px");
    await reject("background-color:#e60012;border:1px solid #000000;clip-path:polygon(0% 0%, 100% 0%, 50% 100%)");
    await reject("background-color:#e60012;background-image:linear-gradient(90deg, #e60012, #ffffff)");
  }, 60000);
});
