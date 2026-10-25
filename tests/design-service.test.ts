// @vitest-environment node
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { convertDesign, syncDesign } from "../src/design/service";
import { projectPage } from "../src/design/projection";
import { nodeText } from "../src/design/model";
import { HeadlessSession } from "../src/headless/session";
import { AutoSaver } from "../src/page/autosave";
import { DesignBridge } from "../src/page/design-bridge";

const base = readFileSync("tools/templates/blank.pptx");
const content = '<div data-kind="text" style="position:absolute;left:80px;top:80px;width:800px;height:100px"><p style="font-family:Arial;font-size:36px;line-height:48px;margin:0">Design title</p></div>';
const result = (r: any) => { expect(r.ok, JSON.stringify(r)).toBe(true); return JSON.parse(r.content[0].text); };

describe("bidirectional design service", () => {
  it("adds HTML, edits in the native editor, projects fresh HTML and exports again with stable IDs", async () => {
    const first = await syncDesign(base);
    const generated = await convertDesign({ operation: "apply", design: first, expectedRevision: 0,
      html: projectPage(first, first.pages[0].id).replace("</body>", content + "</body>") });
    const title = generated.pages[0].nodes.find(n => nodeText(n) === "Design title")!;
    expect(title.origin?.shapeId).toBeDefined();
    const session = await HeadlessSession.open(Buffer.from(generated.source.pptxBase64, "base64"), { quiet: true });
    let bytes: Buffer;
    try {
      const edited = await session.apply({ commands: [{ op: "replaceText", target: { slide: 1, text: "Design title" }, find: "Design title", replace: "Manually revised" }] });
      expect(edited.ok, JSON.stringify(edited)).toBe(true);
      bytes = await session.export();
    } finally { session.dispose(); }
    const refreshed = await syncDesign(bytes!, generated);
    expect(refreshed.pages[0].nodes.find(n => nodeText(n) === "Manually revised")!.id).toBe(title.id);
    expect(projectPage(refreshed, refreshed.pages[0].id)).toContain("Manually revised");
    await expect(convertDesign({ operation: "apply", design: refreshed, expectedRevision: refreshed.revision,
      html: projectPage(generated, generated.pages[0].id) })).resolves.toBeDefined();
    const final = await convertDesign({ operation: "apply", design: refreshed, expectedRevision: refreshed.revision,
      html: projectPage(refreshed, refreshed.pages[0].id).replace("Manually revised", "Final design") });
    const reopened = await syncDesign(Buffer.from(final.source.pptxBase64, "base64"));
    expect(reopened.pages[0].nodes.some(n => nodeText(n) === "Final design")).toBe(true);
  }, 30000);

  it("reflows a native page through Grid/Flex after a manual move without stored layout source", async () => {
    const first = await syncDesign(base);
    const layout = `<div style="position:absolute;left:80px;top:100px;width:1000px;display:grid;grid-template-columns:1fr 1fr;gap:40px">
      <div style="display:flex;flex-direction:column;gap:16px">
        <div data-kind="text" data-id="new:body" style="width:220px"><p style="font-family:Arial;font-size:24px;line-height:32px">First line<br>Second line</p></div>
        <div data-kind="shape" data-id="new:marker" style="width:24px;height:40px;background:#e60012"></div>
      </div>
      <div data-kind="text" data-id="new:peer"><p style="font-family:Arial;font-size:24px;line-height:32px">Other column</p></div>
    </div>`;
    const generated = await convertDesign({ operation:"apply", design:first, expectedRevision:first.revision,
      html:projectPage(first, first.pages[0].id).replace("</body>", layout + "</body>") });
    const page = generated.pages[0];
    const body = page.nodes.find(n => n.id === "new:body")!;
    const marker = page.nodes.find(n => n.id === "new:marker")!;
    const peer = page.nodes.find(n => n.id === "new:peer")!;
    expect(body.appearance.rect.x).toBeCloseTo(80, 1);
    expect(peer.appearance.rect.x).toBeCloseTo(600, 1);
    expect(marker.appearance.rect.y).toBeGreaterThanOrEqual(body.appearance.rect.y + body.appearance.rect.h + 15);
    const session = await HeadlessSession.open(Buffer.from(generated.source.pptxBase64, "base64"), {quiet:true});
    let moved: Buffer;
    try {
      const edited = await session.apply({commands:[{op:"setBounds",target:{slide:1,id:marker.origin!.shapeId},x:12,y:9}]});
      expect(edited.ok, JSON.stringify(edited)).toBe(true);
      moved = await session.export();
    } finally { session.dispose(); }
    const refreshed = await syncDesign(moved!, generated);
    const absolute = projectPage(refreshed, refreshed.pages[0].id);
    const parsed = new DOMParser().parseFromString(absolute, "text/html");
    const nodes = ["new:body","new:marker","new:peer"].map(id => Array.from(parsed.querySelectorAll<HTMLElement>("[data-id]")).find(n => n.dataset.id === id)!);
    expect(nodes.every(n => n.style.position === "absolute")).toBe(true);
    const root = parsed.createElement("div");
    root.style.cssText = "position:absolute;left:80px;top:100px;width:1000px;display:grid;grid-template-columns:1fr 1fr;gap:40px";
    const stack = parsed.createElement("div"); stack.style.cssText = "display:flex;flex-direction:column;gap:24px";
    for (const node of nodes) { for (const key of ["position","left","top","width","height"]) node.style.removeProperty(key); }
    nodes[1].style.cssText += ";width:24px;height:40px";
    stack.append(nodes[0],nodes[1]); root.append(stack,nodes[2]); parsed.body.append(root);
    const reflowed = await convertDesign({operation:"apply", design:refreshed, expectedRevision:refreshed.revision,
      html:"<!doctype html>" + parsed.documentElement.outerHTML});
    const after = reflowed.pages[0].nodes;
    const b = after.find(n => n.id === body.id)!, m = after.find(n => n.id === marker.id)!;
    expect(m.appearance.rect.y).toBeGreaterThanOrEqual(b.appearance.rect.y + b.appearance.rect.h + 23);
    const reopened = await syncDesign(Buffer.from(reflowed.source.pptxBase64,"base64"));
    expect(reopened.pages[0].nodes.some(n => nodeText(n).includes("Second line"))).toBe(true);
  }, 30000);

  it("supports a 20-page deck, deletion and reordering without layout generation", async () => {
    let design = await syncDesign(base);
    design = await convertDesign({ operation: "pages", design, expectedRevision: design.revision, commands: Array.from({ length: 19 }, () => ({ op: "addSlide" as const })) });
    expect(design.pages).toHaveLength(20);
    const lastId = design.pages[19].id;
    design = await convertDesign({ operation: "pages", design, expectedRevision: design.revision, commands: [{ op: "moveSlide", slide: 20, to: 1 }, { op: "deleteSlide", slide: 20 }] });
    expect(design.pages).toHaveLength(19);
    expect(design.pages[0].id).toBe(lastId);
    expect((await syncDesign(Buffer.from(design.source.pptxBase64, "base64"))).pages).toHaveLength(19);
  }, 30000);

  it("saves only PPTX after an HTML replacement and can reopen without a saved design JSON", async () => {
    let binary: Uint8Array = base;
    const viewer = { snapshotPptx: async () => new Blob([binary as BlobPart]), loadPptx: async (buffer: ArrayBuffer) => { binary = new Uint8Array(buffer); }, getCurrentSlideIndex: () => 0, getSlidesCount: () => 1, selectSlide: async () => {} };
    const bridge = new DesignBridge(viewer, convertDesign);
    const start = result(await bridge.run("get_design", { page: 1 }));
    result(await bridge.run("apply_html", { expected_design_revision: start.design_revision, html: start.html.replace("</body>", content + "</body>") }));
    expect(bridge.pending).toBe(true);
    const persisted: Blob[] = [];
    const saver = new AutoSaver({ exportPptx: viewer.snapshotPptx, hasUnsavedEdits: () => bridge.pending }, {
      save: async (blob, expected, source) => { expect(source).toBe("agent"); persisted.push(blob); bridge.saved(); return expected + 1; }
    }, { revision: 7 });
    expect(await saver.saveNow("agent")).toBe(8);
    expect(persisted).toHaveLength(1);
    expect(bridge.pending).toBe(false);
    // A fresh page gets no JSON state from the previous page.
    const reopened = new DesignBridge(viewer, convertDesign);
    expect(result(await reopened.run("get_design", { page: 1 })).html).toContain("Design title");
    saver.dispose();
  }, 20000);

  it("applies the submitted edit without an intervening version gate", async () => {
    let release!: () => void;
    let started!: () => void; const converting = new Promise<void>(resolve => { started = resolve; });
    let loads = 0;
    const viewer = { snapshotPptx: async () => new Blob([base]), loadPptx: async () => { loads++; }, getCurrentSlideIndex: () => 0, getSlidesCount: () => 1, selectSlide: async () => {} };
    const bridge = new DesignBridge(viewer, async request => {
      if (request.operation === "apply") await new Promise<void>(resolve => { release = resolve; started(); });
      return convertDesign(request);
    });
    const before = result(await bridge.run("get_design", { page: 1 }));
    const pending = bridge.run("apply_html", { expected_design_revision: before.design_revision, html: before.html.replace("</body>", content + "</body>") });
    await converting; bridge.edited(); release();
    expect((await pending).ok).toBe(true);
    expect(loads).toBe(1);
    expect(bridge.pending).toBe(true);
    const latest = result(await bridge.run("get_design", { page: 1 }));
    expect(latest.design_revision).toBeGreaterThan(before.design_revision);
  }, 20000);
});
