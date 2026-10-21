// @vitest-environment node
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { createCanvas } from "@napi-rs/canvas";
import { buildDeck } from "./helpers/deck";
import { apiDeckParts } from "./helpers/api-deck";
import { importDesign } from "../src/design/import";
import { exportDesign } from "../src/design/export";
import { projectPage } from "../src/design/projection";
import { cloneDesign, DesignDocument, MeasuredPage, mergeMeasuredPage, nodeText } from "../src/design/model";
import { HeadlessSession } from "../src/headless/session";

async function fixture() {
  const zip = await JSZip.loadAsync(await buildDeck(apiDeckParts()));
  zip.file("customXml/item1.xml", '<private-data owner="keep">business metadata</private-data>');
  const raw = await zip.file("ppt/slides/slide1.xml")!.async("string");
  zip.file("ppt/slides/slide1.xml", raw.replace("</p:sld>", '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite"/></p:par></p:tnLst></p:timing></p:sld>'));
  const bytes = await zip.generateAsync({ type: "uint8array" });
  return { bytes, design: await importDesign(bytes) };
}
const measured = (d: DesignDocument): MeasuredPage => ({ pageId: d.pages[0].id, revision: d.revision, nodes: d.pages[0].nodes.filter(n => !n.hidden).map(n => ({ id: n.id, kind: n.kind, parentId: n.parentId, appearance: cloneDesign(n.appearance) })) });

describe("design layer identity and preservation", () => {
  it("imports stable IDs, explicit groups and an HTML projection", async () => {
    const { bytes, design } = await fixture();
    const repeated = await importDesign(bytes);
    expect(repeated.pages[0].nodes.map(n => n.id)).toEqual(design.pages[0].nodes.map(n => n.id));
    const group = design.pages[0].nodes.find(n => n.name === "数据组")!;
    const child = design.pages[0].nodes.find(n => n.name === "指标甲")!;
    expect(group.kind).toBe("group");
    expect(child.parentId).toBe(group.id);
    const html = new DOMParser().parseFromString(projectPage(design, design.pages[0].id), "text/html");
    const element = [...html.querySelectorAll("[data-id]")].find(n => n.getAttribute("data-id") === child.id)!;
    expect(element.parentElement!.getAttribute("data-id")).toBe(group.id);
    expect(design.source.pptxBase64).toBe(Buffer.from(bytes).toString("base64"));
  });

  it("merges appearance without losing bindings, metadata or group membership", async () => {
    const { design } = await fixture();
    const card = design.pages[0].nodes.find(n => n.name === "卡片")!;
    expect(card.kind).toBe("text");
    card.binding = { semantic: "summary", futureField: { x: 1 } };
    card.metadata = { ...card.metadata, futureUnknown: ["keep"] };
    const input = measured(design), changed = input.nodes.find(n => n.id === card.id)!;
    changed.appearance.text!.paragraphs[0].runs[0].text = "更新后的总结";
    const result = mergeMeasuredPage(design, input), after = result.pages[0].nodes.find(n => n.id === card.id)!;
    expect(nodeText(after)).toBe("更新后的总结");
    expect(after.binding).toEqual(card.binding);
    expect(after.metadata).toEqual(card.metadata);
    expect(nodeText(card)).toBe("增长 30%");
    expect(result.revision).toBe(1);
  });

  it("rejects stale updates, duplicate IDs, reparenting and removal/modification of locked objects atomically", async () => {
    const { design } = await fixture();
    const card = design.pages[0].nodes.find(n => n.name === "卡片")!;
    card.lock = "all";
    const snapshot = JSON.stringify(design);
    const changed = measured(design); changed.nodes.find(n => n.id === card.id)!.appearance.rect.x++;
    expect(() => mergeMeasuredPage(design, changed)).toThrow("Locked");
    expect(() => mergeMeasuredPage(design, { ...measured(design), revision: 999 })).toThrow("STALE");
    const deleted = measured(design); deleted.nodes = deleted.nodes.filter(n => n.id !== card.id);
    expect(() => mergeMeasuredPage(design, deleted)).toThrow("removed");
    const duplicated = measured(design); duplicated.nodes.push(duplicated.nodes[0]);
    expect(() => mergeMeasuredPage(design, duplicated)).toThrow("Duplicate");
    const reparent = measured(design); reparent.nodes.find(n => n.parentId)!.parentId = undefined;
    expect(() => mergeMeasuredPage(design, reparent)).toThrow("membership");
    expect(JSON.stringify(design)).toBe(snapshot);
  });

  it("reads current template reference semantics and can replace complex reference artwork with an image", async () => {
    const { bytes } = await fixture();
    const zip = await JSZip.loadAsync(bytes);
    const xml = new DOMParser().parseFromString(await zip.file("ppt/slides/slide1.xml")!.async("string"), "application/xml");
    const nv = [...xml.getElementsByTagNameNS("*", "cNvPr")].find(n => n.getAttribute("name") === "卡片")!;
    nv.setAttribute("name", "@内容.总结图|用途=参考图|说明=按当前主题重画");
    nv.parentElement!.parentElement!.getElementsByTagNameNS("*", "prstGeom")[0].setAttribute("prst", "ellipse");
    zip.file("ppt/slides/slide1.xml", new XMLSerializer().serializeToString(xml));
    const design = await importDesign(await zip.generateAsync({ type: "uint8array" }));
    const node = design.pages[0].nodes.find(n => n.name.startsWith("@内容.总结图"))!;
    expect(node.kind).toBe("asset");
    expect(node.binding).toMatchObject({ reference: true });
    expect(node.lock).toBeUndefined();
    const canvas = createCanvas(20, 20); canvas.getContext("2d").fillRect(0, 0, 20, 20);
    node.appearance.src = `data:image/png;base64,${canvas.toBuffer("image/png").toString("base64")}`;
    const reopened = await importDesign(await exportDesign(design));
    expect(reopened.pages[0].nodes.find(n => n.id === node.id)?.kind).toBe("image");
    const deletion = measured(design); deletion.nodes = deletion.nodes.filter(n => n.id !== node.id);
    expect(mergeMeasuredPage(design, deletion).pages[0].nodes.some(n => n.id === node.id)).toBe(false);
  }, 20000);

  it("preserves untouched XML and custom package parts through export", async () => {
    const { bytes, design } = await fixture();
    const exported = await exportDesign(design), original = await JSZip.loadAsync(bytes), zip = await JSZip.loadAsync(exported);
    expect(await zip.file("customXml/item1.xml")!.async("string")).toBe(await original.file("customXml/item1.xml")!.async("string"));
    expect(await zip.file("ppt/slides/slide1.xml")!.async("string")).toBe(await original.file("ppt/slides/slide1.xml")!.async("string"));
  });

  it("writes changed Chinese text and grouped geometry without losing placeholders or timing", async () => {
    const { design } = await fixture();
    const title = design.pages[0].nodes.find(n => n.name === "标题 1")!;
    const child = design.pages[0].nodes.find(n => n.name === "指标甲")!;
    title.appearance.text!.paragraphs[0].runs[0].text = "设计层更新";
    child.appearance.rect.x += 40;
    const bytes = await exportDesign(design), zip = await JSZip.loadAsync(bytes);
    const raw = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect(raw).toContain("设计层更新");
    expect(raw).toContain('type="title"');
    expect(raw).toContain("p:timing");
    const reopened = await importDesign(bytes);
    expect(nodeText(reopened.pages[0].nodes.find(n => n.name === "标题 1")!)).toBe("设计层更新");
    expect(reopened.pages[0].nodes.find(n => n.name === "指标甲")!.appearance.rect.x).toBeCloseTo(child.appearance.rect.x, 3);
    expect(reopened.pages[0].nodes.find(n => n.name === "指标甲")!.parentId).toBe(child.parentId);
  });

  it("adds native editable text and a group, exports and reopens them", async () => {
    const { design } = await fixture();
    const p = design.pages[0];
    const style = cloneDesign(p.nodes.find(n => n.name === "卡片")!.appearance.text!);
    style.paragraphs[0].runs[0].text = "新的设计文字";
    p.nodes.push({ id: "new:g", name: "HTML 组合", kind: "group", appearance: { rect: { x: 400, y: 100, w: 400, h: 100 } } });
    p.nodes.push({ id: "new:t", name: "HTML 文字", kind: "text", parentId: "new:g", appearance: { rect: { x: 400, y: 100, w: 400, h: 100 }, text: style, fill: "#eaf2ee" } });
    const session = await HeadlessSession.open(await exportDesign(design), { quiet: true });
    try {
      const outline = await session.outline();
      const group = outline.slides[0].shapes.find(n => n.name === "HTML 组合")!;
      expect(group.children?.find(n => n.name === "HTML 文字")?.text).toBe("新的设计文字");
    } finally { session.dispose(); }
  });
});
