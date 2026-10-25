// @vitest-environment node
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { importDesign } from "../src/design/import";
import { exportDesign } from "../src/design/export";
import { mergeHtml, measureHtml } from "../src/design/measure";
import { projectPage } from "../src/design/projection";
import { materializeDesign, convertDesign } from "../src/design/service";
import { cloneDesign, nodeText, type DesignDocument } from "../src/design/model";

async function fixture(): Promise<DesignDocument> {
  const design = await importDesign(readFileSync("tools/templates/blank.pptx"));
  const text = (id: string, value: string, y: number) => ({ id, kind: "text" as const, name: id,
    appearance: { rect: { x: 81.12345, y, w: 600.23456, h: 100.34567 }, fill: "#eef4ef", text: {
      padding: [12.12345, 20.23456, 14.34567, 22.45678] as [number, number, number, number], verticalAlign: "middle" as const,
      paragraphs: [{ runs: [{ text: value, fontFamily: "Arial", fontSize: 25.23456, color: "#123456", bold: false, italic: false, letterSpacing: 0.12345 }], align: "left" as const, lineHeight: 35.23456, before: 1.12345, after: 2.23456 }]
    } } });
  design.pages[0].nodes.push(text("new:title", "Original title", 80.12345), text("new:untouched", "Leave this alone", 260.23456));
  return materializeDesign(design);
}

const cleanChanges = (value: DesignDocument) => { const result = cloneDesign(value); delete result.changes; return result; };

describe("HTML attribute differences", () => {
  it("keeps native two-line CJK title typography when only its wording changes", async () => {
    const design=await importDesign(readFileSync("tests/fixtures/unicom-manual-title.pptx"));
    const page=design.pages[0], before=page.nodes.find(n=>n.origin?.shapeId==="10")!;
    expect(before.appearance.rect).toEqual({x:70,y:165,w:1000,h:110});
    const html=projectPage(design,page.id);
    const edited=html.replace("为什么总加班？","加班为何集中？").replace("先看交付机制如何“推迟问题”","把问题暴露在交付前");
    const parsed=new DOMParser().parseFromString(edited,"text/html");
    const paragraph=Array.from(parsed.querySelectorAll<HTMLElement>("[data-id]")).find(n=>n.dataset.id===before.id)!.querySelector("p")!;
    expect(paragraph.style.fontSize).toBe("42px");expect(paragraph.style.lineHeight).toBe("56px");
    const measured=await measureHtml(edited,undefined,true);
    expect(measured.issues?.filter(i=>i.nodeId===before.id)).toEqual([]);
    const ink=measured.textPaint![before.id];expect(Math.max(...ink.map(r=>r.y+r.h))).toBeLessThanOrEqual(275);
    const changed=await mergeHtml(design,edited), expected=cloneDesign(before);
    for(const run of expected.appearance.text!.paragraphs[0].runs) run.text=run.text.replace("为什么总加班？","加班为何集中？").replace("先看交付机制如何“推迟问题”","把问题暴露在交付前");
    expect(changed.pages[0].nodes.find(n=>n.id===before.id)).toEqual(expected);
    expect(changed.pages[0].nodes.filter(n=>n.id!==before.id)).toEqual(page.nodes.filter(n=>n.id!==before.id));
    expect(changed.changes!.flatMap(c=>c.fields).every(field=>/^appearance\.text\.paragraphs\.0\.runs\.\d+\.text$/.test(field))).toBe(true);
    // Clipping remains visible as a warning; no fitting or box enlargement is performed.
    const clipped=await mergeHtml(design,edited.replace("height:110px","height:50px"));
    expect(clipped.diagnostics.some(d=>d.severity==="warning" && /clipped|outside/.test(d.message))).toBe(true);
  },30000);

  it("changes one word without replacing native geometry, margins, styles or another node's XML", async () => {
    const design = await fixture(), html = projectPage(design, design.pages[0].id);
    const before = design.pages[0].nodes.find(n => nodeText(n) === "Original title")!;
    const untouched = design.pages[0].nodes.find(n => nodeText(n) === "Leave this alone")!;
    const changed = await mergeHtml(design, html.replace("Original title", "Updated title"));
    const expected = cloneDesign(before); expected.appearance.text!.paragraphs[0].runs[0].text = "Updated title";
    expect(changed.pages[0].nodes.find(n => n.id === before.id)).toEqual(expected);
    expect(changed.pages[0].nodes.find(n => n.id === untouched.id)).toEqual(untouched);
    expect(changed.changes).toEqual([{ pageId: design.pages[0].id, nodeId: before.id, action: "update", fields: ["appearance.text.paragraphs.0.runs.0.text"],
      text: { before: "Original title", after: "Updated title" } }]);
    const saved = await materializeDesign(changed);
    expect(saved.pages[0].nodes.find(n => n.id === untouched.id)?.appearance).toEqual(untouched.appearance);
    const originalZip = await JSZip.loadAsync(Buffer.from(design.source.pptxBase64, "base64"));
    const updatedZip = await JSZip.loadAsync(await exportDesign(saved));
    const element = (xml: string) => {
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      return Array.from(doc.getElementsByTagNameNS("*", "cNvPr")).find(n => n.getAttribute("id") === untouched.origin!.shapeId)!.parentElement!.parentElement!.outerHTML;
    };
    expect(element(await updatedZip.file(untouched.origin!.part)!.async("string"))).toBe(element(await originalZip.file(untouched.origin!.part)!.async("string")));
  }, 15000);

  it("updates a CSS color without normalizing typography or geometry", async () => {
    const design = await fixture(), before = design.pages[0].nodes.find(n => nodeText(n) === "Original title")!;
    const html = projectPage(design, design.pages[0].id).replace("color:#123456", "color:#aa3300");
    const changed = await mergeHtml(design, html), expected = cloneDesign(before);
    expected.appearance.text!.paragraphs[0].runs[0].color = "#aa3300";
    expect(changed.pages[0].nodes.find(n => n.id === before.id)).toEqual(expected);
    expect(changed.changes![0].fields).toEqual(["appearance.text.paragraphs.0.runs.0.color"]);
  }, 15000);

  it("leaves unrelated overflow alone and reports edited overflow without rejecting it", async () => {
    const design = await fixture();
    const untouched = design.pages[0].nodes.find(n => nodeText(n) === "Leave this alone")!;
    untouched.appearance.rect.h = 5; // existing bad manual layout, not part of this request
    const html = projectPage(design, design.pages[0].id);
    const result = await mergeHtml(design, html.replace("Original title", "Updated title"));
    expect(result.pages[0].nodes.find(n => n.id === untouched.id)).toEqual(untouched);
    const changed=await mergeHtml(design, html.replace("Leave this alone", "Edited overflow"));
    expect(changed.diagnostics.some(d=>d.nodeId===untouched.id && d.severity==="warning" && /clipped|outside/.test(d.message))).toBe(true);
  }, 15000);

  it("detects shared stylesheet changes and flex layout dependencies", async () => {
    let design = await importDesign(readFileSync("tools/templates/blank.pptx"));
    let html = projectPage(design, design.pages[0].id).replace("</body>", '<div data-kind="group" data-id="new:flow" style="position:absolute;left:80px;top:80px;width:1100px;height:300px;display:flex;gap:20px"><div data-kind="shape" data-id="new:left" style="width:100px;height:100px;background:#008866"></div><div data-kind="shape" data-id="new:right" style="width:100px;height:100px;background:#008866"></div></div></body>');
    design = await materializeDesign(await mergeHtml(design, html));
    const parsed = new DOMParser().parseFromString(html, "text/html");
    parsed.body.dataset.sourceHash = design.source.sha256; parsed.body.dataset.designRevision = String(design.revision);
    html = '<!doctype html>' + parsed.documentElement.outerHTML;
    const changedHtml = html.replace("</head>", '<style>[data-id="new:left"]{width:200px!important}</style></head>');
    const changed = await mergeHtml(design, changedHtml, undefined, html);
    expect(changed.pages[0].nodes.find(n => n.id === "new:left")!.appearance.rect.w).toBe(200);
    expect(changed.pages[0].nodes.find(n => n.id === "new:right")!.appearance.rect.x).toBe(300);
    expect(changed.changes?.map(c => c.nodeId)).toEqual(["new:left", "new:right"]);
  }, 15000);

  it("does not rewrite a PPTX for semantically unchanged HTML or relax a lock", async () => {
    const design = await fixture(), html = projectPage(design, design.pages[0].id);
    const result = await convertDesign({ operation: "apply", design, expectedRevision: design.revision, html: html.replace("<head>", "<head><!-- formatting only -->") });
    expect(cleanChanges(result)).toEqual(cleanChanges(design));
    design.pages[0].nodes.find(n => nodeText(n) === "Original title")!.lock = "all";
    await expect(mergeHtml(design, projectPage(design, design.pages[0].id).replace("Original title", "Bypass lock"))).rejects.toThrow(/Locked/);
  }, 15000);
});
