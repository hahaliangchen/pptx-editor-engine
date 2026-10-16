import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { emptyStructure, exportEditedPptx, markDirty } from "../src/pptx-exporter";
import { pruneDeletedReferences } from "../src/pptx-references";
import { Slide } from "../src/pptx-virtual-dom";
import { checkPackage } from "./helpers/package-check";

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/** One click-effect for each target id, the way PowerPoint nests them. */
const effect = (id: number, spid: string, preset = 1) =>
  `<p:par><p:cTn id="${id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>`
  + `<p:par><p:cTn id="${id + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
  + `<p:par><p:cTn id="${id + 2}" presetID="${preset}" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
  + `<p:set><p:cBhvr><p:cTn id="${id + 3}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
  + `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;

const timingXml = (spids: string[]) =>
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>`
  + `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>`
  + spids.map((spid, i) => effect(10 + i * 10, spid)).join("")
  + `</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst></p:seq>`
  + `</p:childTnLst></p:cTn></p:par></p:tnLst>`
  + `<p:bldLst>${spids.map(spid => `<p:bldP spid="${spid}" grpId="0"/>`).join("")}</p:bldLst></p:timing>`;

const slideXml = (spids: string[]) =>
  `<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`
  + `<p:sp><p:nvSpPr><p:cNvPr id="2" name="A"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`
  + `<p:sp><p:nvSpPr><p:cNvPr id="3" name="B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`
  + `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4" name="C"/><p:cNvCxnSpPr><a:stCxn id="2" idx="0"/><a:endCxn id="3" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>`
  + `</p:spTree></p:cSld>${timingXml(spids)}</p:sld>`;

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");
const spids = (doc: Document) => Array.from(doc.getElementsByTagName("p:spTgt")).map(n => n.getAttribute("spid"));

describe("pruneDeletedReferences", () => {
  it("drops the effect of a deleted shape, with its empty wrappers, and keeps the others", () => {
    const doc = parse(slideXml(["2", "3", "4"]));
    pruneDeletedReferences(doc, new Set(["3"]));
    expect(spids(doc)).toEqual(["2", "4"]);
    expect(Array.from(doc.getElementsByTagName("p:bldP")).map(n => n.getAttribute("spid"))).toEqual(["2", "4"]);
    // main sequence survives, and no wrapper is left empty
    expect(doc.getElementsByTagName("p:cTn")[1].getAttribute("nodeType")).toBe("mainSeq");
    for (const list of Array.from(doc.getElementsByTagName("p:childTnLst"))) expect(list.children.length).toBeGreaterThan(0);
  });

  it("removes the whole timing when no animation is left", () => {
    const doc = parse(slideXml(["2"]));
    pruneDeletedReferences(doc, new Set(["2"]));
    expect(doc.getElementsByTagName("p:timing")).toHaveLength(0);
  });

  it("drops the whole timing when a video's play controls point at the deleted shape", () => {
    const media = `<p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>`
      + `<p:seq><p:cTn id="2" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cond></p:stCondLst>`
      + `<p:childTnLst><p:par><p:cTn id="3"><p:childTnLst><p:par><p:cTn id="5" presetID="2" presetClass="mediacall" nodeType="clickEffect"><p:childTnLst>`
      + `<p:cmd type="call" cmd="togglePause"><p:cBhvr><p:cTn id="6"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:seq>`
      + `<p:video><p:cMediaNode><p:cTn id="8"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cMediaNode></p:video></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const doc = parse(`<p:sld ${NS}><p:cSld><p:spTree/></p:cSld>${media}</p:sld>`);
    pruneDeletedReferences(doc, new Set(["7"]));
    expect(doc.getElementsByTagName("p:timing")).toHaveLength(0);
  });

  it("lets go of connector ends glued to a deleted shape", () => {
    const doc = parse(slideXml(["4"]));
    pruneDeletedReferences(doc, new Set(["2"]));
    expect(doc.getElementsByTagName("a:stCxn")).toHaveLength(0);
    expect(doc.getElementsByTagName("a:endCxn")).toHaveLength(1);
  });

  it("does nothing when nothing was deleted", () => {
    const xml = slideXml(["2", "3"]);
    const doc = parse(xml);
    pruneDeletedReferences(doc, new Set());
    expect(spids(doc)).toEqual(["2", "3"]);
  });
});

describe("export after deleting an animated shape", () => {
  it("leaves no animation pointing at the deleted shape", async () => {
    const part = "ppt/slides/slide1.xml";
    const zip = new JSZip();
    zip.file(part, slideXml(["2", "3"]));
    const dirty = new Map();
    markDirty(dirty, { part, shapeId: "3", xform: { scaleX: 1, scaleY: 1, offsetX: 0, offsetY: 0 } });
    const source = (id: string) => ({ part, shapeId: id, xform: { scaleX: 1, scaleY: 1, offsetX: 0, offsetY: 0 } });
    const slides: Slide[] = [{ id: "s", rawXml: "", elements: [
      { type: "shape", id: "a", source: source("2"), rect: { x: 0, y: 0, w: 1, h: 1 }, shapeType: "rect", fill: "#fff" }
    ] }];
    const blob = await exportEditedPptx(zip, slides, dirty, new Map(), new Map(), emptyStructure());
    const out = await JSZip.loadAsync(await new Promise<ArrayBuffer>(resolve => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob);
    }));
    const xml = await out.file(part)!.async("string");
    expect(xml).not.toContain('spid="3"');
    expect(xml).toContain('spid="2"');
    expect(xml).not.toContain('id="3" idx="1"');
  });
});

describe("checkPackage", () => {
  it("flags a dangling relationship, an undeclared part and a duplicated shape id", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="x"/></Types>`);
    zip.file("ppt/presentation.xml", `<p:presentation ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`);
    zip.file("ppt/_rels/presentation.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="t/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="t/theme" Target="theme/missing.xml"/></Relationships>`);
    zip.file("ppt/slides/slide1.xml", `<p:sld ${NS}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="a"/></p:nvSpPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="2" name="b"/></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sld>`);
    zip.file("ppt/media/pic.png", new Uint8Array([1]));
    const problems = await checkPackage(await zip.generateAsync({ type: "uint8array" }));
    expect(problems).toEqual(expect.arrayContaining([
      expect.stringContaining("missing ppt/theme/missing.xml"),
      expect.stringContaining("no content type for ppt/media/pic.png"),
      expect.stringContaining("shape id 2 used twice")
    ]));
  });
});
