import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { emptyStructure, exportEditedPptx, materializeSlide, resolveTarget } from "../src/pptx-exporter";
import { Slide } from "../src/pptx-virtual-dom";
import { checkPackage } from "./helpers/package-check";

const P = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const slideXml = (text: string) =>
  `<p:sld ${P}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
const slideRels = (n: number) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
  + `<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`
  + `<Relationship Id="rId2" Type="${REL}/notesSlide" Target="../notesSlides/notesSlide${n}.xml"/></Relationships>`;

const presentation = (sections = false) =>
  `<p:presentation ${P} xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:sldMasterIdLst/>`
  + `<p:sldIdLst><p:sldId id="300" r:id="rId4"/><p:sldId id="301" r:id="rId5"/><p:sldId id="302" r:id="rId6"/></p:sldIdLst>`
  + (sections
    ? `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst>`
      + `<p14:section name="A" id="{A}"><p14:sldIdLst><p14:sldId id="300"/><p14:sldId id="301"/></p14:sldIdLst></p14:section>`
      + `<p14:section name="B" id="{B}"><p14:sldIdLst><p14:sldId id="302"/></p14:sldIdLst></p14:section>`
      + `</p14:sectionLst></p:ext></p:extLst>`
    : "")
  + `</p:presentation>`;

async function makeDeck(sections = false): Promise<JSZip> {
  const zip = new JSZip();
  zip.file("ppt/presentation.xml", presentation(sections));
  zip.file("ppt/_rels/presentation.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + [1, 2, 3].map(n => `<Relationship Id="rId${n + 3}" Type="${REL}/slide" Target="slides/slide${n}.xml"/>`).join("")
    + `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
    + [1, 2, 3].map(n => `<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("")
    + [1, 2, 3].map(n => `<Override PartName="/ppt/notesSlides/notesSlide${n}.xml" ContentType="notes"/>`).join("") + `</Types>`);
  zip.file("ppt/slideMasters/slideMaster1.xml", "<master/>");
  zip.file("ppt/slideLayouts/slideLayout1.xml", "<layout/>");
  for (const n of [1, 2, 3]) {
    zip.file(`ppt/slides/slide${n}.xml`, slideXml(`slide ${n}`));
    zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, slideRels(n));
    zip.file(`ppt/notesSlides/notesSlide${n}.xml`, `<notes/>`);
    zip.file(`ppt/notesSlides/_rels/notesSlide${n}.xml.rels`, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slide" Target="../slides/slide${n}.xml"/></Relationships>`);
  }
  return zip;
}

const toBuffer = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error);
  reader.readAsArrayBuffer(blob);
});
const deckSlide = (n: number): Slide => ({ id: `slide_${n}`, elements: [], rawXml: "", part: `ppt/slides/slide${n}.xml` });
const open = async (blob: Blob) => JSZip.loadAsync(await toBuffer(blob));
const listIds = (xml: string) => Array.from(xml.matchAll(/<p:sldId id="(\d+)"/g)).map(m => m[1]);
const run = (zip: JSZip, slides: Slide[], structure = emptyStructure()) =>
  exportEditedPptx(zip, slides, new Map(), new Map(), new Map(), structure);

describe("resolveTarget", () => {
  it("resolves relative and absolute targets", () => {
    expect(resolveTarget("ppt/slides", "../media/a.png")).toBe("ppt/media/a.png");
    expect(resolveTarget("ppt", "slides/slide1.xml")).toBe("ppt/slides/slide1.xml");
    expect(resolveTarget("ppt/slides", "/ppt/theme/theme1.xml")).toBe("ppt/theme/theme1.xml");
  });
});

describe("slide list sync", () => {
  it("leaves the package alone when the slide list is unchanged", async () => {
    const zip = await makeDeck();
    const out = await open(await run(zip, [1, 2, 3].map(deckSlide)));
    expect(await out.file("ppt/presentation.xml")!.async("string")).toBe(presentation());
  });

  it("reorders slides by rewriting sldIdLst only", async () => {
    const zip = await makeDeck();
    const out = await open(await run(zip, [3, 1, 2].map(deckSlide)));
    expect(listIds(await out.file("ppt/presentation.xml")!.async("string"))).toEqual(["302", "300", "301"]);
    expect(out.file("ppt/slides/slide2.xml")).not.toBeNull();
  });

  it("deletes a slide with its rels, notes page, content types and list entry, and restores them in memory", async () => {
    const zip = await makeDeck();
    const out = await open(await run(zip, [1, 3].map(deckSlide)));
    expect(listIds(await out.file("ppt/presentation.xml")!.async("string"))).toEqual(["300", "302"]);
    expect(out.file("ppt/slides/slide2.xml")).toBeNull();
    expect(out.file("ppt/slides/_rels/slide2.xml.rels")).toBeNull();
    expect(out.file("ppt/notesSlides/notesSlide2.xml")).toBeNull();
    expect(await out.file("ppt/_rels/presentation.xml.rels")!.async("string")).not.toContain("slide2.xml");
    const types = await out.file("[Content_Types].xml")!.async("string");
    expect(types).not.toContain("slide2.xml");
    expect(types).not.toContain("notesSlide2.xml");
    expect(types).toContain("slide3.xml");
    // the live package still has everything, so the delete can be undone
    expect(zip.file("ppt/slides/slide2.xml")).not.toBeNull();
    expect(zip.file("ppt/notesSlides/notesSlide2.xml")).not.toBeNull();
    expect(await zip.file("ppt/presentation.xml")!.async("string")).toBe(presentation());
  });

  it("adds a new slide made only of virtual files, and removes them again after export", async () => {
    const zip = await makeDeck();
    const structure = emptyStructure();
    structure.virtual.set("ppt/slides/slide4.xml", slideXml("copy of 2"));
    structure.virtual.set("ppt/slides/_rels/slide4.xml.rels", slideRels(9).replace(/<Relationship Id="rId2"[^>]*>/, ""));
    const slides = [1, 2, 3].map(deckSlide);
    slides.splice(2, 0, deckSlide(4));
    const out = await open(await run(zip, slides, structure));
    const xml = await out.file("ppt/presentation.xml")!.async("string");
    expect(listIds(xml)).toEqual(["300", "301", "303", "302"]);
    expect(await out.file("ppt/slides/slide4.xml")!.async("string")).toContain("copy of 2");
    const rels = await out.file("ppt/_rels/presentation.xml.rels")!.async("string");
    expect(rels).toContain('Target="slides/slide4.xml"');
    expect(await out.file("[Content_Types].xml")!.async("string")).toContain("/ppt/slides/slide4.xml");
    expect(zip.file("ppt/slides/slide4.xml")).toBeNull();
    expect(await zip.file("[Content_Types].xml")!.async("string")).not.toContain("slide4");
  });

  it("keeps sections valid: a copy joins its neighbour's section, a deleted last slide empties its section", async () => {
    const zip = await makeDeck(true);
    const structure = emptyStructure();
    structure.virtual.set("ppt/slides/slide4.xml", slideXml("copy"));
    structure.virtual.set("ppt/slides/_rels/slide4.xml.rels", slideRels(9));
    const slides = [deckSlide(1), deckSlide(4), deckSlide(2)];
    const out = await open(await run(zip, slides, structure));
    const xml = await out.file("ppt/presentation.xml")!.async("string");
    const sections = Array.from(xml.matchAll(/<p14:section name="(\w)"[^>]*>(.*?)<\/p14:section>/g)).map(m => [m[1], listIds(m[2].replace(/p14:/g, "p:"))]);
    expect(sections).toEqual([["A", ["300", "303", "301"]], ["B", []]]);
  });
});

describe("materializeSlide", () => {
  it("returns the slide with its pending edits and all of its relationships", async () => {
    const zip = await makeDeck();
    const edits = { dirty: new Map(), textEdits: new Map(), styleEdits: new Map(), structure: emptyStructure() };
    const copy = await materializeSlide(zip, "ppt/slides/slide1.xml", [], edits);
    expect(copy!.xml).toContain("slide 1");
    expect(copy!.relationships.map(rel => rel.id).sort()).toEqual(["rId1", "rId2"]);
  });
});

describe("package consistency after slide edits", () => {
  it("leaves no dangling relationships, undeclared parts or orphan notes after reorder, delete and add", async () => {
    const structure = emptyStructure();
    structure.virtual.set("ppt/slides/slide4.xml", slideXml("copy"));
    structure.virtual.set("ppt/slides/_rels/slide4.xml.rels", slideRels(9).replace(/<Relationship Id="rId2"[^>]*>/, ""));
    const cases: Array<[string, Slide[], ReturnType<typeof emptyStructure>?]> = [
      ["untouched", [1, 2, 3].map(deckSlide)],
      ["reordered", [3, 1, 2].map(deckSlide)],
      ["deleted", [1, 3].map(deckSlide)],
      ["added", [deckSlide(1), deckSlide(4), deckSlide(2), deckSlide(3)], structure]
    ];
    for (const [name, slides, edits] of cases) {
      const zip = await makeDeck();
      const blob = await run(zip, slides, edits);
      expect({ name, problems: await checkPackage(await toBuffer(blob)) }).toEqual({ name, problems: [] });
    }
  });
});

describe("links to deleted slides", () => {
  it("drops the relationship and the hyperlink that used it, and keeps links to slides that remain", async () => {
    const zip = await makeDeck();
    const link = (id: string) => `<a:r><a:rPr><a:hlinkClick xmlns:r="${REL}" r:id="${id}" action="ppaction://hlinksldjump"/></a:rPr><a:t>go</a:t></a:r>`;
    zip.file("ppt/slides/slide1.xml", slideXml("one").replace("<a:r>", link("rId7") + link("rId8") + "<a:r>"));
    zip.file("ppt/slides/_rels/slide1.xml.rels", slideRels(1).replace("</Relationships>",
      `<Relationship Id="rId7" Type="${REL}/slide" Target="slide2.xml"/><Relationship Id="rId8" Type="${REL}/slide" Target="slide3.xml"/></Relationships>`));
    const out = await open(await run(zip, [1, 3].map(deckSlide)));
    const xml = await out.file("ppt/slides/slide1.xml")!.async("string");
    const rels = await out.file("ppt/slides/_rels/slide1.xml.rels")!.async("string");
    expect(xml).not.toContain('r:id="rId7"');
    expect(xml).toContain('r:id="rId8"');
    expect(rels).not.toContain("slide2.xml");
    expect(rels).toContain("slide3.xml");
    expect(await checkPackage(await toBuffer(await run(zip, [1, 3].map(deckSlide))))).toEqual([]);
    // the live package keeps the link, so undoing the delete works
    expect(await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string")).toContain("slide2.xml");
  });
});
