import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  emptyStructure,
  exportEditedPptx,
  markDirty,
  materializeShape,
  popTextOp,
  pushItem,
  pushTextOp
} from "../src/pptx-exporter";
import { ElementSource, ShapeElement, Slide, TextElement } from "../src/pptx-virtual-dom";

const PART = "ppt/slides/slide1.xml";
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS}><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="A"/></p:nvSpPr><p:spPr><a:xfrm rot="600000"><a:off x="100" y="200"/><a:ext cx="3000" cy="4000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="B"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
<p:pic><p:nvPicPr><p:cNvPr id="4" name="C"/></p:nvPicPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:pic>
<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="G"/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="6" name="D"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="50" y="50"/><a:ext cx="500" cy="500"/></a:xfrm></p:spPr></p:sp></p:grpSp>
</p:spTree></p:cSld></p:sld>`;

const xform = { scaleX: 0.001, scaleY: 0.001, offsetX: 0, offsetY: 0 };
const source = (shapeId: string, xf = xform): ElementSource => ({ part: PART, shapeId, xform: xf });
const shape = (shapeId: string, rect: ShapeElement["rect"], xf = xform): ShapeElement => ({
  type: "shape", id: `shape_${shapeId}`, source: source(shapeId, xf), rect, shapeType: "rect", fill: "#fff"
});

async function makeZip(): Promise<JSZip> {
  const zip = new JSZip();
  zip.file(PART, SLIDE_XML);
  zip.file("ppt/slides/slide2.xml", "<untouched/>");
  zip.file("ppt/media/image1.png", new Uint8Array([1, 2, 3, 4]));
  return zip;
}
// jsdom's Blob has no arrayBuffer().
const toBuffer = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error);
  reader.readAsArrayBuffer(blob);
});
const read = async (blob: Blob, name: string) => (await JSZip.loadAsync(await toBuffer(blob))).file(name)!.async("string");
const slide = (elements: Slide["elements"]): Slide => ({ id: "slide_1", elements, rawXml: "" });

describe("exportEditedPptx", () => {
  it("rewrites only the edited slide and keeps every other part byte-identical", async () => {
    const zip = await makeZip();
    const dirty = new Map();
    markDirty(dirty, source("2"));
    const blob = await exportEditedPptx(zip, [slide([shape("2", { x: 1, y: 2, w: 3, h: 4 })])], dirty);
    const out = await JSZip.loadAsync(await toBuffer(blob));
    expect(await out.file("ppt/slides/slide2.xml")!.async("string")).toBe("<untouched/>");
    expect(Array.from(await out.file("ppt/media/image1.png")!.async("uint8array"))).toEqual([1, 2, 3, 4]);
  });

  it("converts the logical rect back to EMU and keeps rotation", async () => {
    const zip = await makeZip();
    const dirty = new Map();
    markDirty(dirty, source("2"));
    const blob = await exportEditedPptx(zip, [slide([shape("2", { x: 1, y: 2, w: 3, h: 4 })])], dirty);
    const xml = await read(blob, PART);
    expect(xml).toContain('<a:off x="1000" y="2000"/>');
    expect(xml).toContain('<a:ext cx="3000" cy="4000"/>');
    expect(xml).toContain('rot="600000"');
  });

  it("undoes the group transform for children of a group", async () => {
    const zip = await makeZip();
    const dirty = new Map();
    const groupXform = { scaleX: 0.002, scaleY: 0.002, offsetX: 10, offsetY: 20 };
    markDirty(dirty, source("6", groupXform));
    const blob = await exportEditedPptx(zip, [slide([shape("6", { x: 110, y: 120, w: 40, h: 60 }, groupXform)])], dirty);
    const xml = await read(blob, PART);
    expect(xml).toContain('<a:off x="50000" y="50000"/>');
    expect(xml).toContain('<a:ext cx="20000" cy="30000"/>');
  });

  it("adds an xfrm to a placeholder that inherited its position", async () => {
    const zip = await makeZip();
    const dirty = new Map();
    markDirty(dirty, source("3"));
    const blob = await exportEditedPptx(zip, [slide([shape("3", { x: 5, y: 6, w: 7, h: 8 })])], dirty);
    const xml = await read(blob, PART);
    expect(xml).toMatch(/<p:cNvPr id="3"[\s\S]*?<a:xfrm><a:off x="5000" y="6000"\/><a:ext cx="7000" cy="8000"\/><\/a:xfrm>/);
  });

  it("removes a deleted shape, and a picture is found by its own id", async () => {
    const zip = await makeZip();
    const dirty = new Map();
    markDirty(dirty, source("2"));
    markDirty(dirty, source("4"));
    const blob = await exportEditedPptx(zip, [slide([{
      type: "image", id: "img_4", source: source("4"), rect: { x: 1, y: 1, w: 2, h: 2 }, url: ""
    }])], dirty);
    const xml = await read(blob, PART);
    expect(xml).not.toContain('name="A"');
    expect(xml).toContain('<a:off x="1000" y="1000"/>');
  });

  it("replays text edits and a second export after undo starts from the original", async () => {
    const zip = await makeZip();
    const edits = new Map();
    const target: TextElement = {
      type: "text", id: "t", source: source("2"), rect: { x: 0, y: 0, w: 1, h: 1 }, content: "Hello",
      style: { fontSize: 10, color: "#000", bold: false, align: "left" }
    };
    pushTextOp(edits, target.source!, { start: 5, end: 5, insert: "!" });
    const first = await exportEditedPptx(zip, [slide([target])], new Map(), edits);
    expect(await read(first, PART)).toContain("<a:t>Hello!</a:t>");

    popTextOp(edits, target.source!);
    expect(edits.size).toBe(0);
    const second = await exportEditedPptx(zip, [slide([target])], new Map(), edits);
    const out = await JSZip.loadAsync(await toBuffer(second));
    expect(await out.file(PART)!.async("string")).toBe(SLIDE_XML);
  });

  it("moves and resizes a table: frame box from the union of its cells, grid scaled to match", async () => {
    const tableXml = `<p:sld ${NS}><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="T"/></p:nvGraphicFramePr>`
      + `<p:xfrm><a:off x="0" y="0"/><a:ext cx="4000" cy="2000"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblGrid>`
      + `<a:gridCol w="1000"/><a:gridCol w="3000"/></a:tblGrid><a:tr h="500"><a:tc/><a:tc/></a:tr><a:tr h="1500"><a:tc/><a:tc/></a:tr>`
      + `</a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`;
    const zip = new JSZip();
    zip.file(PART, tableXml);
    const cell = (x: number, y: number, w: number, h: number): ShapeElement => shape("9", { x, y, w, h });
    // Frame moved to (10,20) and doubled in width, same height.
    const cells = [cell(10, 20, 2, 0.5), cell(12, 20, 6, 0.5), cell(10, 20.5, 2, 1.5), cell(12, 20.5, 6, 1.5)];
    const dirty = new Map();
    markDirty(dirty, source("9"));
    const blob = await exportEditedPptx(zip, [slide(cells)], dirty);
    const xml = await read(blob, PART);
    expect(xml).toContain('<a:off x="10000" y="20000"/>');
    expect(xml).toContain('<a:ext cx="8000" cy="2000"/>');
    expect(xml).toContain('<a:gridCol w="2000"/>');
    expect(xml).toContain('<a:gridCol w="6000"/>');
    expect(xml).toContain('<a:tr h="500">');
  });

  it("moves a group by its own off/ext and leaves the children's coordinates alone", async () => {
    const xml = `<p:sld ${NS}><p:cSld><p:spTree><p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="G"/></p:nvGrpSpPr>`
      + `<p:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="4000" cy="2000"/><a:chOff x="0" y="0"/><a:chExt cx="400" cy="200"/></a:xfrm></p:grpSpPr>`
      + `<p:sp><p:nvSpPr><p:cNvPr id="6" name="D"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="200" cy="200"/></a:xfrm></p:spPr></p:sp>`
      + `</p:grpSp></p:spTree></p:cSld></p:sld>`;
    const zip = new JSZip();
    zip.file(PART, xml);
    const ref = { shapeId: "5", xform, origRect: { x: 1, y: 2, w: 4, h: 2 } };
    const child = shape("6", { x: 3, y: 5, w: 4, h: 2 }); // moved by (+2,+3), width unchanged
    child.source = { ...source("6", { scaleX: 0.01, scaleY: 0.01, offsetX: 1, offsetY: 2 }), groups: [ref] };
    const dirty = new Map();
    markDirty(dirty, { part: PART, shapeId: "5", xform });
    const blob = await exportEditedPptx(zip, [slide([child])], dirty);
    const out = await read(blob, PART);
    expect(out).toContain('<a:off x="3000" y="5000"/>');
    expect(out).toContain('<a:ext cx="4000" cy="2000"/>');
    expect(out).toContain('<a:chOff x="0" y="0"/><a:chExt cx="400" cy="200"/>');
    expect(out).toContain('<a:off x="0" y="0"/><a:ext cx="200" cy="200"/>'); // child untouched
  });

  it("writes the rotation back as 60000ths of a degree, and clears it at zero", async () => {
    const zip = await makeZip();
    const dirty = new Map();
    markDirty(dirty, source("2"));
    const turned = { ...shape("2", { x: 1, y: 2, w: 3, h: 4 }), rotation: 90 };
    expect(await read(await exportEditedPptx(zip, [slide([turned])], dirty), PART)).toContain('rot="5400000"');
    const straight = { ...turned, rotation: 0 };
    expect(await read(await exportEditedPptx(zip, [slide([straight])], dirty), PART)).not.toContain('rot="');
  });

  const orderOf = (xml: string) => Array.from(xml.matchAll(/<p:cNvPr id="(\d+)"/g)).map(m => m[1]);

  it("replays z-order changes among sibling shapes", async () => {
    const zip = await makeZip();
    const structure = emptyStructure();
    pushItem(structure.orders, PART, { shapeId: "2", op: "front" });
    let xml = await read(await exportEditedPptx(zip, [], new Map(), new Map(), new Map(), structure), PART);
    expect(orderOf(xml).filter(id => ["2", "3", "4", "5"].includes(id))).toEqual(["3", "4", "5", "2"]);

    const back = emptyStructure();
    pushItem(back.orders, PART, { shapeId: "4", op: "back" });
    xml = await read(await exportEditedPptx(zip, [], new Map(), new Map(), new Map(), back), PART);
    expect(orderOf(xml).filter(id => ["2", "3", "4", "5"].includes(id))).toEqual(["4", "2", "3", "5"]);

    const steps = emptyStructure();
    pushItem(steps.orders, PART, { shapeId: "3", op: "forward" });
    pushItem(steps.orders, PART, { shapeId: "5", op: "backward" });
    xml = await read(await exportEditedPptx(zip, [], new Map(), new Map(), new Map(), steps), PART);
    expect(orderOf(xml).filter(id => ["2", "3", "4", "5"].includes(id))).toEqual(["2", "4", "5", "3"]);
  });

  it("inserts a pasted shape on top, positions it from its elements, and adds its relationships", async () => {
    const zip = await makeZip();
    zip.file("ppt/slides/_rels/slide1.xml.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="t/layout" Target="../slideLayouts/l.xml"/></Relationships>`);
    const structure = emptyStructure();
    pushItem(structure.inserts, PART, {
      id: "9",
      xml: `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="9" name="copy"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="1"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr></p:sp>`
    });
    pushItem(structure.relAdds, PART, { id: "rId2", type: "t/image", target: "../media/image1.png" });
    const dirty = new Map();
    markDirty(dirty, source("9"));
    const blob = await exportEditedPptx(zip, [slide([shape("9", { x: 5, y: 6, w: 7, h: 8 })])], dirty, new Map(), new Map(), structure);
    const xml = await read(blob, PART);
    expect(orderOf(xml).pop()).toBe("9");
    expect(xml).toContain('<a:off x="5000" y="6000"/><a:ext cx="7000" cy="8000"/>');
    const rels = await read(blob, "ppt/slides/_rels/slide1.xml.rels");
    expect(rels).toContain('Id="rId2"');
    expect(rels).toContain('Target="../media/image1.png"');
    expect(rels).toContain('Id="rId1"');
    // the original package is left as it was
    expect(await zip.file("ppt/slides/_rels/slide1.xml.rels")!.async("string")).not.toContain("rId2");
  });

  it("materializes a shape with its pending edits and reports its relationships", async () => {
    const zip = new JSZip();
    zip.file(PART, `<p:sld ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>`
      + `<p:pic><p:nvPicPr><p:cNvPr id="4" name="C"/><p:cNvPicPr/><p:nvPr><p:custDataLst><p:tags r:id="rId9"/></p:custDataLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId5"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:pic>`
      + `<p:sp><p:nvSpPr><p:cNvPr id="2" name="A"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>`
      + `</p:spTree></p:cSld></p:sld>`);
    zip.file("ppt/slides/_rels/slide1.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="t/image" Target="../media/image7.png"/><Relationship Id="rId9" Type="t/tags" Target="../tags/tag1.xml"/></Relationships>`);
    const edits = { dirty: new Map(), textEdits: new Map(), styleEdits: new Map(), structure: emptyStructure() };
    pushTextOpTo(edits, "2");
    const text = await materializeShape(zip, PART, "2", [], edits);
    expect(text!.xml).toContain("<a:t>Hello!</a:t>");
    const picture = await materializeShape(zip, PART, "4", [], edits);
    expect(picture!.relationships).toEqual([{ id: "rId5", type: "t/image", target: "../media/image7.png", targetMode: undefined }]);
    expect(picture!.xml).not.toContain("custDataLst");
  });
});

describe("pictures added to the package", () => {
  it("writes the file, declares its type once, and removes both again afterwards", async () => {
    const zip = await makeZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`);
    const structure = emptyStructure();
    structure.media.push({ path: "ppt/media/image9.png", bytes: new Uint8Array([9, 9, 9]), extension: "png", contentType: "image/png" });
    structure.media.push({ path: "ppt/media/image10.png", bytes: new Uint8Array([8]), extension: "png", contentType: "image/png" });
    const blob = await exportEditedPptx(zip, [], new Map(), new Map(), new Map(), structure);
    const out = await JSZip.loadAsync(await toBuffer(blob));
    expect(Array.from(await out.file("ppt/media/image9.png")!.async("uint8array"))).toEqual([9, 9, 9]);
    expect(Array.from(await out.file("ppt/media/image10.png")!.async("uint8array"))).toEqual([8]);
    const types = await out.file("[Content_Types].xml")!.async("string");
    expect(types.match(/Extension="png"/g)).toHaveLength(1);
    expect(types).toContain('ContentType="image/png"');
    // the live package is untouched
    expect(zip.file("ppt/media/image9.png")).toBeNull();
    expect(await zip.file("[Content_Types].xml")!.async("string")).not.toContain("png");
  });
});

function pushTextOpTo(edits: { textEdits: Map<string, Map<string, unknown[]>> }, shapeId: string) {
  pushTextOp(edits.textEdits as never, { part: PART, shapeId, xform }, { start: 5, end: 5, insert: "!" });
}
