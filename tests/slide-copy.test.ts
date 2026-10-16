import { describe, expect, it } from "vitest";
import { blankSlideXml, layoutPlaceholders, placeholderSlideXml, prepareSlideCopy, relationshipsXml } from "../src/editor/slides";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const rel = (id: string, kind: string, target = `../x/${id}.xml`) => ({ id, type: `${REL}/${kind}`, target });

describe("slide helpers", () => {
  it("builds a well-formed blank slide and relationship file", () => {
    const doc = new DOMParser().parseFromString(blankSlideXml(), "application/xml");
    expect(doc.getElementsByTagName("parsererror")).toHaveLength(0);
    expect(doc.documentElement.localName).toBe("sld");
    const rels = new DOMParser().parseFromString(
      relationshipsXml([{ ...rel("rId1", "slideLayout"), targetMode: undefined }, { id: "rId2", type: `${REL}/hyperlink`, target: "http://a?b=1&c=2", targetMode: "External" }]),
      "application/xml"
    );
    expect(rels.getElementsByTagName("parsererror")).toHaveLength(0);
    expect(rels.getElementsByTagName("Relationship")[1].getAttribute("Target")).toBe("http://a?b=1&c=2");
    expect(rels.getElementsByTagName("Relationship")[1].getAttribute("TargetMode")).toBe("External");
  });

  it("drops notes and tags from a copy, strips tag references, and shares pictures and layout", () => {
    const copy = prepareSlideCopy({
      xml: '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:custDataLst><p:tags r:id="rId9"/></p:custDataLst></p:cSld></p:sld>',
      relationships: [rel("rId1", "slideLayout"), rel("rId2", "notesSlide"), rel("rId3", "tags"), rel("rId4", "image")]
    });
    expect(copy.relationships.map(r => r.id)).toEqual(["rId1", "rId4"]);
    expect(copy.xml).not.toContain("custDataLst");
  });

  it("refuses slides that own charts or other parts", () => {
    expect(() => prepareSlideCopy({ xml: "<p:sld xmlns:p='x'/>", relationships: [rel("rId1", "chart")] })).toThrow(/暂不支持/);
  });
});

describe("placeholders from a layout", () => {
  const layout = '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr></p:sp>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="b"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr></p:sp>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="4" name="s"/><p:cNvSpPr/><p:nvPr><p:ph type="subTitle" idx="13"/></p:nvPr></p:nvSpPr></p:sp>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="5" name="d"/><p:cNvSpPr/><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr></p:sp>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="6" name="n"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr></p:sp>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="7" name="p"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="14"/></p:nvPr></p:nvSpPr></p:sp>'
    + '</p:spTree></p:cSld></p:sldLayout>';

  it("keeps the text placeholders and skips footer, date, number and picture ones", () => {
    expect(layoutPlaceholders(layout)).toEqual([
      { type: "title" },
      { idx: "1" },
      { type: "subTitle", idx: "13" }
    ].map(spec => ({ type: undefined, idx: undefined, orient: undefined, sz: undefined, ...spec })));
  });

  it("builds a slide with one empty, typed placeholder shape per entry and unique ids", () => {
    const doc = new DOMParser().parseFromString(placeholderSlideXml(layoutPlaceholders(layout)), "application/xml");
    expect(doc.getElementsByTagName("parsererror")).toHaveLength(0);
    const shapes = Array.from(doc.getElementsByTagName("p:sp"));
    expect(shapes).toHaveLength(3);
    expect(Array.from(doc.getElementsByTagName("p:cNvPr")).map(n => n.getAttribute("id"))).toEqual(["1", "2", "3", "4"]);
    expect(Array.from(doc.getElementsByTagName("p:ph")).map(n => [n.getAttribute("type"), n.getAttribute("idx")]))
      .toEqual([["title", null], [null, "1"], ["subTitle", "13"]]);
    expect(doc.getElementsByTagName("a:t")).toHaveLength(0);
    expect(shapes[0].getElementsByTagName("p:spPr")[0].children).toHaveLength(0);
  });
});

