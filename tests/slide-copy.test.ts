import { describe, expect, it } from "vitest";
import { blankSlideXml, prepareSlideCopy, relationshipsXml } from "../src/editor/slides";

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
