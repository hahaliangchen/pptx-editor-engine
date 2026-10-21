import { describe, expect, it } from "vitest";
import { NewShapeKind, buildPictureXml, buildShapeElements, buildShapeXml } from "../src/editor/new-shapes";
import { writeTextOps } from "../src/pptx-text-writer";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const box = { x: 100, y: 200, cx: 3000, cy: 1000 };
const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");

describe("buildShapeXml", () => {
  const kinds: NewShapeKind[] = ["textbox", "rect", "roundRect", "ellipse", "triangle", "rightArrow", "leftArrow", "upArrow", "downArrow", "line", "arrow"];

  it("produces well-formed, correctly namespaced nodes for every kind", () => {
    for (const kind of kinds) {
      const doc = parse(buildShapeXml(kind, "7", box));
      expect(doc.getElementsByTagName("parsererror")).toHaveLength(0);
      const root = doc.documentElement;
      expect(root.namespaceURI).toBe(P);
      expect(root.localName).toBe(kind === "line" || kind === "arrow" ? "cxnSp" : "sp");
      expect(doc.getElementsByTagNameNS(P, "cNvPr")[0].getAttribute("id")).toBe("7");
      expect(doc.getElementsByTagNameNS(A, "off")[0].getAttribute("x")).toBe("100");
      expect(Array.from(root.getElementsByTagName("*")).filter(n => n.tagName.startsWith("a:") && n.namespaceURI !== A)).toEqual([]);
    }
  });

  it("gives shapes a fill, border and centred white text, and text boxes neither fill nor border", () => {
    const shape = parse(buildShapeXml("rect", "2", box));
    expect(shape.getElementsByTagNameNS(A, "srgbClr")[0].getAttribute("val")).toBe("4472C4");
    expect(shape.getElementsByTagNameNS(A, "ln")).toHaveLength(1);
    expect(shape.getElementsByTagNameNS(A, "bodyPr")[0].getAttribute("anchor")).toBe("ctr");
    const text = parse(buildShapeXml("textbox", "3", box));
    expect(text.getElementsByTagNameNS(A, "noFill")).toHaveLength(1);
    expect(text.getElementsByTagNameNS(A, "ln")).toHaveLength(0);
    expect(text.getElementsByTagNameNS(P, "cNvSpPr")[0].getAttribute("txBox")).toBe("1");
  });

  it("flips lines drawn up or left and adds an arrow head for arrows", () => {
    const line = parse(buildShapeXml("arrow", "4", { ...box, flipV: true }));
    expect(line.getElementsByTagNameNS(A, "xfrm")[0].getAttribute("flipV")).toBe("1");
    expect(line.getElementsByTagNameNS(A, "tailEnd")[0].getAttribute("type")).toBe("triangle");
    expect(parse(buildShapeXml("line", "5", box)).getElementsByTagNameNS(A, "tailEnd")).toHaveLength(0);
  });

  it("builds a picture node that points at its relationship", () => {
    const doc = parse(buildPictureXml("9", "图片 9", "rId12", box));
    expect(doc.getElementsByTagNameNS(A, "blip")[0].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed")).toBe("rId12");
  });

  it("lets typing into the new shape write a run that keeps its size and colour", () => {
    const doc = parse(buildShapeXml("rect", "2", box));
    writeTextOps(doc.documentElement, [{ start: 0, end: 0, insert: "Hi" }]);
    const rPr = doc.getElementsByTagNameNS(A, "rPr")[0];
    expect(rPr.getAttribute("sz")).toBe("1800");
    expect(rPr.getElementsByTagNameNS(A, "srgbClr")[0].getAttribute("val")).toBe("FFFFFF");
    expect(doc.getElementsByTagNameNS(A, "t")[0].textContent).toBe("Hi");
  });
});

describe("buildShapeElements", () => {
  const source = { part: "ppt/slides/slide1.xml", shapeId: "7", xform: { scaleX: 1, scaleY: 1, offsetX: 0, offsetY: 0 } };
  const rect = { x: 10, y: 20, w: 300, h: 100 };

  it("returns a shape and an empty centred text element that share one source", () => {
    const [shape, text] = buildShapeElements("roundRect", source, rect, { emuPerPx: 6350 });
    expect(shape.type).toBe("shape");
    expect(text.type).toBe("text");
    expect(shape.source).toBe(source);
    expect(text.source).toBe(source);
    if (text.type === "text") {
      expect(text.paragraphs![0].runs[0].content).toBe("");
      expect(text.style.align).toBe("center");
      expect(text.style.fontSize).toBeCloseTo(18 * 12700 / 6350);
    }
    if (shape.type === "shape") expect(shape.cornerRadius).toBeCloseTo(0.16667);
  });

  it("returns a single arrow-headed line element for arrows", () => {
    const elements = buildShapeElements("arrow", source, rect, { emuPerPx: 6350 }, { flipV: true });
    expect(elements).toHaveLength(1);
    const line = elements[0];
    if (line.type === "shape") {
      expect(line.flipV).toBe(true);
      expect(line.computedStyle!.line!.tailEnd).toEqual({ type: "triangle" });
    }
  });
});
