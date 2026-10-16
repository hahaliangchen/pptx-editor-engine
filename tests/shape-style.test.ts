import { describe, expect, it } from "vitest";
import { applyShapeStyleOp, restoreShape, snapshotShape } from "../src/editor/shape-ops";
import { describeSelection } from "../src/editor/properties";
import { applyTextOp } from "../src/editor/text-ops";
import { vdomTextModel } from "../src/editor/text-model-vdom";
import { writeShapeStyleOps } from "../src/pptx-style-writer";
import { writeTextOps } from "../src/pptx-text-writer";
import { ShapeElement, TextParagraph } from "../src/pptx-virtual-dom";

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const parse = (inner: string) => new DOMParser().parseFromString(`<p:sp ${NS}>${inner}</p:sp>`, "application/xml").documentElement;
const names = (el: Element) => Array.from(el.children).map(child => child.localName);

const shape = (extra: Partial<ShapeElement> = {}): ShapeElement => ({
  type: "shape", id: "s", rect: { x: 0, y: 0, w: 10, h: 10 }, shapeType: "rect", fill: "#ffffff",
  computedStyle: { fill: { type: "solid", color: "#ffffff" } }, ...extra
});

describe("shape style on the virtual DOM", () => {
  it("sets and clears the fill, keeping the fallback colour in step", () => {
    const s = shape();
    applyShapeStyleOp(s, { kind: "fill", color: "#e60012" }, 6350);
    expect(s.fill).toBe("#e60012");
    expect(s.computedStyle!.fill).toEqual({ type: "solid", color: "#e60012" });
    applyShapeStyleOp(s, { kind: "fill", color: null }, 6350);
    expect(s.fill).toBe("transparent");
    expect(s.computedStyle!.fill).toEqual({ type: "none" });
  });

  it("adds a border from a width op and converts points to pixels", () => {
    const s = shape();
    applyShapeStyleOp(s, { kind: "line", color: "#000000", widthPt: 2 }, 6350);
    expect(s.computedStyle!.line!.width).toBeCloseTo(4);
    expect(s.border).toEqual({ color: "#000000", width: s.computedStyle!.line!.width });
    applyShapeStyleOp(s, { kind: "line", dash: "dash" }, 6350);
    expect(s.computedStyle!.line!.dash).toBe("dash");
    expect(s.computedStyle!.line!.width).toBeCloseTo(4);
    applyShapeStyleOp(s, { kind: "line", color: null }, 6350);
    expect(s.computedStyle!.line).toBeUndefined();
    expect(s.border).toBeUndefined();
  });

  it("restores an exact snapshot", () => {
    const s = shape({ border: { color: "#111111", width: 2 } });
    const before = snapshotShape(s);
    applyShapeStyleOp(s, { kind: "fill", color: "#00ff00" }, 6350);
    restoreShape(s, before);
    expect(s.fill).toBe("#ffffff");
    expect(s.computedStyle!.fill).toEqual({ type: "solid", color: "#ffffff" });
  });
});

describe("describeSelection bounds", () => {
  it("reports the selection's box in centimetres and its rotation", () => {
    // 6350 EMU per px: 360000 / 6350 = 56.69 px per cm.
    const props = describeSelection([], [], 6350, { rect: { x: 56.69, y: 113.39, w: 283.46, h: 170.08 }, rotation: 30 });
    expect(props.bounds).toEqual({ x: 1, y: 2, w: 5, h: 3 });
    expect(props.rotation).toBe(30);
    expect(describeSelection([], [], 6350).bounds).toBeNull();
  });
});

describe("describeSelection", () => {
  it("reports shared values and mixed ones", () => {
    const a = shape();
    const b = shape({ fill: "#ff0000", computedStyle: { fill: { type: "solid", color: "#ff0000" } } });
    expect(describeSelection([a, a], [], 6350).fill).toEqual({ kind: "solid", color: "#ffffff" });
    expect(describeSelection([a, b], [], 6350).fill).toEqual({ kind: "mixed" });
    expect(describeSelection([a], [], 6350).lineColor).toEqual({ kind: "none" });
  });
});

describe("shape style in OOXML", () => {
  it("puts the fill after the geometry and before a:ln", () => {
    const sp = parse('<p:nvSpPr><p:cNvPr id="2" name="x"/></p:nvSpPr><p:spPr><a:xfrm/><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:effectLst/></p:spPr>');
    writeShapeStyleOps(sp, [{ kind: "fill", color: "#e60012" }]);
    const spPr = sp.getElementsByTagName("p:spPr")[0];
    expect(names(spPr)).toEqual(["xfrm", "prstGeom", "solidFill", "ln", "effectLst"]);
    expect(spPr.getElementsByTagName("a:srgbClr")[0].getAttribute("val")).toBe("E60012");
  });

  it("writes noFill, and creates spPr content when the shape had none", () => {
    const sp = parse('<p:nvSpPr><p:cNvPr id="2" name="x"/></p:nvSpPr><p:spPr/>');
    writeShapeStyleOps(sp, [{ kind: "fill", color: null }]);
    expect(names(sp.getElementsByTagName("p:spPr")[0])).toEqual(["noFill"]);
  });

  it("builds a:ln with width, colour first and dash after it", () => {
    const sp = parse('<p:nvSpPr><p:cNvPr id="2" name="x"/></p:nvSpPr><p:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:spPr>');
    writeShapeStyleOps(sp, [
      { kind: "line", color: "#0070c0", widthPt: 2.25 },
      { kind: "line", dash: "dash" }
    ]);
    const spPr = sp.getElementsByTagName("p:spPr")[0];
    expect(names(spPr)).toEqual(["prstGeom", "solidFill", "ln"]);
    const ln = spPr.getElementsByTagName("a:ln")[0];
    expect(ln.getAttribute("w")).toBe("28575");
    expect(names(ln)).toEqual(["solidFill", "prstDash"]);
    expect(ln.getElementsByTagName("a:prstDash")[0].getAttribute("val")).toBe("dash");
    writeShapeStyleOps(sp, [{ kind: "line", dash: "solid" }]);
    expect(names(ln)).toEqual(["solidFill"]);
  });

  it("sets the vertical anchor on bodyPr", () => {
    const sp = parse('<p:spPr/><p:txBody><a:bodyPr anchor="t"/><a:p/></p:txBody>');
    writeShapeStyleOps(sp, [{ kind: "anchor", value: "middle" }]);
    expect(sp.getElementsByTagName("a:bodyPr")[0].getAttribute("anchor")).toBe("ctr");
  });
});

describe("paragraph alignment", () => {
  const para = (text: string): TextParagraph => ({
    style: { align: "left", level: 0, marginLeft: 0, indent: 0 },
    runs: [{ content: text, style: { fontSize: 10, color: "#000", bold: false, align: "left" } }]
  });

  it("aligns only the paragraphs the range touches (a caret touches one)", () => {
    const paragraphs = [para("one"), para("two"), para("three")];
    applyTextOp(vdomTextModel(paragraphs, 6350), { start: 5, end: 5, insert: "", paragraph: { align: "center" } });
    expect(paragraphs.map(p => p.style.align)).toEqual(["left", "center", "left"]);
    expect(paragraphs[1].runs[0].style.align).toBe("center");
    applyTextOp(vdomTextModel(paragraphs, 6350), { start: 0, end: 13, insert: "", paragraph: { align: "right" } });
    expect(paragraphs.map(p => p.style.align)).toEqual(["right", "right", "right"]);
  });

  it("writes algn on a:pPr, creating it when missing", () => {
    const sp = parse('<p:txBody><a:bodyPr/><a:p><a:pPr algn="l"/><a:r><a:t>one</a:t></a:r></a:p><a:p><a:r><a:t>two</a:t></a:r></a:p></p:txBody>');
    writeTextOps(sp, [{ start: 0, end: 7, insert: "", paragraph: { align: "right" } }]);
    const ps = Array.from(sp.getElementsByTagName("a:p"));
    expect(ps.map(p => p.getElementsByTagName("a:pPr")[0].getAttribute("algn"))).toEqual(["r", "r"]);
    expect(ps[1].firstElementChild!.localName).toBe("pPr");
  });
});

describe("namespaces of generated elements", () => {
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const strays = (root: Element) =>
    Array.from(root.getElementsByTagName("*")).filter(node => node.tagName.startsWith("a:") && node.namespaceURI !== A);

  it("keeps every a:* element in the DrawingML namespace, also after a serialize/parse round trip", () => {
    const sp = parse('<p:nvSpPr><p:cNvPr id="2" name="x"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>hello</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody>');
    writeShapeStyleOps(sp, [
      { kind: "fill", color: "#e60012" },
      { kind: "line", color: "#0070c0", widthPt: 1.5, dash: "dash" }
    ]);
    writeTextOps(sp, [
      { start: 2, end: 4, insert: "", format: { color: "#00b050", bold: true } },
      { start: 5, end: 5, insert: "\n" },
      { start: 0, end: 0, insert: "", paragraph: { align: "center" } }
    ]);
    expect(strays(sp)).toEqual([]);

    const again = new DOMParser().parseFromString(new XMLSerializer().serializeToString(sp), "application/xml").documentElement;
    expect(again.getElementsByTagName("parsererror")).toHaveLength(0);
    expect(again.getElementsByTagNameNS(A, "srgbClr").length).toBeGreaterThanOrEqual(3);
    expect(strays(again)).toEqual([]);
  });
});
