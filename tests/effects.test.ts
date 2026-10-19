import { describe, expect, it } from "vitest";
import { splitColor, withAlpha } from "../src/color";
import { applyShapeStyleOp } from "../src/editor/shape-ops";
import { writeShapeStyleOps } from "../src/pptx-style-writer";
import { ImageElement, ShapeElement } from "../src/pptx-virtual-dom";

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const shape = (spPr: string) =>
  new DOMParser().parseFromString(`<p:sp ${NS}><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${spPr}</p:spPr></p:sp>`, "application/xml").documentElement;
const base = '<a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
const names = (node: Element) => Array.from(node.getElementsByTagNameNS("*", "spPr")[0].children).map(child => child.localName);
const spPrOf = (node: Element) => node.getElementsByTagNameNS("*", "spPr")[0];
const first = (node: Element, tag: string) => node.getElementsByTagNameNS("*", tag)[0];

describe("colours", () => {
  it("splits a see-through colour into plain colour and opacity", () => {
    expect(splitColor("rgba(255,0,0,0.5)")).toEqual({ hex: "#ff0000", opacity: 0.5 });
    expect(splitColor("#C00000")).toEqual({ hex: "#c00000", opacity: 1 });
    expect(splitColor("#f00")).toEqual({ hex: "#ff0000", opacity: 1 });
  });

  it("writes a colour at an opacity, plain when fully opaque", () => {
    expect(withAlpha("#ff0000", 0.25)).toBe("rgba(255,0,0,0.25)");
    expect(withAlpha("rgba(255,0,0,0.25)", 1)).toBe("#ff0000");
    expect(withAlpha("#112233", 1)).toBe("#112233");
  });
});

describe("writing fills and shadows into a shape", () => {
  it("writes a solid fill with its opacity", () => {
    const node = shape(base);
    writeShapeStyleOps(node, [{ kind: "fill", color: "#c00000", opacity: 0.4 }]);
    const color = first(node, "srgbClr");
    expect(color.getAttribute("val")).toBe("C00000");
    expect(first(color, "alpha").getAttribute("val")).toBe("40000");
    expect(names(node)).toEqual(["xfrm", "prstGeom", "solidFill"]);
  });

  it("writes a linear gradient: stops in order with their opacity, and the angle", () => {
    const node = shape(`${base}<a:ln w="9525"><a:noFill/></a:ln>`);
    writeShapeStyleOps(node, [{ kind: "gradient", gradient: { kind: "linear", angle: 45, stops: [
      { position: 0, color: "#ff0000" }, { position: 0.5, color: "#00ff00", opacity: 0.5 }, { position: 1, color: "#0000ff" }
    ] } }]);
    expect(names(node)).toEqual(["xfrm", "prstGeom", "gradFill", "ln"]);
    const stops = Array.from(node.getElementsByTagNameNS("*", "gs"));
    expect(stops.map(stop => stop.getAttribute("pos"))).toEqual(["0", "50000", "100000"]);
    expect(first(stops[1], "alpha").getAttribute("val")).toBe("50000");
    expect(first(node, "lin").getAttribute("ang")).toBe("2700000");
  });

  it("writes a radial gradient as a circular path", () => {
    const node = shape(base);
    writeShapeStyleOps(node, [{ kind: "gradient", gradient: { kind: "radial", angle: 0, stops: [{ position: 0, color: "#ffffff" }, { position: 1, color: "#000000" }] } }]);
    expect(first(node, "path").getAttribute("path")).toBe("circle");
    expect(first(node, "lin")).toBeUndefined();
  });

  it("replaces whatever fill was there with the gradient", () => {
    const node = shape(`${base}<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`);
    writeShapeStyleOps(node, [{ kind: "gradient", gradient: { kind: "linear", angle: 0, stops: [{ position: 0, color: "#000000" }, { position: 1, color: "#ffffff" }] } }]);
    expect(names(node).filter(name => name.endsWith("Fill"))).toEqual(["gradFill"]);
  });

  it("changes only the opacity of a fill the file spells out, theme colours included", () => {
    const node = shape(`${base}<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>`);
    writeShapeStyleOps(node, [{ kind: "opacity", opacity: 0.3 }]);
    const color = first(node, "schemeClr");
    expect(color.getAttribute("val")).toBe("accent1");
    expect(first(color, "lumMod")).toBeTruthy();
    expect(first(color, "alpha").getAttribute("val")).toBe("30000");
    writeShapeStyleOps(node, [{ kind: "opacity", opacity: 0.6 }]);
    expect(color.getElementsByTagNameNS("*", "alpha")).toHaveLength(1);
    expect(first(color, "alpha").getAttribute("val")).toBe("60000");
    writeShapeStyleOps(node, [{ kind: "opacity", opacity: 1 }]);
    expect(color.getElementsByTagNameNS("*", "alpha")).toHaveLength(0);
  });

  it("changes the opacity of every stop of a gradient", () => {
    const node = shape(`${base}<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>`);
    writeShapeStyleOps(node, [{ kind: "opacity", opacity: 0.5 }]);
    expect(Array.from(node.getElementsByTagNameNS("*", "alpha")).map(a => a.getAttribute("val"))).toEqual(["50000", "50000"]);
  });

  it("writes out a fill that only a theme style gave the shape, so its opacity can be set", () => {
    const node = shape(base);
    writeShapeStyleOps(node, [{ kind: "opacity", opacity: 0.5, base: { type: "solid", color: "#4472c4" } }]);
    expect(first(node, "srgbClr").getAttribute("val")).toBe("4472C4");
    expect(first(node, "alpha").getAttribute("val")).toBe("50000");
  });

  it("adds a drop shadow after the outline", () => {
    const node = shape(`${base}<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="9525"><a:noFill/></a:ln>`);
    writeShapeStyleOps(node, [{ kind: "shadow", shadow: { color: "#000000", opacity: 0.4, blurPt: 8, distancePt: 5, angle: 45 } }]);
    expect(names(node)).toEqual(["xfrm", "prstGeom", "solidFill", "ln", "effectLst"]);
    const shadow = first(node, "outerShdw");
    expect(shadow.getAttribute("blurRad")).toBe(String(8 * 12700));
    expect(shadow.getAttribute("dist")).toBe(String(5 * 12700));
    expect(shadow.getAttribute("dir")).toBe("2700000");
    expect(first(shadow, "alpha").getAttribute("val")).toBe("40000");
  });

  it("replaces a shadow and keeps the other effects in the schema's order", () => {
    const node = shape(`${base}<a:effectLst><a:glow rad="1000"><a:srgbClr val="FFFF00"/></a:glow><a:outerShdw blurRad="1" dist="1" dir="0"><a:srgbClr val="000000"/></a:outerShdw><a:softEdge rad="500"/></a:effectLst>`);
    writeShapeStyleOps(node, [{ kind: "shadow", shadow: { color: "#ff0000", opacity: 1, blurPt: 2, distancePt: 3, angle: 90 } }]);
    const list = first(node, "effectLst");
    expect(Array.from(list.children).map(child => child.localName)).toEqual(["glow", "outerShdw", "softEdge"]);
    expect(list.getElementsByTagNameNS("*", "outerShdw")).toHaveLength(1);
    expect(first(list, "outerShdw").getAttribute("dir")).toBe("5400000");
  });

  it("removes a shadow but leaves an empty list, which also switches off one the theme gives", () => {
    const node = shape(`${base}<a:effectLst><a:outerShdw blurRad="1" dist="1" dir="0"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>`);
    writeShapeStyleOps(node, [{ kind: "shadow", shadow: null }]);
    expect(first(node, "effectLst").children).toHaveLength(0);
  });

  it("puts a picture's shadow in its shape properties", () => {
    const pic = new DOMParser().parseFromString(
      `<p:pic ${NS}><p:nvPicPr><p:cNvPr id="3" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr>${base}</p:spPr></p:pic>`, "application/xml").documentElement;
    writeShapeStyleOps(pic, [{ kind: "shadow", shadow: { color: "#000000", opacity: 0.5, blurPt: 4, distancePt: 4, angle: 45 } }]);
    expect(names(pic)).toEqual(["xfrm", "prstGeom", "effectLst"]);
  });
});

describe("the model after a look change", () => {
  const rect = (): ShapeElement => ({
    type: "shape", id: "s", rect: { x: 0, y: 0, w: 10, h: 10 }, shapeType: "rect", fill: "#4472c4",
    computedStyle: { fill: { type: "solid", color: "#4472c4" } }
  });

  it("sets a gradient with sorted stops and the colours' opacity", () => {
    const shapeElement = rect();
    applyShapeStyleOp(shapeElement, { kind: "gradient", gradient: { kind: "linear", angle: 90, stops: [
      { position: 1, color: "#0000ff", opacity: 0.5 }, { position: 0, color: "#ff0000" }
    ] } }, 12700);
    expect(shapeElement.computedStyle!.fill).toMatchObject({
      type: "gradient", kind: "linear", angle: 90,
      stops: [{ position: 0, color: "#ff0000" }, { position: 1, color: "rgba(0,0,255,0.5)" }]
    });
  });

  it("sets the opacity of a solid fill and puts it back", () => {
    const shapeElement = rect();
    applyShapeStyleOp(shapeElement, { kind: "opacity", opacity: 0.5 }, 12700);
    expect(shapeElement.computedStyle!.fill).toEqual({ type: "solid", color: "rgba(68,114,196,0.5)" });
    applyShapeStyleOp(shapeElement, { kind: "opacity", opacity: 1 }, 12700);
    expect(shapeElement.computedStyle!.fill).toEqual({ type: "solid", color: "#4472c4" });
  });

  it("measures a shadow in pixels for the slide it is on", () => {
    const shapeElement = rect();
    applyShapeStyleOp(shapeElement, { kind: "shadow", shadow: { color: "#000000", opacity: 0.4, blurPt: 10, distancePt: 4, angle: 30 } }, 6350);
    expect(shapeElement.computedStyle!.effects!.outerShadow).toEqual({ color: "#000000", opacity: 0.4, blur: 20, distance: 8, direction: 30 });
    applyShapeStyleOp(shapeElement, { kind: "shadow", shadow: null }, 6350);
    expect(shapeElement.computedStyle!.effects).toBeUndefined();
  });

  it("gives a picture a shadow of its own", () => {
    const picture: ImageElement = { type: "image", id: "i", rect: { x: 0, y: 0, w: 5, h: 5 }, url: "x" };
    applyShapeStyleOp(picture, { kind: "shadow", shadow: { color: "#000000", opacity: 0.5, blurPt: 2, distancePt: 2, angle: 45 } }, 12700);
    expect(picture.effects!.outerShadow).toMatchObject({ blur: 2, distance: 2 });
  });
});
