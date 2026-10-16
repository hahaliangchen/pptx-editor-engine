import { describe, expect, it } from "vitest";
import { TextOp } from "../src/editor/text-ops";
import { writeTextOps } from "../src/pptx-text-writer";

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const shapeXml = (paragraphs: string) =>
  `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="4" name="t"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>${paragraphs}</p:txBody></p:sp>`;
const run = (text: string, attrs = "", inner = "") =>
  `<a:r><a:rPr lang="zh-CN" ${attrs}>${inner}</a:rPr><a:t>${text}</a:t></a:r>`;

function apply(paragraphs: string, ops: TextOp[]): Element {
  const doc = new DOMParser().parseFromString(shapeXml(paragraphs), "application/xml");
  const shape = doc.documentElement;
  expect(writeTextOps(shape, ops)).toBe(true);
  return shape;
}
const summary = (shape: Element) =>
  Array.from(shape.getElementsByTagName("a:p")).map(p =>
    Array.from(p.getElementsByTagName("a:r")).map(r => {
      const rPr = r.getElementsByTagName("a:rPr")[0];
      return `${rPr?.getAttribute("b") === "1" ? "B" : "-"}:${r.getElementsByTagName("a:t")[0].textContent}`;
    })
  );

describe("writeTextOps", () => {
  it("edits inside a run and leaves the other runs byte-identical", () => {
    const shape = apply(
      `<a:p>${run("Hello ", 'b="1"')}${run("world")}</a:p>`,
      [{ start: 8, end: 8, insert: "XX" }]
    );
    expect(summary(shape)).toEqual([["B:Hello ", "-:woXXrld"]]);
  });

  it("splits a paragraph and copies its paragraph properties", () => {
    const shape = apply(
      `<a:p><a:pPr marL="171450" indent="-171450"><a:buChar char="•"/></a:pPr>${run("abcdef", 'b="1"')}<a:endParaRPr lang="zh-CN"/></a:p>`,
      [{ start: 3, end: 3, insert: "\n" }]
    );
    expect(summary(shape)).toEqual([["B:abc"], ["B:def"]]);
    const paragraphs = Array.from(shape.getElementsByTagName("a:p"));
    for (const p of paragraphs) {
      expect(p.getElementsByTagName("a:pPr")[0].getAttribute("marL")).toBe("171450");
      expect(p.getElementsByTagName("a:buChar")).toHaveLength(1);
    }
  });

  it("merges paragraphs when a line break is removed", () => {
    const shape = apply(`<a:p>${run("one")}</a:p><a:p>${run("two", 'b="1"')}</a:p>`, [{ start: 3, end: 4, insert: "" }]);
    expect(summary(shape)).toEqual([["-:one", "B:two"]]);
  });

  it("applies several hunks given in descending order", () => {
    const shape = apply(`<a:p>${run("abcdefghij", 'b="1"')}</a:p>`, [
      { start: 8, end: 9, insert: "Z" },
      { start: 2, end: 2, insert: "Y" }
    ]);
    expect(summary(shape)).toEqual([["B:abYcdefghZj"]]);
  });

  it("writes bold, italic, size and colour, with the fill after a:ln", () => {
    const shape = apply(
      `<a:p>${run("abcdef", 'sz="1150"', '<a:ln w="1"/><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="x"/>')}</a:p>`,
      [{ start: 2, end: 4, insert: "", format: { bold: true, italic: true, fontSizePt: 14, color: "#e60012" } }]
    );
    const runs = Array.from(shape.getElementsByTagName("a:r"));
    expect(runs.map(r => r.getElementsByTagName("a:t")[0].textContent)).toEqual(["ab", "cd", "ef"]);
    const middle = runs[1].getElementsByTagName("a:rPr")[0];
    expect(middle.getAttribute("b")).toBe("1");
    expect(middle.getAttribute("i")).toBe("1");
    expect(middle.getAttribute("sz")).toBe("1400");
    const order = Array.from(middle.children).map(child => child.localName);
    expect(order).toEqual(["ln", "solidFill", "latin"]);
    expect(middle.getElementsByTagName("a:srgbClr")[0].getAttribute("val")).toBe("E60012");
    expect(runs[0].getElementsByTagName("a:rPr")[0].getAttribute("b")).toBeNull();
    expect(runs[0].getElementsByTagName("a:schemeClr")).toHaveLength(1);
  });

  it("creates a run from endParaRPr when typing into an empty paragraph", () => {
    const shape = apply(`<a:p><a:endParaRPr lang="zh-CN" sz="2000"/></a:p>`, [{ start: 0, end: 0, insert: "hi" }]);
    const rPr = shape.getElementsByTagName("a:rPr")[0];
    expect(rPr.getAttribute("sz")).toBe("2000");
    expect(shape.getElementsByTagName("a:t")[0].textContent).toBe("hi");
  });
});
