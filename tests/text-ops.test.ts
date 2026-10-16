import { describe, expect, it } from "vitest";
import { applyTextOp, diffText, flatText, TextOp } from "../src/editor/text-ops";
import { normalizeParagraphs, paragraphsText, vdomTextModel } from "../src/editor/text-model-vdom";
import { TextParagraph } from "../src/pptx-virtual-dom";

const style = (extra = {}) => ({ fontSize: 20, color: "#000000", bold: false, align: "left" as const, ...extra });
const para = (...runs: Array<[string, object?]>): TextParagraph => ({
  style: { align: "left", level: 0, marginLeft: 0, indent: 0 },
  runs: runs.map(([content, extra]) => ({ content, style: style(extra) }))
});
const applyAll = (text: string, ops: TextOp[]) =>
  ops.reduce((value, op) => value.slice(0, op.start) + op.insert + value.slice(op.end), text);

describe("diffText", () => {
  it("returns nothing for identical text", () => {
    expect(diffText("abc", "abc")).toEqual([]);
  });

  it("reproduces the target text for random edits (descending hunks)", () => {
    const alphabet = ["a", "b", "c", "好", "事", "\n", "😀", " "];
    const random = (max: number) =>
      Array.from({ length: Math.floor(Math.random() * max) }, () => alphabet[Math.floor(Math.random() * alphabet.length)]).join("");
    for (let i = 0; i < 3000; i++) {
      const before = random(14);
      let after = before;
      for (let e = 0; e < Math.floor(Math.random() * 4); e++) {
        const at = Math.floor(Math.random() * (after.length + 1));
        after = after.slice(0, at) + random(3) + after.slice(at + Math.floor(Math.random() * 3));
      }
      const ops = diffText(before, after);
      expect(applyAll(before, ops)).toBe(after);
      ops.forEach((op, index) => {
        if (index > 0) expect(op.end).toBeLessThanOrEqual(ops[index - 1].start);
      });
    }
  });

  it("keeps edits apart when unchanged text lies between them", () => {
    const before = "abc".repeat(100);
    const after = before.slice(0, 50) + "XYZ" + before.slice(50, 200) + before.slice(220);
    expect(diffText(before, after)).toHaveLength(2);
  });
});

describe("applyTextOp on the virtual DOM", () => {
  it("types inside a run without touching neighbours", () => {
    const paragraphs = [para(["Hello ", { bold: true }], ["world"])];
    applyTextOp(vdomTextModel(paragraphs, 6350), { start: 3, end: 3, insert: "XY" });
    expect(paragraphs[0].runs.map(r => [r.content, r.style.bold])).toEqual([["HelXYlo ", true], ["world", false]]);
  });

  it("replacing text keeps the formatting of the text it replaces", () => {
    const paragraphs = [para(["Hello ", { bold: true, color: "#ff0000" }], ["world"])];
    applyTextOp(vdomTextModel(paragraphs, 6350), { start: 0, end: 5, insert: "Howdy" });
    expect(paragraphs[0].runs.map(r => [r.content, r.style.bold, r.style.color])).toEqual([
      ["Howdy ", true, "#ff0000"], ["world", false, "#000000"]
    ]);
    // replacing everything: the new text takes the first run's look
    const whole = [para(["Old title", { bold: true, fontSize: 40 }]), para(["second"])];
    applyTextOp(vdomTextModel(whole, 6350), { start: 0, end: 16, insert: "New" });
    expect(whole).toHaveLength(1);
    expect(whole[0].runs.map(r => [r.content, r.style.bold, r.style.fontSize])).toEqual([["New", true, 40]]);
  });

  it("splits a paragraph on Enter and keeps the formatting after the caret", () => {
    const paragraphs = [para(["Hello ", { bold: true }], ["world"])];
    applyTextOp(vdomTextModel(paragraphs, 6350), { start: 8, end: 8, insert: "\n" });
    expect(paragraphsText(paragraphs)).toBe("Hello wo\nrld");
    expect(paragraphs[1].runs[0].style.bold).toBe(false);
    expect(paragraphs[0].runs.map(r => r.content)).toEqual(["Hello ", "wo"]);
  });

  it("merges paragraphs when the line break is deleted", () => {
    const paragraphs = [para(["one"]), para(["two", { bold: true }])];
    applyTextOp(vdomTextModel(paragraphs, 6350), { start: 3, end: 4, insert: "" });
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0].runs.map(r => [r.content, r.style.bold])).toEqual([["one", false], ["two", true]]);
  });

  it("deleting everything leaves one empty run the renderer can use", () => {
    const paragraphs = [para(["abc"]), para(["def"])];
    applyTextOp(vdomTextModel(paragraphs, 6350), { start: 0, end: 7, insert: "" });
    normalizeParagraphs(paragraphs, style());
    expect(paragraphs).toHaveLength(1);
    expect(paragraphs[0].runs).toHaveLength(1);
    expect(flatText(vdomTextModel(paragraphs, 6350))).toBe("");
  });

  it("formats only the selected part and splits runs at the edges", () => {
    const paragraphs = [para(["Hello ", { bold: true }], ["big world"]), para(["Second"])];
    applyTextOp(vdomTextModel(paragraphs, 6350), {
      start: 3, end: 15, insert: "", format: { color: "#ff0000", italic: true, fontSizePt: 12 }
    });
    expect(paragraphs[0].runs.map(r => [r.content, r.style.bold, r.style.italic ?? false, r.style.color])).toEqual([
      ["Hel", true, false, "#000000"],
      ["lo ", true, true, "#ff0000"],
      ["big world", false, true, "#ff0000"]
    ]);
    expect(paragraphs[1].runs[0].style.color).toBe("#000000");
    // 12pt at 6350 EMU/px
    expect(paragraphs[0].runs[1].style.fontSize).toBeCloseTo(12 * 12700 / 6350);
  });
});
