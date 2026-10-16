import { describe, expect, it } from "vitest";
import { reorderSiblings } from "../src/editor/arrange";
import { remapShapeXml, shapeIdsIn } from "../src/editor/clone";
import { OrderOp } from "../src/pptx-exporter";

const run = (order: string[], selected: string[], mode: OrderOp) =>
  reorderSiblings(order, new Set(selected), mode);

/** Replaying the moves one by one on a plain list must give the same order. */
function replay(order: string[], moves: Array<{ key: string; op: OrderOp }>): string[] {
  const list = [...order];
  for (const { key, op } of moves) {
    const i = list.indexOf(key);
    list.splice(i, 1);
    if (op === "front") list.push(key);
    else if (op === "back") list.unshift(key);
    else if (op === "forward") list.splice(Math.min(i + 1, list.length), 0, key);
    else list.splice(Math.max(i - 1, 0), 0, key);
  }
  return list;
}

describe("reorderSiblings", () => {
  const order = ["a", "b", "c", "d", "e"];

  it("brings to front and sends to back, keeping the selection's relative order", () => {
    expect(run(order, ["b", "d"], "front").order).toEqual(["a", "c", "e", "b", "d"]);
    expect(run(order, ["b", "d"], "back").order).toEqual(["b", "d", "a", "c", "e"]);
  });

  it("does nothing when already at the edge", () => {
    expect(run(order, ["e"], "front").moves).toEqual([]);
    expect(run(order, ["a"], "back").moves).toEqual([]);
    expect(run(order, ["e"], "forward").moves).toEqual([]);
    expect(run(order, ["a"], "backward").moves).toEqual([]);
  });

  it("steps forward and backward past one unselected neighbour", () => {
    expect(run(order, ["c"], "forward").order).toEqual(["a", "b", "d", "c", "e"]);
    expect(run(order, ["c"], "backward").order).toEqual(["a", "c", "b", "d", "e"]);
    expect(run(order, ["c", "d"], "forward").order).toEqual(["a", "b", "e", "c", "d"]);
    expect(run(order, ["d", "e"], "forward").order).toEqual(order);
  });

  it("produces moves that replay to the same result on the XML side", () => {
    const modes: OrderOp[] = ["front", "back", "forward", "backward"];
    for (let trial = 0; trial < 500; trial++) {
      const keys = ["a", "b", "c", "d", "e", "f"];
      const selected = keys.filter(() => Math.random() < 0.4);
      const mode = modes[trial % 4];
      const result = run(keys, selected, mode);
      expect(replay(keys, result.moves)).toEqual(result.order);
    }
  });
});

const XML = '<p:grpSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  + '<p:nvGrpSpPr><p:cNvPr id="5" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + '<p:pic><p:nvPicPr><p:cNvPr id="6" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/></p:blipFill><p:spPr/></p:pic>'
  + '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="C"/><p:cNvCxnSpPr><a:stCxn id="6" idx="0"/><a:endCxn id="5" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>'
  + '</p:grpSp>';

describe("remapShapeXml", () => {
  it("lists the ids in document order", () => {
    expect(shapeIdsIn(XML)).toEqual(["5", "6", "7"]);
  });

  it("gives every shape a fresh id, remaps relationships and drops connector attachments", () => {
    let next = 100;
    const result = remapShapeXml(XML, () => String(next++), new Map([["rId3", "rId9"]]));
    expect([...result.idMap]).toEqual([["5", "100"], ["6", "101"], ["7", "102"]]);
    expect(shapeIdsIn(result.xml)).toEqual(["100", "101", "102"]);
    expect(result.xml).toContain('r:embed="rId9"');
    expect(result.xml).not.toContain("stCxn");
    expect(result.xml).not.toContain("endCxn");
    // still well-formed with the namespaces intact
    const again = new DOMParser().parseFromString(result.xml, "application/xml");
    expect(again.getElementsByTagName("parsererror")).toHaveLength(0);
    expect(again.documentElement.localName).toBe("grpSp");
  });
});
