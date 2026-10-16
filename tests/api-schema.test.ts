import { describe, expect, it } from "vitest";
import { ApiError } from "../src/api/errors";
import { OPS, batchSchema, parseBatch } from "../src/api/schema";

const fail = (input: unknown): ApiError => {
  try {
    parseBatch(input);
  } catch (error) {
    if (error instanceof ApiError) return error;
    throw error;
  }
  throw new Error("expected the batch to be rejected");
};

describe("batch schema", () => {
  it("accepts the example from the design: label, atomic, aliases, targets of every kind", () => {
    const batch = parseBatch({
      label: "改标题并加图",
      commands: [
        { op: "setText", target: { slide: 1, placeholder: "title" }, text: "2025 年市场分析" },
        { op: "setFill", target: { slide: 1, name: "背景条" }, color: "#E60012" },
        { op: "addImage", slide: 1, src: "data:image/png;base64,AAAA", as: "logo" },
        { op: "setBounds", target: "$logo", x: 28, y: 1, w: 4 },
        { op: "setLine", target: "@selection", color: null },
        { op: "align", targets: [{ slide: 2, id: "5" }, "$logo"], to: "slide", edge: "hcenter" }
      ]
    });
    expect(batch.atomic).toBe(true);
    expect(batch.label).toBe("改标题并加图");
    expect(batch.commands).toHaveLength(6);
  });

  it("takes a bare array of commands, and keeps atomic: false", () => {
    expect(parseBatch([{ op: "deleteSlide", slide: 3 }]).commands).toHaveLength(1);
    expect(parseBatch({ atomic: false, commands: [{ op: "deleteSlide", slide: 3 }] }).atomic).toBe(false);
  });

  it("knows every operation in the design", () => {
    expect(OPS).toEqual(expect.arrayContaining([
      "addSlide", "duplicateSlide", "deleteSlide", "moveSlide", "addText", "addShape", "addImage",
      "addChart", "setChart", "setText", "replaceText", "formatText", "setParagraph", "setBounds", "setRotation", "setFill",
      "setLine", "setAnchor", "align", "arrange", "duplicate", "copyTo", "delete"
    ]));
    expect(OPS).toHaveLength(23);
    expect(batchSchema.properties.commands.items.oneOf).toHaveLength(23);
  });

  it("rejects an unknown op and lists the ones that exist", () => {
    const error = fail([{ op: "explode", slide: 1 }]);
    expect(error.code).toBe("INVALID_SCHEMA");
    expect(error.message).toContain("未知的 op");
    expect(error.message).toContain("setText");
  });

  it("points at the exact field that is wrong", () => {
    expect(fail([{ op: "setFill", target: "$a" }])).toMatchObject({ path: "commands[0]", message: expect.stringContaining('缺少字段 "color"') });
    expect(fail([{ op: "setFill", target: "$a", color: "red" }]).path).toBe("commands[0].color");
    expect(fail([{ op: "setFill", target: "$a", color: "#12345" }]).message).toContain("格式不对");
    expect(fail([{ op: "setBounds", target: "$a", x: "3" }]).path).toBe("commands[0].x");
    expect(fail([{ op: "setParagraph", target: "$a", align: "justify" }]).message).toContain('["left","center","right"]');
  });

  it("rejects a typo in a field name instead of silently ignoring it", () => {
    const error = fail([{ op: "setFill", target: "$a", colour: "#ffffff", color: "#ffffff" }]);
    expect(error.message).toContain('不认识的字段 "colour"');
  });

  it("requires exactly one way to point at a shape", () => {
    expect(fail([{ op: "delete", target: { slide: 1 } }]).message).toContain("指向方式");
    expect(fail([{ op: "delete", target: { slide: 1, id: "3", name: "x" } }]).message).toContain("指向方式");
    expect(() => parseBatch([{ op: "delete", target: { slide: 1, id: 3 } }])).not.toThrow();
    expect(fail([{ op: "delete", target: "nobody" }]).code).toBe("INVALID_SCHEMA");
  });

  it("needs something to change, and text or paragraphs for a text box", () => {
    expect(fail([{ op: "setBounds", target: "$a" }]).message).toContain("至少");
    expect(fail([{ op: "formatText", target: "$a" }]).message).toContain("至少");
    expect(fail([{ op: "addText", slide: 1, rect: { x: 1, y: 1, w: 5, h: 2 } }]).code).toBe("INVALID_SCHEMA");
    expect(() => parseBatch([{ op: "addText", slide: 1, rect: { x: 1, y: 1, w: 5, h: 2 }, paragraphs: [{ text: "a", align: "center" }] }])).not.toThrow();
  });

  it("rejects an empty batch, a bad alias and an image that is not a data URL", () => {
    expect(fail({ commands: [] }).code).toBe("INVALID_SCHEMA");
    expect(fail([{ op: "duplicate", target: "$a", as: "9bad" }]).path).toBe("commands[0].as");
    expect(fail([{ op: "addImage", slide: 1, src: "https://example.com/a.png" }]).path).toBe("commands[0].src");
  });
});
