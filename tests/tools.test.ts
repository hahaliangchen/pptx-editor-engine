// @vitest-environment node
import { afterEach, describe, expect, it } from "vitest";
import Ajv from "ajv";
import { HeadlessSession } from "../src/headless/session";
import { batchSchema, sharedSchemas } from "../src/api/schema";
import { TOOLS, TOOL_GUIDE, ToolHost, callTool } from "../src/api/tools";
import {
  callAnthropicTool, callMcpTool, callOpenAITool, shareDefinitions, toAnthropic, toMcp, toOpenAI
} from "../src/api/adapters";
import { buildDeck } from "./helpers/deck";
import { apiDeckParts } from "./helpers/api-deck";

const sessions: HeadlessSession[] = [];
async function open(): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(await buildDeck(apiDeckParts()));
  sessions.push(session);
  return session;
}
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

const json = (result: { content: Array<{ type: string; text?: string }> }) =>
  JSON.parse(result.content.find(part => part.type === "text")!.text!);
const PNG = [0x89, 0x50, 0x4e, 0x47];

describe("tool definitions", () => {
  it("offers the three tools, each taking an object", () => {
    expect(TOOLS.map(tool => tool.name)).toEqual(["get_outline", "apply_edits", "render_slide"]);
    for (const tool of TOOLS) {
      expect(tool.inputSchema.type).toBe("object");
      expect(tool.description.length).toBeGreaterThan(40);
    }
    expect(TOOLS[1].inputSchema).toBe(batchSchema);
  });

  it("explains how to work with them", () => {
    for (const tool of TOOLS) expect(TOOL_GUIDE).toContain(tool.name);
    expect(TOOL_GUIDE).toContain("从 1 开始");
  });
});

describe("adapters", () => {
  it("gives Anthropic its input_schema", () => {
    const tools = toAnthropic();
    expect(tools.map(tool => Object.keys(tool).sort())).toEqual(Array(3).fill(["description", "input_schema", "name"]));
  });

  it("wraps each tool as an OpenAI function", () => {
    for (const tool of toOpenAI()) {
      expect(tool.type).toBe("function");
      expect(Object.keys(tool.function).sort()).toEqual(["description", "name", "parameters"]);
      expect(tool.function.parameters.type).toBe("object");
    }
  });

  it("gives MCP its inputSchema", () => {
    expect(toMcp().map(tool => Object.keys(tool).sort())).toEqual(Array(3).fill(["description", "inputSchema", "name"]));
  });

  it("hands out private copies without the $schema marker some providers reject", () => {
    for (const tools of [toAnthropic(), toOpenAI().map(tool => tool.function), toMcp()]) {
      const schema = ("input_schema" in tools[1] ? tools[1].input_schema : "parameters" in tools[1] ? tools[1].parameters : (tools[1] as { inputSchema: Record<string, unknown> }).inputSchema) as Record<string, unknown>;
      expect(schema.$schema).toBeUndefined();
      (schema.properties as Record<string, unknown>).tampered = true;
    }
    expect((TOOLS[1].inputSchema.properties as Record<string, unknown>).tampered).toBeUndefined();
    expect(batchSchema.$schema).toBeDefined();
  });

  it("shares repeated fragments as definitions, in half the size or less", () => {
    const shared = toAnthropic()[1].input_schema;
    expect(Object.keys(shared.$defs as object)).toEqual(expect.arrayContaining(["target", "rect", "color"]));
    expect(JSON.stringify(shared).length).toBeLessThan(JSON.stringify(batchSchema).length * 0.6);
  });
});

describe("the shared definitions mean the same thing", () => {
  const original = new Ajv({ strict: false, allErrors: false }).compile(batchSchema as object);
  const shared = new Ajv({ strict: false, allErrors: false }).compile(
    shareDefinitions(JSON.parse(JSON.stringify(batchSchema)), sharedSchemas));
  const target = { slide: 1, name: "卡片" };
  const batches: unknown[] = [
    { commands: [{ op: "addSlide" }] },
    { commands: [{ op: "addSlide", after: 1, layout: "空白" }] },
    { commands: [{ op: "setText", target, text: "x" }] },
    { commands: [{ op: "setText", target: "$a", text: "x" }] },
    { commands: [{ op: "setText", target: "@selection", text: "x" }] },
    { commands: [{ op: "setText", target: { slide: 1, id: 3 }, text: "x" }] },
    { commands: [{ op: "setText", target: { slide: 1, placeholder: "title" }, text: "x" }] },
    { commands: [{ op: "addShape", slide: 1, shape: "roundRect", rect: { x: 1, y: 1, w: 2, h: 2 }, fill: "#ff0000", as: "a" }] },
    { commands: [{ op: "addText", slide: 1, rect: { x: 1, y: 1, w: 2, h: 2 }, paragraphs: [{ text: "a", align: "center" }], style: { bold: true, color: "#000000" } }] },
    { commands: [{ op: "setFill", target, color: null }] },
    { commands: [{ op: "setLine", target, color: "#00ff00", dash: "dash", widthPt: 2 }] },
    { commands: [{ op: "addChart", slide: 1, type: "pie", categories: ["a"], series: [{ name: "s", values: [1, null], color: "#112233" }], legend: "right" }] },
    { commands: [{ op: "setChart", target, title: null }] },
    { commands: [{ op: "addTable", slide: 1, rows: [["a", "b"]], columnWidths: [1, 2] }] },
    { commands: [{ op: "setCell", target, row: 1, col: 1, fill: null, align: "right" }] },
    { commands: [{ op: "insertRow", target, at: 2, text: ["x"] }] },
    { label: "x", atomic: false, commands: [{ op: "delete", target }] },
    // invalid ones
    { commands: [] },
    { commands: [{ op: "explode" }] },
    { commands: [{ op: "setText", target }] },
    { commands: [{ op: "setText", target: { slide: 1 }, text: "x" }] },
    { commands: [{ op: "setText", target: { slide: 1, id: 1, name: "a" }, text: "x" }] },
    { commands: [{ op: "setText", target: "alias", text: "x" }] },
    { commands: [{ op: "setFill", target, color: "red" }] },
    { commands: [{ op: "addShape", slide: 0, shape: "rect", rect: { x: 1, y: 1, w: 2, h: 2 } }] },
    { commands: [{ op: "addShape", slide: 1, shape: "star", rect: { x: 1, y: 1, w: 2, h: 2 } }] },
    { commands: [{ op: "addShape", slide: 1, shape: "rect", rect: { x: 1, y: 1, w: 0, h: 2 } }] },
    { commands: [{ op: "addChart", slide: 1, type: "radar", categories: ["a"], series: [{ name: "s", values: [1] }] }] },
    { commands: [{ op: "addChart", slide: 1, type: "pie", categories: ["a"], series: [{ name: "s", values: ["x"] }] }] },
    { commands: [{ op: "addTable", slide: 1, rows: [] }] },
    { commands: [{ op: "setCell", target, row: 0, col: 1, text: "x" }] },
    { commands: [{ op: "setCell", target, row: 1, col: 1 }] },
    { commands: [{ op: "setText", target, text: "x", extra: 1 }] },
    { commands: "nope" },
    { atomic: "yes", commands: [{ op: "delete", target }] }
  ];

  it("accepts and rejects exactly the same batches", () => {
    const verdicts = batches.map(batch => [Boolean(original(batch)), Boolean(shared(batch))]);
    expect(verdicts.map(([a, b]) => a === b)).toEqual(batches.map(() => true));
    expect(verdicts.filter(([a]) => a).length).toBe(17);
    expect(verdicts.filter(([a]) => !a).length).toBe(batches.length - 17);
  });
});

describe("calling the tools", () => {
  it("reads the outline, all of it or some pages, with a compact text", async () => {
    const session = await open();
    const all = await callTool(session, "get_outline", {});
    expect(all.ok).toBe(true);
    expect(all.content).toHaveLength(1);
    expect(json(all).slides[0].shapes.map((shape: { name: string }) => shape.name)).toContain("卡片");
    expect(all.content[0].type === "text" && all.content[0].text.includes("\n")).toBe(false);
    const some = json(await callTool(session, "get_outline", { slides: [1], maxTextLength: 20 }));
    expect(some.slides).toHaveLength(1);
  });

  it("applies edits and reports the result as JSON", async () => {
    const session = await open();
    const result = await callTool(session, "apply_edits", {
      commands: [{ op: "addText", slide: 1, rect: { x: 1, y: 1, w: 5, h: 2 }, text: "新文字", name: "新", as: "n" }]
    });
    expect(result.ok).toBe(true);
    expect(json(result)).toMatchObject({ ok: true, results: [{ ok: true, created: { name: "新" } }] });
    expect(json(await callTool(session, "get_outline", {})).slides[0].shapes.map((s: { name: string }) => s.name)).toContain("新");
  });

  it("returns a failed batch as a failed result a model can act on", async () => {
    const session = await open();
    const missing = await callTool(session, "apply_edits", { commands: [{ op: "setText", target: { slide: 1, name: "没有这个" }, text: "x" }] });
    expect(missing.ok).toBe(false);
    expect(json(missing)).toMatchObject({ ok: false, failedAt: 0, error: { code: "SHAPE_NOT_FOUND", path: "commands[0].target.name" } });

    const invalid = await callTool(session, "apply_edits", { commands: [{ op: "setFill", target: { slide: 1, name: "卡片" }, color: "red" }] });
    expect(invalid.ok).toBe(false);
    expect(json(invalid).error).toMatchObject({ code: "INVALID_SCHEMA", path: expect.stringContaining("commands[0]") });
  });

  it("renders a page as a PNG image with a caption", async () => {
    const session = await open();
    const result = await callTool(session, "render_slide", { slide: 1, width: 480 });
    expect(result.ok).toBe(true);
    const image = result.content.find(part => part.type === "image")!;
    expect(image).toMatchObject({ type: "image", mimeType: "image/png" });
    expect([...Buffer.from((image as { data: string }).data, "base64").subarray(0, 4)]).toEqual(PNG);
    expect(result.content.find(part => part.type === "text")).toMatchObject({ text: expect.stringContaining("480") });
  });

  it("says what is wrong with an argument instead of throwing", async () => {
    const session = await open();
    for (const [name, args, message] of [
      ["get_outline", { slides: [] }, "slides"],
      ["get_outline", { slides: [0] }, "slides"],
      ["get_outline", { maxTextLength: 3 }, "maxTextLength"],
      ["render_slide", {}, "slide"],
      ["render_slide", { slide: 1.5 }, "slide"],
      ["render_slide", { slide: 1, width: 50 }, "width"],
      ["render_slide", { slide: 99 }, "没有第 99 页"],
      ["no_such_tool", {}, "no_such_tool"]
    ] as const) {
      const result = await callTool(session, name, args);
      expect(result.ok, `${name} ${JSON.stringify(args)}`).toBe(false);
      expect(JSON.stringify(json(result))).toContain(message);
    }
    expect((await callTool(session, "apply_edits", undefined)).ok).toBe(false);
  });

  it("copes with a host that cannot draw, and one that breaks", async () => {
    const host: ToolHost = {
      outline: async () => { throw new Error("坏了"); },
      apply: async () => ({ ok: true, results: [] })
    };
    expect(json(await callTool(host, "render_slide", { slide: 1 })).error.message).toContain("不能渲染");
    const broken = await callTool(host, "get_outline", {});
    expect(broken.ok).toBe(false);
    expect(json(broken).error.message).toBe("坏了");
  });
});

describe("a round trip on each platform", () => {
  const edit = { commands: [{ op: "addChart", slide: 1, type: "column", categories: ["a", "b"], series: [{ name: "s", values: [1, 2] }], name: "图" }] };

  it("Anthropic: a tool_use block in, a tool_result block out", async () => {
    const session = await open();
    const applied = await callAnthropicTool(session, { id: "toolu_1", name: "apply_edits", input: edit });
    expect(applied).toMatchObject({ type: "tool_result", tool_use_id: "toolu_1" });
    expect(applied.is_error).toBeUndefined();
    expect(JSON.parse((applied.content[0] as { text: string }).text).ok).toBe(true);

    const picture = await callAnthropicTool(session, { id: "toolu_2", name: "render_slide", input: { slide: 1 } });
    const image = picture.content.find(part => part.type === "image") as { source: { type: string; media_type: string; data: string } };
    expect(image.source).toMatchObject({ type: "base64", media_type: "image/png" });
    expect([...Buffer.from(image.source.data, "base64").subarray(0, 4)]).toEqual(PNG);

    const failed = await callAnthropicTool(session, { id: "toolu_3", name: "apply_edits", input: { commands: [{ op: "delete", target: { slide: 1, name: "无" } }] } });
    expect(failed.is_error).toBe(true);
  });

  it("OpenAI: arguments arrive as a string, the image follows in a user message", async () => {
    const session = await open();
    const applied = await callOpenAITool(session, { id: "call_1", function: { name: "apply_edits", arguments: JSON.stringify(edit) } });
    expect(applied).toHaveLength(1);
    expect(applied[0]).toMatchObject({ role: "tool", tool_call_id: "call_1" });

    const picture = await callOpenAITool(session, { id: "call_2", function: { name: "render_slide", arguments: '{"slide":1}' } });
    expect(picture.map(message => message.role)).toEqual(["tool", "user"]);
    const user = picture[1] as { content: Array<{ type: string; image_url?: { url: string } }> };
    expect(user.content.find(part => part.type === "image_url")!.image_url!.url).toMatch(/^data:image\/png;base64,/);

    const garbled = await callOpenAITool(session, { id: "call_3", function: { name: "apply_edits", arguments: "{not json" } });
    expect(JSON.parse((garbled[0] as { content: string }).content).error.message).toContain("合法的 JSON");
    const empty = await callOpenAITool(session, { id: "call_4", function: { name: "get_outline", arguments: "" } });
    expect(JSON.parse((empty[0] as { content: string }).content).slides).toHaveLength(1);
  });

  it("MCP: content parts, and isError only when it failed", async () => {
    const session = await open();
    const outline = await callMcpTool(session, "get_outline", {});
    expect(outline.isError).toBeUndefined();
    expect(outline.content[0].type).toBe("text");
    const picture = await callMcpTool(session, "render_slide", { slide: 1 });
    expect(picture.content.find(part => part.type === "image")).toMatchObject({ mimeType: "image/png" });
    expect((await callMcpTool(session, "nope", {})).isError).toBe(true);
  });

  it("a session runs tools itself", async () => {
    const session = await open();
    expect((await session.callTool("apply_edits", edit)).ok).toBe(true);
  });
});
