import { beforeAll, describe, expect, it } from "vitest";
import { PptxParser } from "../src/pptx-parser";
import { PptVirtualDocument } from "../src/pptx-virtual-dom";
import { buildOutline } from "../src/api/outline";
import { ApiError } from "../src/api/errors";
import { ResolveContext, resolveNode, resolveSlide } from "../src/api/resolve";
import { allNodes, buildScene } from "../src/api/scene";
import { apiDeckParts } from "./helpers/api-deck";
import { buildDeck } from "./helpers/deck";

let doc: PptVirtualDocument;

beforeAll(async () => {
  const bytes = await buildDeck(apiDeckParts());
  doc = await new PptxParser().parse(bytes as unknown as ArrayBuffer);
});

const contextOf = (selection: Array<{ part: string; id: string }> = []): ResolveContext => ({
  scenes: buildScene(doc.slides),
  aliases: new Map(),
  selection: () => selection
});

describe("scene tree", () => {
  it("matches the selection pane: top-level shapes, with group members nested", () => {
    const [scene] = buildScene(doc.slides);
    expect(scene.roots.map(node => [node.id, node.name, node.kind])).toEqual([
      ["90", "标题 1", "text"],
      ["10", "卡片", "shape"],
      ["11", "双格式", "text"],
      ["20", "数据组", "group"],
      ["30", "连接线", "line"]
    ]);
    const group = scene.roots.find(node => node.kind === "group")!;
    expect(group.children.map(node => node.name)).toEqual(["指标甲", "指标乙"]);
    expect(group.children[0].parent).toBe(group);
    expect(allNodes(scene).map(node => node.id)).toEqual(["90", "10", "11", "20", "21", "22", "30"]);
  });

  it("does not include the master's and layout's own shapes", () => {
    const [scene] = buildScene(doc.slides);
    expect(allNodes(scene).every(node => node.part === "ppt/slides/slide1.xml")).toBe(true);
  });
});

describe("outline", () => {
  let outline: ReturnType<typeof buildOutline>;
  let shapes: ReturnType<typeof buildOutline>["slides"][number]["shapes"];
  beforeAll(() => {
    outline = buildOutline(doc, { layouts: ["标题和内容"] });
    shapes = outline.slides[0].shapes;
  });

  it("reports sizes in centimetres and the slide's own size", () => {
    expect(outline.slideSize).toEqual({ w: 25.4, h: 14.29 });
    expect(outline.layouts).toEqual(["标题和内容"]);
    const card = shapes.find(shape => shape.name === "卡片")!;
    expect(card.rect).toEqual({ x: 2.54, y: 5.08, w: 10.16, h: 3.81 });
  });

  it("describes fill, border, geometry, text and its style", () => {
    const card = shapes.find(shape => shape.name === "卡片")!;
    expect(card).toMatchObject({ type: "shape", shape: "roundRect", fill: "#4472c4", text: "增长 30%" });
    expect(card.line).toEqual({ color: "#2f528f", widthPt: 1, dash: "solid" });
    const title = shapes.find(shape => shape.placeholder === "title")!;
    expect(title.text).toBe("市场分析");
    expect(title.style).toMatchObject({ fontSizePt: 32, bold: true, align: "center" });
  });

  it("nests group members and can be limited to some slides or shorter text", () => {
    const group = shapes.find(shape => shape.type === "group")!;
    expect(group.children!.map(child => child.text)).toEqual(["收入", "收入"]);
    const short = buildOutline(doc, { maxTextLength: 2 }).slides[0].shapes.find(shape => shape.name === "卡片")!;
    expect(short).toMatchObject({ text: "增长", textTruncated: true });
    expect(buildOutline(doc, { slides: [5] }).slides).toEqual([]);
  });
});

describe("resolving a target", () => {
  it("finds a shape by id, name, placeholder or text, including group members", () => {
    const context = contextOf();
    expect(resolveNode(context, { slide: 1, id: "10" }, "t").node.name).toBe("卡片");
    expect(resolveNode(context, { slide: 1, name: "指标乙" }, "t").node.id).toBe("22");
    expect(resolveNode(context, { slide: 1, placeholder: "title" }, "t").node.id).toBe("90");
    expect(resolveNode(context, { slide: 1, text: "30%" }, "t").node.id).toBe("10");
    expect(resolveNode(context, { slide: "slide_1", id: "20" }, "t").node.kind).toBe("group");
    expect(resolveSlide(context.scenes, "1", "t").index).toBe(1);
  });

  it("reports an ambiguous text match with every candidate rather than picking one", () => {
    const error = catchError(() => resolveNode(contextOf(), { slide: 1, text: "收入" }, "commands[0].target"));
    expect(error.code).toBe("SHAPE_AMBIGUOUS");
    expect(error.path).toBe("commands[0].target.text");
    expect(error.candidates!.map(candidate => candidate.name)).toEqual(["指标甲", "指标乙"]);
  });

  it("reports a missing shape with the shapes that do exist", () => {
    const error = catchError(() => resolveNode(contextOf(), { slide: 1, name: "不存在" }, "t"));
    expect(error.code).toBe("SHAPE_NOT_FOUND");
    expect(error.candidates!.map(candidate => candidate.name)).toContain("卡片");
  });

  it("reports a slide that is out of range with the page count", () => {
    const error = catchError(() => resolveNode(contextOf(), { slide: 4, id: "1" }, "t"));
    expect(error.code).toBe("SLIDE_NOT_FOUND");
    expect(error.message).toContain("共 1 页");
  });

  it("resolves aliases and the selection, and says so when they cannot be", () => {
    const context = contextOf([{ part: "ppt/slides/slide1.xml", id: "10" }]);
    context.aliases.set("hero", { part: "ppt/slides/slide1.xml", id: "30" });
    expect(resolveNode(context, "$hero", "t").node.name).toBe("连接线");
    expect(resolveNode(context, "@selection", "t").node.name).toBe("卡片");
    expect(catchError(() => resolveNode(context, "$nope", "t")).code).toBe("SHAPE_NOT_FOUND");
    expect(catchError(() => resolveNode(contextOf(), "@selection", "t")).code).toBe("SHAPE_NOT_FOUND");
    expect(catchError(() => resolveNode(contextOf([
      { part: "ppt/slides/slide1.xml", id: "10" }, { part: "ppt/slides/slide1.xml", id: "30" }
    ]), "@selection", "t")).code).toBe("SHAPE_AMBIGUOUS");
  });
});

function catchError(fn: () => unknown): ApiError {
  try {
    fn();
  } catch (error) {
    if (error instanceof ApiError) return error;
    throw error;
  }
  throw new Error("expected an ApiError");
}
