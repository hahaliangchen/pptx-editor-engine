import { beforeEach, describe, expect, it } from "vitest";
import { apiDeckParts } from "./helpers/api-deck";
import { openHarness } from "./helpers/harness";

type Harness = Awaited<ReturnType<typeof openHarness>>;
let h: Harness;
beforeEach(async () => { h = await openHarness(apiDeckParts()); });

const shape = async (name: string) => (await h.api.outline()).slides[0].shapes.flatMap(s => [s, ...(s.children || [])]).find(s => s.name === name)!;
/** The same XML regardless of how it was written out (declaration, spacing). */
const canon = (xml: string) => new XMLSerializer().serializeToString(new DOMParser().parseFromString(xml, "application/xml"));
const xmlOf = (xml: string, id: string) => {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const node = Array.from(doc.getElementsByTagName("*")).find(n => n.localName === "cNvPr" && n.getAttribute("id") === id)!.parentNode!.parentNode as Element;
  return { node, text: Array.from(node.getElementsByTagName("a:t")).map(t => t.textContent).join("|") };
};

describe("text commands", () => {
  it("setText replaces the words and keeps the placeholder's size, boldness and alignment", async () => {
    const result = await h.api.apply([{ op: "setText", target: { slide: 1, placeholder: "title" }, text: "2025 年度总结" }]);
    expect(result.ok).toBe(true);
    const title = await shape("标题 1");
    expect(title.text).toBe("2025 年度总结");
    expect(title.style).toMatchObject({ fontSizePt: 32, bold: true, align: "center" });
    expect(xmlOf(await h.exported(), "90").text).toBe("2025 年度总结");
  });

  it("replaceText changes only the matched words and leaves each run's formatting", async () => {
    await h.api.apply([{ op: "replaceText", target: { slide: 1, name: "双格式" }, find: "三成", replace: "30%" }]);
    expect((await shape("双格式")).text).toBe("重点：营收增长30%");
    const { node } = xmlOf(await h.exported(), "11");
    const runs = Array.from(node.getElementsByTagName("a:r"));
    expect(runs.map(r => r.getElementsByTagName("a:t")[0].textContent)).toEqual(["重点：", "营收增长30%"]);
    expect(runs[0].getElementsByTagName("a:rPr")[0].getAttribute("b")).toBe("1");
    expect(runs[0].getElementsByTagName("a:srgbClr")[0].getAttribute("val")).toBe("FF0000");
  });

  it("replaceText replaces every match with all: true, and reports a missing one", async () => {
    await h.api.apply([{ op: "setText", target: { slide: 1, name: "卡片" }, text: "增长 增长 增长" }]);
    await h.api.apply([{ op: "replaceText", target: { slide: 1, name: "卡片" }, find: "增长", replace: "下降", all: true }]);
    expect((await shape("卡片")).text).toBe("下降 下降 下降");
    const miss = await h.api.apply([{ op: "replaceText", target: { slide: 1, name: "卡片" }, find: "不存在", replace: "x" }]);
    expect(miss.ok).toBe(false);
    expect(miss.error!.message).toContain("没有找到");
  });

  it("formatText styles a character range and splits the run there", async () => {
    await h.api.apply([{ op: "formatText", target: { slide: 1, name: "双格式" }, range: { from: 5, to: 7 }, bold: true, color: "#0070C0" }]);
    const { node } = xmlOf(await h.exported(), "11");
    const runs = Array.from(node.getElementsByTagName("a:r"));
    expect(runs.map(r => r.getElementsByTagName("a:t")[0].textContent)).toEqual(["重点：", "营收", "增长", "三成"]);
    expect(runs[2].getElementsByTagName("a:rPr")[0].getAttribute("b")).toBe("1");
    expect(runs[2].getElementsByTagName("a:srgbClr")[0].getAttribute("val")).toBe("0070C0");
    expect(runs[1].getElementsByTagName("a:rPr")[0].getAttribute("b")).toBeNull();
  });

  it("setParagraph aligns the paragraphs", async () => {
    await h.api.apply([{ op: "setParagraph", target: { slide: 1, name: "双格式" }, align: "right" }]);
    expect((await shape("双格式")).style!.align).toBe("right");
    expect(xmlOf(await h.exported(), "11").node.getElementsByTagName("a:pPr")[0].getAttribute("algn")).toBe("r");
  });
});

describe("shape commands", () => {
  it("setFill, setLine and setAnchor are visible in the outline and in the file", async () => {
    await h.api.apply([
      { op: "setFill", target: { slide: 1, name: "卡片" }, color: "#E60012" },
      { op: "setLine", target: { slide: 1, name: "卡片" }, color: "#000000", widthPt: 3, dash: "dash" },
      { op: "setAnchor", target: { slide: 1, name: "卡片" }, anchor: "bottom" }
    ]);
    const card = await shape("卡片");
    expect(card.fill).toBe("#e60012");
    expect(card.line).toEqual({ color: "#000000", widthPt: 3, dash: "dash" });
    const { node } = xmlOf(await h.exported(), "10");
    expect(node.getElementsByTagName("p:spPr")[0].getElementsByTagName("a:srgbClr")[0].getAttribute("val")).toBe("E60012");
    expect(node.getElementsByTagName("a:ln")[0].getAttribute("w")).toBe("38100");
    expect(node.getElementsByTagName("a:prstDash")[0].getAttribute("val")).toBe("dash");
    expect(node.getElementsByTagName("a:bodyPr")[0].getAttribute("anchor")).toBe("b");
  });

  it("setFill with null removes the fill", async () => {
    await h.api.apply([{ op: "setFill", target: { slide: 1, name: "卡片" }, color: null }]);
    expect((await shape("卡片")).fill).toBeNull();
    expect(xmlOf(await h.exported(), "10").node.getElementsByTagName("a:noFill")).toHaveLength(1);
  });

  it("setBounds takes centimetres and keeps what is not mentioned", async () => {
    await h.api.apply([{ op: "setBounds", target: { slide: 1, name: "卡片" }, x: 5, y: 4 }]);
    expect((await shape("卡片")).rect).toEqual({ x: 5, y: 4, w: 10.16, h: 3.81 });
    const off = xmlOf(await h.exported(), "10").node.getElementsByTagName("a:off")[0];
    expect([off.getAttribute("x"), off.getAttribute("y")]).toEqual(["1800000", "1440000"]);
  });

  it("moves a group as a whole and a member inside it", async () => {
    await h.api.apply([{ op: "setBounds", target: { slide: 1, name: "数据组" }, x: 10 }]);
    expect((await shape("数据组")).rect.x).toBe(10);
    expect((await shape("指标甲")).rect.x).toBeCloseTo(10, 1);
    await h.api.apply([{ op: "setBounds", target: { slide: 1, name: "指标乙" }, y: 12 }]);
    expect((await shape("指标乙")).rect.y).toBe(12);
  });

  it("setRotation, arrange, align and delete work on shapes picked by id or text", async () => {
    await h.api.apply([
      { op: "setRotation", target: { slide: 1, id: "10" }, degrees: 30 },
      { op: "arrange", target: { slide: 1, id: "10" }, mode: "back" },
      { op: "align", targets: [{ slide: 1, text: "30%" }], to: "slide", edge: "hcenter" },
      { op: "delete", target: { slide: 1, name: "连接线" } }
    ]);
    const outline = (await h.api.outline()).slides[0].shapes;
    const card = outline.find(s => s.name === "卡片")!;
    expect(card.rotation).toBe(30);
    expect(card.rect.x + card.rect.w / 2).toBeCloseTo(12.7, 1);
    expect(outline[0].name).toBe("卡片"); // sent to the back
    expect(outline.some(s => s.name === "连接线")).toBe(false);
    expect(await h.exported()).not.toContain('name="连接线"');
  });

  it("refuses commands that make no sense for the target", async () => {
    const group = await h.api.apply([{ op: "setText", target: { slide: 1, name: "数据组" }, text: "x" }]);
    expect(group.error!.code).toBe("NOT_APPLICABLE");
    const line = await h.api.apply([{ op: "setText", target: { slide: 1, name: "连接线" }, text: "x" }]);
    expect(line.error!.code).toBe("NOT_APPLICABLE");
    const turn = await h.api.apply([{ op: "setRotation", target: { slide: 1, name: "数据组" }, degrees: 10 }]);
    expect(turn.error!.code).toBe("NOT_APPLICABLE");
  });
});

describe("new content", () => {
  it("addShape then refer to it by alias in the same batch", async () => {
    const result = await h.api.apply({
      label: "加一个色块",
      commands: [
        { op: "addShape", slide: 1, shape: "roundRect", rect: { x: 2, y: 10, w: 6, h: 2 }, fill: "#00B050", text: "新增", as: "badge" },
        { op: "setBounds", target: "$badge", x: 4 },
        { op: "setLine", target: "$badge", color: null }
      ]
    });
    expect(result.ok).toBe(true);
    const created = result.results[0].created!;
    expect(created).toMatchObject({ slide: 1, name: expect.stringContaining("圆角矩形") });
    const badge = await shape(created.name);
    expect(badge).toMatchObject({ type: "shape", shape: "roundRect", fill: "#00b050", text: "新增", line: null });
    expect(badge.rect).toEqual({ x: 4, y: 10, w: 6, h: 2 });
    const xml = await h.exported();
    expect(xml).toContain(`id="${created.id}"`);
    expect(xmlOf(xml, created.id).text).toBe("新增");
    expect(h.steps().at(-1)).toBe("加一个色块");
  });

  it("addText writes paragraphs with their own alignment and an overall style", async () => {
    const result = await h.api.apply([{
      op: "addText", slide: 1, rect: { x: 2, y: 14, w: 12, h: 3 }, as: "note",
      paragraphs: [{ text: "第一段", align: "center" }, { text: "第二段" }],
      style: { fontSizePt: 24, bold: true, color: "#C00000" }
    }]);
    expect(result.ok).toBe(true);
    const note = await shape(result.results[0].created!.name);
    expect(note.text).toBe("第一段\n第二段");
    expect(note.style).toMatchObject({ fontSizePt: 24, bold: true, color: "#c00000", align: "center" });
    const { node } = xmlOf(await h.exported(), result.results[0].created!.id);
    const aligns = Array.from(node.getElementsByTagName("a:p")).map(p => p.getElementsByTagName("a:pPr")[0]?.getAttribute("algn") ?? null);
    expect(aligns).toEqual(["ctr", null]);
    expect(node.getElementsByTagName("a:rPr")[0].getAttribute("sz")).toBe("2400");
  });

  it("addShape draws lines and arrows, which carry no text", async () => {
    const ok = await h.api.apply([{ op: "addShape", slide: 1, shape: "arrow", rect: { x: 2, y: 8, w: 8, h: 0 }, line: { color: "#C00000", widthPt: 2 } }]);
    expect(ok.ok).toBe(true);
    const bad = await h.api.apply([{ op: "addShape", slide: 1, shape: "line", rect: { x: 2, y: 8, w: 8, h: 0 }, text: "nope" }]);
    expect(bad.error!.code).toBe("NOT_APPLICABLE");
  });

  it("addImage places a picture, optionally at a given rectangle", async () => {
    const result = await h.api.apply([{
      op: "addImage", slide: 1, src: "data:image/png;base64,iVBORw0KGgo=", rect: { x: 20, y: 2, w: 4, h: 3 }, name: "标志", as: "logo"
    }]);
    expect(result.ok).toBe(true);
    const logo = await shape("标志");
    expect(logo.type).toBe("image");
    expect(logo.rect).toEqual({ x: 20, y: 2, w: 4, h: 3 });
    expect(h.edits.structure.media).toHaveLength(1);
  });

  it("duplicate and copyTo give the copy an alias and a new id", async () => {
    const result = await h.api.apply([
      { op: "duplicate", target: { slide: 1, name: "卡片" }, offset: 1, as: "copy" },
      { op: "setFill", target: "$copy", color: "#7030A0" }
    ]);
    expect(result.ok).toBe(true);
    const created = result.results[0].created!;
    expect(created.id).not.toBe("10");
    const outline = (await h.api.outline()).slides[0].shapes;
    expect(outline.find(s => s.id === created.id)!.fill).toBe("#7030a0");
    expect(outline.find(s => s.id === "10")!.fill).toBe("#4472c4");
    expect(outline.find(s => s.id === created.id)!.rect.x).toBeCloseTo(3.54, 1);
  });
});

describe("slides", () => {
  it("adds, duplicates, moves and deletes slides, and says when a page does not exist", async () => {
    expect((await h.api.apply([{ op: "duplicateSlide", slide: 1 }])).ok).toBe(true);
    expect(h.doc.slides).toHaveLength(2);
    expect((await h.api.apply([{ op: "moveSlide", slide: 2, to: 1 }])).ok).toBe(true);
    expect((await h.api.apply([{ op: "deleteSlide", slide: 1 }])).ok).toBe(true);
    expect(h.doc.slides).toHaveLength(1);
    const last = await h.api.apply([{ op: "deleteSlide", slide: 1 }]);
    expect(last.error!.code).toBe("NOT_APPLICABLE");
    const missing = await h.api.apply([{ op: "duplicateSlide", slide: 7 }]);
    expect(missing.error!.code).toBe("SLIDE_NOT_FOUND");
  });

  it("rejects an unknown layout name and lists the real ones", async () => {
    const result = await h.api.apply([{ op: "addSlide", layout: "不存在" }]);
    expect(result.error!.message).toContain("标题和内容");
  });
});

describe("batches", () => {
  it("is one step in the history and one undo takes everything back", async () => {
    const before = await h.exported();
    await h.api.apply({ label: "一批修改", commands: [
      { op: "setText", target: { slide: 1, name: "卡片" }, text: "改过" },
      { op: "setFill", target: { slide: 1, name: "卡片" }, color: "#000000" },
      { op: "addShape", slide: 1, shape: "rect", rect: { x: 1, y: 1, w: 2, h: 2 } }
    ] });
    expect(h.steps()).toEqual(["一批修改"]);
    expect(await h.exported()).not.toBe(before);
    h.editor.undo();
    expect((await shape("卡片")).text).toBe("增长 30%");
    expect((await shape("卡片")).fill).toBe("#4472c4");
    expect(canon(await h.exported())).toBe(canon(before));
    h.editor.redo();
    expect((await shape("卡片")).text).toBe("改过");
  });

  it("atomic (the default): a failing command undoes the ones before it and leaves no step", async () => {
    const before = await h.exported();
    const result = await h.api.apply([
      { op: "setText", target: { slide: 1, name: "卡片" }, text: "不应保留" },
      { op: "addShape", slide: 1, shape: "rect", rect: { x: 1, y: 1, w: 2, h: 2 } },
      { op: "setFill", target: { slide: 1, name: "背景条" }, color: "#E60012" }
    ]);
    expect(result).toMatchObject({ ok: false, failedAt: 2, error: { code: "SHAPE_NOT_FOUND", path: "commands[2].target.name" } });
    expect(result.error!.candidates!.map(c => c.name)).toContain("卡片");
    expect((await shape("卡片")).text).toBe("增长 30%");
    expect(h.steps()).toEqual([]);
    expect(canon(await h.exported())).toBe(canon(before));
  });

  it("atomic: false skips the failing command and keeps the rest as one step", async () => {
    const result = await h.api.apply({ atomic: false, label: "尽量改", commands: [
      { op: "setFill", target: { slide: 1, name: "没有这个" }, color: "#E60012" },
      { op: "setFill", target: { slide: 1, name: "卡片" }, color: "#000000" }
    ] });
    expect(result.ok).toBe(false);
    expect(result.results.map(r => r.ok)).toEqual([false, true]);
    expect((await shape("卡片")).fill).toBe("#000000");
    expect(h.steps()).toEqual(["尽量改"]);
  });

  it("reports a malformed batch without touching the deck", async () => {
    const result = await h.api.apply([{ op: "setFill", target: { slide: 1, name: "卡片" }, color: "red" }]);
    expect(result).toMatchObject({ ok: false, error: { code: "INVALID_SCHEMA", path: "commands[0].color" } });
    expect(h.steps()).toEqual([]);
  });

  it("an ambiguous target is an error with every candidate, not a guess", async () => {
    const result = await h.api.apply([{ op: "setFill", target: { slide: 1, text: "收入" }, color: "#000000" }]);
    expect(result.error).toMatchObject({ code: "SHAPE_AMBIGUOUS" });
    expect(result.error!.candidates!.map(c => c.name)).toEqual(["指标甲", "指标乙"]);
  });
});
