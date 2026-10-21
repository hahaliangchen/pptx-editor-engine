// @vitest-environment node
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it } from "vitest";
import { HeadlessSession } from "../src/headless/session";
import { checkLayout, describeIssues } from "../src/api/layout-check";
import type { Outline, OutlineShape } from "../src/api/outline";

const sessions: HeadlessSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

async function deckWith(commands: unknown[]): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), { quiet: true });
  sessions.push(session);
  const result = await session.apply({ commands });
  expect(result.ok).toBe(true);
  return session;
}
const text = (rect: { x: number; y: number; w: number; h: number }, value: string, sizePt = 18) =>
  ({ op: "addText", slide: 1, rect, text: value, style: { fontSizePt: sizePt } });
const kinds = async (session: HeadlessSession) => checkLayout(await session.outline()).map(issue => issue.kind);

describe("layout check", () => {
  it("measures the real mixed-size cover by native paragraph spacing without false overflow", async () => {
    // Saved product project 820 v4: 37.5pt two-line title followed by an 18pt
    // subtitle, not three lines all set in the title's first-run size.
    const session = await HeadlessSession.open(readFileSync("tests/fixtures/unicom-mixed-text-metrics.pptx"), { quiet: true });
    sessions.push(session);
    const outline = await session.outline({ slides: [1] });
    const title = outline.slides[0].shapes.find(s => s.text?.startsWith("减少无效加班"))!;
    expect(title.textBox?.wrap).toBe("none");
    expect(title.textParagraphs?.map(p => p.runs[0].fontSizePt)).toEqual([37.5, 18]);
    expect(title.textParagraphs?.map(p => p.lineSpacing)).toEqual([{ unit: "points", value: 51 }, { unit: "points", value: 28.5 }]);
    expect(title.textParagraphs?.[1].spaceBeforePt).toBe(15);
    const original = structuredClone(outline);
    expect(checkLayout(outline).filter(i => i.kind === "text_overflow" && i.ids.includes(title.id))).toEqual([]);
    expect(outline).toEqual(original);
    // The same native content must still report actual height and no-wrap width
    // violations, rather than suppressing diagnostics for mixed-size text.
    title.rect.h = 2;
    expect(checkLayout(outline).some(i => i.kind === "text_overflow" && i.ids.includes(title.id))).toBe(true);
    title.rect.h = original.slides[0].shapes.find(s => s.id === title.id)!.rect.h;
    title.rect.w = 4;
    expect(checkLayout(outline).some(i => i.kind === "text_overflow" && i.ids.includes(title.id))).toBe(true);
  });

  it("uses each run's width and preserves explicit breaks, no-wrap and paragraph gaps", () => {
    const shape: OutlineShape = {
      id: "mixed", name: "Mixed sizes", type: "text", rotation: 0,
      rect: { x: 2, y: 3, w: 8, h: 2 }, text: "大后面是较小字号的说明文字", style: { fontSizePt: 28 },
      textBox: { margins: { left: 0, right: 0, top: 0, bottom: 0 }, autoFit: "none", wrap: "none" },
      textParagraphs: [{ runs: [{ text: "大", fontSizePt: 28, letterSpacingPt: 0 }, { text: "后面是较小字号的说明文字", fontSizePt: 10, letterSpacingPt: 0 }], spaceBeforePt: 0, spaceAfterPt: 0, marginLeftPt: 0, indentPt: 0 }]
    };
    const outline: Outline = { slideSize: { w: 33.87, h: 19.05 }, layouts: [], slides: [{ index: 1, id: "1", shapes: [shape] }] };
    expect(checkLayout(outline).filter(i => i.kind === "text_overflow")).toEqual([]);
    shape.textParagraphs![0].spaceBeforePt = 50;
    expect(checkLayout(outline).some(i => i.kind === "text_overflow")).toBe(true);
    shape.textParagraphs![0].spaceBeforePt = 0;
    shape.textParagraphs![0].runs[1].text = "\u2028第二行\n第三行";
    expect(checkLayout(outline).some(i => i.kind === "text_overflow")).toBe(true);
  });

  it("reports a light caption partly escaping its dark panel onto the white page", async () => {
    const session=await deckWith([
      {op:"addShape",slide:1,shape:"roundRect",rect:{x:3.5,y:12,w:19.3,h:1.6},fill:"#1A1A2E",line:null},
      {...text({x:4,y:12.45,w:20,h:1.5},"判定原则：没有数据先补记录；试点未改善，就回到机制假设而非扩大动作。",16),style:{fontSizePt:16,bold:true,color:"#FFFFFF"}}
    ]);
    const before=await session.outline();
    expect(checkLayout(before).some(i=>i.kind==="low_contrast"&&i.message.includes("面板外"))).toBe(true);
    expect(await session.outline()).toEqual(before);
    const caption=before.slides[0].shapes.find(s=>s.text?.startsWith("判定原则"))!;
    await session.apply([{op:"setText",target:{slide:1,id:caption.id},text:"先记录，再试点，最后复盘。"}]);
    expect(await kinds(session)).not.toContain("low_contrast");
  });
  it("reports wrapped height even when the widest original line only slightly exceeds the box", async () => {
    const session = await deckWith([text({x: 7.6, y: 5.1, w: 5.6, h: 2.2}, "02｜排期与资源挤压\n并行项目争用关键人", 17)]);
    expect(await kinds(session)).toContain("text_overflow");
  });
  it("detects dark-on-dark text against a painted background and later opaque covering panels without moving anything", async () => {
    const session = await deckWith([
      {op: "addShape", slide: 1, shape: "rect", rect: {x: 0, y: 0, w: 33.87, h: 19.05}, fill: "#1A1A2E", line: null},
      {...text({x: 3, y: 5, w: 12, h: 2}, "深色背景上的深色正文", 18), style: {fontSizePt: 18, color: "#1A1A2E"}},
      {...text({x: 18, y: 5, w: 12, h: 2}, "白字被后绘制面板遮盖", 18), style: {fontSizePt: 18, color: "#FFFFFF"}},
      {op: "addShape", slide: 1, shape: "rect", rect: {x: 17, y: 4, w: 14, h: 4}, fill: "#FFFFFF", line: null},
    ]);
    const before = await session.outline();
    const issues = checkLayout(before);
    expect(issues.some(issue => issue.kind === "low_contrast" && issue.message.includes("#1a1a2e"))).toBe(true);
    expect(issues.some(issue => issue.kind === "occluded_text")).toBe(true);
    expect(await session.outline()).toEqual(before);
  });
  it("flags a tall one-sided list and font-dependent icons without changing their layout", async () => {
    const session = await deckWith([
      text({x: 3, y: 4, w: 12, h: 1}, "需求与排期", 20),
      text({x: 3, y: 8, w: 12, h: 1}, "协同与返工", 20),
      text({x: 3, y: 12, w: 12, h: 1}, "管理改进", 20),
      text({x: 3, y: 14, w: 1, h: 1}, "💡", 20),
    ]);
    const before = await session.outline();
    expect(checkLayout(before).some(issue => issue.kind === "unbalanced" && issue.message.includes("一侧"))).toBe(true);
    expect(checkLayout(before).some(issue => issue.kind === "font_icon")).toBe(true);
    expect(await session.outline()).toEqual(before);
  });
  it("finds nothing wrong in a tidy page", async () => {
    const session = await deckWith([
      text({ x: 2, y: 2, w: 12, h: 2 }, "标题", 28),
      text({ x: 2, y: 5, w: 12, h: 4 }, "一段不长的说明文字", 18),
      text({ x: 16, y: 5, w: 12, h: 4 }, "另一段说明", 18)
    ]);
    expect(await kinds(session)).toEqual([]);
  });

  it("finds a text that does not fit its box, and says how much room it needs", async () => {
    const session = await deckWith([text({ x: 2, y: 2, w: 6, h: 1 }, "这是一段明显放不进这个小框里的很长很长的文字内容", 24)]);
    const [issue] = checkLayout(await session.outline());
    expect(issue.kind).toBe("text_overflow");
    expect(issue.slide).toBe(1);
    expect(issue.message).toContain("缩小字号");
  });

  it("finds content hugging the page edge, but not the header's own decoration", async () => {
    const cramped = await deckWith([text({ x: 0.3, y: 5, w: 12, h: 3 }, "贴着左边的内容", 20)]);
    expect(await kinds(cramped)).toContain("edge_margin");
    const header = await deckWith([text({ x: 0, y: 0, w: 4, h: 2 }, "01", 28)]);
    expect(await kinds(header)).not.toContain("edge_margin");
  });

  it("finds a paragraph set much larger than the other text on the page", async () => {
    const session = await deckWith([
      text({ x: 2, y: 4, w: 14, h: 6 }, "表格里的文字偏小，比如说明一二三四五六七八九十", 10),
      text({ x: 2, y: 11, w: 14, h: 3 }, "表格里的另一格文字也很小，比如说明一二三四五六", 10),
      text({ x: 18, y: 4, w: 14, h: 6 }, "旁边这一大段文字被设成了很大的字号，读起来比正文还抢眼，主次就颠倒了。", 26)
    ]);
    const issues = checkLayout(await session.outline());
    expect(issues.map(issue => issue.kind)).toContain("type_hierarchy");
    expect(describeIssues(issues)).toContain("主次颠倒");
  });

  it("finds content that sits at the top with a large gap below", async () => {
    const session = await deckWith([
      text({ x: 2, y: 1, w: 12, h: 2 }, "标题", 28),
      text({ x: 2, y: 3.2, w: 12, h: 3 }, "第一块内容文字，说明情况如下", 20),
      text({ x: 16, y: 3.2, w: 12, h: 3 }, "第二块内容文字，说明情况如下", 20),
      { op: "addChart", slide: 1, type: "column", categories: ["甲", "乙"], series: [{ name: "s", values: [1, 2] }], rect: { x: 2, y: 7.5, w: 26, h: 6 } }
    ]);
    expect(await kinds(session)).toContain("unbalanced");
  });

  it("knows where a vertically centred text sits, so a centred block is not taken for one at the top", async () => {
    const session = await deckWith([
      text({ x: 2, y: 1, w: 12, h: 2 }, "标题", 28),
      { op: "addText", slide: 1, rect: { x: 2, y: 4, w: 26, h: 12 }, text: "居中的一段说明，上下留白相等。", style: { fontSizePt: 22 }, as: "mid" },
      { op: "setAnchor", target: "$mid", anchor: "middle" },
      { op: "addChart", slide: 1, type: "column", categories: ["甲", "乙"], series: [{ name: "s", values: [1, 2] }], rect: { x: 2, y: 4, w: 10, h: 12 } }
    ]);
    const outline = await session.outline();
    const mid = (outline.slides[0].shapes as Array<{ text?: string; style?: { anchor?: string } }>).find(shape => shape.text?.startsWith("居中"))!;
    expect(mid.style!.anchor).toBe("middle");
    expect(checkLayout(outline).map(issue => issue.kind)).not.toContain("unbalanced");
  });

  it("finds content off the page", async () => {
    const session = await deckWith([text({ x: 30, y: 2, w: 10, h: 2 }, "出界", 18)]);
    expect(await kinds(session)).toContain("out_of_bounds");
  });

  it("finds two texts on top of each other, but not a text beside another", async () => {
    const stacked = await deckWith([text({ x: 2, y: 2, w: 10, h: 3 }, "第一段文字内容", 20), text({ x: 3, y: 2.5, w: 10, h: 3 }, "第二段文字内容", 20)]);
    expect(await kinds(stacked)).toContain("overlap");
    const beside = await deckWith([text({ x: 2, y: 2, w: 10, h: 3 }, "第一段文字内容", 20), text({ x: 13, y: 2, w: 10, h: 3 }, "第二段文字内容", 20)]);
    expect(await kinds(beside)).not.toContain("overlap");
  });

  it("finds a page whose content stops in its upper part", async () => {
    const session = await deckWith([
      text({ x: 2, y: 3, w: 10, h: 1.5 }, "要点一", 18), text({ x: 2, y: 5, w: 10, h: 1.5 }, "要点二", 18), text({ x: 2, y: 7, w: 10, h: 1.5 }, "要点三", 18)
    ]);
    const issues = checkLayout(await session.outline());
    expect(issues.map(issue => issue.kind)).toContain("empty_space");
    expect(describeIssues(issues)).toContain("第 1 页");
  });

  it("is a tool, which can be limited to some pages, and build_deck-style results can carry its issues", async () => {
    const session = await deckWith([text({ x: 30, y: 2, w: 10, h: 2 }, "出界", 18)]);
    const all = JSON.parse(((await session.callTool("check_layout", {})).content[0] as { text: string }).text);
    expect(all.issues.map((issue: { kind: string }) => issue.kind)).toContain("out_of_bounds");
    const none = JSON.parse(((await session.callTool("check_layout", { slides: [2] })).content[0] as { text: string }).text);
    expect(none.layout_passed).toBe(true);
    expect(none.scope).toBe("geometry_and_visibility");
    expect(none.not_checked).toContain("content completeness");
    expect((await session.callTool("check_layout", { slides: [0] })).ok).toBe(false);
  });
});
