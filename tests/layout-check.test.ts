// @vitest-environment node
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it } from "vitest";
import { HeadlessSession } from "../src/headless/session";
import { checkLayout, describeIssues } from "../src/api/layout-check";

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
    expect(none.ok).toBe(true);
    expect((await session.callTool("check_layout", { slides: [0] })).ok).toBe(false);
  });
});
