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
