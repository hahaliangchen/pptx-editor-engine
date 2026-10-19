// @vitest-environment node
import { existsSync, readFileSync } from "node:fs";
import { afterEach, describe, expect, it } from "vitest";
import { HeadlessSession } from "../src/headless/session";
import { checkTemplate, copySlideAsRole, slideRoles, suggestPlaceholders, templateDesignFacts } from "../src/api/template-authoring";
import { inferTemplate } from "../src/api/template";

const sessions: HeadlessSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

const text = (slide: number, x: number, y: number, w: number, h: number, value: string, sizePt = 20) =>
  ({ op: "addText", slide, rect: { x, y, w, h }, text: value, style: { fontSizePt: sizePt, align: "left" } });

const open = async (file: string | Buffer) => {
  const session = await HeadlessSession.open(typeof file === "string" ? readFileSync(file) : file, { quiet: true });
  sessions.push(session);
  return session;
};
const run = async (session: HeadlessSession, commands: unknown[]) => {
  const result = await session.apply({ commands, label: "t" });
  expect(result.ok, JSON.stringify(result.error)).toBe(true);
};

/** What an uploaded deck looks like: content pages only, each with a title, text, and a table. */
async function plainDeck(): Promise<HeadlessSession> {
  const session = await open("tools/templates/blank.pptx");
  await run(session, [
    { op: "addSlide", after: 1 },
    ...[1, 2].flatMap(n => [
      { op: "addShape", slide: n, shape: "rect", rect: { x: 0, y: 0, w: 33.87, h: 1.2 }, fill: "#1f4e79" },
      text(n, 2, 1.6, 20, 1.5, `第${n}页的标题`, 28),
      text(n, 2, 4, 14, 8, "这里是一大段示例正文，讲了很多内容。", 16),
      { op: "addTable", slide: n, rows: [["a", "b"], ["1", "2"]], rect: { x: 18, y: 4, w: 12, h: 4 } },
      text(n, 2, 17.5, 6, 0.8, "示例公司 2026", 10),
      text(n, 30, 17.5, 2, 0.8, `${n}`, 10)
    ])
  ]);
  return session;
}

describe("checkTemplate", () => {
  it("reports every missing role of a plain deck and offers no role to its pages", async () => {
    const issues = checkTemplate(await (await plainDeck()).outline());
    expect(issues.filter(i => i.kind === "missing_role").map(i => i.message)).toHaveLength(4);
    expect(issues.filter(i => i.kind === "unassigned").map(i => i.slide)).toEqual([1, 2]);
  });

  it("finds a keyword in a painted shape, two fields in one box, a bad spelling, and an unnumbered contents page", async () => {
    const session = await open("tools/templates/blank.pptx");
    await run(session, [
      { op: "addSlide", after: 1 },
      { op: "addShape", slide: 1, shape: "roundRect", rect: { x: 1, y: 1, w: 8, h: 2 }, fill: "#eeeeee", text: "{{TITLE}}" },
      text(1, 1, 5, 8, 1, "{{SUBTITLE}}{{AUTHOR}}", 14),
      text(1, 1, 7, 8, 1, "{{ title }}", 14),
      text(1, 1, 9, 8, 1, "{{COVER_TITLE}}", 14),
      text(2, 1, 1, 8, 1, "{{TOC_ITEM_2_TITLE}}", 14),
      text(2, 1, 3, 8, 1, "{{TOC_ITEM_3_TITLE}}", 14)
    ]);
    const kinds = checkTemplate(await session.outline()).map(i => i.kind);
    for (const kind of ["token_in_shape", "multiple_fields", "bad_token", "renamed_token", "toc_numbering", "toc_header"]) {
      expect(kinds).toContain(kind);
    }
  });
});

/** The user's path: content page in place, the other roles as copies, unassigned pages deleted, pages in order. */
async function makeTemplate(session: HeadlessSession, contentPage: number, basePage: number) {
  await run(session, suggestPlaceholders(await session.outline(), contentPage, "content"));
  for (const role of ["end", "toc", "cover"] as const) {
    await run(session, copySlideAsRole(await session.outline(), basePage, role));
  }
  for (;;) {
    const unassigned = slideRoles(await session.outline()).find(entry => !entry.role);
    if (!unassigned) break;
    await run(session, [{ op: "deleteSlide", slide: unassigned.slide }]);
  }
  const order = ["cover", "toc", "section", "content", "end"];
  for (let target = 1; target <= (await session.outline()).slides.length; target++) {
    const roles = slideRoles(await session.outline());
    const want = order.filter(role => roles.some(entry => entry.role === role))[target - 1];
    const at = roles.find(entry => entry.role === want)!.slide;
    if (at !== target) await run(session, [{ op: "moveSlide", slide: at, to: target }]);
  }
}

const sample = {
  slides: [
    { page: "cover", title: "智能体方案", subtitle: "行业落地" },
    { page: "toc", items: ["背景", "方案", "计划", "预算", "风险"] },
    { page: "content", title: "背景", blocks: [{ type: "text", text: "说明一" }, { type: "bullets", items: ["要点一", "要点二"] }] },
    { page: "end", title: "谢谢" }
  ]
};

async function expectUsable(session: HeadlessSession) {
  const outline = await session.outline();
  const errors = checkTemplate(outline).filter(issue => issue.severity === "error");
  expect(errors, JSON.stringify(errors, null, 1)).toEqual([]);
  expect(slideRoles(outline).map(entry => entry.role)).toEqual(["cover", "toc", "content", "end"]);
  const built = await session.callTool("build_deck", sample);
  expect(built.ok, JSON.stringify(built.content)).toBe(true);
}

describe("making a template from a plain deck", () => {
  it("ends with a template that has no errors and that build_deck fills", async () => {
    const session = await plainDeck();
    await makeTemplate(session, 1, 2);
    await expectUsable(session);
  }, 60_000);

  it("suggests the content page's keywords: title, page number, a dashed area in place of the sample content", async () => {
    const session = await plainDeck();
    await run(session, suggestPlaceholders(await session.outline(), 1, "content"));
    const outline = await session.outline();
    const page = inferTemplate(outline).pages.content!;
    expect(page.slots.title).toBeDefined();
    expect(page.area!.w).toBeGreaterThan(20);
    const shapes = outline.slides[0].shapes;
    expect(shapes.some(shape => shape.type === "table")).toBe(false);
    expect(shapes.some(shape => shape.text?.includes("示例公司"))).toBe(false);
    expect(shapes.some(shape => shape.text?.includes("{{PAGE_NUM}}"))).toBe(true);
    expect(shapes.some(shape => shape.fill === "#1f4e79")).toBe(true);
  });

  it("does not copy the content-area guide onto a cover made from a content page", async () => {
    const session = await plainDeck();
    await run(session, suggestPlaceholders(await session.outline(), 1, "content"));
    await run(session, copySlideAsRole(await session.outline(), 1, "cover"));
    const outline = await session.outline();
    expect(outline.slides[0].shapes.some(shape => shape.name === "内容区")).toBe(true);
    expect(outline.slides[1].shapes.some(shape => shape.name === "内容区")).toBe(false);
    expect(inferTemplate(outline).pages.cover?.slide).toBe(2);
    expect(outline.slides[1].shapes.some(shape => shape.fill === "#1f4e79")).toBe(true);
  });

  it("states the design facts", async () => {
    const session = await plainDeck();
    await makeTemplate(session, 1, 2);
    const facts = templateDesignFacts(await session.outline());
    expect(facts.pages.map(page => page.role)).toEqual(["cover", "toc", "content", "end"]);
    expect(facts.contentArea!.w).toBeGreaterThan(20);
    expect(facts.tocItems).toBeGreaterThanOrEqual(1);
  }, 60_000);

  it("check_template is a tool", async () => {
    const session = await plainDeck();
    const result = JSON.parse(((await session.callTool("check_template", {})).content[0] as { text: string }).text);
    expect(result.ok).toBe(false);
    expect(result.issues.length).toBeGreaterThan(0);
  });
});

const REAL = "元景行业智能体方案-20260713(2) - 副本.pptx";
describe.skipIf(!existsSync(REAL))("the real deck of the design doc", () => {
  it("becomes a template that checkTemplate accepts and build_deck fills", async () => {
    const session = await open(REAL);
    expect((await session.outline()).slides).toHaveLength(8);
    await makeTemplate(session, 2, 3);
    await expectUsable(session);
    const built = JSON.parse(((await session.callTool("check_layout", {})).content[0] as { text: string }).text);
    console.log("layout issues:", JSON.stringify(built.issues.map((i: { slide: number; kind: string }) => `${i.slide}:${i.kind}`)));
  }, 240_000);
});
