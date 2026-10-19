// @vitest-environment node
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it } from "vitest";
import { HeadlessSession } from "../src/headless/session";
import { ApiError } from "../src/api/errors";
import { compileDeck, fitFontSize } from "../src/api/deck-spec";
import { describeTemplate, inferTemplate, slotName, tokensIn } from "../src/api/template";

const sessions: HeadlessSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

const text = (slide: number, x: number, y: number, w: number, h: number, value: string, sizePt = 20, name?: string) =>
  ({ op: "addText", slide, rect: { x, y, w, h }, text: value, style: { fontSizePt: sizePt, align: "left" }, ...(name ? { name } : {}) });

/** A template made of the deck's own tools: cover, contents (3 entries), section, content, end. */
async function template(): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), { quiet: true });
  sessions.push(session);
  const added = await session.apply({
    commands: [
      { op: "addSlide", after: 1 }, { op: "addSlide", after: 2 }, { op: "addSlide", after: 3 }, { op: "addSlide", after: 4 },
      // cover
      text(1, 2, 6, 8, 2, "{{TITLE}}", 40), text(1, 2, 9, 8, 1.2, "{{SUBTITLE}}"), text(1, 2, 12, 6, 1, "{{AUTHOR}}", 14),
      text(1, 2, 16, 12, 1, "示例品牌口号", 12),
      // contents: three entries, each a number, a title token and a description
      text(2, 2, 1, 6, 1, "目录", 28),
      ...[0, 1, 2].flatMap(i => [
        { op: "addShape", slide: 2, shape: "rect", rect: { x: 2, y: 4 + i * 3, w: 20, h: 2.4 }, fill: "#eef2f7" },
        text(2, 3, 4.2 + i * 3, 1.4, 1, `0${i + 1}`, 20),
        text(2, 5, 4.2 + i * 3, 8, 1, `{{TOC_ITEM_${i + 1}_TITLE}}`, 18),
        text(2, 5, 5.3 + i * 3, 8, 0.8, `第 ${i + 1} 项的示例说明`, 12)
      ]),
      // section
      text(3, 2, 6, 8, 2, "PART {{CHAPTER_NUM}}", 20), text(3, 2, 8, 12, 2, "{{CHAPTER_TITLE}}", 36), text(3, 2, 11, 12, 1, "{{CHAPTER_DESC}}", 16),
      // content: a title, a dashed frame with a label, the page number, a footer
      text(4, 2, 1, 12, 1.5, "{{PAGE_TITLE}}", 28),
      { op: "addShape", slide: 4, shape: "rect", rect: { x: 2, y: 4, w: 29.8, h: 12 }, fill: null, line: { color: "#999999", widthPt: 1, dash: "dash" } },
      text(4, 14, 9.5, 6, 1, "{{CONTENT_AREA}}", 18),
      text(4, 28, 17.8, 3, 0.8, "{{PAGE_NUM}} / {{TOTAL_PAGES}}", 10), text(4, 2, 17.8, 8, 0.8, "示例公司名称", 10),
      // end
      text(5, 2, 6, 8, 2, "{{THANK_YOU}}", 40), text(5, 2, 9, 8, 1, "{{CONTACT_INFO}}", 14)
    ]
  });
  expect(added.ok).toBe(true);
  return session;
}

const textsOf = async (session: HeadlessSession, slide: number): Promise<string[]> =>
  (await session.outline({ slides: [slide] })).slides[0].shapes
    .filter((shape: { type: string; text?: string }) => shape.type === "text" || shape.text)
    .map((shape: { text?: string }) => shape.text ?? "");

describe("reading a template", () => {
  it("finds the page kinds from their tokens and names what can be filled", async () => {
    const info = inferTemplate(await (await template()).outline());
    expect(Object.keys(info.pages)).toEqual(["cover", "toc", "section", "content", "end"]);
    expect(Object.keys(info.pages.cover!.slots)).toEqual(["title", "subtitle", "author"]);
    expect(Object.keys(info.pages.section!.slots)).toEqual(["title", "desc"]);
    expect(Object.keys(info.pages.end!.slots)).toEqual(["title", "contact"]);
    expect(info.pages.cover!.fixed.map(item => item.text)).toEqual(["示例品牌口号"]);
  });

  it("groups the contents entries with everything that belongs to each", async () => {
    const toc = inferTemplate(await (await template()).outline()).pages.toc!;
    expect(toc.entries).toHaveLength(3);
    // rectangle, number, title, description: four shapes each; the header is none of them
    expect(toc.entries!.map(entry => entry.shapes.length)).toEqual([4, 4, 4]);
    expect(toc.entries![0].desc).toBeDefined();
  });

  it("takes the content area from the frame around its marker", async () => {
    const info = inferTemplate(await (await template()).outline());
    expect(info.pages.content!.area).toMatchObject({ x: 2, y: 4, w: 29.8, h: 12 });
    expect(info.pages.content!.areaShapes).toHaveLength(2);
    expect(describeTemplate(info)).toMatchObject({ pages: { content: { fields: ["title"] }, toc: { entries: 3 } } });
  });

  it("names tokens", () => {
    expect(tokensIn("PART {{CHAPTER_NUM}} {{X}}")).toEqual(["CHAPTER_NUM", "X"]);
    expect(slotName("PAGE_TITLE")).toBe("title");
    expect(slotName("CLOSING_MESSAGE")).toBe("message");
    expect(slotName("STAT_1_LABEL")).toBe("stat1Label");
  });
});

const spec = {
  replace: { 示例公司名称: "新能源研究院" },
  slides: [
    { page: "cover", title: "2025 新能源汽车市场分析与未来趋势展望报告", subtitle: "规模与格局", replace: { 示例品牌口号: "" } },
    { page: "toc", items: ["市场规模", { title: "竞争格局", desc: "头部品牌" }] },
    { page: "section", title: "市场规模", desc: "销量与渗透率" },
    {
      page: "content", title: "核心指标", blocks: [
        { type: "cards", items: [{ value: "1290万", title: "销量", text: "同比 +35%" }, { value: "54%", title: "渗透率" }] },
        { type: "chart", kind: "column", categories: ["Q1", "Q2"], series: [{ name: "销量", values: [1, 2] }] }
      ]
    },
    { page: "end", title: "谢谢" }
  ]
};

describe("building a deck from a spec", () => {
  it("makes one page per entry, in order, and leaves no template page or token behind", async () => {
    const session = await template();
    const built = await session.apply(compileDeck(spec, inferTemplate(await session.outline())));
    expect(built.ok).toBe(true);
    const outline = await session.outline();
    expect(outline.slides).toHaveLength(5);
    const all = (await Promise.all([1, 2, 3, 4, 5].map(page => textsOf(session, page)))).flat();
    expect(all.some(line => line.includes("{{"))).toBe(false);
    expect(await textsOf(session, 3)).toEqual(expect.arrayContaining(["PART 01", "市场规模", "销量与渗透率"]));
    expect(await textsOf(session, 4)).toEqual(expect.arrayContaining(["核心指标", "4 / 5", "新能源研究院"]));
  });

  it("fills the contents and removes the entries that are not used, with their shapes", async () => {
    const session = await template();
    await session.apply(compileDeck(spec, inferTemplate(await session.outline())));
    const shapes = (await session.outline({ slides: [2] })).slides[0].shapes;
    const lines = shapes.map((shape: { text?: string }) => shape.text).filter(Boolean);
    expect(lines).toEqual(expect.arrayContaining(["市场规模", "竞争格局", "头部品牌"]));
    expect(lines).not.toContain("第 3 项的示例说明");
    // 1 header + 2 entries of (rectangle, number, title) plus the one description given
    // (the new slides also carry the layout's two empty placeholders, which are not content)
    expect(shapes.filter((shape: { type: string; text?: string }) => shape.type !== "text" || shape.text)).toHaveLength(1 + 2 * 3 + 1);
  });

  it("removes a sample text on request and shrinks a title that is too wide for its box", async () => {
    const session = await template();
    await session.apply(compileDeck(spec, inferTemplate(await session.outline())));
    const cover = (await session.outline({ slides: [1] })).slides[0].shapes as Array<{ text?: string; style?: { fontSizePt?: number }; rect: { x: number; w: number } }>;
    expect(cover.find(shape => shape.text === "示例品牌口号")).toBeUndefined();
    const title = cover.find(shape => shape.text?.startsWith("2025"))!;
    expect(title.style!.fontSizePt!).toBeLessThan(40);
    expect(title.style!.fontSizePt!).toBeGreaterThanOrEqual(28);
    // widened to the right, inside the page
    expect(title.rect.w).toBeGreaterThan(8);
    expect(title.rect.x + title.rect.w).toBeLessThanOrEqual(33.87 - 0.5 + 0.01);
  });

  it("lays blocks into the content area: inside it, side by side or stacked, charts native", async () => {
    const session = await template();
    await session.apply(compileDeck(spec, inferTemplate(await session.outline())));
    const shapes = (await session.outline({ slides: [4] })).slides[0].shapes as Array<{ type: string; rect: { x: number; y: number; w: number; h: number } }>;
    const chart = shapes.find(shape => shape.type === "chart")!;
    expect(chart).toBeDefined();
    for (const shape of shapes.filter(item => item.type === "chart" || item.type === "shape")) {
      expect(shape.rect.x).toBeGreaterThanOrEqual(2 - 0.01);
      expect(shape.rect.x + shape.rect.w).toBeLessThanOrEqual(2 + 29.8 + 0.01);
      expect(shape.rect.y).toBeGreaterThanOrEqual(4 - 0.01);
      expect(shape.rect.y + shape.rect.h).toBeLessThanOrEqual(4 + 12 + 0.01);
    }
    // cards and chart share the area: stacked, the chart is below the cards
    const card = shapes.find(shape => shape.type === "shape")!;
    expect(chart.rect.y).toBeGreaterThan(card.rect.y + card.rect.h - 0.01);
  });

  it("is atomic: a bad page builds nothing", async () => {
    const session = await template();
    const info = inferTemplate(await session.outline());
    const before = JSON.stringify(await session.outline());
    expect(() => compileDeck({ slides: [{ page: "cover", title: "x", colour: "red" }] }, info)).toThrow(/没有 "colour"/);
    expect(() => compileDeck({ slides: [{ page: "gallery" }] }, info)).toThrow(/没有 "gallery"/);
    expect(() => compileDeck({ slides: [{ page: "toc", items: ["1", "2", "3", "4"] }] }, info)).toThrow(/最多 3 项/);
    expect(() => compileDeck({ slides: [{ page: "content", title: "x", blocks: [{ type: "chart", kind: "pie", categories: ["a", "b"], series: [{ name: "s", values: [1] }] }] }] }, info))
      .toThrow(/一样多/);
    expect(() => compileDeck({ slides: [{ page: "cover", title: "x", blocks: [{ type: "text", text: "y" }] }] }, info)).toThrow(ApiError);
    expect(JSON.stringify(await session.outline())).toBe(before);
  });

  it("reports a bad spec with the path of the field", async () => {
    const info = inferTemplate(await (await template()).outline());
    try {
      compileDeck({ slides: [{ page: "cover", title: "x" }, { page: "cover", titel: "y" }] }, info);
      throw new Error("expected a failure");
    } catch (error) {
      expect(error).toBeInstanceOf(ApiError);
      expect((error as ApiError).path).toBe("slides[1].titel");
    }
  });

  it("is offered as tools: get_template and build_deck, and refuses on a deck that is no template", async () => {
    const session = await template();
    const described = await session.callTool("get_template", {});
    expect(JSON.parse((described.content[0] as { text: string }).text).pages.cover.fields).toContain("title");
    const built = await session.callTool("build_deck", spec);
    expect(built.ok).toBe(true);
    expect(JSON.parse((built.content[0] as { text: string }).text).slides).toBe(5);
    // the template pages are gone now
    const again = await session.callTool("build_deck", spec);
    expect(again.ok).toBe(false);
    const bad = await (await template()).callTool("build_deck", { slides: [{ page: "cover", nope: "x" }] });
    expect(bad.ok).toBe(false);
    expect((bad.content[0] as { text: string }).text).toContain("slides[0].nope");
  });
});

describe("fitting text", () => {
  it("picks a size that fits and never goes below the floor", () => {
    expect(fitFontSize(["短"], { w: 10, h: 5 }, 24)).toBe(24);
    const long = "很长很长的一段话".repeat(30);
    const size = fitFontSize([long], { w: 10, h: 4 }, 24, 12);
    expect(size).toBeLessThan(24);
    expect(fitFontSize([long.repeat(10)], { w: 10, h: 2 }, 24, 12)).toBe(12);
  });
});
