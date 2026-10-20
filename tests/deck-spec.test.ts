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
    // the frame is 2, 4, 29.8 × 12; content keeps 0.5 cm clear of the header above it and the page's footer below
    expect(info.pages.content!.area).toMatchObject({ x: 2, y: 4.5, w: 29.8, h: 11.5 });
    expect(info.pages.content!.areaShapes).toHaveLength(2);
    expect(describeTemplate(info)).toMatchObject({ pages: { content: { fields: ["title"] }, toc: { maxEntries: 3 } } });
  });

  it("keeps a rule under the contents heading out of the first entry, so fewer entries leave it in place", async () => {
    const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), { quiet: true });
    sessions.push(session);
    const added = await session.apply({
      commands: [
        text(1, 2, 1, 6, 1, "目录", 28),
        { op: "addShape", slide: 1, shape: "line", rect: { x: 2, y: 2.6, w: 8, h: 0 }, line: { color: "#cccccc", widthPt: 1 }, name: "分隔线" },
        ...[0, 1, 2].flatMap(i => [
          text(1, 3, 4 + i * 3, 1.4, 1, `0${i + 1}`, 20),
          text(1, 5, 4 + i * 3, 8, 1, `{{TOC_ITEM_${i + 1}_TITLE}}`, 18)
        ])
      ]
    });
    expect(added.ok).toBe(true);
    const info = inferTemplate(await session.outline());
    const rule = (await session.outline()).slides[0].shapes.find((shape: { name: string }) => shape.name === "分隔线")!;
    expect(info.pages.toc!.entries!.map(entry => entry.shapes.length)).toEqual([2, 2, 2]);
    expect(info.pages.toc!.entries!.flatMap(entry => entry.shapes)).not.toContain(rule.id);
    expect(describeTemplate(info)).toMatchObject({ pages: { toc: { maxEntries: 3, entriesNote: expect.stringContaining("1 到 3 条") } } });

    expect((await session.apply(compileDeck({ slides: [{ page: "toc", items: ["唯一一章"] }] }, info))).ok).toBe(true);
    const shapes = (await session.outline()).slides[0].shapes;
    expect(shapes.find((shape: { name: string }) => shape.name === "分隔线")!.rect.y).toBeCloseTo(2.6, 1);
    expect(shapes.filter((shape: { text?: string }) => shape.text === "唯一一章")).toHaveLength(1);
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

/** A template whose shapes are bound by name and keep real sample text: no {{TOKEN}} anywhere. */
async function boundTemplate(): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), { quiet: true });
  sessions.push(session);
  const badge = (slide: number, name: string, value: string) =>
    ({ op: "addShape", slide, shape: "roundRect", rect: { x: 29, y: 1, w: 3, h: 2 }, fill: "#c00000", text: value, name });
  const added = await session.apply({
    commands: [
      { op: "addSlide", after: 1 }, { op: "addSlide", after: 2 }, { op: "addSlide", after: 3 }, { op: "addSlide", after: 4 },
      text(1, 2, 6, 14, 2, "联通产互年度汇报", 40, "@封面.标题"), text(1, 2, 9, 8, 1, "2026.10", 16, "@封面.日期"),
      text(1, 2, 12, 8, 1, "产品互联网部", 14, "@封面.作者"), text(1, 2, 16, 12, 1, "CHINA UNICOM", 12),
      text(2, 2, 1, 6, 1, "目录", 28),
      ...[0, 1, 2].flatMap(i => [
        text(2, 3, 4.2 + i * 3, 1.4, 1, `0${i + 1}`, 20),
        text(2, 5, 4.2 + i * 3, 12, 1, ["市场背景", "核心发现", "行动建议"][i], 18, `@目录.${i + 1}.标题`)
      ]),
      text(3, 2, 8, 14, 2, "第一章 市场", 36, "@章节.标题"), text(3, 2, 11, 14, 1, "规模与格局", 16, "@章节.说明"), badge(3, "@章节号", "1"),
      text(4, 4, 1, 12, 1.5, "核心指标一览", 28, "@内容.标题"), text(4, 4, 2.6, 12, 1, "一句话副标题", 14, "@内容.副标题"),
      { op: "addShape", slide: 4, shape: "rect", rect: { x: 2, y: 4, w: 29.8, h: 12 }, fill: null, line: { color: "#999999", widthPt: 1, dash: "dash" }, name: "@内容.区域" },
      text(4, 30, 17.8, 2, 0.8, "3", 10, "@页码"), badge(4, "@章节号", "1"),
      text(5, 2, 6, 14, 2, "谢谢观看", 40, "@结尾.标题"), text(5, 2, 9, 14, 1, "联系：张三 138-0000-0000", 14, "@结尾.联系方式")
    ]
  });
  expect(added.ok).toBe(true);
  return session;
}

describe("a template bound by shape names", () => {
  it("reads the same pages and fields as a template of {{TOKEN}}s, with the sample text and how much fits", async () => {
    const info = inferTemplate(await (await boundTemplate()).outline());
    expect(Object.keys(info.pages)).toEqual(["cover", "toc", "section", "content", "end"]);
    expect(Object.keys(info.pages.cover!.slots)).toEqual(["title", "date", "author"]);
    expect(Object.keys(info.pages.content!.slots)).toEqual(["title", "副标题"]);
    expect(info.pages.toc!.entries).toHaveLength(3);
    expect(info.pages.content!.area).toMatchObject({ x: 2, w: 29.8 });
    expect(info.pages.content!.areaShapes).toHaveLength(1);
    const described = describeTemplate(info) as { pages: Record<string, { fields: string[]; fieldInfo: Record<string, { sample?: string; fitChars: number; lines: number }> }> };
    expect(described.pages.cover.fieldInfo.title).toMatchObject({ sample: "联通产互年度汇报", lines: 1 });
    expect(described.pages.cover.fieldInfo.title.fitChars).toBeGreaterThan(8);
    expect(described.pages.content.fieldInfo["副标题"].sample).toBe("一句话副标题");
    // the decoration that is not bound is left alone and listed as sample wording
    expect(info.pages.cover!.fixed.map(item => item.text)).toEqual(["CHINA UNICOM"]);
  });

  it("builds a deck: bound shapes take the content, unused ones go, numbers keep the sample's look", async () => {
    const session = await boundTemplate();
    const info = inferTemplate(await session.outline());
    const built = await session.apply(compileDeck({
      slides: [
        { page: "cover", title: "年度经营分析", date: "2026.12" },
        { page: "toc", items: ["市场背景", "核心发现"] },
        { page: "section", title: "第一章 市场", desc: "规模与格局" },
        { page: "content", title: "销量创新高", "副标题": "同比增长 35%", blocks: [{ type: "text", text: "说明" }] },
        { page: "end", title: "谢谢", contact: "李四" }
      ]
    }, info));
    expect(built.ok).toBe(true);
    const all = (await Promise.all([1, 2, 3, 4, 5].map(page => textsOf(session, page)))).flat();
    expect(all.some(line => line.includes("{{") || line.startsWith("@"))).toBe(false);
    expect(await textsOf(session, 1)).toEqual(expect.arrayContaining(["年度经营分析", "2026.12", "CHINA UNICOM"]));
    expect(await textsOf(session, 1)).not.toContain("产品互联网部");
    expect(await textsOf(session, 1)).not.toContain("联通产互年度汇报");
    // the sample "1" in the badge stays one digit; the page number follows the page
    expect(await textsOf(session, 3)).toContain("1");
    expect(await textsOf(session, 4)).toEqual(expect.arrayContaining(["销量创新高", "同比增长 35%", "说明", "4"]));
    expect(await textsOf(session, 5)).toEqual(expect.arrayContaining(["谢谢", "李四"]));
    // the frame that marked the content area is gone
    const content = (await session.outline({ slides: [4] })).slides[0].shapes as Array<{ name: string }>;
    expect(content.some(shape => shape.name === "@内容.区域")).toBe(false);
  });

  it("refuses a field the page does not have, naming the ones it has", async () => {
    const session = await boundTemplate();
    const info = inferTemplate(await session.outline());
    expect(() => compileDeck({ slides: [{ page: "cover", title: "x", 备注: "y" }] }, info)).toThrow(/title、date、author/);
  });
});

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

  it("keeps content clear of the page edges even when the template's frame runs edge to edge", async () => {
    const session = await template();
    const frame = inferTemplate(await session.outline()).pages.content!.areaShapes![0];
    expect((await session.apply({ commands: [{ op: "setBounds", target: { slide: 4, id: frame }, x: 0, w: 33.87 }] })).ok).toBe(true);
    const info = inferTemplate(await session.outline());
    expect(info.pages.content!.area!.x).toBeCloseTo(1.2, 2);
    expect(info.pages.content!.area!.x + info.pages.content!.area!.w).toBeCloseTo(33.87 - 1.2, 2);
  });

  it("numbers content pages 01, 02… when the deck has no section pages, instead of 00", async () => {
    const session = await template();
    await session.apply({ commands: [text(4, 28, 1, 3, 1, "{{CHAPTER_NUM}}", 20)] });
    const info = inferTemplate(await session.outline());
    const content = (title: string) => ({ page: "content", title, blocks: [{ type: "text", text: "说明" }] });
    expect((await session.apply(compileDeck({ slides: [content("甲"), content("乙")] }, info))).ok).toBe(true);
    expect(await textsOf(session, 1)).toContain("01");
    expect(await textsOf(session, 2)).toContain("02");
  });

  it("fills the room: cards stand in the middle of the area, large, and a lone list becomes numbered rows", async () => {
    const session = await template();
    const info = inferTemplate(await session.outline());
    const cards = { page: "content", title: "指标", blocks: [{ type: "cards", items: [{ value: "9.1h", title: "峰值", text: "出现在 Q3" }, { value: "-41%", title: "回落", text: "排期评审后" }, { value: "3 类", title: "诱因" }] }] };
    const list = { page: "content", title: "路线", blocks: [{ type: "bullets", items: ["统一优先级入口", "前置确认点", "关键节点前移", "每月复盘"] }] };
    expect((await session.apply(compileDeck({ slides: [cards, list] }, info))).ok).toBe(true);
    const shapes = (n: number) => session.outline({ slides: [n] }).then(o => o.slides[0].shapes as Array<{ type: string; text?: string; rect: { x: number; y: number; w: number; h: number }; style?: { fontSizePt?: number } }>);

    const first = await shapes(1);
    const boxes = first.filter(shape => shape.type === "shape");
    expect(boxes).toHaveLength(3);
    const middle = boxes[0].rect.y + boxes[0].rect.h / 2;
    expect(middle).toBeGreaterThan(4.5 + 11.5 * 0.35);
    expect(middle).toBeLessThan(4.5 + 11.5 * 0.65);
    expect(boxes[0].rect.h).toBeGreaterThan(4.5);
    const value = first.find(shape => shape.text === "9.1h")!;
    expect(value.style!.fontSizePt!).toBeGreaterThanOrEqual(26);

    const rows = await shapes(2);
    expect(rows.filter(shape => shape.text && /^[1-4]$/.test(shape.text))).toHaveLength(4);
    const point = rows.find(shape => shape.text === "统一优先级入口")!;
    expect(point.style!.fontSizePt!).toBeGreaterThanOrEqual(16);
  });

  it("styles a table: header, bands, widths that follow the content", async () => {
    const session = await template();
    const info = inferTemplate(await session.outline());
    const table = { type: "table", rows: [["项目", "说明"], ["甲", "一段比较长的说明文字，需要更宽的列才能放得下"], ["乙", "短"]] };
    expect((await session.apply(compileDeck({ slides: [{ page: "content", title: "表", blocks: [table] }] }, info))).ok).toBe(true);
    const found = (await session.outline({ slides: [1] })).slides[0].shapes.find((shape: { type: string }) => shape.type === "table") as { rect: { x: number; w: number } };
    expect(found).toBeDefined();
    expect(found.rect.x).toBeGreaterThanOrEqual(2 - 0.01);
    expect(found.rect.x + found.rect.w).toBeLessThanOrEqual(2 + 29.8 + 0.01);
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
