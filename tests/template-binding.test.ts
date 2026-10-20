// @vitest-environment node
import { readFileSync } from "node:fs";
import { createCanvas } from "@napi-rs/canvas";
import { afterEach, describe, expect, it } from "vitest";
import { HeadlessSession } from "../src/headless/session";
import { compileDeck } from "../src/api/deck-spec";
import { checkTemplate } from "../src/api/template-authoring";
import { describeTemplate, inferTemplate, parseBinding } from "../src/api/template";

const sessions: HeadlessSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

const text = (slide: number, x: number, y: number, w: number, h: number, value: string, sizePt = 20, name?: string) =>
  ({ op: "addText", slide, rect: { x, y, w, h }, text: value, style: { fontSizePt: sizePt, align: "left" }, ...(name ? { name } : {}) });
const badge = (slide: number, name: string, value: string) =>
  ({ op: "addShape", slide, shape: "roundRect", rect: { x: 29, y: 1, w: 3, h: 2 }, fill: "#c00000", text: value, name });
const frame = (slide: number, name: string) =>
  ({ op: "addShape", slide, shape: "rect", rect: { x: 2, y: 4, w: 29.8, h: 12 }, fill: null, line: { color: "#999999", widthPt: 1, dash: "dash" }, name });

/** A template whose shapes are bound by name and keep real sample text: no {{TOKEN}} anywhere. */
async function boundTemplate(extra: unknown[] = [], slides = 5): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), { quiet: true });
  sessions.push(session);
  const added = await session.apply({
    commands: [
      ...Array.from({ length: slides - 1 }, (_, i) => ({ op: "addSlide", after: i + 1 })),
      text(1, 2, 6, 14, 2, "联通产互年度汇报", 40, "@封面.标题"), text(1, 2, 9, 8, 1, "2026.10", 16, "@封面.日期"),
      text(1, 2, 12, 8, 1, "产品互联网部", 14, "@封面.作者"), text(1, 2, 16, 12, 1, "CHINA UNICOM", 12),
      text(2, 2, 1, 6, 1, "目录", 28),
      ...[0, 1, 2].flatMap(i => [
        text(2, 3, 4.2 + i * 3, 1.4, 1, `0${i + 1}`, 20),
        text(2, 5, 4.2 + i * 3, 12, 1, ["市场背景", "核心发现", "行动建议"][i], 18, `@目录.${i + 1}.标题`)
      ]),
      text(3, 2, 8, 14, 2, "第一章 市场", 36, "@章节.标题"), text(3, 2, 11, 14, 1, "规模与格局", 16, "@章节.说明"), badge(3, "@章节号", "1"),
      text(4, 4, 1, 12, 1.5, "核心指标一览", 28, "@内容.标题"), text(4, 4, 2.6, 12, 1, "一句话副标题", 14, "@内容.副标题"),
      frame(4, "@内容.区域"), text(4, 30, 17.8, 2, 0.8, "3", 10, "@页码"), badge(4, "@章节号", "1"),
      text(5, 2, 6, 14, 2, "谢谢观看", 40, "@结尾.标题"), text(5, 2, 9, 14, 1, "联系：张三 138-0000-0000", 14, "@结尾.联系方式"),
      ...extra
    ]
  });
  expect(added.ok).toBe(true);
  return session;
}

const textsOf = async (session: HeadlessSession, slide: number): Promise<string[]> =>
  (await session.outline({ slides: [slide] })).slides[0].shapes
    .filter((shape: { type: string; text?: string }) => shape.type === "text" || shape.text)
    .map((shape: { text?: string }) => shape.text ?? "");

describe("the name grammar", () => {
  it("reads roles, fields, automatic fields, the content frame, layouts and options", () => {
    expect(parseBinding("矩形 3")).toBeUndefined();
    expect(parseBinding("@")).toBeUndefined();
    expect(parseBinding("@封面.标题")).toEqual({ role: "cover", token: "TITLE", field: "title" });
    expect(parseBinding("@cover.title")).toEqual({ role: "cover", token: "TITLE", field: "title" });
    expect(parseBinding("＠内容.标题")).toEqual({ role: "content", token: "PAGE_TITLE", field: "title" });
    expect(parseBinding("@内容.副标题")).toMatchObject({ role: "content", field: "副标题", token: expect.stringMatching(/^U_/) });
    expect(parseBinding("@内容.区域")).toEqual({ role: "content", area: true });
    expect(parseBinding("@页码")).toEqual({ token: "PAGE_NUM" });
    expect(parseBinding("@章节号")).toEqual({ token: "CHAPTER_NUM" });
    expect(parseBinding("@目录.2.标题")).toEqual({ role: "toc", token: "TOC_ITEM_2_TITLE", field: "title" });
    expect(parseBinding("@目录.2.说明")).toEqual({ role: "toc", token: "TOC_ITEM_2_DESC", field: "desc" });
    expect(parseBinding("@结尾.联系方式")).toEqual({ role: "end", token: "CONTACT_INFO", field: "contact" });
    // a layout name after the role, and options after "|"
    expect(parseBinding("@内容:三圆图.标题|适用=三个并列要点")).toEqual({
      role: "content", variant: "三圆图", token: "PAGE_TITLE", field: "title", options: { 适用: "三个并列要点" }
    });
    expect(parseBinding("@内容:三圆图.区域")).toEqual({ role: "content", area: true, variant: "三圆图" });
  });
});

describe("a template bound by shape names", () => {
  it("reads the same pages and fields as a template of {{TOKEN}}s, with the sample text and how much fits", async () => {
    const info = inferTemplate(await (await boundTemplate()).outline());
    expect(Object.keys(info.pages)).toEqual(["cover", "toc", "section", "content", "end"]);
    expect(Object.keys(info.pages.cover!.slots)).toEqual(["title", "date", "author"]);
    expect(Object.keys(info.pages.content!.slots)).toEqual(["title", "副标题"]);
    expect(info.pages.toc!.entries).toHaveLength(3);
    expect(info.pages.content!.area).toMatchObject({ x: 2, w: 29.8 });
    expect(info.pages.content!.areaShapes).toHaveLength(1);
    const described = describeTemplate(info) as { pages: Record<string, { fieldInfo: Record<string, { sample?: string; fitChars: number; lines: number }> }> };
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

describe("several pages of one role (layouts)", () => {
  /** Two more content layouts: one named for cards, one for picture and text. */
  const layouts = [
    { op: "addSlide", after: 5 }, { op: "addSlide", after: 6 },
    text(6, 4, 1, 12, 1.5, "三个要点", 28, "@内容:三卡片.标题|适用=三个并列要点，每个一两句话"), frame(6, "@内容:三卡片.区域"),
    text(7, 4, 1, 12, 1.5, "图文说明", 28, "@内容:图文.标题|适用=一段说明配一张图"), text(7, 4, 3, 12, 1, "图注", 14, "@内容:图文.图注"), frame(7, "@内容:图文.区域")
  ];

  it("lists every named layout with when to use it, and builds the one a page asks for", async () => {
    const session = await boundTemplate(layouts, 7);
    const info = inferTemplate(await session.outline());
    expect(info.variants.map(page => [page.kind, page.variant])).toEqual([
      ["cover", undefined], ["toc", undefined], ["section", undefined], ["content", undefined], ["end", undefined],
      ["content", "三卡片"], ["content", "图文"]
    ]);
    const described = describeTemplate(info) as { pages: { content: { variants: Array<{ variant?: string; when?: string; fields: string[] }>; variantNote: string } } };
    expect(described.pages.content.variants.map(item => item.variant)).toEqual([undefined, "三卡片", "图文"]);
    expect(described.pages.content.variants[1].when).toBe("三个并列要点，每个一两句话");
    expect(described.pages.content.variants[2].fields).toEqual(["title", "图注"]);

    const built = await session.apply(compileDeck({
      slides: [
        { page: "content", variant: "图文", title: "产品外观", 图注: "正面视角", blocks: [{ type: "text", text: "说明" }] },
        { page: "content", title: "默认版式", blocks: [{ type: "text", text: "说明" }] },
        { page: "content", variant: "三卡片", title: "三点", blocks: [{ type: "text", text: "说明" }] }
      ]
    }, info));
    expect(built.ok).toBe(true);
    expect(await textsOf(session, 1)).toEqual(expect.arrayContaining(["产品外观", "正面视角"]));
    expect(await textsOf(session, 2)).toContain("默认版式");
    expect(await textsOf(session, 3)).toContain("三点");
  });

  it("refuses an unknown layout and lists the ones there are", async () => {
    const session = await boundTemplate(layouts, 7);
    const info = inferTemplate(await session.outline());
    expect(() => compileDeck({ slides: [{ page: "content", variant: "不存在", title: "x" }] }, info)).toThrow(/三卡片、图文/);
  });

  it("is accepted by the template check when each extra page names itself once, and refused when not", async () => {
    const good = checkTemplate(await (await boundTemplate(layouts, 7)).outline());
    expect(good.filter(issue => issue.kind === "duplicate_role")).toEqual([]);

    const unnamed = await boundTemplate([
      { op: "addSlide", after: 5 }, text(6, 4, 1, 12, 1.5, "又一页", 28, "@内容.标题"), frame(6, "@内容.区域")
    ], 6);
    expect(checkTemplate(await unnamed.outline()).map(issue => issue.kind)).toContain("duplicate_role");

    const twice = await boundTemplate([
      { op: "addSlide", after: 5 }, { op: "addSlide", after: 6 },
      text(6, 4, 1, 12, 1.5, "甲", 28, "@内容:同名.标题"), frame(6, "@内容:同名.区域"),
      text(7, 4, 1, 12, 1.5, "乙", 28, "@内容:同名.标题"), frame(7, "@内容:同名.区域")
    ], 7);
    expect(checkTemplate(await twice.outline()).filter(issue => issue.kind === "duplicate_role")).toHaveLength(1);
  });
});

describe("repeating units", () => {
  /** A content layout with three cards in a row; each is a background, a number, a title and a text. */
  const cards = (extra: Record<string, unknown>[] = []) => [
    { op: "addSlide", after: 5 },
    text(6, 4, 1, 12, 1.5, "三个要点", 28, "@内容:卡片.标题|适用=并列的几个要点"), frame(6, "@内容:卡片.区域"),
    ...[0, 1, 2].flatMap(i => {
      const x = 4 + i * 9.5;
      const name = (part?: string) => `@内容:卡片.卡片${i + 1}${part ? "." + part : ""}${i === 0 && !part ? "|数量=2-6" : ""}`;
      return [
        { op: "addShape", slide: 6, shape: "roundRect", rect: { x, y: 6, w: 8, h: 8 }, fill: "#f2f6f2", name: name() },
        text(6, x + 0.5, 6.4, 2, 1, `0${i + 1}`, 24, name("序号")),
        text(6, x + 0.5, 8, 7, 1.2, ["质量", "成本", "交付"][i], 22, name("标题")),
        text(6, x + 0.5, 9.5, 7, 3, ["做稳每一个细节", "把成本降下来", "按时交付结果"][i], 16, name("说明"))
      ];
    }),
    ...extra
  ];
  const setup = async (extra: Record<string, unknown>[] = []) => {
    const session = await boundTemplate(cards(extra), 6);
    return { session, info: inferTemplate(await session.outline()) };
  };
  const item = (title: string, desc: string) => ({ 标题: title, 说明: desc });

  it("finds the unit, its fields, its count range and the even spacing", async () => {
    const { info } = await setup();
    const unit = info.variants.find(page => page.variant === "卡片")!.units["卡片"];
    expect(unit.instances).toHaveLength(3);
    expect(unit.fields).toEqual(["标题", "说明"]);
    expect(unit.auto).toEqual(["序号"]);
    expect(unit.problems).toEqual([]);
    expect(unit.pitch!.dx).toBeCloseTo(9.5, 1);
    expect(unit.fixed).toBe(false);
    expect(unit.min).toBe(2);
    expect(unit.max).toBeGreaterThanOrEqual(5);
    const described = describeTemplate(info) as { pages: { content: { variants: Array<{ groups?: Record<string, { count: { min: number; max: number }; fieldInfo: Record<string, { sample?: string }> }> }> } } };
    const groups = described.pages.content.variants[1].groups!;
    expect(groups["卡片"].count.min).toBe(2);
    expect(groups["卡片"].fieldInfo["标题"].sample).toBe("质量");
  });

  it("builds the template's own number, fewer, and more (the instances shrink to fit)", async () => {
    for (const count of [3, 2, 5]) {
      const { session, info } = await setup();
      const entries = Array.from({ length: count }, (_, i) => item(`标题${i + 1}`, `说明文字${i + 1}`));
      const built = await session.apply(compileDeck({ slides: [{ page: "content", variant: "卡片", title: "要点", 卡片: entries }] }, info));
      expect(built.ok, `${count} cards`).toBe(true);
      const lines = await textsOf(session, 1);
      for (let i = 0; i < count; i++) expect(lines).toEqual(expect.arrayContaining([`标题${i + 1}`, `说明文字${i + 1}`, `0${i + 1}`]));
      expect(lines).not.toContain("质量");
      expect(lines.filter(line => /^标题\d$/.test(line))).toHaveLength(count);
      // every card stays on the page, inside the margins
      const shapes = (await session.outline({ slides: [1] })).slides[0].shapes as Array<{ rect: { x: number; w: number }; text?: string; style?: { fontSizePt?: number } }>;
      for (const shape of shapes) {
        expect(shape.rect.x).toBeGreaterThanOrEqual(-0.01);
        expect(shape.rect.x + shape.rect.w).toBeLessThanOrEqual(33.87 + 0.01);
      }
      if (count === 5) {
        const title = shapes.find(shape => shape.text === "标题1")!;
        expect(title.style!.fontSizePt!).toBeLessThan(22);
        // five cards side by side, evenly: the fifth ends where the third did
        const fifth = shapes.find(shape => shape.text === "标题5")!;
        expect(fifth.rect.x + fifth.rect.w).toBeLessThanOrEqual(4 + 2 * 9.5 + 8 + 0.2);
      }
    }
  });

  it("refuses a count outside the range, an unknown field, a missing unit; and drops a field an item leaves out", async () => {
    const { session, info } = await setup();
    const page = (cards: unknown) => ({ slides: [{ page: "content", variant: "卡片", title: "x", 卡片: cards }] });
    expect(() => compileDeck(page([item("a", "b")]), info)).toThrow(/2 到/);
    expect(() => compileDeck(page(Array.from({ length: 40 }, () => item("a", "b"))), info)).toThrow(/拆成两页/);
    expect(() => compileDeck(page([{ 标题: "a", 颜色: "红" }, item("b", "c")]), info)).toThrow(/标题、说明/);
    expect(() => compileDeck({ slides: [{ page: "content", variant: "卡片", title: "x" }] }, info)).toThrow(/需要“卡片”/);
    // an item without 说明 loses that text box
    const built = await session.apply(compileDeck(page([{ 标题: "只有标题" }, item("甲", "乙")]), info));
    expect(built.ok).toBe(true);
    const lines = await textsOf(session, 1);
    expect(lines).toContain("只有标题");
    expect(lines).not.toContain("做稳每一个细节");
  });

  it("reports how a unit is wrongly built", async () => {
    // the second card has no 说明, and the cards do not stand evenly
    const { info } = await setup([
      { op: "addSlide", after: 6 },
      text(7, 4, 1, 12, 1.5, "坏版式", 28, "@内容:坏.标题"), frame(7, "@内容:坏.区域"),
      text(7, 4, 6, 4, 1, "甲", 20, "@内容:坏.块1.标题"), text(7, 4, 7.2, 4, 1, "说明", 16, "@内容:坏.块1.说明"),
      text(7, 12, 9, 4, 1, "乙", 20, "@内容:坏.块2.标题"),
      text(7, 25, 6, 4, 1, "丙", 20, "@内容:坏.块3.标题"), text(7, 25, 7.2, 4, 1, "说明", 16, "@内容:坏.块3.说明")
    ]);
    const unit = info.variants.find(page => page.variant === "坏")!.units["块"];
    expect(unit.fixed).toBe(true);
    // the template check says so too
    const issues = checkTemplate(await (await boundTemplate(cards([
      { op: "addSlide", after: 6 },
      text(7, 4, 1, 12, 1.5, "坏版式", 28, "@内容:坏.标题"), frame(7, "@内容:坏.区域"),
      text(7, 4, 6, 4, 1, "甲", 20, "@内容:坏.块1.标题"), text(7, 4, 7.2, 4, 1, "说明", 16, "@内容:坏.块1.说明"),
      text(7, 12, 9, 4, 1, "乙", 20, "@内容:坏.块2.标题"),
      text(7, 25, 6, 4, 1, "丙", 20, "@内容:坏.块3.标题"), text(7, 25, 7.2, 4, 1, "说明", 16, "@内容:坏.块3.说明")
    ]), 7)).outline());
    expect(issues.filter(issue => issue.kind === "unit").map(issue => issue.message).join("\n")).toMatch(/不一样/);
    expect(unit.problems.join("\n")).toMatch(/字段和“块1”不一样/);
    expect(unit.problems.join("\n")).toMatch(/不规则/);
  });
});

describe("picture slots", () => {
  const png = (color: string, w = 120, h = 80) => {
    const canvas = createCanvas(w, h);
    const context = canvas.getContext("2d");
    context.fillStyle = color; context.fillRect(0, 0, w, h);
    return canvas.toDataURL("image/png");
  };
  /** A layout with one big picture slot, and a layout with a row of three pictures with captions. */
  const layouts = [
    { op: "addSlide", after: 5 }, { op: "addSlide", after: 6 },
    text(6, 4, 1, 12, 1.5, "图文", 28, "@内容:图文.标题"), text(6, 18, 6, 10, 1, "图注", 14, "@内容:图文.图注"),
    { op: "addImage", slide: 6, src: png("#888888"), rect: { x: 3, y: 4, w: 14, h: 10 }, name: "@内容:图文.图片" },
    text(7, 4, 1, 12, 1.5, "三张图", 28, "@内容:图墙.标题"),
    ...[0, 1, 2].flatMap(i => [
      { op: "addImage", slide: 7, src: png("#aaaaaa", 80, 80), rect: { x: 3 + i * 10, y: 4, w: 8, h: 8 }, name: `@内容:图墙.图${i + 1}.图片${i === 0 ? "|数量=2-4" : ""}` },
      text(7, 3 + i * 10, 12.5, 8, 1, "说明", 14, `@内容:图墙.图${i + 1}.说明`)
    ])
  ];

  it("lists them with their proportions and how to give a picture", async () => {
    const session = await boundTemplate(layouts, 7);
    const info = inferTemplate(await session.outline());
    const described = describeTemplate(info) as { pages: { content: { variants: Array<{ variant?: string; images?: Record<string, { aspect: number; note: string }>; groups?: Record<string, { fieldInfo: Record<string, { kind?: string; aspect?: number }> }> }> } } };
    const single = described.pages.content.variants.find(item => item.variant === "图文")!;
    expect(single.images!["图片"].aspect).toBeCloseTo(1.4, 1);
    expect(single.images!["图片"].note).toContain("placeholder");
    const wall = described.pages.content.variants.find(item => item.variant === "图墙")!;
    expect(wall.groups!["图"].fieldInfo["图片"]).toMatchObject({ kind: "image", aspect: 1 });
  });

  it("puts a given picture in the slot, cut to its proportions, and a stand-in for placeholder", async () => {
    const session = await boundTemplate(layouts, 7);
    const info = inferTemplate(await session.outline());
    const built = await session.apply(compileDeck({
      slides: [
        { page: "content", variant: "图文", title: "产品", 图注: "正面", 图片: png("#336699", 300, 100) },
        { page: "content", variant: "图墙", title: "三张", 图: [{ 说明: "甲", 图片: png("#cc3333") }, { 说明: "乙", 图片: "placeholder" }, { 说明: "丙", 图片: "占位" }] }
      ]
    }, info));
    expect(built.ok).toBe(true);
    const shapes = async (n: number) => (await session.outline({ slides: [n] })).slides[0].shapes as Array<{ type: string; name: string; rect: { x: number; y: number; w: number; h: number } }>;
    const first = (await shapes(1)).filter(shape => shape.type === "image");
    expect(first).toHaveLength(1);
    expect(first[0].rect).toMatchObject({ x: 3, w: 14 });
    expect(first[0].rect.h).toBeCloseTo(10, 0);
    // the sample picture is gone, not left behind the new one
    expect((await shapes(1)).filter(shape => shape.name === "@内容:图文.图片")).toHaveLength(0);
    expect((await shapes(2)).filter(shape => shape.type === "image")).toHaveLength(3);
  });

  it("accepts project pictures (the page fetches them) and refuses a missing picture or a wrong path", async () => {
    const session = await boundTemplate(layouts, 7);
    const info = inferTemplate(await session.outline());
    const compiled = compileDeck({ slides: [{ page: "content", variant: "图文", title: "产品", 图注: "正面", 图片: "images/ab12_logo.png" }] }, info);
    expect(JSON.stringify(compiled.commands)).toContain("project:images/ab12_logo.png");
    expect(() => compileDeck({ slides: [{ page: "content", variant: "图文", title: "产品", 图注: "正面" }] }, info)).toThrow(/需要图片“图片”/);
    expect(() => compileDeck({ slides: [{ page: "content", variant: "图文", title: "产品", 图注: "正面", 图片: "C:/a.png" }] }, info)).toThrow(/项目图片路径/);
    expect(() => compileDeck({ slides: [{ page: "content", variant: "图墙", title: "x", 图: [{ 说明: "甲" }, { 说明: "乙", 图片: "placeholder" }] }] }, info)).toThrow(/缺少图片/);
  });
});
