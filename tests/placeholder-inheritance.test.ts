import { describe, expect, it } from "vitest";
import { PX_PER_PT, box, para, parseTexts, ph, slidePh } from "./helpers/deck";

const lvl = (n: number, attrs: string, inner = "") => `<a:lvl${n}pPr ${attrs}>${inner}</a:lvl${n}pPr>`;
const defRPr = (attrs: string, inner = "") => `<a:defRPr ${attrs}>${inner}</a:defRPr>`;
const tx1 = '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>';
const bullet = (char: string) => `<a:buFont typeface="Arial"/><a:buChar char="${char}"/>`;

// The shape every PowerPoint master has: titleStyle / bodyStyle by level.
const txStyles = (title = "", body = "") =>
  `<p:titleStyle>${lvl(1, 'algn="l"', defRPr('sz="4400"', tx1) + "<a:buNone/>") + title}</p:titleStyle>`
  + `<p:bodyStyle>${lvl(1, 'marL="342900" indent="-342900" algn="l"', bullet("•") + defRPr('sz="2800"', tx1))
    + lvl(2, 'marL="742950" indent="-285750" algn="l"', bullet("–") + defRPr('sz="2400"', tx1)) + body}</p:bodyStyle>`;

const pt = (n: number) => n * PX_PER_PT;
/** Sizes are rounded to whole pixels by the parser. */
const expectPx = (actual: number, expected: number) => expect(Math.abs(actual - expected)).toBeLessThan(0.6);
const emu = (n: number) => (n * 1920) / 9144000;

describe("placeholder text inheritance", () => {
  it("title: size and alignment come from the master's titleStyle, and it is not bold unless told", async () => {
    const [title] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("title", null, "", box(0, 0, 9144000, 900000)),
      layoutShapes: ph("title", null),
      slideShapes: slidePh("title", null, para("Hello"))
    });
    const p = title.paragraphs![0];
    expectPx(p.runs[0].style.fontSize, pt(44));
    expect(p.runs[0].style.bold).toBe(false);
    expect(p.style.align).toBe("left");
    expect(p.bullet).toBeUndefined();
  });

  it("a bold titleStyle makes the title bold", async () => {
    const [title] = await parseTexts({
      txStyles: `<p:titleStyle>${lvl(1, 'algn="ctr"', defRPr('sz="3600" b="1"'))}</p:titleStyle>`,
      masterShapes: ph("title", null, "", box(0, 0, 9144000, 900000)),
      layoutShapes: ph("title", null),
      slideShapes: slidePh("title", null, para("Hello"))
    });
    expect(title.paragraphs![0].runs[0].style.bold).toBe(true);
    expect(title.paragraphs![0].style.align).toBe("center");
  });

  it("layout overrides master: size and alignment of the layout placeholder win", async () => {
    const [title] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("title", null, "", box(0, 0, 9144000, 900000)),
      layoutShapes: ph("title", null, lvl(1, 'algn="ctr"', defRPr('sz="3000"'))),
      slideShapes: slidePh("title", null, para("Hello"))
    });
    expectPx(title.paragraphs![0].runs[0].style.fontSize, pt(30));
    expect(title.paragraphs![0].style.align).toBe("center");
  });

  it("slide lstStyle and run properties win over the layout", async () => {
    const [title] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("title", null, "", box(0, 0, 9144000, 900000)),
      layoutShapes: ph("title", null, lvl(1, 'algn="ctr"', defRPr('sz="3000"'))),
      slideShapes: slidePh("title", null,
        para("Hello", '<a:pPr algn="r"/>', '<a:rPr sz="2000"/>') + para("There"),
        lvl(1, 'algn="l"', defRPr('sz="2600"')))
    });
    const [first, second] = title.paragraphs!;
    expectPx(first.runs[0].style.fontSize, pt(20)); // run
    expect(first.style.align).toBe("right"); // paragraph
    expectPx(second.runs[0].style.fontSize, pt(26)); // slide lstStyle
    expect(second.style.align).toBe("left");
  });

  it("ctrTitle inherits from the master's title placeholder", async () => {
    const [title] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("title", null, lvl(1, 'algn="ctr"', defRPr('sz="5000"')), box(0, 0, 9144000, 900000)),
      layoutShapes: ph("ctrTitle", null),
      slideShapes: slidePh("ctrTitle", null, para("Big"))
    });
    expectPx(title.paragraphs![0].runs[0].style.fontSize, pt(50));
    expect(title.paragraphs![0].style.align).toBe("center");
  });

  it("subTitle inherits from the master's body placeholder lstStyle", async () => {
    const [sub] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("body", 1, lvl(1, 'marL="0" indent="0" algn="ctr"', "<a:buNone/>" + defRPr('sz="3000"')), box(0, 900000, 9144000, 3000000)),
      layoutShapes: ph("subTitle", 1),
      slideShapes: slidePh("subTitle", 1, para("Sub"))
    });
    const p = sub.paragraphs![0];
    expectPx(p.runs[0].style.fontSize, pt(30));
    expect(p.style.align).toBe("center");
    expect(p.bullet).toBeUndefined();
    expect(p.style.marginLeft).toBeCloseTo(0);
  });

  it("subTitle: a layout buNone and indent reset beat the master's bullets", async () => {
    const [sub] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("body", 1, "", box(0, 900000, 9144000, 3000000)),
      layoutShapes: ph("subTitle", 1, lvl(1, 'marL="0" indent="0" algn="ctr"', "<a:buNone/>" + defRPr('sz="3200"'))),
      slideShapes: slidePh("subTitle", 1, para("Sub"))
    });
    const p = sub.paragraphs![0];
    expectPx(p.runs[0].style.fontSize, pt(32));
    expect(p.bullet).toBeUndefined();
    expect(p.style.marginLeft).toBeCloseTo(0);
    expect(p.style.indent).toBeCloseTo(0);
    expect(p.style.align).toBe("center");
  });

  it("body: level 1 and level 2 take their own size, bullet and hanging indent from bodyStyle", async () => {
    const [body] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("body", 1, "", box(0, 900000, 9144000, 3000000)),
      layoutShapes: ph(null, 1),
      slideShapes: slidePh(null, 1, para("One") + para("Two", '<a:pPr lvl="1"/>') + para("Three"))
    });
    const [one, two, three] = body.paragraphs!;
    expectPx(one.runs[0].style.fontSize, pt(28));
    expect(one.bullet?.char).toBe("•");
    expectPx(one.style.marginLeft, emu(342900));
    expectPx(one.style.indent, emu(-342900));
    expectPx(two.runs[0].style.fontSize, pt(24));
    expect(two.bullet?.char).toBe("–");
    expectPx(two.style.marginLeft, emu(742950));
    expect(three.bullet?.char).toBe("•");
  });

  it("layout-level bullet size/colour tweaks apply even when the bullet itself comes from the master", async () => {
    const [body] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("body", 1, "", box(0, 900000, 9144000, 3000000)),
      layoutShapes: ph(null, 1, lvl(1, "", '<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="50000"/>')),
      slideShapes: slidePh(null, 1, para("One"))
    });
    const bulletOf = body.paragraphs![0].bullet!;
    expect(bulletOf.char).toBe("•");
    expect(bulletOf.color.toLowerCase()).toBe("#ff0000");
    expectPx(bulletOf.fontSize, pt(28) * 0.5);
  });

  it("falls back to the presentation's default text size, not a made-up per-type size", async () => {
    const [sub] = await parseTexts({
      defaultTextStyle: lvl(1, 'algn="l"', defRPr('sz="2000"')),
      txStyles: "",
      masterShapes: ph("title", null, "", box(0, 0, 9144000, 900000)),
      layoutShapes: ph("title", null),
      slideShapes: slidePh("title", null, para("Plain"))
    });
    const style = sub.paragraphs![0].runs[0].style;
    expectPx(style.fontSize, pt(20));
    expect(style.bold).toBe(false);
  });

  it("numbered lists count up, restart after an interruption, and keep a separate count per level", async () => {
    const num = (scheme: string, extra = "") => `<a:buAutoNum type="${scheme}"${extra}/>`;
    const [body] = await parseTexts({
      txStyles: txStyles(),
      masterShapes: ph("body", 1, "", box(0, 900000, 9144000, 3000000)),
      layoutShapes: ph(null, 1),
      slideShapes: slidePh(null, 1, [
        para("a", `<a:pPr>${num("arabicPeriod")}</a:pPr>`),
        para("b", `<a:pPr>${num("arabicPeriod")}</a:pPr>`),
        para("b1", `<a:pPr lvl="1">${num("alphaLcParenR")}</a:pPr>`),
        para("b2", `<a:pPr lvl="1">${num("alphaLcParenR")}</a:pPr>`),
        para("c", `<a:pPr>${num("arabicPeriod")}</a:pPr>`),
        para("plain", "<a:pPr><a:buNone/></a:pPr>"),
        para("d", `<a:pPr>${num("arabicPeriod")}</a:pPr>`),
        para("e", `<a:pPr>${num("romanUcPeriod", ' startAt="4"')}</a:pPr>`)
      ].join(""))
    });
    expect(body.paragraphs!.map(p => p.bullet?.char ?? "")).toEqual(["1.", "2.", "a)", "b)", "3.", "", "1.", "IV."]);
  });
});

describe("formatAutoNumber", () => {
  it("covers the common numbering schemes", async () => {
    const { formatAutoNumber } = await import("../src/pptx-parser");
    expect(formatAutoNumber("arabicPeriod", 3)).toBe("3.");
    expect(formatAutoNumber("arabicParenR", 3)).toBe("3)");
    expect(formatAutoNumber("arabicParenBoth", 3)).toBe("(3)");
    expect(formatAutoNumber("arabicPlain", 3)).toBe("3");
    expect(formatAutoNumber("alphaLcPeriod", 28)).toBe("ab.");
    expect(formatAutoNumber("alphaUcParenR", 2)).toBe("B)");
    expect(formatAutoNumber("romanLcPeriod", 9)).toBe("ix.");
    expect(formatAutoNumber("romanUcPeriod", 14)).toBe("XIV.");
    expect(formatAutoNumber("circleNumDbPlain", 2)).toBe("②");
  });
});
