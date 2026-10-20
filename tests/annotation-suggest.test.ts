import { describe, expect, it } from "vitest";
import type { Outline, OutlineShape, OutlineSlide } from "../src/api/outline";
import { looksLikeSample, suggestAnnotations, suggestRoles } from "../src/api/annotation-suggest";

const shape = (id: string, name: string, x: number, y: number, w: number, h: number, extra: Partial<OutlineShape> = {}): OutlineShape =>
  ({ id, name, type: extra.text ? "text" : "shape", rotation: 0, rect: { x, y, w, h }, ...extra });
const text = (id: string, x: number, y: number, w: number, h: number, value: string, size = 18) =>
  shape(id, `文本 ${id}`, x, y, w, h, { text: value, style: { fontSizePt: size } });

/** A page of three cards (a round picture, a title and a text each), and a corner mark that every page has. */
function deck(): Outline {
  const corner = (index: number) => shape(`k${index}`, "角标", 30, 1, 3, 1);
  const cards = [0, 1, 2].flatMap(i => [
    shape(`p${i}`, `椭圆 ${i}`, 4 + i * 10, 4, 6, 6, { shape: "ellipse", fill: { other: "picture" } }),
    text(`t${i}`, 4 + i * 10, 10.5, 6, 1, "单击此处输入标题", 20),
    text(`d${i}`, 4 + i * 10, 11.8, 6, 3, "If they throw stones at you, do not throw back", 14)
  ]);
  const page = (index: number, shapes: OutlineShape[]): OutlineSlide => ({ index, id: `s${index}`, shapes: [corner(index), ...shapes] });
  return {
    slideSize: { w: 33.9, h: 19.1 }, layouts: [],
    slides: [
      page(1, [text("a", 2, 6, 20, 2, "年度报告", 48)]),
      page(2, [text("h", 2, 1, 6, 1, "目录", 28), text("n1", 4, 4, 2, 1, "01"), text("n2", 4, 7, 2, 1, "02"), text("n3", 4, 10, 2, 1, "03")]),
      page(3, [text("n", 12, 5, 6, 3, "01", 66), text("c", 8, 9, 14, 2, "工作完成情况", 36)]),
      page(4, [text("ti", 12, 1, 8, 1.2, "单击此处输入标题文本", 24), ...cards]),
      page(5, [text("e", 8, 7, 18, 3, "THANK YOU", 66)])
    ]
  };
}

describe("suggesting roles", () => {
  it("tells the cover, contents, section, content and end pages apart", () => {
    expect(suggestRoles(deck()).map(page => page.role)).toEqual(["cover", "toc", "section", "content", "end"]);
  });
});

describe("suggesting annotations for a page", () => {
  it("finds the unit of three cards, names each part, and leaves the repeated corner mark alone", () => {
    const result = suggestAnnotations(deck(), 4, { variant: "三图文", unit: "卡片" });
    expect(result.decoration).toContain("k4");
    expect(result.units).toHaveLength(1);
    expect(result.units[0]).toMatchObject({ name: "卡片", count: 3, axis: "row" });
    const names = Object.fromEntries(result.commands.map(command => [command.target.id, command.name]));
    expect(names.p0).toBe("@内容:三图文.卡片1.图片");
    expect(names.t1).toBe("@内容:三图文.卡片2.标题");
    expect(names.d2).toBe("@内容:三图文.卡片3.说明");
    // the page title is not part of the unit
    expect(names.ti).toBe("@内容:三图文.标题");
    expect(names.k4).toBeUndefined();
  });

  it("recognises sample wording", () => {
    expect(looksLikeSample("单击此处输入标题")).toBe(true);
    expect(looksLikeSample("If they throw stones at you, do not throw back")).toBe(true);
    expect(looksLikeSample("01")).toBe(true);
    expect(looksLikeSample("2025 年销量同比增长 35%")).toBe(false);
  });

  it("says so when there is no repeated structure, and refuses a page that is not there", () => {
    const result = suggestAnnotations(deck(), 1);
    expect(result.units).toEqual([]);
    expect(result.notes.join("")).toContain("没有找到等距重复的结构");
    expect(result.commands.map(command => command.name)).toEqual(["@封面.标题"]);
    expect(() => suggestAnnotations(deck(), 9)).toThrow(/没有第 9 页/);
  });
});
