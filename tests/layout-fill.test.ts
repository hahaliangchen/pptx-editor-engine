import { describe, expect, it } from "vitest";
import { describeLayout, fillLayout, layoutIssues } from "../src/design/layout-fill";

const P = (t: string, color = "#333333") => `<p data-paragraph="true" style="font-size:13px"><span style="font-size:13px;color:${color}">${t}</span></p>`;
const bullet = (t: string) => `<p data-paragraph="true" style="font-size:12px"><span style="color:#E60012;font-size:9px">■</span><span style="font-size:12px;color:#333333"> ${t}</span></p>`;
const card = (n: number, x: number, y: number, w = 300, h = 120) => `<div data-id="card${n}" data-kind="group" data-name="g${n}" data-purpose="content" data-slot="要点卡#1" style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px">
  <div data-id="bg${n}" data-kind="shape" data-name="bg${n}" data-purpose="content" style="position:absolute;left:0px;top:0px;width:${w}px;height:${h}px;background-image:linear-gradient(180deg, #ffffff 0%, #fdf3f3 100%);border-radius:6px;box-shadow:0px 1px 4px rgba(0, 0, 0, 0.09)"></div>
  <div data-id="ico${n}" data-kind="group" data-name="i${n}" data-purpose="reference" data-slot="icon" style="position:absolute;left:10px;top:37px;width:46px;height:46px">
    <div data-id="tile${n}" data-kind="shape" data-name="t${n}" data-purpose="reference" style="position:absolute;left:0px;top:0px;width:46px;height:46px;background-color:#FDEDEE;border-radius:10px"></div>
    <div data-id="ring${n}" data-kind="shape" data-name="r${n}" data-purpose="reference" style="position:absolute;left:12px;top:12px;width:22px;height:22px;border:2px dashed #E60012;border-radius:50%"></div>
  </div>
  <div data-id="txt${n}" data-kind="text" data-name="x${n}" data-purpose="content" data-slot="text" style="position:absolute;left:67px;top:20px;width:${w - 77}px;height:80px">${P("要点标题", "#E60012")}${bullet("说明文字占位")}</div>
</div>`;
const sep = (n: number, x: number) => `<div data-id="sep${n}" data-kind="text" data-name="s${n}" data-purpose="content" data-slot="sep#1" style="position:absolute;left:${x}px;top:150px;width:14px;height:26px">${P("»")}</div>`;
const page = (body: string) => `<!doctype html><html><head></head><body data-page-id="p1" data-page-fill="layout"><div data-id="title" data-kind="text" data-name="TextBox 8" data-purpose="content" style="position:absolute;left:50px;top:20px;width:700px;height:46px">${P("页面标题——副标题占位")}</div>${body}</body></html>`;
const row = page(card(1, 40, 100) + sep(1, 353) + card(2, 380, 100) + sep(2, 693) + card(3, 720, 100));
const doc = (html: string) => new DOMParser().parseFromString(html, "text/html");
const svg = `<svg viewBox="0 0 24 24"><path d="M4 4h16v16H4z" fill="none" stroke="#E60012" stroke-width="2"/></svg>`;

describe("layout fill", () => {
  it("summarises sets, fields, icons and loose texts compactly", () => {
    const summary = describeLayout(row);
    expect(summary.texts).toEqual([{ id: "title", text: ["页面标题——副标题占位"] }]);
    expect(summary.sets).toHaveLength(1);
    expect(summary.sets[0]).toMatchObject({ set: "要点卡#1", block: "要点卡", arrangement: "row", columns: 3 });
    expect(summary.sets[0].items[0]).toEqual({ fields: { text: ["要点标题", "■\u00a0说明文字占位"] }, icons: ["icon"], sets: [] });
  });

  it("drops a block, re-flows the row, keeps separators between items and fills text and icons", () => {
    const { html, notes } = fillLayout(row, { texts: { title: "1、建设背景——一期成果" }, sets: [{ set: "要点卡#1", items: [
      { fields: { text: ["知识检索", "已上线", "支持全文检索"] }, icons: { icon: svg } },
      { fields: { text: ["智能问答"] } }
    ] }] });
    const d = doc(html);
    expect(notes).toEqual([]);
    expect(d.querySelector('[data-id="title"]')!.textContent).toBe("1、建设背景——一期成果");
    const cards = Array.from(d.querySelectorAll<HTMLElement>('[data-slot="要点卡#1"]'));
    expect(cards.map(c => c.dataset.id)).toEqual(["card1", "card2"]);
    // Two cards share the old three-card region: 40..1020 with the original 40px gap.
    expect(cards.map(c => [c.style.left, c.style.width])).toEqual([["40px", "470px"], ["550px", "470px"]]);
    // Wide children stretch with the card, the icon keeps its place, paint is untouched.
    expect(d.querySelector<HTMLElement>('[data-id="bg1"]')!.style.width).toBe("470px");
    expect(d.querySelector<HTMLElement>('[data-id="bg1"]')!.style.boxShadow).toBe("0px 1px 4px rgba(0, 0, 0, 0.09)");
    expect(d.querySelector<HTMLElement>('[data-id="ico1"]')!.style.left).toBe("10px");
    expect(d.querySelector<HTMLElement>('[data-id="txt1"]')!.style.width).toBe("393px");
    const seps = Array.from(d.querySelectorAll<HTMLElement>('[data-slot="sep#1"]'));
    expect(seps.map(s => s.dataset.id)).toEqual(["sep1"]);
    expect(seps[0].style.left).toBe("523px");
    // Paragraph styles follow the template: first like the title, the rest like the bullet, marker kept.
    const paragraphs = Array.from(d.querySelector('[data-id="txt1"]')!.querySelectorAll("p")).map(p => p.textContent);
    expect(paragraphs).toEqual(["知识检索", "■ 已上线", "■ 支持全文检索"]);
    expect(d.querySelector('[data-id="txt1"] p span')!.getAttribute("style")).toContain("#E60012");
    // The ring becomes the icon at the ring's box; the tile stays.
    expect(d.querySelector('[data-id="ring1"]')).toBeNull();
    const art = d.querySelector<HTMLElement>('[data-id="ico1"] svg')!;
    expect(art.dataset.kind).toBe("image");
    expect([art.style.left, art.style.width]).toEqual(["12px", "22px"]);
    expect(d.querySelector('[data-id="tile1"]')).not.toBeNull();
    expect(d.querySelector('[data-id="card3"]')).toBeNull();
  });

  it("takes an icon slot marked on a single dashed shape as its own ring", () => {
    const single = row.replace(/<div data-id="ico1"[\s\S]*?<\/div>\s*<\/div>/, '<div data-id="ico1" data-kind="shape" data-name="i1" data-purpose="content" data-slot="icon" style="position:absolute;left:10px;top:37px;width:46px;height:46px;border:2px dashed #E60012"></div>');
    const icon = doc(fillLayout(single, { sets: [{ set: "要点卡#1", items: [{ icons: { icon: svg } }] }] }).html).querySelector('[data-id="card1"] svg')!;
    expect([icon.getAttribute("style"), icon.querySelector("path")!.getAttribute("stroke")]).toEqual(["position:absolute;left:10px;top:37px;width:46px;height:46px", "#E60012"]);
  });

  it("reports what keeps a layout page from being filled", () => {
    expect(layoutIssues(row)).toEqual([]);
    const messages = (html: string) => layoutIssues(html).map(i => i.message).join("\n");
    expect(layoutIssues(page(""))).toMatchObject([{ severity: "info", message: expect.stringMatching(/没有块/) }]);
    expect(messages(row.replace('data-slot="text"', ""))).toMatch(/没有字段名/);
    expect(messages(row.replace('data-id="bg1" data-kind="shape" data-name="bg1" data-purpose="content"', '$& data-slot="卡#2"'))).toMatch(/块必须是组合/);
    expect(messages(row.replace('data-slot="sep#1" style="position:absolute;left:353px', 'data-slot="sep#3" style="position:absolute;left:353px'))).toMatch(/没有组号为 3 的块/);
  });

  it("paints an icon in its placeholder ring's color, white on a red tile", () => {
    const white = row.replace(/border:2px dashed #E60012/g, "border:2px dashed #FFFFFF");
    const art = `<svg viewBox="0 0 24 24" stroke="#FFFFFF"><circle cx="12" cy="12" r="8" fill="#FBD3D7" fill-opacity=".4"/><path d="M8 12h8" style="stroke:#333333;fill:none"/><rect x="2" y="2" width="4" height="4"/></svg>`;
    const fill = (html: string) => doc(fillLayout(html, { sets: [{ set: "要点卡#1", items: [{ icons: { icon: art } }] }] }).html).querySelector('[data-id="ico1"] svg')!;
    const red = fill(row);
    expect(red.getAttribute("stroke")).toBe("#E60012");
    expect(red.getAttribute("fill")).toBe("#E60012");
    expect(red.querySelector("circle")!.getAttribute("fill")).toBe("#E60012");
    expect(red.querySelector("circle")!.getAttribute("fill-opacity")).toBe(".4");
    expect(red.querySelector<SVGElement>("path")!.style.stroke).toMatch(/#E60012|rgb\(230, 0, 18\)/i);
    expect(red.querySelector<SVGElement>("path")!.style.fill).toBe("none");
    expect(fill(white).getAttribute("stroke")).toBe("#FFFFFF");
  });

  it("copies the last block for extra items as new objects and lays a grid out by columns", () => {
    const grid = page(card(1, 40, 100, 200, 100) + card(2, 260, 100, 200, 100) + card(3, 480, 100, 200, 100) + card(4, 40, 220, 200, 100) + card(5, 260, 220, 200, 100) + card(6, 480, 220, 200, 100));
    expect(describeLayout(grid).sets[0]).toMatchObject({ arrangement: "grid", columns: 3 });
    const four = doc(fillLayout(grid, { sets: [{ set: "要点卡", items: [1, 2, 3, 4].map(i => ({ fields: { text: `能力${i}` } })) }] }).html);
    const kept = Array.from(four.querySelectorAll<HTMLElement>('[data-slot="要点卡#1"]'));
    expect(kept).toHaveLength(4);
    expect(kept.map(c => [c.style.left, c.style.top])).toEqual([["40px", "100px"], ["370px", "100px"], ["40px", "220px"], ["370px", "220px"]]);
    const more = doc(fillLayout(row, { sets: [{ set: "要点卡#1", items: [1, 2, 3, 4].map(i => ({ fields: { text: `能力${i}` } })) }] }).html);
    expect(more.querySelectorAll('[data-slot="要点卡#1"]')).toHaveLength(3);
    const copy = Array.from(more.body.children).find(c => c.getAttribute("data-kind") === "group" && !c.hasAttribute("data-id")) as HTMLElement;
    expect(copy.querySelector("p")!.textContent).toBe("能力4");
    expect(copy.querySelector("[data-id],[data-slot],[data-purpose],[data-note]")).toBeNull();
    expect(copy.innerHTML).toContain("linear-gradient(180deg, #ffffff 0%, #fdf3f3 100%)");
    expect(more.querySelectorAll('[data-kind="text"]')).toHaveLength(1 + 4 + 3);
  });

  it("keeps icons in proportion and shrinks them only when a block gets too small", () => {
    const column = page(card(1, 40, 100, 300, 120) + card(2, 40, 230, 300, 120));
    const wide = doc(fillLayout(column, { sets: [{ set: "要点卡#1", items: [{ fields: { text: "一" } }] }] }).html);
    expect(wide.querySelector<HTMLElement>('[data-id="card1"]')!.style.height).toBe("250px");
    const icon = wide.querySelector<HTMLElement>('[data-id="ico1"]')!;
    expect([icon.style.width, icon.style.height, icon.style.top]).toEqual(["46px", "46px", "102px"]);
    // Four cards are 55px tall: the centred 46px icon gives up its margins but keeps four fifths of the card
    // (it used to shrink to a dot whenever the block lost more height than the icon's margins).
    const four = doc(fillLayout(column, { sets: [{ set: "要点卡#1", items: [1, 2, 3, 4].map(i => ({ fields: { text: `${i}` } })) }] }).html);
    expect(four.querySelector<HTMLElement>('[data-id="ico1"]')!.style.width).toBe("44px");
    const tight = doc(fillLayout(column, { sets: [{ set: "要点卡#1", items: [1, 2, 3, 4, 5, 6, 7, 8].map(i => ({ fields: { text: `${i}` } })) }] }).html);
    const small = tight.querySelector<HTMLElement>('[data-id="ico1"]')!;
    expect(small.style.width).toBe(small.style.height);
    expect(parseFloat(small.style.width)).toBeLessThan(46);
    expect(parseFloat(tight.querySelector<HTMLElement>('[data-id="ring1"]')!.style.width)).toBeLessThan(22);
  });

  it("does not repeat a list marker the content already has, and widens one-line labels with their frame", () => {
    const marked = doc(fillLayout(row, { sets: [{ set: "要点卡#1", items: [{ fields: { text: ["标题", "■ 已有标记"] } }] }] }).html);
    expect(Array.from(marked.querySelectorAll('[data-id="txt1"] p')).map(p => p.textContent)).toEqual(["标题", "■\u00a0已有标记"]);
    const ribbon = page(`<div data-id="rib" data-kind="group" data-name="r" data-purpose="content" data-slot="飘带标题#2" style="position:absolute;left:22px;top:300px;width:1236px;height:32px">
      <div data-id="trap" data-kind="shape" data-name="t" data-purpose="content" style="position:absolute;left:518px;top:0px;width:200px;height:32px;background-color:#B8000E"></div>
      <div data-id="label" data-kind="text" data-name="l" data-purpose="content" data-slot="text" style="position:absolute;left:582px;top:4px;width:72px;height:24px"><p data-paragraph="true" style="font-size:18px;line-height:24px"><span style="font-size:18px;color:#FFFFFF">飘带标题</span></p></div></div>`);
    const wide = doc(fillLayout(ribbon, { sets: [{ set: "飘带标题#2", items: [{ fields: { text: "以阶段验证保障上线质量" } }] }] }).html);
    const label = wide.querySelector<HTMLElement>('[data-id="label"]')!, trap = wide.querySelector<HTMLElement>('[data-id="trap"]')!;
    const w = parseFloat(label.style.width);
    expect(w).toBeGreaterThanOrEqual(18 * 11);
    // Still centred, and the red frame grew by the same amount around the same centre.
    expect(parseFloat(label.style.left) + w / 2).toBeCloseTo(618, 0);
    expect(parseFloat(trap.style.width) - 200).toBeCloseTo(w - 72, 0);
    expect(parseFloat(trap.style.left) + parseFloat(trap.style.width) / 2).toBeCloseTo(618, 0);
  });

  it("removes a whole set with an empty list and reports blocks left as sample content", () => {
    const gone = fillLayout(row, { sets: [{ set: "要点卡#1", items: [] }] });
    expect(doc(gone.html).querySelector('[data-slot="要点卡#1"],[data-slot="sep#1"]')).toBeNull();
    expect(fillLayout(row, {}).notes.join()).toContain("要点卡#1");
    expect(() => fillLayout(row, { sets: [{ set: "不存在", items: [] }] })).toThrow(/FILL_LAYOUT/);
    expect(() => fillLayout(row, { sets: [{ set: "要点卡#1", items: [{ icons: { icon: "<b>x</b>" } }] }] })).toThrow(/svg/);
  });

  const box = (id: string, slot: string, x: number, y: number, w: number, h: number, inner = "") =>
    `<div data-id="${id}" data-kind="group" data-name="${id}" data-purpose="content" data-slot="${slot}" style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px">${inner}</div>`;
  const textBox = (id: string, slot: string, x: number, y: number, w: number, h: number, html: string, align = "left") =>
    `<div data-id="${id}" data-kind="text" data-name="${id}" data-purpose="content"${slot ? ` data-slot="${slot}"` : ""} style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px">${html.replace(/<p /g, `<p style="text-align:${align}" `)}</div>`;
  const shapeBox = (x: number, y: number, w: number, h: number, fill: string) =>
    `<div data-kind="shape" data-purpose="content" style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;background-color:${fill}"></div>`;
  const big = (t: string) => `<p data-paragraph="true"><span style="font-size:56px">${t}</span></p>`;

  it("keeps a long cover title on one line: widens it in the page margins, then shrinks its font", () => {
    const cover = page(textBox("cover", "", 490, 340, 300, 70, big("智能应用建设方案"), "center"));
    const fit = doc(fillLayout(cover, { texts: { cover: "雅砻江知识运营平台二期需求汇报" } }).html).querySelector<HTMLElement>('[data-id="cover"]')!;
    const w = parseFloat(fit.style.width), x = parseFloat(fit.style.left);
    expect(w).toBeGreaterThanOrEqual(15 * 56);
    expect(x + w / 2).toBeCloseTo(640, 0);
    expect(fit.querySelector("span")!.style.fontSize).toBe("56px");
    const long = doc(fillLayout(cover, { texts: { cover: "雅砻江流域水电开发有限公司知识运营管理平台二期需求汇报" } }).html).querySelector<HTMLElement>('[data-id="cover"]')!;
    expect([long.style.left, long.style.width]).toEqual(["48px", "1184px"]);
    expect(parseFloat(long.querySelector("span")!.style.fontSize)).toBeGreaterThanOrEqual(42);
    expect(parseFloat(long.querySelector("span")!.style.fontSize)).toBeLessThan(56);
  });

  it("keeps a long page title clear of the brand logo on its line, shrinking the font instead", () => {
    const logo = `<div data-id="logo" data-kind="shape" data-purpose="fixed" style="position:absolute;left:755px;top:-6px;width:500px;height:124px;background-color:#E60012"></div>`;
    const titled = page(logo).replace('width:700px;height:46px">', 'width:700px;height:46px">').replace(P("页面标题——副标题占位"), `<p data-paragraph="true" style="line-height:46px"><span style="font-size:38px">页面标题——副标题占位</span></p>`);
    const d = doc(fillLayout(titled, { texts: { title: "一、出发点——不压缩上下文，用函数作用域分流" } }).html);
    const title = d.querySelector<HTMLElement>('[data-id="title"]')!;
    // The template already ends the box beside the logo: it is not widened under it.
    expect(title.style.width).toBe("700px");
    expect(parseFloat(title.querySelector("span")!.style.fontSize)).toBeLessThan(38);
    expect(parseFloat(title.querySelector("span")!.style.fontSize)).toBeGreaterThanOrEqual(28.5);
  });

  it("removes a row's own decoration with the row before re-spacing the page", () => {
    const arc = `<div data-id="arc" data-kind="shape" data-purpose="content" style="position:absolute;left:222px;top:282px;width:836px;height:30px;background-color:#FDE1E4"></div>`;
    const bands = page(box("a", "导语条#1", 22, 100, 1236, 60) + box("b", "红头卡片#2", 22, 170, 1236, 60) + box("s1", "图标场景项#3", 400, 240, 100, 60) + box("s2", "图标场景项#3", 600, 240, 100, 60) + arc + box("d", "结论条#4", 22, 320, 1236, 60));
    const d = doc(fillLayout(bands, { sets: [{ set: "图标场景项#3", items: [] }] }).html);
    expect(d.querySelector('[data-id="arc"]')).toBeNull();
    expect(d.querySelector<HTMLElement>('[data-id="b"]')!.style.top).not.toBe("170px");
    expect(d.querySelector<HTMLElement>('[data-id="d"]')!.style.top).toBe("320px");
  });

  it("sizes re-laid texts by their new words: a short list stays under its header and the header icon keeps its size", () => {
    const tall = Array.from({ length: 7 }, () => bullet("功能项占位")).join("");
    const panel = (n: number) => box(`panel${n}`, "红头卡片#1", 22 + n * 310, 100, 290, 330,
      shapeBox(0, 0, 290, 40, "#E60012") +
      box(`hico${n}`, "icon", 14, 6, 28, 28, shapeBox(0, 0, 28, 28, "#FFFFFF")) +
      textBox(`head${n}`, "title", 52, 8, 220, 24, P("栏目标题")) +
      textBox(`body${n}`, "body", 20, 50, 250, 270, tall));
    const four = page([0, 1, 2, 3].map(panel).join(""));
    const d = doc(fillLayout(four, { sets: [{ set: "红头卡片#1", columns: 2, items: [0, 1, 2, 3].map(i => ({ fields: { title: `能力${i}`, body: ["■ 第一条", "■ 第二条"] } })) }] }).html);
    const body = d.querySelector<HTMLElement>('[data-id="body0"]')!;
    expect(parseFloat(d.querySelector<HTMLElement>('[data-id="panel0"]')!.style.height)).toBeLessThan(170);
    expect(body.style.top).toBe("50px");
    expect(parseFloat(body.style.height)).toBeLessThan(60);
    expect(d.querySelector<HTMLElement>('[data-id="hico0"]')!.style.width).toBe("28px");
  });

  it("re-spaces the page when a full-width set is removed, keeping the content's top and bottom and the page number", () => {
    const bands = page(box("a", "导语条#1", 22, 100, 1236, 60) + box("b", "飘带标题#2", 22, 170, 1236, 60) + box("c", "要点卡#3", 22, 240, 1236, 60) + box("d", "结论条#4", 22, 310, 1236, 60)
      + textBox("num", "", 1180, 680, 40, 20, P("3")));
    const d = doc(fillLayout(bands, { sets: [{ set: "飘带标题#2", items: [] }] }).html);
    const top = (id: string) => d.querySelector<HTMLElement>(`[data-id="${id}"]`)!.style.top;
    expect(d.querySelector('[data-id="b"]')).toBeNull();
    expect([top("a"), top("c"), top("d"), top("num")]).toEqual(["100px", "205px", "310px", "680px"]);
  });

  it("widens a block that hugs a short label before shrinking the label's font", () => {
    const item = (n: number) => box(`v${n}`, "价值项#1", 94 + n * 300, 8, 120, 40,
      box(`vi${n}`, "icon", 0, 0, 40, 40, shapeBox(0, 0, 40, 40, "#FDEDEE")) + textBox(`vt${n}`, "text", 52, 8, 68, 23, `<p data-paragraph="true" style="line-height:23px"><span style="font-size:17px">价值点</span></p>`));
    const bar = page(`<div data-kind="group" data-purpose="content" style="position:absolute;left:22px;top:620px;width:1236px;height:56px">${item(0)}${item(1)}${item(2)}</div>`);
    const d = doc(fillLayout(bar, { sets: [{ set: "价值项#1", items: ["问题即问即答", "知识可沉淀", "审查更高效"].map(t => ({ fields: { text: t } })) }] }).html);
    const label = d.querySelector<HTMLElement>('[data-id="vt0"]')!;
    expect(label.querySelector("span")!.style.fontSize).toBe("17px");
    expect(parseFloat(label.style.width)).toBeGreaterThanOrEqual(6 * 17);
    expect(parseFloat(d.querySelector<HTMLElement>('[data-id="v0"]')!.style.width)).toBeGreaterThan(120);
  });

  it("spreads a packed row of icon-over-label blocks along its band when their labels need the room", () => {
    // Five 44px steps packed in the middle of an empty band, a 12px label under each icon.
    const step = (n: number) => box(`s${n}`, "图标环节项#1", 422 + n * 98, 500, 44, 63,
      box(`si${n}`, "icon", 0, 0, 44, 44, shapeBox(0, 0, 44, 44, "#FFFFFF")) + textBox(`st${n}`, "text", 4, 47, 36, 16, `<p data-paragraph="true" style="line-height:16px"><span style="font-size:12px">环节</span></p>`, "center"));
    const row = page([0, 1, 2, 3, 4].map(step).join(""));
    const labels = ["每次调用独立上下文", "返回即整体释放", "捕获外层变量", "文件即共享状态", "静态与动态库"];
    const d = doc(fillLayout(row, { sets: [{ set: "图标环节项#1", items: labels.map(t => ({ fields: { text: t } })) }] }).html);
    const blocks = [0, 1, 2, 3, 4].map(n => d.querySelector<HTMLElement>(`[data-id="s${n}"]`)!);
    const text = d.querySelector<HTMLElement>('[data-id="st0"]')!, icon = d.querySelector<HTMLElement>('[data-id="si0"]')!;
    expect(parseFloat(blocks[0].style.width)).toBeGreaterThanOrEqual(9 * 12);
    expect(parseFloat(blocks[4].style.left) + parseFloat(blocks[4].style.width)).toBeLessThanOrEqual(1280 - 24);
    expect(text.querySelector("span")!.style.fontSize).toBe("12px");
    expect(parseFloat(text.style.top)).toBeGreaterThanOrEqual(44);
    expect(parseFloat(icon.style.left) + 22).toBeCloseTo(parseFloat(blocks[0].style.width) / 2, 0);
    // Cards that the ordinary re-layout already makes wide enough are not touched by it.
    const panels = page([0, 1, 2, 3].map(n => box(`c${n}`, "红头卡片#1", 22 + n * 310, 100, 290, 200, textBox(`ct${n}`, "title", 52, 8, 220, 24, P("栏目标题")))).join(""));
    const two = doc(fillLayout(panels, { sets: [{ set: "红头卡片#1", items: ["@virtual 虚函数", "@module 模块"].map(t => ({ fields: { title: t } })) }] }).html);
    expect(parseFloat(two.querySelector<HTMLElement>('[data-id="c0"]')!.style.width)).toBeGreaterThan(500);
  });

  it("lets longer words take the free room below their box in the block before they overflow", () => {
    const lines = ["标题", ...Array.from({ length: 7 }, (_, i) => `■ 第${i + 1}条说明`)];
    const d = doc(fillLayout(row, { sets: [{ set: "要点卡#1", items: [0, 1, 2].map(() => ({ fields: { text: lines } })) }] }).html);
    const text = d.querySelector<HTMLElement>('[data-id="txt1"]')!;
    // The sample box is 80px tall in a 120px card with nothing under it: eight lines grow it, not past the card.
    expect(parseFloat(text.style.height)).toBeGreaterThan(80);
    expect(parseFloat(text.style.top) + parseFloat(text.style.height)).toBeLessThanOrEqual(118);
  });

  it("keeps repeated title bars one width: a long title shrinks inside its bar instead of stretching it", () => {
    const tocRow = (n: number, y: number) => box(`row${n}`, "目录项#1", 450, y, 740, 80,
      shapeBox(0, 0, 140, 80, "#8C8C8C") + shapeBox(200, 0, 540, 80, "#8C8C8C") +
      textBox(`title${n}`, "title", 218, 23, 504, 34, `<p data-paragraph="true" style="line-height:34px"><span style="font-size:28px">目录条目</span></p>`, "center"));
    const toc = page(tocRow(0, 162) + tocRow(1, 298) + tocRow(2, 434) + tocRow(3, 570));
    const titles = ["出发点：最稳的是朴素循环，瓶颈只在上下文", "两种声明：虚函数交给 AI，模块写成代码", "实现约定"];
    const d = doc(fillLayout(toc, { sets: [{ set: "目录项#1", items: titles.map(t => ({ fields: { title: t } })) }] }).html);
    const widths = [0, 1, 2].map(n => d.querySelector<HTMLElement>(`[data-id="row${n}"] [data-kind="shape"]:nth-child(2)`)!.style.width);
    expect(new Set(widths)).toEqual(new Set(["540px"]));
    const first = d.querySelector<HTMLElement>('[data-id="title0"]')!;
    expect(first.style.width).toBe("504px");
    expect(parseFloat(first.querySelector("span")!.style.fontSize)).toBeLessThan(28);
    expect(d.querySelector<HTMLElement>('[data-id="title2"] span')!.style.fontSize).toBe("28px");
  });

  it("keeps a centred label centred when its block widens for it", () => {
    const ribbon = page(`<div data-kind="group" data-purpose="content" style="position:absolute;left:400px;top:300px;width:480px;height:40px">` + box("rib", "飘带标题#1", 140, 0, 200, 32,
      shapeBox(0, 0, 200, 32, "#B8000E") + textBox("label", "text", 20, 4, 160, 24, `<p data-paragraph="true" style="line-height:24px"><span style="font-size:18px">飘带标题</span></p>`, "center")) + `</div>`);
    const d = doc(fillLayout(ribbon, { sets: [{ set: "飘带标题#1", items: [{ fields: { text: "已废弃、不得恢复：折叠上下文" } }] }] }).html);
    const rib = d.querySelector<HTMLElement>('[data-id="rib"]')!, label = d.querySelector<HTMLElement>('[data-id="label"]')!;
    const ribCentre = parseFloat(rib.style.left) + parseFloat(rib.style.width) / 2;
    expect(ribCentre).toBeCloseTo(240, 0);
    expect(parseFloat(label.style.left) + parseFloat(label.style.width) / 2).toBeCloseTo(parseFloat(rib.style.width) / 2, 0);
  });
});
