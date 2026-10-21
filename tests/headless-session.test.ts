// @vitest-environment node
import { afterEach, describe, expect, it } from "vitest";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { HeadlessSession } from "../src/headless/session";
import { buildDeck } from "./helpers/deck";
import { apiDeckParts } from "./helpers/api-deck";
import { checkPackage } from "./helpers/package-check";
import { readFileSync } from "node:fs";

const sessions: HeadlessSession[] = [];
async function open(): Promise<HeadlessSession> {
  const session = await HeadlessSession.open(await buildDeck(apiDeckParts()));
  sessions.push(session);
  return session;
}
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

async function pixelAt(png: Buffer, x: number, y: number): Promise<number[]> {
  const image = await loadImage(png);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  return [...ctx.getImageData(Math.round(x), Math.round(y), 1, 1).data];
}

const pngDataUrl = (color: string) => {
  const canvas = createCanvas(40, 30);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, 40, 30);
  return `data:image/png;base64,${canvas.toBuffer("image/png").toString("base64")}`;
};

/** A point inside the blue card, clear of its rounded corner and its text. */
const insideCard = (width: number) => {
  const pxPerEmu = width / 9144000;
  return [(914400 + 400000) * pxPerEmu, (1828800 + 100000) * pxPerEmu];
};

describe("headless session", () => {
  it("draws upward and vertical native connectors from exact endpoints through export", async () => {
    const session=await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"),{quiet:true});sessions.push(session);
    expect((await session.apply([{op:"addConnector",slide:1,from:{x:4,y:9},to:{x:8,y:5},arrow:false,line:{color:"#D97757",widthPt:3}},{op:"addConnector",slide:1,from:{x:12,y:5},to:{x:12,y:9},arrow:false,line:{color:"#D97757",widthPt:3}}])).ok).toBe(true);
    const reopened=await HeadlessSession.open(await session.export(),{quiet:true});sessions.push(reopened);
    for(const deck of [session,reopened]){
      const png=await deck.render(1,{width:1280}),scale=1280/(await deck.outline()).slideSize.w;
      for(const [x,y]of[[4.5,8.5],[7.5,5.5],[12,7]])expect((await pixelAt(png,x*scale,y*scale)).slice(0,3)).toEqual([217,119,87]);
      expect((await pixelAt(png,4.5*scale,5.5*scale)).slice(0,3)).toEqual([255,255,255]);
    }
    expect(await session.apply([{op:"addConnector",slide:1,from:{x:2,y:2},to:{x:2,y:2}}])).toMatchObject({ok:false,rolledBack:true,appliedCommands:0});
  });
  it("duplicates after the requested page, preserving independent native text through export", async () => {
    const session=await open();
    expect((await session.apply([{op:"setText",target:{slide:1,name:"卡片"},text:"A"},{op:"duplicateSlide",slide:1},{op:"setText",target:{slide:2,name:"卡片"},text:"B"}])).ok).toBe(true);
    expect(await session.apply([{op:"duplicateSlide",slide:1,after:2}])).toMatchObject({ok:true,results:[{slide:3}]});
    expect(await session.apply([{op:"duplicateSlide",slide:2,after:0}])).toMatchObject({ok:true,results:[{slide:1}]});
    const reopened=await HeadlessSession.open(await session.export(),{quiet:true});sessions.push(reopened);
    for(const deck of [session,reopened])expect((await deck.outline()).slides.map(s=>s.shapes.find(x=>x.name==="卡片")?.text)).toEqual(["B","A","B","A"]);
  });
  it.each([
    {shape: "leftArrow", rect: {x: 3,y: 5,w: 5,h: 2}, points: [[3.4,6],[7.6,6]]},
    {shape: "upArrow", rect: {x: 3,y: 5,w: 2,h: 5}, points: [[4,5.4],[4,9.6]]},
    {shape: "downArrow", rect: {x: 3,y: 5,w: 2,h: 5}, points: [[4,5.4],[4,9.6]]}
  ])("creates and reopens a visible native $shape", async ({shape,rect,points}) => {
    const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), {quiet: true});
    sessions.push(session);
    expect((await session.apply({commands: [{op: "addShape", slide: 1, shape, rect, fill: "#D97757", line: null}]})).ok).toBe(true);
    const reopened = await HeadlessSession.open(await session.export(), {quiet: true});
    sessions.push(reopened);
    for (const deck of [session,reopened]) {
      const picture=await deck.render(1,{width:1280});
      const scale=1280/(await deck.outline()).slideSize.w;
      for(const [x,y] of points)expect((await pixelAt(picture,x*scale,y*scale)).slice(0,3)).toEqual([217,119,87]);
    }
  });
  it("draws a native arrow shaft and head inside its bounds before and after PPTX export", async () => {
    const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), {quiet: true});
    sessions.push(session);
    expect((await session.apply({commands: [{op: "addShape", slide: 1, name: "箭头", shape: "rightArrow", rect: {x: 3, y: 5, w: 5, h: 2}, fill: "#D97757", line: null}]})).ok).toBe(true);
    const reopened = await HeadlessSession.open(await session.export(), {quiet: true});
    sessions.push(reopened);
    for (const deck of [session, reopened]) {
      const picture = await deck.render(1, {width: 1280});
      const scale = 1280 / (await deck.outline()).slideSize.w;
      expect((await pixelAt(picture, 3.4 * scale, 6 * scale)).slice(0, 3)).toEqual([217,119,87]);
      expect((await pixelAt(picture, 7.6 * scale, 6 * scale)).slice(0, 3)).toEqual([217,119,87]);
      expect((await pixelAt(picture, 3.4 * scale, 12 * scale)).slice(0, 3)).toEqual([255,255,255]);
    }
  });
  it("reads the deck's structure", async () => {
    const outline = await (await open()).outline();
    expect(outline.slides).toHaveLength(1);
    expect(outline.slides[0].shapes.map((shape: { name: string }) => shape.name)).toContain("卡片");
  });

  it("applies commands and the next render shows them", async () => {
    const session = await open();
    const [x, y] = insideCard(960);
    expect((await pixelAt(await session.render(1, { width: 960 }), x, y)).slice(0, 3)).toEqual([0x44, 0x72, 0xc4]);

    const result = await session.apply([{ op: "setFill", target: { slide: 1, name: "卡片" }, color: "#e60012" }]);
    expect(result.ok).toBe(true);
    expect((await pixelAt(await session.render(1, { width: 960 }), x, y)).slice(0, 3)).toEqual([0xe6, 0x00, 0x12]);
  });

  it("undo restores the previous look", async () => {
    const session = await open();
    await session.apply([{ op: "setFill", target: { slide: 1, name: "卡片" }, color: "#e60012" }]);
    session.undo();
    const [x, y] = insideCard(960);
    expect((await pixelAt(await session.render(1, { width: 960 }), x, y)).slice(0, 3)).toEqual([0x44, 0x72, 0xc4]);
  });

  it("rolls a failing batch back as one step", async () => {
    const session = await open();
    const result = await session.apply([
      { op: "setFill", target: { slide: 1, name: "卡片" }, color: "#e60012" },
      { op: "setFill", target: { slide: 1, name: "不存在的形状" }, color: "#000000" }
    ]);
    expect(result.ok).toBe(false);
    const [x, y] = insideCard(960);
    expect((await pixelAt(await session.render(1, { width: 960 }), x, y)).slice(0, 3)).toEqual([0x44, 0x72, 0xc4]);
  });

  it("a new page with a picture exports to a package that opens again", async () => {
    const session = await open();
    const result = await session.apply([
      { op: "addSlide", after: 1 },
      { op: "addImage", slide: 2, src: pngDataUrl("#ff6a00"), rect: { x: 2, y: 2, w: 6, h: 4 }, name: "橙色块" },
      { op: "addText", slide: 2, rect: { x: 2, y: 7, w: 10, h: 2 }, text: "第二页" }
    ]);
    expect(result.ok).toBe(true);

    const bytes = await session.export();
    // The bare test fixture never declares a content type for .rels parts; that is not the export's doing.
    const problems = (await checkPackage(bytes)).filter(message => !/^no content type for .*\.rels$/.test(message));
    expect(problems).toEqual([]);

    const reopened = await HeadlessSession.open(bytes);
    sessions.push(reopened);
    const outline = await reopened.outline();
    expect(outline.slides).toHaveLength(2);
    const names = outline.slides[1].shapes.map((shape: { name: string }) => shape.name);
    expect(names).toContain("橙色块");
    // The picture is on the page, in colour.
    const png = await reopened.render(2, { width: 960 });
    const pxPerCm = 960 / (9144000 / 360000);
    expect((await pixelAt(png, 5 * pxPerCm, 4 * pxPerCm)).slice(0, 3)).toEqual([0xff, 0x6a, 0x00]);
  });

  it("names new text boxes and shapes, so later commands can find them by name", async () => {
    const session = await open();
    const result = await session.apply([
      { op: "addText", slide: 1, rect: { x: 1, y: 1, w: 5, h: 2 }, text: "说明", name: "说明文字" },
      { op: "addShape", slide: 1, shape: "rect", rect: { x: 8, y: 1, w: 3, h: 2 }, name: "色块" },
      { op: "setText", target: { slide: 1, name: "说明文字" }, text: "改过" }
    ]);
    expect(result.ok).toBe(true);
    const reopened = await HeadlessSession.open(await session.export());
    sessions.push(reopened);
    const shapes = (await reopened.outline()).slides[0].shapes;
    expect(shapes.find((shape: { name: string }) => shape.name === "说明文字")).toMatchObject({ text: "改过" });
    expect(shapes.some((shape: { name: string }) => shape.name === "色块")).toBe(true);
  });

  it("two sessions in one process stay independent", async () => {
    const first = await open();
    const second = await open();
    await first.apply([{ op: "setFill", target: { slide: 1, name: "卡片" }, color: "#e60012" }]);
    const [x, y] = insideCard(960);
    expect((await pixelAt(await second.render(1, { width: 960 }), x, y)).slice(0, 3)).toEqual([0x44, 0x72, 0xc4]);
  });
});
