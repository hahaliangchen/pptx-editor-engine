// @vitest-environment node
import { readFileSync } from "node:fs";
import { afterEach, expect, it } from "vitest";
import { HeadlessSession } from "../src/headless/session";

const sessions: HeadlessSession[] = [];
afterEach(() => { for (const session of sessions.splice(0)) session.dispose(); });

it("renames a shape and a group, keeps the name through export, and undoes", async () => {
  const session = await HeadlessSession.open(readFileSync("tools/templates/blank.pptx"), { quiet: true });
  sessions.push(session);
  const made = await session.apply({ commands: [
    { op: "addText", slide: 1, rect: { x: 2, y: 2, w: 8, h: 2 }, text: "标题", as: "t" },
    { op: "addShape", slide: 1, shape: "rect", rect: { x: 2, y: 6, w: 4, h: 3 }, fill: "#336699", as: "r" },
    { op: "setName", target: "$t", name: "@封面.标题" },
    { op: "setName", target: "$r", name: "@内容.卡片1" }
  ] });
  expect(made.ok).toBe(true);
  const names = async (s: HeadlessSession) => (await s.outline()).slides[0].shapes.map((shape: { name: string }) => shape.name);
  expect(await names(session)).toEqual(expect.arrayContaining(["@封面.标题", "@内容.卡片1"]));

  // the name is in the saved file
  const reopened = await HeadlessSession.open(await session.export(), { quiet: true });
  sessions.push(reopened);
  expect(await names(reopened)).toEqual(expect.arrayContaining(["@封面.标题", "@内容.卡片1"]));

  session.undo();
  expect(await names(session)).not.toContain("@内容.卡片1");
});
