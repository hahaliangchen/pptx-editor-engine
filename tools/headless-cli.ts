import { readFileSync, writeFileSync } from "node:fs";
import { HeadlessSession } from "../src/headless/session";
import { TOOLS, TOOL_GUIDE } from "../src/api/tools";
import { toAnthropic, toMcp, toOpenAI } from "../src/api/adapters";

const USAGE = `用法：
  npm run headless -- outline <文稿.pptx> [页码]
  npm run headless -- render  <文稿.pptx> <页码> <输出.png> [宽度像素，默认 1280]
  npm run headless -- apply   <文稿.pptx> <指令.json> <输出.pptx> [--render 输出目录]
  npm run headless -- tools   [anthropic|openai|mcp|guide]        打印给模型用的工具定义（guide 为系统提示片段）
  npm run headless -- call    <文稿.pptx> <工具名> '<参数 JSON>'     直接调用一个工具，图片写到 ./tool-image.png
需要先运行 npm run build:wasm:node。`;

async function main(): Promise<void> {
  const [command, deck, ...rest] = process.argv.slice(2);
  if (command === "tools") {
    const format = deck || "anthropic";
    if (format === "guide") console.log(TOOL_GUIDE);
    else if (format === "anthropic") console.log(JSON.stringify(toAnthropic(), null, 2));
    else if (format === "openai") console.log(JSON.stringify(toOpenAI(), null, 2));
    else if (format === "mcp") console.log(JSON.stringify(toMcp(), null, 2));
    else console.log(JSON.stringify(TOOLS, null, 2));
    return;
  }
  if (!command || !deck) {
    console.log(USAGE);
    process.exit(1);
  }
  const session = await HeadlessSession.open(readFileSync(deck), { quiet: true });
  try {
    if (command === "outline") {
      const page = rest[0] ? [Number(rest[0])] : undefined;
      console.log(JSON.stringify(await session.outline({ slides: page }), null, 2));
    } else if (command === "render") {
      const [page, out, width] = rest;
      writeFileSync(out, await session.render(Number(page), { width: width ? Number(width) : 1280 }));
      console.log(`已写入 ${out}`);
    } else if (command === "call") {
      const [name, json] = rest;
      const result = await session.callTool(name, json ? JSON.parse(json) : {});
      for (const part of result.content) {
        if (part.type === "text") console.log(part.text);
        else {
          writeFileSync("tool-image.png", Buffer.from(part.data, "base64"));
          console.log("[图片已写入 tool-image.png]");
        }
      }
      if (!result.ok) process.exitCode = 2;
    } else if (command === "apply") {
      const [commands, out, flag, dir] = rest;
      const result = await session.apply(JSON.parse(readFileSync(commands, "utf8")));
      console.log(JSON.stringify(result, null, 2));
      if (!result.ok) process.exitCode = 2;
      writeFileSync(out, await session.export());
      console.log(`已写入 ${out}`);
      if (flag === "--render" && dir) {
        const count = (await session.outline()).slides.length;
        for (let page = 1; page <= count; page++) writeFileSync(`${dir}/page${page}.png`, await session.render(page, { width: 960 }));
        console.log(`已渲染 ${count} 页到 ${dir}`);
      }
    } else {
      console.log(USAGE);
      process.exit(1);
    }
  } finally {
    session.dispose();
  }
}

main().catch(error => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
