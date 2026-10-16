import { readFileSync, writeFileSync } from "node:fs";
import { HeadlessSession } from "../src/headless/session";

const USAGE = `用法：
  npm run headless -- outline <文稿.pptx> [页码]
  npm run headless -- render  <文稿.pptx> <页码> <输出.png> [宽度像素，默认 1280]
  npm run headless -- apply   <文稿.pptx> <指令.json> <输出.pptx> [--render 输出目录]
需要先运行 npm run build:wasm:node。`;

async function main(): Promise<void> {
  const [command, deck, ...rest] = process.argv.slice(2);
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
