/**
 * Writes the page tools (definitions with `executor: "page"`) and the usage
 * guide as JSON, for a plugin backend to declare in its handshake:
 *
 *   npm run page-tools -- path/to/page_tools.json
 */
import { writeFileSync } from "node:fs";
import { PAGE_TOOLS } from "../src/page/page-tools";
import { TOOL_GUIDE } from "../src/api/tools";

const target = process.argv[2];
if (!target) {
  console.error("用法：npm run page-tools -- 输出文件.json");
  process.exit(1);
}
writeFileSync(target, JSON.stringify({ tools: PAGE_TOOLS, guide: TOOL_GUIDE }, null, 2) + "\n");
console.log(`已写入 ${PAGE_TOOLS.length} 个页面工具：${target}`);
