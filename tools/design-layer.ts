import { readFile, writeFile, mkdir } from "node:fs/promises";
import { dirname } from "node:path";
import { importDesign } from "../src/design/import";
import { exportDesign } from "../src/design/export";
import { projectPage } from "../src/design/projection";
import { mergeHtml } from "../src/design/measure";
import { DesignDocument } from "../src/design/model";

const [command, input, output, extra] = process.argv.slice(2).filter(arg => arg !== "--");
async function save(file: string, data: string | Buffer): Promise<void> { await mkdir(dirname(file), { recursive: true }); await writeFile(file, data); }
async function main() {
  if (!input || !output) throw new Error("Usage: npm run design -- import input.pptx design.json | project design.json page.html [page-number] | merge design.json updated.json page.html | export design.json output.pptx");
  if (command === "import") {
    const doc = await importDesign(await readFile(input)); await save(output, JSON.stringify(doc, null, 2));
    console.log(JSON.stringify({ output, pages: doc.pages.length, nodes: doc.pages.reduce((n,p) => n + p.nodes.length, 0), diagnostics: doc.diagnostics })); return;
  }
  const doc: DesignDocument = JSON.parse(await readFile(input, "utf8"));
  if (command === "project") { const page = doc.pages[Number(extra ?? 1) - 1]; if (!page) throw new Error("Invalid page number"); await save(output, projectPage(doc, page.id)); }
  else if (command === "merge") { if (!extra) throw new Error("merge requires the edited HTML file"); await save(output, JSON.stringify(await mergeHtml(doc, await readFile(extra, "utf8")), null, 2)); }
  else if (command === "export") await save(output, await exportDesign(doc));
  else throw new Error(`Unknown command: ${command}`);
  console.log(JSON.stringify({ output }));
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
