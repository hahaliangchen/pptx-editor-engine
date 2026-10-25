// @vitest-environment node
import { beforeAll, afterAll, describe, expect, it } from "vitest";
import { createRequire } from "node:module";
import { createServer, type Server } from "node:http";
import { readFileSync, writeFileSync, mkdtempSync, rmSync, existsSync } from "node:fs";
import { join, resolve } from "node:path";
import { tmpdir } from "node:os";
import { chromium, type Browser, type Page } from "playwright";

const require = createRequire(import.meta.url);
let browser: Browser, server: Server, page: Page, output: string;
const network: string[] = [];
const browserNetwork: string[] = [];
const failedRequests: Array<{ url: string; error: string | undefined }> = [];
const base64 = readFileSync("tools/templates/blank.pptx").toString("base64");

beforeAll(async () => {
  output = mkdtempSync(join(tmpdir(), "ppt-client-test-"));
  const rspack = require("@rspack/core");
  await new Promise<void>((resolveBuild, reject) => {
    const compiler = rspack({ mode: "development", devtool: false,
      entry: resolve("tests/fixtures/design-client.ts"), output: { path: output, filename: "app.js", publicPath: "/" },
      resolve: { extensions: [".ts", ".js", ".wasm"] },
      module: { rules: [{ test: /\.ts$/, exclude: /node_modules/, use: { loader: "builtin:swc-loader", options: { jsc: { parser: { syntax: "typescript" }, target: "es2020" } } }, type: "javascript/auto" }] },
      experiments: { asyncWebAssembly: true }, optimization: { minimize: false }
    });
    compiler.run((error: Error | null, stats: any) => compiler.close(() => {
      if (error || stats.hasErrors()) reject(error ?? new Error(stats.toString({ all: false, errors: true })));
      else resolveBuild();
    }));
  });
  server = createServer((req, res) => {
    network.push(`${req.method} ${req.url}`);
    const url = new URL(req.url!, "http://localhost");
    if (url.pathname === "/") { res.setHeader("Content-Type", "text/html; charset=utf-8"); res.end('<html><body><div id="editor" style="width:1280px;height:720px"></div><script src="/app.js"></script></body></html>'); }
    else if (url.pathname === "/api/fonts") { res.setHeader("Content-Type", "application/json"); res.end(JSON.stringify(["Noto Sans CJK SC", "DejaVu Sans"])); }
    else if (url.pathname === "/api/font") { res.end(readFileSync(url.searchParams.get("family") === "Noto Sans CJK SC" ? "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc" : "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")); }
    else {
      const file = join(output, url.pathname.replace(/^\//, ""));
      if (!file.startsWith(output + "/") || !existsSync(file)) { res.statusCode = 404; res.end(); return; }
      res.setHeader("Content-Type", file.endsWith(".wasm") ? "application/wasm" : "application/javascript; charset=utf-8");
      res.end(readFileSync(file));
    }
  });
  await new Promise<void>(resolveListen => server.listen(0, "127.0.0.1", resolveListen));
  browser = await chromium.launch();
  page = await browser.newPage({ viewport: { width: 1360, height: 850 } });
  page.on("pageerror", error => console.error("Browser runtime:", error.stack));
  page.on("request", request => browserNetwork.push(request.url()));
  page.on("requestfailed", request => failedRequests.push({ url: request.url(), error: request.failure()?.errorText }));
  await page.goto(`http://127.0.0.1:${(server.address() as any).port}`);
  await page.waitForFunction(() => !!(window as any).designTest);
  // LAN HTTP does not provide crypto.subtle; no conversion may rely on it.
  await page.evaluate(async base => {
    Object.defineProperty(crypto, "subtle", { value: undefined });
    const w = window as any, t = w.designTest;
    w.viewer = new t.PptViewer({ container: document.querySelector("#editor"), fontBackendUrl: location.origin,
      onDocumentEdited: () => w.autoBridge?.edited(), onEditInteractionEnd: () => w.autoBridge?.scheduleRefresh() });
    await w.viewer.loadPptx(t.base64ToBytes(base).buffer);
    w.converter = t.createBrowserDesignConverter(w.viewer);
  }, base64);
}, 60000);

afterAll(async () => {
  await browser?.close();
  if (server) await new Promise<void>(resolveClose => server.close(() => resolveClose()));
  if (output) rmSync(output, { recursive: true, force: true });
});

describe("conversion on the user's browser", () => {
  it("uses the actual viewer and WASM assets, edits HTML locally, and reopens only PPTX with groups and text", async () => {
    const result = await page.evaluate(async () => {
      const w = window as any, t = w.designTest;
      // An existing complex shape must become an independent preview asset.
      const added = await w.viewer.api.apply({ commands: [{ op: "addShape", slide: 1, shape: "ellipse", rect: { x: 28, y: 2, w: 3, h: 3 }, fill: "#1e8060" }] });
      if (!added.ok) throw new Error(JSON.stringify(added));
      const bridge = new t.DesignBridge(w.viewer, w.converter.convertDesign);
      const read = (r: any) => { if (!r.ok) throw new Error(JSON.stringify(r)); return JSON.parse(r.content[0].text); };
      const start = read(await bridge.run("get_design", { page: 1 }));
      const html = start.html.replace("</body>", `<div data-kind="group" style="position:absolute;left:80px;top:180px;width:900px;height:200px;display:flex;gap:24px"><div data-kind="text" style="width:438px;height:200px;padding:24px;background:#eaf2ee;display:flex;flex-direction:column;justify-content:center"><p style="font-size:28px;line-height:40px;font-family:'Noto Sans CJK SC';margin:0">浏览器本地转换<br>中文换行保持完整</p></div><div data-kind="text" style="width:438px;height:200px;padding:24px;background:#183d32;display:flex;flex-direction:column;justify-content:center"><p style="font-size:28px;line-height:40px;font-family:'Noto Sans CJK SC';color:#ffffff;margin:0">独立素材与<strong>可编辑文字</strong></p></div></div></body>`);
      read(await bridge.run("apply_html", { expected_design_revision: start.design_revision, html }));
      const stale = await bridge.run("apply_html", { expected_design_revision: start.design_revision, html });
      const persisted = await w.viewer.snapshotPptx();
      const reopened = await w.converter.syncDesign(new Uint8Array(await persisted.arrayBuffer()));
      const nodes = reopened.pages[0].nodes;
      w.lastDesign = reopened;
      const png = await w.viewer.renderPagePng(1, 1280);
      w.previewPng = t.bytesToBase64(png);
      return { texts: nodes.map(t.nodeText), grouped: nodes.filter((n: any) => n.parentId).length,
        assets: nodes.filter((n: any) => n.kind === "asset" && n.appearance.src?.startsWith("data:image/png")).length,
        staleRejected: !stale.ok, frames: document.querySelectorAll("iframe").length,
        nodeAvailable: typeof (window as any).Buffer, pending: bridge.pending, pngBytes: png.length,
        bold: nodes.some((n: any) => n.appearance.text?.paragraphs.some((p: any) => p.runs.some((r: any) => r.bold && r.text.includes("可编辑")))) };
    });
    expect(result.texts).toContain("浏览器本地转换\u2028中文换行保持完整");
    expect(result.grouped).toBe(2);
    expect(result.assets).toBeGreaterThan(0);
    expect(result.staleRejected).toBe(false);
    expect(result.frames).toBe(0);
    expect(result.nodeAvailable).toBe("undefined");
    expect(result.pending).toBe(true);
    expect(result.pngBytes).toBeGreaterThan(1000);
    expect(result.bold).toBe(true);
    expect(network.some(request => request.startsWith("POST") || request.includes("/design"))).toBe(false);
    if (process.env.PPT_DESIGN_TEST_OUTPUT) writeFileSync(process.env.PPT_DESIGN_TEST_OUTPUT,
      Buffer.from(await page.evaluate(() => (window as any).previewPng), "base64"));
  }, 60000);

  it("adds, moves and removes pages locally in a 20-page document", async () => {
    const result = await page.evaluate(async base => {
      const w = window as any;
      let design = await w.converter.convertDesign({ operation: "sync", pptxBase64: base });
      design = await w.converter.convertDesign({ operation: "pages", design, expectedRevision: design.revision, commands: Array.from({ length: 19 }, () => ({ op: "addSlide" })) });
      const count = design.pages.length, lastId = design.pages[19].id;
      design = await w.converter.convertDesign({ operation: "pages", design, expectedRevision: design.revision, commands: [{ op: "moveSlide", slide: 20, to: 1 }, { op: "deleteSlide", slide: 20 }] });
      return { count, remaining: design.pages.length, moved: design.pages[0].id === lastId };
    }, base64);
    expect(result).toEqual({ count: 20, remaining: 19, moved: true });
  }, 30000);

  it("invalidates HTML on manual commit and automatically refreshes it after edit, undo and redo", async () => {
    const initial = await page.evaluate(async () => {
      const w = window as any, t = w.designTest;
      w.autoBridge = new t.DesignBridge(w.viewer, w.converter.convertDesign);
      const reply = await w.autoBridge.run("get_design", { page: 1 });
      const before = JSON.parse(reply.content[0].text);
      const result = await w.viewer.api.apply({ commands: [{ op: "replaceText", target: { slide: 1, text: "浏览器本地转换" }, find: "浏览器本地转换", replace: "用户手动修改" }] });
      if (!result.ok) throw new Error(JSON.stringify(result));
      const invalidated = w.autoBridge.peekHtml(1) === undefined;
      const stale = await w.autoBridge.run("apply_html", { expected_design_revision: before.design_revision, html: before.html });
      return { invalidated, staleRejected: !stale.ok, revision: before.design_revision };
    });
    expect(initial.invalidated).toBe(true);
    expect(initial.staleRejected).toBe(false);
    await page.waitForFunction(() => (window as any).autoBridge.peekHtml(1)?.includes("用户手动修改"));
    const after = await page.evaluate(() => {
      const w = window as any;
      const revision = Number(new DOMParser().parseFromString(w.autoBridge.peekHtml(1), "text/html").body.dataset.designRevision);
      w.viewer.undo();
      return { revision, invalidated: w.autoBridge.peekHtml(1) === undefined };
    });
    expect(after.revision).toBeGreaterThan(initial.revision);
    expect(after.invalidated).toBe(true);
    await page.waitForFunction(() => (window as any).autoBridge.peekHtml(1)?.includes("浏览器本地转换"));
    await page.evaluate(() => (window as any).viewer.redo());
    await page.waitForFunction(() => (window as any).autoBridge.peekHtml(1)?.includes("用户手动修改"));
    await page.evaluate(() => { const w = window as any; w.autoBridge.dispose(); w.autoBridge = null; });
  }, 30000);

  it("blocks scripts and external assets and cleans up the isolated frame after invalid layouts", async () => {
    const result = await page.evaluate(async () => {
      const w = window as any, t = w.designTest;
      const base = t.jsonToHtml(w.lastDesign, w.lastDesign.pages[0].id), errors = [];
      for (const extra of ['<script>parent.designInjected=true</script>', '<img data-kind="image" src="https://example.invalid/probe.png">', '<div data-kind="text" style="width:100px;height:5px"><p style="font-size:40px">溢出</p></div>', '<style>@import url("https://example.invalid/probe.css");</style><div data-kind="shape" style="width:80px;height:80px;filter:blur(3px)"></div>']) {
        try { await t.measureHtmlInBrowser(base.replace("</body>", extra + "</body>")); errors.push(""); }
        catch (error) { errors.push(String(error)); }
      }
      return { errors, injected: !!w.designInjected, frames: document.querySelectorAll("iframe").length };
    });
    expect(result.errors.every(message => message.length > 0)).toBe(true);
    expect(result.injected).toBe(false);
    expect(result.frames).toBe(0);
    expect(browserNetwork.filter(url => url.includes("probe.png"))).toEqual([]);
    expect(failedRequests.filter(request => request.url.includes("probe.css"))).toEqual([
      { url: "https://example.invalid/probe.css", error: expect.stringMatching(/csp/i) }
    ]);
  }, 30000);
});
