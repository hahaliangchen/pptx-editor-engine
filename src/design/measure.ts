import { prepareTemplateSemantics, finishTemplateSemantics, type TemplateSemanticAuthorization } from "./template-semantics";
import { measureDesignDOM } from "./measure-dom";
import { createRequire } from "node:module";
import { randomUUID } from "node:crypto";
import { MeasuredPage, DesignDocument } from "./model";
import { baselineFor, mergeHtmlDifference } from "./html-diff";

/** Chromium resolves layout; this function never tries to infer flex/grid geometry itself. */
export async function measureHtml(html: string, screenshot?: string, collectIssues = false): Promise<MeasuredPage> {
  const { chromium } = createRequire(import.meta.url)("playwright");
  const browser = await chromium.launch({ ...(process.env.PPT_DESIGN_BROWSER ? { executablePath: process.env.PPT_DESIGN_BROWSER } : {}) });
  try {
    const context = await browser.newContext({ viewport: { width: 1280, height: 720 }, deviceScaleFactor: 1, javaScriptEnabled: false });
    await context.route("**/*", (route: any) => /^(data:|about:)/.test(route.request().url()) ? route.continue() : route.abort());
    const page = await context.newPage();
    await page.setContent(html, { waitUntil: "load", timeout: 15000 });
    const result = await page.evaluate(measureDesignDOM, { nonce: randomUUID(), collectIssues });
    if (screenshot) await page.screenshot({ path: screenshot });
    return result;
  } finally { await browser.close(); }
}

export async function mergeHtml(design: DesignDocument, html: string, screenshot?: string, baselineHtml?: string, authorization?: TemplateSemanticAuthorization): Promise<DesignDocument> {
  const baseline = baselineFor(design, html, baselineHtml);
  const prepared = prepareTemplateSemantics(design, baseline, html, authorization);
  const before = await measureHtml(baseline, undefined, true);
  const after = await measureHtml(html, screenshot, true);
  return finishTemplateSemantics(design, mergeHtmlDifference(prepared.design, before, after), prepared.proposal);
}
