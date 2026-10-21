import { prepareTemplateSemantics, finishTemplateSemantics, type TemplateSemanticAuthorization } from "./template-semantics";
import { measureDesignDOM } from "./measure-dom";
import { type DesignDocument, type MeasuredPage } from "./model";
import { baselineFor, mergeHtmlDifference } from "./html-diff";
import { parseDesignHtml } from "./html-document";

export type PrepareDesignFonts = (document: Document) => Promise<void>;

/** No scripts or network in the frame; the trusted parent reads only layout. */
export async function measureHtmlInBrowser(html: string, prepareFonts?: PrepareDesignFonts, collectIssues = false): Promise<MeasuredPage> {
  const parsed = parseDesignHtml(html);
  for (const img of Array.from(parsed.images)) {
    if (!img.getAttribute("src")?.startsWith("data:image/") || img.hasAttribute("srcset")) throw new Error("Images must use embedded project assets");
  }
  const policy = parsed.createElement("meta");
  policy.httpEquiv = "Content-Security-Policy";
  policy.content = "default-src 'none'; script-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src 'none'; connect-src 'none'; form-action 'none'; base-uri 'none'";
  parsed.head.prepend(policy);
  const frame = document.createElement("iframe");
  frame.setAttribute("sandbox", "allow-same-origin");
  frame.setAttribute("aria-hidden", "true");
  frame.tabIndex = -1;
  frame.style.cssText = "position:fixed;left:-20000px;top:0;width:1280px;height:720px;border:0;pointer-events:none";
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    const loaded = new Promise<void>((resolve, reject) => {
      frame.onload = () => resolve();
      frame.onerror = () => reject(new Error("Cannot load the HTML design frame"));
      timer = setTimeout(() => reject(new Error("HTML design frame timed out")), 15000);
    });
    frame.srcdoc = "<!doctype html>" + parsed.documentElement.outerHTML;
    document.body.appendChild(frame);
    await loaded;
    clearTimeout(timer);
    const doc = frame.contentDocument!;
    // This policy was inserted by us, not supplied by the model. Keep it active
    // while removing the element before the shared static-markup validation.
    doc.querySelector("meta[http-equiv]")?.remove();
    const measured = (async () => {
      if (prepareFonts) await prepareFonts(doc);
      else document.fonts.forEach(face => (doc.fonts as FontFaceSet & { add(face: FontFace): void }).add(face));
      const nonce = Array.from(crypto.getRandomValues(new Uint8Array(16)), n => n.toString(16).padStart(2, "0")).join("");
      return measureDesignDOM({ nonce, collectIssues }, doc);
    })();
    return await Promise.race([measured, new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new Error("HTML design measurement timed out")), 30000);
    })]);
  } finally {
    clearTimeout(timer);
    frame.remove();
  }
}

export async function mergeHtmlInBrowser(design: DesignDocument, html: string, prepareFonts?: PrepareDesignFonts, baselineHtml?: string, authorization?: TemplateSemanticAuthorization): Promise<DesignDocument> {
  const baseline = baselineFor(design, html, baselineHtml);
  const prepared = prepareTemplateSemantics(design, baseline, html, authorization);
  const before = await measureHtmlInBrowser(baseline, prepareFonts, true);
  const after = await measureHtmlInBrowser(html, prepareFonts, true);
  return finishTemplateSemantics(design, mergeHtmlDifference(prepared.design, before, after), prepared.proposal);
}
