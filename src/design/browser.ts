import { PptxParser, type PptVirtualDocument, type Slide, type SlideElement } from "../pptx-parser";
import { DocumentSession } from "../document-session";
import { importDocument } from "./import-core";
import { exportDesignBytes } from "./export-core";
import { createDesignConverter } from "./service-core";
import { mergeHtmlInBrowser, type PrepareDesignFonts } from "./measure-browser";
import { bytesToBase64 } from "./bytes";
import type { DesignRect } from "./model";

export interface BrowserDesignHost {
  prepareDesignDocument(document: PptVirtualDocument): Promise<void>;
  prepareDesignFonts: PrepareDesignFonts;
  renderDesignAsset(document: PptVirtualDocument, slide: Slide, elements: SlideElement[], bounds: DesignRect): Promise<string>;
}

async function embeddedImage(url: string): Promise<string> {
  if (url.startsWith("data:image/")) return url;
  if (!url.startsWith("blob:") && !url.startsWith("data:")) throw new Error("Source image must be embedded");
  const response = await fetch(url);
  const bytes = new Uint8Array(await response.arrayBuffer());
  const header = new TextDecoder().decode(bytes.subarray(0, 1024));
  const mime = bytes[0] === 137 && bytes[1] === 80 ? "image/png"
    : bytes[0] === 255 && bytes[1] === 216 ? "image/jpeg"
    : header.startsWith("GIF") ? "image/gif" : /<svg[\s>]/i.test(header) ? "image/svg+xml" : undefined;
  if (!mime) throw new Error("Unsupported embedded image type; supply a PNG/JPEG preview");
  return `data:${mime};base64,${bytesToBase64(bytes)}`;
}

function releaseImages(doc?: PptVirtualDocument): void {
  const urls = new Set<string>();
  for (const slide of doc?.slides ?? []) for (const element of slide.elements) {
    if (element.type === "image") urls.add(element.url);
    if (element.type === "shape" && element.computedStyle?.fill?.type === "picture" && element.computedStyle.fill.url) urls.add(element.computedStyle.fill.url);
  }
  for (const url of urls) if (url.startsWith("blob:")) URL.revokeObjectURL(url);
}

/** All conversions run in the user's browser; no conversion HTTP endpoint. */
export function createBrowserDesignConverter(host: BrowserDesignHost) {
  return createDesignConverter({
    async importDesign(bytes) {
      const doc = await new PptxParser().parse(bytes.slice().buffer as ArrayBuffer);
      try {
        await host.prepareDesignDocument(doc);
        return await importDocument(bytes, doc, {
          dataUrl: embeddedImage,
          preview: (slide, elements, bounds) => host.renderDesignAsset(doc, slide, elements, bounds)
        });
      } finally { releaseImages(doc); }
    },
    exportDesign: exportDesignBytes,
    mergeHtml: (design, html, baseline, authorization) => mergeHtmlInBrowser(design, html, doc => host.prepareDesignFonts(doc), baseline, authorization),
    async pages(bytes, commands) {
      const session = new DocumentSession({ container: document.createElement("div"),
        view: { invalidate() {}, beginDrag() {}, endDrag() {}, slideEdited() {}, beginTextEdit() {}, endTextEdit() {},
          async slidesChanged() {}, async currentChanged() {}, async prepareSlide() {} },
        async probeImage(bytes, contentType) {
          const image = await createImageBitmap(new Blob([bytes.slice().buffer as ArrayBuffer], { type: contentType }));
          try { return { width: image.width, height: image.height }; } finally { image.close(); }
        }
      });
      try {
        await session.open(bytes.slice().buffer as ArrayBuffer);
        const result = await session.api.apply({ commands });
        if (!result.ok) throw new Error(JSON.stringify(result));
        if (!session.presentation?.slides.length) throw new Error("Keep at least one page");
        return new Uint8Array(await (await session.exportPptx(false)).arrayBuffer());
      } finally { releaseImages(session.presentation ?? undefined); session.dispose(); }
    }
  });
}

export { projectPage as jsonToHtml } from "./projection";
export { exportDesignBytes as jsonToPptx } from "./export-core";
export { mergeHtmlInBrowser as htmlToJson } from "./measure-browser";
export type { DesignRequest } from "./service-core";
