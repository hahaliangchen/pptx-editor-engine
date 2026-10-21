import { resolveObjectURL } from "node:buffer";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { HeadlessSession } from "../headless/session";
import { HeadlessRenderer } from "../headless/renderer";
import type { Slide, SlideElement } from "../pptx-virtual-dom";
import type { DesignDocument, DesignRect } from "./model";
import { importDocument } from "./import-core";
async function dataUrl(url: string): Promise<string> {
  if (url.startsWith("data:image/")) return url;
  if (url.startsWith("data:")) {
    const bytes = Buffer.from(url.slice(url.indexOf(",") + 1), "base64");
    return embeddedImage(bytes);
  }
  const blob = url.startsWith("blob:") ? resolveObjectURL(url) : undefined;
  if (!blob) throw new Error("Source picture has no embedded data");
  return embeddedImage(Buffer.from(await blob.arrayBuffer()));
}
function embeddedImage(bytes: Buffer): string {
  const type = bytes.subarray(0, 8).equals(Buffer.from([137,80,78,71,13,10,26,10])) ? "image/png"
    : bytes[0] === 255 && bytes[1] === 216 ? "image/jpeg"
    : bytes.subarray(0, 3).toString() === "GIF" ? "image/gif"
    : /<svg[\s>]/i.test(bytes.subarray(0, 1024).toString()) ? "image/svg+xml" : undefined;
  if (!type) throw new Error("Unsupported embedded image type; supply a PNG/JPEG preview");
  return `data:${type};base64,${bytes.toString("base64")}`;
}

export async function importDesign(bytes: Uint8Array): Promise<DesignDocument> {
  const session = await HeadlessSession.open(bytes, { quiet: true });
  try {
    const doc = session.document;
    const renderer = new HeadlessRenderer({ quiet: true });
    renderer.prepare(doc);
    const preview = async (slide: Slide, elements: SlideElement[], bounds: DesignRect): Promise<string> => {
      const image = await loadImage(await renderer.renderSlide({ ...doc, slides: [{ ...slide, elements }] }, 0, { width: 1280 }));
      const canvas = createCanvas(Math.max(1, Math.ceil(bounds.w * 2)), Math.max(1, Math.ceil(bounds.h * 2)));
      canvas.getContext("2d").drawImage(image, bounds.x, bounds.y, Math.max(1, bounds.w), Math.max(1, bounds.h), 0, 0, canvas.width, canvas.height);
      return `data:image/png;base64,${canvas.toBuffer("image/png").toString("base64")}`;
    };
    return await importDocument(bytes, doc, { dataUrl, preview });
  } finally { session.dispose(); }
}
