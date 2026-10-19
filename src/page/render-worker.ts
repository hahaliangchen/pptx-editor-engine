/// <reference lib="webworker" />
/**
 * Draws slides to PNG off the main thread, so rendering for the Agent neither
 * blocks the editor nor stalls when the tab is in the background (workers are
 * not paused like requestAnimationFrame).
 *
 * The Rust renderer is written against `web_sys`: it looks up
 * `window.document`, makes helper canvases with `createElement("canvas")` and
 * type-checks with `instanceof`. A worker has none of that, so OffscreenCanvas
 * stands in, the same trick the Node renderer plays with Skia canvases.
 */
import type { RenderRequest, WorkerReply } from "./render-protocol";

const scope = self as any;

function installWorkerEnv(): void {
  class WorkerWindow {
    static [Symbol.hasInstance](value: unknown) { return value === self; }
  }
  scope.Window = WorkerWindow;
  scope.window = self;
  scope.HTMLCanvasElement = OffscreenCanvas;
  scope.CanvasRenderingContext2D = OffscreenCanvasRenderingContext2D;
  // Images arrive as ImageBitmaps; the renderer reads naturalWidth/Height.
  scope.HTMLImageElement = ImageBitmap;
  for (const name of ["naturalWidth", "naturalHeight"] as const) {
    if (!(name in ImageBitmap.prototype)) {
      const size = name === "naturalWidth" ? "width" : "height";
      Object.defineProperty(ImageBitmap.prototype, name, { get(this: ImageBitmap) { return this[size]; } });
    }
  }
  scope.document = { createElement: (tag: string) => {
    if (tag.toLowerCase() !== "canvas") throw new Error(`worker cannot create <${tag}>`);
    return new OffscreenCanvas(1, 1);
  } };
  scope.Document = Object;
}

installWorkerEnv();

const canvas = new OffscreenCanvas(1, 1);
const ctx = canvas.getContext("2d")!;
const images = new Map<string, Promise<ImageBitmap | null>>();
let rendererPromise: Promise<any> | null = null;
const pendingFonts: Uint8Array[] = [];

function renderer(): Promise<any> {
  rendererPromise ??= import("../../rust-engine/pkg/ppt_engine").then(wasm => {
    const created = quietly(() => new wasm.RustPptRenderer(ctx as any));
    for (const bytes of pendingFonts.splice(0)) quietly(() => created.register_font(bytes));
    return created;
  });
  return rendererPromise;
}

/** The WASM logs every font face and slide; keep the console readable. */
function quietly<T>(run: () => T): T {
  const log = console.log;
  console.log = () => undefined;
  try { return run(); } finally { console.log = log; }
}

function image(url: string): Promise<ImageBitmap | null> {
  let loading = images.get(url);
  if (!loading) {
    // blob: URLs made by the page are readable here (same origin).
    loading = fetch(url).then(r => r.blob()).then(blob => createImageBitmap(blob)).catch(error => {
      console.warn(`[render-worker] 图片无法解码，已跳过：${url}`, error);
      return null;
    });
    images.set(url, loading);
  }
  return loading;
}

async function render(request: RenderRequest): Promise<ArrayBuffer> {
  const engine = await renderer();
  const bitmaps: Record<string, ImageBitmap> = {};
  await Promise.all(request.images.map(async url => {
    const bitmap = await image(url);
    if (bitmap) bitmaps[url] = bitmap;
  }));
  const scale = request.width / request.slideSize.width;
  canvas.width = request.width;
  canvas.height = Math.max(1, Math.round(request.slideSize.height * scale));
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  quietly(() => engine.render_slide(request.slideJson, bitmaps));
  const blob = await canvas.convertToBlob({ type: "image/png" });
  return blob.arrayBuffer();
}

scope.onmessage = async (event: MessageEvent) => {
  const message = event.data;
  if (message.type === "font") {
    if (rendererPromise) (await rendererPromise).register_font(message.bytes);
    else pendingFonts.push(message.bytes);
    return;
  }
  if (message.type === "forget-images") {
    images.clear();
    return;
  }
  if (message.type === "render") {
    let reply: WorkerReply;
    try {
      const png = await render(message);
      reply = { id: message.id, png };
      scope.postMessage(reply, [png]);
    } catch (error) {
      reply = { id: message.id, error: error instanceof Error ? error.message : String(error) };
      scope.postMessage(reply);
    }
  }
};
