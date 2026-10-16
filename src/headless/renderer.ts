import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { resolveObjectURL } from "node:buffer";
import { createCanvas, loadImage } from "@napi-rs/canvas";
import { expandForRender } from "../charts/layout";
import { PptVirtualDocument, Slide } from "../pptx-virtual-dom";
import { FontResolver, collectFontRequests } from "./fonts";
import { installNodeEnv } from "./node-env";

export interface HeadlessRenderOptions {
  /** Extra font files registered up front (e.g. fonts shipped with the product). */
  fonts?: string[];
  /** The WASM prints every font face it loads; hidden unless this is false. */
  quiet?: boolean;
}

export interface RenderSlideOptions {
  /** Output width in pixels; the height follows the slide's aspect ratio. */
  width?: number;
}

/**
 * Draws slides to PNG without a browser: the same Rust renderer the viewer uses,
 * on a Skia canvas. Register a document's fonts once with `prepare`, then render
 * any of its slides.
 */
export class HeadlessRenderer {
  private canvas = createCanvas(1, 1);
  private ctx = this.canvas.getContext("2d");
  private renderer: any;
  private fonts = new FontResolver();
  private registered = new Set<string>();
  private images = new Map<string, unknown>();

  constructor(private options: HeadlessRenderOptions = {}) {
    installNodeEnv();
    let wasm: any;
    try {
      wasm = createRequire(import.meta.url)("../../rust-engine/pkg-node/ppt_engine.js");
    } catch (error) {
      throw new Error(`找不到 Node 版渲染引擎，请先运行 npm run build:wasm:node：${error}`);
    }
    this.renderer = this.quietly(() => new wasm.RustPptRenderer(this.ctx));
    for (const file of options.fonts || []) this.registerFile(file);
  }

  /** Register the fonts `doc` uses. Safe to call again for another document. */
  public prepare(doc: PptVirtualDocument): void {
    for (const request of collectFontRequests(doc)) this.registerFile(this.fonts.resolve(request));
  }

  public async renderSlide(doc: PptVirtualDocument, index: number, options: RenderSlideOptions = {}): Promise<Buffer> {
    const slide = doc.slides[index];
    if (!slide) throw new Error(`没有第 ${index + 1} 页。`);
    const width = Math.max(1, Math.round(options.width ?? 1280));
    const scale = width / doc.size.width;
    const height = Math.max(1, Math.round(doc.size.height * scale));
    this.canvas.width = width;
    this.canvas.height = height;
    this.ctx.setTransform(scale, 0, 0, scale, 0, 0);

    const images = await this.loadImages(slide);
    this.quietly(() => this.renderer.render_slide(JSON.stringify({ id: slide.id, elements: expandForRender(slide.elements) }), images));
    return this.canvas.toBuffer("image/png");
  }

  private registerFile(file: string): void {
    if (this.registered.has(file)) return;
    this.registered.add(file);
    const bytes = new Uint8Array(readFileSync(file));
    this.quietly(() => this.renderer.register_font(bytes));
  }

  private async loadImages(slide: Slide): Promise<Record<string, unknown>> {
    const urls = new Set<string>();
    for (const element of slide.elements) {
      if (element.type === "image") urls.add(element.url);
      else if (element.type === "shape") {
        const fill = element.computedStyle?.fill;
        if (fill?.type === "picture" && fill.url) urls.add(fill.url);
      }
    }
    const result: Record<string, unknown> = {};
    for (const url of urls) {
      if (!this.images.has(url)) this.images.set(url, await this.decode(url));
      const image = this.images.get(url);
      if (image) result[url] = image;
    }
    return result;
  }

  private async decode(url: string): Promise<unknown> {
    try {
      if (url.startsWith("blob:")) {
        const blob = resolveObjectURL(url);
        return blob ? await loadImage(Buffer.from(await blob.arrayBuffer())) : null;
      }
      return await loadImage(url);
    } catch (error) {
      console.warn(`[headless] 图片无法解码，已跳过：${String(error)}`);
      return null;
    }
  }

  private quietly<T>(run: () => T): T {
    if (this.options.quiet === false) return run();
    const log = console.log;
    console.log = () => undefined;
    try { return run(); } finally { console.log = log; }
  }
}
