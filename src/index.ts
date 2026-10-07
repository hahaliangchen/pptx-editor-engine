import {
  ImageElement,
  PptVirtualDocument,
  PptxParser,
  PresentationAST,
  Slide,
  ShapeElement,
  PresentationSize
} from "./pptx-parser";
import * as wasm from "../rust-engine/pkg/ppt_engine";

export interface ViewerOptions {
  container: HTMLElement;
  width?: number; // CSS width
  height?: number; // CSS height
  fontBackendUrl?: string;
  debugTextBoxes?: boolean;
  onSlideChange?: (slideIndex: number, slide: Slide) => void;
  onLoadComplete?: (document: PptVirtualDocument) => void | Promise<void>;
}

interface FontRequest {
  family: string;
  bold: boolean;
  italic: boolean;
  eastAsian: boolean;
}

interface RenderedSlideSnapshot {
  width: number;
  height: number;
  canvas: HTMLCanvasElement;
}

export default class PptViewer {
  private container: HTMLElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private renderer: wasm.RustPptRenderer | null = null;
  private cacheCanvas: HTMLCanvasElement | null = null;
  private cacheContext: CanvasRenderingContext2D | null = null;
  private cacheRenderer: wasm.RustPptRenderer | null = null;
  private parser: PptxParser;
  private presentation: PptVirtualDocument | null = null;
  private currentSlideIndex: number = 0;
  private imageCache: Record<string, HTMLImageElement> = {};
  private registeredFontBytes: Uint8Array[] = [];
  private slideJsonCache = new Map<string, string>();
  private renderedSlideCache = new Map<string, RenderedSlideSnapshot>();
  private fontBackendUrl: string;
  private engineReady: Promise<void>;
  private fontCatalogPromise: Promise<string[]> | null = null;
  private loadedFontKeys = new Set<string>();
  private fontLoads = new Map<string, Promise<void>>();
  private preferredWidth?: number;
  private preferredHeight?: number;
  private debugTextBoxes: boolean;
  private renderEpoch = 0;
  private isRenderingThumbnails = false;
  private cacheWarmupEpoch = 0;
  private loadSequence = 0;
  private selectedElementId: string | null = null;
  private selectedSlideId: string | null = null;
  private dragState: {
    pointerId: number;
    slideId: string;
    elementId: string;
    startX: number;
    startY: number;
    originX: number;
    originY: number;
  } | null = null;

  // Callbacks
  private onSlideChange?: (slideIndex: number, slide: Slide) => void;
  private onLoadComplete?: (document: PptVirtualDocument) => void | Promise<void>;

  constructor(options: ViewerOptions) {
    this.container = options.container;
    this.onSlideChange = options.onSlideChange;
    this.onLoadComplete = options.onLoadComplete;
    this.preferredWidth = options.width;
    this.preferredHeight = options.height;
    this.debugTextBoxes = options.debugTextBoxes ?? false;
    this.fontBackendUrl = (options.fontBackendUrl || "http://127.0.0.1:8080").replace(/\/$/, "");

    // Create Canvas element
    this.canvas = document.createElement("canvas");
    this.canvas.style.width = "100%";
    this.canvas.style.height = "100%";
    this.canvas.style.display = "block";
    this.canvas.style.boxShadow = "0 8px 24px rgba(0,0,0,0.15)";
    this.canvas.style.borderRadius = "8px";
    this.canvas.style.backgroundColor = "#ffffff";
    this.container.appendChild(this.canvas);

    const context = this.canvas.getContext("2d");
    if (!context) throw new Error("Failed to get 2D canvas context");
    this.ctx = context;

    this.parser = new PptxParser();
    this.engineReady = this.initEngine();

    this.canvas.style.touchAction = "none";
    this.canvas.addEventListener("pointerdown", event => this.onCanvasPointerDown(event));
    this.canvas.addEventListener("pointermove", event => this.onCanvasPointerMove(event));
    this.canvas.addEventListener("pointerup", event => this.onCanvasPointerEnd(event));
    this.canvas.addEventListener("pointercancel", event => this.onCanvasPointerEnd(event));

    // Listen to resize to keep canvas crisp
    window.addEventListener("resize", () => this.resizeAndRedraw());
  }

  private async initEngine() {
    try {
      const wasmModule = await import("../rust-engine/pkg/ppt_engine");
      this.renderer = new wasmModule.RustPptRenderer(this.ctx);
      console.log("PPT WASM Engine initialized successfully.");

    } catch (err) {
      console.error("Failed to load Rust WASM Engine:", err);
      throw err;
    }
  }

  private async fetchFontBackend(path: string): Promise<Response> {
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), 10000);
    console.log(`[Font Backend] request ${path}`);
    try {
      const response = await fetch(`${this.fontBackendUrl}${path}`, {
        signal: controller.signal
      });
      console.log(`[Font Backend] response ${path} status=${response.status}`);
      return response;
    } catch (err) {
      if (controller.signal.aborted) {
        console.error(`[Font Backend] timeout ${path}`);
        throw new Error("字体后端请求超时，请确认 8080 端口服务正在运行。");
      }
      console.error(`[Font Backend] failed ${path}`, err);
      throw new Error("无法连接字体后端，请确认 8080 端口服务正在运行。");
    } finally {
      window.clearTimeout(timeout);
    }
  }

  private getFontCatalog(): Promise<string[]> {
    if (!this.fontCatalogPromise) {
      this.fontCatalogPromise = this.fetchFontBackend("/api/fonts")
        .then(response => {
          if (!response.ok) throw new Error(`Font backend returned HTTP ${response.status}`);
          return response.json() as Promise<string[]>;
        })
        .then(families => {
          console.log(`[Font Manager] Backend exposes ${families.length} font families.`);
          return families;
        })
        .catch(err => {
          this.fontCatalogPromise = null;
          throw err;
        });
    }
    return this.fontCatalogPromise;
  }

  private resolveBackendFamily(request: FontRequest, catalog: string[]): string {
    const byLowerName = new Map(catalog.map(family => [family.toLowerCase(), family]));
    const requested = request.family.trim();
    const exact = requested ? byLowerName.get(requested.toLowerCase()) : undefined;
    if (exact && !["sans-serif", "serif", "monospace"].includes(requested.toLowerCase())) {
      return exact;
    }

    const candidates = request.eastAsian
      ? ["Microsoft YaHei", "DengXian", "SimSun", "Noto Sans CJK SC", "Source Han Sans CN"]
      : ["Aptos", "Calibri", "Arial", "Segoe UI", "Times New Roman"];
    for (const candidate of candidates) {
      const available = byLowerName.get(candidate.toLowerCase());
      if (available) return available;
    }
    if (catalog[0]) return catalog[0];
    throw new Error("Font backend returned an empty font catalog");
  }

  private collectPresentationFonts(document: PptVirtualDocument): FontRequest[] {
    const requests = new Map<string, FontRequest>();
    const addFamily = (
      family: string | undefined,
      eastAsian: boolean,
      bold = false,
      italic = false
    ) => {
      if (!family) return;
      const request: FontRequest = { family, bold, italic, eastAsian };
      const key = `${family.toLowerCase()}|${request.bold}|${request.italic}|${eastAsian}`;
      requests.set(key, request);
    };
    const addStyle = (style: { fontFamily?: string; eastAsianFontFamily?: string; bold: boolean; italic?: boolean }) => {
      addFamily(style.fontFamily, false, style.bold, style.italic || false);
      addFamily(style.eastAsianFontFamily, true, style.bold, style.italic || false);
    };

    for (const slide of document.slides) {
      for (const element of slide.elements) {
        if (element.type !== "text") continue;
        addStyle(element.style);
        for (const paragraph of element.paragraphs || []) {
          for (const run of paragraph.runs) addStyle(run.style);
          addFamily(paragraph.bullet?.fontFamily, false);
        }
      }
    }
    if (requests.size === 0) {
      requests.set("sans-serif|false|false|false", {
        family: "sans-serif",
        bold: false,
        italic: false,
        eastAsian: false
      });
    }
    return [...requests.values()];
  }

  private loadBackendFont(family: string, bold: boolean, italic: boolean): Promise<void> {
    const weight = bold ? 700 : 400;
    const key = `${family.toLowerCase()}|${weight}|${italic}`;
    if (this.loadedFontKeys.has(key)) return Promise.resolve();
    const pending = this.fontLoads.get(key);
    if (pending) return pending;

    const load = (async () => {
      if (!this.renderer) throw new Error("WASM renderer is not initialized");
      console.log(`[Font Manager] loading ${family} (${weight}, italic=${italic})`);
      const query = new URLSearchParams({
        family,
        weight: weight.toString(),
        italic: italic.toString()
      });
      const response = await this.fetchFontBackend(`/api/font?${query}`);
      if (!response.ok) {
        throw new Error(`Font backend could not provide ${family} (${weight}, italic=${italic})`);
      }
      const fontBytes = new Uint8Array(await response.arrayBuffer());
      this.renderer.register_font(fontBytes);
      this.registeredFontBytes.push(fontBytes);
      if (this.cacheRenderer) {
        this.cacheRenderer.register_font(fontBytes);
      }
      this.loadedFontKeys.add(key);
      console.log(`[Font Manager] Registered backend font in Rust: ${family} (${weight}, italic=${italic}).`);
    })().finally(() => this.fontLoads.delete(key));

    this.fontLoads.set(key, load);
    return load;
  }

  private async ensurePresentationFonts(document: PptVirtualDocument): Promise<void> {
    const catalog = await this.getFontCatalog();
    const resolved = new Map<string, { family: string; bold: boolean; italic: boolean }>();
    const latinFallbacks = new Map<string, string>();
    const eastAsianFallbacks = new Map<string, string>();
    for (const request of this.collectPresentationFonts(document)) {
      const family = this.resolveBackendFamily(request, catalog);
      (request.eastAsian ? eastAsianFallbacks : latinFallbacks)
        .set(request.family.toLowerCase(), family);
      if (family.toLowerCase() !== request.family.toLowerCase()) {
        console.warn(`[Font Manager] ${request.family} is unavailable; backend fallback is ${family}.`);
      }
      resolved.set(`${family.toLowerCase()}|${request.bold}|${request.italic}`, {
        family,
        bold: request.bold,
        italic: request.italic
      });
    }

    const applyBackendFamily = (style: { fontFamily?: string; eastAsianFontFamily?: string }) => {
      if (style.fontFamily) {
        style.fontFamily = latinFallbacks.get(style.fontFamily.toLowerCase()) || style.fontFamily;
      }
      if (style.eastAsianFontFamily) {
        style.eastAsianFontFamily = eastAsianFallbacks.get(style.eastAsianFontFamily.toLowerCase())
          || style.eastAsianFontFamily;
      }
    };
    for (const slide of document.slides) {
      for (const element of slide.elements) {
        if (element.type !== "text") continue;
        applyBackendFamily(element.style);
        for (const paragraph of element.paragraphs || []) {
          for (const run of paragraph.runs) applyBackendFamily(run.style);
          if (paragraph.bullet?.fontFamily) {
            paragraph.bullet.fontFamily = latinFallbacks.get(paragraph.bullet.fontFamily.toLowerCase())
              || paragraph.bullet.fontFamily;
          }
        }
      }
    }

    await Promise.all([...resolved.values()].map(font =>
      this.loadBackendFont(font.family, font.bold, font.italic)
    ));
  }

  // Load PPTX file from ArrayBuffer
  public async loadPptx(buffer: ArrayBuffer): Promise<PptVirtualDocument> {
    const loadId = ++this.loadSequence;
    const startedAt = performance.now();
    const log = (message: string, ...args: unknown[]) => {
      console.log(
        `[PPTX Load #${loadId}] +${Math.round(performance.now() - startedAt)}ms ${message}`,
        ...args
      );
    };
    log(`start bytes=${buffer.byteLength}`);

    try {
      log("parse:start");
      this.presentation = await this.parser.parse(buffer);
      this.selectedElementId = null;
      this.selectedSlideId = null;
      this.dragState = null;
      log(`parse:done slides=${this.presentation.slides.length}`);

      log("wasm:init:wait");
      await this.engineReady;
      log("wasm:init:ready");

      log("fonts:start");
      await this.ensurePresentationFonts(this.presentation);
      log("fonts:done");

      this.currentSlideIndex = 0;
      this.imageCache = {}; // Clear old image cache
      this.slideJsonCache.clear();
      this.renderedSlideCache.clear();

      log("first-slide:start");
      await this.renderCurrentSlide();
      log("first-slide:done");

      if (this.onLoadComplete) {
        log("onLoadComplete:start");
        await this.onLoadComplete(this.presentation);
        log("onLoadComplete:done");
      }

      log("complete");
      return this.presentation;
    } catch (error) {
      log("failed", error);
      throw error;
    }
  }

  // Load a pre-built PPT Virtual DOM (for demos / testing).
  public async loadVirtualDocument(document: PptVirtualDocument): Promise<void> {
    this.presentation = document;
    this.selectedElementId = null;
    this.selectedSlideId = null;
    this.dragState = null;
    await this.engineReady;
    await this.ensurePresentationFonts(document);
    this.currentSlideIndex = 0;
    this.imageCache = {};
    this.slideJsonCache.clear();
    this.renderedSlideCache.clear();

    // Paint the active slide first so it is on screen, then run any
    // onLoadComplete work (e.g. thumbnail generation, which repurposes the
    // shared canvas) with the main canvas already in its final state.
    await this.renderCurrentSlide();
    if (this.onLoadComplete) {
      await this.onLoadComplete(this.presentation);
    }
  }

  /**
   * Render every slide to an offscreen-sized bitmap and return PNG data URLs,
   * used to build real page thumbnails for the sidebar filmstrip. The Rust
   * renderer is bound to a single 2D context, so this temporarily repurposes
   * the main canvas, then restores and repaints the active slide.
   */
  public async renderThumbnails(targetCssWidth = 240): Promise<string[]> {
    if (!this.presentation || !this.renderer) return [];
    await this.engineReady;

    const startedAt = performance.now();
    console.log(`[Thumbnails] start slides=${this.presentation.slides.length}`);

    const logicalSize = this.presentation.size;
    const ratio = logicalSize.width / logicalSize.height;
    const dpr = window.devicePixelRatio || 1;
    const pixelWidth = Math.max(1, Math.round(targetCssWidth * dpr));
    const pixelHeight = Math.max(1, Math.round(pixelWidth / ratio));

    const savedWidth = this.canvas.width;
    const savedHeight = this.canvas.height;

    // Block any concurrent resizeAndRedraw (e.g. from a window resize event)
    // from clobbering the shared canvas while it is repurposed for thumbnails.
    this.isRenderingThumbnails = true;
    // Hide the working canvas while it is repurposed for thumbnail rasters.
    this.canvas.style.visibility = "hidden";
    this.canvas.width = pixelWidth;
    this.canvas.height = pixelHeight;

    const scale = Math.min(pixelWidth / logicalSize.width, pixelHeight / logicalSize.height);
    const thumbnails: string[] = [];

    try {
      for (const slide of this.presentation.slides) {
        console.log(`[Thumbnails] render:start slide=${slide.id}`);
        // Preload images referenced by this slide so the thumbnail is complete.
        await this.preloadSlideImages(slide, "thumbnail");

        this.ctx.setTransform(scale, 0, 0, scale, 0, 0);
        try {
          this.renderer.render_slide(this.serializeSlideForRenderer(slide), this.imageCache);
          thumbnails.push(this.canvas.toDataURL("image/png"));
          console.log(`[Thumbnails] render:done slide=${slide.id}`);
        } catch (err) {
          console.error("Error while rendering thumbnail:", err);
          thumbnails.push("");
        }
      }
    } finally {
      // Restore the main canvas and repaint the active slide, then release the
      // resize guard even if a slide threw mid-loop.
      this.canvas.width = savedWidth;
      this.canvas.height = savedHeight;
      this.canvas.style.visibility = "visible";
      this.renderedSlideCache.clear();
      this.isRenderingThumbnails = false;
      await this.renderCurrentSlide();
    }

    console.log(`[Thumbnails] complete +${Math.round(performance.now() - startedAt)}ms`);
    return thumbnails;
  }

  /** @deprecated Use loadVirtualDocument. */
  public async loadAST(ast: PresentationAST): Promise<void> {
    await this.loadVirtualDocument({
      ...ast,
      styleRegistry: ast.styleRegistry || { rules: {} }
    });
  }

  public getSlidesCount(): number {
    return this.presentation?.slides.length || 0;
  }

  public getCurrentSlideIndex(): number {
    return this.currentSlideIndex;
  }

  public getPresentationSize(): PresentationSize | null {
    return this.presentation?.size || null;
  }

  public async selectSlide(index: number): Promise<void> {
    if (!this.presentation || index < 0 || index >= this.presentation.slides.length) return;
    this.currentSlideIndex = index;
    await this.renderCurrentSlide();
  }

  public async nextSlide(): Promise<void> {
    if (!this.presentation) return;
    if (this.currentSlideIndex < this.presentation.slides.length - 1) {
      this.currentSlideIndex++;
      await this.renderCurrentSlide();
    }
  }

  public async prevSlide(): Promise<void> {
    if (!this.presentation) return;
    if (this.currentSlideIndex > 0) {
      this.currentSlideIndex--;
      await this.renderCurrentSlide();
    }
  }

  public getCurrentSlideAST(): Slide | null {
    if (!this.presentation || this.currentSlideIndex < 0 || this.currentSlideIndex >= this.presentation.slides.length) {
      return null;
    }
    return this.presentation.slides[this.currentSlideIndex];
  }

  private getSlidePoint(event: PointerEvent): { x: number; y: number } | null {
    if (!this.presentation) return null;
    const bounds = this.canvas.getBoundingClientRect();
    if (bounds.width <= 0 || bounds.height <= 0) return null;

    const logicalSize = this.presentation.size;
    const scale = Math.min(
      this.canvas.width / logicalSize.width,
      this.canvas.height / logicalSize.height
    );
    if (!Number.isFinite(scale) || scale <= 0) return null;

    const canvasX = (event.clientX - bounds.left) * this.canvas.width / bounds.width;
    const canvasY = (event.clientY - bounds.top) * this.canvas.height / bounds.height;
    const offsetX = (this.canvas.width - logicalSize.width * scale) / 2;
    const offsetY = (this.canvas.height - logicalSize.height * scale) / 2;
    return {
      x: (canvasX - offsetX) / scale,
      y: (canvasY - offsetY) / scale
    };
  }

  private onCanvasPointerDown(event: PointerEvent): void {
    const slide = this.getCurrentSlideAST();
    const point = this.getSlidePoint(event);
    if (!slide || !point) return;
    event.preventDefault();

    const size = this.presentation!.size;
    const withinSlide = point.x >= 0 && point.y >= 0
      && point.x <= size.width && point.y <= size.height;
    const element = withinSlide
      ? [...slide.elements].reverse().find(candidate =>
        point.x >= candidate.rect.x
        && point.x <= candidate.rect.x + candidate.rect.w
        && point.y >= candidate.rect.y
        && point.y <= candidate.rect.y + candidate.rect.h
      )
      : undefined;

    if (!element) {
      this.selectedElementId = null;
      this.selectedSlideId = null;
      this.dragState = null;
      this.resizeAndRedraw(slide);
      return;
    }

    this.selectedElementId = element.id;
    this.selectedSlideId = slide.id;
    this.dragState = {
      pointerId: event.pointerId,
      slideId: slide.id,
      elementId: element.id,
      startX: point.x,
      startY: point.y,
      originX: element.rect.x,
      originY: element.rect.y
    };
    this.canvas.setPointerCapture(event.pointerId);
    this.resizeAndRedraw(slide);
  }

  private onCanvasPointerMove(event: PointerEvent): void {
    const drag = this.dragState;
    if (!drag || event.pointerId !== drag.pointerId) return;
    const slide = this.presentation?.slides.find(candidate => candidate.id === drag.slideId);
    const element = slide?.elements.find(candidate => candidate.id === drag.elementId);
    const point = this.getSlidePoint(event);
    if (!slide || !element || !point) return;

    const x = drag.originX + point.x - drag.startX;
    const y = drag.originY + point.y - drag.startY;
    if (element.rect.x === x && element.rect.y === y) return;
    element.rect.x = x;
    element.rect.y = y;
    this.slideJsonCache.delete(slide.id);
    this.renderedSlideCache.delete(slide.id);
    this.resizeAndRedraw(slide);
  }

  private onCanvasPointerEnd(event: PointerEvent): void {
    if (this.dragState?.pointerId === event.pointerId) {
      this.dragState = null;
    }
  }

  /** Replace the plain text of a top-level text element on a slide. */
  public async updateTextElement(slideId: string, elementId: string, content: string): Promise<void> {
    if (!this.presentation) {
      throw new Error("No presentation is loaded.");
    }

    const slide = this.presentation.slides.find(candidate => candidate.id === slideId);
    if (!slide) {
      throw new Error(`Slide with id "${slideId}" was not found.`);
    }

    const element = slide.elements.find(candidate => candidate.id === elementId);
    if (!element) {
      throw new Error(`Element with id "${elementId}" was not found on slide "${slideId}".`);
    }
    if (element.type !== "text") {
      throw new Error(`Element "${elementId}" on slide "${slideId}" is not a text element.`);
    }

    const firstParagraph = element.paragraphs?.[0];
    const paragraphStyle = firstParagraph?.style ?? {
      align: element.style.align,
      level: 0,
      marginLeft: 0,
      indent: 0
    };
    const runStyle = firstParagraph?.runs[0]?.style ?? element.style;

    element.content = content;
    element.paragraphs = [{
      style: paragraphStyle,
      runs: [{ content, style: runStyle }],
      ...(firstParagraph?.bullet ? { bullet: firstParagraph.bullet } : {})
    }];

    this.slideJsonCache.delete(slide.id);
    this.renderedSlideCache.delete(slide.id);

    if (this.getCurrentSlideAST() === slide) {
      await this.renderCurrentSlide();
    }
  }

  public setDebugTextBoxes(enabled: boolean): void {
    if (this.debugTextBoxes === enabled) return;
    this.debugTextBoxes = enabled;
    this.resizeAndRedraw();
  }

  public isDebugTextBoxesEnabled(): boolean {
    return this.debugTextBoxes;
  }

  private async renderCurrentSlide() {
    if (!this.presentation || !this.renderer) return;

    const slide = this.getCurrentSlideAST();
    if (!slide) return;
    const renderEpoch = ++this.renderEpoch;

    // Trigger callback
    if (this.onSlideChange) {
      this.onSlideChange(this.currentSlideIndex, slide);
    }

    // 1. Preload all images in the slide
    await this.preloadSlideImages(slide, "slide");

    // A rapid click sequence may finish image loading out of order. Drop a
    // stale request so it cannot render over the newest selected slide.
    if (renderEpoch !== this.renderEpoch) return;

    // 2. Setup canvas dimensions and scale
    this.resizeAndRedraw(slide);
  }

  private async preloadSlideImages(slide: Slide, context: "slide" | "thumbnail" | "cache") {
    const imageElements = slide.elements.filter(el => el.type === "image") as ImageElement[];
    const shapeFillUrls = slide.elements
      .filter((el): el is ShapeElement => el.type === "shape")
      .map(el => el.computedStyle?.fill)
      .flatMap(fill => fill?.type === "picture" && fill.url ? [fill.url] : []);
    const imageUrls = [
      ...imageElements.map(img => img.url),
      ...shapeFillUrls
    ];
    await Promise.all(imageUrls.map(async (url) => {
      if (this.imageCache[url]) return;
      try {
        this.imageCache[url] = await this.loadImage(url);
      } catch (err) {
        console.error(`Failed to preload ${context} image: ${url}`, err);
      }
    }));
  }

  private serializeSlideForRenderer(slide: Slide): string {
    const cached = this.slideJsonCache.get(slide.id);
    if (cached) return cached;

    // rawXml is only for the inspector. Rust deserializes id/elements and
    // should not receive the complete source XML on every page switch.
    const json = JSON.stringify({ id: slide.id, elements: slide.elements });
    this.slideJsonCache.set(slide.id, json);
    return json;
  }

  private cacheRenderedSlide(
    slide: Slide,
    sourceCanvas: HTMLCanvasElement = this.canvas
  ): void {
    const snapshot = document.createElement("canvas");
    snapshot.width = sourceCanvas.width;
    snapshot.height = sourceCanvas.height;
    const snapshotContext = snapshot.getContext("2d");
    if (!snapshotContext) return;
    snapshotContext.drawImage(sourceCanvas, 0, 0);
    this.renderedSlideCache.set(slide.id, {
      width: sourceCanvas.width,
      height: sourceCanvas.height,
      canvas: snapshot
    });
  }

  private ensureCacheRenderer(width: number, height: number): void {
    if (
      this.cacheRenderer
      && this.cacheCanvas
      && this.cacheCanvas.width === width
      && this.cacheCanvas.height === height
    ) {
      return;
    }

    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Failed to create the background slide cache context");

    const renderer = new wasm.RustPptRenderer(context);
    for (const fontBytes of this.registeredFontBytes) {
      renderer.register_font(fontBytes);
    }
    this.cacheCanvas = canvas;
    this.cacheContext = context;
    this.cacheRenderer = renderer;
  }

  private scheduleSlideCacheWarmup(): void {
    if (!this.presentation || !this.renderer || this.isRenderingThumbnails) return;

    const epoch = ++this.cacheWarmupEpoch;
    window.setTimeout(() => {
      void this.warmNextSlide(epoch);
    }, 0);
  }

  private async warmNextSlide(epoch: number): Promise<void> {
    if (
      epoch !== this.cacheWarmupEpoch
      || !this.presentation
      || !this.renderer
      || this.isRenderingThumbnails
    ) {
      return;
    }

    const width = this.canvas.width;
    const height = this.canvas.height;
    if (width < 1 || height < 1) return;

    const logicalSize = this.presentation.size;
    const scale = Math.min(width / logicalSize.width, height / logicalSize.height);
    const offsetX = (width - logicalSize.width * scale) / 2;
    const offsetY = (height - logicalSize.height * scale) / 2;
    const orderedIndices = this.presentation.slides
      .map((_, index) => index)
      .sort((a, b) => {
        const distanceA = Math.abs(a - this.currentSlideIndex);
        const distanceB = Math.abs(b - this.currentSlideIndex);
        return distanceA - distanceB;
      });
    const slide = orderedIndices
      .map(index => this.presentation?.slides[index])
      .find(candidate => {
        if (!candidate) return false;
        const cached = this.renderedSlideCache.get(candidate.id);
        return !cached || cached.width !== width || cached.height !== height;
      });
    if (!slide) return;

    this.ensureCacheRenderer(width, height);
    await this.preloadSlideImages(slide, "cache");
    if (
      epoch !== this.cacheWarmupEpoch
      || !this.presentation
      || !this.cacheRenderer
      || !this.cacheContext
    ) {
      return;
    }

    this.cacheContext.setTransform(scale, 0, 0, scale, offsetX, offsetY);
    this.cacheRenderer.render_slide(this.serializeSlideForRenderer(slide), this.imageCache);
    this.cacheRenderedSlide(slide, this.cacheCanvas!);

    // Yield between pages so a large deck does not monopolize the main thread.
    window.setTimeout(() => {
      void this.warmNextSlide(epoch);
    }, 0);
  }

  private resizeAndRedraw(slideOverride?: Slide) {
    if (!this.presentation || !this.renderer) return;
    // The shared canvas is temporarily resized for thumbnail rasters; ignore
    // resize-driven repaints until that finishes to avoid clobbering it.
    if (this.isRenderingThumbnails) return;

    const slide = slideOverride || this.getCurrentSlideAST();
    if (!slide) return;

    const logicalSize = this.presentation.size;
    const slideRatio = logicalSize.width / logicalSize.height;

    const viewport = this.container.parentElement;
    const viewportStyle = viewport ? getComputedStyle(viewport) : null;
    const horizontalPadding = viewportStyle
      ? parseFloat(viewportStyle.paddingLeft) + parseFloat(viewportStyle.paddingRight)
      : 0;
    const verticalPadding = viewportStyle
      ? parseFloat(viewportStyle.paddingTop) + parseFloat(viewportStyle.paddingBottom)
      : 0;
    const controls = viewport?.querySelector<HTMLElement>(".control-bar");
    const controlsHeight = controls ? controls.offsetHeight + 16 : 0;
    const availableWidth = Math.max(
      1,
      Math.min(
        (viewport?.clientWidth || this.container.clientWidth || logicalSize.width) - horizontalPadding,
        this.preferredWidth || Number.POSITIVE_INFINITY
      )
    );
    const availableHeight = Math.max(
      1,
      Math.min(
        (viewport?.clientHeight || this.container.clientHeight || logicalSize.height)
          - verticalPadding
          - controlsHeight,
        this.preferredHeight || Number.POSITIVE_INFINITY
      )
    );
    const cssWidth = Math.min(availableWidth, availableHeight * slideRatio);
    const cssHeight = cssWidth / slideRatio;

    this.container.style.width = `${cssWidth}px`;
    this.container.style.height = `${cssHeight}px`;
    this.container.style.aspectRatio = `${logicalSize.width} / ${logicalSize.height}`;

    // Device Pixel Ratio scaling for Retina screens (sharp rendering)
    const dpr = window.devicePixelRatio || 1;
    const pixelWidth = Math.max(1, Math.round(cssWidth * dpr));
    const pixelHeight = Math.max(1, Math.round(cssHeight * dpr));
    if (this.canvas.width !== pixelWidth || this.canvas.height !== pixelHeight) {
      this.canvas.width = pixelWidth;
      this.canvas.height = pixelHeight;
      this.renderedSlideCache.clear();
    }

    this.ctx.save();
    
    // Always use a uniform scale. Any spare pixels become centered letterboxing.
    const scale = Math.min(
      this.canvas.width / logicalSize.width,
      this.canvas.height / logicalSize.height
    );
    const offsetX = (this.canvas.width - logicalSize.width * scale) / 2;
    const offsetY = (this.canvas.height - logicalSize.height * scale) / 2;
    const cached = this.renderedSlideCache.get(slide.id);
    const hasCachedSlide = cached
      && cached.width === this.canvas.width
      && cached.height === this.canvas.height;

    if (hasCachedSlide) {
      this.ctx.setTransform(1, 0, 0, 1, 0, 0);
      this.ctx.drawImage(cached.canvas, 0, 0);
      this.ctx.setTransform(scale, 0, 0, scale, offsetX, offsetY);
    } else {
      this.ctx.setTransform(scale, 0, 0, scale, offsetX, offsetY);

      try {
        this.renderer.render_slide(this.serializeSlideForRenderer(slide), this.imageCache);
        this.cacheRenderedSlide(slide);
      } catch (err) {
        console.error("Error during Rust rendering:", err);
      }
    }

    if (this.debugTextBoxes) {
      this.drawDebugTextBoxes(slide);
    }
    this.drawSelectionBox(slide);

    this.ctx.restore();
    this.scheduleSlideCacheWarmup();
  }

  private drawDebugTextBoxes(slide: Slide): void {
    this.ctx.save();
    this.ctx.strokeStyle = "#000000";
    const transform = this.ctx.getTransform();
    const deviceScale = Math.hypot(transform.a, transform.b);
    this.ctx.lineWidth = 1 / Math.max(deviceScale, 0.1);
    this.ctx.setLineDash([5, 3]);
    for (const element of slide.elements) {
      if (element.type !== "text" || !element.content.trim()) continue;
      this.ctx.strokeRect(
        element.rect.x,
        element.rect.y,
        element.rect.w,
        element.rect.h
      );
    }
    this.ctx.restore();
  }

  private drawSelectionBox(slide: Slide): void {
    if (this.selectedSlideId !== slide.id || !this.selectedElementId) return;
    const element = slide.elements.find(candidate => candidate.id === this.selectedElementId);
    if (!element) return;

    this.ctx.save();
    const transform = this.ctx.getTransform();
    const scale = Math.max(Math.hypot(transform.a, transform.b), 0.1);
    this.ctx.strokeStyle = "#1683ff";
    this.ctx.lineWidth = 2 / scale;
    this.ctx.setLineDash([]);
    this.ctx.strokeRect(
      element.rect.x,
      element.rect.y,
      element.rect.w,
      element.rect.h
    );
    this.ctx.restore();
  }

  private loadImage(url: string): Promise<HTMLImageElement> {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = () => resolve(img);
      img.onerror = (e) => reject(new Error(`Failed to load image at ${url}`));
      img.src = url;
    });
  }
}
