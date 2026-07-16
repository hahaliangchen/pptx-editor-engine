import {
  ImageElement,
  PptVirtualDocument,
  PptxParser,
  PresentationAST,
  Slide,
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
  onLoadComplete?: (document: PptVirtualDocument) => void;
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
  private parser: PptxParser;
  private presentation: PptVirtualDocument | null = null;
  private currentSlideIndex: number = 0;
  private imageCache: Record<string, HTMLImageElement> = {};
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
  
  // Callbacks
  private onSlideChange?: (slideIndex: number, slide: Slide) => void;
  private onLoadComplete?: (document: PptVirtualDocument) => void;

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

  private getFontCatalog(): Promise<string[]> {
    if (!this.fontCatalogPromise) {
      this.fontCatalogPromise = fetch(`${this.fontBackendUrl}/api/fonts`)
        .then(response => {
          if (!response.ok) throw new Error(`Font backend returned HTTP ${response.status}`);
          return response.json() as Promise<string[]>;
        })
        .then(families => {
          console.log(`[Font Manager] Backend exposes ${families.length} font families.`);
          return families;
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
      const query = new URLSearchParams({
        family,
        weight: weight.toString(),
        italic: italic.toString()
      });
      const response = await fetch(`${this.fontBackendUrl}/api/font?${query}`);
      if (!response.ok) {
        throw new Error(`Font backend could not provide ${family} (${weight}, italic=${italic})`);
      }
      const buffer = await response.arrayBuffer();
      this.renderer.register_font(new Uint8Array(buffer));
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
    this.presentation = await this.parser.parse(buffer);
    await this.engineReady;
    await this.ensurePresentationFonts(this.presentation);
    this.currentSlideIndex = 0;
    this.imageCache = {}; // Clear old image cache
    this.slideJsonCache.clear();
    this.renderedSlideCache.clear();
    
    if (this.onLoadComplete) {
      this.onLoadComplete(this.presentation);
    }

    await this.renderCurrentSlide();
    return this.presentation;
  }

  // Load a pre-built PPT Virtual DOM (for demos / testing).
  public async loadVirtualDocument(document: PptVirtualDocument): Promise<void> {
    this.presentation = document;
    await this.engineReady;
    await this.ensurePresentationFonts(document);
    this.currentSlideIndex = 0;
    this.imageCache = {};
    this.slideJsonCache.clear();
    this.renderedSlideCache.clear();
    
    if (this.onLoadComplete) {
      this.onLoadComplete(this.presentation);
    }
    
    await this.renderCurrentSlide();
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
    const imageElements = slide.elements.filter(el => el.type === "image") as ImageElement[];
    const loadPromises = imageElements.map(async (img) => {
      if (this.imageCache[img.url]) return;
      try {
        const loadedImg = await this.loadImage(img.url);
        this.imageCache[img.url] = loadedImg;
      } catch (err) {
        console.error(`Failed to preload image: ${img.url}`, err);
      }
    });

    await Promise.all(loadPromises);

    // A rapid click sequence may finish image loading out of order. Drop a
    // stale request so it cannot render over the newest selected slide.
    if (renderEpoch !== this.renderEpoch) return;

    // 2. Setup canvas dimensions and scale
    this.resizeAndRedraw(slide);
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

  private cacheRenderedSlide(slide: Slide): void {
    const snapshot = document.createElement("canvas");
    snapshot.width = this.canvas.width;
    snapshot.height = this.canvas.height;
    const snapshotContext = snapshot.getContext("2d");
    if (!snapshotContext) return;
    snapshotContext.drawImage(this.canvas, 0, 0);
    this.renderedSlideCache.set(slide.id, {
      width: this.canvas.width,
      height: this.canvas.height,
      canvas: snapshot
    });
  }

  private resizeAndRedraw(slideOverride?: Slide) {
    if (!this.presentation || !this.renderer) return;

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

    this.ctx.restore();
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
