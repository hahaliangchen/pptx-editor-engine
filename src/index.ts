import {
  ImageElement,
  PptVirtualDocument,
  PptxParser,
  PresentationAST,
  Slide,
  SlideElement,
  TextElement,
  ShapeElement,
  PresentationSize
} from "./pptx-parser";
import * as wasm from "../rust-engine/pkg/ppt_engine";
import { EditorState, SlideEditor } from "./editor/slide-editor";
import { Command } from "./editor/history";
import { blankSlideXml, prepareSlideCopy, relationshipsXml } from "./editor/slides";
import {
  DirtyShapes,
  TextEdits,
  exportEditedPptx,
  markDirty,
  RelAdd,
  StructureEdits,
  StyleEdits,
  emptyStructure,
  materializeShape,
  materializeSlide,
  readRelationships,
  popOp,
  popTextOp,
  pushOp,
  pushTextOp
} from "./pptx-exporter";

export interface ViewerOptions {
  container: HTMLElement;
  width?: number; // CSS width
  height?: number; // CSS height
  fontBackendUrl?: string;
  debugTextBoxes?: boolean;
  onSlideChange?: (slideIndex: number, slide: Slide) => void;
  onLoadComplete?: (document: PptVirtualDocument) => void | Promise<void>;
  /** Fired when selection or undo/redo availability changes. */
  onEditStateChange?: (state: EditorState) => void;
  /** Fired when slides were added, removed or reordered (also by undo/redo). */
  onSlidesChanged?: (document: PptVirtualDocument, currentIndex: number) => void;
  /** Fired when the zoom changes; `fit` is true while the slide follows the window size. */
  onZoomChange?: (percent: number, fit: boolean) => void;
  /** Fired (debounced) after an edit so thumbnails can be refreshed. */
  onSlideEdited?: (slideIndex: number, slide: Slide) => void;
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
  private editor: SlideEditor;
  private dragRenderer: wasm.RustPptRenderer | null = null;
  private dragScratch: HTMLCanvasElement | null = null;
  private dragContext: CanvasRenderingContext2D | null = null;
  private dragLayers: {
    slide: Slide;
    moving: SlideElement[];
    below: HTMLCanvasElement;
    above: HTMLCanvasElement;
    scale: number;
    offsetX: number;
    offsetY: number;
  } | null = null;
  private dirtyShapes: DirtyShapes = new Map();
  private textEdits: TextEdits = new Map();
  private styleEdits: StyleEdits = new Map();
  private structureEdits: StructureEdits = emptyStructure();
  private shapeIdCounters = new Map<string, number>();
  private hiddenElements = new Set<SlideElement>();
  private editNotifyTimer = 0;
  private onSlideEdited?: (slideIndex: number, slide: Slide) => void;
  private onSlidesChanged?: (document: PptVirtualDocument, currentIndex: number) => void;
  private onZoomChange?: (percent: number, fit: boolean) => void;
  /** "fit" follows the window; a number is a percentage where 100 is the slide's real size. */
  private zoomMode: "fit" | number = "fit";
  private zoomPercent = 100;
  private slideSerial = 0;
  private originalSlideParts: Array<string | undefined> = [];

  // Callbacks
  private onSlideChange?: (slideIndex: number, slide: Slide) => void;
  private onLoadComplete?: (document: PptVirtualDocument) => void | Promise<void>;

  constructor(options: ViewerOptions) {
    this.container = options.container;
    this.onSlideChange = options.onSlideChange;
    this.onLoadComplete = options.onLoadComplete;
    this.onSlideEdited = options.onSlideEdited;
    this.onSlidesChanged = options.onSlidesChanged;
    this.onZoomChange = options.onZoomChange;
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

    this.container.style.touchAction = "none";
    this.editor = new SlideEditor({
      container: this.container,
      getSlide: () => this.getCurrentSlideAST(),
      getSlideSize: () => this.presentation?.size || null,
      invalidate: slide => this.invalidateSlide(slide),
      beginDrag: (slide, elements) => this.beginDragLayers(slide, elements),
      endDrag: () => this.endDragLayers(),
      showSlide: slide => {
        const index = this.presentation?.slides.indexOf(slide) ?? -1;
        if (index >= 0) this.currentSlideIndex = index;
      },
      onCommit: slide => this.notifySlideEdited(slide),
      markDirty: source => markDirty(this.dirtyShapes, source),
      beginTextEdit: (slide, element) => this.setHiddenElements(slide, [element]),
      endTextEdit: slide => this.setHiddenElements(slide, []),
      recordTextOp: (source, op) => pushTextOp(this.textEdits, source, op),
      discardTextOp: source => popTextOp(this.textEdits, source),
      structure: () => this.structureEdits,
      materialize: source => this.materializeSource(source),
      allocateShapeId: part => this.allocateShapeId(part),
      deckFont: () => this.deckFont(),
      allocateRelationshipId: part => this.allocateRelationshipId(part),
      allocateMediaPath: extension => this.allocateMediaPath(extension),
      resolveRelationships: (target, source, relationships) => this.resolveRelationships(target, source, relationships),
      recordStyleOp: (source, op) => pushOp(this.styleEdits, source, op),
      discardStyleOp: source => popOp(this.styleEdits, source)
    });
    if (options.onEditStateChange) this.editor.subscribe(options.onEditStateChange);

    // Listen to resize to keep canvas crisp
    window.addEventListener("resize", () => this.resizeAndRedraw());

    // Ctrl + wheel (and trackpad pinch) zooms around the pointer; plain wheel scrolls the stage.
    (this.container.parentElement ?? this.container).addEventListener("wheel", event => {
      if (!(event.ctrlKey || event.metaKey) || !this.presentation) return;
      event.preventDefault();
      this.setZoom(this.zoomPercent * Math.exp(-event.deltaY * 0.0015), { x: event.clientX, y: event.clientY });
    }, { passive: false });
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
      try {
        // The in-place text editor is DOM based, so it needs the same font.
        const face = new FontFace(family, fontBytes.slice().buffer, {
          weight: String(weight),
          style: italic ? "italic" : "normal"
        });
        document.fonts.add(await face.load());
      } catch (err) {
        console.warn(`[Font Manager] DOM font registration failed for ${family}`, err);
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
      this.resetEditing();
      this.zoomMode = "fit";
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
    this.parser.discardPackage();
    this.resetEditing();
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
  public async renderThumbnails(targetCssWidth = 240, only?: number | number[]): Promise<string[]> {
    const wanted = only === undefined ? null : new Set(Array.isArray(only) ? only : [only]);
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
      for (const [slideIndex, slide] of this.presentation.slides.entries()) {
        if (wanted && !wanted.has(slideIndex)) {
          thumbnails.push("");
          continue;
        }
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

  public getEditorState(): EditorState {
    return this.editor.getState();
  }

  public undo(): void {
    this.editor.undo();
  }

  public redo(): void {
    this.editor.redo();
  }

  public deleteSelection(): void {
    this.editor.deleteSelection();
  }

  public applyShapeStyle(op: Parameters<SlideEditor["applyShapeStyle"]>[0]): void {
    this.editor.applyShapeStyle(op);
  }

  public getZoom(): { percent: number; fit: boolean } {
    return { percent: this.zoomPercent, fit: this.zoomMode === "fit" };
  }

  /**
   * Zoom to a percentage (100 = real size) or back to "fit". With an anchor
   * (client coordinates) the point under it stays put while the slide grows.
   */
  public setZoom(mode: "fit" | number, anchor?: { x: number; y: number }): void {
    if (!this.presentation) return;
    const next = mode === "fit" ? "fit" : Math.min(400, Math.max(10, mode));
    const stage = this.container.parentElement;
    const before = this.container.getBoundingClientRect();
    const fraction = anchor && before.width > 0
      ? { x: (anchor.x - before.left) / before.width, y: (anchor.y - before.top) / before.height }
      : null;

    this.zoomMode = next;
    this.renderedSlideCache.clear();
    this.resizeAndRedraw();

    if (stage && anchor && fraction) {
      const after = this.container.getBoundingClientRect();
      stage.scrollLeft += after.left + fraction.x * after.width - anchor.x;
      stage.scrollTop += after.top + fraction.y * after.height - anchor.y;
      this.editor.layout();
    }
  }

  /** One notch up or down the usual zoom ladder. */
  public zoomStep(direction: 1 | -1): void {
    const ladder = [25, 50, 66, 75, 100, 125, 150, 200, 300, 400];
    const current = this.zoomPercent;
    const target = direction > 0
      ? ladder.find(step => step > current + 0.5)
      : [...ladder].reverse().find(step => step < current - 0.5);
    if (target !== undefined) this.setZoom(target);
  }

  public setTool(tool: Parameters<SlideEditor["setTool"]>[0]): void {
    this.editor.setTool(tool);
  }

  /** Place a picture file on the current slide. */
  public async insertImage(file: File): Promise<void> {
    const types: Record<string, { extension: string; contentType: string }> = {
      "image/png": { extension: "png", contentType: "image/png" },
      "image/jpeg": { extension: "jpeg", contentType: "image/jpeg" },
      "image/gif": { extension: "gif", contentType: "image/gif" }
    };
    const kind = types[file.type];
    if (!kind) throw new Error("仅支持 PNG、JPEG、GIF 图片。");
    const url = URL.createObjectURL(file);
    const size = await new Promise<{ width: number; height: number }>((resolve, reject) => {
      const probe = new Image();
      probe.onload = () => resolve({ width: probe.naturalWidth, height: probe.naturalHeight });
      probe.onerror = () => reject(new Error("无法读取这张图片。"));
      probe.src = url;
    });
    // Pixels become slide units at 96 dpi: 1 px = 9525 EMU.
    const emuPerPx = this.presentation?.slides[0]?.emuPerPx || 6350;
    await this.editor.insertImage({
      bytes: new Uint8Array(await file.arrayBuffer()),
      ...kind,
      url,
      width: (size.width * 9525) / emuPerPx,
      height: (size.height * 9525) / emuPerPx,
      name: file.name
    });
  }

  public arrange(mode: Parameters<SlideEditor["arrange"]>[0]): void {
    this.editor.arrange(mode);
  }

  public copy(): Promise<boolean> {
    return this.editor.copy();
  }

  public cut(): Promise<void> {
    return this.editor.cut();
  }

  public paste(): Promise<void> {
    return this.editor.paste();
  }

  public duplicate(): Promise<void> {
    return this.editor.duplicate();
  }

  public setBounds(change: Parameters<SlideEditor["setBounds"]>[0], keepRatio = false): void {
    this.editor.setBounds(change, keepRatio);
  }

  public setRotation(degrees: number): void {
    this.editor.setRotationDegrees(degrees);
  }

  public alignParagraphs(align: "left" | "center" | "right"): void {
    this.editor.alignParagraphs(align);
  }

  public alignToSlide(edge: Parameters<SlideEditor["alignToSlide"]>[0]): void {
    this.editor.alignToSlide(edge);
  }

  /** True when edits can be written back to an opened .pptx package. */
  public canExport(): boolean {
    return this.parser.getZip() !== null;
  }

  public hasUnsavedEdits(): boolean {
    return this.dirtyShapes.size > 0 || this.textEdits.size > 0 || this.styleEdits.size > 0
      || this.structureEdits.orders.size > 0 || this.structureEdits.inserts.size > 0
      || this.structureEdits.media.length > 0
      || (this.presentation?.slides || []).some((slide, index) => slide.part !== this.originalSlideParts[index])
      || (this.presentation?.slides.length ?? 0) !== this.originalSlideParts.length;
  }

  /** Build a new .pptx with every edit written back into the original package. */
  public async exportPptx(): Promise<Blob> {
    const zip = this.parser.getZip();
    if (!zip || !this.presentation) throw new Error("请先打开一个 .pptx 文件。");
    return exportEditedPptx(zip, this.presentation.slides, this.dirtyShapes, this.textEdits, this.styleEdits, this.structureEdits);
  }

  // ------------------------------------------------------------ slide list

  private editSet() {
    return {
      dirty: this.dirtyShapes,
      textEdits: this.textEdits,
      styleEdits: this.styleEdits,
      structure: this.structureEdits
    };
  }

  private allocateSlidePart(): string {
    const taken = [
      ...Object.keys(this.parser.getZip()?.files || {}),
      ...this.structureEdits.virtual.keys(),
      ...(this.presentation?.slides || []).map(slide => slide.part || "")
    ];
    let highest = 0;
    for (const path of taken) {
      const match = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
      if (match) highest = Math.max(highest, parseInt(match[1], 10));
    }
    return `ppt/slides/slide${highest + 1}.xml`;
  }

  private cloneElements(elements: SlideElement[], part?: string): SlideElement[] {
    return elements.map(element => {
      const copy: SlideElement = JSON.parse(JSON.stringify(element));
      if (part && copy.source) copy.source.part = part;
      return copy;
    });
  }

  /** Swap the slide list (and current page) and tell the UI; used by every slide command. */
  public async setSlides(slides: Slide[], index: number): Promise<void> {
    if (!this.presentation) return;
    this.presentation.slides = slides;
    this.currentSlideIndex = Math.max(0, Math.min(index, slides.length - 1));
    this.editor.clearSelection();
    await this.renderCurrentSlide();
    this.onSlidesChanged?.(this.presentation, this.currentSlideIndex);
  }

  private recordSlideChange(label: string, before: Slide[], beforeIndex: number, after: Slide[], afterIndex: number): void {
    const command: Command = {
      label,
      apply: () => { void this.setSlides([...after], afterIndex); },
      revert: () => { void this.setSlides([...before], beforeIndex); }
    };
    command.apply();
    this.editor.record(command);
  }

  private requirePackage(): void {
    if (!this.parser.getZip() || !this.presentation) throw new Error("请先打开一个 .pptx 文件。");
  }

  /** A blank page with the same layout as `index`, right after it. */
  public async addSlide(index = this.currentSlideIndex): Promise<void> {
    this.requirePackage();
    const zip = this.parser.getZip()!;
    const slides = this.presentation!.slides;
    const reference = slides[index];
    if (!reference?.part) return;

    const known = await readRelationships(zip, reference.part, this.structureEdits.relAdds.get(reference.part), this.structureEdits.virtual);
    const layout = [...known.values()].find(rel => rel.type.endsWith("/slideLayout"));
    if (!layout) throw new Error("找不到这一页使用的版式。");

    const part = this.allocateSlidePart();
    const xml = blankSlideXml();
    this.structureEdits.virtual.set(part, xml);
    this.structureEdits.virtual.set(
      `${part.slice(0, part.lastIndexOf("/"))}/_rels/${part.slice(part.lastIndexOf("/") + 1)}.rels`,
      relationshipsXml([{ ...layout, id: "rId1" }])
    );
    const slide: Slide = {
      id: `slide_${++this.slideSerial}`,
      // Decoration from the master and layout comes along; shapes do not.
      elements: this.cloneElements(reference.elements.filter(element => !element.source)),
      rawXml: xml,
      emuPerPx: reference.emuPerPx,
      part,
      xform: reference.xform
    };
    const after = [...slides];
    after.splice(index + 1, 0, slide);
    this.recordSlideChange("新建幻灯片", [...slides], this.currentSlideIndex, after, index + 1);
  }

  /** A copy of slide `index`, with the edits made so far, right after it. */
  public async duplicateSlide(index = this.currentSlideIndex): Promise<void> {
    this.requirePackage();
    const zip = this.parser.getZip()!;
    const slides = this.presentation!.slides;
    const source = slides[index];
    if (!source?.part) return;

    const snapshot = await materializeSlide(zip, source.part, slides, this.editSet());
    if (!snapshot) return;
    const copy = prepareSlideCopy(snapshot);
    const part = this.allocateSlidePart();
    this.structureEdits.virtual.set(part, copy.xml);
    this.structureEdits.virtual.set(
      `${part.slice(0, part.lastIndexOf("/"))}/_rels/${part.slice(part.lastIndexOf("/") + 1)}.rels`,
      relationshipsXml(copy.relationships)
    );
    const slide: Slide = {
      id: `slide_${++this.slideSerial}`,
      elements: this.cloneElements(source.elements, part),
      rawXml: copy.xml,
      emuPerPx: source.emuPerPx,
      part,
      xform: source.xform
    };
    const after = [...slides];
    after.splice(index + 1, 0, slide);
    this.recordSlideChange("复制幻灯片", [...slides], this.currentSlideIndex, after, index + 1);
  }

  public deleteSlide(index = this.currentSlideIndex): void {
    this.requirePackage();
    const slides = this.presentation!.slides;
    if (slides.length <= 1) throw new Error("至少需要保留一页幻灯片。");
    const after = slides.filter((_, i) => i !== index);
    this.recordSlideChange("删除幻灯片", [...slides], this.currentSlideIndex, after, Math.min(index, after.length - 1));
  }

  public moveSlide(from: number, to: number): void {
    this.requirePackage();
    const slides = this.presentation!.slides;
    if (from === to || from < 0 || from >= slides.length || to < 0 || to >= slides.length) return;
    const after = [...slides];
    const [moved] = after.splice(from, 1);
    after.splice(to, 0, moved);
    this.recordSlideChange("移动幻灯片", [...slides], this.currentSlideIndex, after, to);
  }

  private materializeSource(source: { part: string; shapeId: string }) {
    const zip = this.parser.getZip();
    if (!zip || !this.presentation) return Promise.resolve(null);
    return materializeShape(zip, source.part, source.shapeId, this.presentation.slides, {
      dirty: this.dirtyShapes,
      textEdits: this.textEdits,
      styleEdits: this.styleEdits,
      structure: this.structureEdits
    });
  }

  /** Next free p:cNvPr id of a slide, seeded from the highest id in its XML. */
  private allocateShapeId(part: string): string {
    let counter = this.shapeIdCounters.get(part);
    if (counter === undefined) {
      const xml = this.presentation?.slides.find(slide => slide.part === part)?.rawXml || "";
      counter = 1;
      for (const match of xml.matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)) {
        counter = Math.max(counter, parseInt(match[1], 10));
      }
    }
    counter += 1;
    this.shapeIdCounters.set(part, counter);
    return String(counter);
  }

  /** Body font of the deck, taken from the first text that has one. */
  private deckFont(): { fontFamily?: string; eastAsianFontFamily?: string } {
    for (const slide of this.presentation?.slides || []) {
      for (const element of slide.elements) {
        if (element.type !== "text" || !element.content.trim()) continue;
        const { fontFamily, eastAsianFontFamily } = element.style;
        return {
          ...(fontFamily && fontFamily !== "sans-serif" ? { fontFamily } : {}),
          ...(eastAsianFontFamily ? { eastAsianFontFamily } : {})
        };
      }
    }
    return {};
  }

  private async allocateRelationshipId(part: string): Promise<string> {
    const zip = this.parser.getZip();
    const existing = zip ? await readRelationships(zip, part, this.structureEdits.relAdds.get(part)) : new Map();
    let counter = 0;
    for (const id of existing.keys()) counter = Math.max(counter, parseInt(String(id).replace(/\D/g, ""), 10) || 0);
    return `rId${counter + 1}`;
  }

  private allocateMediaPath(extension: string): string {
    const taken = [
      ...Object.keys(this.parser.getZip()?.files || {}),
      ...this.structureEdits.media.map(item => item.path)
    ];
    let highest = 0;
    for (const path of taken) {
      const match = path.match(/^ppt\/media\/image(\d+)\./i);
      if (match) highest = Math.max(highest, parseInt(match[1], 10));
    }
    return `ppt/media/image${highest + 1}.${extension}`;
  }

  /** Relationship ids for pasted shapes: reuse what the target slide has, add what it lacks. */
  private async resolveRelationships(
    targetPart: string,
    sourcePart: string,
    relationships: RelAdd[]
  ): Promise<{ ids: Map<string, string>; adds: RelAdd[] }> {
    const ids = new Map<string, string>();
    const adds: RelAdd[] = [];
    const zip = this.parser.getZip();
    if (targetPart === sourcePart || !zip || relationships.length === 0) return { ids, adds };

    const existing = await readRelationships(zip, targetPart, this.structureEdits.relAdds.get(targetPart));
    let counter = 0;
    for (const id of existing.keys()) counter = Math.max(counter, parseInt(id.replace(/\D/g, ""), 10) || 0);
    for (const rel of relationships) {
      const same = [...existing.values()].find(candidate =>
        candidate.type === rel.type && candidate.target === rel.target && candidate.targetMode === rel.targetMode);
      if (same) {
        ids.set(rel.id, same.id);
        continue;
      }
      counter += 1;
      const added: RelAdd = { ...rel, id: `rId${counter}` };
      existing.set(added.id, added);
      adds.push(added);
      ids.set(rel.id, added.id);
    }
    return { ids, adds };
  }

  private resetEditing(): void {
    this.originalSlideParts = (this.presentation?.slides || []).map(slide => slide.part);
    this.slideSerial = this.presentation?.slides.length ?? 0;
    this.dirtyShapes = new Map();
    this.textEdits = new Map();
    this.styleEdits = new Map();
    this.structureEdits = emptyStructure();
    this.shapeIdCounters.clear();
    this.hiddenElements.clear();
    // Fonts differ per document, so the drag renderer is rebuilt on demand.
    this.dragRenderer = null;
    this.dragScratch = null;
    this.dragLayers = null;
    this.editor.reset();
  }

  /** Elements drawn by the DOM text editor instead of the canvas while editing. */
  private setHiddenElements(slide: Slide, elements: SlideElement[]): void {
    this.hiddenElements = new Set(elements);
    this.invalidateSlide(slide);
  }

  private invalidateSlide(slide: Slide): void {
    this.slideJsonCache.delete(slide.id);
    this.renderedSlideCache.delete(slide.id);
    if (this.dragLayers?.slide === slide) {
      this.drawDragFrame(this.dragLayers);
    } else if (this.getCurrentSlideAST() === slide) {
      this.resizeAndRedraw(slide);
    }
  }

  /**
   * While an element is dragged, everything underneath and above it is rendered
   * once into bitmaps; each frame then re-renders only the moving elements and
   * composites the three layers, instead of re-laying-out the whole slide.
   */
  private beginDragLayers(slide: Slide, moving: SlideElement[]): void {
    if (!this.presentation || !this.renderer || this.isRenderingThumbnails) return;
    const width = this.canvas.width;
    const height = this.canvas.height;
    if (width < 1 || height < 1) return;

    const movingSet = new Set(moving);
    const firstMoving = slide.elements.findIndex(element => movingSet.has(element));
    if (firstMoving < 0) return;

    this.ensureDragRenderer(width, height);
    const scratch = this.dragScratch!;
    const scratchContext = this.dragContext!;
    const logical = this.presentation.size;
    const scale = Math.min(width / logical.width, height / logical.height);
    const offsetX = (width - logical.width * scale) / 2;
    const offsetY = (height - logical.height * scale) / 2;

    const makeLayer = (elements: SlideElement[], withBackground: boolean): HTMLCanvasElement => {
      const layer = document.createElement("canvas");
      layer.width = width;
      layer.height = height;
      scratchContext.setTransform(1, 0, 0, 1, 0, 0);
      scratchContext.clearRect(0, 0, width, height);
      scratchContext.setTransform(scale, 0, 0, scale, offsetX, offsetY);
      const json = JSON.stringify({ id: slide.id, elements });
      if (withBackground) this.dragRenderer!.render_slide(json, this.imageCache);
      else if (elements.length > 0) this.dragRenderer!.render_slide_overlay(json, this.imageCache);
      layer.getContext("2d")!.drawImage(scratch, 0, 0);
      return layer;
    };

    const rest = slide.elements.filter(element => !movingSet.has(element));
    const belowCount = slide.elements.slice(0, firstMoving).length;
    this.dragLayers = {
      slide,
      moving: slide.elements.filter(element => movingSet.has(element)),
      below: makeLayer(rest.slice(0, belowCount), true),
      above: makeLayer(rest.slice(belowCount), false),
      scale,
      offsetX,
      offsetY
    };
  }

  private ensureDragRenderer(width: number, height: number): void {
    if (this.dragRenderer && this.dragScratch?.width === width && this.dragScratch.height === height) return;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Failed to create the drag layer context");
    const renderer = new wasm.RustPptRenderer(context);
    for (const fontBytes of this.registeredFontBytes) renderer.register_font(fontBytes);
    this.dragScratch = canvas;
    this.dragContext = context;
    this.dragRenderer = renderer;
  }

  private drawDragFrame(layers: NonNullable<PptViewer["dragLayers"]>): void {
    const scratch = this.dragScratch!;
    const scratchContext = this.dragContext!;
    scratchContext.setTransform(1, 0, 0, 1, 0, 0);
    scratchContext.clearRect(0, 0, scratch.width, scratch.height);
    scratchContext.setTransform(layers.scale, 0, 0, layers.scale, layers.offsetX, layers.offsetY);
    this.dragRenderer!.render_slide_overlay(
      JSON.stringify({ id: layers.slide.id, elements: layers.moving }),
      this.imageCache
    );

    this.ctx.save();
    this.ctx.setTransform(1, 0, 0, 1, 0, 0);
    this.ctx.drawImage(layers.below, 0, 0);
    this.ctx.drawImage(scratch, 0, 0);
    this.ctx.drawImage(layers.above, 0, 0);
    this.ctx.restore();
  }

  private endDragLayers(): void {
    this.dragLayers = null;
  }

  private notifySlideEdited(slide: Slide): void {
    window.clearTimeout(this.editNotifyTimer);
    this.editNotifyTimer = window.setTimeout(() => {
      const index = this.presentation?.slides.indexOf(slide) ?? -1;
      if (index >= 0) this.onSlideEdited?.(index, slide);
    }, 350);
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
    if (this.hiddenElements.size > 0) {
      return JSON.stringify({
        id: slide.id,
        elements: slide.elements.filter(element => !this.hiddenElements.has(element))
      });
    }
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
    // Pre-rendering neighbours at a large zoom would only burn memory.
    if (this.canvas.width > 2600) return;

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
    const fitWidth = Math.min(availableWidth, availableHeight * slideRatio);
    // 100% is the slide at its real size: 96 css px per inch, 9525 EMU per px.
    const realWidth = Math.max(1, logicalSize.width * (slide.emuPerPx || 6350) / 9525);
    const cssWidth = this.zoomMode === "fit" ? fitWidth : realWidth * this.zoomMode / 100;
    const cssHeight = cssWidth / slideRatio;
    const percent = cssWidth / realWidth * 100;
    if (Math.abs(percent - this.zoomPercent) > 0.01) {
      this.zoomPercent = percent;
      this.onZoomChange?.(percent, this.zoomMode === "fit");
    }

    this.container.style.width = `${cssWidth}px`;
    this.container.style.height = `${cssHeight}px`;
    this.container.style.aspectRatio = `${logicalSize.width} / ${logicalSize.height}`;

    // Device Pixel Ratio scaling for Retina screens (sharp rendering)
    // Past ~4k wide a bitmap costs more than it shows; the browser scales it up.
    const dpr = Math.min(window.devicePixelRatio || 1, 4096 / cssWidth);
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
    this.editor.layout();

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
