import { PptxParser, PresentationAST, Slide, PresentationSize } from "./pptx-parser";
import * as wasm from "../rust-engine/pkg/ppt_engine";

export interface ViewerOptions {
  container: HTMLElement;
  width?: number; // CSS width
  height?: number; // CSS height
  onSlideChange?: (slideIndex: number, ast: Slide) => void;
  onLoadComplete?: (ast: PresentationAST) => void;
}

export default class PptViewer {
  private container: HTMLElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private renderer: wasm.RustPptRenderer | null = null;
  private parser: PptxParser;
  private presentation: PresentationAST | null = null;
  private currentSlideIndex: number = 0;
  private imageCache: Record<string, HTMLImageElement> = {};
  
  // Callbacks
  private onSlideChange?: (slideIndex: number, ast: Slide) => void;
  private onLoadComplete?: (ast: PresentationAST) => void;

  constructor(options: ViewerOptions) {
    this.container = options.container;
    this.onSlideChange = options.onSlideChange;
    this.onLoadComplete = options.onLoadComplete;

    // Create Canvas element
    this.canvas = document.createElement("canvas");
    this.canvas.style.width = options.width ? `${options.width}px` : "100%";
    this.canvas.style.height = options.height ? `${options.height}px` : "100%";
    this.canvas.style.display = "block";
    this.canvas.style.boxShadow = "0 8px 24px rgba(0,0,0,0.15)";
    this.canvas.style.borderRadius = "8px";
    this.canvas.style.backgroundColor = "#ffffff";
    this.container.appendChild(this.canvas);

    const context = this.canvas.getContext("2d");
    if (!context) throw new Error("Failed to get 2D canvas context");
    this.ctx = context;

    this.parser = new PptxParser();
    this.initEngine();

    // Listen to resize to keep canvas crisp
    window.addEventListener("resize", () => this.resizeAndRedraw());
  }

  private async initEngine() {
    try {
      const wasmModule = await import("../rust-engine/pkg/ppt_engine");
      this.renderer = new wasmModule.RustPptRenderer(this.ctx);
      console.log("PPT WASM Engine initialized successfully.");

      // Load and register layout fonts in WASM
      await this.loadAndRegisterFonts();
      
      // If a presentation was loaded before engine was ready, render it now
      if (this.presentation) {
        this.renderCurrentSlide();
      }
    } catch (err) {
      console.error("Failed to load Rust WASM Engine:", err);
    }
  }

  private async loadAndRegisterFonts() {
    if (!this.renderer) return;
    console.log("Downloading standard layout fonts...");
    const fontUrls = {
      regular: "https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.3.3/fonts/Roboto/Roboto-Regular.ttf",
      bold: "https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.3.3/fonts/Roboto/Roboto-Medium.ttf"
    };

    try {
      const loadFont = async (url: string) => {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`HTTP ${response.status} fetching font`);
        const buffer = await response.arrayBuffer();
        const bytes = new Uint8Array(buffer);
        this.renderer?.register_font(bytes);
      };

      // Download both fonts in parallel
      await Promise.all([
        loadFont(fontUrls.regular),
        loadFont(fontUrls.bold)
      ]);
      console.log("Standard fonts registered in WASM successfully.");
    } catch (err) {
      console.error("Failed to fetch and register layout fonts in WASM:", err);
    }
  }

  // Load PPTX file from ArrayBuffer
  public async loadPptx(buffer: ArrayBuffer): Promise<PresentationAST> {
    this.presentation = await this.parser.parse(buffer);
    this.currentSlideIndex = 0;
    this.imageCache = {}; // Clear old image cache
    
    if (this.onLoadComplete) {
      this.onLoadComplete(this.presentation);
    }

    await this.renderCurrentSlide();
    return this.presentation;
  }

  // Load manual AST structure (for demos / testing)
  public async loadAST(ast: PresentationAST): Promise<void> {
    this.presentation = ast;
    this.currentSlideIndex = 0;
    this.imageCache = {};
    
    if (this.onLoadComplete) {
      this.onLoadComplete(this.presentation);
    }
    
    await this.renderCurrentSlide();
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

  private async renderCurrentSlide() {
    if (!this.presentation || !this.renderer) return;

    const slide = this.getCurrentSlideAST();
    if (!slide) return;

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

    // 2. Setup canvas dimensions and scale
    this.resizeAndRedraw();
  }

  private resizeAndRedraw() {
    if (!this.presentation || !this.renderer) return;

    const slide = this.getCurrentSlideAST();
    if (!slide) return;

    const logicalSize = this.presentation.size;

    // Get display sizes in CSS pixels
    const rect = this.canvas.getBoundingClientRect();
    const cssWidth = rect.width || this.canvas.clientWidth || 960;
    const cssHeight = rect.height || this.canvas.clientHeight || 540;

    // Device Pixel Ratio scaling for Retina screens (sharp rendering)
    const dpr = window.devicePixelRatio || 1;
    this.canvas.width = cssWidth * dpr;
    this.canvas.height = cssHeight * dpr;

    this.ctx.save();
    
    // Scale coordinate system from logical (e.g. 1920x1080) to physical pixel viewport
    const scaleX = this.canvas.width / logicalSize.width;
    const scaleY = this.canvas.height / logicalSize.height;
    this.ctx.setTransform(scaleX, 0, 0, scaleY, 0, 0);

    // Call Rust WASM renderer
    try {
      this.renderer.render_slide(JSON.stringify(slide), this.imageCache);
    } catch (err) {
      console.error("Error during Rust rendering:", err);
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
