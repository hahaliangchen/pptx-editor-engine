import JSZip from "jszip";

// AST Interfaces
export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface TextStyle {
  fontSize: number;
  color: string;
  bold: boolean;
  align: "left" | "center" | "right";
  fontFamily?: string;
  eastAsianFontFamily?: string;
  italic?: boolean;
}

export interface TextRun {
  content: string;
  style: TextStyle;
}

export interface ParagraphStyle {
  align: "left" | "center" | "right";
  level: number;
  marginLeft: number;
  indent: number;
  lineSpacing?: {
    unit: "percent" | "points";
    value: number;
  };
  spaceBefore?: number;
  spaceAfter?: number;
}

export interface TextParagraph {
  runs: TextRun[];
  style: ParagraphStyle;
  bullet?: {
    char: string;
    color: string;
  };
}

export interface TextBodyProperties {
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  verticalAnchor: "top" | "middle" | "bottom";
  autoFit: "none" | "shape" | "shrink";
  fontScale?: number;
}

export interface TextElement {
  type: "text";
  id: string;
  rect: Rect;
  content: string;
  style: TextStyle;
  paragraphs?: TextParagraph[];
  body?: TextBodyProperties;
}

export interface Border {
  color: string;
  width: number;
}

export interface ShapeElement {
  type: "shape";
  id: string;
  rect: Rect;
  shapeType: "rect" | "roundRect" | "ellipse" | "triangle";
  fill: string;
  border?: Border;
}

export interface ImageElement {
  type: "image";
  id: string;
  rect: Rect;
  url: string;
  crop?: ImageCrop;
}

export interface ImageCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export type SlideElement = TextElement | ShapeElement | ImageElement;

export interface Slide {
  id: string;
  elements: SlideElement[];
  rawXml: string; // Raw slide XML text
}

export interface PresentationSize {
  width: number;
  height: number;
}

export interface PresentationAST {
  size: PresentationSize;
  slides: Slide[];
}

interface PlaceholderDescriptor {
  idx: string | null;
  type: string;
}

interface PlaceholderContext {
  layoutShapes: globalThis.Element[];
  masterShapes: globalThis.Element[];
  masterTextStyles: globalThis.Element | null;
}

interface TextInheritanceContext {
  textBodies: globalThis.Element[];
  masterTextStyle: globalThis.Element | null;
}

interface ComputedRunStyle {
  fontSize: number;
  color: string;
  bold: boolean;
  fontFamily: string;
  eastAsianFontFamily: string;
  italic: boolean;
}

// XML Namespaces query helper
function querySelector(parent: Element | Document, selectors: string): Element | null {
  const parts = selectors.split(",").map(s => s.trim());
  for (const part of parts) {
    try {
      const el = parent.querySelector(part);
      if (el) return el as Element;
    } catch (e) {}
  }
  // LocalName fallback
  for (const part of parts) {
    const cleanTag = part.replace(/^.*\\:/, "").replace(/^.*:/, "");
    const els = parent.getElementsByTagName(cleanTag);
    if (els.length > 0) return els[0] as Element;
  }
  return null;
}

function querySelectorAll(parent: Element | Document, selectors: string): Element[] {
  const parts = selectors.split(",").map(s => s.trim());
  for (const part of parts) {
    try {
      const list = parent.querySelectorAll(part);
      if (list.length > 0) return Array.from(list) as Element[];
    } catch (e) {}
  }
  // LocalName fallback
  for (const part of parts) {
    const cleanTag = part.replace(/^.*\\:/, "").replace(/^.*:/, "");
    const els = parent.getElementsByTagName(cleanTag);
    if (els.length > 0) return Array.from(els) as Element[];
  }
  return [];
}

export class PptxParser {
  private zip: JSZip | null = null;
  private themeColors: Record<string, string> = {
    "dk1": "#000000",
    "lt1": "#ffffff",
    "dk2": "#1f497d",
    "lt2": "#eef1f5",
    "accent1": "#4f81bd",
    "accent2": "#c0504d",
    "accent3": "#9bbb59",
    "accent4": "#8064a2",
    "accent5": "#4bacc6",
    "accent6": "#f79646",
    "hlink": "#0000ff",
    "folHlink": "#800080",
    "window": "#ffffff",
    "bg1": "#ffffff",
    "tx1": "#000000",
    "bg2": "#eef1f5",
    "tx2": "#1f497d"
  };
  private themeFonts: Record<string, string> = {
    "+mj-lt": "",
    "+mj-ea": "",
    "+mj-cs": "",
    "+mn-lt": "",
    "+mn-ea": "",
    "+mn-cs": "",
    "+mj-Hans": "",
    "+mj-Hant": "",
    "+mj-Jpan": "",
    "+mj-Hang": "",
    "+mn-Hans": "",
    "+mn-Hant": "",
    "+mn-Jpan": "",
    "+mn-Hang": ""
  };

  public async parse(buffer: ArrayBuffer): Promise<PresentationAST> {
    this.zip = await JSZip.loadAsync(buffer);

    // 1. Parse Theme Colors
    await this.parseThemeColors();
    await this.parseThemeFonts();

    // 2. Parse Presentation Properties (slide size)
    const size = await this.parsePresentationSize();

    // 3. Find and sort slide files
    const slidePaths = await this.getOrderedSlidePaths();

    // 4. Parse individual slides
    const slides: Slide[] = [];
    for (let i = 0; i < slidePaths.length; i++) {
      const path = slidePaths[i];
      try {
        const slide = await this.parseSlide(path, `slide_${i + 1}`, size);
        slides.push(slide);
      } catch (err) {
        console.error(`Error parsing slide ${path}:`, err);
      }
    }

    return { size, slides };
  }

  private async parseThemeColors() {
    const themeFile = this.zip?.file("ppt/theme/theme1.xml");
    if (!themeFile) return;

    const xmlText = await themeFile.async("text");
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");

    // Standard schemes
    const schemeNode = querySelector(doc, "a\\:clrScheme, clrScheme");
    if (!schemeNode) return;

    const colorKeys = [
      "dk1", "lt1", "dk2", "lt2",
      "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
      "hlink", "folHlink"
    ];

    for (const key of colorKeys) {
      // Find node under schemeNode with localName = key or prefix:key
      const colorNode = querySelector(schemeNode, `a\\:${key}, ${key}`);
      if (colorNode) {
        const hex = this.extractHexColor(colorNode);
        if (hex) {
          this.themeColors[key] = hex;
          // Set text and bg aliases
          if (key === "dk1") this.themeColors["tx1"] = hex;
          if (key === "lt1") this.themeColors["bg1"] = hex;
          if (key === "dk2") this.themeColors["tx2"] = hex;
          if (key === "lt2") this.themeColors["bg2"] = hex;
        }
      }
    }
  }

  private async parseThemeFonts() {
    const themeFile = this.zip?.file("ppt/theme/theme1.xml");
    if (!themeFile) return;

    const xmlText = await themeFile.async("text");
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");
    const major = querySelector(doc, "a\\:majorFont, majorFont");
    const minor = querySelector(doc, "a\\:minorFont, minorFont");

    const readTypeface = (root: Element | null, selector: string) =>
      root ? querySelector(root, selector)?.getAttribute("typeface") || "" : "";

    this.themeFonts["+mj-lt"] = readTypeface(major, "a\\:latin, latin");
    this.themeFonts["+mj-ea"] = readTypeface(major, "a\\:ea, ea");
    this.themeFonts["+mj-cs"] = readTypeface(major, "a\\:cs, cs");
    this.themeFonts["+mn-lt"] = readTypeface(minor, "a\\:latin, latin");
    this.themeFonts["+mn-ea"] = readTypeface(minor, "a\\:ea, ea");
    this.themeFonts["+mn-cs"] = readTypeface(minor, "a\\:cs, cs");

    for (const [prefix, root] of [["+mj", major], ["+mn", minor]] as const) {
      if (!root) continue;
      for (const font of querySelectorAll(root, "a\\:font, font")) {
        const script = font.getAttribute("script");
        const typeface = font.getAttribute("typeface");
        if (script && typeface && ["Hans", "Hant", "Jpan", "Hang"].includes(script)) {
          this.themeFonts[`${prefix}-${script}`] = typeface;
        }
      }
    }
  }

  private resolveThemeTypeface(typeface: string): string {
    if (!typeface.startsWith("+")) return typeface;
    const resolved = this.themeFonts[typeface];
    if (resolved) return resolved;
    if (typeface.endsWith("-ea") || typeface.endsWith("-cs")) {
      return this.themeFonts[`${typeface.substring(0, 3)}-lt`] || "sans-serif";
    }
    return "sans-serif";
  }

  private extractHexColor(node: Element): string | null {
    let hex: string | null = null;
    let colorNode: Element | null = null;

    // 1. Direct srgbClr
    const srgbNode = querySelector(node, "a\\:srgbClr, srgbClr");
    if (srgbNode) {
      const val = srgbNode.getAttribute("val");
      if (val) {
        hex = `#${val}`;
        colorNode = srgbNode;
      }
    }

    // 2. sysClr
    if (!hex) {
      const sysNode = querySelector(node, "a\\:sysClr, sysClr");
      if (sysNode) {
        const val = sysNode.getAttribute("lastClr");
        if (val) {
          hex = `#${val}`;
          colorNode = sysNode;
        }
      }
    }

    // 3. schemeClr (nested reference)
    if (!hex) {
      const schemeNode = querySelector(node, "a\\:schemeClr, schemeClr");
      if (schemeNode) {
        const val = schemeNode.getAttribute("val");
        if (val && this.themeColors[val]) {
          hex = this.themeColors[val];
          colorNode = schemeNode;
        }
      }
    }

    if (!hex || !colorNode) return hex;

    // Apply color transformations (tint, shade, lumMod, lumOff) under the colorNode
    try {
      let r = parseInt(hex.substring(1, 3), 16);
      let g = parseInt(hex.substring(3, 5), 16);
      let b = parseInt(hex.substring(5, 7), 16);

      const children = Array.from(colorNode.childNodes) as Element[];
      for (const child of children) {
        if (child.nodeType !== 1) continue;
        const tagName = child.localName || child.nodeName;
        const valAttr = child.getAttribute("val");
        if (!valAttr) continue;
        const val = parseInt(valAttr, 10) / 100000; // OOXML values are in 1000ths of a percent

        if (tagName.endsWith("tint")) {
          // Tint: blend with white. E.g. val=0.8 means 80% color, 20% white
          r = Math.round(r * val + 255 * (1 - val));
          g = Math.round(g * val + 255 * (1 - val));
          b = Math.round(b * val + 255 * (1 - val));
        } else if (tagName.endsWith("shade")) {
          // Shade: blend with black. E.g. val=0.8 means 80% color, 20% black
          r = Math.round(r * val);
          g = Math.round(g * val);
          b = Math.round(b * val);
        } else if (tagName.endsWith("lumMod")) {
          r = Math.round(r * val);
          g = Math.round(g * val);
          b = Math.round(b * val);
        } else if (tagName.endsWith("lumOff")) {
          const offset = Math.round(255 * val);
          r = Math.min(255, r + offset);
          g = Math.min(255, g + offset);
          b = Math.min(255, b + offset);
        }
      }

      // Clamp values to 0-255
      r = Math.max(0, Math.min(255, r));
      g = Math.max(0, Math.min(255, g));
      b = Math.max(0, Math.min(255, b));

      return "#" + r.toString(16).padStart(2, '0') + g.toString(16).padStart(2, '0') + b.toString(16).padStart(2, '0');
    } catch (e) {
      return hex;
    }
  }

  private approximateGradientColor(gradient: Element): string | null {
    const stops = querySelectorAll(gradient, "a\\:gs, gs")
      .map(stop => this.extractHexColor(stop))
      .filter((color): color is string => color !== null);
    if (stops.length === 0) return null;
    if (stops.length === 1) return stops[0];

    const rgb = stops.map(color => ({
      r: parseInt(color.slice(1, 3), 16),
      g: parseInt(color.slice(3, 5), 16),
      b: parseInt(color.slice(5, 7), 16)
    }));
    const average = (channel: "r" | "g" | "b") =>
      Math.round(rgb.reduce((sum, color) => sum + color[channel], 0) / rgb.length)
        .toString(16)
        .padStart(2, "0");
    return `#${average("r")}${average("g")}${average("b")}`;
  }

  private async parsePresentationSize(): Promise<PresentationSize> {
    const defaultSize: PresentationSize = { width: 1920, height: 1080 };
    const presFile = this.zip?.file("ppt/presentation.xml");
    if (!presFile) return defaultSize;

    const xmlText = await presFile.async("text");
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");

    const sldSzNode = querySelector(doc, "p\\:sldSz, sldSz");
    if (!sldSzNode) return defaultSize;

    const cxStr = sldSzNode.getAttribute("cx");
    const cyStr = sldSzNode.getAttribute("cy");
    if (!cxStr || !cyStr) return defaultSize;

    const cx = parseInt(cxStr, 10);
    const cy = parseInt(cyStr, 10);

    // Compute standard aspect ratio scale (base width is 1920)
    const logicalWidth = 1920;
    const logicalHeight = Math.round((cy / cx) * 1920);

    return { width: logicalWidth, height: logicalHeight };
  }

  private async getOrderedSlidePaths(): Promise<string[]> {
    const presFile = this.zip?.file("ppt/presentation.xml");
    const relsFile = this.zip?.file("ppt/_rels/presentation.xml.rels");
    if (!presFile || !relsFile) return [];

    // Parse relations
    const relsText = await relsFile.async("text");
    const relsDoc = new DOMParser().parseFromString(relsText, "application/xml");
    const relationships = querySelectorAll(relsDoc, "Relationship");
    const relMap: Record<string, string> = {};
    for (const rel of relationships) {
      const id = rel.getAttribute("Id");
      const target = rel.getAttribute("Target");
      if (id && target) {
        relMap[id] = target;
      }
    }

    // Parse presentation to get slide ordering
    const presText = await presFile.async("text");
    const presDoc = new DOMParser().parseFromString(presText, "application/xml");
    const slideIds = querySelectorAll(presDoc, "p\\:sldId, sldId");

    const paths: string[] = [];
    for (const slideIdNode of slideIds) {
      const rId = slideIdNode.getAttribute("r:id") || slideIdNode.getAttribute("id");
      if (rId && relMap[rId]) {
        // Resolve Target relative to ppt/
        let targetPath = relMap[rId];
        if (!targetPath.startsWith("ppt/")) {
          targetPath = "ppt/" + targetPath;
        }
        paths.push(targetPath);
      }
    }

    // Fallback if presentation.xml parsing fails or lists no slides
    if (paths.length === 0) {
      const files = Object.keys(this.zip?.files || {});
      const slideFiles = files
        .filter(f => f.startsWith("ppt/slides/slide") && f.endsWith(".xml"))
        .sort((a, b) => {
          const numA = parseInt(a.replace(/[^\d]/g, ""), 10) || 0;
          const numB = parseInt(b.replace(/[^\d]/g, ""), 10) || 0;
          return numA - numB;
        });
      return slideFiles;
    }

    return paths;
  }

  private async loadRelationships(relsPath: string, baseDir: string): Promise<{
    imgRelMap: Record<string, string>;
    layoutPath?: string;
    masterPath?: string;
  }> {
    const relMap: Record<string, string> = {};
    let layoutPath: string | undefined = undefined;
    let masterPath: string | undefined = undefined;

    const relsFile = this.zip?.file(relsPath);
    if (relsFile) {
      const relsText = await relsFile.async("text");
      const relsDoc = new DOMParser().parseFromString(relsText, "application/xml");
      const relationships = querySelectorAll(relsDoc, "Relationship");
      for (const rel of relationships) {
        const id = rel.getAttribute("Id");
        const target = rel.getAttribute("Target");
        const type = rel.getAttribute("Type");
        if (id && target && type) {
          if (type.includes("image")) {
            relMap[id] = this.resolveRelativePath(baseDir, target);
          } else if (type.includes("slideLayout")) {
            layoutPath = this.resolveRelativePath(baseDir, target);
          } else if (type.includes("slideMaster")) {
            masterPath = this.resolveRelativePath(baseDir, target);
          }
        }
      }
    }
    return { imgRelMap: relMap, layoutPath, masterPath };
  }

  private async parseSlide(slidePath: string, slideId: string, viewSize: PresentationSize): Promise<Slide> {
    const slideFile = this.zip?.file(slidePath);
    if (!slideFile) throw new Error(`Slide file not found: ${slidePath}`);

    const xmlText = await slideFile.async("text");
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");

    // Get EMU to Logical Pixel Scale
    // We get the original slide dimensions by searching ppt/presentation.xml again, or we can assume standard widescreen.
    // Let's retrieve sldSz from presentation.xml or default it. We'll read the scale factors.
    // In parsePresentationSize we mapped cx/cy to width/height.
    // Standard PPTX widescreen uses cx=9144000, cy=5142864.
    // Let's look up slide dimensions from presentation.xml to calculate correct EMU scaling.
    const presFile = this.zip?.file("ppt/presentation.xml");
    let emuWidth = 9144000;
    let emuHeight = 5142864;
    if (presFile) {
      const presText = await presFile.async("text");
      const presDoc = new DOMParser().parseFromString(presText, "application/xml");
      const sldSzNode = querySelector(presDoc, "p\\:sldSz, sldSz");
      if (sldSzNode) {
        emuWidth = parseInt(sldSzNode.getAttribute("cx") || "9144000", 10);
        emuHeight = parseInt(sldSzNode.getAttribute("cy") || "5142864", 10);
      }
    }

    const scaleX = viewSize.width / emuWidth;
    const scaleY = viewSize.height / emuHeight;

    // Load relationships for images
    const slideDir = slidePath.substring(0, slidePath.lastIndexOf("/"));
    const slideFileName = slidePath.substring(slidePath.lastIndexOf("/") + 1);
    const slideRelsPath = `${slideDir}/_rels/${slideFileName}.rels`;
    const { imgRelMap, layoutPath } = await this.loadRelationships(slideRelsPath, slideDir);

    let layoutImgRelMap: Record<string, string> = {};
    let masterImgRelMap: Record<string, string> = {};
    let layoutDoc: Document | null = null;
    let masterDoc: Document | null = null;

    // Load layout if present
    if (layoutPath) {
      const layoutDir = layoutPath.substring(0, layoutPath.lastIndexOf("/"));
      const layoutFileName = layoutPath.substring(layoutPath.lastIndexOf("/") + 1);
      const layoutRelsPath = `${layoutDir}/_rels/${layoutFileName}.rels`;
      const layoutRels = await this.loadRelationships(layoutRelsPath, layoutDir);
      layoutImgRelMap = layoutRels.imgRelMap;

      const layoutFile = this.zip?.file(layoutPath);
      if (layoutFile) {
        const text = await layoutFile.async("text");
        layoutDoc = new DOMParser().parseFromString(text, "application/xml");
      }

      // Load master if present in layout relations
      if (layoutRels.masterPath) {
        const masterPath = layoutRels.masterPath;
        const masterDir = masterPath.substring(0, masterPath.lastIndexOf("/"));
        const masterRelsPath = `${masterDir}/_rels/${masterPath.substring(masterPath.lastIndexOf("/") + 1)}.rels`;
        const masterRels = await this.loadRelationships(masterRelsPath, masterDir);
        masterImgRelMap = masterRels.imgRelMap;

        const masterFile = this.zip?.file(masterPath);
        if (masterFile) {
          const text = await masterFile.async("text");
          masterDoc = new DOMParser().parseFromString(text, "application/xml");
        }
      }
    }

    const elements: SlideElement[] = [];
    const placeholderContext: PlaceholderContext = {
      layoutShapes: layoutDoc ? this.collectPlaceholderShapes(layoutDoc) : [],
      masterShapes: masterDoc ? this.collectPlaceholderShapes(masterDoc) : [],
      masterTextStyles: masterDoc
        ? querySelector(masterDoc, "p\\:txStyles, txStyles") as unknown as globalThis.Element | null
        : null
    };

    // Initialize coordinate transform (EMU to logical pixels)
    const initialTransform: TransformState = {
      scaleX: scaleX,
      scaleY: scaleY,
      absoluteUnitScale: Math.min(scaleX, scaleY),
      offsetX: 0,
      offsetY: 0
    };

    // First: Parse Slide Master background elements
    if (masterDoc) {
      const sldMaster = querySelector(masterDoc, "p\\:sldMaster, sldMaster");
      const cSld = sldMaster ? querySelector(sldMaster, "p\\:cSld, cSld") : null;
      const spTree = cSld ? querySelector(cSld, "p\\:spTree, spTree") : null;
      if (spTree) {
        await this.parseElementsRecursive(spTree, initialTransform, masterImgRelMap, elements, true);
      }
    }

    // Second: Parse Slide Layout midground elements
    if (layoutDoc) {
      const sldLayout = querySelector(layoutDoc, "p\\:sldLayout, sldLayout");
      const cSld = sldLayout ? querySelector(sldLayout, "p\\:cSld, cSld") : null;
      const spTree = cSld ? querySelector(cSld, "p\\:spTree, spTree") : null;
      if (spTree) {
        await this.parseElementsRecursive(spTree, initialTransform, layoutImgRelMap, elements, true);
      }
    }

    // Third: Parse Slide shapes (foreground layer)
    const sld = querySelector(doc, "p\\:sld, sld");
    const cSld = sld ? querySelector(sld, "p\\:cSld, cSld") : null;
    const spTree = cSld ? querySelector(cSld, "p\\:spTree, spTree") : null;
    if (spTree) {
      await this.parseElementsRecursive(spTree, initialTransform, imgRelMap, elements, false, placeholderContext);
    }

    return { id: slideId, elements, rawXml: xmlText };
  }

  private resolveRelativePath(baseDir: string, relativePath: string): string {
    const baseParts = baseDir.split("/");
    const relParts = relativePath.split("/");
    for (const part of relParts) {
      if (part === "..") {
        baseParts.pop();
      } else if (part !== ".") {
        baseParts.push(part);
      }
    }
    return baseParts.join("/");
  }

  private collectPlaceholderShapes(doc: Document): globalThis.Element[] {
    return querySelectorAll(doc, "p\\:sp, sp")
      .filter(shape => this.getPlaceholderDescriptor(shape) !== null);
  }

  private getPlaceholderDescriptor(shape: globalThis.Element): PlaceholderDescriptor | null {
    const nvSpPr = querySelector(shape, "p\\:nvSpPr, nvSpPr");
    const nvPr = nvSpPr ? querySelector(nvSpPr, "p\\:nvPr, nvPr") : null;
    const ph = nvPr ? querySelector(nvPr, "p\\:ph, ph") : null;
    if (!ph) return null;

    return {
      idx: ph.getAttribute("idx"),
      type: ph.getAttribute("type") || "obj"
    };
  }

  private findMatchingPlaceholder(
    descriptor: PlaceholderDescriptor,
    candidates: globalThis.Element[]
  ): globalThis.Element | null {
    if (descriptor.idx !== null) {
      const byIndex = candidates.find(candidate =>
        this.getPlaceholderDescriptor(candidate)?.idx === descriptor.idx
      );
      if (byIndex) return byIndex;
    }

    return candidates.find(candidate =>
      this.getPlaceholderDescriptor(candidate)?.type === descriptor.type
    ) || null;
  }

  private resolvePlaceholderChain(
    shape: globalThis.Element,
    context: PlaceholderContext | null
  ): globalThis.Element[] {
    const chain = [shape];
    if (!context) return chain;

    const slideDescriptor = this.getPlaceholderDescriptor(shape);
    if (!slideDescriptor) return chain;

    const layoutShape = this.findMatchingPlaceholder(slideDescriptor, context.layoutShapes);
    if (layoutShape) chain.push(layoutShape);

    const masterDescriptor = layoutShape
      ? this.getPlaceholderDescriptor(layoutShape) || slideDescriptor
      : slideDescriptor;
    const masterShape = this.findMatchingPlaceholder(masterDescriptor, context.masterShapes);
    if (masterShape) chain.push(masterShape);

    return chain;
  }

  private selectMasterTextStyle(
    txStyles: globalThis.Element | null,
    placeholder: PlaceholderDescriptor | null
  ): globalThis.Element | null {
    if (!txStyles) return null;
    const type = placeholder?.type || "obj";
    if (type === "title" || type === "ctrTitle") {
      return querySelector(txStyles, "p\\:titleStyle, titleStyle");
    }
    if (type === "body" || type === "obj" || type === "subTitle") {
      return querySelector(txStyles, "p\\:bodyStyle, bodyStyle");
    }
    return querySelector(txStyles, "p\\:otherStyle, otherStyle");
  }

  private getLevelParagraphProperties(
    container: globalThis.Element,
    level: number
  ): globalThis.Element | null {
    const styleRoot = (container.localName || container.nodeName).endsWith("txBody")
      ? querySelector(container, "a\\:lstStyle, lstStyle")
      : container;
    if (!styleRoot) return null;
    const levelName = `lvl${Math.max(0, Math.min(8, level)) + 1}pPr`;
    return querySelector(styleRoot, `a\\:${levelName}, ${levelName}`);
  }

  private resolveTextBodyProperties(
    textBodies: globalThis.Element[],
    scaleX: number,
    scaleY: number
  ): TextBodyProperties {
    const emuToX = (value: string | null, fallback: number) =>
      (value === null ? fallback : parseInt(value, 10)) * scaleX;
    const emuToY = (value: string | null, fallback: number) =>
      (value === null ? fallback : parseInt(value, 10)) * scaleY;

    const result: TextBodyProperties = {
      marginLeft: 91440 * scaleX,
      marginRight: 91440 * scaleX,
      marginTop: 45720 * scaleY,
      marginBottom: 45720 * scaleY,
      verticalAnchor: "top",
      autoFit: "none",
      fontScale: 1
    };

    for (const txBody of [...textBodies].reverse()) {
      const bodyPr = querySelector(txBody, "a\\:bodyPr, bodyPr");
      if (!bodyPr) continue;
      if (bodyPr.hasAttribute("lIns")) result.marginLeft = emuToX(bodyPr.getAttribute("lIns"), 91440);
      if (bodyPr.hasAttribute("rIns")) result.marginRight = emuToX(bodyPr.getAttribute("rIns"), 91440);
      if (bodyPr.hasAttribute("tIns")) result.marginTop = emuToY(bodyPr.getAttribute("tIns"), 45720);
      if (bodyPr.hasAttribute("bIns")) result.marginBottom = emuToY(bodyPr.getAttribute("bIns"), 45720);

      const anchor = bodyPr.getAttribute("anchor");
      if (anchor === "ctr") result.verticalAnchor = "middle";
      else if (anchor === "b") result.verticalAnchor = "bottom";
      else if (anchor === "t") result.verticalAnchor = "top";

      if (querySelector(bodyPr, "a\\:spAutoFit, spAutoFit")) result.autoFit = "shape";
      else if (querySelector(bodyPr, "a\\:normAutofit, normAutofit")) {
        result.autoFit = "shrink";
        const normAutofit = querySelector(bodyPr, "a\\:normAutofit, normAutofit");
        const fontScale = normAutofit?.getAttribute("fontScale");
        if (fontScale) result.fontScale = parseInt(fontScale, 10) / 100000;
      }
      else if (querySelector(bodyPr, "a\\:noAutofit, noAutofit")) result.autoFit = "none";
    }
    return result;
  }

  private async parseElementsRecursive(
    parent: Element,
    transform: TransformState,
    imgRelMap: Record<string, string>,
    elements: SlideElement[],
    isTemplate: boolean = false,
    placeholderContext: PlaceholderContext | null = null
  ): Promise<void> {
    const children = Array.from(parent.childNodes) as Element[];
    for (const node of children) {
      if (node.nodeType !== 1) continue; // Only elements
      const tagName = node.localName || node.nodeName;

      if (tagName.endsWith("sp") || tagName.endsWith("cxnSp")) {
        // Shape or connection shape
        await this.parseShapeNode(node, transform, elements, isTemplate, placeholderContext);
      } else if (tagName.endsWith("pic")) {
        // Picture element
        await this.parsePicNode(node, transform, imgRelMap, elements, isTemplate);
      } else if (tagName.endsWith("grpSp")) {
        // Group shape - compute nested transform and recurse
        const grpSpPr = querySelector(node, "p\\:grpSpPr, grpSpPr");
        let childTransform = transform;
        if (grpSpPr) {
          const xfrm = querySelector(grpSpPr, "a\\:xfrm, xfrm");
          if (xfrm) {
            const off = querySelector(xfrm, "a\\:off, off");
            const ext = querySelector(xfrm, "a\\:ext, ext");
            const chOff = querySelector(xfrm, "a\\:chOff, chOff");
            const chExt = querySelector(xfrm, "a\\:chExt, chExt");

            if (off && ext && chOff && chExt) {
              const ox = parseInt(off.getAttribute("x") || "0", 10);
              const oy = parseInt(off.getAttribute("y") || "0", 10);
              const cx = parseInt(ext.getAttribute("cx") || "1", 10);
              const cy = parseInt(ext.getAttribute("cy") || "1", 10);
              const cox = parseInt(chOff.getAttribute("x") || "0", 10);
              const coy = parseInt(chOff.getAttribute("y") || "0", 10);
              const ccx = parseInt(chExt.getAttribute("cx") || "1", 10);
              const ccy = parseInt(chExt.getAttribute("cy") || "1", 10);

              const groupScaleX = cx / ccx;
              const groupScaleY = cy / ccy;
              const groupOffsetX = ox - cox * groupScaleX;
              const groupOffsetY = oy - coy * groupScaleY;

              childTransform = {
                scaleX: transform.scaleX * groupScaleX,
                scaleY: transform.scaleY * groupScaleY,
                absoluteUnitScale: transform.absoluteUnitScale,
                offsetX: transform.offsetX + groupOffsetX * transform.scaleX,
                offsetY: transform.offsetY + groupOffsetY * transform.scaleY
              };
            }
          }
        }
        await this.parseElementsRecursive(node, childTransform, imgRelMap, elements, isTemplate, placeholderContext);
      } else if (tagName.endsWith("graphicFrame")) {
        // Table graphic frame
        await this.parseGraphicFrameNode(node, transform, elements, isTemplate);
      }
    }
  }

  private parseRunProperties(
    rPr: Element,
    absoluteUnitScale: number,
    defaults: ComputedRunStyle
  ): ComputedRunStyle {
    const res = { ...defaults };
    const szAttr = rPr.getAttribute("sz");
    if (szAttr) {
      const ptSize = parseInt(szAttr, 10) / 100;
      res.fontSize = Math.round(ptSize * 12700 * absoluteUnitScale);
    }
    const boldAttr = rPr.getAttribute("b");
    if (boldAttr === "1" || boldAttr === "true") {
      res.bold = true;
    } else if (boldAttr === "0" || boldAttr === "false") {
      res.bold = false;
    }
    const italicAttr = rPr.getAttribute("i");
    if (italicAttr === "1" || italicAttr === "true") {
      res.italic = true;
    } else if (italicAttr === "0" || italicAttr === "false") {
      res.italic = false;
    }
    const latin = querySelector(rPr, "a\\:latin, latin");
    const typeface = latin?.getAttribute("typeface");
    if (typeface) res.fontFamily = this.resolveThemeTypeface(typeface);
    const eastAsian = querySelector(rPr, "a\\:ea, ea");
    const eastAsianTypeface = eastAsian?.getAttribute("typeface");
    if (eastAsianTypeface) res.eastAsianFontFamily = this.resolveThemeTypeface(eastAsianTypeface);
    const runColor = this.extractHexColor(rPr);
    if (runColor) {
      res.color = runColor;
    }
    return res;
  }

  private async parseShapeNode(
    node: Element,
    transform: TransformState,
    elements: SlideElement[],
    isTemplate: boolean = false,
    placeholderContext: PlaceholderContext | null = null
  ) {
    const nvSpPr = querySelector(node, "p\\:nvSpPr, nvSpPr");
    const nvPr = nvSpPr ? querySelector(nvSpPr, "p\\:nvPr, nvPr") : null;
    const ph = nvPr ? querySelector(nvPr, "p\\:ph, ph") : null;

    // Skip empty placeholder templates from master/layout
    if (isTemplate && ph) {
      return;
    }

    const cNvPr = nvSpPr ? querySelector(nvSpPr, "p\\:cNvPr, cNvPr") : null;
    const id = cNvPr?.getAttribute("id") || `shape_${Math.random().toString(36).substr(2, 9)}`;

    const shapeChain = this.resolvePlaceholderChain(node, placeholderContext);
    const spPrNodes = shapeChain
      .map(shape => querySelector(shape, "p\\:spPr, spPr"))
      .filter((spPr): spPr is globalThis.Element => spPr !== null);
    const spPr = spPrNodes[0];
    if (!spPr) return;

    const rect = spPrNodes
      .map(candidate => this.parseRect(candidate, transform))
      .find((candidate): candidate is Rect => candidate !== null) || null;
    if (!rect) return;

    // Check geometry
    const prstGeom = querySelector(spPr, "a\\:prstGeom, prstGeom");
    let shapeType: "rect" | "roundRect" | "ellipse" | "triangle" = "rect";
    if (prstGeom) {
      const prst = prstGeom.getAttribute("prst");
      if (prst === "ellipse" || prst === "oval") {
        shapeType = "ellipse";
      } else if (prst === "roundRect") {
        shapeType = "roundRect";
      } else if (prst === "triangle") {
        shapeType = "triangle";
      }
    }

    // Fill Color
    let fill = "#cccccc"; // Default shape fill
    const solidFill = querySelector(spPr, "a\\:solidFill, solidFill");
    const gradientFill = querySelector(spPr, "a\\:gradFill, gradFill");
    if (solidFill) {
      const color = this.extractHexColor(solidFill);
      if (color) fill = color;
    } else if (gradientFill) {
      const color = this.approximateGradientColor(gradientFill);
      if (color) fill = color;
    } else {
      // Check style reference
      const styleNode = querySelector(node, "p\\:style, style");
      const fillRef = styleNode ? querySelector(styleNode, "a\\:fillRef, fillRef") : null;
      if (fillRef) {
        const color = this.extractHexColor(fillRef);
        if (color) fill = color;
      }
    }

    // Border (Outline) Color and Width
    let border: Border | undefined = undefined;
    const ln = querySelector(spPr, "a\\:ln, ln");
    let borderColor: string | null = null;
    let emuWidth = 12700; // 1 pt default

    if (ln) {
      const lnNoFill = querySelector(ln, "a\\:noFill, noFill");
      if (!lnNoFill) {
        const lnSolidFill = querySelector(ln, "a\\:solidFill, solidFill");
        if (lnSolidFill) {
          borderColor = this.extractHexColor(lnSolidFill);
        }
        const wAttr = ln.getAttribute("w");
        if (wAttr) {
          emuWidth = parseInt(wAttr, 10);
        }
      }
    }

    // Fallback: Check shape style lnRef for border outline color
    if (!borderColor) {
      const styleNode = querySelector(node, "p\\:style, style");
      const lnRef = styleNode ? querySelector(styleNode, "a\\:lnRef, lnRef") : null;
      if (lnRef) {
        borderColor = this.extractHexColor(lnRef);
      }
    }

    if (borderColor) {
      const strokeWidth = Math.max(1, emuWidth * transform.absoluteUnitScale);
      border = { color: borderColor, width: strokeWidth };
    }

    // Fix: only look for noFill directly under spPr, not inside ln
    const noFill = querySelector(spPr, "a\\:noFill, noFill");
    const isShapeNoFill = noFill && (
      noFill.parentElement === spPr ||
      noFill.parentElement?.localName === "spPr" ||
      noFill.parentElement?.nodeName.endsWith("spPr")
    );
    const hasFill = !isShapeNoFill && (solidFill || gradientFill || querySelector(node, "p\\:style, style"));
    const hasBorder = border !== undefined;

    if (hasFill || hasBorder) {
      elements.push({
        type: "shape",
        id: `shape_bg_${id}`,
        rect: { ...rect },
        shapeType,
        fill: hasFill ? fill : "transparent",
        border
      });
    }

    // Default text color and font size from shape placeholder / style
    let defaultTextColor = "#333333";
    let defaultFontSize = 18; // Default in points (18pt)
    let defaultBold = false;
    let defaultFontFamily = "sans-serif";
    let defaultEastAsianFontFamily = "";

    // Check placeholder type
    if (ph) {
      const type = ph.getAttribute("type");
      if (type === "title" || type === "ctrTitle") {
        defaultFontSize = 40; // 40pt
        defaultBold = true;
      } else if (type === "subTitle") {
        defaultFontSize = 24; // 24pt
      } else if (type === "body") {
        defaultFontSize = 18; // 18pt
      }
    }

    // Check style node for fontRef
    const styleNode = querySelector(node, "p\\:style, style");
    if (styleNode) {
      const fontRef = querySelector(styleNode, "a\\:fontRef, fontRef");
      if (fontRef) {
        const color = this.extractHexColor(fontRef);
        if (color) defaultTextColor = color;
        const fontScheme = fontRef.getAttribute("idx");
        if (fontScheme === "major" || fontScheme === "minor") {
          const prefix = fontScheme === "major" ? "+mj" : "+mn";
          defaultFontFamily = this.resolveThemeTypeface(`${prefix}-lt`);
          defaultEastAsianFontFamily = this.themeFonts[`${prefix}-Hans`]
            || this.resolveThemeTypeface(`${prefix}-ea`);
        }
      }
    }

    const defaultFontSizePx = Math.round(defaultFontSize * 12700 * transform.absoluteUnitScale);

    // Text Content
    const textBodies = shapeChain
      .map(shape => querySelector(shape, "p\\:txBody, txBody"))
      .filter((txBody): txBody is globalThis.Element => txBody !== null);
    const txBody = textBodies[0] || null;
    if (txBody) {
      const inheritance: TextInheritanceContext = {
        textBodies,
        masterTextStyle: this.selectMasterTextStyle(
          placeholderContext?.masterTextStyles || null,
          this.getPlaceholderDescriptor(node)
        )
      };
      const body = this.resolveTextBodyProperties(textBodies, transform.scaleX, transform.scaleY);
      await this.parseTextBody(
        txBody,
        rect.x,
        rect.y,
        rect.w,
        rect.h,
        transform.absoluteUnitScale,
        transform.scaleY,
        id,
        elements,
        defaultTextColor,
        defaultFontSizePx,
        defaultBold,
        defaultFontFamily,
        defaultEastAsianFontFamily,
        inheritance,
        body
      );
    }
  }

  private async parseGraphicFrameNode(
    node: Element,
    transform: TransformState,
    elements: SlideElement[],
    isTemplate: boolean = false
  ) {
    // If it's a template placeholder graphicFrame, skip
    const nvGraphicFramePr = querySelector(node, "p\\:nvGraphicFramePr, nvGraphicFramePr");
    const nvPr = nvGraphicFramePr ? querySelector(nvGraphicFramePr, "p\\:nvPr, nvPr") : null;
    const ph = nvPr ? querySelector(nvPr, "p\\:ph, ph") : null;
    if (isTemplate && ph) {
      return;
    }

    const tbl = querySelector(node, "a\\:tbl, tbl");
    if (!tbl) return;

    // Get position of the graphic frame
    const xfrm = querySelector(node, "a\\:xfrm, xfrm");
    if (!xfrm) return;
    const off = querySelector(xfrm, "a\\:off, off");
    const ext = querySelector(xfrm, "a\\:ext, ext");
    if (!off || !ext) return null;

    const ox = parseInt(off.getAttribute("x") || "0", 10);
    const oy = parseInt(off.getAttribute("y") || "0", 10);
    const cx = parseInt(ext.getAttribute("cx") || "0", 10);
    const cy = parseInt(ext.getAttribute("cy") || "0", 10);

    const fx = transform.offsetX + ox * transform.scaleX;
    const fy = transform.offsetY + oy * transform.scaleY;
    const fcx = cx * transform.scaleX;
    const fcy = cy * transform.scaleY;

    // Read column widths from tblGrid
    const gridCols = querySelectorAll(tbl, "a\\:gridCol, gridCol");
    const colWidths = gridCols.map(col => parseInt(col.getAttribute("w") || "0", 10) * transform.scaleX);

    // Read rows
    const rows = querySelectorAll(tbl, "a\\:tr, tr");
    let currentY = fy;

    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      const hAttr = row.getAttribute("h");
      const rowHeight = hAttr ? parseInt(hAttr, 10) * transform.scaleY : (fcy / rows.length);

      const cells = querySelectorAll(row, "a\\:tc, tc");
      let currentX = fx;

      for (let c = 0; c < cells.length; c++) {
        const cell = cells[c];
        const cellWidth = colWidths[c] || (fcx / cells.length);

        // Check merge (gridSpan)
        const gridSpan = parseInt(cell.getAttribute("gridSpan") || "1", 10);
        let actualWidth = 0;
        for (let i = 0; i < gridSpan; i++) {
          actualWidth += colWidths[c + i] || cellWidth;
        }

        // Cell background fill
        const tcPr = querySelector(cell, "a\\:tcPr, tcPr");
        let fill = "transparent";
        let cellBorderColor = "#cccccc";

        if (tcPr) {
          const solidFill = querySelector(tcPr, "a\\:solidFill, solidFill");
          if (solidFill) {
            const color = this.extractHexColor(solidFill);
            if (color) fill = color;
          }
          // Border color styling (accent-based or custom)
          const lnL = querySelector(tcPr, "a\\:lnL, lnL");
          if (lnL) {
            const solidFillLn = querySelector(lnL, "a\\:solidFill, solidFill");
            if (solidFillLn) {
              const color = this.extractHexColor(solidFillLn);
              if (color) cellBorderColor = color;
            }
          }
        }

        // Add cell shape
        const cellId = `tbl_cell_${r}_${c}_${Math.random().toString(36).substr(2, 5)}`;
        elements.push({
          type: "shape",
          id: `shape_${cellId}`,
          rect: { x: currentX, y: currentY, w: actualWidth, h: rowHeight },
          shapeType: "rect",
          fill,
          border: { color: cellBorderColor, width: 1 }
        });

        // Cell text content
        const txBody = querySelector(cell, "a\\:txBody, txBody");
        if (txBody) {
          await this.parseTextBody(
            txBody,
            currentX,
            currentY,
            actualWidth,
            rowHeight,
            transform.absoluteUnitScale,
            transform.scaleY,
            cellId,
            elements,
            "#333333"
          );
        }

        currentX += actualWidth;
      }
      currentY += rowHeight;
    }
  }

  private getBulletChar(buCharNode: Element | null, buFontNode: Element | null): string {
    if (!buCharNode) return "•";
    const char = buCharNode.getAttribute("char") || "•";
    const font = buFontNode ? buFontNode.getAttribute("typeface") : "";
    if (font === "Wingdings" || font === "Wingdings 2" || font === "Wingdings 3") {
      if (char === "n") return "■";
      if (char === "u") return "◆";
      if (char === "o") return "○";
    }
    return char;
  }

  private async parseTextBody(
    txBody: Element,
    x: number,
    y: number,
    w: number,
    h: number,
    absoluteUnitScale: number,
    layoutScaleY: number,
    id: string,
    elements: SlideElement[],
    defaultTextColor: string = "#333333",
    defaultFontSizePx: number = 20,
    defaultBold: boolean = false,
    defaultFontFamily: string = "sans-serif",
    defaultEastAsianFontFamily: string = "",
    inheritance: TextInheritanceContext | null = null,
    resolvedBody: TextBodyProperties | null = null
  ) {
    const paragraphs = querySelectorAll(txBody, "a\\:p, p");
    const body = resolvedBody || {
      marginLeft: 8,
      marginRight: 8,
      marginTop: 4,
      marginBottom: 4,
      verticalAnchor: "top" as const,
      autoFit: "none" as const,
      fontScale: 1
    };
    const computedParagraphs: TextParagraph[] = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const pPr = querySelector(p, "a\\:pPr, pPr");
      const level = Math.max(0, Math.min(8, parseInt(pPr?.getAttribute("lvl") || "0", 10)));
      let pDefaults: ComputedRunStyle = {
        fontSize: defaultFontSizePx,
        color: defaultTextColor,
        bold: defaultBold,
        fontFamily: defaultFontFamily,
        eastAsianFontFamily: defaultEastAsianFontFamily,
        italic: false
      };

      const inheritedParagraphProperties: globalThis.Element[] = [];
      if (inheritance?.masterTextStyle) {
        const masterLevel = this.getLevelParagraphProperties(inheritance.masterTextStyle, level);
        if (masterLevel) inheritedParagraphProperties.push(masterLevel);
      }
      if (inheritance) {
        for (const inheritedBody of [...inheritance.textBodies].reverse()) {
          const levelProperties = this.getLevelParagraphProperties(inheritedBody, level);
          if (levelProperties) inheritedParagraphProperties.push(levelProperties);
        }
      }
      if (pPr) inheritedParagraphProperties.push(pPr);

      let align: "left" | "center" | "right" = "left";
      let marginLeft = 0;
      let indent = 0;
      let lineSpacing: ParagraphStyle["lineSpacing"];
      let spaceBefore = 0;
      let spaceAfter = 0;

      const readSpacing = (
        properties: globalThis.Element,
        selector: string
      ): { unit: "percent" | "points"; value: number } | null => {
        const spacing = querySelector(properties, selector);
        if (!spacing) return null;
        const percent = querySelector(spacing, "a\\:spcPct, spcPct")?.getAttribute("val");
        if (percent !== null && percent !== undefined) {
          return { unit: "percent", value: parseInt(percent, 10) / 100000 };
        }
        const points = querySelector(spacing, "a\\:spcPts, spcPts")?.getAttribute("val");
        if (points !== null && points !== undefined) {
          return { unit: "points", value: (parseInt(points, 10) / 100) * 12700 * absoluteUnitScale };
        }
        return null;
      };

      for (const inheritedPPr of inheritedParagraphProperties) {
        const algnAttr = inheritedPPr.getAttribute("algn");
        if (algnAttr === "ctr") align = "center";
        else if (algnAttr === "r") align = "right";
        else if (algnAttr === "l" || algnAttr === "just" || algnAttr === "dist") align = "left";

        const marL = inheritedPPr.getAttribute("marL");
        const indentAttr = inheritedPPr.getAttribute("indent");
        if (marL !== null) marginLeft = parseInt(marL, 10) * layoutScaleY;
        if (indentAttr !== null) indent = parseInt(indentAttr, 10) * layoutScaleY;

        const defRPr = querySelector(inheritedPPr, "a\\:defRPr, defRPr");
        if (defRPr) {
          pDefaults = this.parseRunProperties(defRPr, absoluteUnitScale, pDefaults);
        }

        const inheritedLineSpacing = readSpacing(inheritedPPr, "a\\:lnSpc, lnSpc");
        if (inheritedLineSpacing) lineSpacing = inheritedLineSpacing;
        const inheritedSpaceBefore = readSpacing(inheritedPPr, "a\\:spcBef, spcBef");
        if (inheritedSpaceBefore) {
          spaceBefore = inheritedSpaceBefore.unit === "points"
            ? inheritedSpaceBefore.value
            : pDefaults.fontSize * inheritedSpaceBefore.value;
        }
        const inheritedSpaceAfter = readSpacing(inheritedPPr, "a\\:spcAft, spcAft");
        if (inheritedSpaceAfter) {
          spaceAfter = inheritedSpaceAfter.unit === "points"
            ? inheritedSpaceAfter.value
            : pDefaults.fontSize * inheritedSpaceAfter.value;
        }
      }

      // Check if bullet exists
      let bulletChar = "";
      let bulletColor = pDefaults.color;

      for (const inheritedPPr of inheritedParagraphProperties) {
        if (querySelector(inheritedPPr, "a\\:buNone, buNone")) {
          bulletChar = "";
        }
        const buChar = querySelector(inheritedPPr, "a\\:buChar, buChar");
        if (buChar) {
          const buFont = querySelector(inheritedPPr, "a\\:buFont, buFont");
          bulletChar = this.getBulletChar(buChar, buFont);

          const buClr = querySelector(inheritedPPr, "a\\:buClr, buClr");
          if (buClr) {
            const color = this.extractHexColor(buClr);
            if (color) bulletColor = color;
          }
        }
      }

      const computedRuns: TextRun[] = [];
      const runs = querySelectorAll(p, "a\\:r, r, a\\:fld, fld, a\\:br, br");
      for (const run of runs) {
        const runName = run.localName || run.nodeName;
        let runStyle = pDefaults;
        const rPr = querySelector(run, "a\\:rPr, rPr");
        if (rPr) runStyle = this.parseRunProperties(rPr, absoluteUnitScale, pDefaults);

        const content = runName.endsWith("br")
          ? "\n"
          : querySelector(run, "a\\:t, t")?.textContent || "";
        if (content) {
          computedRuns.push({
            content,
            style: {
              fontSize: runStyle.fontSize,
              color: runStyle.color,
              bold: runStyle.bold,
              align,
              fontFamily: runStyle.fontFamily,
              eastAsianFontFamily: runStyle.eastAsianFontFamily || undefined,
              italic: runStyle.italic
            }
          });
        }
      }

      if (computedRuns.length === 0) {
        const endParaRPr = querySelector(p, "a\\:endParaRPr, endParaRPr");
        const emptyStyle = endParaRPr
          ? this.parseRunProperties(endParaRPr, absoluteUnitScale, pDefaults)
          : pDefaults;
        computedRuns.push({
          content: "",
          style: {
            fontSize: emptyStyle.fontSize,
            color: emptyStyle.color,
            bold: emptyStyle.bold,
            align,
            fontFamily: emptyStyle.fontFamily,
            eastAsianFontFamily: emptyStyle.eastAsianFontFamily || undefined,
            italic: emptyStyle.italic
          }
        });
      }

      computedParagraphs.push({
        style: {
          align,
          level,
          marginLeft,
          indent,
          lineSpacing,
          spaceBefore,
          spaceAfter
        },
        bullet: bulletChar ? { char: bulletChar, color: bulletColor } : undefined,
        runs: computedRuns
      });
    }

    if (computedParagraphs.length === 0) return;
    const fallbackRun = computedParagraphs
      .flatMap(paragraph => paragraph.runs)
      .find(run => run.content.length > 0) || computedParagraphs[0].runs[0];
    const content = computedParagraphs
      .map(paragraph => paragraph.runs.map(run => run.content).join(""))
      .join("\n");

    elements.push({
      type: "text",
      id: `text_${id}`,
      rect: { x, y, w, h },
      content,
      style: fallbackRun.style,
      paragraphs: computedParagraphs,
      body
    });
  }

  private async parsePicNode(
    node: Element,
    transform: TransformState,
    imgRelMap: Record<string, string>,
    elements: SlideElement[],
    isTemplate: boolean = false
  ) {
    // If it's a template placeholder picture, skip
    const nvPicPr = querySelector(node, "p\\:nvPicPr, nvPicPr");
    const nvPr = nvPicPr ? querySelector(nvPicPr, "p\\:nvPr, nvPr") : null;
    const ph = nvPr ? querySelector(nvPr, "p\\:ph, ph") : null;
    if (isTemplate && ph) {
      return;
    }

    const cNvPr = nvPicPr ? querySelector(nvPicPr, "p\\:cNvPr, cNvPr") : null;
    const id = cNvPr?.getAttribute("id") || `pic_${Math.random().toString(36).substr(2, 9)}`;

    const spPr = querySelector(node, "p\\:spPr, spPr");
    if (!spPr) return;

    const rect = this.parseRect(spPr, transform);
    if (!rect) return;

    const blipFill = querySelector(node, "p\\:blipFill, blipFill");
    if (!blipFill) return;

    const blip = querySelector(blipFill, "a\\:blip, blip");
    if (!blip) return;

    const embedId = blip.getAttribute("r:embed") || blip.getAttribute("embed");
    if (!embedId || !imgRelMap[embedId]) return;

    const imgPath = imgRelMap[embedId];
    const imgZipFile = this.zip?.file(imgPath);
    if (!imgZipFile) return;

    // Convert zip file image data to a Blob URL
    const blob = await imgZipFile.async("blob");
    const blobUrl = URL.createObjectURL(blob);
    const srcRect = querySelector(blipFill, "a\\:srcRect, srcRect");
    const crop = srcRect ? {
      left: parseInt(srcRect.getAttribute("l") || "0", 10) / 100000,
      top: parseInt(srcRect.getAttribute("t") || "0", 10) / 100000,
      right: parseInt(srcRect.getAttribute("r") || "0", 10) / 100000,
      bottom: parseInt(srcRect.getAttribute("b") || "0", 10) / 100000
    } : undefined;

    elements.push({
      type: "image",
      id: `img_${id}`,
      rect,
      url: blobUrl,
      crop
    });
  }

  private parseRect(spPr: Element, transform: TransformState): Rect | null {
    const xfrm = querySelector(spPr, "a\\:xfrm, xfrm");
    if (!xfrm) return null;

    const off = querySelector(xfrm, "a\\:off, off");
    const ext = querySelector(xfrm, "a\\:ext, ext");
    if (!off || !ext) return null;

    const x = parseInt(off.getAttribute("x") || "0", 10);
    const y = parseInt(off.getAttribute("y") || "0", 10);
    const cx = parseInt(ext.getAttribute("cx") || "0", 10);
    const cy = parseInt(ext.getAttribute("cy") || "0", 10);

    return {
      x: transform.offsetX + x * transform.scaleX,
      y: transform.offsetY + y * transform.scaleY,
      w: cx * transform.scaleX,
      h: cy * transform.scaleY
    };
  }
}

// Coordinate transform state interface
export interface TransformState {
  // Cumulative group-space transform for geometry coordinates.
  scaleX: number;
  scaleY: number;
  // Page-level conversion for pt/EMU visual properties; never multiplied by group scale.
  absoluteUnitScale: number;
  offsetX: number;
  offsetY: number;
}
