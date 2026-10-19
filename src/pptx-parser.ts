import { parseChartXml } from "./charts/parse";
import { TableCell, TableElement, TableModel } from "./tables/model";
import {
  DEFAULT_TABLE_STYLE_ID,
  StyleBorder,
  TableStyleDef,
  TableStylePart,
  TableStylePartName,
  builtinDefaultStyle,
  resolveCellStyle
} from "./tables/styles";
import JSZip from "jszip";

import {
  Border,
  ComputedRunStyle,
  ComputedShapeStyle,
  EffectStyle,
  ElementSource,
  GroupRef,
  FillStyle,
  GradientStop,
  ImageElement,
  LineStyle,
  ParagraphStyle,
  PlaceholderContext,
  PlaceholderDescriptor,
  PptVirtualDocument,
  PresentationSize,
  Rect,
  ReflectionStyle,
  ShapeElement,
  ShapeStyleProperty,
  Slide,
  SlideElement,
  StyleRegistry,
  StyleRule,
  StyleSourceKind,
  TextBodyProperties,
  TextElement,
  TextInheritanceContext,
  TextParagraph,
  TextRun,
  TextStyle,
  ThemeStyleMatrix,
  TransformState
} from "./pptx-virtual-dom";
import {
  getDirectChild,
  getDirectChildren,
  hasLocalName,
  querySelector,
  querySelectorAll
} from "./pptx-xml";
import { PptxStyleResolver } from "./pptx-style-resolver";
import { parseCustomGeometry } from "./pptx-custom-geometry";

export * from "./pptx-virtual-dom";

/** What East Asian text is set in when neither the theme nor the run names a font. */
export const DEFAULT_EAST_ASIAN_FONT = "Microsoft YaHei";

export class PptxParser {
  private zip: JSZip | null = null;
  private currentSlidePart = "";
  /** Table styles of the deck, by style id. */
  private tableStyles = new Map<string, TableStyleDef>();
  /** Chart parts the slide being parsed refers to, by relationship id. */
  private currentChartRels: Record<string, string> = {};
  /** presentation.xml defaultTextStyle: the bottom of every placeholder's inheritance chain. */
  private defaultTextStyle: globalThis.Element | null = null;
  private groupStack: GroupRef[] = [];
  private styleRegistry: StyleRegistry = { rules: {} };
  private nextStyleRuleId = 1;
  private themeStyleMatrix: ThemeStyleMatrix = { fills: [], backgroundFills: [], lines: [], effects: [] };
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

  private get styleResolver(): PptxStyleResolver {
    return new PptxStyleResolver({
      themeColors: this.themeColors,
      themeFonts: this.themeFonts,
      themeStyleMatrix: this.themeStyleMatrix,
      registerStyleRule: rule => this.registerStyleRule(rule)
    });
  }

  /** The theme's six accent colours, in order. */
  public themeAccents(): string[] {
    return ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map(name => this.themeColors[name]);
  }

  /** The opened package, kept so edits can be written back on export. */
  public getZip(): JSZip | null {
    return this.zip;
  }

  /**
   * Parse a slide that exists only in the editing session. Its files are put
   * into the package just long enough to be read like any other slide.
   */
  public async parseVirtualSlide(
    part: string,
    xml: string,
    relsXml: string,
    slideId: string,
    size: PresentationSize
  ): Promise<Slide> {
    if (!this.zip) throw new Error("No package is open.");
    const slash = part.lastIndexOf("/");
    const relsPath = `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`;
    this.zip.file(part, xml);
    this.zip.file(relsPath, relsXml);
    try {
      return await this.parseSlide(part, slideId, size);
    } finally {
      this.zip.remove(part);
      this.zip.remove(relsPath);
    }
  }

  /** Forget the opened package (a non-PPTX document replaced it). */
  public discardPackage(): void {
    this.zip = null;
  }

  private makeSource(
    isTemplate: boolean,
    shapeId: string | null | undefined,
    transform: TransformState,
    name?: string | null
  ): ElementSource | undefined {
    if (isTemplate || !shapeId || !this.currentSlidePart) return undefined;
    return {
      part: this.currentSlidePart,
      shapeId,
      ...(name ? { name } : {}),
      xform: {
        scaleX: transform.scaleX,
        scaleY: transform.scaleY,
        offsetX: transform.offsetX,
        offsetY: transform.offsetY
      },
      ...(this.groupStack.length ? { groups: [...this.groupStack] } : {})
    };
  }

  public async parse(buffer: ArrayBuffer): Promise<PptVirtualDocument> {
    this.zip = await JSZip.loadAsync(buffer);
    this.styleRegistry = { rules: {} };
    this.nextStyleRuleId = 1;

    // 1. Parse Theme Colors
    await this.parseThemeColors();
    await this.parseThemeFonts();
    await this.parseThemeStyleMatrix();
    await this.parseTableStyles();

    // 2. Parse Presentation Properties (slide size)
    const size = await this.parsePresentationSize();
    await this.parseDefaultTextStyle();

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

    return { size, slides, styleRegistry: this.styleRegistry };
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
        const hex = this.styleResolver.extractHexColor(colorNode);
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

  private async parseThemeStyleMatrix() {
    this.themeStyleMatrix = { fills: [], backgroundFills: [], lines: [], effects: [] };
    const themeFile = this.zip?.file("ppt/theme/theme1.xml");
    if (!themeFile) return;

    const xmlText = await themeFile.async("text");
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");
    const formatScheme = querySelector(doc, "a\\:fmtScheme, fmtScheme");
    if (!formatScheme) return;

    const fillList = getDirectChild(formatScheme, "fillStyleLst");
    const backgroundFillList = getDirectChild(formatScheme, "bgFillStyleLst");
    const lineList = getDirectChild(formatScheme, "lnStyleLst");
    const effectList = getDirectChild(formatScheme, "effectStyleLst");
    this.themeStyleMatrix.fills = fillList ? getDirectChildren(fillList, "solidFill", "gradFill", "noFill", "pattFill") : [];
    this.themeStyleMatrix.backgroundFills = backgroundFillList
      ? getDirectChildren(backgroundFillList, "solidFill", "gradFill", "noFill", "pattFill")
      : [];
    this.themeStyleMatrix.lines = lineList ? getDirectChildren(lineList, "ln") : [];
    this.themeStyleMatrix.effects = effectList
      ? getDirectChildren(effectList, "effectStyle")
      : [];
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

  private extractHexColor(node: Element, placeholderColor?: string | null): string | null {
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

    // 3. preset colors, commonly used by shadows (for example prstClr="black")
    if (!hex) {
      const presetNode = querySelector(node, "a\\:prstClr, prstClr");
      if (presetNode) {
        const presetColors: Record<string, string> = {
          black: "#000000",
          white: "#ffffff",
          red: "#ff0000",
          green: "#008000",
          blue: "#0000ff",
          yellow: "#ffff00",
          gray: "#808080",
          grey: "#808080",
          dkGray: "#404040",
          ltGray: "#c0c0c0"
        };
        const val = presetNode.getAttribute("val") || "";
        if (presetColors[val]) {
          hex = presetColors[val];
          colorNode = presetNode;
        }
      }
    }

    // 4. scRGB colors use percentages rather than 8-bit channels.
    if (!hex) {
      const scRgbNode = querySelector(node, "a\\:scrgbClr, scrgbClr");
      if (scRgbNode) {
        const channel = (name: string) => Math.round(
          Math.max(0, Math.min(100000, parseInt(scRgbNode.getAttribute(name) || "0", 10))) * 255 / 100000
        );
        hex = `#${channel("r").toString(16).padStart(2, "0")}${channel("g").toString(16).padStart(2, "0")}${channel("b").toString(16).padStart(2, "0")}`;
        colorNode = scRgbNode;
      }
    }

    // 5. schemeClr (nested reference)
    if (!hex) {
      const schemeNode = querySelector(node, "a\\:schemeClr, schemeClr");
      if (schemeNode) {
        const val = schemeNode.getAttribute("val");
        if (val === "phClr" && placeholderColor) {
          hex = placeholderColor;
          colorNode = schemeNode;
        } else if (val && this.themeColors[val]) {
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

  private extractOpacity(node: Element): number {
    const alpha = querySelector(node, "a\\:alpha, alpha");
    const value = alpha?.getAttribute("val");
    return value ? Math.max(0, Math.min(1, parseInt(value, 10) / 100000)) : 1;
  }

  private applyColorOpacity(color: string, opacity: number): string {
    if (opacity >= 1 || !/^#[0-9a-f]{6}$/i.test(color)) return color;
    const r = parseInt(color.slice(1, 3), 16);
    const g = parseInt(color.slice(3, 5), 16);
    const b = parseInt(color.slice(5, 7), 16);
    return `rgba(${r},${g},${b},${opacity})`;
  }

  private parseFillStyle(container: Element, placeholderColor?: string | null): FillStyle | undefined {
    return this.styleResolver.parseFillStyle(container, placeholderColor);
  }

  private parseLineStyle(
    line: Element,
    absoluteUnitScale: number,
    placeholderColor?: string | null
  ): LineStyle | undefined {
    const fill = this.parseFillStyle(line, placeholderColor);
    if (fill?.type === "none") return undefined;
    const resolvedFill = fill || (placeholderColor ? { type: "solid" as const, color: placeholderColor } : undefined);
    if (!resolvedFill) return undefined;

    const width = Math.max(1, parseInt(line.getAttribute("w") || "12700", 10) * absoluteUnitScale);
    const dashNode = getDirectChild(line, "prstDash");
    const joinNode = getDirectChild(line, "round", "bevel", "miter");
    const cap = line.getAttribute("cap");
    return {
      fill: resolvedFill,
      width,
      dash: dashNode?.getAttribute("val") || undefined,
      cap: cap === "rnd" ? "round" : cap === "sq" ? "square" : cap === "flat" ? "butt" : undefined,
      join: joinNode ? (joinNode.localName || joinNode.nodeName.replace(/^.*:/, "")) : undefined,
      compound: line.getAttribute("cmpd") || undefined
    };
  }

  private parseEffectStyle(
    container: Element,
    absoluteUnitScale: number,
    placeholderColor?: string | null
  ): EffectStyle | undefined {
    return this.styleResolver.parseEffectStyle(container, absoluteUnitScale, placeholderColor);
  }

  private registerStyleRule(rule: Omit<StyleRule, "id">): string {
    const id = `style_${this.nextStyleRuleId++}`;
    this.styleRegistry.rules[id] = { id, ...rule };
    return id;
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

  private async parseDefaultTextStyle(): Promise<void> {
    this.defaultTextStyle = null;
    const file = this.zip?.file("ppt/presentation.xml");
    if (!file) return;
    const doc = new DOMParser().parseFromString(await file.async("text"), "application/xml");
    this.defaultTextStyle = querySelector(doc, "p\\:defaultTextStyle, defaultTextStyle") as unknown as globalThis.Element | null;
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
    chartRelMap: Record<string, string>;
    layoutPath?: string;
    masterPath?: string;
  }> {
    const relMap: Record<string, string> = {};
    const chartRelMap: Record<string, string> = {};
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
          } else if (type.endsWith("/chart")) {
            chartRelMap[id] = this.resolveRelativePath(baseDir, target);
          } else if (type.includes("slideLayout")) {
            layoutPath = this.resolveRelativePath(baseDir, target);
          } else if (type.includes("slideMaster")) {
            masterPath = this.resolveRelativePath(baseDir, target);
          }
        }
      }
    }
    return { imgRelMap: relMap, chartRelMap, layoutPath, masterPath };
  }

  private async parseSlide(slidePath: string, slideId: string, viewSize: PresentationSize): Promise<Slide> {
    const slideFile = this.zip?.file(slidePath);
    if (!slideFile) throw new Error(`Slide file not found: ${slidePath}`);

    this.currentSlidePart = slidePath;
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
    const { imgRelMap, chartRelMap, layoutPath } = await this.loadRelationships(slideRelsPath, slideDir);
    this.currentChartRels = chartRelMap;

    let layoutImgRelMap: Record<string, string> = {};
    let masterImgRelMap: Record<string, string> = {};
    let layoutDoc: Document | null = null;
    let masterDoc: Document | null = null;
    let masterPart: string | null = null;

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
        masterPart = masterPath;
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

    this.applyColorMap(masterDoc, layoutDoc, doc);

    const elements: SlideElement[] = [];
    const placeholderContext: PlaceholderContext = {
      layoutShapes: layoutDoc ? this.collectPlaceholderShapes(layoutDoc) : [],
      masterShapes: masterDoc ? this.collectPlaceholderShapes(masterDoc) : [],
      masterTextStyles: masterDoc
        ? querySelector(masterDoc, "p\\:txStyles, txStyles") as unknown as globalThis.Element | null
        : null,
      slidePart: slidePath,
      layoutPart: layoutPath || null,
      masterPart
    };

    // Initialize coordinate transform (EMU to logical pixels)
    const initialTransform: TransformState = {
      scaleX: scaleX,
      scaleY: scaleY,
      absoluteUnitScale: Math.min(scaleX, scaleY),
      offsetX: 0,
      offsetY: 0
    };

    // The page's own background, else the layout's, else the master's, lies under everything.
    const background = await this.parseBackground(
      [[doc, imgRelMap], [layoutDoc, layoutImgRelMap], [masterDoc, masterImgRelMap]]
    );
    if (background) {
      elements.push({
        type: "shape",
        id: "slide_background",
        rect: { x: 0, y: 0, w: viewSize.width, h: viewSize.height },
        shapeType: "rect",
        fill: background.type === "solid" ? background.color : background.type === "gradient" ? background.stops[0].color : "transparent",
        computedStyle: { fill: background }
      });
    }

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

    return {
      id: slideId,
      elements,
      rawXml: xmlText,
      emuPerPx: 1 / Math.min(scaleX, scaleY),
      part: slidePath,
      layout: layoutDoc
        ? querySelector(layoutDoc, "p\\:cSld, cSld")?.getAttribute("name") || undefined
        : undefined,
      xform: { scaleX, scaleY }
    };
  }

  /**
   * Which theme colour "bg1", "tx1", "bg2" and "tx2" stand for on this slide:
   * the master's colour map, unless the layout or the slide overrides it. A
   * dark master maps bg1 to the theme's dark colour and tx1 to its light one.
   */
  private applyColorMap(master: Document | null, layout: Document | null, slide: Document): void {
    const identity: Record<string, string> = { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" };
    const read = (doc: Document | null): Record<string, string> | null => {
      const map = doc && querySelector(doc, "p\\:clrMap, clrMap");
      if (!map) return null;
      return Object.fromEntries(Object.keys(identity).map(name => [name, map.getAttribute(name) || identity[name]]));
    };
    const override = (doc: Document | null): Record<string, string> | null => {
      const node = doc && querySelector(doc, "p\\:clrMapOvr, clrMapOvr");
      const explicit = node && getDirectChild(node, "overrideClrMapping");
      if (!explicit) return null;
      return Object.fromEntries(Object.keys(identity).map(name => [name, explicit.getAttribute(name) || identity[name]]));
    };
    const map = override(slide) ?? override(layout) ?? read(master) ?? identity;
    for (const [name, source] of Object.entries(map)) {
      const color = this.themeColors[source];
      if (color) this.themeColors[name] = color;
    }
  }

  /** The fill of the first `p:bg` found, looking at the slide, then its layout, then its master. */
  private async parseBackground(levels: Array<[Document | null, Record<string, string>]>): Promise<FillStyle | null> {
    for (const [doc, relMap] of levels) {
      const cSld = doc && querySelector(doc, "p\\:cSld, cSld");
      const bg = cSld && getDirectChild(cSld, "bg");
      if (!bg) continue;

      const properties = getDirectChild(bg, "bgPr");
      if (properties) {
        let fill = this.styleResolver.parseFillStyle(properties);
        if (fill?.type === "picture" && fill.embed && relMap[fill.embed]) {
          const file = this.zip?.file(relMap[fill.embed]);
          if (file) fill = { ...fill, url: URL.createObjectURL(await file.async("blob")) };
        }
        if (fill) return fill;
        continue;
      }

      // A reference into the theme's background styles, painted with the colour given beside it.
      const reference = getDirectChild(bg, "bgRef");
      if (reference) {
        const index = parseInt(reference.getAttribute("idx") || "0", 10);
        const color = this.extractHexColor(reference);
        const source = index >= 1000 ? this.themeStyleMatrix.backgroundFills[index - 1001] : this.themeStyleMatrix.fills[index - 1];
        const fill = index <= 0
          ? { type: "none" as const }
          : source
            ? this.parseFillStyle(source, color)
            : (color ? { type: "solid" as const, color } : undefined);
        if (fill) return fill;
      }
    }
    return null;
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

  /**
   * Placeholder kinds that share a master placeholder: a centred title uses the
   * master's title, and subtitles and content placeholders use its body.
   */
  private placeholderClass(type: string): string {
    if (type === "ctrTitle") return "title";
    if (type === "subTitle" || type === "obj") return "body";
    return type;
  }

  private findMatchingPlaceholder(
    descriptor: PlaceholderDescriptor,
    candidates: globalThis.Element[],
    matchByIndex = true
  ): globalThis.Element | null {
    if (matchByIndex && descriptor.idx !== null) {
      const byIndex = candidates.find(candidate =>
        this.getPlaceholderDescriptor(candidate)?.idx === descriptor.idx
      );
      if (byIndex) return byIndex;
    }

    const exact = candidates.find(candidate =>
      this.getPlaceholderDescriptor(candidate)?.type === descriptor.type
    );
    if (exact) return exact;
    const kind = this.placeholderClass(descriptor.type);
    return candidates.find(candidate => {
      const type = this.getPlaceholderDescriptor(candidate)?.type;
      return type !== undefined && this.placeholderClass(type) === kind;
    }) || null;
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
    // The master is matched by kind only: its indices belong to the master itself.
    const masterShape = this.findMatchingPlaceholder(masterDescriptor, context.masterShapes, false);
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
      wrap: "square",
      verticalOverflow: "overflow",
      horizontalOverflow: "overflow",
      fontScale: 1,
      textDirection: "horz"
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

      const textDirection = bodyPr.getAttribute("vert");
      if (textDirection) {
        result.textDirection = textDirection as NonNullable<TextBodyProperties["textDirection"]>;
      }

      const wrap = bodyPr.getAttribute("wrap");
      if (wrap === "none" || wrap === "square") {
        result.wrap = wrap;
      }

      if (querySelector(bodyPr, "a\\:spAutoFit, spAutoFit")) {
        result.autoFit = "shape";
      }
      else if (querySelector(bodyPr, "a\\:normAutofit, normAutofit")) {
        result.autoFit = "shrink";
        const normAutofit = querySelector(bodyPr, "a\\:normAutofit, normAutofit");
        const fontScale = normAutofit?.getAttribute("fontScale");
        if (fontScale) result.fontScale = parseInt(fontScale, 10) / 100000;
      }
      else if (querySelector(bodyPr, "a\\:noAutofit, noAutofit")) result.autoFit = "none";

      const verticalOverflow = bodyPr.getAttribute("vertOverflow");
      if (verticalOverflow === "clip" || verticalOverflow === "ellipsis" || verticalOverflow === "overflow") {
        result.verticalOverflow = verticalOverflow;
      }
      const horizontalOverflow = bodyPr.getAttribute("horzOverflow");
      if (horizontalOverflow === "clip" || horizontalOverflow === "ellipsis" || horizontalOverflow === "overflow") {
        result.horizontalOverflow = horizontalOverflow;
      }
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
        await this.parseShapeNode(node, transform, elements, imgRelMap, isTemplate, placeholderContext);
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
        const groupNode = querySelector(node, "p\\:cNvPr, cNvPr");
        const groupId = groupNode?.getAttribute("id");
        const group: GroupRef | null = !isTemplate && groupId
          ? {
              shapeId: groupId,
              ...(groupNode?.getAttribute("name") ? { name: groupNode.getAttribute("name")! } : {}),
              xform: {
                scaleX: transform.scaleX,
                scaleY: transform.scaleY,
                offsetX: transform.offsetX,
                offsetY: transform.offsetY
              },
              origRect: { x: 0, y: 0, w: 0, h: 0 }
            }
          : null;
        const firstMember = elements.length;
        if (group) this.groupStack.push(group);
        try {
          await this.parseElementsRecursive(node, childTransform, imgRelMap, elements, isTemplate, placeholderContext);
        } finally {
          if (group) this.groupStack.pop();
        }
        if (group) {
          const members = elements.slice(firstMember);
          if (members.length > 0) {
            const left = Math.min(...members.map(m => m.rect.x));
            const top = Math.min(...members.map(m => m.rect.y));
            const right = Math.max(...members.map(m => m.rect.x + m.rect.w));
            const bottom = Math.max(...members.map(m => m.rect.y + m.rect.h));
            group.origRect = { x: left, y: top, w: right - left, h: bottom - top };
          }
        }
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
    const spacingAttr = rPr.getAttribute("spc");
    if (spacingAttr) {
      res.letterSpacing = (parseInt(spacingAttr, 10) / 100) * 12700 * absoluteUnitScale;
    }
    const language = rPr.getAttribute("lang") || rPr.getAttribute("altLang") || undefined;
    const latin = querySelector(rPr, "a\\:latin, latin");
    const typeface = latin?.getAttribute("typeface");
    if (typeface) res.fontFamily = this.styleResolver.resolveThemeTypeface(typeface, language);
    const eastAsian = querySelector(rPr, "a\\:ea, ea");
    const eastAsianTypeface = eastAsian?.getAttribute("typeface");
    if (eastAsianTypeface) {
      res.eastAsianFontFamily = this.styleResolver.resolveThemeTypeface(eastAsianTypeface, language);
    }
    const reflection = querySelector(rPr, "a\\:reflection, reflection");
    if (reflection) {
      res.reflection = {
        blurRadius: parseInt(reflection.getAttribute("blurRad") || "0", 10) * absoluteUnitScale,
        startAlpha: Math.max(0, Math.min(1, parseInt(reflection.getAttribute("stA") || "0", 10) / 100000)),
        endAlpha: Math.max(0, Math.min(1, parseInt(reflection.getAttribute("endA") || "0", 10) / 100000)),
        endPosition: Math.max(0, Math.min(1, parseInt(reflection.getAttribute("endPos") || "100000", 10) / 100000)),
        direction: parseInt(reflection.getAttribute("dir") || "0", 10) / 60000,
        distance: parseInt(reflection.getAttribute("dist") || "0", 10) * absoluteUnitScale,
        scaleY: parseInt(reflection.getAttribute("sy") || "-100000", 10) / 100000,
        alignment: reflection.getAttribute("algn") || undefined,
        rotationWithShape: reflection.getAttribute("rotWithShape") === "1"
          || reflection.getAttribute("rotWithShape") === "true"
      } satisfies ReflectionStyle;
    }
    const textEffects = this.styleResolver.parseEffectStyle(rPr, absoluteUnitScale);
    if (textEffects?.outerShadow) {
      res.shadow = textEffects.outerShadow;
    }
    const outlineNode = getDirectChild(rPr, "ln");
    if (outlineNode) {
      const outline = this.styleResolver.parseLineStyle(outlineNode, absoluteUnitScale, res.color);
      if (outline?.fill.type === "solid") {
        res.outline = { color: outline.fill.color, width: outline.width };
      }
    }
    // A run's fill and its effect list are siblings. Reading the first color
    // anywhere below rPr can accidentally select outerShdw's color instead of
    // the text fill.
    const runColor = this.styleResolver.extractDirectFillColorWithOpacity(rPr);
    if (runColor) {
      res.color = runColor;
    }
    return res;
  }

  private resolveThemeStyleRules(
    styleNode: Element | null,
    absoluteUnitScale: number
  ): Array<{ id: string; property: ShapeStyleProperty; value: FillStyle | LineStyle | EffectStyle }> {
    if (!styleNode) return [];
    const result: Array<{ id: string; property: ShapeStyleProperty; value: FillStyle | LineStyle | EffectStyle }> = [];

    const fillRef = getDirectChild(styleNode, "fillRef");
    if (fillRef) {
      const index = parseInt(fillRef.getAttribute("idx") || "0", 10);
      const placeholderColor = this.extractHexColor(fillRef);
      const source = index >= 1000
        ? this.themeStyleMatrix.backgroundFills[index - 1001]
        : this.themeStyleMatrix.fills[index - 1];
      const fill = index <= 0
        ? { type: "none" as const }
        : source
          ? this.parseFillStyle(source, placeholderColor)
          : (placeholderColor ? { type: "solid" as const, color: placeholderColor } : undefined);
      if (fill) {
        const id = this.registerStyleRule({
          source: { kind: "theme", part: "ppt/theme/theme1.xml", nodeId: `fillRef:${index}` },
          parents: [],
          declarations: { fill }
        });
        result.push({ id, property: "fill", value: fill });
      }
    }

    const lineRef = getDirectChild(styleNode, "lnRef");
    if (lineRef) {
      const index = parseInt(lineRef.getAttribute("idx") || "0", 10);
      const placeholderColor = this.extractHexColor(lineRef);
      const source = this.themeStyleMatrix.lines[index - 1];
      const line = source
        ? this.parseLineStyle(source, absoluteUnitScale, placeholderColor)
        : undefined;
      if (line) {
        const id = this.registerStyleRule({
          source: { kind: "theme", part: "ppt/theme/theme1.xml", nodeId: `lnRef:${index}` },
          parents: [],
          declarations: { line }
        });
        result.push({ id, property: "line", value: line });
      }
    }

    const effectRef = getDirectChild(styleNode, "effectRef");
    if (effectRef) {
      const index = parseInt(effectRef.getAttribute("idx") || "0", 10);
      const placeholderColor = this.extractHexColor(effectRef);
      const source = this.themeStyleMatrix.effects[index - 1];
      const effects = source
        ? this.parseEffectStyle(source, absoluteUnitScale, placeholderColor)
        : undefined;
      if (effects) {
        const id = this.registerStyleRule({
          source: { kind: "theme", part: "ppt/theme/theme1.xml", nodeId: `effectRef:${index}` },
          parents: [],
          declarations: { effects }
        });
        result.push({ id, property: "effects", value: effects });
      }
    }
    return result;
  }

  private async parseShapeNode(
    node: Element,
    transform: TransformState,
    elements: SlideElement[],
    imgRelMap: Record<string, string>,
    isTemplate: boolean = false,
    placeholderContext: PlaceholderContext | null = null
  ) {
    const nvSpPr = getDirectChild(node, "nvSpPr", "nvCxnSpPr");
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

    // Geometry is structural: only a direct child of spPr can define it.
    const prstGeom = spPrNodes
      .map(candidate => getDirectChild(candidate, "prstGeom"))
      .find((candidate): candidate is globalThis.Element => candidate !== null) || null;
    const customGeometry = getDirectChild(spPr, "custGeom");
    const parsedCustomGeometry = customGeometry ? parseCustomGeometry(customGeometry) : undefined;
    let shapeType = prstGeom?.getAttribute("prst") || (customGeometry ? "custom" : "rect");
    const xfrm = spPrNodes
      .map(candidate => getDirectChild(candidate, "xfrm"))
      .find((candidate): candidate is globalThis.Element => candidate !== null) || null;
    const rotation = xfrm ? parseInt(xfrm.getAttribute("rot") || "0", 10) / 60000 : 0;
    const flipH = xfrm?.getAttribute("flipH") === "1" || xfrm?.getAttribute("flipH") === "true";
    const flipV = xfrm?.getAttribute("flipV") === "1" || xfrm?.getAttribute("flipV") === "true";
    const adjustments: Record<string, number> = {};
    if (prstGeom) {
      const avLst = getDirectChild(prstGeom, "avLst");
      for (const adjustment of avLst ? getDirectChildren(avLst, "gd") : []) {
        const name = adjustment.getAttribute("name");
        const value = adjustment.getAttribute("fmla")?.match(/(?:val|pin)\s+(-?\d+(?:\.\d+)?)/)?.[1];
        if (name && value !== undefined) adjustments[name] = parseFloat(value) / 100000;
      }
    }
    const cornerRadius = shapeType === "roundRect" && adjustments.adj !== undefined
      ? Math.max(0, Math.min(0.5, adjustments.adj))
      : undefined;
    let arrowHeadHeight: number | undefined;
    let arrowShaftWidth: number | undefined;
    if (shapeType === "upArrow") {
      if (adjustments.adj1 !== undefined) {
        arrowHeadHeight = Math.max(0.001, Math.min(0.5, adjustments.adj1));
      }
      if (adjustments.adj2 !== undefined) {
        arrowShaftWidth = Math.max(0.001, Math.min(1, adjustments.adj2));
      }
    }

    const computedStyle: ComputedShapeStyle = { fill: { type: "none" } };
    const styleRefs: string[] = [];
    const styleTrace: Partial<Record<ShapeStyleProperty, string>> = {};
    let previousSourceRule: string | null = null;

    // Apply low-priority sources first: master -> layout -> slide shape.
    const cascade = [...shapeChain].reverse();
    for (let sourceIndex = 0; sourceIndex < cascade.length; sourceIndex++) {
      const sourceShape = cascade[sourceIndex];
      const sourceSpPr = getDirectChild(sourceShape, "spPr");
      const sourceStyle = getDirectChild(sourceShape, "style");
      const sourceKind: StyleSourceKind = sourceShape === node
        ? (isTemplate ? "placeholder" : "shape")
        : placeholderContext?.layoutShapes.includes(sourceShape)
          ? "layout"
          : placeholderContext?.masterShapes.includes(sourceShape)
            ? "master"
            : "placeholder";
      const sourcePart = sourceKind === "shape"
        ? placeholderContext?.slidePart || "slide"
        : sourceKind === "layout"
          ? placeholderContext?.layoutPart || "layout"
          : sourceKind === "master"
            ? placeholderContext?.masterPart || "master"
            : sourceKind;
      const sourceNodeId = querySelector(sourceShape, "p\\:cNvPr, cNvPr")?.getAttribute("id") || undefined;
      const themeRules = this.styleResolver.resolveThemeStyleRules(sourceStyle, transform.absoluteUnitScale);
      const declarations: Partial<ComputedShapeStyle> = {};

      for (const themeRule of themeRules) {
        styleRefs.push(themeRule.id);
        if (themeRule.property === "fill") computedStyle.fill = themeRule.value as FillStyle;
        else if (themeRule.property === "line") computedStyle.line = themeRule.value as LineStyle;
        else computedStyle.effects = themeRule.value as EffectStyle;
        styleTrace[themeRule.property] = themeRule.id;
      }

      if (sourceSpPr) {
        let directFill = this.styleResolver.parseFillStyle(sourceSpPr);
        if (directFill?.type === "picture" && directFill.embed && imgRelMap[directFill.embed]) {
          const imgZipFile = this.zip?.file(imgRelMap[directFill.embed]);
          if (imgZipFile) {
            const blob = await imgZipFile.async("blob");
            directFill = { ...directFill, url: URL.createObjectURL(blob) };
          }
        }
        const directLineNode = getDirectChild(sourceSpPr, "ln");
        const directEffectNode = getDirectChild(sourceSpPr, "effectLst", "effectDag");
        const directEffects = this.styleResolver.parseEffectStyle(sourceSpPr, transform.absoluteUnitScale);
        if (directFill) declarations.fill = directFill;
        if (directLineNode) {
          const directLine = this.styleResolver.parseLineStyle(directLineNode, transform.absoluteUnitScale);
          // An explicit a:ln/a:noFill clears an inherited line.
          declarations.line = directLine || { fill: { type: "none" }, width: 0 };
        }
        if (directEffectNode) declarations.effects = directEffects || {};
      }

      const parents = [
        ...(previousSourceRule ? [previousSourceRule] : []),
        ...themeRules.map(rule => rule.id)
      ];
      if (Object.keys(declarations).length > 0 || themeRules.length > 0) {
        const sourceRule = this.registerStyleRule({
          source: {
            kind: sourceKind,
            part: sourcePart,
            nodeId: sourceNodeId
          },
          parents,
          declarations
        });
        styleRefs.push(sourceRule);
        previousSourceRule = sourceRule;
        for (const property of ["fill", "line", "effects"] as ShapeStyleProperty[]) {
          const value = declarations[property];
          if (value === undefined) continue;
          if (property === "fill") computedStyle.fill = value as FillStyle;
          else if (property === "line") computedStyle.line = value as LineStyle;
          else computedStyle.effects = value as EffectStyle;
          styleTrace[property] = sourceRule;
        }
      }
    }

    const fallbackFill = computedStyle.fill.type === "solid"
      ? computedStyle.fill.color
      : computedStyle.fill.type === "gradient"
        ? computedStyle.fill.stops[0]?.color || "#cccccc"
        : "transparent";
    const lineFill = computedStyle.line?.fill;
    const borderColor = lineFill?.type === "solid" ? lineFill.color : null;
    const border: Border | undefined = borderColor && computedStyle.line
      ? { color: borderColor, width: computedStyle.line.width }
      : undefined;
    const hasFill = computedStyle.fill.type !== "none";
    const hasBorder = !!computedStyle.line && computedStyle.line.fill.type !== "none";
    const hasEffects = !!(
      computedStyle.effects?.outerShadow
      || computedStyle.effects?.innerShadow
      || computedStyle.effects?.glow
      || computedStyle.effects?.reflection
      || computedStyle.effects?.softEdge
      || computedStyle.effects?.blur
      || computedStyle.effects?.fillOverlay
    );

    // Slide shapes always get a shape element, even a see-through one, so the
    // editor has something to give a fill or border to later.
    const editableShape = !isTemplate && !!cNvPr?.getAttribute("id");
    if (hasFill || hasBorder || hasEffects || editableShape) {
      elements.push({
        type: "shape",
        id: `shape_bg_${id}`,
        source: this.makeSource(isTemplate, cNvPr?.getAttribute("id"), transform, cNvPr?.getAttribute("name")),
        rect: { ...rect },
        shapeType,
        rotation,
        flipH,
        flipV,
        cornerRadius,
        arrowHeadHeight,
        arrowShaftWidth,
        adjustments,
        customGeometry: parsedCustomGeometry,
        fill: fallbackFill,
        border,
        styleRefs,
        computedStyle,
        styleTrace
      });
    }

    // Default text color and font size from shape placeholder / style
    let defaultTextColor = "#333333";
    let defaultFontSize = 18; // Default in points (18pt)
    let defaultBold = false;
    let defaultFontFamily = "sans-serif";
    let defaultEastAsianFontFamily = "";

    // Check style node for fontRef
    const styleNode = querySelector(node, "p\\:style, style");
    if (styleNode) {
      const fontRef = querySelector(styleNode, "a\\:fontRef, fontRef");
      if (fontRef) {
        const color = this.styleResolver.extractHexColor(fontRef);
        if (color) defaultTextColor = color;
        const fontScheme = fontRef.getAttribute("idx");
        if (fontScheme === "major" || fontScheme === "minor") {
          const prefix = fontScheme === "major" ? "+mj" : "+mn";
          defaultFontFamily = this.styleResolver.resolveThemeTypeface(`${prefix}-lt`);
          defaultEastAsianFontFamily = this.themeFonts[`${prefix}-Hans`]
            || this.styleResolver.resolveThemeTypeface(`${prefix}-ea`);
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
      const placeholder = this.getPlaceholderDescriptor(node);
      const inheritance: TextInheritanceContext = {
        textBodies,
        // Master txStyles participate in placeholder inheritance. Ordinary shapes
        // carry their defaults in their own txBody/lstStyle and must not inherit
        // unrelated master bullet or run defaults.
        masterTextStyle: placeholder
          ? this.selectMasterTextStyle(placeholderContext?.masterTextStyles || null, placeholder)
          : null,
        defaultTextStyle: placeholder ? this.defaultTextStyle : null
      };
      const body = this.resolveTextBodyProperties(
        textBodies,
        transform.absoluteUnitScale,
        transform.absoluteUnitScale
      );
      const firstTextElement = elements.length;
      await this.parseTextBody(
        txBody,
        rect.x,
        rect.y,
        rect.w,
        rect.h,
        transform.absoluteUnitScale,
        id,
        elements,
        defaultTextColor,
        defaultFontSizePx,
        defaultBold,
        defaultFontFamily,
        defaultEastAsianFontFamily,
        inheritance,
        body,
        cNvPr?.getAttribute("id") || undefined,
        placeholderContext?.slidePart,
        this.makeSource(isTemplate, cNvPr?.getAttribute("id"), transform, cNvPr?.getAttribute("name"))
      );
      for (const element of elements.slice(firstTextElement)) {
        if (element.type !== "text") continue;
        if (rotation) element.rotation = rotation;
        if (ph && !isTemplate) element.placeholder = { type: ph.getAttribute("type") || "body" };
      }
    }
  }

  /** A graphicFrame that holds a chart becomes one chart element; returns false for any other frame. */
  private async parseChartFrame(
    node: Element,
    transform: TransformState,
    elements: SlideElement[],
    isTemplate: boolean
  ): Promise<boolean> {
    const graphicData = querySelector(node, "a\\:graphicData, graphicData");
    if (!graphicData || !(graphicData.getAttribute("uri") || "").endsWith("/chart")) return false;
    const chartRef = getDirectChild(graphicData, "chart");
    const relId = chartRef?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
      || chartRef?.getAttribute("r:id");
    const xfrm = getDirectChild(node, "xfrm");
    const off = xfrm && getDirectChild(xfrm, "off");
    const ext = xfrm && getDirectChild(xfrm, "ext");
    if (!relId || !off || !ext) return true;

    const frame = querySelector(node, "p\\:cNvPr, cNvPr");
    const frameId = frame?.getAttribute("id") || `${elements.length}`;
    const rect = {
      x: transform.offsetX + parseInt(off.getAttribute("x") || "0", 10) * transform.scaleX,
      y: transform.offsetY + parseInt(off.getAttribute("y") || "0", 10) * transform.scaleY,
      w: parseInt(ext.getAttribute("cx") || "0", 10) * transform.scaleX,
      h: parseInt(ext.getAttribute("cy") || "0", 10) * transform.scaleY
    };

    const part = this.currentChartRels[relId];
    const file = part ? this.zip?.file(part) : null;
    const accents = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map(name => this.themeColors[name]);
    const context = {
      solidFill: (target: Element) => this.styleResolver.extractDirectFillColor(target),
      typeface: (name: string) => this.styleResolver.resolveThemeTypeface(name),
      palette: accents,
      defaultFont: {
        family: this.themeFonts["+mn-lt"] || undefined,
        eastAsianFamily: this.themeFonts["+mn-ea"] || undefined
      }
    };
    let chart;
    try {
      chart = parseChartXml(file ? await file.async("string") : "", part || "", context);
    } catch (error) {
      chart = parseChartXml("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>", part || "", context);
    }
    elements.push({
      type: "chart",
      id: `chart_${frameId}`,
      source: this.makeSource(isTemplate, frameId, transform, frame?.getAttribute("name")),
      rect,
      chart,
      pxPerPt: 12700 * transform.scaleX
    });
    return true;
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

    if (await this.parseChartFrame(node, transform, elements, isTemplate)) return;

    const tbl = querySelector(node, "a\\:tbl, tbl");
    if (!tbl) return;
    const frameId = querySelector(node, "p\\:cNvPr, cNvPr")?.getAttribute("id") || `${elements.length}`;
    const source = this.makeSource(isTemplate, frameId, transform, querySelector(node, "p\\:cNvPr, cNvPr")?.getAttribute("name"));
    const element = await this.parseTableElement(node, tbl, transform, frameId, source);
    if (element) elements.push(element);
  }

  /** The deck's table styles (ppt/tableStyles.xml), as plain data. */
  private async parseTableStyles(): Promise<void> {
    this.tableStyles = new Map();
    const file = this.zip?.file("ppt/tableStyles.xml");
    if (!file) return;
    const doc = new DOMParser().parseFromString(await file.async("text"), "application/xml");
    const colorOf = (holder: Element): string | null => {
      // A colour written straight under the node (not inside a nested font reference).
      const direct = Array.from(holder.childNodes).find((child): child is Element =>
        child.nodeType === 1 && /^(schemeClr|srgbClr|sysClr|prstClr|scrgbClr)$/.test((child as Element).localName || ""));
      if (!direct) return null;
      const wrapper = doc.createElementNS("http://schemas.openxmlformats.org/drawingml/2006/main", "a:solidFill");
      wrapper.appendChild(direct.cloneNode(true));
      return this.styleResolver.extractHexColor(wrapper);
    };
    const names: TableStylePartName[] = [
      "wholeTbl", "band1H", "band2H", "band1V", "band2V", "firstCol", "lastCol",
      "firstRow", "lastRow", "nwCell", "neCell", "swCell", "seCell"
    ];
    for (const style of Array.from(doc.getElementsByTagNameNS("*", "tblStyle"))) {
      const id = style.getAttribute("styleId");
      if (!id) continue;
      const def: TableStyleDef = {};
      for (const name of names) {
        const node = getDirectChild(style, name);
        if (!node) continue;
        const part: TableStylePart = {};
        const text = getDirectChild(node, "tcTxStyle");
        if (text) {
          const color = colorOf(text);
          if (color) part.textColor = color;
          if (text.getAttribute("b")) part.bold = text.getAttribute("b") === "on";
          if (text.getAttribute("i")) part.italic = text.getAttribute("i") === "on";
        }
        const cell = getDirectChild(node, "tcStyle");
        if (cell) {
          const fill = getDirectChild(cell, "fill");
          if (fill) {
            const solid = getDirectChild(fill, "solidFill");
            if (solid) part.fill = this.styleResolver.extractHexColor(solid);
            else if (getDirectChild(fill, "noFill")) part.fill = null;
          }
          const borders = getDirectChild(cell, "tcBdr");
          if (borders) {
            part.borders = {};
            for (const side of ["left", "right", "top", "bottom", "insideH", "insideV"] as const) {
              const line = getDirectChild(borders, side);
              const ln = line && getDirectChild(line, "ln");
              if (!ln) continue;
              const solid = getDirectChild(ln, "solidFill");
              const color = solid ? this.styleResolver.extractHexColor(solid) : null;
              part.borders[side] = color
                ? { color, width: parseInt(ln.getAttribute("w") || "12700", 10) }
                : null;
            }
          }
        }
        def[name] = part;
      }
      this.tableStyles.set(id, def);
    }
  }

  private tableStyleFor(id: string | undefined): TableStyleDef | undefined {
    if (!id) return undefined;
    const known = this.tableStyles.get(id);
    if (known) return known;
    if (id !== DEFAULT_TABLE_STYLE_ID) return undefined;
    const tint = (hex: string, factor: number) => {
      const channel = (offset: number) => Math.round(parseInt(hex.slice(offset, offset + 2), 16) * factor + 255 * (1 - factor));
      return `#${[1, 3, 5].map(offset => channel(offset).toString(16).padStart(2, "0")).join("")}`;
    };
    return builtinDefaultStyle(this.themeColors.accent1, this.themeColors.dk1, this.themeColors.lt1, tint);
  }

  /**
   * A table frame as one element: its grid, merged cells, and for each cell the
   * fill, borders and text as the table style and the cell itself give them.
   */
  private async parseTableElement(
    frame: Element,
    tbl: Element,
    transform: TransformState,
    frameId: string,
    source: ElementSource | undefined
  ): Promise<TableElement | null> {
    const xfrm = getDirectChild(frame, "xfrm");
    const off = xfrm && getDirectChild(xfrm, "off");
    const ext = xfrm && getDirectChild(xfrm, "ext");
    if (!off || !ext) return null;
    const rect = {
      x: transform.offsetX + parseInt(off.getAttribute("x") || "0", 10) * transform.scaleX,
      y: transform.offsetY + parseInt(off.getAttribute("y") || "0", 10) * transform.scaleY,
      w: parseInt(ext.getAttribute("cx") || "0", 10) * transform.scaleX,
      h: parseInt(ext.getAttribute("cy") || "0", 10) * transform.scaleY
    };
    const scale = transform.absoluteUnitScale;

    const properties = getDirectChild(tbl, "tblPr");
    const on = (name: string) => ["1", "true"].includes(properties?.getAttribute(name) || "");
    const flags = {
      firstRow: on("firstRow"), lastRow: on("lastRow"), firstColumn: on("firstCol"),
      lastColumn: on("lastCol"), bandRow: on("bandRow"), bandColumn: on("bandCol")
    };
    const style = this.tableStyleFor(getDirectChild(properties || tbl, "tableStyleId")?.textContent?.trim());

    const grid = getDirectChild(tbl, "tblGrid");
    const columns = (grid ? getDirectChildren(grid, "gridCol") : []).map(col => parseInt(col.getAttribute("w") || "0", 10) * transform.scaleX);
    const rowNodes = getDirectChildren(tbl, "tr");
    const family = this.styleResolver.resolveThemeTypeface("+mn-lt");
    const eastAsian = this.themeFonts["+mn-Hans"] || this.styleResolver.resolveThemeTypeface("+mn-ea");
    const toBorder = (border: StyleBorder | null | undefined) =>
      border ? { color: border.color, width: Math.max(1, border.width * scale) } : null;

    const rows: TableModel["rows"] = [];
    let merged = false;
    for (let r = 0; r < rowNodes.length; r++) {
      const cells: TableCell[] = [];
      const cellNodes = getDirectChildren(rowNodes[r], "tc");
      for (let c = 0; c < columns.length; c++) {
        const node = cellNodes[c];
        const colSpan = parseInt(node?.getAttribute("gridSpan") || "1", 10) || 1;
        const rowSpan = parseInt(node?.getAttribute("rowSpan") || "1", 10) || 1;
        const covered = ["1", "true"].includes(node?.getAttribute("hMerge") || "") || ["1", "true"].includes(node?.getAttribute("vMerge") || "");
        if (colSpan > 1 || rowSpan > 1 || covered) merged = true;

        const resolved = resolveCellStyle(style, flags, { row: r, col: c, rowSpan, colSpan }, rowNodes.length, columns.length);
        const tcPr = node && getDirectChild(node, "tcPr");
        let fill = resolved.fill;
        const borders = {
          left: toBorder(resolved.borders.left), right: toBorder(resolved.borders.right),
          top: toBorder(resolved.borders.top), bottom: toBorder(resolved.borders.bottom)
        };
        if (tcPr) {
          if (getDirectChild(tcPr, "noFill")) fill = null;
          else if (getDirectChild(tcPr, "solidFill")) fill = this.styleResolver.extractDirectFillColorWithOpacity(tcPr) ?? fill;
          for (const [side, tag] of [["left", "lnL"], ["right", "lnR"], ["top", "lnT"], ["bottom", "lnB"]] as const) {
            const line = getDirectChild(tcPr, tag);
            if (!line) continue;
            if (getDirectChild(line, "noFill")) {
              borders[side] = null;
              continue;
            }
            const color = this.styleResolver.extractDirectFillColorWithOpacity(line);
            if (color) borders[side] = { color, width: Math.max(1, parseInt(line.getAttribute("w") || "12700", 10) * scale) };
          }
        }

        const margin = (name: string, fallback: number) => parseInt(tcPr?.getAttribute(name) || String(fallback), 10) * scale;
        const anchor = tcPr?.getAttribute("anchor");
        const body = {
          marginLeft: margin("marL", 91440), marginRight: margin("marR", 91440),
          marginTop: margin("marT", 45720), marginBottom: margin("marB", 45720),
          verticalAnchor: anchor === "ctr" ? "middle" as const : anchor === "b" ? "bottom" as const : "top" as const,
          autoFit: "none" as const, wrap: "square" as const, verticalOverflow: "overflow" as const,
          horizontalOverflow: "overflow" as const, fontScale: 1, textDirection: "horz" as const
        };
        const txBody = node && getDirectChild(node, "txBody");
        const scratch: SlideElement[] = [];
        if (txBody && !covered) {
          await this.parseTextBody(
            txBody, 0, 0, columns[c] || 1, 1, scale, `cell_${r}_${c}`, scratch,
            resolved.textColor ?? this.themeColors.dk1,
            Math.round(18 * 12700 * scale),
            resolved.bold ?? false,
            family, eastAsian, null, body
          );
        }
        const text = scratch.find((element): element is TextElement => element.type === "text");
        cells.push({
          colSpan, rowSpan, covered, fill, borders,
          paragraphs: text?.paragraphs ?? [],
          style: text?.style ?? {
            fontSize: Math.round(18 * 12700 * scale), color: resolved.textColor ?? this.themeColors.dk1,
            bold: resolved.bold ?? false, align: "left", fontFamily: family, eastAsianFontFamily: eastAsian
          },
          body: text?.body ?? body,
          text: text?.content ?? ""
        });
      }
      rows.push({ height: parseInt(rowNodes[r].getAttribute("h") || "0", 10) * transform.scaleY, cells });
    }
    if (columns.length === 0 || rows.length === 0) return null;

    return {
      type: "table",
      id: `table_${frameId}`,
      source,
      rect,
      table: { columns, rows, firstRow: flags.firstRow, bandRow: flags.bandRow, firstColumn: flags.firstColumn, merged },
      pxPerPt: 12700 * scale
    };
  }

  /**
   * Read a table frame from its XML again (after an edit changed it), as it sits
   * on the slide that `source` names.
   */
  public async parseTableXml(xml: string, source: ElementSource): Promise<TableElement | null> {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const frame = doc.documentElement;
    const tbl = Array.from(frame.getElementsByTagNameNS("*", "tbl"))[0];
    if (!tbl) return null;
    const transform: TransformState = { ...source.xform, absoluteUnitScale: source.xform.scaleX };
    const element = await this.parseTableElement(frame, tbl, transform, source.shapeId, source);
    return element;
  }

  /**
   * Symbol fonts (Wingdings, Webdings, Symbol) store bullets as private-use
   * ASCII codes. Those fonts are not available in the renderer, so the code
   * is translated to the equivalent Unicode glyph and drawn with the normal
   * font fallback chain instead of showing the raw letter (e.g. "n").
   */
  private static readonly SYMBOL_BULLET_MAP: Record<string, Record<string, string>> = {
    wingdings: {
      l: "●", m: "○", n: "■", o: "□", p: "◻", q: "❑", r: "❒", s: "⬧", t: "⧫",
      u: "◆", v: "❖", w: "⬥", x: "⌧", y: "⍓", z: "⌘",
      "§": "▪", "¨": "◻", "©": "◈", "¡": "○", "¢": "●",
      "Ø": "➢", "Ü": "➔", "è": "➔", "ü": "✓", "û": "✗", "þ": "☑", "ý": "☒",
      "J": "☺", "L": "☹", "*": "✎", "&": "☎", "4": "✌", "!": "✏"
    },
    "wingdings 2": { l: "●", n: "■", p: "◻", "¢": "●", "£": "◻", "R": "☑", "Q": "☐" },
    "wingdings 3": { p: "▲", q: "▼", t: "►", u: "◄", "}": "▶", "~": "◀" },
    symbol: { "·": "•", "§": "♣", "¨": "♦", "ª": "♠", "Ø": "⇒" },
    webdings: { "=": "•", "4": "▶", "a": "✓" }
  };

  private getBulletDescriptor(
    charValue: string,
    typeface: string
  ): { char: string; fontFamily?: string } {
    const mapped = PptxParser.SYMBOL_BULLET_MAP[typeface.trim().toLowerCase()]?.[charValue];
    if (mapped) return { char: mapped };
    return {
      char: charValue,
      fontFamily: typeface ? this.styleResolver.resolveThemeTypeface(typeface) : undefined
    };
  }

  private async parseTextBody(
    txBody: Element,
    x: number,
    y: number,
    w: number,
    h: number,
    absoluteUnitScale: number,
    id: string,
    elements: SlideElement[],
    defaultTextColor: string = "#333333",
    defaultFontSizePx: number = 20,
    defaultBold: boolean = false,
    defaultFontFamily: string = "sans-serif",
    defaultEastAsianFontFamily: string = "",
    inheritance: TextInheritanceContext | null = null,
    resolvedBody: TextBodyProperties | null = null,
    sourceShapeId?: string,
    sourceSlidePart?: string,
    source?: ElementSource
  ) {
    const paragraphs = querySelectorAll(txBody, "a\\:p, p");
    const body = resolvedBody || {
      marginLeft: 8,
      marginRight: 8,
      marginTop: 4,
      marginBottom: 4,
      verticalAnchor: "top" as const,
      autoFit: "none" as const,
      wrap: "square" as const,
      verticalOverflow: "overflow" as const,
      horizontalOverflow: "overflow" as const,
      fontScale: 1,
      textDirection: "horz" as const
    };
    const computedParagraphs: TextParagraph[] = [];
    /** Running count of numbered paragraphs, per level. */
    const numbering: Array<{ type: string; count: number } | undefined> = [];

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const pPr = querySelector(p, "a\\:pPr, pPr");
      const level = Math.max(0, Math.min(8, parseInt(pPr?.getAttribute("lvl") || "0", 10)));
      let pDefaults: ComputedRunStyle = {
        fontSize: defaultFontSizePx,
        color: defaultTextColor,
        bold: defaultBold,
        fontFamily: defaultFontFamily,
        // A theme that names no East Asian font leaves it to the application: PowerPoint in a Chinese locale
        // uses Microsoft YaHei, whose tall line height is what Chinese text is laid out with.
        eastAsianFontFamily: defaultEastAsianFontFamily || DEFAULT_EAST_ASIAN_FONT,
        italic: false,
        letterSpacing: 0,
        shadow: undefined
      };

      const inheritedParagraphProperties: globalThis.Element[] = [];
      if (inheritance?.defaultTextStyle) {
        const base = this.getLevelParagraphProperties(inheritance.defaultTextStyle, level);
        if (base) inheritedParagraphProperties.push(base);
      }
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
      let eastAsianLineBreak = false;
      let hangingPunctuation = false;
      let fontAlignment: ParagraphStyle["fontAlignment"] = "auto";
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
        // Paragraph margins and hanging indents are EMU values in the
        // presentation coordinate system. They are not coordinates inside a
        // group, so groupScale must never be applied here.
        if (marL !== null) marginLeft = parseInt(marL, 10) * absoluteUnitScale;
        if (indentAttr !== null) indent = parseInt(indentAttr, 10) * absoluteUnitScale;

        const eaLnBrk = inheritedPPr.getAttribute("eaLnBrk");
        if (eaLnBrk !== null) eastAsianLineBreak = eaLnBrk === "1" || eaLnBrk === "true";
        const hangingPunct = inheritedPPr.getAttribute("hangingPunct");
        if (hangingPunct !== null) {
          hangingPunctuation = hangingPunct === "1" || hangingPunct === "true";
        }
        const fontAlgn = inheritedPPr.getAttribute("fontAlgn");
        if (fontAlgn === "base") fontAlignment = "baseline";
        else if (fontAlgn === "t") fontAlignment = "top";
        else if (fontAlgn === "ctr") fontAlignment = "center";
        else if (fontAlgn === "b") fontAlignment = "bottom";
        else if (fontAlgn === "auto") fontAlignment = "auto";

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

      // Bullet properties inherit one by one: a layout may change only the size
      // or colour of a bullet whose character comes from the master.
      let bulletKind: "none" | "char" | "auto" | null = null;
      let bulletCharValue = "";
      let bulletTypeface = "";
      let bulletAuto: { type: string; startAt: number } | null = null;
      let bulletSizePct: number | null = null;
      let bulletSizePts: number | null = null;
      let bulletColorValue: string | null = null;

      for (const inheritedPPr of inheritedParagraphProperties) {
        if (getDirectChild(inheritedPPr, "buNone")) bulletKind = "none";
        const buChar = getDirectChild(inheritedPPr, "buChar");
        if (buChar) {
          bulletKind = "char";
          bulletCharValue = buChar.getAttribute("char") || "•";
        }
        const buAutoNum = getDirectChild(inheritedPPr, "buAutoNum");
        if (buAutoNum) {
          bulletKind = "auto";
          bulletAuto = {
            type: buAutoNum.getAttribute("type") || "arabicPeriod",
            startAt: parseInt(buAutoNum.getAttribute("startAt") || "1", 10)
          };
        }
        const buFont = getDirectChild(inheritedPPr, "buFont");
        if (buFont) bulletTypeface = buFont.getAttribute("typeface") || "";
        if (getDirectChild(inheritedPPr, "buFontTx")) bulletTypeface = "";

        const buSzPct = getDirectChild(inheritedPPr, "buSzPct")?.getAttribute("val");
        const buSzPts = getDirectChild(inheritedPPr, "buSzPts")?.getAttribute("val");
        if (buSzPct !== null && buSzPct !== undefined) {
          bulletSizePct = Math.max(0, parseInt(buSzPct, 10) / 100000);
          bulletSizePts = null;
        } else if (buSzPts !== null && buSzPts !== undefined) {
          bulletSizePts = (parseInt(buSzPts, 10) / 100) * 12700 * absoluteUnitScale;
          bulletSizePct = null;
        }
        if (getDirectChild(inheritedPPr, "buSzTx")) {
          bulletSizePct = null;
          bulletSizePts = null;
        }

        const buClr = getDirectChild(inheritedPPr, "buClr");
        if (buClr) bulletColorValue = this.styleResolver.extractHexColor(buClr) || bulletColorValue;
        if (getDirectChild(inheritedPPr, "buClrTx")) bulletColorValue = null;
      }

      let bulletChar = "";
      let bulletColor = bulletColorValue || pDefaults.color;
      let bulletFontFamily: string | undefined;
      let bulletFontSize: number | undefined;
      const hasExplicitBulletColor = bulletColorValue !== null;
      if (bulletKind === "char") {
        const descriptor = this.getBulletDescriptor(bulletCharValue, bulletTypeface);
        bulletChar = descriptor.char;
        bulletFontFamily = descriptor.fontFamily;
      } else if (bulletKind === "auto" && bulletAuto) {
        // Consecutive numbered paragraphs count up; any other paragraph at the
        // same level ends the run, and a shallower one restarts deeper levels.
        for (let deeper = level + 1; deeper < numbering.length; deeper++) numbering[deeper] = undefined;
        const previous = numbering[level];
        const count = previous && previous.type === bulletAuto.type ? previous.count + 1 : bulletAuto.startAt;
        numbering[level] = { type: bulletAuto.type, count };
        bulletChar = formatAutoNumber(bulletAuto.type, count);
        if (bulletTypeface) bulletFontFamily = this.getBulletDescriptor("", bulletTypeface).fontFamily;
      }
      if (bulletKind !== "auto") {
        for (let deeper = level + 1; deeper < numbering.length; deeper++) numbering[deeper] = undefined;
        numbering[level] = undefined;
      }

      // Title paragraphs in this deck inherit the level-1 hanging indent
      // from lstStyle, but explicitly use text-only bullet properties
      // (buClrTx/buSzTx/buFontTx) and have no bullet character. In that
      // combination WPS resets the inherited hanging indent; keeping it
      // would move the first title glyph outside the text box and clip it.
      const explicitTextOnlyBulletProperties = pPr &&
        !getDirectChild(pPr, "buChar") &&
        (getDirectChild(pPr, "buNone") ||
          getDirectChild(pPr, "buClrTx") ||
          getDirectChild(pPr, "buSzTx") ||
          getDirectChild(pPr, "buFontTx"));
      if (!bulletChar && explicitTextOnlyBulletProperties) {
        indent = 0;
      }

      const computedRuns: TextRun[] = [];
      const runs = querySelectorAll(p, "a\\:r, r, a\\:fld, fld, a\\:br, br");
      for (const run of runs) {
        const runName = run.localName || run.nodeName;
        let runStyle = pDefaults;
        const rPr = querySelector(run, "a\\:rPr, rPr");
        if (rPr) runStyle = this.parseRunProperties(rPr, absoluteUnitScale, pDefaults);
        const directRunColor = rPr ? this.styleResolver.extractDirectFillColorWithOpacity(rPr) : null;
        if (
          bulletChar
          && !directRunColor
          && (runStyle.color.toLowerCase() === "#ffffff" || runStyle.color.toLowerCase() === "#fff")
        ) {
          // A colored bullet can inherit lt1 from a template while its body
          // text remains the ordinary dark body style.
          runStyle = { ...runStyle, color: "#333333" };
        }

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
              italic: runStyle.italic,
              letterSpacing: runStyle.letterSpacing,
              reflection: runStyle.reflection,
              shadow: runStyle.shadow,
              outline: runStyle.outline
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
            italic: emptyStyle.italic,
            letterSpacing: emptyStyle.letterSpacing,
            reflection: emptyStyle.reflection,
            shadow: emptyStyle.shadow,
            outline: emptyStyle.outline
          }
        });
      }

      if (bulletChar) {
        // A percentage is of the paragraph's first run, which may differ from the level default.
        const firstRunSize = computedRuns[0].style.fontSize;
        if (bulletSizePct !== null) bulletFontSize = firstRunSize * bulletSizePct;
        else if (bulletSizePts !== null) bulletFontSize = bulletSizePts;
      }
      if (bulletChar && !hasExplicitBulletColor) {
        bulletColor = computedRuns[0].style.color;
      }

      computedParagraphs.push({
        style: {
          align,
          level,
          marginLeft,
          indent,
          eastAsianLineBreak,
          hangingPunctuation,
          fontAlignment,
          lineSpacing,
          spaceBefore,
          spaceAfter
        },
        bullet: bulletChar
          ? {
              char: bulletChar,
              color: bulletColor,
              fontFamily: bulletFontFamily,
              fontSize: bulletFontSize
            }
          : undefined,
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
      sourceShapeId,
      sourceSlidePart,
      source,
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
    const xfrm = getDirectChild(spPr, "xfrm");
    const rotation = xfrm ? parseInt(xfrm.getAttribute("rot") || "0", 10) / 60000 : 0;
    const flipH = xfrm?.getAttribute("flipH") === "1" || xfrm?.getAttribute("flipH") === "true";
    const flipV = xfrm?.getAttribute("flipV") === "1" || xfrm?.getAttribute("flipV") === "true";

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
    const effects = this.styleResolver.parseEffectStyle(spPr, transform.absoluteUnitScale);

    elements.push({
      type: "image",
      id: `img_${id}`,
      source: this.makeSource(isTemplate, cNvPr?.getAttribute("id"), transform, cNvPr?.getAttribute("name")),
      rect,
      url: blobUrl,
      crop,
      rotation,
      flipH,
      flipV,
      effects
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


const ROMAN: Array<[number, string]> = [
  [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
  [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]
];

function toRoman(value: number): string {
  let rest = Math.max(1, Math.min(3999, value));
  let result = "";
  for (const [size, symbol] of ROMAN) {
    while (rest >= size) {
      result += symbol;
      rest -= size;
    }
  }
  return result;
}

function toLetters(value: number): string {
  let rest = Math.max(1, value);
  let result = "";
  while (rest > 0) {
    rest -= 1;
    result = String.fromCharCode(97 + (rest % 26)) + result;
    rest = Math.floor(rest / 26);
  }
  return result;
}

/** The text of an auto-number bullet, e.g. ("alphaLcParenR", 2) -> "b)". */
export function formatAutoNumber(type: string, count: number): string {
  const lower = type.toLowerCase();
  let body: string;
  if (lower.startsWith("alphalc")) body = toLetters(count);
  else if (lower.startsWith("alphauc")) body = toLetters(count).toUpperCase();
  else if (lower.startsWith("romanlc")) body = toRoman(count).toLowerCase();
  else if (lower.startsWith("romanuc")) body = toRoman(count);
  else if (lower.startsWith("circlenum")) body = String.fromCodePoint(count >= 1 && count <= 20 ? 0x2460 + count - 1 : 0x2460);
  else body = String(count);

  if (lower.endsWith("parenboth")) return `(${body})`;
  if (lower.endsWith("parenr")) return `${body})`;
  if (lower.endsWith("period")) return `${body}.`;
  if (lower.endsWith("plain")) return body;
  return `${body}.`;
}
