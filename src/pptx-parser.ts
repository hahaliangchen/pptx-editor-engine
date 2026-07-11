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
}

export interface TextElement {
  type: "text";
  id: string;
  rect: Rect;
  content: string;
  style: TextStyle;
}

export interface ShapeElement {
  type: "shape";
  id: string;
  rect: Rect;
  shapeType: "rect" | "ellipse" | "triangle";
  fill: string;
}

export interface ImageElement {
  type: "image";
  id: string;
  rect: Rect;
  url: string;
}

export type Element = TextElement | ShapeElement | ImageElement;

export interface Slide {
  id: string;
  elements: Element[];
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

  public async parse(buffer: ArrayBuffer): Promise<PresentationAST> {
    this.zip = await JSZip.loadAsync(buffer);

    // 1. Parse Theme Colors
    await this.parseThemeColors();

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

  private extractHexColor(node: Element): string | null {
    // 1. Direct srgbClr
    const srgbNode = querySelector(node, "a\\:srgbClr, srgbClr");
    if (srgbNode) {
      const val = srgbNode.getAttribute("val");
      if (val) return `#${val}`;
    }

    // 2. sysClr
    const sysNode = querySelector(node, "a\\:sysClr, sysClr");
    if (sysNode) {
      const val = sysNode.getAttribute("lastClr");
      if (val) return `#${val}`;
    }

    // 3. schemeClr (nested reference)
    const schemeNode = querySelector(node, "a\\:schemeClr, schemeClr");
    if (schemeNode) {
      const val = schemeNode.getAttribute("val");
      if (val && this.themeColors[val]) {
        return this.themeColors[val];
      }
    }

    return null;
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
    const relsPath = `${slideDir}/_rels/${slideFileName}.rels`;
    const relsFile = this.zip?.file(relsPath);
    const imgRelMap: Record<string, string> = {};

    if (relsFile) {
      const relsText = await relsFile.async("text");
      const relsDoc = new DOMParser().parseFromString(relsText, "application/xml");
      const relationships = querySelectorAll(relsDoc, "Relationship");
      for (const rel of relationships) {
        const id = rel.getAttribute("Id");
        const target = rel.getAttribute("Target");
        const type = rel.getAttribute("Type");
        if (id && target && type && type.includes("image")) {
          // Resolve target relative to slide directory
          // e.g. Target="../media/image1.png" relative to "ppt/slides" -> "ppt/media/image1.png"
          const resolvedTarget = this.resolveRelativePath(slideDir, target);
          imgRelMap[id] = resolvedTarget;
        }
      }
    }

    const elements: Element[] = [];

    // Parse shape tree
    const spTree = querySelector(doc, "p\\:spTree, spTree");
    if (spTree) {
      const children = Array.from(spTree.childNodes) as Element[];
      for (const node of children) {
        if (node.nodeType !== 1) continue; // Only elements
        const tagName = node.localName || node.nodeName;

        if (tagName.endsWith("sp")) {
          // Shape or text box
          await this.parseShapeNode(node, scaleX, scaleY, elements);
        } else if (tagName.endsWith("pic")) {
          // Picture element
          await this.parsePicNode(node, scaleX, scaleY, imgRelMap, elements);
        }
      }
    }

    return { id: slideId, elements, rawXml: xmlText };
  }

  private resolveRelativePath(baseDir: string, relativePath: string): string {
    // baseDir: "ppt/slides", relativePath: "../media/image1.png"
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

  private async parseShapeNode(node: Element, scaleX: number, scaleY: number, elements: Element[]) {
    const nvSpPr = querySelector(node, "p\\:nvSpPr, nvSpPr");
    const cNvPr = nvSpPr ? querySelector(nvSpPr, "p\\:cNvPr, cNvPr") : null;
    const id = cNvPr?.getAttribute("id") || `shape_${Math.random().toString(36).substr(2, 9)}`;
    const name = cNvPr?.getAttribute("name") || id;

    const spPr = querySelector(node, "p\\:spPr, spPr");
    if (!spPr) return;

    const rect = this.parseXfrm(spPr, scaleX, scaleY);
    if (!rect) return;

    // Check geometry
    const prstGeom = querySelector(spPr, "a\\:prstGeom, prstGeom");
    let shapeType: "rect" | "ellipse" | "triangle" = "rect";
    if (prstGeom) {
      const prst = prstGeom.getAttribute("prst");
      if (prst === "ellipse" || prst === "oval") {
        shapeType = "ellipse";
      } else if (prst === "triangle") {
        shapeType = "triangle";
      }
    }

    // Fill Color
    let fill = "#cccccc"; // Default shape fill
    const solidFill = querySelector(spPr, "a\\:solidFill, solidFill");
    if (solidFill) {
      const color = this.extractHexColor(solidFill);
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

    // Check if it has solid fill or style, draw background shape if it is not transparent
    // PPTX shape can be a placeholder with no outline/fill (pure text container).
    // Let's output shape only if there is a fill or if it's explicitly drawn.
    // If no solidFill and no fillRef, we might have noFill.
    const noFill = querySelector(spPr, "a\\:noFill, noFill");
    const hasBg = !noFill && (solidFill || querySelector(node, "p\\:style, style"));

    if (hasBg) {
      elements.push({
        type: "shape",
        id: `shape_bg_${id}`,
        rect: { ...rect },
        shapeType,
        fill
      });
    }

    // Text Content
    const txBody = querySelector(node, "p\\:txBody, txBody");
    if (txBody) {
      const paragraphs = querySelectorAll(txBody, "a\\:p, p");
      let fullText = "";
      let fontSize = 24; // Default text size (px)
      let color = "#333333"; // Default text color
      let bold = false;
      let align: "left" | "center" | "right" = "left";

      for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        
        // Paragraph alignment
        const pPr = querySelector(p, "a\\:pPr, pPr");
        if (pPr) {
          const algnAttr = pPr.getAttribute("algn");
          if (algnAttr === "ctr") align = "center";
          else if (algnAttr === "r") align = "right";
        }

        // Text Runs
        const runs = querySelectorAll(p, "a\\:r, r, a\\:fld, fld, a\\:br, br");
        let paraText = "";

        for (const run of runs) {
          const runName = run.localName || run.nodeName;
          if (runName.endsWith("br")) {
            paraText += "\n";
            continue;
          }

          const txtNode = querySelector(run, "a\\:t, t");
          if (txtNode) {
            paraText += txtNode.textContent || "";
          }

          // Extract text style from the first run that defines them
          const rPr = querySelector(run, "a\\:rPr, rPr");
          if (rPr) {
            const szAttr = rPr.getAttribute("sz");
            if (szAttr) {
              // sz is in 100ths of a point. E.g. sz="2400" -> 24pt
              // Convert pt to pixels: 1pt = 1.333px.
              // Wait, let's keep font size proportional!
              const ptSize = parseInt(szAttr, 10) / 100;
              fontSize = Math.round(ptSize * 1.33 * scaleY * 1.3); // Scale text size proportionally
            }

            const boldAttr = rPr.getAttribute("b");
            if (boldAttr === "1" || boldAttr === "true") {
              bold = true;
            }

            const runColor = this.extractHexColor(rPr);
            if (runColor) {
              color = runColor;
            }
          }
        }

        if (paraText) {
          fullText += (fullText ? "\n" : "") + paraText;
        }
      }

      if (fullText.trim()) {
        elements.push({
          type: "text",
          id: `text_${id}`,
          rect: {
            x: rect.x + 8, // slight margin
            y: rect.y + 4,
            w: Math.max(10, rect.w - 16),
            h: Math.max(10, rect.h - 8)
          },
          content: fullText,
          style: {
            fontSize: Math.max(10, fontSize),
            color,
            bold,
            align
          }
        });
      }
    }
  }

  private async parsePicNode(node: Element, scaleX: number, scaleY: number, imgRelMap: Record<string, string>, elements: Element[]) {
    const nvPicPr = querySelector(node, "p\\:nvPicPr, nvPicPr");
    const cNvPr = nvPicPr ? querySelector(nvPicPr, "p\\:cNvPr, cNvPr") : null;
    const id = cNvPr?.getAttribute("id") || `pic_${Math.random().toString(36).substr(2, 9)}`;

    const spPr = querySelector(node, "p\\:spPr, spPr");
    if (!spPr) return;

    const rect = this.parseXfrm(spPr, scaleX, scaleY);
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

    elements.push({
      type: "image",
      id: `img_${id}`,
      rect,
      url: blobUrl
    });
  }

  private parseXfrm(spPr: Element, scaleX: number, scaleY: number): Rect | null {
    const xfrm = querySelector(spPr, "a\\:xfrm, xfrm");
    if (!xfrm) return null;

    const off = querySelector(xfrm, "a\\:off, off");
    const ext = querySelector(xfrm, "a\\:ext, ext");
    if (!off || !ext) return null;

    const xStr = off.getAttribute("x");
    const yStr = off.getAttribute("y");
    const cxStr = ext.getAttribute("cx");
    const cyStr = ext.getAttribute("cy");

    if (!xStr || !yStr || !cxStr || !cyStr) return null;

    const x = parseInt(xStr, 10);
    const y = parseInt(yStr, 10);
    const cx = parseInt(cxStr, 10);
    const cy = parseInt(cyStr, 10);

    return {
      x: x * scaleX,
      y: y * scaleY,
      w: cx * scaleX,
      h: cy * scaleY
    };
  }
}
