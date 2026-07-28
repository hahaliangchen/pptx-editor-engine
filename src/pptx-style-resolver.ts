import {
  ComputedShapeStyle,
  EffectStyle,
  FillStyle,
  GradientStop,
  LineEndStyle,
  LineStyle,
  ShadowStyle,
  ShapeStyleProperty,
  StyleRule,
  ThemeStyleMatrix
} from "./pptx-virtual-dom";
import { getDirectChild, getDirectChildren, hasLocalName, querySelector, querySelectorAll } from "./pptx-xml";

export interface StyleResolverContext {
  themeColors: Record<string, string>;
  themeFonts: Record<string, string>;
  themeStyleMatrix: ThemeStyleMatrix;
  registerStyleRule: (rule: Omit<StyleRule, "id">) => string;
}

export class PptxStyleResolver {
  constructor(private readonly context: StyleResolverContext) {}

  resolveThemeTypeface(typeface: string): string {
    if (!typeface.startsWith("+")) return typeface;
    const resolved = this.context.themeFonts[typeface];
    if (resolved) return resolved;
    if (typeface.endsWith("-ea") || typeface.endsWith("-cs")) {
      return this.context.themeFonts[`${typeface.substring(0, 3)}-lt`] || "sans-serif";
    }
    return "sans-serif";
  }

  extractDirectFillColor(node: Element, placeholderColor?: string | null): string | null {
    const solidFill = getDirectChild(node, "solidFill");
    return solidFill ? this.extractHexColor(solidFill, placeholderColor) : null;
  }

  extractHexColor(node: Element, placeholderColor?: string | null): string | null {
    const colorNames = ["srgbClr", "sysClr", "prstClr", "scrgbClr", "schemeClr"];
    const colorNode = hasLocalName(node, colorNames)
      ? node
      : getDirectChild(node, ...colorNames);
    if (!colorNode) return null;

    const colorType = colorNode.localName || colorNode.nodeName.replace(/^.*:/, "");
    let hex: string | null = null;
    if (colorType === "srgbClr") {
      const value = colorNode.getAttribute("val");
      if (value) hex = `#${value}`;
    } else if (colorType === "sysClr") {
      const value = colorNode.getAttribute("lastClr");
      if (value) hex = `#${value}`;
    } else if (colorType === "prstClr") {
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
      hex = presetColors[colorNode.getAttribute("val") || ""] || null;
    } else if (colorType === "scrgbClr") {
      const channel = (name: string) => Math.round(
        Math.max(0, Math.min(100000, parseInt(colorNode.getAttribute(name) || "0", 10))) * 255 / 100000
      );
      hex = `#${channel("r").toString(16).padStart(2, "0")}${channel("g").toString(16).padStart(2, "0")}${channel("b").toString(16).padStart(2, "0")}`;
    } else if (colorType === "schemeClr") {
      const value = colorNode.getAttribute("val");
      if (value === "phClr" && placeholderColor) {
        hex = placeholderColor;
      } else if (value && this.context.themeColors[value]) {
        hex = this.context.themeColors[value];
      }
    }

    if (!hex) return null;

    try {
      let r = parseInt(hex.substring(1, 3), 16);
      let g = parseInt(hex.substring(3, 5), 16);
      let b = parseInt(hex.substring(5, 7), 16);
      const children = Array.from(colorNode.childNodes) as Element[];
      for (const child of children) {
        if (child.nodeType !== 1) continue;
        const tagName = child.localName || child.nodeName;
        const value = child.getAttribute("val");
        if (!value) continue;
        const factor = parseInt(value, 10) / 100000;
        if (tagName.endsWith("tint")) {
          r = Math.round(r * factor + 255 * (1 - factor));
          g = Math.round(g * factor + 255 * (1 - factor));
          b = Math.round(b * factor + 255 * (1 - factor));
        } else if (tagName.endsWith("shade") || tagName.endsWith("lumMod")) {
          r = Math.round(r * factor);
          g = Math.round(g * factor);
          b = Math.round(b * factor);
        } else if (tagName.endsWith("lumOff")) {
          const offset = Math.round(255 * factor);
          r = Math.min(255, r + offset);
          g = Math.min(255, g + offset);
          b = Math.min(255, b + offset);
        }
      }
      r = Math.max(0, Math.min(255, r));
      g = Math.max(0, Math.min(255, g));
      b = Math.max(0, Math.min(255, b));
      return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
    } catch (_error) {
      return hex;
    }
  }

  extractOpacity(node: Element): number {
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

  parseFillStyle(container: Element, placeholderColor?: string | null): FillStyle | undefined {
    const fillNode = hasLocalName(container, ["solidFill", "gradFill", "noFill", "pattFill", "blipFill"])
      ? container
      : getDirectChild(container, "noFill", "solidFill", "gradFill", "pattFill", "blipFill");
    if (!fillNode) return undefined;
    if (hasLocalName(fillNode, ["noFill"])) return { type: "none" };

    if (hasLocalName(fillNode, ["solidFill"])) {
      const color = this.extractHexColor(fillNode, placeholderColor);
      return color
        ? { type: "solid", color: this.applyColorOpacity(color, this.extractOpacity(fillNode)) }
        : undefined;
    }

    if (hasLocalName(fillNode, ["pattFill"])) {
      const preset = getDirectChild(fillNode, "prst")?.getAttribute("val") || "5Percent";
      const foreground = this.extractHexColor(getDirectChild(fillNode, "fgClr") || fillNode, placeholderColor);
      const background = this.extractHexColor(getDirectChild(fillNode, "bgClr") || fillNode, placeholderColor);
      if (!foreground && !background) return undefined;
      return {
        type: "pattern",
        preset,
        foreground: foreground || "#000000",
        background: background || "#ffffff"
      };
    }

    if (hasLocalName(fillNode, ["blipFill"])) {
      const blip = getDirectChild(fillNode, "blip");
      const srcRect = getDirectChild(fillNode, "srcRect");
      const readCrop = (name: string) => Math.max(0, Math.min(1, parseInt(srcRect?.getAttribute(name) || "0", 10) / 100000));
      return {
        type: "picture",
        embed: blip?.getAttribute("r:embed") || blip?.getAttribute("embed") || undefined,
        mode: getDirectChild(fillNode, "tile") ? "tile" : "stretch",
        srcRect: srcRect
          ? { left: readCrop("l"), top: readCrop("t"), right: readCrop("r"), bottom: readCrop("b") }
          : undefined
      };
    }

    const stopList = getDirectChild(fillNode, "gsLst");
    const stops = (stopList ? getDirectChildren(stopList, "gs") : [])
      .map(stop => {
        const color = this.extractHexColor(stop, placeholderColor);
        if (!color) return null;
        return {
          position: Math.max(0, Math.min(1, parseInt(stop.getAttribute("pos") || "0", 10) / 100000)),
          color: this.applyColorOpacity(color, this.extractOpacity(stop))
        };
      })
      .filter((stop): stop is GradientStop => stop !== null);
    if (stops.length === 0) return undefined;
    // OOXML does not guarantee that gradient stops are serialized in offset
    // order. Canvas gradients are order-sensitive in some engines, so make
    // the effective stop list deterministic before it reaches Rust.
    stops.sort((left, right) => left.position - right.position);

    const linear = getDirectChild(fillNode, "lin");
    const path = getDirectChild(fillNode, "path");
    return {
      type: "gradient",
      kind: path ? "radial" : "linear",
      stops,
      angle: linear ? parseInt(linear.getAttribute("ang") || "0", 10) / 60000 : undefined,
      rotateWithShape: fillNode.getAttribute("rotWithShape") !== "0"
    };
  }

  parseLineStyle(
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
    const parseLineEnd = (name: "headEnd" | "tailEnd"): LineEndStyle | undefined => {
      const end = getDirectChild(line, name);
      if (!end) return undefined;
      const width = end.getAttribute("w");
      const length = end.getAttribute("len");
      return {
        type: end.getAttribute("type") || "none",
        width: width === "sm" || width === "med" || width === "lg" ? width : undefined,
        length: length === "sm" || length === "med" || length === "lg" ? length : undefined
      };
    };
    return {
      fill: resolvedFill,
      width,
      dash: dashNode?.getAttribute("val") || undefined,
      cap: cap === "rnd" ? "round" : cap === "sq" ? "square" : cap === "flat" ? "butt" : undefined,
      join: joinNode ? (joinNode.localName || joinNode.nodeName.replace(/^.*:/, "")) : undefined,
      compound: line.getAttribute("cmpd") || undefined,
      headEnd: parseLineEnd("headEnd"),
      tailEnd: parseLineEnd("tailEnd")
    };
  }

  parseEffectStyle(
    container: Element,
    absoluteUnitScale: number,
    placeholderColor?: string | null
  ): EffectStyle | undefined {
    const effectRoot = hasLocalName(container, ["effectLst", "effectDag"])
      ? container
      : getDirectChild(container, "effectLst", "effectDag") || container;
    const shadow = querySelector(effectRoot, "a\\:outerShdw, outerShdw");
    const innerShadow = querySelector(effectRoot, "a\\:innerShdw, innerShdw");
    const glow = querySelector(effectRoot, "a\\:glow, glow");
    const reflection = querySelector(effectRoot, "a\\:reflection, reflection");
    const softEdge = querySelector(effectRoot, "a\\:softEdge, softEdge");
    const blur = querySelector(effectRoot, "a\\:blur, blur");
    const fillOverlay = querySelector(effectRoot, "a\\:fillOverlay, fillOverlay");
    const result: EffectStyle = {};

    const parseShadow = (node: Element): ShadowStyle | undefined => {
      const color = this.extractHexColor(node, placeholderColor);
      if (!color) return undefined;
      return {
        color,
        opacity: this.extractOpacity(node),
        blur: parseInt(node.getAttribute("blurRad") || "0", 10) * absoluteUnitScale,
        distance: parseInt(node.getAttribute("dist") || "0", 10) * absoluteUnitScale,
        direction: parseInt(node.getAttribute("dir") || "0", 10) / 60000,
        scaleX: parseInt(node.getAttribute("sx") || "100000", 10) / 100000,
        scaleY: parseInt(node.getAttribute("sy") || "100000", 10) / 100000,
        alignment: node.getAttribute("algn") || undefined
      };
    };

    if (shadow) {
      result.outerShadow = parseShadow(shadow);
    }
    if (innerShadow) {
      result.innerShadow = parseShadow(innerShadow);
    }
    if (glow) {
      const color = this.extractHexColor(glow, placeholderColor);
      if (color) {
        result.glow = {
          color,
          opacity: this.extractOpacity(glow),
          radius: parseInt(glow.getAttribute("rad") || "0", 10) * absoluteUnitScale
        };
      }
    }
    if (reflection) {
      result.reflection = {
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
      };
    }
    if (softEdge) result.softEdge = { radius: parseInt(softEdge.getAttribute("rad") || "0", 10) * absoluteUnitScale };
    if (blur) result.blur = { radius: parseInt(blur.getAttribute("rad") || "0", 10) * absoluteUnitScale };
    if (fillOverlay) result.fillOverlay = this.parseFillStyle(fillOverlay, placeholderColor);
    return result.outerShadow || result.innerShadow || result.glow || result.reflection
      || result.softEdge || result.blur || result.fillOverlay ? result : undefined;
  }

  resolveThemeStyleRules(
    styleNode: Element | null,
    absoluteUnitScale: number
  ): Array<{ id: string; property: ShapeStyleProperty; value: FillStyle | LineStyle | EffectStyle }> {
    if (!styleNode) return [];
    const result: Array<{ id: string; property: ShapeStyleProperty; value: FillStyle | LineStyle | EffectStyle }> = [];
    const register = (
      property: ShapeStyleProperty,
      nodeId: string,
      value: FillStyle | LineStyle | EffectStyle
    ) => {
      const id = this.context.registerStyleRule({
        source: { kind: "theme", part: "ppt/theme/theme1.xml", nodeId },
        parents: [],
        declarations: property === "fill"
          ? { fill: value as FillStyle }
          : property === "line"
            ? { line: value as LineStyle }
            : { effects: value as EffectStyle }
      });
      result.push({ id, property, value });
    };

    const fillRef = getDirectChild(styleNode, "fillRef");
    if (fillRef) {
      const index = parseInt(fillRef.getAttribute("idx") || "0", 10);
      const placeholderColor = this.extractHexColor(fillRef);
      const source = index >= 1000
        ? this.context.themeStyleMatrix.backgroundFills[index - 1001]
        : this.context.themeStyleMatrix.fills[index - 1];
      const fill = index <= 0
        ? { type: "none" as const }
        : source
          ? this.parseFillStyle(source, placeholderColor)
          : (placeholderColor ? { type: "solid" as const, color: placeholderColor } : undefined);
      if (fill) register("fill", `fillRef:${index}`, fill);
    }

    const lineRef = getDirectChild(styleNode, "lnRef");
    if (lineRef) {
      const index = parseInt(lineRef.getAttribute("idx") || "0", 10);
      const placeholderColor = this.extractHexColor(lineRef);
      const source = this.context.themeStyleMatrix.lines[index - 1];
      const line = source ? this.parseLineStyle(source, absoluteUnitScale, placeholderColor) : undefined;
      if (line) register("line", `lnRef:${index}`, line);
    }

    const effectRef = getDirectChild(styleNode, "effectRef");
    if (effectRef) {
      const index = parseInt(effectRef.getAttribute("idx") || "0", 10);
      const placeholderColor = this.extractHexColor(effectRef);
      const source = this.context.themeStyleMatrix.effects[index - 1];
      const effects = source
        ? this.parseEffectStyle(source, absoluteUnitScale, placeholderColor)
        : undefined;
      if (effects) register("effects", `effectRef:${index}`, effects);
    }
    return result;
  }
}
