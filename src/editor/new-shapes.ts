import { ElementSource, Rect, ShapeElement, SlideElement, TextElement, TextRun } from "../pptx-virtual-dom";

export type NewShapeKind =
  | "textbox"
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "rightArrow"
  | "line"
  | "arrow";

/** Look of freshly drawn shapes, close to what PowerPoint starts with. */
export const SHAPE_FILL = "#4472c4";
export const SHAPE_LINE = "#2f528f";
export const SHAPE_TEXT = "#ffffff";
export const TEXTBOX_TEXT = "#000000";
const SHAPE_LINE_PT = 1;
const LINE_PT = 1.5;
const FONT_PT = 18;
const PT_EMU = 12700;
const INSET_X_EMU = 91440;
const INSET_Y_EMU = 45720;

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
  + 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const PRESET: Record<Exclude<NewShapeKind, "textbox" | "line" | "arrow">, string> = {
  rect: "rect",
  roundRect: "roundRect",
  ellipse: "ellipse",
  triangle: "triangle",
  rightArrow: "rightArrow"
};

const NAMES: Record<NewShapeKind, string> = {
  textbox: "文本框",
  rect: "矩形",
  roundRect: "圆角矩形",
  ellipse: "椭圆",
  triangle: "三角形",
  rightArrow: "右箭头",
  line: "直线",
  arrow: "箭头"
};

export const isLine = (kind: NewShapeKind): boolean => kind === "line" || kind === "arrow";
export const displayName = (kind: NewShapeKind, id: string): string => `${NAMES[kind]} ${id}`;

const hex = (color: string) => color.replace("#", "").toUpperCase();
const escapeXml = (text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");

export interface EmuBox {
  x: number;
  y: number;
  cx: number;
  cy: number;
  flipH?: boolean;
  flipV?: boolean;
}

const xfrm = (box: EmuBox) =>
  `<a:xfrm${box.flipH ? ' flipH="1"' : ""}${box.flipV ? ' flipV="1"' : ""}>`
  + `<a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></a:xfrm>`;

const solid = (color: string) => `<a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill>`;

/** The p:sp / p:cxnSp node for a new shape. */
export function buildShapeXml(kind: NewShapeKind, id: string, box: EmuBox): string {
  const name = escapeXml(displayName(kind, id));
  if (isLine(kind)) {
    const tail = kind === "arrow" ? '<a:tailEnd type="triangle"/>' : "";
    return `<p:cxnSp ${NS}><p:nvCxnSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>`
      + `<p:spPr>${xfrm(box)}<a:prstGeom prst="${kind === "arrow" ? "straightConnector1" : "line"}"><a:avLst/></a:prstGeom>`
      + `<a:ln w="${Math.round(LINE_PT * PT_EMU)}">${solid(SHAPE_FILL)}${tail}</a:ln></p:spPr></p:cxnSp>`;
  }

  const isText = kind === "textbox";
  const geometry = `<a:prstGeom prst="${isText ? "rect" : PRESET[kind as keyof typeof PRESET]}"><a:avLst/></a:prstGeom>`;
  const paint = isText
    ? "<a:noFill/>"
    : `${solid(SHAPE_FILL)}<a:ln w="${Math.round(SHAPE_LINE_PT * PT_EMU)}">${solid(SHAPE_LINE)}</a:ln>`;
  const textColor = isText ? "" : solid(SHAPE_TEXT);
  const body = isText
    ? '<a:bodyPr wrap="square" rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr>'
    : '<a:bodyPr rtlCol="0" anchor="ctr"/>';
  const align = isText ? "" : '<a:pPr algn="ctr"/>';
  return `<p:sp ${NS}><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr${isText ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${xfrm(box)}${geometry}${paint}</p:spPr>`
    + `<p:txBody>${body}<a:lstStyle/><a:p>${align}<a:endParaRPr lang="zh-CN" altLang="en-US" sz="${FONT_PT * 100}" dirty="0">${textColor}</a:endParaRPr></a:p></p:txBody></p:sp>`;
}

/** The p:pic node for a picture whose file is related to the slide as `relId`. */
export function buildPictureXml(id: string, name: string, relId: string, box: EmuBox): string {
  return `<p:pic ${NS}><p:nvPicPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>`
    + `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`
    + `<p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
    + `<p:spPr>${xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

export interface NewShapeContext {
  emuPerPx: number;
  fontFamily?: string;
  eastAsianFontFamily?: string;
}

/** What the renderer needs to draw the shape that `buildShapeXml` describes. */
export function buildShapeElements(
  kind: NewShapeKind,
  source: ElementSource,
  rect: Rect,
  context: NewShapeContext,
  flip: { flipH?: boolean; flipV?: boolean } = {}
): SlideElement[] {
  const px = (emu: number) => emu / context.emuPerPx;
  const id = source.shapeId;

  if (isLine(kind)) {
    const width = LINE_PT * PT_EMU / context.emuPerPx;
    const line: NonNullable<ShapeElement["computedStyle"]>["line"] = {
      fill: { type: "solid", color: SHAPE_FILL },
      width,
      ...(kind === "arrow" ? { tailEnd: { type: "triangle" } } : {})
    };
    return [{
      type: "shape",
      id: `shape_bg_${id}`,
      source,
      rect: { ...rect },
      shapeType: kind === "arrow" ? "straightConnector1" : "line",
      flipH: !!flip.flipH,
      flipV: !!flip.flipV,
      fill: "transparent",
      border: { color: SHAPE_FILL, width },
      computedStyle: { fill: { type: "none" }, line }
    }];
  }

  const isText = kind === "textbox";
  const lineWidth = SHAPE_LINE_PT * PT_EMU / context.emuPerPx;
  const shape: ShapeElement = {
    type: "shape",
    id: `shape_bg_${id}`,
    source,
    rect: { ...rect },
    shapeType: isText ? "rect" : PRESET[kind as keyof typeof PRESET],
    ...(kind === "roundRect" ? { cornerRadius: 0.16667 } : {}),
    fill: isText ? "transparent" : SHAPE_FILL,
    ...(isText ? {} : { border: { color: SHAPE_LINE, width: lineWidth } }),
    computedStyle: isText
      ? { fill: { type: "none" } }
      : {
        fill: { type: "solid", color: SHAPE_FILL },
        line: { fill: { type: "solid", color: SHAPE_LINE }, width: lineWidth }
      }
  };

  const align = isText ? "left" : "center";
  const style: TextRun["style"] = {
    fontSize: FONT_PT * PT_EMU / context.emuPerPx,
    color: isText ? TEXTBOX_TEXT : SHAPE_TEXT,
    bold: false,
    align,
    ...(context.fontFamily ? { fontFamily: context.fontFamily } : {}),
    ...(context.eastAsianFontFamily ? { eastAsianFontFamily: context.eastAsianFontFamily } : {})
  };
  const text: TextElement = {
    type: "text",
    id: `text_${id}`,
    source,
    rect: { ...rect },
    content: "",
    style,
    paragraphs: [{
      style: { align, level: 0, marginLeft: 0, indent: 0 },
      runs: [{ content: "", style: { ...style } }]
    }],
    body: {
      marginLeft: px(INSET_X_EMU),
      marginRight: px(INSET_X_EMU),
      marginTop: px(INSET_Y_EMU),
      marginBottom: px(INSET_Y_EMU),
      verticalAnchor: isText ? "top" : "middle",
      autoFit: "none",
      wrap: "square",
      fontScale: 1,
      textDirection: "horz"
    }
  };
  return [shape, text];
}
