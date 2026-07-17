export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ReflectionStyle {
  blurRadius: number;
  startAlpha: number;
  endAlpha: number;
  endPosition: number;
  direction: number;
  distance: number;
  scaleY: number;
  alignment?: string;
  rotationWithShape?: boolean;
}

export interface TextStyle {
  fontSize: number;
  color: string;
  bold: boolean;
  align: "left" | "center" | "right";
  fontFamily?: string;
  eastAsianFontFamily?: string;
  italic?: boolean;
  letterSpacing?: number;
  reflection?: ReflectionStyle;
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
  eastAsianLineBreak?: boolean;
  hangingPunctuation?: boolean;
  fontAlignment?: "auto" | "baseline" | "top" | "center" | "bottom";
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
    /** OOXML buFont/typeface used to draw the bullet character. */
    fontFamily?: string;
    /** Absolute pixel size after resolving buSzPct/buSzPts. */
    fontSize?: number;
  };
}

export interface TextBodyProperties {
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  verticalAnchor: "top" | "middle" | "bottom";
  autoFit: "none" | "shape" | "shrink";
  verticalOverflow?: "overflow" | "clip" | "ellipsis";
  horizontalOverflow?: "overflow" | "clip" | "ellipsis";
  fontScale?: number;
  /** DrawingML bodyPr@vert writing direction. */
  textDirection?: "horz" | "vert" | "vert270" | "eaVert" | "mongolianVert" | "wordArtVert" | "wordArtVertRtl";
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

export interface GradientStop {
  position: number;
  color: string;
}

export type FillStyle =
  | { type: "none" }
  | { type: "solid"; color: string }
  | {
      type: "gradient";
      kind: "linear" | "radial";
      stops: GradientStop[];
      angle?: number;
      rotateWithShape?: boolean;
    };

export interface LineStyle {
  fill: FillStyle;
  width: number;
  dash?: string;
  cap?: string;
  join?: string;
  compound?: string;
}

export interface ShadowStyle {
  color: string;
  opacity: number;
  blur: number;
  distance: number;
  direction: number;
  scaleX?: number;
  scaleY?: number;
  alignment?: string;
}

export interface GlowStyle {
  color: string;
  opacity: number;
  radius: number;
}

export interface EffectStyle {
  outerShadow?: ShadowStyle;
  glow?: GlowStyle;
}

export interface ComputedShapeStyle {
  fill: FillStyle;
  line?: LineStyle;
  effects?: EffectStyle;
}

export type StyleSourceKind = "theme" | "master" | "layout" | "placeholder" | "shape";
export type ShapeStyleProperty = "fill" | "line" | "effects";

export interface StyleRule {
  id: string;
  source: {
    kind: StyleSourceKind;
    part: string;
    nodeId?: string;
  };
  parents: string[];
  declarations: Partial<ComputedShapeStyle>;
}

export interface StyleRegistry {
  rules: Record<string, StyleRule>;
}

export interface ShapeElement {
  type: "shape";
  id: string;
  rect: Rect;
  shapeType: "rect" | "roundRect" | "ellipse" | "triangle" | "line" | "mathPlus" | "upArrow";
  /** xfrm rotation, in canvas degrees. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** Normalized OOXML roundRect adjustment (adj / 100000). */
  cornerRadius?: number;
  /** Normalized OOXML upArrow adjustments. */
  arrowHeadHeight?: number;
  arrowShaftWidth?: number;
  fill: string;
  border?: Border;
  styleRefs?: string[];
  computedStyle?: ComputedShapeStyle;
  styleTrace?: Partial<Record<ShapeStyleProperty, string>>;
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
  rawXml: string;
}

export interface PresentationSize {
  width: number;
  height: number;
}

export interface PptVirtualDocument {
  size: PresentationSize;
  slides: Slide[];
  styleRegistry: StyleRegistry;
}

/** @deprecated Use PptVirtualDocument. */
export type PresentationAST = Omit<PptVirtualDocument, "styleRegistry"> & {
  styleRegistry?: StyleRegistry;
};

export interface PlaceholderDescriptor {
  idx: string | null;
  type: string;
}

export interface PlaceholderContext {
  layoutShapes: globalThis.Element[];
  masterShapes: globalThis.Element[];
  masterTextStyles: globalThis.Element | null;
  slidePart: string;
  layoutPart: string | null;
  masterPart: string | null;
}

export interface TextInheritanceContext {
  textBodies: globalThis.Element[];
  masterTextStyle: globalThis.Element | null;
}

export interface ComputedRunStyle {
  fontSize: number;
  color: string;
  bold: boolean;
  fontFamily: string;
  eastAsianFontFamily: string;
  italic: boolean;
  letterSpacing: number;
  reflection?: ReflectionStyle;
}

export interface ThemeStyleMatrix {
  fills: globalThis.Element[];
  backgroundFills: globalThis.Element[];
  lines: globalThis.Element[];
  effects: globalThis.Element[];
}

export interface TransformState {
  scaleX: number;
  scaleY: number;
  absoluteUnitScale: number;
  offsetX: number;
  offsetY: number;
}
