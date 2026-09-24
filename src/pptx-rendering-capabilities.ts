/**
 * Static PPTX rendering capability catalog.
 *
 * This is intentionally metadata-only. It is the inventory for the parser,
 * virtual DOM and Rust renderer; adding an entry here does not claim that a
 * renderer has implemented it. Animation and transition capabilities are
 * deliberately outside this catalog.
 */

export type RenderingCapabilityStatus = "supported" | "partial" | "planned";

export type RenderingCapabilityDomain =
  | "text"
  | "shape"
  | "fill"
  | "line"
  | "effect"
  | "image"
  | "composition";

export enum RenderingCapabilityId {
  TextFont = "text.font",
  TextRunStyle = "text.run-style",
  TextParagraphStyle = "text.paragraph-style",
  TextSpacing = "text.spacing",
  TextWrapping = "text.wrapping",
  TextBullet = "text.bullet",
  TextBody = "text.body",
  TextDirection = "text.direction",
  TextReflection = "text.reflection",
  TextShadow = "text.shadow",
  TextStroke = "text.stroke",
  TextWarp = "text.warp",

  ShapeBasicGeometry = "shape.basic-geometry",
  ShapePresetGeometry = "shape.preset-geometry",
  ShapeAdjustments = "shape.adjustments",
  ShapeTransform = "shape.transform",
  ShapeCustomGeometry = "shape.custom-geometry",
  ShapeConnector = "shape.connector",

  FillNone = "fill.none",
  FillSolid = "fill.solid",
  FillGradient = "fill.gradient",
  FillPattern = "fill.pattern",
  FillPicture = "fill.picture",

  LineStyle = "line.style",
  LineDash = "line.dash",
  LineCompound = "line.compound",
  LineHeadTail = "line.head-tail",

  EffectOuterShadow = "effect.outer-shadow",
  EffectInnerShadow = "effect.inner-shadow",
  EffectGlow = "effect.glow",
  EffectReflection = "effect.reflection",
  EffectSoftEdge = "effect.soft-edge",
  EffectBlur = "effect.blur",
  EffectFillOverlay = "effect.fill-overlay",
  EffectPreset = "effect.preset",

  ImageCrop = "image.crop",
  ImageTransform = "image.transform",
  ImageFill = "image.fill",

  CompositionOrder = "composition.order",
  CompositionClip = "composition.clip",
  CompositionGroupTransform = "composition.group-transform",
  CompositionStyleCascade = "composition.style-cascade",
};

export interface RenderingCapability {
  readonly id: RenderingCapabilityId;
  readonly domain: RenderingCapabilityDomain;
  readonly status: RenderingCapabilityStatus;
  /** DrawingML nodes or attributes that provide the source value. */
  readonly xml: readonly string[];
  /** Current owner of the behavior, when it is implemented. */
  readonly module?: string;
  readonly note: string;
}

export interface ShapeGeometryCapability {
  readonly name: string;
  readonly preset: string;
  readonly family: string;
  readonly status: RenderingCapabilityStatus;
  readonly note: string;
}

export const STATIC_RENDERING_CAPABILITIES: readonly RenderingCapability[] = [
  {
    id: RenderingCapabilityId.TextFont,
    domain: "text",
    status: "supported",
    xml: ["a:rPr@latin", "a:rPr@ea", "a:rPr@cs", "a:latin", "a:ea", "a:cs"],
    module: "src/pptx-parser.ts + rust-engine/src/text_layout.rs",
    note: "Font family, East Asian fallback, size, color, bold, italic and letter spacing.",
  },
  {
    id: RenderingCapabilityId.TextRunStyle,
    domain: "text",
    status: "supported",
    xml: ["a:r", "a:rPr", "a:endParaRPr", "a:defRPr"],
    module: "src/pptx-virtual-dom.ts + rust-engine/src/font_renderer.rs",
    note: "Run-level style is preserved instead of flattening a rich paragraph to one style.",
  },
  {
    id: RenderingCapabilityId.TextParagraphStyle,
    domain: "text",
    status: "supported",
    xml: ["a:p", "a:pPr", "a:defPPr", "a:lvlNpPr"],
    module: "src/pptx-parser.ts + rust-engine/src/text_layout.rs",
    note: "Paragraph alignment, level, margin, indent, before and after spacing.",
  },
  {
    id: RenderingCapabilityId.TextSpacing,
    domain: "text",
    status: "supported",
    xml: ["a:lnSpc", "a:spcPct", "a:spcPts", "a:spcBef", "a:spcAft"],
    module: "src/pptx-parser.ts + rust-engine/src/text_layout.rs",
    note: "Percentage and fixed line spacing use the font metrics and paragraph values.",
  },
  {
    id: RenderingCapabilityId.TextWrapping,
    domain: "text",
    status: "supported",
    xml: ["a:eaLnBrk", "a:hangingPunct", "a:fontAlgn"],
    module: "rust-engine/src/text_layout.rs",
    note: "Measured glyph advances drive line breaks; East Asian line breaking and hanging punctuation are retained.",
  },
  {
    id: RenderingCapabilityId.TextBullet,
    domain: "text",
    status: "partial",
    xml: ["a:buChar", "a:buFont", "a:buSzPct", "a:buSzPts", "a:buClr"],
    module: "src/pptx-parser.ts + rust-engine/src/text_layout.rs",
    note: "Separate bullet layout, font, size and color are supported; complete Office inheritance still needs more cases.",
  },
  {
    id: RenderingCapabilityId.TextBody,
    domain: "text",
    status: "partial",
    xml: ["p:txBody", "a:bodyPr", "a:bodyPr@lIns", "a:bodyPr@rIns", "a:bodyPr@tIns", "a:bodyPr@bIns"],
    module: "src/pptx-parser.ts + rust-engine/src/text_layout.rs",
    note: "Insets, vertical anchor, autofit and overflow are represented; full normAutofit iteration is not complete.",
  },
  {
    id: RenderingCapabilityId.TextDirection,
    domain: "text",
    status: "partial",
    xml: ["a:bodyPr@vert"],
    module: "src/pptx-parser.ts + rust-engine/src/text_layout.rs",
    note: "Horizontal and vertical directions are handled; every DrawingML writing mode still needs independent validation.",
  },
  {
    id: RenderingCapabilityId.TextReflection,
    domain: "text",
    status: "supported",
    xml: ["a:reflection"],
    module: "src/pptx-parser.ts + rust-engine/src/font_renderer.rs",
    note: "Per-line glyph reflection, endPos clipping, alpha gradient and text-box clipping are rendered in Rust.",
  },
  {
    id: RenderingCapabilityId.TextShadow,
    domain: "text",
    status: "partial",
    xml: ["a:rPr/a:effectLst/a:outerShdw"],
    module: "src/pptx-parser.ts + rust-engine/src/font_renderer.rs",
    note: "Run-level outer shadows combine a low-intensity shifted glyph-alpha core with a weaker Gaussian penumbra; XML color, opacity, distance and direction are parsed, with a calibrated 0.35 visible text-offset scale for WPS-like overlap. Multi-run stacking and exact Office blur profiles remain approximate.",
  },
  {
    id: RenderingCapabilityId.TextStroke,
    domain: "text",
    status: "planned",
    xml: ["a:rPr/a:ln", "a:rPr/a:solidFill"],
    module: "rust-engine/src/font_renderer.rs",
    note: "Text outline is distinct from shape line and must be implemented at glyph-mask level.",
  },
  {
    id: RenderingCapabilityId.TextWarp,
    domain: "text",
    status: "planned",
    xml: ["a:prstTxWarp", "a:custGeom"],
    note: "WordArt and transformed text geometry are static effects, but are not yet in the virtual DOM.",
  },

  {
    id: RenderingCapabilityId.ShapeBasicGeometry,
    domain: "shape",
    status: "supported",
    xml: ["a:prstGeom@prst"],
    module: "src/pptx-parser.ts + rust-engine/src/shape_renderer.rs",
    note: "rect, roundRect, ellipse, triangle, line, mathPlus and upArrow are currently rendered.",
  },
  {
    id: RenderingCapabilityId.ShapePresetGeometry,
    domain: "shape",
    status: "partial",
    xml: ["a:prstGeom", "a:prstGeom/a:avLst"],
    module: "src/pptx-parser.ts + rust-engine/src/shape_renderer.rs",
    note: "The catalog below records the broader preset geometry families; only the current basic set is wired to paths.",
  },
  {
    id: RenderingCapabilityId.ShapeAdjustments,
    domain: "shape",
    status: "partial",
    xml: ["a:avLst", "a:gd@name", "a:gd@fmla"],
    module: "src/pptx-parser.ts + rust-engine/src/shape_renderer.rs",
    note: "Preset adjustments are retained; the custom-geometry parser evaluates common DrawingML guide formulas.",
  },
  {
    id: RenderingCapabilityId.ShapeTransform,
    domain: "shape",
    status: "supported",
    xml: ["a:xfrm", "a:xfrm@rot", "a:xfrm@flipH", "a:xfrm@flipV"],
    module: "src/pptx-parser.ts + rust-engine/src/shape_renderer.rs",
    note: "Rotation and horizontal/vertical flips are applied around the shape center.",
  },
  {
    id: RenderingCapabilityId.ShapeCustomGeometry,
    domain: "shape",
    status: "partial",
    xml: ["a:custGeom", "a:avLst", "a:gdLst", "a:pathLst", "a:path", "a:moveTo", "a:lnTo", "a:quadBezTo", "a:cubicBezTo", "a:arcTo"],
    module: "src/pptx-custom-geometry.ts + rust-engine/src/shape_geometry/custom.rs",
    note: "Custom paths, guide formulas, Bezier curves, ellipse arcs, path-level fill/stroke switches and lighten/darken fill modes are rendered with Canvas blend approximations; exact arcTo semantics still need more Office samples.",
  },
  {
    id: RenderingCapabilityId.ShapeConnector,
    domain: "shape",
    status: "planned",
    xml: ["p:cxnSp", "a:cxnSpPr", "a:stCxn", "a:endCxn"],
    note: "Connectors need endpoint attachment and route-aware geometry.",
  },

  {
    id: RenderingCapabilityId.FillNone,
    domain: "fill",
    status: "supported",
    xml: ["a:noFill", "a:fillRef@idx=0"],
    module: "src/pptx-style-resolver.ts + rust-engine/src/shape_renderer.rs",
    note: "No fill is kept distinct from a theme or accent solid fill.",
  },
  {
    id: RenderingCapabilityId.FillSolid,
    domain: "fill",
    status: "supported",
    xml: ["a:solidFill", "a:srgbClr", "a:schemeClr", "a:prstClr", "a:scrgbClr"],
    module: "src/pptx-style-resolver.ts + rust-engine/src/shape_renderer.rs",
    note: "Direct, preset and theme colors are resolved with alpha modifiers.",
  },
  {
    id: RenderingCapabilityId.FillGradient,
    domain: "fill",
    status: "partial",
    xml: ["a:gradFill", "a:gsLst", "a:gs", "a:lin", "a:path"],
    module: "src/pptx-style-resolver.ts + rust-engine/src/shape_renderer.rs",
    note: "Linear and radial gradients are supported; advanced path-gradient geometry needs more work.",
  },
  {
    id: RenderingCapabilityId.FillPattern,
    domain: "fill",
    status: "partial",
    xml: ["a:pattFill", "a:prst", "a:fgClr", "a:bgClr"],
    module: "src/pptx-style-resolver.ts + rust-engine/src/shape_renderer.rs",
    note: "Pattern metadata is preserved; the current Canvas fallback paints the foreground color until tiled pattern surfaces are added.",
  },
  {
    id: RenderingCapabilityId.FillPicture,
    domain: "fill",
    status: "partial",
    module: "src/pptx-style-resolver.ts + rust-engine/src/shape_renderer.rs",
    xml: ["a:blipFill", "a:blip", "a:srcRect"],
    note: "Relationship and crop metadata are parsed; raster image sampling inside arbitrary shape paths remains pending.",
  },

  {
    id: RenderingCapabilityId.LineStyle,
    domain: "line",
    status: "supported",
    xml: ["a:ln", "a:ln/a:solidFill", "a:ln@w", "a:ln@cap", "a:ln@cmpd", "a:ln@algn"],
    module: "src/pptx-style-resolver.ts + rust-engine/src/shape_renderer.rs",
    note: "Width, fill, cap, join and basic line alignment are carried into Canvas.",
  },
  {
    id: RenderingCapabilityId.LineDash,
    domain: "line",
    status: "supported",
    xml: ["a:prstDash", "a:custDash"],
    module: "rust-engine/src/shape_renderer.rs",
    note: "Preset dash patterns are mapped to Canvas line dash arrays; custom dash stops need exact conversion.",
  },
  {
    id: RenderingCapabilityId.LineCompound,
    domain: "line",
    status: "partial",
    xml: ["a:ln@cmpd"],
    module: "src/pptx-style-resolver.ts",
    note: "The value is preserved in the model, but Canvas does not yet reproduce multi-line compounds exactly.",
  },
  {
    id: RenderingCapabilityId.LineHeadTail,
    domain: "line",
    status: "partial",
    module: "src/pptx-style-resolver.ts + rust-engine/src/shape_renderer.rs",
    xml: ["a:headEnd", "a:tailEnd"],
    note: "Basic triangle, diamond, oval and open markers are rendered for straight lines; exact stealth sizing and routed connectors remain.",
  },

  {
    id: RenderingCapabilityId.EffectOuterShadow,
    domain: "effect",
    status: "supported",
    xml: ["a:outerShdw", "a:effectLst", "a:effectDag"],
    module: "src/pptx-style-resolver.ts + rust-engine/src/effects.rs",
    note: "Custom alpha mask, blur, scale, offset, direction and z-order composition are used.",
  },
  {
    id: RenderingCapabilityId.EffectInnerShadow,
    domain: "effect",
    status: "partial",
    module: "src/pptx-style-resolver.ts + rust-engine/src/effects.rs",
    xml: ["a:innerShdw"],
    note: "Parsed and rendered through the shape mask pipeline; complex effect stacking is still simplified.",
  },
  {
    id: RenderingCapabilityId.EffectGlow,
    domain: "effect",
    status: "partial",
    xml: ["a:glow"],
    module: "src/pptx-style-resolver.ts + rust-engine/src/effects.rs",
    note: "Glow uses the same custom mask pipeline with its radius, color and opacity.",
  },
  {
    id: RenderingCapabilityId.EffectReflection,
    domain: "effect",
    status: "supported",
    xml: ["a:reflection"],
    module: "src/pptx-parser.ts + rust-engine/src/font_renderer.rs",
    note: "Text reflection is implemented; shape/image reflection remains separate.",
  },
  {
    id: RenderingCapabilityId.EffectSoftEdge,
    domain: "effect",
    status: "partial",
    module: "src/pptx-style-resolver.ts + rust-engine/src/effects.rs",
    xml: ["a:softEdge"],
    note: "A blurred destination-in mask provides the 2D edge falloff; exact Office edge profiles may differ.",
  },
  {
    id: RenderingCapabilityId.EffectBlur,
    domain: "effect",
    status: "partial",
    module: "src/pptx-style-resolver.ts + rust-engine/src/ast.rs",
    xml: ["a:blur"],
    note: "Blur parameters are preserved in the AST; source bitmap blur is not yet applied to every element type.",
  },
  {
    id: RenderingCapabilityId.EffectFillOverlay,
    domain: "effect",
    status: "partial",
    module: "src/pptx-style-resolver.ts + rust-engine/src/ast.rs",
    xml: ["a:fillOverlay"],
    note: "Overlay fill is parsed; compositing multiple fills is intentionally deferred until image/pattern fill surfaces are complete.",
  },
  {
    id: RenderingCapabilityId.EffectPreset,
    domain: "effect",
    status: "planned",
    xml: ["a:prstShdw", "a:prstClr", "a:scene3d", "a:sp3d"],
    note: "Preset shadow and 3D scene/material lighting are static effects but are not yet modeled.",
  },

  {
    id: RenderingCapabilityId.ImageCrop,
    domain: "image",
    status: "supported",
    xml: ["p:pic", "p:blipFill", "a:srcRect"],
    module: "src/pptx-parser.ts + rust-engine/src/image_renderer.rs",
    note: "Source cropping and high-quality Canvas smoothing are supported.",
  },
  {
    id: RenderingCapabilityId.ImageTransform,
    domain: "image",
    status: "supported",
    xml: ["a:xfrm", "a:xfrm@rot", "a:xfrm@flipH", "a:xfrm@flipV"],
    module: "src/pptx-parser.ts + rust-engine/src/image_renderer.rs",
    note: "Image rotation and horizontal/vertical flips are applied around the image center.",
  },
  {
    id: RenderingCapabilityId.ImageFill,
    domain: "image",
    status: "partial",
    module: "src/pptx-parser.ts + rust-engine/src/shape_renderer.rs",
    xml: ["a:blipFill", "a:blip", "a:srcRect"],
    note: "Relationship images are loaded and clipped to the shape path; tile mode and full fill transform semantics remain.",
  },

  {
    id: RenderingCapabilityId.CompositionOrder,
    domain: "composition",
    status: "supported",
    xml: ["p:spTree", "p:spTree child order"],
    module: "src/pptx-parser.ts + rust-engine/src/lib.rs",
    note: "Elements are painted in slide tree order, so later elements cover earlier ones.",
  },
  {
    id: RenderingCapabilityId.CompositionClip,
    domain: "composition",
    status: "partial",
    xml: ["a:bodyPr@vertOverflow", "a:bodyPr@horzOverflow", "a:srcRect"],
    module: "rust-engine/src/lib.rs + rust-engine/src/font_renderer.rs",
    note: "Text and reflection clipping are handled; general clip paths are not yet universal.",
  },
  {
    id: RenderingCapabilityId.CompositionGroupTransform,
    domain: "composition",
    status: "partial",
    xml: ["p:grpSp", "p:grpSpPr", "a:xfrm"],
    module: "src/pptx-parser.ts",
    note: "Group coordinate scaling is resolved for current content; nested group editing needs more model support.",
  },
  {
    id: RenderingCapabilityId.CompositionStyleCascade,
    domain: "composition",
    status: "supported",
    xml: ["p:style", "a:fillRef", "a:lnRef", "a:effectRef", "p:ph"],
    module: "src/pptx-style-resolver.ts",
    note: "Theme, master, layout, placeholder and shape sources are kept in the style trace.",
  },
] as const;

/**
 * Shape geometry inventory. The status describes the current renderer, not
 * whether the preset exists in the OOXML standard.
 */
export const PPTX_PRESET_SHAPE_CATALOG: readonly ShapeGeometryCapability[] = [
  { name: "Rectangle", preset: "rect", family: "basic", status: "supported", note: "Axis-aligned rectangle." },
  { name: "Rounded rectangle", preset: "roundRect", family: "basic", status: "supported", note: "Uses the OOXML adj value when present." },
  { name: "Ellipse", preset: "ellipse", family: "basic", status: "supported", note: "Ellipse or circle from the shape bounds." },
  { name: "Triangle", preset: "triangle", family: "basic", status: "supported", note: "Preset triangle path." },
  { name: "Line", preset: "line", family: "basic", status: "supported", note: "Open path with line styling." },
  { name: "Plus", preset: "mathPlus", family: "basic", status: "supported", note: "Math plus geometry used by the sample deck." },
  { name: "Up arrow", preset: "upArrow", family: "arrow", status: "supported", note: "Uses short-side geometry for narrow arrows." },
  { name: "Arrow family", preset: "downArrow,leftArrow,rightArrow,leftRightArrow,upDownArrow", family: "arrow", status: "planned", note: "Shared arrow adjustment and shaft/head path logic." },
  { name: "Chevron and notched arrows", preset: "chevron,homePlate,notchedRightArrow", family: "arrow", status: "planned", note: "Preset paths and adjustment guides are not yet evaluated." },
  { name: "Basic polygons", preset: "diamond,parallelogram,trapezoid,pentagon,hexagon,heptagon,octagon,decagon,dodecagon", family: "basic", status: "planned", note: "Polygon path generation can share a regular-polygon helper." },
  { name: "Circular and pie shapes", preset: "arc,chord,pie,pieWedge,blockArc,donut", family: "basic", status: "planned", note: "Requires angle and adjustment guide evaluation." },
  { name: "Stars and banners", preset: "star4,star5,star6,star8,star10,star12,star16,star24,star32,ribbon,banner", family: "decorative", status: "planned", note: "Preset geometry family for static decorative shapes." },
  { name: "Callouts", preset: "wedgeRectCallout,wedgeRoundRectCallout,wedgeEllipseCallout,cloudCallout", family: "callout", status: "planned", note: "Callout tail location depends on adjustment guides." },
  { name: "Flowchart shapes", preset: "flowChartProcess,flowChartDecision,flowChartDocument,flowChartTerminator", family: "flowchart", status: "planned", note: "The catalog starts with common flowchart presets; remaining presets follow the same path registry." },
  { name: "Equation shapes", preset: "mathMinus,mathMultiply,mathDivide,mathEqual,mathNotEqual", family: "equation", status: "planned", note: "Static equation symbols are distinct from text glyphs." },
  { name: "Custom geometry", preset: "custGeom", family: "custom", status: "planned", note: "Arbitrary paths are the fallback for unsupported presets." },
] as const;

export function getRenderingCapability(id: RenderingCapabilityId): RenderingCapability | undefined {
  return STATIC_RENDERING_CAPABILITIES.find(capability => capability.id === id);
}

export function getRenderingCapabilitiesByDomain(
  domain: RenderingCapabilityDomain
): readonly RenderingCapability[] {
  return STATIC_RENDERING_CAPABILITIES.filter(capability => capability.domain === domain);
}
