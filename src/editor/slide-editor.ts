import { ChartAdd, ChartElement, ChartModel, cloneChart } from "../charts/model";
import { CHART_REL_TYPE, graphicFrameXml } from "../charts/write";
import { TableElement, TableModel } from "../tables/model";
import { defaultTextColor, isDarkSlide } from "./slide-tone";
import { NewTable, buildTableXml } from "../tables/xml";
import { ElementSource, GroupRef, Rect, ShapeElement, Slide, SlideElement, TextElement, TextParagraph } from "../pptx-virtual-dom";
import { Command, History, HistoryEntry } from "./history";
import { LINE_BREAK, TextFormat, TextOp, applyTextOp, diffText } from "./text-ops";
import { reorderSiblings } from "./arrange";
import { remapShapeXml, shapeIdsIn } from "./clone";
import { InsertedShape, MediaAdd, OrderOp, RelAdd, StructureEdits, popItem, pushItem } from "../pptx-exporter";
import {
  NewShapeKind,
  buildPictureXml,
  buildShapeElements,
  buildShapeXml,
  displayName,
  isLine
} from "./new-shapes";
import { FormatState, FormatToolbar } from "./format-toolbar";
import { SelectionProperties, describeSelection } from "./properties";
import { placeholderPrompt } from "./slides";
import {
  BackgroundSpec,
  ShapeSnapshot,
  ShapeStyleOp,
  Styled,
  gradientFill,
  applyAnchorOp,
  applyShapeStyleOp,
  restoreShape,
  snapshotShape
} from "./shape-ops";
import { normalizeParagraphs, paragraphsText, vdomTextModel } from "./text-model-vdom";

export interface EditorHost {
  /** Element that wraps the canvas; the editor overlay is mounted inside it. */
  container: HTMLElement;
  getSlide(): Slide | null;
  getSlideSize(): { width: number; height: number } | null;
  /** Drop render caches for the slide and repaint it. */
  invalidate(slide: Slide): void;
  /** Start a drag: the host may cache everything except `elements` for cheap repaints. */
  beginDrag(slide: Slide, elements: SlideElement[]): void;
  endDrag(): void;
  /** Make a slide current (used when undo/redo touches another page). */
  showSlide(slide: Slide): void;
  /** Called after an edit is finished so thumbnails can refresh. */
  onCommit(slide: Slide): void;
  markDirty(source: ElementSource | undefined): void;
  /** Hide an element's canvas text while it is edited in the DOM overlay. */
  beginTextEdit(slide: Slide, element: TextElement): void;
  endTextEdit(slide: Slide): void;
  recordTextOp(source: ElementSource, op: TextOp): void;
  discardTextOp(source: ElementSource): void;
  recordStyleOp(source: ElementSource, op: ShapeStyleOp): void;
  discardStyleOp(source: ElementSource): void;
  /** Live record of shapes added, reordered or given new relationships. */
  structure(): StructureEdits;
  /** The XML a shape has right now (edits applied), for copying. */
  materialize(source: ElementSource): Promise<{ xml: string; relationships: RelAdd[] } | null>;
  allocateShapeId(part: string): string;
  /** Font the deck uses for body text, so new text looks like the rest. */
  deckFont(): { fontFamily?: string; eastAsianFontFamily?: string };
  allocateRelationshipId(part: string): Promise<string>;
  allocateMediaPath(extension: string): string;
  /** Free part paths for a new chart and its embedded workbook. */
  allocateChartParts(): { part: string; workbook: string };
  /** The deck's six accent colours, for charts that name no colours. */
  themePalette(): string[];
  /** Read a table frame's XML (after it was edited) into a table element placed like `source`. */
  parseTable(xml: string, source: ElementSource): Promise<TableElement | null>;
  /** Give a copy of a chart element's model a chart part of its own (see `cloneChartParts`). */
  cloneChart(model: ChartModel): Promise<{ part: string; add?: ChartAdd }>;
  /** Relationship ids to use in `targetPart` for relationships that belong to `sourcePart`. */
  resolveRelationships(
    targetPart: string,
    sourcePart: string,
    relationships: RelAdd[]
  ): Promise<{ ids: Map<string, string>; adds: RelAdd[] }>;
}

export interface EditorState {
  canUndo: boolean;
  canRedo: boolean;
  selectionCount: number;
  /** Shape being drawn with the pointer, if a tool is armed. */
  tool: NewShapeKind | null;
  /** Look of the selection, for the properties panel; null when nothing is selected. */
  properties: SelectionProperties | null;
  /** True when the deck differs from what was last exported. */
  dirty: boolean;
  /** Every step made so far, and how many of them are currently applied. */
  history: { entries: HistoryEntry[]; position: number };
}

interface SceneObject {
  key: string;
  /** What gets written back on export: the shape itself, or the group. */
  source: ElementSource;
  isGroup: boolean;
  /** Ids of the groups the object sits in (its siblings share this). */
  parent: string;
  /** Groups entered above the object. */
  depth: number;
  elements: SlideElement[];
  rect: Rect;
  rotation: number;
  order: number;
}

type HandleId = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";
const HANDLES: HandleId[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

interface DragItem {
  element: SlideElement;
  before: Rect;
  owner: ElementSource;
}

type Drag =
  | { kind: "move"; slide: Slide; start: Point; items: DragItem[]; bbox: Rect; keys: Set<string>; active: boolean; collapseTo: string | null }
  | { kind: "resize"; slide: Slide; start: Point; items: DragItem[]; bbox: Rect; handle: HandleId; lockRatio: boolean; rotation: number }
  | { kind: "rotate"; slide: Slide; items: Array<{ element: SlideElement; before: number; owner: ElementSource }>; center: Point; startAngle: number; base: number }
  | { kind: "marquee"; slide: Slide; start: Point; base: Set<string> }
  | { kind: "draw"; slide: Slide; start: Point; shape: NewShapeKind; end: Point; constrain: boolean };

interface Point {
  x: number;
  y: number;
}

interface Guide {
  axis: "x" | "y";
  pos: number;
  from: number;
  to: number;
}

const MIN_SIZE = 8;
const SNAP_PX = 6;
const DRAG_THRESHOLD_PX = 3;

const rotationOf = (element: SlideElement): number => element.rotation || 0;

const union = (rects: Rect[]): Rect => {
  const left = Math.min(...rects.map(r => r.x));
  const top = Math.min(...rects.map(r => r.y));
  const right = Math.max(...rects.map(r => r.x + r.w));
  const bottom = Math.max(...rects.map(r => r.y + r.h));
  return { x: left, y: top, w: right - left, h: bottom - top };
};

const cloneRect = (rect: Rect): Rect => ({ x: rect.x, y: rect.y, w: rect.w, h: rect.h });

class RectsCommand implements Command {
  public readonly label: string;
  public mergeKey?: string;

  constructor(
    label: string,
    private slide: Slide,
    private items: Array<{ element: SlideElement; before: Rect; after: Rect; owner: ElementSource }>,
    private editor: SlideEditor,
    mergeKey?: string
  ) {
    this.label = label;
    this.mergeKey = mergeKey;
  }

  apply(): void {
    for (const { element, after } of this.items) element.rect = cloneRect(after);
    this.editor.afterCommand(this.slide, this.items.map(item => item.owner));
  }

  revert(): void {
    for (const { element, before } of this.items) element.rect = cloneRect(before);
    this.editor.afterCommand(this.slide, this.items.map(item => item.owner));
  }

  merge(next: Command): void {
    if (!(next instanceof RectsCommand)) return;
    for (const item of next.items) {
      const mine = this.items.find(candidate => candidate.element === item.element);
      if (mine) mine.after = item.after;
    }
  }
}

const setRotation = (element: SlideElement, degrees: number): void => {
  element.rotation = degrees;
};

class RotateCommand implements Command {
  public readonly label = "旋转";

  constructor(
    private slide: Slide,
    private items: Array<{ element: SlideElement; before: number; after: number; owner: ElementSource }>,
    private editor: SlideEditor
  ) {}

  apply(): void {
    for (const { element, after } of this.items) setRotation(element, after);
    this.editor.afterCommand(this.slide, this.items.map(item => item.owner));
  }

  revert(): void {
    for (const { element, before } of this.items) setRotation(element, before);
    this.editor.afterCommand(this.slide, this.items.map(item => item.owner));
  }
}

class DeleteCommand implements Command {
  public readonly label = "删除";
  private removed: Array<{ element: SlideElement; index: number }>;

  constructor(
    private slide: Slide,
    elements: SlideElement[],
    private owners: ElementSource[],
    private editor: SlideEditor
  ) {
    this.removed = elements
      .map(element => ({ element, index: slide.elements.indexOf(element) }))
      .filter(item => item.index >= 0)
      .sort((a, b) => a.index - b.index);
  }

  apply(): void {
    const gone = new Set(this.removed.map(item => item.element));
    this.slide.elements = this.slide.elements.filter(element => !gone.has(element));
    this.editor.afterCommand(this.slide, this.owners);
  }

  revert(): void {
    for (const { element, index } of this.removed) this.slide.elements.splice(index, 0, element);
    this.editor.afterCommand(this.slide, this.owners);
  }
}

class ShapeStyleCommand implements Command {
  get label(): string {
    if (this.op.kind === "fill") return "填充颜色";
    if (this.op.kind === "gradient") return "渐变填充";
    if (this.op.kind === "opacity") return "透明度";
    if (this.op.kind === "shadow") return "阴影";
    if (this.op.kind === "line") {
      if (this.op.color !== undefined) return "边框颜色";
      return this.op.widthPt !== undefined ? "边框粗细" : "边框线型";
    }
    return "形状样式";
  }

  constructor(
    private slide: Slide,
    private items: Array<{ shape: Styled; before: ShapeSnapshot; after: ShapeSnapshot; op?: ShapeStyleOp }>,
    private op: ShapeStyleOp,
    private editor: SlideEditor,
    private host: EditorHost
  ) {}

  apply(): void {
    for (const { shape, after, op } of this.items) {
      restoreShape(shape, after);
      if (shape.source) this.host.recordStyleOp(shape.source, op ?? this.op);
    }
    this.editor.afterCommand(this.slide, [], false);
  }

  revert(): void {
    for (const { shape, before } of this.items) {
      restoreShape(shape, before);
      if (shape.source) this.host.discardStyleOp(shape.source);
    }
    this.editor.afterCommand(this.slide, [], false);
  }
}

class AnchorCommand implements Command {
  public readonly label = "垂直对齐";

  constructor(
    private slide: Slide,
    private items: Array<{ text: TextElement; before: TextElement["body"]; value: "top" | "middle" | "bottom" }>,
    private editor: SlideEditor,
    private host: EditorHost
  ) {}

  apply(): void {
    for (const { text, value } of this.items) {
      applyAnchorOp(text, value);
      if (text.source) this.host.recordStyleOp(text.source, { kind: "anchor", value });
    }
    this.editor.afterCommand(this.slide, [], false);
  }

  revert(): void {
    for (const { text, before } of this.items) {
      text.body = before;
      if (text.source) this.host.discardStyleOp(text.source);
    }
    this.editor.afterCommand(this.slide, [], false);
  }
}

class ArrangeCommand implements Command {
  public readonly label = "调整层级";

  constructor(
    private slide: Slide,
    private before: SlideElement[],
    private after: SlideElement[],
    private moves: Array<{ part: string; shapeId: string; op: OrderOp }>,
    private editor: SlideEditor,
    private host: EditorHost
  ) {}

  apply(): void {
    this.slide.elements = [...this.after];
    for (const move of this.moves) {
      pushItem(this.host.structure().orders, move.part, { shapeId: move.shapeId, op: move.op });
    }
    this.editor.afterCommand(this.slide, [], false);
  }

  revert(): void {
    this.slide.elements = [...this.before];
    for (const move of this.moves) popItem(this.host.structure().orders, move.part);
    this.editor.afterCommand(this.slide, [], false);
  }
}

interface PastedObject {
  elements: SlideElement[];
  inserted: InsertedShape;
  adds: RelAdd[];
  owner: ElementSource;
  media?: MediaAdd;
  chart?: ChartAdd;
}

class PasteCommand implements Command {
  constructor(
    private slide: Slide,
    private part: string,
    private pasted: PastedObject[],
    private editor: SlideEditor,
    private host: EditorHost,
    public readonly label = "粘贴"
  ) {}

  apply(): void {
    const structure = this.host.structure();
    for (const item of this.pasted) {
      this.slide.elements.push(...item.elements);
      for (const add of item.adds) pushItem(structure.relAdds, this.part, add);
      if (item.media) structure.media.push(item.media);
      if (item.chart) structure.charts.push(item.chart);
      pushItem(structure.inserts, this.part, item.inserted);
    }
    this.editor.afterCommand(this.slide, this.pasted.map(item => item.owner));
  }

  revert(): void {
    const structure = this.host.structure();
    const gone = new Set(this.pasted.flatMap(item => item.elements));
    this.slide.elements = this.slide.elements.filter(element => !gone.has(element));
    for (const item of [...this.pasted].reverse()) {
      popItem(structure.inserts, this.part);
      if (item.media) structure.media.pop();
      if (item.chart) structure.charts.pop();
      for (const _ of item.adds) popItem(structure.relAdds, this.part);
    }
    this.editor.afterCommand(this.slide, [], false);
  }
}

interface ClipboardEntry {
  object: { source: ElementSource; isGroup: boolean; depth: number };
  xml: string;
  relationships: RelAdd[];
  elements: SlideElement[];
}

class ChartCommand implements Command {
  constructor(
    private slide: Slide,
    private element: ChartElement,
    private before: ChartModel,
    private after: ChartModel,
    private editor: SlideEditor,
    private host: EditorHost,
    public readonly label: string
  ) {}

  private write(model: ChartModel): void {
    // The model object is shared with the export's chart list, so it is changed in place.
    const target = this.element.chart;
    for (const key of Object.keys(target)) delete (target as unknown as Record<string, unknown>)[key];
    Object.assign(target, cloneChart(model));
    if (target.part) this.host.structure().chartEdits.set(target.part, target);
    this.editor.afterCommand(this.slide, this.element.source ? [this.element.source] : [], false);
  }

  apply(): void { this.write(this.after); }
  revert(): void { this.write(this.before); }
}

class TableCommand implements Command {
  private previous: string | undefined;

  constructor(
    private slide: Slide,
    private element: TableElement,
    private before: { table: TableModel; rect: Rect },
    private after: { table: TableModel; rect: Rect; xml: string },
    private editor: SlideEditor,
    private host: EditorHost,
    public readonly label: string
  ) {}

  private replaced(): Map<string, string> {
    const structure = this.host.structure();
    const part = this.element.source!.part;
    let shapes = structure.replaced.get(part);
    if (!shapes) {
      shapes = new Map();
      structure.replaced.set(part, shapes);
    }
    return shapes;
  }

  apply(): void {
    const shapes = this.replaced();
    const id = this.element.source!.shapeId;
    this.previous = shapes.get(id);
    shapes.set(id, this.after.xml);
    this.element.table = JSON.parse(JSON.stringify(this.after.table));
    this.element.rect = { ...this.after.rect };
    this.editor.afterCommand(this.slide, [this.element.source!], false);
  }

  revert(): void {
    const shapes = this.replaced();
    const id = this.element.source!.shapeId;
    if (this.previous === undefined) shapes.delete(id);
    else shapes.set(id, this.previous);
    this.element.table = JSON.parse(JSON.stringify(this.before.table));
    this.element.rect = { ...this.before.rect };
    this.editor.afterCommand(this.slide, [this.element.source!], false);
  }
}

class BackgroundCommand implements Command {
  public readonly label = "页面背景";
  private before: { element: ShapeElement; fill: string; computedStyle: ShapeElement["computedStyle"] } | null = null;
  private previous: BackgroundSpec | undefined;
  private created: ShapeElement | null = null;

  constructor(
    private slide: Slide,
    private spec: BackgroundSpec,
    private size: { width: number; height: number },
    private editor: SlideEditor,
    private host: EditorHost
  ) {}

  private fill(): NonNullable<ShapeElement["computedStyle"]>["fill"] {
    return this.spec.gradient ? gradientFill(this.spec.gradient) : { type: "solid", color: this.spec.color ?? "#ffffff" };
  }

  apply(): void {
    const fill = this.fill();
    const paint = fill.type === "gradient" ? fill.stops[0].color : fill.type === "solid" ? fill.color : "transparent";
    const existing = this.slide.elements.find((element): element is ShapeElement => element.type === "shape" && element.id === "slide_background");
    if (existing) {
      this.before = { element: existing, fill: existing.fill, computedStyle: JSON.parse(JSON.stringify(existing.computedStyle ?? null)) ?? undefined };
      existing.fill = paint;
      existing.computedStyle = { fill };
    } else {
      this.created = {
        type: "shape", id: "slide_background", rect: { x: 0, y: 0, w: this.size.width, h: this.size.height },
        shapeType: "rect", fill: paint, computedStyle: { fill }
      };
      this.slide.elements.unshift(this.created);
    }
    const structure = this.host.structure();
    const part = this.slide.part!;
    this.previous = structure.backgrounds.get(part);
    structure.backgrounds.set(part, this.spec);
    this.editor.afterCommand(this.slide, [], false);
  }

  revert(): void {
    if (this.before) {
      this.before.element.fill = this.before.fill;
      this.before.element.computedStyle = this.before.computedStyle ?? undefined;
    }
    if (this.created) this.slide.elements = this.slide.elements.filter(element => element !== this.created);
    this.created = null;
    const structure = this.host.structure();
    const part = this.slide.part!;
    if (this.previous) structure.backgrounds.set(part, this.previous);
    else structure.backgrounds.delete(part);
    this.editor.afterCommand(this.slide, [], false);
  }
}

class CompositeCommand implements Command {
  constructor(public readonly label: string, private commands: Command[]) {}
  apply(): void { this.commands.forEach(command => command.apply()); }
  revert(): void { [...this.commands].reverse().forEach(command => command.revert()); }
}

/** What a program gets back after adding a shape or picture. */
export interface CreatedShape {
  id: string;
  name: string;
  part: string;
  /** The text element of a new text-carrying shape. */
  text?: TextElement;
}

const cloneParagraphs = (paragraphs: TextParagraph[]): TextParagraph[] =>
  JSON.parse(JSON.stringify(paragraphs));

class TextCommand implements Command {
  get label(): string {
    if (this.ops.some(op => op.format)) return "文字格式";
    if (this.ops.some(op => op.paragraph)) return "段落对齐";
    return "编辑文字";
  }

  constructor(
    private slide: Slide,
    private element: TextElement,
    private before: TextParagraph[],
    private after: TextParagraph[],
    private ops: TextOp[],
    private editor: SlideEditor,
    private host: EditorHost
  ) {}

  apply(): void {
    this.element.paragraphs = cloneParagraphs(this.after);
    this.element.content = paragraphsText(this.after).split(LINE_BREAK).join("\n");
    if (this.element.source) for (const op of this.ops) this.host.recordTextOp(this.element.source, op);
    this.editor.afterCommand(this.slide, [], false);
  }

  revert(): void {
    this.element.paragraphs = cloneParagraphs(this.before);
    this.element.content = paragraphsText(this.before).split(LINE_BREAK).join("\n");
    if (this.element.source) for (const _ of this.ops) this.host.discardTextOp(this.element.source);
    this.editor.afterCommand(this.slide, [], false);
  }
}

/**
 * Interactive layer on top of the rendered canvas: selection, move, resize,
 * marquee, smart guides, keyboard shortcuts and undo/redo. All geometry is kept
 * in slide logical pixels; the overlay converts to CSS pixels when it paints.
 */
export class SlideEditor {
  public readonly history = new History();

  private selection = new Set<string>();
  private selectionSlideId: string | null = null;
  private hoverKey: string | null = null;
  private drag: Drag | null = null;
  private guides: Guide[] = [];
  private repaintQueued = false;
  private listeners = new Set<(state: EditorState) => void>();

  private layer: HTMLDivElement;
  private hoverBox: HTMLDivElement;
  private selectBox: HTMLDivElement;
  private marqueeBox: HTMLDivElement;
  private badge: HTMLDivElement;
  private guideLayer: HTMLDivElement;
  private handleNodes = new Map<HandleId, HTMLDivElement>();
  private toolbar: FormatToolbar;
  /** Groups double-clicked into, outermost first; their children are selectable. */
  private entered: string[] = [];
  private groupFrame: HTMLDivElement;
  private hintLayer: HTMLDivElement;
  private tool: NewShapeKind | null = null;
  private drawBox: HTMLDivElement;
  private clipboard: ClipboardEntry[] = [];
  private pasteCount = 0;
  private menu: HTMLDivElement | null = null;
  private textSession: {
    slide: Slide;
    element: TextElement;
    shell: HTMLDivElement;
    body: HTMLDivElement;
    before: TextParagraph[];
  } | null = null;

  constructor(private host: EditorHost) {
    const container = host.container;
    this.layer = this.div("edit-layer");
    this.guideLayer = this.div("edit-guides");
    this.hoverBox = this.div("edit-hover");
    this.groupFrame = this.div("edit-group-frame");
    this.drawBox = this.div("edit-draw-box");
    this.hintLayer = this.div("edit-hints");
    this.selectBox = this.div("edit-selection");
    this.marqueeBox = this.div("edit-marquee");
    this.badge = this.div("edit-badge");
    for (const id of HANDLES) {
      const handle = this.div(`edit-handle edit-handle-${id}`);
      handle.dataset.handle = id;
      this.selectBox.appendChild(handle);
      this.handleNodes.set(id, handle);
    }
    const rotate = this.div("edit-rotate");
    rotate.dataset.handle = "rotate";
    rotate.title = "旋转 (Shift 吸附 15°)";
    this.selectBox.appendChild(rotate);
    this.layer.append(this.hintLayer, this.guideLayer, this.groupFrame, this.drawBox, this.hoverBox, this.selectBox, this.marqueeBox, this.badge);
    container.appendChild(this.layer);
    this.toolbar = new FormatToolbar(this.layer, format => this.applyFormat(format));

    container.addEventListener("pointerdown", e => this.onPointerDown(e));
    container.addEventListener("pointermove", e => this.onPointerMove(e));
    container.addEventListener("pointerup", e => this.onPointerUp(e));
    container.addEventListener("pointercancel", e => this.onPointerUp(e));
    container.addEventListener("pointerleave", () => this.setHover(null));
    container.addEventListener("dblclick", e => this.onDoubleClick(e));
    container.addEventListener("contextmenu", e => this.openMenu(e));
    document.addEventListener("keydown", this.onDocumentKeyDown);
    document.addEventListener("selectionchange", this.onDocumentSelectionChange);
    this.history.subscribe(() => this.emit());
  }

  private onDocumentKeyDown = (e: KeyboardEvent) => this.onKeyDown(e);

  private onDocumentSelectionChange = () => {
    if (this.textSession) this.updateToolbar();
  };

  /** Detach from the page: remove the overlay and the document-level listeners. */
  public dispose(): void {
    document.removeEventListener("keydown", this.onDocumentKeyDown);
    document.removeEventListener("selectionchange", this.onDocumentSelectionChange);
    this.layer.remove();
    this.listeners.clear();
  }

  // ---------------------------------------------------------------- public

  public subscribe(listener: (state: EditorState) => void): void {
    this.listeners.add(listener);
  }

  public getState(): EditorState {
    return {
      canUndo: this.history.canUndo,
      canRedo: this.history.canRedo,
      selectionCount: this.selection.size,
      tool: this.tool,
      dirty: this.history.isDirty,
      history: { entries: this.history.entries(), position: this.history.position },
      properties: this.describeProperties()
    };
  }

  private describeProperties(): SelectionProperties | null {
    const slide = this.host.getSlide();
    if (!slide || this.selection.size === 0) return null;
    const objects = this.selectedObjects(slide);
    const single = objects.length === 1 && !objects[0].isGroup ? objects[0] : null;
    return describeSelection(this.selectedShapes(slide), this.selectedTexts(slide), this.emuPerPx(slide), {
      rect: objects.length ? union(objects.map(object => object.rect)) : null,
      rotation: single ? single.rotation : null
    });
  }

  /** Shapes whose fill and border the panel may change (table cells excluded). */
  private selectedShapes(slide: Slide): ShapeElement[] {
    return this.selectedObjects(slide).flatMap(object => object.elements)
      .filter((element): element is ShapeElement => element.type === "shape" && !!element.source && !element.source.readOnlyText);
  }

  private selectedStyled(slide: Slide): Styled[] {
    return this.selectedObjects(slide).flatMap(object => object.elements)
      .filter((element): element is Styled =>
        (element.type === "shape" || element.type === "image") && !!element.source && !element.source.readOnlyText);
  }

  private selectedTexts(slide: Slide): TextElement[] {
    return this.selectedObjects(slide).flatMap(object => object.elements)
      .filter((element): element is TextElement => element.type === "text" && !!element.source && !element.source.readOnlyText);
  }

  /** Call after a (re)paint or resize so the overlay follows the slide. */
  public layout(): void {
    const slide = this.host.getSlide();
    if (slide && this.selectionSlideId !== slide.id) {
      this.selection.clear();
      this.entered = [];
      this.selectionSlideId = slide.id;
      this.emit();
    }
    this.paintOverlay();
  }

  public reset(): void {
    this.cancelTextSession();
    this.entered = [];
    this.selection.clear();
    this.selectionSlideId = null;
    this.drag = null;
    this.guides = [];
    this.history.clear();
    this.paintOverlay();
    this.emit();
  }

  /** The deck was exported: undoing back to here counts as "no unsaved changes". */
  public markSaved(): void {
    if (this.textSession) this.endTextEdit();
    this.history.markSaved();
  }

  public isDirty(): boolean {
    return this.history.isDirty;
  }

  /** Go back or forward through the history until `position` steps are applied. */
  public jumpTo(position: number): void {
    if (this.textSession) this.endTextEdit();
    this.history.goTo(position);
  }

  public undo(): void {
    if (this.textSession) this.endTextEdit();
    this.history.undo();
  }

  public redo(): void {
    if (this.textSession) this.endTextEdit();
    this.history.redo();
  }

  public deleteSelection(): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    const objects = this.selectedObjects(slide);
    if (objects.length === 0) return;
    const command = new DeleteCommand(
      slide,
      objects.flatMap(object => object.elements),
      objects.map(object => object.source),
      this
    );
    command.apply();
    this.history.record(command);
    this.selection.clear();
    this.host.onCommit(slide);
    this.paintOverlay();
    this.emit();
  }

  /**
   * Position and size typed into the panel, in centimetres. The selection is
   * moved and scaled as one box, exactly as dragging its handles would.
   */
  public setBounds(change: Partial<{ x: number; y: number; w: number; h: number }>, keepRatio = false): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    if (this.textSession) this.endTextEdit();
    const objects = this.selectedObjects(slide);
    if (objects.length === 0) return;

    const emu = this.emuPerPx(slide);
    const px = (cm: number) => cm * 360000 / emu;
    const box = union(objects.map(object => object.rect));
    let w = change.w !== undefined ? Math.max(1, px(change.w)) : box.w;
    let h = change.h !== undefined ? Math.max(1, px(change.h)) : box.h;
    if (keepRatio && box.w > 0 && box.h > 0) {
      if (change.w !== undefined && change.h === undefined) h = w * box.h / box.w;
      else if (change.h !== undefined && change.w === undefined) w = h * box.w / box.h;
    }
    const next: Rect = {
      x: change.x !== undefined ? px(change.x) : box.x,
      y: change.y !== undefined ? px(change.y) : box.y,
      w,
      h
    };
    // A line has no width or height to scale from; give each member the new size outright.
    const sx = box.w > 0.5 ? next.w / box.w : null;
    const sy = box.h > 0.5 ? next.h / box.h : null;

    const items = objects.flatMap(object => object.elements.map(element => {
      const r = element.rect;
      return {
        element,
        before: cloneRect(r),
        after: {
          x: next.x + (r.x - box.x) * (sx ?? 1),
          y: next.y + (r.y - box.y) * (sy ?? 1),
          w: sx === null ? next.w : r.w * sx,
          h: sy === null ? next.h : r.h * sy
        },
        owner: object.source
      };
    })).filter(item => ["x", "y", "w", "h"].some(key =>
      Math.abs((item.after as Record<string, number>)[key] - (item.before as unknown as Record<string, number>)[key]) > 0.01));
    if (items.length === 0) return;
    const command = new RectsCommand("设置位置与大小", slide, items, this);
    command.apply();
    this.history.record(command);
  }

  /** Rotation typed into the panel, in degrees. */
  public setRotationDegrees(degrees: number): void {
    const slide = this.host.getSlide();
    if (!slide || !Number.isFinite(degrees)) return;
    const objects = this.selectedObjects(slide);
    if (objects.length !== 1 || objects[0].isGroup) return;
    const target = Math.round((((degrees % 360) + 360) % 360) * 10) / 10;
    const items = objects[0].elements
      .filter(element => rotationOf(element) !== target)
      .map(element => ({ element, before: rotationOf(element), after: target, owner: objects[0].source }));
    if (items.length === 0) return;
    const command = new RotateCommand(slide, items, this);
    command.apply();
    this.history.record(command);
  }

  /** Fill / border changes from the properties panel. */
  public applyShapeStyle(requested: ShapeStyleOp): void {
    let op = requested;
    const slide = this.host.getSlide();
    if (!slide) return;
    if (this.textSession) this.endTextEdit();
    if (requested.kind === "anchor") {
      const value = requested.value;
      const items = this.selectedTexts(slide).map(text => ({ text, before: text.body, value }));
      if (items.length === 0) return;
      const command = new AnchorCommand(slide, items, this, this.host);
      command.apply();
      this.history.record(command);
      return;
    }

    const emu = this.emuPerPx(slide);
    // A drop shadow also suits a picture; every other look change is for shapes.
    const shapes: Styled[] = op.kind === "shadow" ? this.selectedStyled(slide) : this.selectedShapes(slide);
    // Changing width or dash of a shape without a border should show one.
    const hasBorder = (shape: Styled) =>
      shape.type === "image" ? false
        : shape.computedStyle?.line ? shape.computedStyle.line.fill.type !== "none" : !!shape.border;
    if (op.kind === "line" && op.color === undefined && shapes.length > 0 && !shapes.some(hasBorder)) {
      op = { ...op, color: "#000000" };
    }
    const items = shapes.map(shape => {
      // A fill that comes from a theme style is written out from what is drawn now.
      const own: ShapeStyleOp = op.kind === "opacity" && shape.type === "shape"
        ? { ...op, base: shape.computedStyle?.fill ? JSON.parse(JSON.stringify(shape.computedStyle.fill)) : undefined }
        : op;
      const before = snapshotShape(shape);
      applyShapeStyleOp(shape, own, emu);
      const after = snapshotShape(shape);
      restoreShape(shape, before);
      return { shape, before, after, op: own };
    });
    if (items.length === 0) return;
    const command = new ShapeStyleCommand(slide, items, op, this, this.host);
    command.apply();
    this.history.record(command);
  }

  /** Left / centre / right alignment of paragraphs (the caret's, the selection's, or all of the shape's). */
  public alignParagraphs(align: "left" | "center" | "right"): void {
    this.applyRangeOp({ paragraph: { align } }, true);
  }

  /** Line a selection up with an edge or the middle of the slide. */
  public alignToSlide(edge: "left" | "hcenter" | "right" | "top" | "vcenter" | "bottom"): void {
    const slide = this.host.getSlide();
    const size = this.host.getSlideSize();
    if (!slide || !size) return;
    if (this.textSession) this.endTextEdit();
    const objects = this.selectedObjects(slide);
    if (objects.length === 0) return;

    const items = objects.flatMap(object => {
      let dx = 0;
      let dy = 0;
      const r = object.rect;
      if (edge === "left") dx = -r.x;
      else if (edge === "hcenter") dx = (size.width - r.w) / 2 - r.x;
      else if (edge === "right") dx = size.width - r.w - r.x;
      else if (edge === "top") dy = -r.y;
      else if (edge === "vcenter") dy = (size.height - r.h) / 2 - r.y;
      else dy = size.height - r.h - r.y;
      if (Math.abs(dx) < 0.01 && Math.abs(dy) < 0.01) return [];
      return object.elements.map(element => ({
        element,
        before: cloneRect(element.rect),
        after: { ...element.rect, x: element.rect.x + dx, y: element.rect.y + dy },
        owner: object.source
      }));
    });
    if (items.length === 0) return;
    const command = new RectsCommand("对齐", slide, items, this);
    command.apply();
    this.history.record(command);
  }

  // ------------------------------------------------- layers, copy and paste

  /** Change the stacking order of the selection among its siblings. */
  public arrange(mode: OrderOp): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    if (this.textSession) this.endTextEdit();
    const objects = this.objects(slide);
    const selected = this.selectedObjects(slide);
    if (selected.length === 0) return;

    const before = [...slide.elements];
    const next = [...slide.elements];
    const moves: Array<{ part: string; shapeId: string; op: OrderOp }> = [];
    for (const parent of new Set(selected.map(object => object.parent))) {
      const siblings = objects.filter(object => object.parent === parent);
      const byKey = new Map(siblings.map(object => [object.key, object]));
      const result = reorderSiblings(
        siblings.map(object => object.key),
        new Set(selected.filter(object => object.parent === parent).map(object => object.key)),
        mode
      );
      if (result.moves.length === 0) continue;

      // The siblings keep occupying the same slots; only which one sits where changes.
      const slots = siblings.flatMap(object => object.elements.map(element => before.indexOf(element))).sort((a, b) => a - b);
      const sequence = result.order.flatMap(key => [...byKey.get(key)!.elements].sort((a, b) => before.indexOf(a) - before.indexOf(b)));
      slots.forEach((slot, index) => { next[slot] = sequence[index]; });
      for (const move of result.moves) {
        const object = byKey.get(move.key)!;
        moves.push({ part: object.source.part, shapeId: object.source.shapeId, op: move.op });
      }
    }
    if (moves.length === 0) return;
    const command = new ArrangeCommand(slide, before, next, moves, this, this.host);
    command.apply();
    this.history.record(command);
  }

  /** Record a command that something outside the editor (the slide list) has already applied. */
  public record(command: Command): void {
    this.history.record(command);
  }

  public hasClipboard(): boolean {
    return this.clipboard.length > 0;
  }

  /** Snapshot the selection (with pending edits applied) for pasting. */
  private async collect(slide: Slide): Promise<ClipboardEntry[]> {
    const entries: ClipboardEntry[] = [];
    for (const object of this.selectedObjects(slide)) {
      // A group nested inside an entered group lives in another coordinate space.
      if (object.isGroup && object.depth > 0) continue;
      const snapshot = await this.host.materialize(object.source);
      if (!snapshot) continue;
      entries.push({
        object: { source: object.source, isGroup: object.isGroup, depth: object.depth },
        xml: snapshot.xml,
        relationships: snapshot.relationships,
        elements: JSON.parse(JSON.stringify(object.elements))
      });
    }
    return entries;
  }

  public async copy(): Promise<boolean> {
    const slide = this.host.getSlide();
    if (!slide) return false;
    if (this.textSession) this.endTextEdit();
    const entries = await this.collect(slide);
    if (entries.length === 0) return false;
    this.clipboard = entries;
    this.pasteCount = 0;
    this.emit();
    return true;
  }

  public async cut(): Promise<void> {
    if (await this.copy()) this.deleteSelection();
  }

  public async paste(): Promise<void> {
    const slide = this.host.getSlide();
    if (!slide || this.clipboard.length === 0) return;
    if (this.textSession) this.endTextEdit();
    // Pasting again on the same slide steps the copy further from the original.
    const sameSlide = this.clipboard[0].object.source.part === slide.part;
    this.pasteCount += 1;
    await this.insert(slide, this.clipboard, sameSlide ? 24 * this.pasteCount : 0);
  }

  /** Copy the selection in place, offset a little, without touching the clipboard. */
  public async duplicate(): Promise<void> {
    const slide = this.host.getSlide();
    if (!slide) return;
    if (this.textSession) this.endTextEdit();
    const entries = await this.collect(slide);
    if (entries.length > 0) await this.insert(slide, entries, 24, "创建副本");
  }

  private async insert(slide: Slide, entries: ClipboardEntry[], offset: number, label = "粘贴"): Promise<void> {
    const part = slide.part;
    if (!part) return;
    const base = slide.xform || { scaleX: 1 / this.emuPerPx(slide), scaleY: 1 / this.emuPerPx(slide) };
    const pasted: PastedObject[] = [];

    for (const entry of entries) {
      const isChart = (rel: RelAdd) => rel.type.endsWith("/chart");
      const { ids, adds } = await this.host.resolveRelationships(
        part, entry.object.source.part, entry.relationships.filter(rel => !isChart(rel)));
      // A chart is never shared between two frames: the copy gets a chart part of its own.
      let chart: ChartAdd | undefined;
      let chartCopy: ChartModel | undefined;
      const chartOriginal = entry.elements.find((element): element is ChartElement => element.type === "chart");
      const chartRel = entry.relationships.find(isChart);
      if (chartOriginal && chartRel) {
        chartCopy = cloneChart(chartOriginal.chart);
        const cloned = await this.host.cloneChart(chartCopy);
        chart = cloned.add;
        const pending = [...pasted.flatMap(item => item.adds), ...adds].length;
        const base = parseInt((await this.host.allocateRelationshipId(part)).replace(/\D/g, ""), 10) || 1;
        const relId = `rId${base + pending}`;
        adds.push({ ...chartRel, id: relId, target: `../charts/${cloned.part.split("/").pop()}` });
        ids.set(chartRel.id, relId);
      }
      const remapped = remapShapeXml(entry.xml, () => this.host.allocateShapeId(part), ids);
      const topId = remapped.idMap.get(entry.object.source.shapeId);
      if (!topId) continue;

      const refs = new Map<string, GroupRef>();
      const stamp = Date.now().toString(36);
      const elements = entry.elements.map((original, index) => {
        const element: SlideElement = JSON.parse(JSON.stringify(original));
        const src = original.source!;
        const groups = (src.groups || []).slice(entry.object.depth);
        const rebased = { ...base, offsetX: 0, offsetY: 0 };
        const newGroups = groups.map((group, level) => {
          let ref = refs.get(group.shapeId);
          if (!ref) {
            ref = {
              shapeId: remapped.idMap.get(group.shapeId) || group.shapeId,
              xform: level === 0 ? rebased : group.xform,
              origRect: { x: 0, y: 0, w: 0, h: 0 }
            };
            refs.set(group.shapeId, ref);
          }
          return ref;
        });
        element.source = {
          part,
          shapeId: remapped.idMap.get(src.shapeId) || src.shapeId,
          // A leaf lifted out of an entered group is placed on the slide itself.
          xform: groups.length === 0 && entry.object.depth > 0 ? rebased : src.xform,
          ...(src.readOnlyText ? { readOnlyText: true } : {}),
          ...(newGroups.length ? { groups: newGroups } : {})
        };
        element.id = `${element.id}_c${stamp}${index}`;
        if (element.type === "chart" && chartCopy) element.chart = chartCopy;
        return element;
      });

      for (const [oldId, ref] of refs) {
        const members = elements.filter(element => element.source?.groups?.some(group => group.shapeId === ref.shapeId));
        if (members.length === 0) continue;
        void oldId;
        const left = Math.min(...members.map(m => m.rect.x));
        const top = Math.min(...members.map(m => m.rect.y));
        ref.origRect = {
          x: left,
          y: top,
          w: Math.max(...members.map(m => m.rect.x + m.rect.w)) - left,
          h: Math.max(...members.map(m => m.rect.y + m.rect.h)) - top
        };
      }
      for (const element of elements) {
        element.rect = { ...element.rect, x: element.rect.x + offset, y: element.rect.y + offset };
      }

      const top = entry.object.isGroup
        ? { part, shapeId: topId, xform: elements[0].source!.groups![0].xform }
        : elements[0].source!;
      pasted.push({ elements, inserted: { id: topId, xml: remapped.xml }, adds, owner: top, ...(chart ? { chart } : {}) });
    }
    if (pasted.length === 0) return;

    this.entered = [];
    const command = new PasteCommand(slide, part, pasted, this, this.host, label);
    command.apply();
    this.history.record(command);
    this.selection = new Set(pasted.map(item => `${part}|${item.inserted.id}`));
    this.paintOverlay();
    this.emit();
  }

  // ------------------------------------------------- driving the editor from code

  /** Make exactly these shapes (leaf or group, by file id) the selection. False if one is missing. */
  public selectShapes(refs: Array<{ part: string; id: string }>): boolean {
    const slide = this.host.getSlide();
    if (!slide || refs.length === 0) return false;
    if (this.textSession) this.endTextEdit();
    const keys: string[] = [];
    let entered: string[] | null = null;
    for (const ref of refs) {
      const hit = slide.elements.find(element =>
        element.source?.part === ref.part
        && (element.source.shapeId === ref.id || element.source.groups?.some(group => group.shapeId === ref.id)));
      if (!hit) return false;
      const groups = hit.source!.groups || [];
      const own = groups.findIndex(group => group.shapeId === ref.id);
      // Members of a group are picked from inside it, so step into the groups above the target.
      if (entered === null) entered = (own >= 0 ? groups.slice(0, own) : groups).map(group => group.shapeId);
      keys.push(`${ref.part}|${ref.id}`);
    }
    this.entered = entered ?? [];
    this.selection = new Set(keys);
    this.selectionSlideId = slide.id;
    this.paintOverlay();
    this.emit();
    return true;
  }

  public selectionRefs(): Array<{ part: string; id: string }> {
    return [...this.selection].map(key => {
      const at = key.indexOf("|");
      return { part: key.slice(0, at), id: key.slice(at + 1) };
    });
  }

  /**
   * Apply text operations (replace, restyle, paragraph format) to one text
   * element as a single undoable step. Operations must be in descending order
   * of position when there are several.
   */
  public applyTextOps(element: TextElement, ops: TextOp[]): void {
    const slide = this.host.getSlide();
    if (!slide || !element.source || ops.length === 0) return;
    if (this.textSession) this.endTextEdit();
    const before = cloneParagraphs(element.paragraphs || []);
    const after = cloneParagraphs(before);
    const model = vdomTextModel(after, this.emuPerPx(slide));
    for (const op of ops) applyTextOp(model, op);
    normalizeParagraphs(after, element.style);
    const command = new TextCommand(slide, element, before, after, ops, this, this.host);
    command.apply();
    this.history.record(command);
  }

  // ------------------------------------------------------------ inserting

  /** Arm (or disarm) a drawing tool; the next drag on the slide creates the shape. */
  public setTool(tool: NewShapeKind | null): void {
    if (this.tool === tool) return;
    this.tool = tool;
    if (tool && this.textSession) this.endTextEdit();
    this.host.container.classList.toggle("is-drawing", tool !== null);
    this.setHover(null);
    this.emit();
  }

  private applyDraw(drag: Extract<Drag, { kind: "draw" }>, point: Point, event: PointerEvent): void {
    drag.end = point;
    drag.constrain = event.shiftKey;
    const rect = this.drawRect(drag);
    const k = this.scale() || 1;
    Object.assign(this.drawBox.style, {
      display: "block",
      left: `${rect.x * k}px`,
      top: `${rect.y * k}px`,
      width: `${rect.w * k}px`,
      height: `${rect.h * k}px`
    });
    this.showBadge(drag.slide, rect, `${this.cm(drag.slide, rect.w)} × ${this.cm(drag.slide, rect.h)} cm`);
  }

  private drawRect(drag: Extract<Drag, { kind: "draw" }>): Rect {
    const size = this.host.getSlideSize();
    let dx = drag.end.x - drag.start.x;
    let dy = drag.end.y - drag.start.y;
    if (drag.constrain) {
      if (isLine(drag.shape)) {
        // Lines snap to 45 degree steps.
        const length = Math.hypot(dx, dy);
        const angle = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4);
        dx = Math.cos(angle) * length;
        dy = Math.sin(angle) * length;
      } else {
        const side = Math.max(Math.abs(dx), Math.abs(dy));
        dx = Math.sign(dx || 1) * side;
        dy = Math.sign(dy || 1) * side;
      }
    }
    let x2 = drag.start.x + dx;
    let y2 = drag.start.y + dy;
    if (size) {
      x2 = Math.min(Math.max(x2, 0), size.width);
      y2 = Math.min(Math.max(y2, 0), size.height);
    }
    return {
      x: Math.min(drag.start.x, x2),
      y: Math.min(drag.start.y, y2),
      w: Math.abs(x2 - drag.start.x),
      h: Math.abs(y2 - drag.start.y)
    };
  }

  private finishDraw(drag: Extract<Drag, { kind: "draw" }>): void {
    const size = this.host.getSlideSize();
    let rect = this.drawRect(drag);
    let flipH = drag.end.x < drag.start.x;
    let flipV = drag.end.y < drag.start.y;
    if (drag.constrain && isLine(drag.shape)) {
      flipH = drag.end.x < drag.start.x;
      flipV = drag.end.y < drag.start.y;
    }

    // A plain click drops a default-sized shape centred on the pointer.
    const dragged = Math.hypot(drag.end.x - drag.start.x, drag.end.y - drag.start.y) * (this.scale() || 1) >= 6;
    if (!dragged) {
      const defaults = drag.shape === "textbox" ? { w: 320, h: 64 } : isLine(drag.shape) ? { w: 300, h: 0 } : { w: 240, h: 160 };
      rect = { x: drag.start.x - defaults.w / 2, y: drag.start.y - defaults.h / 2, ...defaults };
      flipH = false;
      flipV = false;
    }
    if (size) {
      rect.x = Math.min(Math.max(rect.x, 0), Math.max(0, size.width - rect.w));
      rect.y = Math.min(Math.max(rect.y, 0), Math.max(0, size.height - rect.h));
    }
    this.insertShape(drag.slide, drag.shape, rect, { flipH, flipV });
  }

  private insertShape(
    slide: Slide,
    kind: NewShapeKind,
    rect: Rect,
    flip: { flipH: boolean; flipV: boolean },
    edit = true,
    nameOverride?: string
  ): CreatedShape | null {
    const part = slide.part;
    this.setTool(null);
    if (!part) return null;
    const id = this.host.allocateShapeId(part);
    const name = nameOverride || displayName(kind, id);
    const base = this.slideTransform(slide);
    const source: ElementSource = { part, shapeId: id, name, xform: base };
    const font = this.host.deckFont();
    const elements = buildShapeElements(kind, source, rect, { emuPerPx: this.emuPerPx(slide), textColor: defaultTextColor(slide), ...font }, flip);
    const xml = buildShapeXml(kind, id, {
      x: Math.round(rect.x / base.scaleX),
      y: Math.round(rect.y / base.scaleY),
      cx: Math.round(rect.w / base.scaleX),
      cy: Math.round(rect.h / base.scaleY),
      ...flip
    }, name);
    this.commitInserted(slide, part, [{ elements, inserted: { id, xml }, adds: [], owner: source }], "插入形状");

    // A new text box drawn by hand is for typing: go straight into it.
    const text = elements.find((element): element is TextElement => element.type === "text");
    if (edit && kind === "textbox" && text) this.startTextEdit(slide, text);
    return { id, name, part, text };
  }

  /** A shape made by a program: same result as drawing it, without entering text editing. */
  public createShape(
    kind: NewShapeKind,
    rect: Rect,
    flip: { flipH: boolean; flipV: boolean } = { flipH: false, flipV: false },
    name?: string
  ): CreatedShape | null {
    const slide = this.host.getSlide();
    return slide ? this.insertShape(slide, kind, rect, flip, false, name) : null;
  }

  /** Add a picture file to the deck and place it on the slide. */
  public async insertImage(image: {
    bytes: Uint8Array;
    extension: string;
    contentType: string;
    url: string;
    width: number;
    height: number;
    name: string;
  }): Promise<CreatedShape | null> {
    const slide = this.host.getSlide();
    const size = this.host.getSlideSize();
    const part = slide?.part;
    if (!slide || !size || !part) return null;
    if (this.textSession) this.endTextEdit();
    this.setTool(null);

    const id = this.host.allocateShapeId(part);
    const relId = await this.host.allocateRelationshipId(part);
    const path = this.host.allocateMediaPath(image.extension);
    const base = this.slideTransform(slide);

    // Fit inside 60% of the slide, keep the picture's proportions, centre it.
    const fit = Math.min(1, (size.width * 0.6) / image.width, (size.height * 0.6) / image.height);
    const w = image.width * fit;
    const h = image.height * fit;
    const rect = { x: (size.width - w) / 2, y: (size.height - h) / 2, w, h };

    const source: ElementSource = { part, shapeId: id, name: image.name, xform: base };
    const element: SlideElement = { type: "image", id: `img_${id}`, source, rect, url: image.url };
    const xml = buildPictureXml(id, image.name, relId, {
      x: Math.round(rect.x / base.scaleX),
      y: Math.round(rect.y / base.scaleY),
      cx: Math.round(rect.w / base.scaleX),
      cy: Math.round(rect.h / base.scaleY)
    });
    this.commitInserted(slide, part, [{
      elements: [element],
      inserted: { id, xml },
      adds: [{
        id: relId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        target: `../media/${path.split("/").pop()}`
      }],
      owner: source,
      media: { path, bytes: image.bytes, extension: image.extension, contentType: image.contentType }
    }], "插入图片");
    return { id, name: image.name, part };
  }

  /** Add a chart to the current slide; `rect` is in slide pixels, centred and sized to the slide when absent. */
  public async createChart(spec: { chart: ChartModel; rect?: Rect; name?: string }): Promise<CreatedShape | null> {
    const slide = this.host.getSlide();
    const size = this.host.getSlideSize();
    const part = slide?.part;
    if (!slide || !size || !part) return null;
    if (this.textSession) this.endTextEdit();
    this.setTool(null);

    const id = this.host.allocateShapeId(part);
    const relId = await this.host.allocateRelationshipId(part);
    const parts = this.host.allocateChartParts();
    const base = this.slideTransform(slide);
    const rect = spec.rect ?? {
      x: size.width * 0.15, y: size.height * 0.18, w: size.width * 0.7, h: size.height * 0.64
    };
    const name = spec.name || `图表 ${id}`;

    const chart = cloneChart(spec.chart);
    chart.palette = this.host.themePalette();
    const font = this.host.deckFont();
    chart.font = {
      ...chart.font,
      // Grey text on a light page, light grey on a dark one.
      color: isDarkSlide(slide) ? "#f2f2f2" : chart.font.color,
      family: chart.font.family ?? font.fontFamily,
      eastAsianFamily: chart.font.eastAsianFamily ?? font.eastAsianFontFamily
    };

    const source: ElementSource = { part, shapeId: id, name, xform: base };
    const element: ChartElement = {
      type: "chart", id: `chart_${id}`, source, rect: { ...rect }, chart, pxPerPt: 12700 * base.scaleX
    };
    const xml = graphicFrameXml(id, name, relId, {
      x: Math.round(rect.x / base.scaleX),
      y: Math.round(rect.y / base.scaleY),
      cx: Math.round(rect.w / base.scaleX),
      cy: Math.round(rect.h / base.scaleY)
    });
    this.commitInserted(slide, part, [{
      elements: [element],
      inserted: { id, xml },
      adds: [{ id: relId, type: CHART_REL_TYPE, target: `../charts/${parts.part.split("/").pop()}` }],
      owner: source,
      chart: { part: parts.part, workbook: parts.workbook, model: chart }
    }], "插入图表");
    return { id, name, part };
  }

  /** Replace a chart's data or settings in one undoable step. */
  public updateChart(element: ChartElement, next: ChartModel, label = "修改图表"): void {
    const slide = this.host.getSlide();
    if (!slide || !slide.elements.includes(element)) return;
    const command = new ChartCommand(slide, element, cloneChart(element.chart), cloneChart(next), this, this.host, label);
    command.apply();
    this.history.record(command);
  }

  /** Add a table to the current slide; `rect` is in slide pixels (height is a minimum), centred when absent. */
  public async createTable(spec: {
    rows: string[][];
    header: boolean;
    fontPt: number;
    rect?: Rect;
    columnWidthsPx?: number[];
    name?: string;
  }): Promise<CreatedShape | null> {
    const slide = this.host.getSlide();
    const size = this.host.getSlideSize();
    const part = slide?.part;
    if (!slide || !size || !part) return null;
    if (this.textSession) this.endTextEdit();
    this.setTool(null);

    const id = this.host.allocateShapeId(part);
    const base = this.slideTransform(slide);
    const rect = spec.rect ?? { x: size.width * 0.1, y: size.height * 0.2, w: size.width * 0.8, h: 0 };
    const name = spec.name || `表格 ${id}`;
    const table: NewTable = {
      rows: spec.rows,
      header: spec.header,
      width: Math.round(rect.w / base.scaleX),
      ...(rect.h > 0 ? { height: Math.round(rect.h / base.scaleY) } : {}),
      ...(spec.columnWidthsPx ? { columnWidths: spec.columnWidthsPx.map(width => Math.round(width / base.scaleX)) } : {}),
      fontPt: spec.fontPt,
      accent: this.host.themePalette()[0] || "#4472c4"
    };
    const xml = buildTableXml(id, name, { x: Math.round(rect.x / base.scaleX), y: Math.round(rect.y / base.scaleY) }, table);
    const source: ElementSource = { part, shapeId: id, name, xform: base };
    const element = await this.host.parseTable(xml, source);
    if (!element) return null;
    this.commitInserted(slide, part, [{ elements: [element], inserted: { id, xml }, adds: [], owner: source }], "插入表格");
    return { id, name, part };
  }

  /** Change a table by rewriting its XML; the table is read again so it is drawn as the file will show it. */
  public async updateTable(element: TableElement, edit: (xml: string) => string, label = "修改表格"): Promise<void> {
    const slide = this.host.getSlide();
    const source = element.source;
    if (!slide || !source || !slide.elements.includes(element)) return;
    if (source.groups?.length) throw new Error("暂不支持修改组合里的表格");
    const snapshot = await this.host.materialize(source);
    if (!snapshot) throw new Error("找不到这个表格");
    const xml = edit(snapshot.xml);
    const parsed = await this.host.parseTable(xml, source);
    if (!parsed) throw new Error("表格改动后无法读取");
    const command = new TableCommand(
      slide, element,
      { table: JSON.parse(JSON.stringify(element.table)), rect: { ...element.rect } },
      { table: parsed.table, rect: parsed.rect, xml },
      this, this.host, label
    );
    command.apply();
    this.history.record(command);
  }

  /** Give the current slide its own background colour or gradient. */
  public setBackground(spec: BackgroundSpec): void {
    const slide = this.host.getSlide();
    const size = this.host.getSlideSize();
    if (!slide?.part || !size) return;
    if (this.textSession) this.endTextEdit();
    const command = new BackgroundCommand(slide, spec, size, this, this.host);
    command.apply();
    this.history.record(command);
  }

  private slideTransform(slide: Slide): ElementSource["xform"] {
    const scale = 1 / this.emuPerPx(slide);
    return { scaleX: slide.xform?.scaleX ?? scale, scaleY: slide.xform?.scaleY ?? scale, offsetX: 0, offsetY: 0 };
  }

  private commitInserted(slide: Slide, part: string, pasted: PastedObject[], label: string): void {
    this.entered = [];
    const command = new PasteCommand(slide, part, pasted, this, this.host, label);
    command.apply();
    this.history.record(command);
    this.selection = new Set(pasted.map(item => `${part}|${item.inserted.id}`));
    this.paintOverlay();
    this.emit();
  }

  // ------------------------------------------------------------ context menu

  private openMenu(event: MouseEvent): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    event.preventDefault();
    this.closeMenu();
    const box = this.box();
    const k = this.scale() || 1;
    const point = { x: (event.clientX - box.left) / k, y: (event.clientY - box.top) / k };
    const object = this.hitInScope(slide, point);
    if (object && !this.selection.has(object.key)) {
      this.selection = new Set([object.key]);
      this.paintOverlay();
      this.emit();
    } else if (!object) {
      this.selection.clear();
      this.paintOverlay();
      this.emit();
    }

    const hasSelection = this.selection.size > 0;
    const item = (label: string, hint: string, enabled: boolean, run: () => void) => ({ label, hint, enabled, run });
    const groups: Array<Array<ReturnType<typeof item>>> = [
      [
        item("剪切", "Ctrl+X", hasSelection, () => void this.cut()),
        item("复制", "Ctrl+C", hasSelection, () => void this.copy()),
        item("粘贴", "Ctrl+V", this.hasClipboard(), () => void this.paste()),
        item("创建副本", "Ctrl+D", hasSelection, () => void this.duplicate())
      ],
      [
        item("置于顶层", "Ctrl+Shift+]", hasSelection, () => this.arrange("front")),
        item("上移一层", "Ctrl+]", hasSelection, () => this.arrange("forward")),
        item("下移一层", "Ctrl+[", hasSelection, () => this.arrange("backward")),
        item("置于底层", "Ctrl+Shift+[", hasSelection, () => this.arrange("back"))
      ],
      [item("删除", "Delete", hasSelection, () => this.deleteSelection())]
    ];

    const menu = this.div("edit-menu");
    groups.forEach((group, index) => {
      if (index > 0) menu.appendChild(this.div("edit-menu-divider"));
      for (const entry of group) {
        const row = document.createElement("button");
        row.type = "button";
        row.className = "edit-menu-item";
        row.disabled = !entry.enabled;
        const label = document.createElement("span");
        label.textContent = entry.label;
        const hint = document.createElement("kbd");
        hint.textContent = entry.hint;
        row.append(label, hint);
        row.addEventListener("click", () => {
          this.closeMenu();
          entry.run();
        });
        menu.appendChild(row);
      }
    });
    menu.addEventListener("pointerdown", e => e.stopPropagation());
    menu.addEventListener("contextmenu", e => e.preventDefault());
    this.layer.appendChild(menu);
    this.menu = menu;

    const width = menu.offsetWidth;
    const height = menu.offsetHeight;
    menu.style.left = `${Math.max(4, Math.min(event.clientX - box.left, box.width - width - 4))}px`;
    menu.style.top = `${Math.max(4, Math.min(event.clientY - box.top, Math.max(4, box.height - height - 4)))}px`;

    const dismiss = (e: Event) => {
      if (e instanceof KeyboardEvent && e.key !== "Escape") return;
      if (e.target instanceof Node && menu.contains(e.target)) return;
      this.closeMenu();
      document.removeEventListener("pointerdown", dismiss, true);
      document.removeEventListener("keydown", dismiss, true);
    };
    document.addEventListener("pointerdown", dismiss, true);
    document.addEventListener("keydown", dismiss, true);
  }

  private closeMenu(): void {
    this.menu?.remove();
    this.menu = null;
  }

  public selectAll(): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    this.selection = new Set(this.objects(slide).map(object => object.key));
    this.paintOverlay();
    this.emit();
  }

  public clearSelection(): void {
    if (this.selection.size === 0) return;
    this.selection.clear();
    this.paintOverlay();
    this.emit();
  }

  /** Used by commands after they change the model. */
  public afterCommand(slide: Slide, owners: ElementSource[], geometry = true): void {
    if (geometry) for (const owner of owners) this.host.markDirty(owner);
    if (this.host.getSlide() !== slide) this.host.showSlide(slide);
    this.selectionSlideId = slide.id;
    this.host.invalidate(slide);
    this.host.onCommit(slide);
    this.selection = new Set(
      [...this.selection].filter(key => this.objects(slide).some(object => object.key === key))
    );
    this.paintOverlay();
    this.emit();
  }

  // ----------------------------------------------------------------- model

  /** Which selectable object an element belongs to, given the groups entered so far. */
  private refFor(element: SlideElement): {
    key: string;
    source: ElementSource;
    isGroup: boolean;
    parent: string;
    depth: number;
  } | null {
    const source = element.source;
    if (!source) return null;
    const groups = source.groups || [];
    let depth = 0;
    while (depth < this.entered.length && depth < groups.length && groups[depth].shapeId === this.entered[depth]) depth++;
    if (groups.length > depth) {
      const group = groups[depth];
      return {
        key: `${source.part}|${group.shapeId}`,
        source: { part: source.part, shapeId: group.shapeId, xform: group.xform },
        isGroup: true,
        parent: groups.slice(0, depth).map(g => g.shapeId).join("/"),
        depth
      };
    }
    return {
      key: `${source.part}|${source.shapeId}`,
      source,
      isGroup: false,
      parent: groups.map(g => g.shapeId).join("/"),
      depth: groups.length
    };
  }

  private objects(slide: Slide): SceneObject[] {
    const byKey = new Map<string, SceneObject>();
    slide.elements.forEach((element, index) => {
      const ref = this.refFor(element);
      if (!ref) return;
      const existing = byKey.get(ref.key);
      if (existing) {
        existing.elements.push(element);
        existing.order = index;
      } else {
        byKey.set(ref.key, {
          key: ref.key,
          source: ref.source,
          isGroup: ref.isGroup,
          parent: ref.parent,
          depth: ref.depth,
          elements: [element],
          rect: element.rect,
          rotation: rotationOf(element),
          order: index
        });
      }
    });
    const result = [...byKey.values()];
    for (const object of result) {
      object.rect = union(object.elements.map(element => element.rect));
      object.rotation = object.isGroup ? 0 : object.elements.reduce((acc, element) => acc || rotationOf(element), 0);
    }
    return result.sort((a, b) => a.order - b.order);
  }

  /** Clicking something outside the entered group leaves it. */
  private hitInScope(slide: Slide, point: Point): SceneObject | null {
    let object = this.hit(slide, point);
    if (this.entered.length > 0 && !this.isInsideEntered(object)) {
      this.entered = [];
      object = this.hit(slide, point);
    }
    return object;
  }

  private isInsideEntered(object: SceneObject | null): boolean {
    if (!object) return false;
    const groups = object.elements[0].source?.groups || [];
    return this.entered.every((id, index) => groups[index]?.shapeId === id);
  }

  private selectedObjects(slide: Slide): SceneObject[] {
    return this.objects(slide).filter(object => this.selection.has(object.key));
  }

  private contains(element: SlideElement, point: Point, pad: number): boolean {
    const { x, y, w, h } = element.rect;
    let px = point.x;
    let py = point.y;
    const rotation = rotationOf(element);
    if (rotation) {
      const radians = (-rotation * Math.PI) / 180;
      const cx = x + w / 2;
      const cy = y + h / 2;
      const dx = px - cx;
      const dy = py - cy;
      px = cx + dx * Math.cos(radians) - dy * Math.sin(radians);
      py = cy + dx * Math.sin(radians) + dy * Math.cos(radians);
    }
    return px >= x - pad && px <= x + w + pad && py >= y - pad && py <= y + h + pad;
  }

  private hit(slide: Slide, point: Point): SceneObject | null {
    const pad = 3 / this.scale();
    const objects = this.objects(slide);
    for (let i = objects.length - 1; i >= 0; i--) {
      if (objects[i].elements.some(element => this.contains(element, point, pad))) return objects[i];
    }
    return null;
  }

  // --------------------------------------------------------------- geometry

  private box(): DOMRect {
    return this.host.container.getBoundingClientRect();
  }

  /** CSS pixels per logical slide pixel. */
  private scale(): number {
    const size = this.host.getSlideSize();
    const box = this.box();
    return size && size.width > 0 ? box.width / size.width : 1;
  }

  private point(event: PointerEvent): Point {
    const box = this.box();
    const k = this.scale() || 1;
    return { x: (event.clientX - box.left) / k, y: (event.clientY - box.top) / k };
  }

  private emuPerPx(slide: Slide): number {
    const size = this.host.getSlideSize();
    return slide.emuPerPx || (size ? 12192000 / size.width : 6350);
  }

  private cm(slide: Slide, px: number): string {
    return (px * this.emuPerPx(slide) / 360000).toFixed(2);
  }

  // ----------------------------------------------------------------- input

  private onPointerDown(event: PointerEvent): void {
    if (event.button !== 0) return;
    if ((event.target as Element).closest?.(".edit-text, .format-toolbar")) return;
    if (this.textSession) this.endTextEdit();
    const slide = this.host.getSlide();
    if (!slide) return;
    event.preventDefault();
    const point = this.point(event);

    if (this.tool) {
      this.drag = { kind: "draw", slide, start: point, end: point, shape: this.tool, constrain: event.shiftKey };
      this.host.container.setPointerCapture(event.pointerId);
      return;
    }
    const handle = (event.target as HTMLElement).closest<HTMLElement>("[data-handle]")?.dataset.handle as HandleId | "rotate" | undefined;
    const selected = this.selectedObjects(slide);

    if (handle === "rotate" && selected.length === 1 && !selected[0].isGroup) {
      const object = selected[0];
      const center = { x: object.rect.x + object.rect.w / 2, y: object.rect.y + object.rect.h / 2 };
      this.drag = {
        kind: "rotate",
        slide,
        items: object.elements.map(element => ({ element, before: rotationOf(element), owner: object.source })),
        center,
        startAngle: Math.atan2(point.y - center.y, point.x - center.x) * 180 / Math.PI,
        base: object.rotation
      };
      this.host.beginDrag(slide, object.elements);
    } else if (handle && handle !== "rotate" && selected.length > 0) {
      const items = selected.flatMap(object =>
        object.elements.map(element => ({ element, before: cloneRect(element.rect), owner: object.source })));
      this.drag = {
        kind: "resize",
        slide,
        start: point,
        items,
        bbox: union(items.map(item => item.before)),
        handle,
        lockRatio: items.every(item => item.element.type === "image"),
        rotation: selected.length === 1 ? selected[0].rotation : 0
      };
      this.host.beginDrag(slide, items.map(item => item.element));
    } else {
      const object = this.hitInScope(slide, point);
      if (object) {
        let collapseTo: string | null = null;
        if (event.shiftKey) {
          if (this.selection.has(object.key)) this.selection.delete(object.key);
          else this.selection.add(object.key);
        } else if (!this.selection.has(object.key)) {
          this.selection = new Set([object.key]);
        } else if (this.selection.size > 1) {
          collapseTo = object.key;
        }
        this.selectionSlideId = slide.id;
        const moving = this.selectedObjects(slide);
        const items = moving.flatMap(o =>
          o.elements.map(element => ({ element, before: cloneRect(element.rect), owner: o.source })));
        this.drag = {
          kind: "move",
          slide,
          start: point,
          items,
          bbox: items.length ? union(items.map(item => item.before)) : { x: 0, y: 0, w: 0, h: 0 },
          keys: new Set(moving.map(o => o.key)),
          active: false,
          collapseTo
        };
      } else {
        this.entered = [];
        const base = event.shiftKey ? new Set(this.selection) : new Set<string>();
        this.selection = new Set(base);
        this.drag = { kind: "marquee", slide, start: point, base };
      }
    }

    this.host.container.setPointerCapture(event.pointerId);
    this.paintOverlay();
    this.emit();
  }

  private onPointerMove(event: PointerEvent): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    const point = this.point(event);

    if (!this.drag) {
      if (this.textSession) return;
      const hovered = this.hit(slide, point);
      this.setHover(hovered ? hovered.key : null);
      this.host.container.classList.toggle(
        "is-over-selection",
        !!hovered && this.selection.has(hovered.key)
      );
      return;
    }

    const drag = this.drag;
    if (drag.kind === "move") this.applyMove(drag, point, event);
    else if (drag.kind === "resize") this.applyResize(drag, point, event);
    else if (drag.kind === "rotate") this.applyRotate(drag, point, event);
    else if (drag.kind === "draw") this.applyDraw(drag, point, event);
    else this.applyMarquee(drag, point);
  }

  private onPointerUp(event: PointerEvent): void {
    const drag = this.drag;
    if (!drag) return;
    this.drag = null;
    this.guides = [];
    this.host.endDrag();
    if (this.host.container.hasPointerCapture(event.pointerId)) {
      this.host.container.releasePointerCapture(event.pointerId);
    }

    if (drag.kind === "draw") {
      this.drawBox.style.display = "none";
      this.badge.style.display = "none";
      this.finishDraw(drag);
      return;
    }
    if (drag.kind === "marquee") {
      this.marqueeBox.style.display = "none";
    } else if (drag.kind === "rotate") {
      const items = drag.items
        .filter(item => rotationOf(item.element) !== item.before)
        .map(item => ({ ...item, after: rotationOf(item.element) }));
      if (items.length > 0) {
        this.history.record(new RotateCommand(drag.slide, items, this));
        items.forEach(item => this.host.markDirty(item.owner));
        this.host.onCommit(drag.slide);
      }
    } else {
      const items = drag.items
        .filter(item => item.element.rect.x !== item.before.x || item.element.rect.y !== item.before.y
          || item.element.rect.w !== item.before.w || item.element.rect.h !== item.before.h)
        .map(item => ({ element: item.element, before: item.before, after: cloneRect(item.element.rect), owner: item.owner }));
      if (items.length > 0) {
        this.history.record(new RectsCommand(
          drag.kind === "move" ? "移动" : "调整大小", drag.slide, items, this
        ));
        items.forEach(item => this.host.markDirty(item.owner));
        this.host.onCommit(drag.slide);
      } else if (drag.kind === "move" && drag.collapseTo) {
        this.selection = new Set([drag.collapseTo]);
      }
    }
    this.badge.style.display = "none";
    this.paintOverlay();
    this.emit();
  }

  private onKeyDown(event: KeyboardEvent): void {
    const target = event.target;
    if (target instanceof Element && target.closest("input, textarea, select, [contenteditable='true']")) return;
    const slide = this.host.getSlide();
    if (!slide) return;
    const mod = event.metaKey || event.ctrlKey;
    const key = event.key.toLowerCase();

    if (mod && key === "z") {
      event.preventDefault();
      if (event.shiftKey) this.redo(); else this.undo();
    } else if (mod && key === "y") {
      event.preventDefault();
      this.redo();
    } else if (mod && key === "a") {
      event.preventDefault();
      this.selectAll();
    } else if (key === "delete" || key === "backspace") {
      if (this.selection.size === 0) return;
      event.preventDefault();
      this.deleteSelection();
    } else if (mod && key === "c" && this.selection.size > 0) {
      event.preventDefault();
      void this.copy();
    } else if (mod && key === "x" && this.selection.size > 0) {
      event.preventDefault();
      void this.cut();
    } else if (mod && key === "v" && this.hasClipboard()) {
      event.preventDefault();
      void this.paste();
    } else if (mod && key === "d" && this.selection.size > 0) {
      event.preventDefault();
      void this.duplicate();
    } else if (mod && (event.code === "BracketRight" || event.code === "BracketLeft") && this.selection.size > 0) {
      event.preventDefault();
      const up = event.code === "BracketRight";
      this.arrange(event.shiftKey ? (up ? "front" : "back") : (up ? "forward" : "backward"));
    } else if (mod && !event.altKey && (key === "b" || key === "i") && this.selection.size > 0) {
      event.preventDefault();
      this.toggleStyle(key === "b" ? "bold" : "italic");
    } else if (key === "escape") {
      if (this.tool) this.setTool(null);
      else this.clearSelection();
    } else if (key.startsWith("arrow") && this.selection.size > 0 && !mod) {
      event.preventDefault();
      const step = event.shiftKey ? 10 : 1;
      const dx = key === "arrowleft" ? -step : key === "arrowright" ? step : 0;
      const dy = key === "arrowup" ? -step : key === "arrowdown" ? step : 0;
      this.nudge(slide, dx, dy);
    }
  }

  private nudge(slide: Slide, dx: number, dy: number): void {
    const objects = this.selectedObjects(slide);
    if (objects.length === 0) return;
    const items = objects.flatMap(object => object.elements.map(element => ({
      element,
      before: cloneRect(element.rect),
      after: { ...element.rect, x: element.rect.x + dx, y: element.rect.y + dy },
      owner: object.source
    })));
    const command = new RectsCommand("移动", slide, items, this, `nudge:${slide.id}:${[...this.selection].join(",")}`);
    command.apply();
    this.history.record(command);
  }

  // ----------------------------------------------------------- drag logic

  private applyMove(drag: Extract<Drag, { kind: "move" }>, point: Point, event: PointerEvent): void {
    let dx = point.x - drag.start.x;
    let dy = point.y - drag.start.y;
    const k = this.scale() || 1;
    if (!drag.active) {
      if (Math.hypot(dx, dy) * k < DRAG_THRESHOLD_PX) return;
      drag.active = true;
      if (drag.items.length > 0) this.host.beginDrag(drag.slide, drag.items.map(item => item.element));
    }
    if (drag.items.length === 0) return;
    if (event.shiftKey) {
      if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0;
    }

    const snap = event.altKey ? { dx: 0, dy: 0, guides: [] as Guide[] } : this.snapMove(drag, dx, dy);
    dx += snap.dx;
    dy += snap.dy;
    this.guides = snap.guides;

    for (const item of drag.items) {
      item.element.rect = { ...item.before, x: item.before.x + dx, y: item.before.y + dy };
    }
    const moved = { ...drag.bbox, x: drag.bbox.x + dx, y: drag.bbox.y + dy };
    this.showBadge(drag.slide, moved, `${this.cm(drag.slide, moved.x)}, ${this.cm(drag.slide, moved.y)} cm`);
    this.queueRepaint(drag.slide);
  }

  private snapMove(drag: Extract<Drag, { kind: "move" }>, dx: number, dy: number) {
    const size = this.host.getSlideSize();
    const empty = { dx: 0, dy: 0, guides: [] as Guide[] };
    if (!size) return empty;
    const threshold = SNAP_PX / (this.scale() || 1);
    const box = { x: drag.bbox.x + dx, y: drag.bbox.y + dy, w: drag.bbox.w, h: drag.bbox.h };

    type Target = { pos: number; from: number; to: number };
    const xTargets: Target[] = [0, size.width / 2, size.width].map(pos => ({ pos, from: 0, to: size.height }));
    const yTargets: Target[] = [0, size.height / 2, size.height].map(pos => ({ pos, from: 0, to: size.width }));
    for (const object of this.objects(drag.slide)) {
      if (drag.keys.has(object.key)) continue;
      const r = object.rect;
      if (r.w >= size.width * 0.95 && r.h >= size.height * 0.95) continue;
      for (const pos of [r.x, r.x + r.w / 2, r.x + r.w]) xTargets.push({ pos, from: r.y, to: r.y + r.h });
      for (const pos of [r.y, r.y + r.h / 2, r.y + r.h]) yTargets.push({ pos, from: r.x, to: r.x + r.w });
    }

    const best = (edges: number[], targets: Target[]) => {
      let delta = 0;
      let distance = threshold;
      for (const edge of edges) {
        for (const target of targets) {
          const d = target.pos - edge;
          if (Math.abs(d) < distance) {
            distance = Math.abs(d);
            delta = d;
          }
        }
      }
      return distance < threshold ? delta : 0;
    };

    const sx = best([box.x, box.x + box.w / 2, box.x + box.w], xTargets);
    const sy = best([box.y, box.y + box.h / 2, box.y + box.h], yTargets);
    const snapped = { x: box.x + sx, y: box.y + sy, w: box.w, h: box.h };

    const guides: Guide[] = [];
    const eps = 0.5;
    for (const edge of [snapped.x, snapped.x + snapped.w / 2, snapped.x + snapped.w]) {
      for (const target of xTargets) {
        if (Math.abs(target.pos - edge) < eps) {
          guides.push({
            axis: "x", pos: target.pos,
            from: Math.min(target.from, snapped.y), to: Math.max(target.to, snapped.y + snapped.h)
          });
        }
      }
    }
    for (const edge of [snapped.y, snapped.y + snapped.h / 2, snapped.y + snapped.h]) {
      for (const target of yTargets) {
        if (Math.abs(target.pos - edge) < eps) {
          guides.push({
            axis: "y", pos: target.pos,
            from: Math.min(target.from, snapped.x), to: Math.max(target.to, snapped.x + snapped.w)
          });
        }
      }
    }
    return { dx: sx, dy: sy, guides };
  }

  private applyRotate(drag: Extract<Drag, { kind: "rotate" }>, point: Point, event: PointerEvent): void {
    const angle = Math.atan2(point.y - drag.center.y, point.x - drag.center.x) * 180 / Math.PI;
    let degrees = drag.base + (angle - drag.startAngle);
    if (event.shiftKey) degrees = Math.round(degrees / 15) * 15;
    degrees = ((degrees % 360) + 360) % 360;
    for (const item of drag.items) setRotation(item.element, Math.round(degrees * 10) / 10);

    const k = this.scale() || 1;
    this.badge.textContent = `${Math.round(degrees)}°`;
    Object.assign(this.badge.style, {
      display: "block",
      left: `${drag.center.x * k}px`,
      top: `${drag.center.y * k + 40}px`
    });
    this.paintOverlay();
    this.queueRepaint(drag.slide);
  }

  private applyResize(drag: Extract<Drag, { kind: "resize" }>, point: Point, event: PointerEvent): void {
    // For a rotated shape the handles live in its own, turned, frame.
    const theta = (drag.rotation * Math.PI) / 180;
    const rawX = point.x - drag.start.x;
    const rawY = point.y - drag.start.y;
    const dx = rawX * Math.cos(theta) + rawY * Math.sin(theta);
    const dy = -rawX * Math.sin(theta) + rawY * Math.cos(theta);
    const b = drag.bbox;
    const h = drag.handle;
    const fromCenter = event.altKey;
    const isCorner = h.length === 2;

    let left = b.x;
    let top = b.y;
    let right = b.x + b.w;
    let bottom = b.y + b.h;
    if (h.includes("w")) { left += dx; if (fromCenter) right -= dx; }
    if (h.includes("e")) { right += dx; if (fromCenter) left -= dx; }
    if (h.includes("n")) { top += dy; if (fromCenter) bottom -= dy; }
    if (h.includes("s")) { bottom += dy; if (fromCenter) top -= dy; }

    let width = Math.max(MIN_SIZE, right - left);
    let height = Math.max(MIN_SIZE, bottom - top);
    if (isCorner && (event.shiftKey !== drag.lockRatio)) {
      const ratio = b.w / Math.max(b.h, 0.0001);
      if (width / ratio > height) height = width / ratio; else width = height * ratio;
    }

    const cx = b.x + b.w / 2;
    const cy = b.y + b.h / 2;
    const next: Rect = {
      x: fromCenter ? cx - width / 2 : h.includes("w") ? b.x + b.w - width : h.includes("e") ? b.x : left,
      y: fromCenter ? cy - height / 2 : h.includes("n") ? b.y + b.h - height : h.includes("s") ? b.y : top,
      w: width,
      h: height
    };
    // Edge handles only change one axis.
    if (h === "n" || h === "s") { next.x = b.x; next.w = b.w; }
    if (h === "e" || h === "w") { next.y = b.y; next.h = b.h; }

    if (drag.rotation) {
      // The turned box keeps its pivot at the centre, so move the centre by the
      // local shift rotated back into slide space.
      const shiftX = next.x + next.w / 2 - cx;
      const shiftY = next.y + next.h / 2 - cy;
      const centreX = cx + shiftX * Math.cos(theta) - shiftY * Math.sin(theta);
      const centreY = cy + shiftX * Math.sin(theta) + shiftY * Math.cos(theta);
      next.x = centreX - next.w / 2;
      next.y = centreY - next.h / 2;
    }

    const sx = next.w / Math.max(b.w, 0.0001);
    const sy = next.h / Math.max(b.h, 0.0001);
    for (const item of drag.items) {
      const r = item.before;
      item.element.rect = {
        x: next.x + (r.x - b.x) * sx,
        y: next.y + (r.y - b.y) * sy,
        w: r.w * sx,
        h: r.h * sy
      };
    }
    this.showBadge(drag.slide, next, `${this.cm(drag.slide, next.w)} × ${this.cm(drag.slide, next.h)} cm`);
    this.queueRepaint(drag.slide);
  }

  private applyMarquee(drag: Extract<Drag, { kind: "marquee" }>, point: Point): void {
    const rect: Rect = {
      x: Math.min(drag.start.x, point.x),
      y: Math.min(drag.start.y, point.y),
      w: Math.abs(point.x - drag.start.x),
      h: Math.abs(point.y - drag.start.y)
    };
    const next = new Set(drag.base);
    for (const object of this.objects(drag.slide)) {
      const r = object.rect;
      const size = this.host.getSlideSize();
      if (size && r.w >= size.width * 0.95 && r.h >= size.height * 0.95) continue;
      if (r.x < rect.x + rect.w && r.x + r.w > rect.x && r.y < rect.y + rect.h && r.y + r.h > rect.y) {
        next.add(object.key);
      }
    }
    this.selection = next;
    const k = this.scale();
    Object.assign(this.marqueeBox.style, {
      display: "block",
      left: `${rect.x * k}px`,
      top: `${rect.y * k}px`,
      width: `${rect.w * k}px`,
      height: `${rect.h * k}px`
    });
    this.paintOverlay();
    this.emit();
  }

  // ------------------------------------------------------------ text editing

  private onDoubleClick(event: MouseEvent): void {
    if ((event.target as Element).closest?.(".edit-text")) return;
    const slide = this.host.getSlide();
    if (!slide) return;
    const box = this.box();
    const k = this.scale() || 1;
    const point = { x: (event.clientX - box.left) / k, y: (event.clientY - box.top) / k };
    let object = this.hitInScope(slide, point);

    // A double-click on a group steps into it and picks the child under the cursor.
    if (object?.isGroup) {
      event.preventDefault();
      this.entered = [...this.entered, object.source.shapeId];
      object = this.hit(slide, point);
      this.selection = object ? new Set([object.key]) : new Set();
      this.paintOverlay();
      this.emit();
      return;
    }
    const element = object?.elements.find((candidate): candidate is TextElement =>
      candidate.type === "text" && !!candidate.paragraphs?.length && !!candidate.source
      && !candidate.source.readOnlyText);
    if (!object || !element) return;
    event.preventDefault();
    this.selection = new Set([object.key]);
    this.startTextEdit(slide, element);
  }

  private startTextEdit(slide: Slide, element: TextElement): void {
    if (this.textSession) this.endTextEdit();
    const paragraphs = element.paragraphs || [];
    const k = this.scale() || 1;
    const body = element.body;

    const shell = this.div("edit-text");
    const anchor = body?.verticalAnchor === "middle" ? "center" : body?.verticalAnchor === "bottom" ? "flex-end" : "flex-start";
    Object.assign(shell.style, {
      left: `${element.rect.x * k}px`,
      top: `${element.rect.y * k}px`,
      width: `${element.rect.w * k}px`,
      height: `${element.rect.h * k}px`,
      justifyContent: anchor,
      transform: element.rotation ? `rotate(${element.rotation}deg)` : "",
      padding: `${(body?.marginTop ?? 4) * k}px ${(body?.marginRight ?? 8) * k}px ${(body?.marginBottom ?? 4) * k}px ${(body?.marginLeft ?? 8) * k}px`
    });

    const editable = this.div("edit-text-body");
    editable.contentEditable = "true";
    editable.spellcheck = false;
    editable.style.whiteSpace = body?.wrap === "none" ? "pre" : "pre-wrap";
    paragraphs.forEach(paragraph => editable.appendChild(this.paragraphNode(paragraph, k)));
    shell.appendChild(editable);

    editable.addEventListener("keydown", event => {
      event.stopPropagation();
      if ((event.ctrlKey || event.metaKey) && !event.altKey && ["b", "i"].includes(event.key.toLowerCase())) {
        event.preventDefault();
        this.toggleStyle(event.key.toLowerCase() === "b" ? "bold" : "italic");
        return;
      }
      if (event.key === "Escape") {
        event.preventDefault();
        this.endTextEdit();
      }
    });
    editable.addEventListener("paste", event => {
      event.preventDefault();
      const text = event.clipboardData?.getData("text/plain") ?? "";
      text.replace(/\r\n?/g, "\n").split("\n").forEach((line, index) => {
        if (index > 0) document.execCommand("insertParagraph");
        if (line) document.execCommand("insertText", false, line);
      });
    });
    editable.addEventListener("dragstart", event => event.preventDefault());

    this.layer.appendChild(shell);
    this.textSession = { slide, element, shell, body: editable, before: cloneParagraphs(paragraphs) };
    this.host.beginTextEdit(slide, element);
    this.selectBox.classList.add("is-editing");
    this.paintOverlay();
    this.emit();

    editable.focus();
    const range = document.createRange();
    range.selectNodeContents(editable);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }

  private paragraphNode(paragraph: TextParagraph, k: number): HTMLElement {
    const node = document.createElement("div");
    node.className = "edit-p";
    const firstRun = paragraph.runs[0];
    const style = paragraph.style;
    const lineHeight = style.lineSpacing
      ? style.lineSpacing.unit === "percent" ? 1.2 * style.lineSpacing.value : style.lineSpacing.value * k
      : 1.2;
    Object.assign(node.style, {
      textAlign: style.align,
      paddingLeft: `${Math.max(0, style.marginLeft) * k}px`,
      textIndent: `${style.indent * k}px`,
      lineHeight: typeof lineHeight === "number" && lineHeight < 4 ? String(lineHeight) : `${lineHeight}px`,
      marginTop: `${(style.spaceBefore || 0) * k}px`,
      marginBottom: `${(style.spaceAfter || 0) * k}px`,
      fontSize: `${(firstRun?.style.fontSize || 18) * k}px`
    });

    if (paragraph.bullet) {
      node.dataset.bullet = paragraph.bullet.char;
      node.style.setProperty("--bullet-color", paragraph.bullet.color);
      node.style.setProperty("--bullet-size", `${(paragraph.bullet.fontSize || firstRun?.style.fontSize || 18) * k}px`);
      node.style.setProperty("--bullet-width", `${Math.max(0, -style.indent) * k}px`);
    }

    let hasContent = false;
    for (const run of paragraph.runs) {
      if (run.content === "\n") {
        node.appendChild(document.createElement("br"));
        hasContent = true;
        continue;
      }
      if (!run.content) continue;
      const span = document.createElement("span");
      // East Asian first: the browser would otherwise pick a system CJK font
      // for Chinese glyphs and look nothing like the Rust-rendered text.
      const families = [run.style.eastAsianFontFamily, run.style.fontFamily]
        .filter((family): family is string => !!family && family !== "sans-serif")
        .map(family => `"${family}"`);
      Object.assign(span.style, {
        fontSize: `${run.style.fontSize * k}px`,
        color: run.style.color,
        fontWeight: run.style.bold ? "700" : "400",
        fontStyle: run.style.italic ? "italic" : "normal",
        fontFamily: [...families, "sans-serif"].join(", "),
        letterSpacing: run.style.letterSpacing ? `${run.style.letterSpacing * k}px` : ""
      });
      span.textContent = run.content;
      node.appendChild(span);
      hasContent = true;
    }
    if (!hasContent) node.appendChild(document.createElement("br"));
    return node;
  }

  private readEditedParagraphs(body: HTMLElement): string {
    const lines: string[] = [];
    const readBlock = (block: Node): string => {
      let text = "";
      const walk = (node: Node) => {
        if (node.nodeType === Node.TEXT_NODE) {
          text += node.textContent || "";
        } else if (node.nodeName === "BR") {
          text += LINE_BREAK;
        } else {
          node.childNodes.forEach(walk);
        }
      };
      walk(block);
      // A lone <br> only keeps an empty line open in the browser.
      if (text === LINE_BREAK) text = "";
      return text.replace(/\u00a0/g, " ");
    };

    let inline = "";
    let hasInline = false;
    const flushInline = () => {
      if (hasInline) lines.push(inline);
      inline = "";
      hasInline = false;
    };
    body.childNodes.forEach(node => {
      if (node.nodeName === "DIV" || node.nodeName === "P") {
        flushInline();
        lines.push(readBlock(node));
      } else {
        inline += readBlock(node);
        hasInline = true;
      }
    });
    flushInline();
    return (lines.length ? lines : [""]).join("\n");
  }

  /** Turn typing done so far into an undoable command; the session keeps going. */
  private flushText(session: NonNullable<SlideEditor["textSession"]>): void {
    const afterText = this.readEditedParagraphs(session.body);
    const ops = diffText(paragraphsText(session.before), afterText);
    if (ops.length === 0 || !session.element.source) return;
    const after = cloneParagraphs(session.before);
    const model = vdomTextModel(after, this.emuPerPx(session.slide));
    for (const op of ops) applyTextOp(model, op);
    normalizeParagraphs(after, session.element.style);
    const command = new TextCommand(session.slide, session.element, session.before, after, ops, this, this.host);
    command.apply();
    this.history.record(command);
    session.before = cloneParagraphs(after);
  }

  private endTextEdit(): void {
    const session = this.textSession;
    if (!session) return;
    this.flushText(session);
    this.textSession = null;
    session.shell.remove();
    this.selectBox.classList.remove("is-editing");
    this.host.endTextEdit(session.slide);
    this.paintOverlay();
    this.emit();
  }

  private cancelTextSession(): void {
    const session = this.textSession;
    if (!session) return;
    this.textSession = null;
    session.shell.remove();
    this.selectBox.classList.remove("is-editing");
    this.host.endTextEdit(session.slide);
  }

  // ------------------------------------------------------------- formatting

  /** Flat text offset of a DOM position inside the text editor. */
  private flatIndex(body: HTMLElement, node: Node, offset: number): number {
    if (node === body) {
      const children = Array.from(body.children);
      const holder = document.createElement("div");
      children.slice(0, offset).forEach(child => holder.appendChild(child.cloneNode(true)));
      const length = this.readEditedParagraphs(holder).length;
      return offset > 0 && offset < children.length ? length + 1 : length;
    }
    const range = document.createRange();
    range.setStart(body, 0);
    range.setEnd(node, offset);
    const holder = document.createElement("div");
    holder.appendChild(range.cloneContents());
    return this.readEditedParagraphs(holder).length;
  }

  /** Put the DOM selection on the flat offsets [start, end). */
  private selectFlat(body: HTMLElement, start: number, end: number): void {
    const locate = (position: number): [Node, number] => {
      let cursor = 0;
      const paragraphs = Array.from(body.children);
      for (let index = 0; index < paragraphs.length; index++) {
        const paragraph = paragraphs[index];
        const text = this.readEditedParagraphs(this.wrap(paragraph));
        if (position <= cursor + text.length || index === paragraphs.length - 1) {
          let rest = Math.max(0, position - cursor);
          const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            if (node.nodeType === Node.TEXT_NODE) {
              const length = node.textContent?.length || 0;
              if (rest <= length) return [node, rest];
              rest -= length;
            } else if (node.nodeName === "BR" && text.length > 0) {
              if (rest === 0) return [node.parentNode!, Array.from(node.parentNode!.childNodes).indexOf(node as ChildNode)];
              rest -= 1;
            }
          }
          return [paragraph, paragraph.childNodes.length];
        }
        cursor += text.length + 1;
      }
      return [body, 0];
    };
    const [startNode, startOffset] = locate(start);
    const [endNode, endOffset] = locate(end);
    const range = document.createRange();
    range.setStart(startNode, startOffset);
    range.setEnd(endNode, endOffset);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }

  private wrap(node: Element): HTMLElement {
    const holder = document.createElement("div");
    holder.appendChild(node.cloneNode(true));
    return holder;
  }

  private textElementsOf(slide: Slide): TextElement[] {
    return this.selectedObjects(slide)
      .flatMap(object => object.elements)
      .filter((element): element is TextElement =>
        element.type === "text" && !!element.source && !element.source.readOnlyText
        && (element.paragraphs || []).some(paragraph =>
          paragraph.runs.some(run => run.content !== "")));
  }

  private formatOfElements(slide: Slide, elements: TextElement[]): FormatState {
    const px = this.emuPerPx(slide) / 12700;
    const runs = elements.flatMap(element => element.paragraphs || [])
      .flatMap(paragraph => paragraph.runs).filter(run => run.content !== "");
    return this.mergeFormat(runs.map(run => ({
      bold: run.style.bold,
      italic: !!run.style.italic,
      color: run.style.color,
      fontSizePt: run.style.fontSize * px
    })));
  }

  private formatOfDomSelection(range: Range): FormatState {
    const body = this.textSession!.body;
    const k = this.scale() || 1;
    const slide = this.textSession!.slide;
    const walker = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
    const samples: FormatState[] = [];
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (!(node.textContent || "").length) continue;
      // Skip nodes that only touch the selection edge.
      const whole = document.createRange();
      whole.selectNodeContents(node);
      if (
        whole.compareBoundaryPoints(Range.START_TO_END, range) <= 0
        || whole.compareBoundaryPoints(Range.END_TO_START, range) >= 0
      ) continue;
      const style = getComputedStyle((node.parentElement || body));
      samples.push({
        bold: Number(style.fontWeight) >= 600,
        italic: style.fontStyle === "italic",
        color: this.cssToHex(style.color),
        fontSizePt: parseFloat(style.fontSize) / k * (this.emuPerPx(slide) / 12700)
      });
    }
    return this.mergeFormat(samples);
  }

  private mergeFormat(samples: FormatState[]): FormatState {
    const merged: FormatState = { bold: null, italic: null, color: null, fontSizePt: null };
    if (samples.length === 0) return merged;
    const same = <T,>(pick: (sample: FormatState) => T, equal: (a: T, b: T) => boolean) =>
      samples.every(sample => equal(pick(sample), pick(samples[0]))) ? pick(samples[0]) : null;
    merged.bold = same(sample => sample.bold, (a, b) => a === b);
    merged.italic = same(sample => sample.italic, (a, b) => a === b);
    merged.color = same(sample => sample.color, (a, b) => a === b);
    merged.fontSizePt = same(sample => sample.fontSizePt, (a, b) => Math.abs((a ?? 0) - (b ?? 0)) < 0.3);
    return merged;
  }

  private cssToHex(color: string): string | null {
    const match = color.match(/\d+(\.\d+)?/g);
    if (!match || match.length < 3) return null;
    return `#${match.slice(0, 3).map(value => Math.round(Number(value)).toString(16).padStart(2, "0")).join("")}`;
  }

  /** Show, move or hide the floating format toolbar for the current context. */
  private updateToolbar(): void {
    const slide = this.host.getSlide();
    const box = this.box();
    const k = this.scale() || 1;
    const bounds = { w: box.width, h: box.height };
    if (!slide || this.drag) {
      this.toolbar.hide();
      return;
    }

    if (this.textSession) {
      const selection = window.getSelection();
      const range = selection && selection.rangeCount ? selection.getRangeAt(0) : null;
      if (!range || range.collapsed || !this.textSession.body.contains(range.commonAncestorContainer)) {
        this.toolbar.hide();
        return;
      }
      const rect = range.getBoundingClientRect();
      this.toolbar.show(
        { x: rect.left - box.left, y: rect.top - box.top, w: rect.width, h: rect.height },
        this.formatOfDomSelection(range),
        bounds
      );
      return;
    }

    const elements = this.textElementsOf(slide);
    if (elements.length === 0) {
      this.toolbar.hide();
      return;
    }
    const rect = union(this.selectedObjects(slide).map(object => object.rect));
    // Leave room for the rotation handle that sits above a single selection.
    const lift = this.selection.size === 1 ? 30 : 0;
    this.toolbar.show(
      { x: rect.x * k, y: rect.y * k - lift, w: rect.w * k, h: rect.h * k + lift },
      this.formatOfElements(slide, elements),
      bounds
    );
  }

  /** Ctrl+B / Ctrl+I: flip the style the same way the toolbar button would. */
  private toggleStyle(kind: "bold" | "italic"): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    let state: FormatState | null = null;
    if (this.textSession) {
      const selection = window.getSelection();
      const range = selection && selection.rangeCount ? selection.getRangeAt(0) : null;
      if (range && !range.collapsed && this.textSession.body.contains(range.commonAncestorContainer)) {
        state = this.formatOfDomSelection(range);
      }
    } else {
      const elements = this.textElementsOf(slide);
      if (elements.length > 0) state = this.formatOfElements(slide, elements);
    }
    if (!state) return;
    this.applyFormat({ [kind]: state[kind] !== true });
  }

  private applyFormat(format: TextFormat): void {
    this.applyRangeOp({ format }, false);
  }

  /**
   * Run a restyling op over the text selection while editing, or over all of
   * the text of the selected shapes otherwise.
   */
  private applyRangeOp(extra: Pick<TextOp, "format" | "paragraph">, allowCollapsed: boolean): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    const session = this.textSession;

    if (session) {
      const selection = window.getSelection();
      const range = selection && selection.rangeCount ? selection.getRangeAt(0) : null;
      if (!range || !session.body.contains(range.commonAncestorContainer)) return;
      if (range.collapsed && !allowCollapsed) return;
      const start = this.flatIndex(session.body, range.startContainer, range.startOffset);
      const end = this.flatIndex(session.body, range.endContainer, range.endOffset);
      this.flushText(session);

      const op: TextOp = { start, end, insert: "", ...extra };
      const after = cloneParagraphs(session.before);
      applyTextOp(vdomTextModel(after, this.emuPerPx(slide)), op);
      const command = new TextCommand(slide, session.element, session.before, after, [op], this, this.host);
      command.apply();
      this.history.record(command);
      session.before = cloneParagraphs(after);

      const k = this.scale() || 1;
      session.body.replaceChildren(...after.map(paragraph => this.paragraphNode(paragraph, k)));
      session.body.focus();
      this.selectFlat(session.body, start, end);
      this.updateToolbar();
      return;
    }

    const targets = extra.paragraph ? this.selectedTexts(slide) : this.textElementsOf(slide);
    const commands = targets.map(element => {
      const before = cloneParagraphs(element.paragraphs || []);
      const after = cloneParagraphs(before);
      const op: TextOp = { start: 0, end: paragraphsText(before).length, insert: "", ...extra };
      applyTextOp(vdomTextModel(after, this.emuPerPx(slide)), op);
      return new TextCommand(slide, element, before, after, [op], this, this.host);
    });
    if (commands.length === 0) return;
    const command = commands.length === 1 ? commands[0] : new CompositeCommand("设置格式", commands);
    command.apply();
    this.history.record(command);
  }

  // ---------------------------------------------------------------- painting

  private queueRepaint(slide: Slide): void {
    if (this.repaintQueued) return;
    this.repaintQueued = true;
    requestAnimationFrame(() => {
      this.repaintQueued = false;
      this.host.invalidate(slide);
    });
  }

  private setHover(key: string | null): void {
    if (this.hoverKey === key) return;
    this.hoverKey = key;
    this.paintOverlay();
  }

  private place(node: HTMLElement, rect: Rect, rotation = 0): void {
    const k = this.scale();
    Object.assign(node.style, {
      display: "block",
      left: `${rect.x * k}px`,
      top: `${rect.y * k}px`,
      width: `${rect.w * k}px`,
      height: `${rect.h * k}px`,
      transform: rotation ? `rotate(${rotation}deg)` : ""
    });
  }

  private paintOverlay(): void {
    const slide = this.host.getSlide();
    if (!slide) {
      this.hoverBox.style.display = "none";
      this.selectBox.style.display = "none";
      return;
    }
    const objects = this.objects(slide);
    const selected = objects.filter(object => this.selection.has(object.key));
    this.paintGroupFrame(slide);
    this.paintPlaceholderHints(slide);

    const hovered = this.hoverKey && !this.selection.has(this.hoverKey)
      ? objects.find(object => object.key === this.hoverKey) : undefined;
    if (hovered && !this.drag) this.place(this.hoverBox, hovered.rect, hovered.rotation);
    else this.hoverBox.style.display = "none";

    if (selected.length === 0) {
      this.selectBox.style.display = "none";
    } else {
      const rect = union(selected.map(object => object.rect));
      const rotation = selected.length === 1 ? selected[0].rotation : 0;
      this.place(this.selectBox, rect, rotation);
      this.selectBox.classList.toggle("is-multi", selected.length > 1);
      const canResize = selected.length === 1 || selected.every(object => object.rotation === 0);
      this.selectBox.classList.toggle("no-resize", !canResize);
      // Groups have no rotation of their own to write back yet.
      this.selectBox.classList.toggle("no-rotate", !(selected.length === 1 && !selected[0].isGroup));
      // Small boxes drop the edge handles so corners stay reachable.
      const k = this.scale();
      this.selectBox.classList.toggle("is-compact", rect.w * k < 44 || rect.h * k < 44);
    }

    this.paintGuides();
    this.updateToolbar();
  }

  /** Empty placeholders show their prompt, as in PowerPoint's editing view (never exported). */
  private paintPlaceholderHints(slide: Slide): void {
    const k = this.scale() || 1;
    const editing = this.textSession?.element;
    const hints = slide.elements
      .filter((element): element is TextElement =>
        element.type === "text" && !!element.placeholder && !!element.source && element !== editing
        && element.content.trim() === "" && element.rect.w > 0 && element.rect.h > 0)
      .map(element => {
        const hint = this.div("edit-placeholder");
        hint.textContent = placeholderPrompt(element.placeholder!.type);
        Object.assign(hint.style, {
          left: `${element.rect.x * k}px`,
          top: `${element.rect.y * k}px`,
          width: `${element.rect.w * k}px`,
          height: `${element.rect.h * k}px`,
          fontSize: `${Math.min(24, Math.max(11, element.rect.h * k * 0.3))}px`,
          transform: element.rotation ? `rotate(${element.rotation}deg)` : ""
        });
        return hint;
      });
    this.hintLayer.replaceChildren(...hints);
  }

  /** Faint outline around the group the user has stepped into. */
  private paintGroupFrame(slide: Slide): void {
    if (this.entered.length === 0) {
      this.groupFrame.style.display = "none";
      return;
    }
    const members = slide.elements.filter(element =>
      this.entered.every((id, index) => element.source?.groups?.[index]?.shapeId === id));
    if (members.length === 0) {
      this.groupFrame.style.display = "none";
      return;
    }
    this.place(this.groupFrame, union(members.map(member => member.rect)));
  }

  private paintGuides(): void {
    const k = this.scale();
    this.guideLayer.replaceChildren(...this.guides.map(guide => {
      const line = this.div(`edit-guide edit-guide-${guide.axis}`);
      if (guide.axis === "x") {
        Object.assign(line.style, { left: `${guide.pos * k}px`, top: `${guide.from * k}px`, height: `${(guide.to - guide.from) * k}px` });
      } else {
        Object.assign(line.style, { top: `${guide.pos * k}px`, left: `${guide.from * k}px`, width: `${(guide.to - guide.from) * k}px` });
      }
      return line;
    }));
  }

  private showBadge(slide: Slide, rect: Rect, text: string): void {
    const k = this.scale();
    this.badge.textContent = text;
    Object.assign(this.badge.style, {
      display: "block",
      left: `${(rect.x + rect.w / 2) * k}px`,
      top: `${(rect.y + rect.h) * k + 14}px`
    });
    this.paintOverlay();
    void slide;
  }

  private emit(): void {
    const state = this.getState();
    this.listeners.forEach(listener => listener(state));
  }

  private div(className: string): HTMLDivElement {
    const node = document.createElement("div");
    node.className = className;
    return node;
  }
}
