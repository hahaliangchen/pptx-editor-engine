import { ElementSource, Rect, Slide, SlideElement } from "../pptx-virtual-dom";
import { Command, History } from "./history";

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
}

export interface EditorState {
  canUndo: boolean;
  canRedo: boolean;
  selectionCount: number;
}

interface SceneObject {
  key: string;
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
}

type Drag =
  | { kind: "move"; slide: Slide; start: Point; items: DragItem[]; bbox: Rect; keys: Set<string>; active: boolean; collapseTo: string | null }
  | { kind: "resize"; slide: Slide; start: Point; items: DragItem[]; bbox: Rect; handle: HandleId; lockRatio: boolean }
  | { kind: "marquee"; slide: Slide; start: Point; base: Set<string> };

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

const elementKey = (element: SlideElement): string | null =>
  element.source ? `${element.source.part}|${element.source.shapeId}` : null;

const rotationOf = (element: SlideElement): number =>
  element.type === "text" ? 0 : element.rotation || 0;

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
    private items: Array<{ element: SlideElement; before: Rect; after: Rect }>,
    private editor: SlideEditor,
    mergeKey?: string
  ) {
    this.label = label;
    this.mergeKey = mergeKey;
  }

  apply(): void {
    for (const { element, after } of this.items) element.rect = cloneRect(after);
    this.editor.afterCommand(this.slide, this.items.map(item => item.element));
  }

  revert(): void {
    for (const { element, before } of this.items) element.rect = cloneRect(before);
    this.editor.afterCommand(this.slide, this.items.map(item => item.element));
  }

  merge(next: Command): void {
    if (!(next instanceof RectsCommand)) return;
    for (const item of next.items) {
      const mine = this.items.find(candidate => candidate.element === item.element);
      if (mine) mine.after = item.after;
    }
  }
}

class DeleteCommand implements Command {
  public readonly label = "删除";
  private removed: Array<{ element: SlideElement; index: number }>;

  constructor(private slide: Slide, elements: SlideElement[], private editor: SlideEditor) {
    this.removed = elements
      .map(element => ({ element, index: slide.elements.indexOf(element) }))
      .filter(item => item.index >= 0)
      .sort((a, b) => a.index - b.index);
  }

  apply(): void {
    const gone = new Set(this.removed.map(item => item.element));
    this.slide.elements = this.slide.elements.filter(element => !gone.has(element));
    this.editor.afterCommand(this.slide, this.removed.map(item => item.element));
  }

  revert(): void {
    for (const { element, index } of this.removed) this.slide.elements.splice(index, 0, element);
    this.editor.afterCommand(this.slide, this.removed.map(item => item.element));
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

  constructor(private host: EditorHost) {
    const container = host.container;
    this.layer = this.div("edit-layer");
    this.guideLayer = this.div("edit-guides");
    this.hoverBox = this.div("edit-hover");
    this.selectBox = this.div("edit-selection");
    this.marqueeBox = this.div("edit-marquee");
    this.badge = this.div("edit-badge");
    for (const id of HANDLES) {
      const handle = this.div(`edit-handle edit-handle-${id}`);
      handle.dataset.handle = id;
      this.selectBox.appendChild(handle);
      this.handleNodes.set(id, handle);
    }
    this.layer.append(this.guideLayer, this.hoverBox, this.selectBox, this.marqueeBox, this.badge);
    container.appendChild(this.layer);

    container.addEventListener("pointerdown", e => this.onPointerDown(e));
    container.addEventListener("pointermove", e => this.onPointerMove(e));
    container.addEventListener("pointerup", e => this.onPointerUp(e));
    container.addEventListener("pointercancel", e => this.onPointerUp(e));
    container.addEventListener("pointerleave", () => this.setHover(null));
    document.addEventListener("keydown", e => this.onKeyDown(e));
    this.history.subscribe(() => this.emit());
  }

  // ---------------------------------------------------------------- public

  public subscribe(listener: (state: EditorState) => void): void {
    this.listeners.add(listener);
  }

  public getState(): EditorState {
    return {
      canUndo: this.history.canUndo,
      canRedo: this.history.canRedo,
      selectionCount: this.selection.size
    };
  }

  /** Call after a (re)paint or resize so the overlay follows the slide. */
  public layout(): void {
    const slide = this.host.getSlide();
    if (slide && this.selectionSlideId !== slide.id) {
      this.selection.clear();
      this.selectionSlideId = slide.id;
      this.emit();
    }
    this.paintOverlay();
  }

  public reset(): void {
    this.selection.clear();
    this.selectionSlideId = null;
    this.drag = null;
    this.guides = [];
    this.history.clear();
    this.paintOverlay();
    this.emit();
  }

  public undo(): void {
    this.history.undo();
  }

  public redo(): void {
    this.history.redo();
  }

  public deleteSelection(): void {
    const slide = this.host.getSlide();
    if (!slide) return;
    const objects = this.selectedObjects(slide);
    if (objects.length === 0) return;
    const command = new DeleteCommand(slide, objects.flatMap(object => object.elements), this);
    command.apply();
    this.history.record(command);
    this.selection.clear();
    this.host.onCommit(slide);
    this.paintOverlay();
    this.emit();
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
  public afterCommand(slide: Slide, elements: SlideElement[]): void {
    for (const element of elements) this.host.markDirty(element.source);
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

  private objects(slide: Slide): SceneObject[] {
    const byKey = new Map<string, SceneObject>();
    slide.elements.forEach((element, index) => {
      const key = elementKey(element);
      if (!key) return;
      const existing = byKey.get(key);
      if (existing) {
        existing.elements.push(element);
        existing.order = index;
      } else {
        byKey.set(key, { key, elements: [element], rect: element.rect, rotation: rotationOf(element), order: index });
      }
    });
    const result = [...byKey.values()];
    for (const object of result) {
      object.rect = union(object.elements.map(element => element.rect));
      object.rotation = object.elements.reduce((acc, element) => acc || rotationOf(element), 0);
    }
    return result.sort((a, b) => a.order - b.order);
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
    const slide = this.host.getSlide();
    if (!slide) return;
    event.preventDefault();
    const point = this.point(event);
    const handle = (event.target as HTMLElement).closest<HTMLElement>("[data-handle]")?.dataset.handle as HandleId | undefined;
    const selected = this.selectedObjects(slide);

    if (handle && selected.length > 0) {
      const items = selected.flatMap(object => object.elements)
        .map(element => ({ element, before: cloneRect(element.rect) }));
      this.drag = {
        kind: "resize",
        slide,
        start: point,
        items,
        bbox: union(items.map(item => item.before)),
        handle,
        lockRatio: items.every(item => item.element.type === "image")
      };
      this.host.beginDrag(slide, items.map(item => item.element));
    } else {
      const object = this.hit(slide, point);
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
        const items = moving.flatMap(o => o.elements).map(element => ({ element, before: cloneRect(element.rect) }));
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

    if (drag.kind === "marquee") {
      this.marqueeBox.style.display = "none";
    } else {
      const items = drag.items
        .filter(item => item.element.rect.x !== item.before.x || item.element.rect.y !== item.before.y
          || item.element.rect.w !== item.before.w || item.element.rect.h !== item.before.h)
        .map(item => ({ element: item.element, before: item.before, after: cloneRect(item.element.rect) }));
      if (items.length > 0) {
        this.history.record(new RectsCommand(
          drag.kind === "move" ? "移动" : "调整大小", drag.slide, items, this
        ));
        items.forEach(item => this.host.markDirty(item.element.source));
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
    } else if (key === "escape") {
      this.clearSelection();
    } else if (key.startsWith("arrow") && this.selection.size > 0 && !mod) {
      event.preventDefault();
      const step = event.shiftKey ? 10 : 1;
      const dx = key === "arrowleft" ? -step : key === "arrowright" ? step : 0;
      const dy = key === "arrowup" ? -step : key === "arrowdown" ? step : 0;
      this.nudge(slide, dx, dy);
    }
  }

  private nudge(slide: Slide, dx: number, dy: number): void {
    const elements = this.selectedObjects(slide).flatMap(object => object.elements);
    if (elements.length === 0) return;
    const items = elements.map(element => ({
      element,
      before: cloneRect(element.rect),
      after: { ...element.rect, x: element.rect.x + dx, y: element.rect.y + dy }
    }));
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

  private applyResize(drag: Extract<Drag, { kind: "resize" }>, point: Point, event: PointerEvent): void {
    const dx = point.x - drag.start.x;
    const dy = point.y - drag.start.y;
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
      const canResize = rotation === 0 && selected.every(object => object.rotation === 0);
      this.selectBox.classList.toggle("no-resize", !canResize);
      // Small boxes drop the edge handles so corners stay reachable.
      const k = this.scale();
      this.selectBox.classList.toggle("is-compact", rect.w * k < 44 || rect.h * k < 44);
    }

    this.paintGuides();
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
