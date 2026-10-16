import { CreatedShape } from "../editor/slide-editor";
import { NewShapeKind } from "../editor/new-shapes";
import { ShapeStyleOp } from "../editor/shape-ops";
import { TextOp } from "../editor/text-ops";
import { paragraphsText } from "../editor/text-model-vdom";
import { OrderOp } from "../pptx-exporter";
import { PptVirtualDocument, Rect, TextElement } from "../pptx-virtual-dom";
import { ApiError } from "./errors";
import { Outline, OutlineOptions, buildOutline, emuPerPxOf } from "./outline";
import { ResolveContext, Target, resolveNode, resolveSlide } from "./resolve";
import { SceneNode, SlideScene, buildScene } from "./scene";
import { CommandBatch, parseBatch } from "./schema";

/** What the API needs from the editor; SlideEditor provides all of it. */
export interface EditorPort {
  history: { begin(label: string): void; commit(): number; rollback(): void };
  selectShapes(refs: Array<{ part: string; id: string }>): boolean;
  selectionRefs(): Array<{ part: string; id: string }>;
  clearSelection(): void;
  applyTextOps(element: TextElement, ops: TextOp[]): void;
  applyShapeStyle(op: ShapeStyleOp): void;
  setBounds(change: Partial<{ x: number; y: number; w: number; h: number }>, keepRatio?: boolean): void;
  setRotationDegrees(degrees: number): void;
  alignToSlide(edge: "left" | "hcenter" | "right" | "top" | "vcenter" | "bottom"): void;
  arrange(mode: OrderOp): void;
  duplicate(): Promise<void>;
  copy(): Promise<boolean>;
  paste(): Promise<void>;
  deleteSelection(): void;
  createShape(kind: NewShapeKind, rect: Rect, flip?: { flipH: boolean; flipV: boolean }): CreatedShape | null;
  insertImage(image: DecodedImage & { name: string }): Promise<CreatedShape | null>;
}

export interface DecodedImage {
  bytes: Uint8Array;
  extension: string;
  contentType: string;
  url: string;
  /** Logical slide units, as the editor places pictures. */
  width: number;
  height: number;
}

/** What the API needs from the page around the editor: the deck and its slides. */
export interface DeckPort {
  document(): PptVirtualDocument | null;
  /** 0-based index of the slide on screen. */
  currentIndex(): number;
  showSlide(index: number): Promise<void>;
  addSlide(after: number, layout?: string): Promise<void>;
  duplicateSlide(index: number): Promise<void>;
  deleteSlide(index: number): void;
  moveSlide(from: number, to: number): void;
  layouts(): Promise<string[]>;
  decodeImage(dataUrl: string): Promise<DecodedImage>;
}

export interface CommandResult {
  ok: boolean;
  created?: { id: string; name: string; slide: number };
  slide?: number;
  error?: ReturnType<ApiError["toJSON"]>;
}

export interface ApplyResult {
  ok: boolean;
  /** Index of the first command that failed (atomic batches stop there). */
  failedAt?: number;
  error?: ReturnType<ApiError["toJSON"]>;
  results: CommandResult[];
}

type Cmd = Record<string, any> & { op: string };

const CM = 360000;
const DASH = { solid: "solid", dash: "dash", dot: "sysDot" } as const;

/**
 * The JSON edit interface: read the deck (`outline`), change it with a batch of
 * commands (`apply`). Every command goes through the same editor operations the
 * mouse and keyboard use, so history, undo and export behave identically.
 */
export class EditApi {
  constructor(private editor: EditorPort, private deck: DeckPort) {}

  public async outline(options: OutlineOptions = {}): Promise<Outline> {
    const doc = this.requireDocument();
    return buildOutline(doc, { ...options, layouts: options.layouts ?? (await this.deck.layouts()) });
  }

  public async apply(input: unknown): Promise<ApplyResult> {
    let batch: CommandBatch;
    try {
      batch = parseBatch(input);
    } catch (error) {
      return { ok: false, error: this.asError(error).toJSON(), results: [] };
    }

    const doc = this.requireDocument();
    const startIndex = this.deck.currentIndex();
    const aliases = new Map<string, { part: string; id: string }>();
    const results: CommandResult[] = [];
    this.editor.history.begin(batch.label || "AI 修改");

    for (let index = 0; index < batch.commands.length; index++) {
      try {
        results.push({ ok: true, ...(await this.run(batch.commands[index] as Cmd, `commands[${index}]`, aliases)) });
      } catch (error) {
        const failure = this.asError(error, `commands[${index}]`);
        if (batch.atomic) {
          this.editor.history.rollback();
          this.editor.clearSelection();
          await this.deck.showSlide(Math.min(startIndex, doc.slides.length - 1));
          return { ok: false, failedAt: index, error: failure.toJSON(), results };
        }
        results.push({ ok: false, error: failure.toJSON() });
      }
    }

    this.editor.history.commit();
    await this.deck.showSlide(Math.min(startIndex, this.requireDocument().slides.length - 1));
    const failed = results.findIndex(result => !result.ok);
    return failed >= 0
      ? { ok: false, failedAt: failed, error: results[failed].error, results }
      : { ok: true, results };
  }

  // ------------------------------------------------------------------ plumbing

  private requireDocument(): PptVirtualDocument {
    const doc = this.deck.document();
    if (!doc) throw new ApiError("FAILED", "还没有打开任何演示文稿");
    return doc;
  }

  private asError(error: unknown, path?: string): ApiError {
    if (error instanceof ApiError) return error;
    return new ApiError("FAILED", error instanceof Error ? error.message : String(error), path);
  }

  private context(aliases: Map<string, { part: string; id: string }>): ResolveContext {
    return {
      scenes: buildScene(this.requireDocument().slides),
      aliases,
      selection: () => this.editor.selectionRefs()
    };
  }

  private emuPerPx(scene: SlideScene): number {
    return emuPerPxOf(scene.slide, this.requireDocument().size.width);
  }

  private toPx(scene: SlideScene, cm: number): number {
    return (cm * CM) / this.emuPerPx(scene);
  }

  private rectPx(scene: SlideScene, rect: { x: number; y: number; w: number; h: number }): Rect {
    return { x: this.toPx(scene, rect.x), y: this.toPx(scene, rect.y), w: this.toPx(scene, rect.w), h: this.toPx(scene, rect.h) };
  }

  /** Put the shape's slide on screen and make the shape the selection. */
  private async focus(scene: SlideScene, node: SceneNode, path: string): Promise<void> {
    await this.deck.showSlide(scene.index - 1);
    if (!this.editor.selectShapes([{ part: node.part, id: node.id }])) {
      throw new ApiError("SHAPE_NOT_FOUND", `第 ${scene.index} 页上找不到 id 为 ${node.id} 的形状`, path);
    }
  }

  private textElementOf(node: SceneNode, path: string): TextElement {
    const element = node.elements.find((candidate): candidate is TextElement =>
      candidate.type === "text" && !candidate.source?.readOnlyText);
    if (node.kind === "group" || node.kind === "table" || !element) {
      const why = node.kind === "group" ? "组合本身没有文字，请指向组合里的具体形状"
        : node.kind === "table" ? "暂不支持修改表格里的文字"
        : "这个对象没有文字框";
      throw new ApiError("NOT_APPLICABLE", `${node.name}：${why}`, path);
    }
    return element;
  }

  private requirePaintable(node: SceneNode, path: string): void {
    if (node.kind === "image" || node.kind === "table") {
      throw new ApiError("NOT_APPLICABLE", `${node.name} 是${node.kind === "image" ? "图片" : "表格"}，没有填充和边框可以设置`, path);
    }
  }

  private register(aliases: Map<string, { part: string; id: string }>, alias: unknown, created: { part: string; id: string } | null) {
    if (typeof alias === "string" && created) aliases.set(alias, { part: created.part, id: created.id });
  }

  private createdResult(scene: SlideScene | null, created: { id: string; name: string; part: string } | null): CommandResult {
    if (!created) throw new ApiError("FAILED", "创建失败");
    const index = scene?.index ?? this.deck.currentIndex() + 1;
    return { ok: true, created: { id: created.id, name: created.name, slide: index } };
  }

  // ------------------------------------------------------------------ commands

  private async run(command: Cmd, path: string, aliases: Map<string, { part: string; id: string }>): Promise<Partial<CommandResult>> {
    const context = this.context(aliases);
    const at = (field: string) => `${path}.${field}`;
    const node = async (target: Target, field = "target") => {
      const found = resolveNode(this.context(aliases), target, at(field));
      await this.focus(found.scene, found.node, at(field));
      return found;
    };

    switch (command.op) {
      // ---- slides
      case "addSlide": {
        const total = context.scenes.length;
        const after = command.after ?? this.deck.currentIndex() + 1;
        if (after > total) throw new ApiError("SLIDE_NOT_FOUND", `没有第 ${after} 页，当前共 ${total} 页`, at("after"));
        if (command.layout) {
          const layouts = await this.deck.layouts();
          if (!layouts.includes(command.layout)) {
            throw new ApiError("INVALID_SCHEMA", `没有名为 "${command.layout}" 的版式，可用的有：${layouts.join("、")}`, at("layout"));
          }
        }
        // "after 0" means before the first slide: add after it, then move it up.
        await this.deck.addSlide(Math.max(0, after - 1), command.layout);
        if (after === 0) this.deck.moveSlide(1, 0);
        return { slide: after + 1 };
      }
      case "duplicateSlide": {
        const scene = resolveSlide(context.scenes, command.slide, at("slide"));
        await this.deck.duplicateSlide(scene.index - 1);
        return { slide: scene.index + 1 };
      }
      case "deleteSlide": {
        const scene = resolveSlide(context.scenes, command.slide, at("slide"));
        if (context.scenes.length <= 1) throw new ApiError("NOT_APPLICABLE", "至少需要保留一页幻灯片", at("slide"));
        this.deck.deleteSlide(scene.index - 1);
        return {};
      }
      case "moveSlide": {
        const scene = resolveSlide(context.scenes, command.slide, at("slide"));
        if (command.to > context.scenes.length) {
          throw new ApiError("SLIDE_NOT_FOUND", `目标位置 ${command.to} 超出页数，当前共 ${context.scenes.length} 页`, at("to"));
        }
        this.deck.moveSlide(scene.index - 1, command.to - 1);
        return { slide: command.to };
      }

      // ---- new content
      case "addText": {
        const scene = resolveSlide(context.scenes, command.slide, at("slide"));
        await this.deck.showSlide(scene.index - 1);
        const created = this.editor.createShape("textbox", this.rectPx(scene, command.rect));
        if (!created?.text) throw new ApiError("FAILED", "无法创建文本框", path);
        const paragraphs: Array<{ text: string; align?: string }> = command.paragraphs ?? [{ text: command.text }];
        const text = paragraphs.map(paragraph => paragraph.text).join("\n");
        this.setTextOn(created.text, text);
        this.formatParagraphs(created.text, paragraphs, command.style);
        this.register(aliases, command.as, created);
        return this.createdResult(scene, created);
      }
      case "addShape": {
        const scene = resolveSlide(context.scenes, command.slide, at("slide"));
        await this.deck.showSlide(scene.index - 1);
        const created = this.editor.createShape(
          command.shape,
          this.rectPx(scene, command.rect),
          { flipH: !!command.flipH, flipV: !!command.flipV }
        );
        if (!created) throw new ApiError("FAILED", "无法创建形状", path);
        this.editor.selectShapes([{ part: created.part, id: created.id }]);
        if (command.fill !== undefined) this.editor.applyShapeStyle({ kind: "fill", color: command.fill });
        if (command.line) this.editor.applyShapeStyle(this.lineOp(command.line));
        if (command.text) {
          if (!created.text) throw new ApiError("NOT_APPLICABLE", "直线和箭头没有文字", at("text"));
          this.setTextOn(created.text, command.text);
        }
        this.register(aliases, command.as, created);
        return this.createdResult(scene, created);
      }
      case "addImage": {
        const scene = resolveSlide(context.scenes, command.slide, at("slide"));
        await this.deck.showSlide(scene.index - 1);
        let image: DecodedImage;
        try {
          image = await this.deck.decodeImage(command.src);
        } catch (error) {
          throw new ApiError("INVALID_SCHEMA", `图片无法读取：${error instanceof Error ? error.message : error}`, at("src"));
        }
        const created = await this.editor.insertImage({ ...image, name: command.name || "图片" });
        if (!created) throw new ApiError("FAILED", "无法插入图片", path);
        if (command.rect) {
          this.editor.selectShapes([{ part: created.part, id: created.id }]);
          this.editor.setBounds(command.rect);
        }
        this.register(aliases, command.as, created);
        return this.createdResult(scene, created);
      }

      // ---- text
      case "setText": {
        const { node: found } = await node(command.target);
        this.setTextOn(this.textElementOf(found, at("target")), command.text);
        return {};
      }
      case "replaceText": {
        const { node: found } = await node(command.target);
        const element = this.textElementOf(found, at("target"));
        const flat = paragraphsText(element.paragraphs || []);
        const ops: TextOp[] = [];
        for (let from = flat.indexOf(command.find); from >= 0; from = flat.indexOf(command.find, from + command.find.length)) {
          ops.unshift({ start: from, end: from + command.find.length, insert: command.replace });
          if (!command.all) break;
        }
        if (ops.length === 0) {
          throw new ApiError("FAILED", `${found.name} 里没有找到 "${command.find}"`, at("find"));
        }
        this.editor.applyTextOps(element, ops);
        return {};
      }
      case "formatText": {
        const { node: found } = await node(command.target);
        const element = this.textElementOf(found, at("target"));
        const length = paragraphsText(element.paragraphs || []).length;
        const from = command.range?.from ?? 0;
        const to = command.range?.to ?? length;
        if (from > to || to > length) {
          throw new ApiError("INVALID_SCHEMA", `范围 ${from}–${to} 超出文字长度 ${length}`, at("range"));
        }
        const { bold, italic, color, fontSizePt } = command;
        this.editor.applyTextOps(element, [{ start: from, end: to, insert: "", format: { bold, italic, color, fontSizePt } }]);
        return {};
      }
      case "setParagraph": {
        const { node: found } = await node(command.target);
        const element = this.textElementOf(found, at("target"));
        const length = paragraphsText(element.paragraphs || []).length;
        this.editor.applyTextOps(element, [{ start: 0, end: length, insert: "", paragraph: { align: command.align } }]);
        return {};
      }

      // ---- shapes
      case "setBounds": {
        await node(command.target);
        this.editor.setBounds({ x: command.x, y: command.y, w: command.w, h: command.h }, !!command.keepRatio);
        return {};
      }
      case "setRotation": {
        const { node: found } = await node(command.target);
        if (found.kind === "group") throw new ApiError("NOT_APPLICABLE", `${found.name} 是组合，暂不支持旋转`, at("target"));
        this.editor.setRotationDegrees(command.degrees);
        return {};
      }
      case "setFill": {
        const { node: found } = await node(command.target);
        this.requirePaintable(found, at("target"));
        this.editor.applyShapeStyle({ kind: "fill", color: command.color });
        return {};
      }
      case "setLine": {
        const { node: found } = await node(command.target);
        this.requirePaintable(found, at("target"));
        this.editor.applyShapeStyle(this.lineOp(command as { color?: string | null; widthPt?: number; dash?: keyof typeof DASH }));
        return {};
      }
      case "setAnchor": {
        const { node: found } = await node(command.target);
        this.textElementOf(found, at("target"));
        this.editor.applyShapeStyle({ kind: "anchor", value: command.anchor });
        return {};
      }
      case "align": {
        for (const [index, target] of (command.targets as Target[]).entries()) {
          await node(target, `targets[${index}]`);
          this.editor.alignToSlide(command.edge);
        }
        return {};
      }
      case "arrange": {
        await node(command.target);
        this.editor.arrange(command.mode);
        return {};
      }
      case "duplicate": {
        const { scene, node: found } = await node(command.target);
        const x0 = found.rect.x;
        await this.editor.duplicate();
        const created = this.editor.selectionRefs()[0];
        if (!created) throw new ApiError("FAILED", "创建副本失败", path);
        if (command.offset !== undefined) {
          const px = this.toPx(scene, command.offset);
          this.editor.setBounds({
            x: ((x0 + px) * this.emuPerPx(scene)) / CM,
            y: ((found.rect.y + px) * this.emuPerPx(scene)) / CM
          });
        }
        this.register(aliases, command.as, created);
        return this.createdResult(scene, { id: created.id, name: found.name, part: created.part });
      }
      case "copyTo": {
        const { node: found } = await node(command.target);
        const destination = resolveSlide(this.context(aliases).scenes, command.slide, at("slide"));
        if (!(await this.editor.copy())) throw new ApiError("UNSUPPORTED", `${found.name} 无法复制`, at("target"));
        await this.deck.showSlide(destination.index - 1);
        await this.editor.paste();
        const created = this.editor.selectionRefs()[0];
        if (!created) throw new ApiError("FAILED", "复制到目标页失败", path);
        this.register(aliases, command.as, created);
        return this.createdResult(destination, { id: created.id, name: found.name, part: created.part });
      }
      case "delete": {
        await node(command.target);
        this.editor.deleteSelection();
        return {};
      }
      default:
        throw new ApiError("INVALID_SCHEMA", `未知的 op "${command.op}"`, `${path}.op`);
    }
  }

  // ------------------------------------------------------------------ helpers

  private lineOp(line: { color?: string | null; widthPt?: number; dash?: keyof typeof DASH }): ShapeStyleOp {
    return {
      kind: "line",
      ...(line.color !== undefined ? { color: line.color } : {}),
      ...(line.widthPt !== undefined ? { widthPt: line.widthPt } : {}),
      ...(line.dash !== undefined ? { dash: DASH[line.dash] } : {})
    };
  }

  /** Replace all of an element's text; the new text takes the formatting of what it replaces. */
  private setTextOn(element: TextElement, text: string): void {
    const current = paragraphsText(element.paragraphs || []);
    if (current === text) return;
    this.editor.applyTextOps(element, [{ start: 0, end: current.length, insert: text }]);
  }

  /** Per-paragraph alignment and whole-text style for a freshly written text box. */
  private formatParagraphs(
    element: TextElement,
    paragraphs: Array<{ text: string; align?: string }>,
    style?: { fontSizePt?: number; bold?: boolean; italic?: boolean; color?: string; align?: string }
  ): void {
    const flat = paragraphsText(element.paragraphs || []);
    const ops: TextOp[] = [];
    if (style && (style.fontSizePt !== undefined || style.bold !== undefined || style.italic !== undefined || style.color !== undefined)) {
      ops.push({
        start: 0, end: flat.length, insert: "",
        format: { bold: style.bold, italic: style.italic, color: style.color, fontSizePt: style.fontSizePt }
      });
    }
    let cursor = 0;
    paragraphs.forEach(paragraph => {
      const align = (paragraph.align ?? style?.align) as "left" | "center" | "right" | undefined;
      if (align) ops.push({ start: cursor, end: cursor, insert: "", paragraph: { align } });
      cursor += paragraph.text.length + 1;
    });
    if (ops.length > 0) this.editor.applyTextOps(element, ops);
  }
}
