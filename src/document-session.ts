import { PptVirtualDocument, PptxParser, Slide, SlideElement } from "./pptx-parser";
import { SlideEditor } from "./editor/slide-editor";
import { Command } from "./editor/history";
import { DeckPort, DecodedImage, EditApi } from "./api/engine";
import {
  blankSlideXml,
  layoutPlaceholders,
  placeholderSlideXml,
  prepareSlideCopy,
  relationshipsXml
} from "./editor/slides";
import {
  DirtyShapes,
  RelAdd,
  StructureEdits,
  StyleEdits,
  TextEdits,
  emptyStructure,
  exportEditedPptx,
  markDirty,
  materializeShape,
  materializeSlide,
  popOp,
  popTextOp,
  pushOp,
  pushTextOp,
  readRelationships,
  resolveTarget
} from "./pptx-exporter";

/** What the session needs from whoever shows the deck. Headless callers pass no-ops. */
export interface SessionView {
  /** A slide's content changed; redraw it if it is on screen. */
  invalidate(slide: Slide): void;
  beginDrag(slide: Slide, elements: SlideElement[]): void;
  endDrag(): void;
  /** An edit was committed to history. */
  slideEdited(slide: Slide): void;
  /** An element is being edited in place; the view should not draw its text. */
  beginTextEdit(slide: Slide, element: SlideElement): void;
  endTextEdit(slide: Slide): void;
  /** The slide list or the current page changed: redraw and tell the UI. */
  slidesChanged(): Promise<void>;
  /** Another page became the current one (the list is unchanged): redraw it. */
  currentChanged(): Promise<void>;
  /** A slide was just built (new page): load what it needs to draw, e.g. fonts. */
  prepareSlide(slide: Slide): Promise<void>;
}

/** Pixel size of a picture, which only an image decoder knows. */
export type ImageProbe = (bytes: Uint8Array, contentType: string, url: string) => Promise<{ width: number; height: number }>;

export interface SessionOptions {
  /** The editor mounts its overlay here (a real element in a browser, a detached one in Node). */
  container: HTMLElement;
  view: SessionView;
  probeImage: ImageProbe;
}

const IMAGE_TYPES: Record<string, { extension: string; contentType: string }> = {
  "image/png": { extension: "png", contentType: "image/png" },
  "image/jpeg": { extension: "jpeg", contentType: "image/jpeg" },
  "image/gif": { extension: "gif", contentType: "image/gif" }
};

/**
 * Everything about an opened deck that is not drawing: the parsed document, the
 * edit log that gets replayed on export, page management, history and the JSON
 * edit interface. The browser viewer and the Node headless session share it.
 */
export class DocumentSession {
  public readonly editor: SlideEditor;
  /** 0-based index of the page being edited. */
  public currentIndex = 0;

  private parser = new PptxParser();
  private doc: PptVirtualDocument | null = null;
  private dirtyShapes: DirtyShapes = new Map();
  private textEdits: TextEdits = new Map();
  private styleEdits: StyleEdits = new Map();
  private structureEdits: StructureEdits = emptyStructure();
  private shapeIdCounters = new Map<string, number>();
  private slideSerial = 0;
  private edit?: EditApi;
  private view: SessionView;
  private probeImage: ImageProbe;

  constructor(options: SessionOptions) {
    this.view = options.view;
    this.probeImage = options.probeImage;
    this.editor = new SlideEditor({
      container: options.container,
      getSlide: () => this.currentSlide(),
      getSlideSize: () => this.doc?.size || null,
      invalidate: slide => this.view.invalidate(slide),
      beginDrag: (slide, elements) => this.view.beginDrag(slide, elements),
      endDrag: () => this.view.endDrag(),
      showSlide: slide => {
        const index = this.doc?.slides.indexOf(slide) ?? -1;
        if (index >= 0) this.currentIndex = index;
      },
      onCommit: slide => this.view.slideEdited(slide),
      markDirty: source => markDirty(this.dirtyShapes, source),
      beginTextEdit: (slide, element) => this.view.beginTextEdit(slide, element),
      endTextEdit: slide => this.view.endTextEdit(slide),
      recordTextOp: (source, op) => pushTextOp(this.textEdits, source, op),
      discardTextOp: source => popTextOp(this.textEdits, source),
      structure: () => this.structureEdits,
      materialize: source => this.materializeSource(source),
      allocateShapeId: part => this.allocateShapeId(part),
      deckFont: () => this.deckFont(),
      allocateRelationshipId: part => this.allocateRelationshipId(part),
      allocateMediaPath: extension => this.allocateMediaPath(extension),
      resolveRelationships: (target, source, relationships) => this.resolveRelationships(target, source, relationships),
      recordStyleOp: (source, op) => pushOp(this.styleEdits, source, op),
      discardStyleOp: source => popOp(this.styleEdits, source)
    });
  }

  // ------------------------------------------------------------------ open

  public get presentation(): PptVirtualDocument | null {
    return this.doc;
  }

  public currentSlide(): Slide | null {
    if (!this.doc || this.currentIndex < 0 || this.currentIndex >= this.doc.slides.length) return null;
    return this.doc.slides[this.currentIndex];
  }

  /** Parse a .pptx package and start with a clean edit log. */
  public async open(buffer: ArrayBuffer): Promise<PptVirtualDocument> {
    this.doc = await this.parser.parse(buffer);
    this.reset();
    return this.doc;
  }

  /** Use a pre-built document (demos, tests); there is no package to export into. */
  public openVirtual(document: PptVirtualDocument): void {
    this.doc = document;
    this.parser.discardPackage();
    this.reset();
  }

  private reset(): void {
    this.currentIndex = 0;
    this.slideSerial = this.doc?.slides.length ?? 0;
    this.dirtyShapes = new Map();
    this.textEdits = new Map();
    this.styleEdits = new Map();
    this.structureEdits = emptyStructure();
    this.shapeIdCounters.clear();
    this.editor.reset();
  }

  // ---------------------------------------------------------------- export

  /** True when edits can be written back to an opened .pptx package. */
  public canExport(): boolean {
    return this.parser.getZip() !== null;
  }

  /** Build a new .pptx with every edit written back into the original package. */
  public async exportPptx(): Promise<Blob> {
    const zip = this.parser.getZip();
    if (!zip || !this.doc) throw new Error("请先打开一个 .pptx 文件。");
    const blob = await exportEditedPptx(zip, this.doc.slides, this.dirtyShapes, this.textEdits, this.styleEdits, this.structureEdits);
    this.editor.markSaved();
    return blob;
  }

  // ------------------------------------------------------------ slide list

  private editSet() {
    return {
      dirty: this.dirtyShapes,
      textEdits: this.textEdits,
      styleEdits: this.styleEdits,
      structure: this.structureEdits
    };
  }

  private allocateSlidePart(): string {
    const taken = [
      ...Object.keys(this.parser.getZip()?.files || {}),
      ...this.structureEdits.virtual.keys(),
      ...(this.doc?.slides || []).map(slide => slide.part || "")
    ];
    let highest = 0;
    for (const path of taken) {
      const match = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
      if (match) highest = Math.max(highest, parseInt(match[1], 10));
    }
    return `ppt/slides/slide${highest + 1}.xml`;
  }

  private cloneElements(elements: SlideElement[], part?: string): SlideElement[] {
    return elements.map(element => {
      const copy: SlideElement = JSON.parse(JSON.stringify(element));
      if (part && copy.source) copy.source.part = part;
      return copy;
    });
  }

  /** Swap the slide list (and current page) and tell the view; used by every slide command. */
  public async setSlides(slides: Slide[], index: number): Promise<void> {
    if (!this.doc) return;
    this.doc.slides = slides;
    this.currentIndex = Math.max(0, Math.min(index, slides.length - 1));
    this.editor.clearSelection();
    await this.view.slidesChanged();
  }

  private recordSlideChange(label: string, before: Slide[], beforeIndex: number, after: Slide[], afterIndex: number): void {
    const command: Command = {
      label,
      apply: () => { void this.setSlides([...after], afterIndex); },
      revert: () => { void this.setSlides([...before], beforeIndex); }
    };
    command.apply();
    this.editor.record(command);
  }

  private requirePackage(): void {
    if (!this.parser.getZip() || !this.doc) throw new Error("请先打开一个 .pptx 文件。");
  }

  /** A blank page with the same layout as `index`, right after it. */
  public async addSlide(index = this.currentIndex, layoutName?: string): Promise<void> {
    this.requirePackage();
    const zip = this.parser.getZip()!;
    const slides = this.doc!.slides;
    const reference = slides[index];
    if (!reference?.part) return;

    const known = await readRelationships(zip, reference.part, this.structureEdits.relAdds.get(reference.part), this.structureEdits.virtual);
    let layout = [...known.values()].find(rel => rel.type.endsWith("/slideLayout"));
    if (!layout) throw new Error("找不到这一页使用的版式。");
    if (layoutName) {
      const chosen = (await this.layoutParts()).find(entry => entry.name === layoutName);
      if (!chosen) throw new Error(`没有名为 "${layoutName}" 的版式。`);
      layout = { ...layout, target: `../slideLayouts/${chosen.part.split("/").pop()}` };
    }

    const part = this.allocateSlidePart();
    // The new page gets the layout's title and text placeholders, empty, like PowerPoint's "New Slide".
    const layoutPart = resolveTarget(reference.part.slice(0, reference.part.lastIndexOf("/")), layout.target);
    const layoutXml = await zip.file(layoutPart)?.async("string");
    const specs = layoutXml ? layoutPlaceholders(layoutXml) : [];
    const xml = specs.length > 0 ? placeholderSlideXml(specs) : blankSlideXml();
    const relsXml = relationshipsXml([{ ...layout, id: "rId1" }]);
    const relsPath = `${part.slice(0, part.lastIndexOf("/"))}/_rels/${part.slice(part.lastIndexOf("/") + 1)}.rels`;

    const slide = await this.parser.parseVirtualSlide(part, xml, relsXml, `slide_${++this.slideSerial}`, this.doc!.size);
    await this.view.prepareSlide(slide);
    this.structureEdits.virtual.set(part, xml);
    this.structureEdits.virtual.set(relsPath, relsXml);
    const after = [...slides];
    after.splice(index + 1, 0, slide);
    this.recordSlideChange("新建幻灯片", [...slides], this.currentIndex, after, index + 1);
  }

  /** A copy of slide `index`, with the edits made so far, right after it. */
  public async duplicateSlide(index = this.currentIndex): Promise<void> {
    this.requirePackage();
    const zip = this.parser.getZip()!;
    const slides = this.doc!.slides;
    const source = slides[index];
    if (!source?.part) return;

    const snapshot = await materializeSlide(zip, source.part, slides, this.editSet());
    if (!snapshot) return;
    const copy = prepareSlideCopy(snapshot);
    const part = this.allocateSlidePart();
    this.structureEdits.virtual.set(part, copy.xml);
    this.structureEdits.virtual.set(
      `${part.slice(0, part.lastIndexOf("/"))}/_rels/${part.slice(part.lastIndexOf("/") + 1)}.rels`,
      relationshipsXml(copy.relationships)
    );
    const slide: Slide = {
      id: `slide_${++this.slideSerial}`,
      elements: this.cloneElements(source.elements, part),
      rawXml: copy.xml,
      emuPerPx: source.emuPerPx,
      part,
      xform: source.xform
    };
    const after = [...slides];
    after.splice(index + 1, 0, slide);
    this.recordSlideChange("复制幻灯片", [...slides], this.currentIndex, after, index + 1);
  }

  public deleteSlide(index = this.currentIndex): void {
    this.requirePackage();
    const slides = this.doc!.slides;
    if (slides.length <= 1) throw new Error("至少需要保留一页幻灯片。");
    const after = slides.filter((_, i) => i !== index);
    this.recordSlideChange("删除幻灯片", [...slides], this.currentIndex, after, Math.min(index, after.length - 1));
  }

  public moveSlide(from: number, to: number): void {
    this.requirePackage();
    const slides = this.doc!.slides;
    if (from === to || from < 0 || from >= slides.length || to < 0 || to >= slides.length) return;
    const after = [...slides];
    const [moved] = after.splice(from, 1);
    after.splice(to, 0, moved);
    this.recordSlideChange("移动幻灯片", [...slides], this.currentIndex, after, to);
  }

  /** The deck's layouts: the name shown in PowerPoint and the part that holds it. */
  public async layoutParts(): Promise<Array<{ name: string; part: string }>> {
    const zip = this.parser.getZip();
    if (!zip) return [];
    const parts = Object.keys(zip.files)
      .filter(path => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path))
      .sort((a, b) => parseInt(a.replace(/\D/g, ""), 10) - parseInt(b.replace(/\D/g, ""), 10));
    const found: Array<{ name: string; part: string }> = [];
    for (const part of parts) {
      const xml = await zip.file(part)!.async("string");
      const name = xml.match(/<p:cSld\b[^>]*\bname="([^"]*)"/)?.[1];
      if (name) found.push({ name, part });
    }
    return found;
  }

  /** Layout names a new slide can use. */
  public async listLayouts(): Promise<string[]> {
    return [...new Set((await this.layoutParts()).map(entry => entry.name))];
  }

  // ---------------------------------------------------------------- images

  /** Pixels become slide units at 96 dpi: 1 px = 9525 EMU. */
  private decodePicture(bytes: Uint8Array, kind: { extension: string; contentType: string }, name?: string): Promise<DecodedImage> {
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: kind.contentType }));
    return this.probeImage(bytes, kind.contentType, url).then(size => {
      const emuPerPx = this.doc?.slides[0]?.emuPerPx || 6350;
      return {
        bytes, ...kind, url, name,
        width: (size.width * 9525) / emuPerPx,
        height: (size.height * 9525) / emuPerPx
      };
    });
  }

  /** Place a picture on the current slide. */
  public async insertPicture(bytes: Uint8Array, mimeType: string, name?: string): Promise<void> {
    const kind = IMAGE_TYPES[mimeType];
    if (!kind) throw new Error("仅支持 PNG、JPEG、GIF 图片。");
    await this.editor.insertImage(await this.decodePicture(bytes, kind, name) as Parameters<SlideEditor["insertImage"]>[0]);
  }

  /** Decode a picture given as a data: URL, in the units the editor places pictures in. */
  public async decodeDataUrl(dataUrl: string): Promise<DecodedImage> {
    const match = dataUrl.match(/^data:image\/(png|jpeg|gif);base64,(.*)$/s);
    if (!match) throw new Error("需要 PNG、JPEG 或 GIF 的 data: URL。");
    const binary = atob(match[2]);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return this.decodePicture(bytes, IMAGE_TYPES[`image/${match[1]}`]);
  }

  // ------------------------------------------------------------------- API

  /** The JSON interface: `api.outline()` to read the deck, `api.apply(commands)` to change it. */
  public get api(): EditApi {
    if (!this.edit) {
      const deck: DeckPort = {
        document: () => this.doc,
        currentIndex: () => this.currentIndex,
        showSlide: async index => {
          if (index === this.currentIndex || !this.doc || index < 0 || index >= this.doc.slides.length) return;
          this.currentIndex = index;
          await this.view.currentChanged();
        },
        addSlide: (after, layout) => this.addSlide(after, layout),
        duplicateSlide: index => this.duplicateSlide(index),
        deleteSlide: index => this.deleteSlide(index),
        moveSlide: (from, to) => this.moveSlide(from, to),
        layouts: () => this.listLayouts(),
        decodeImage: dataUrl => this.decodeDataUrl(dataUrl)
      };
      this.edit = new EditApi(this.editor, deck);
    }
    return this.edit;
  }

  // ------------------------------------------------- editor host callbacks

  private materializeSource(source: { part: string; shapeId: string }) {
    const zip = this.parser.getZip();
    if (!zip || !this.doc) return Promise.resolve(null);
    return materializeShape(zip, source.part, source.shapeId, this.doc.slides, this.editSet());
  }

  /** Next free p:cNvPr id of a slide, seeded from the highest id in its XML. */
  private allocateShapeId(part: string): string {
    let counter = this.shapeIdCounters.get(part);
    if (counter === undefined) {
      const xml = this.doc?.slides.find(slide => slide.part === part)?.rawXml || "";
      counter = 1;
      for (const match of xml.matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)) {
        counter = Math.max(counter, parseInt(match[1], 10));
      }
    }
    counter += 1;
    this.shapeIdCounters.set(part, counter);
    return String(counter);
  }

  /** Body font of the deck, taken from the first text that has one. */
  private deckFont(): { fontFamily?: string; eastAsianFontFamily?: string } {
    for (const slide of this.doc?.slides || []) {
      for (const element of slide.elements) {
        if (element.type !== "text" || !element.content.trim()) continue;
        const { fontFamily, eastAsianFontFamily } = element.style;
        return {
          ...(fontFamily && fontFamily !== "sans-serif" ? { fontFamily } : {}),
          ...(eastAsianFontFamily ? { eastAsianFontFamily } : {})
        };
      }
    }
    return {};
  }

  private async allocateRelationshipId(part: string): Promise<string> {
    const zip = this.parser.getZip();
    const existing = zip ? await readRelationships(zip, part, this.structureEdits.relAdds.get(part), this.structureEdits.virtual) : new Map();
    let counter = 0;
    for (const id of existing.keys()) counter = Math.max(counter, parseInt(String(id).replace(/\D/g, ""), 10) || 0);
    return `rId${counter + 1}`;
  }

  private allocateMediaPath(extension: string): string {
    const taken = [
      ...Object.keys(this.parser.getZip()?.files || {}),
      ...this.structureEdits.media.map(item => item.path)
    ];
    let highest = 0;
    for (const path of taken) {
      const match = path.match(/^ppt\/media\/image(\d+)\./i);
      if (match) highest = Math.max(highest, parseInt(match[1], 10));
    }
    return `ppt/media/image${highest + 1}.${extension}`;
  }

  /** Relationship ids for pasted shapes: reuse what the target slide has, add what it lacks. */
  private async resolveRelationships(
    targetPart: string,
    sourcePart: string,
    relationships: RelAdd[]
  ): Promise<{ ids: Map<string, string>; adds: RelAdd[] }> {
    const ids = new Map<string, string>();
    const adds: RelAdd[] = [];
    const zip = this.parser.getZip();
    if (targetPart === sourcePart || !zip || relationships.length === 0) return { ids, adds };

    const existing = await readRelationships(zip, targetPart, this.structureEdits.relAdds.get(targetPart), this.structureEdits.virtual);
    let counter = 0;
    for (const id of existing.keys()) counter = Math.max(counter, parseInt(id.replace(/\D/g, ""), 10) || 0);
    for (const rel of relationships) {
      const same = [...existing.values()].find(candidate =>
        candidate.type === rel.type && candidate.target === rel.target && candidate.targetMode === rel.targetMode);
      if (same) {
        ids.set(rel.id, same.id);
        continue;
      }
      counter += 1;
      const added: RelAdd = { ...rel, id: `rId${counter}` };
      existing.set(added.id, added);
      adds.push(added);
      ids.set(rel.id, added.id);
    }
    return { ids, adds };
  }
}
