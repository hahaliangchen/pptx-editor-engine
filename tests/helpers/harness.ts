import JSZip from "jszip";
import { EditApi, DeckPort, DecodedImage } from "../../src/api/engine";
import { SlideEditor } from "../../src/editor/slide-editor";
import { Command } from "../../src/editor/history";
import {
  EditSet,
  emptyStructure,
  exportEditedPptx,
  markDirty,
  materializeShape,
  popOp,
  popTextOp,
  pushOp,
  pushTextOp
} from "../../src/pptx-exporter";
import { PptxParser } from "../../src/pptx-parser";
import { PptVirtualDocument, Slide } from "../../src/pptx-virtual-dom";
import { DeckParts, buildDeck } from "./deck";

/** What the viewer does around the editor, minus rendering: enough to run the JSON API for real. */
export async function openHarness(parts: DeckParts) {
  const parser = new PptxParser();
  const doc: PptVirtualDocument = await parser.parse((await buildDeck(parts)) as unknown as ArrayBuffer);
  const zip = parser.getZip() as JSZip;
  let current = 0;

  const edits: EditSet = {
    dirty: new Map(),
    textEdits: new Map(),
    styleEdits: new Map(),
    structure: emptyStructure()
  };
  const counters = new Map<string, number>();

  const container = document.createElement("div");
  document.body.appendChild(container);
  const editor = new SlideEditor({
    container,
    getSlide: () => doc.slides[current] ?? null,
    getSlideSize: () => doc.size,
    invalidate: () => undefined,
    beginDrag: () => undefined,
    endDrag: () => undefined,
    showSlide: slide => { current = doc.slides.indexOf(slide); },
    onCommit: () => undefined,
    markDirty: source => markDirty(edits.dirty, source),
    beginTextEdit: () => undefined,
    endTextEdit: () => undefined,
    recordTextOp: (source, op) => pushTextOp(edits.textEdits, source, op),
    discardTextOp: source => popTextOp(edits.textEdits, source),
    recordStyleOp: (source, op) => pushOp(edits.styleEdits, source, op),
    discardStyleOp: source => popOp(edits.styleEdits, source),
    structure: () => edits.structure,
    materialize: source => materializeShape(zip, source.part, source.shapeId, doc.slides, edits),
    allocateShapeId: part => {
      let counter = counters.get(part);
      if (counter === undefined) {
        const xml = doc.slides.find(slide => slide.part === part)?.rawXml || "";
        counter = Math.max(1, ...Array.from(xml.matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)).map(match => parseInt(match[1], 10)));
      }
      counters.set(part, counter + 1);
      return String(counter + 1);
    },
    deckFont: () => ({}),
    allocateRelationshipId: async () => "rId9",
    allocateMediaPath: extension => `ppt/media/image1.${extension}`,
    resolveRelationships: async () => ({ ids: new Map(), adds: [] })
  });

  const record = (label: string, before: Slide[], after: Slide[], index: number) => {
    const command: Command = {
      label,
      apply: () => { doc.slides = [...after]; current = index; },
      revert: () => { doc.slides = [...before]; current = 0; }
    };
    command.apply();
    editor.record(command);
  };

  const deck: DeckPort = {
    document: () => doc,
    currentIndex: () => current,
    showSlide: async index => { current = index; },
    addSlide: async after => {
      const slide: Slide = { id: `slide_new_${doc.slides.length + 1}`, elements: [], rawXml: "", part: `ppt/slides/slide${doc.slides.length + 1}.xml`, emuPerPx: doc.slides[0].emuPerPx, xform: doc.slides[0].xform };
      const next = [...doc.slides];
      next.splice(after + 1, 0, slide);
      record("新建幻灯片", [...doc.slides], next, after + 1);
    },
    duplicateSlide: async index => {
      const next = [...doc.slides];
      next.splice(index + 1, 0, { ...doc.slides[index], id: `slide_copy_${doc.slides.length + 1}`, elements: [...doc.slides[index].elements] });
      record("复制幻灯片", [...doc.slides], next, index + 1);
    },
    deleteSlide: index => record("删除幻灯片", [...doc.slides], doc.slides.filter((_, i) => i !== index), 0),
    moveSlide: (from, to) => {
      const next = [...doc.slides];
      const [moved] = next.splice(from, 1);
      next.splice(to, 0, moved);
      record("移动幻灯片", [...doc.slides], next, to);
    },
    layouts: async () => ["标题和内容", "空白"],
    decodeImage: async (): Promise<DecodedImage> => ({
      bytes: new Uint8Array([137, 80, 78, 71]), extension: "png", contentType: "image/png", url: "blob:test", width: 400, height: 300
    })
  };

  const api = new EditApi(editor, deck);
  const toBuffer = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });

  return {
    api, editor, doc, zip, edits,
    /** The slide part's XML as an export would write it right now. */
    async exported(part = "ppt/slides/slide1.xml"): Promise<string> {
      const blob = await exportEditedPptx(zip, doc.slides, edits.dirty, edits.textEdits, edits.styleEdits, edits.structure);
      return (await JSZip.loadAsync(await toBuffer(blob))).file(part)!.async("string");
    },
    steps: () => editor.history.entries().map(entry => entry.label)
  };
}
