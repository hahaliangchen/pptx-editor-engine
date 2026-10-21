import { parseBinding } from "../api/template";
import type { ElementSource, ShapeElement, Slide, SlideElement, TextElement, PptVirtualDocument } from "../pptx-virtual-dom";
import { cloneDesign, DesignAppearance, DesignDocument, DesignNode, DesignParagraph, DesignRect, DesignPage } from "./model";
import { applyDefaultPolicies, readTemplateRecord } from "./template-policy";
import { bytesToBase64, sha256Hex } from "./bytes";
export interface ImportRuntime {
  dataUrl(url: string): Promise<string>;
  preview(slide: Slide, elements: SlideElement[], bounds: DesignRect): Promise<string>;
}
const identity = (part: string, shape: string) => `${part}#${shape}`;
const xmlShapeId = (node: Element): string | undefined => Array.from(node.children).find(n => n.localName.startsWith("nv"))?.getElementsByTagNameNS("*", "cNvPr")[0]?.getAttribute("id") ?? undefined;

function textAppearance(text: TextElement, scale: number): NonNullable<DesignAppearance["text"]> {
  const paragraphs: DesignParagraph[] = (text.paragraphs ?? [{ runs: [{ content: text.content, style: text.style }], style: { align: text.style.align, level: 0, marginLeft: 0, indent: 0 } }]).map(p => {
    const first = p.runs[0]?.style ?? text.style;
    const size = first.fontSize * scale;
    const line = p.style.lineSpacing;
    return {
      runs: p.runs.map(run => ({ text: run.content.replace(/\n/g, "\u2028"), fontFamily: run.style.eastAsianFontFamily || run.style.fontFamily || "Arial", fontSize: run.style.fontSize * scale, color: run.style.color, bold: run.style.bold, italic: !!run.style.italic, letterSpacing: (run.style.letterSpacing ?? 0) * scale })),
      align: p.style.align,
      // Parser normalizes percent to a ratio and point spacing to logical pixels.
      lineHeight: line?.unit === "percent" ? size * line.value : line?.unit === "points" ? line.value * scale : size * 1.2,
      before: (p.style.spaceBefore ?? 0) * scale,
      after: (p.style.spaceAfter ?? 0) * scale
    };
  });
  return { paragraphs, padding: [text.body?.marginTop ?? 0, text.body?.marginRight ?? 0, text.body?.marginBottom ?? 0, text.body?.marginLeft ?? 0].map(n => n * scale) as [number, number, number, number], verticalAlign: text.body?.verticalAnchor ?? "top" };
}


/** Project a parsed package into transient design JSON on either runtime. */
export async function importDocument(bytes: Uint8Array, doc: PptVirtualDocument, runtime: ImportRuntime): Promise<DesignDocument> {
    if (Math.abs(doc.size.width / doc.size.height - 16 / 9) > 0.005) throw new Error("Only 16:9 PPTX is supported by the 1280x720 design layer");
    const scale = 1280 / doc.size.width;
    const emuPerPx = doc.slides[0]?.emuPerPx;
    if (!emuPerPx) throw new Error("PPTX slide units are missing");
    const design: DesignDocument = { schemaVersion: 1, revision: 0, canvas: { width: 1280, height: 720, unit: "css-px" }, source: { pptxBase64: bytesToBase64(bytes), sha256: sha256Hex(bytes), widthEMU: Math.round(doc.size.width * emuPerPx), heightEMU: Math.round(doc.size.height * emuPerPx) }, pages: [], diagnostics: [] };
    const rect = (r: DesignRect): DesignRect => ({ x: r.x * scale, y: r.y * scale, w: r.w * scale, h: r.h * scale });
    for (const slide of doc.slides) {
      await new Promise<void>(resolve => setTimeout(resolve, 0));
      const part = slide.part!;
      const page: DesignPage = { id: part, part, nodes: [] as DesignNode[] };
      design.pages.push(page);
      const byShape = new Map<string, SlideElement[]>();
      for (const element of slide.elements) if (element.source?.part === part) byShape.set(element.source.shapeId, [...(byShape.get(element.source.shapeId) ?? []), element]);
      const decoration = slide.elements.filter(e => !e.source || e.source.part !== part);
      if (decoration.length) page.nodes.push({ id: `${part}#backdrop`, name: "Master and background", kind: "asset", lock: "all", required: true, metadata: { backdrop: true, protectedRegions: decoration.filter(e => e.rect.w * scale < 1200 || e.rect.h * scale < 650).map(e => rect(e.rect)) }, appearance: { rect: { x: 0, y: 0, w: 1280, h: 720 }, src: await runtime.preview(slide, decoration, { x: 0, y: 0, w: 1280, h: 720 }) } });
      const xml = new DOMParser().parseFromString(slide.rawXml, "application/xml");
      const record = readTemplateRecord(xml);
      page.semantics = record?.page;
      const tree = xml.getElementsByTagNameNS("*", "spTree")[0];
      const walk = async (parent: Element, parentId?: string, referenceParent = false): Promise<void> => {
        for (const raw of Array.from(parent.children)) {
          if (!["sp", "pic", "grpSp", "graphicFrame", "cxnSp", "AlternateContent", "contentPart"].includes(raw.localName)) continue;
          const shapeId = xmlShapeId(raw) ?? `opaque-${page.nodes.length}`;
          const id = identity(part, shapeId);
          const elements = byShape.get(shapeId) ?? [];
          const source: ElementSource | undefined = elements[0]?.source;
          const name = raw.getElementsByTagNameNS("*", "cNvPr")[0]?.getAttribute("name") ?? shapeId;
          const binding = parseBinding(name);
          const referenceVisual = referenceParent || !!binding?.reference;
          if (raw.localName === "grpSp") {
            const reference = slide.elements.flatMap(e => e.source?.groups ?? []).find(g => g.shapeId === shapeId);
            const bounds = reference ? rect(reference.origRect) : { x: 0, y: 0, w: 1280, h: 720 };
            page.nodes.push({ id, name, kind: "group", parentId, binding, role: binding?.role, appearance: { rect: bounds }, metadata: { sourceShapeId: shapeId }, origin: { part, shapeId, appearance: { rect: cloneDesign(bounds) }, transform: { scaleX: scale, scaleY: scale, offsetX: 0, offsetY: 0 } } });
            await walk(raw, id, referenceVisual);
            continue;
          }
          if (!source || !elements.length) {
            page.nodes.push({ id, name, kind: "asset", parentId, lock: "all", hidden: true, appearance: { rect: { x: 0, y: 0, w: 0, h: 0 } }, metadata: { sourceShapeId: shapeId, originalElement: raw.localName } });
            design.diagnostics.push({ pageId: part, nodeId: id, message: "Original object preserved; no preview is available. Supply an asset preview before redesigning this object." });
            continue;
          }
          const shape = elements.find((e): e is ShapeElement => e.type === "shape");
          const text = elements.find((e): e is TextElement => e.type === "text");
          const picture = elements.find(e => e.type === "image");
          const bounds = rect(elements[0].rect);
          const appearance: DesignAppearance = { rect: bounds };
          let kind: DesignNode["kind"] = picture ? "image" : text ? "text" : "shape";
          // Unsupported visual grammar stays one asset; its source XML remains untouched.
          const complex = elements.some(e => e.type === "chart" || e.type === "table" || !!e.rotation)
            || !!shape && (!(["rect", "roundRect"].includes(shape.shapeType)) || !!shape.customGeometry || !!shape.computedStyle?.effects || !!shape.computedStyle?.fill && !["solid", "none"].includes(shape.computedStyle.fill.type))
            || !!picture && (!!picture.crop || !!picture.effects || picture.flipH || picture.flipV)
            || !!text && (!!text.body?.textDirection && text.body.textDirection !== "horz" || !!text.style.shadow || !!text.style.outline || !!text.paragraphs?.some(p => p.bullet || p.style.level || p.style.indent || p.style.marginLeft));
          if (complex) {
            kind = "asset";
            appearance.src = await runtime.preview(slide, elements, bounds);
          } else if (picture?.type === "image") appearance.src = await runtime.dataUrl(picture.url);
          else {
            if (shape) {
              const fill = shape.computedStyle?.fill;
              if (fill?.type === "solid") appearance.fill = fill.color;
              else if (!fill && shape.fill !== "transparent") appearance.fill = shape.fill;
              const line = shape.computedStyle?.line;
              if (line?.fill.type === "solid" && line.width > 0) appearance.border = { color: line.fill.color, width: line.width * scale };
              if (shape.shapeType === "roundRect") appearance.radius = Math.min(bounds.w, bounds.h) * (shape.cornerRadius ?? 0.16667);
            }
            if (text) appearance.text = textAppearance(text, scale);
          }
          page.nodes.push({ id, name, parentId, kind, appearance, binding, role: binding?.role ?? text?.placeholder?.type,
            metadata: { sourceShapeId: shapeId, originalName: name, ...(referenceVisual ? { referenceVisual: true } : {}), ...(text?.placeholder ? { placeholder: cloneDesign(text.placeholder) } : {}) },
            origin: { part, shapeId, appearance: cloneDesign(appearance), transform: { scaleX: source.xform.scaleX * scale, scaleY: source.xform.scaleY * scale, offsetX: source.xform.offsetX * scale, offsetY: source.xform.offsetY * scale } } });
        }
      };
      const bg = tree?.parentElement?.getElementsByTagNameNS("*", "bgPr")[0];
      const solid = bg && Array.from(bg.children).find(e => e.localName === "solidFill");
      const rgb = solid?.getElementsByTagNameNS("*", "srgbClr")[0]?.getAttribute("val");
      if (rgb) { page.background = `#${rgb.toLowerCase()}`; page.originBackground = page.background; }
      if (tree) await walk(tree);
      for (const node of page.nodes) {
        const stored = record?.nodes[node.origin?.shapeId ?? String(node.metadata?.sourceShapeId ?? "")];
        if (stored && (!stored.name || stored.name === node.name)) { node.purpose = stored.purpose; node.note = stored.note; }
      }
    }
    applyDefaultPolicies(design);
    return design;
}
