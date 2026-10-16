import { LINE_BREAK, TextFormat, TextModel, TextOp, applyTextOp } from "./editor/text-ops";
import { getDirectChild, getDirectChildren } from "./pptx-xml";

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RUN_TAGS = ["r", "fld", "br"];
const isRunNode = (node: Element) => RUN_TAGS.includes(node.localName);

const runText = (node: Element): string =>
  node.localName === "br" ? LINE_BREAK : getDirectChild(node, "t")?.textContent || "";

/** Adapter over a:p / a:r nodes of one p:txBody. */
function xmlTextModel(txBody: Element): TextModel<Element, Element> {
  const paragraphs = () => getDirectChildren(txBody, "p");
  const textRuns = (paragraph: Element) =>
    getDirectChildren(paragraph, ...RUN_TAGS).filter(node => runText(node) !== "");
  const doc = txBody.ownerDocument;

  return {
    paragraphs,
    runs: textRuns,
    text: runText,
    setText: (run, text) => {
      if (run.localName === "br") return;
      let t = getDirectChild(run, "t");
      if (!t) {
        t = doc.createElementNS(A_NS, "a:t");
        run.appendChild(t);
      }
      t.textContent = text;
    },
    removeRun: (paragraph, run) => {
      if (run.parentNode === paragraph) paragraph.removeChild(run);
    },
    cloneRun: (run, text) => {
      const copy = run.cloneNode(true) as Element;
      if (copy.localName === "br") return copy;
      let t = getDirectChild(copy, "t");
      if (!t) {
        t = doc.createElementNS(A_NS, "a:t");
        copy.appendChild(t);
      }
      t.textContent = text;
      return copy;
    },
    newRun: (paragraph, text) => {
      const ns = A_NS;
      const run = doc.createElementNS(ns, "a:r");
      // The paragraph's end properties describe the formatting of new text.
      const end = getDirectChild(paragraph, "endParaRPr");
      if (end) {
        const rPr = end.cloneNode(true) as Element;
        const renamed = doc.createElementNS(ns, "a:rPr");
        for (const attr of Array.from(rPr.attributes)) renamed.setAttribute(attr.name, attr.value);
        while (rPr.firstChild) renamed.appendChild(rPr.firstChild);
        run.appendChild(renamed);
      }
      const t = doc.createElementNS(ns, "a:t");
      t.textContent = text;
      run.appendChild(t);
      return run;
    },
    insertRun: (paragraph, index, run) => {
      const before = textRuns(paragraph)[index];
      const anchor = before || getDirectChild(paragraph, "endParaRPr");
      paragraph.insertBefore(run, anchor || null);
    },
    newParagraphAfter: paragraph => {
      const created = paragraph.cloneNode(false) as Element;
      const pPr = getDirectChild(paragraph, "pPr");
      if (pPr) created.appendChild(pPr.cloneNode(true));
      const end = getDirectChild(paragraph, "endParaRPr");
      if (end) created.appendChild(end.cloneNode(true));
      paragraph.parentNode!.insertBefore(created, paragraph.nextSibling);
      return created;
    },
    removeParagraph: paragraph => {
      if (getDirectChildren(txBody, "p").length > 1) paragraph.parentNode?.removeChild(paragraph);
    },
    setFormat: (run, format) => writeRunFormat(run, format),
    setParagraphFormat: (paragraph, format) => {
      if (format.align === undefined) return;
      let pPr = getDirectChild(paragraph, "pPr");
      if (!pPr) {
        pPr = doc.createElementNS(A_NS, "a:pPr");
        paragraph.insertBefore(pPr, paragraph.firstChild);
      }
      pPr.setAttribute("algn", { left: "l", center: "ctr", right: "r" }[format.align]);
    }
  };
}

const FILL_TAGS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];

function writeRunFormat(run: Element, format: TextFormat): void {
  const doc = run.ownerDocument;
  const ns = A_NS;
  let rPr = getDirectChild(run, "rPr");
  if (!rPr) {
    rPr = doc.createElementNS(ns, "a:rPr");
    run.insertBefore(rPr, run.firstChild);
  }
  if (format.bold !== undefined) rPr.setAttribute("b", format.bold ? "1" : "0");
  if (format.italic !== undefined) rPr.setAttribute("i", format.italic ? "1" : "0");
  if (format.fontSizePt !== undefined) rPr.setAttribute("sz", String(Math.round(format.fontSizePt * 100)));
  if (format.color !== undefined) {
    for (const fill of getDirectChildren(rPr, ...FILL_TAGS)) rPr.removeChild(fill);
    const solid = doc.createElementNS(ns, "a:solidFill");
    const rgb = doc.createElementNS(ns, "a:srgbClr");
    rgb.setAttribute("val", format.color.replace("#", "").toUpperCase());
    solid.appendChild(rgb);
    // Schema order: a:ln comes first, the fill right after it.
    const line = getDirectChild(rPr, "ln");
    rPr.insertBefore(solid, line ? line.nextSibling : rPr.firstChild);
  }
}

/** Replay recorded edits, in order, on the text body of a shape node. */
export function writeTextOps(shape: Element, ops: TextOp[]): boolean {
  const txBody = getDirectChild(shape, "txBody");
  if (!txBody) return false;
  const model = xmlTextModel(txBody);
  for (const op of ops) applyTextOp(model, op);
  return true;
}

export { isRunNode };
