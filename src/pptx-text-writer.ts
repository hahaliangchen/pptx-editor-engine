import { LINE_BREAK, TextModel, TextOp, applyTextOp } from "./editor/text-ops";
import { getDirectChild, getDirectChildren } from "./pptx-xml";

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
        t = doc.createElementNS(run.namespaceURI, "a:t");
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
        t = doc.createElementNS(copy.namespaceURI, "a:t");
        copy.appendChild(t);
      }
      t.textContent = text;
      return copy;
    },
    newRun: (paragraph, text) => {
      const ns = paragraph.namespaceURI;
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
    }
  };
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
