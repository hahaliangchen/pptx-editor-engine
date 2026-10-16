import { TextParagraph, TextRun } from "../pptx-virtual-dom";
import { LINE_BREAK, TextModel } from "./text-ops";

const PT_EMU = 12700;

const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value));

export const paragraphsText = (paragraphs: TextParagraph[]): string =>
  paragraphs
    .map(paragraph => paragraph.runs.map(run => run.content === "\n" ? LINE_BREAK : run.content).join(""))
    .join("\n");

/**
 * Adapter over the renderer's paragraphs. The parser keeps one empty
 * placeholder run in a paragraph without text; that run is invisible to the
 * model and restored by `normalizeParagraphs`.
 */
export function vdomTextModel(
  paragraphs: TextParagraph[],
  emuPerPx: number
): TextModel<TextParagraph, TextRun> {
  const textRuns = (paragraph: TextParagraph) => paragraph.runs.filter(run => run.content !== "");
  return {
    paragraphs: () => paragraphs,
    runs: textRuns,
    text: run => run.content === "\n" ? LINE_BREAK : run.content,
    setText: (run, text) => { run.content = text === LINE_BREAK ? "\n" : text; },
    removeRun: (paragraph, run) => {
      const index = paragraph.runs.indexOf(run);
      if (index >= 0) paragraph.runs.splice(index, 1);
    },
    cloneRun: (run, text) => ({ content: text === LINE_BREAK ? "\n" : text, style: clone(run.style) }),
    newRun: (paragraph, text) => {
      const basis = paragraph.runs[paragraph.runs.length - 1];
      return { content: text, style: basis ? clone(basis.style) : { fontSize: 20, color: "#000000", bold: false, align: paragraph.style.align } };
    },
    insertRun: (paragraph, index, run) => {
      const target = textRuns(paragraph)[index];
      const at = target ? paragraph.runs.indexOf(target) : -1;
      if (at >= 0) paragraph.runs.splice(at, 0, run);
      else paragraph.runs.push(run);
    },
    newParagraphAfter: paragraph => {
      const created: TextParagraph = {
        style: clone(paragraph.style),
        bullet: paragraph.bullet ? clone(paragraph.bullet) : undefined,
        runs: paragraph.runs.length ? [{ content: "", style: clone(paragraph.runs[paragraph.runs.length - 1].style) }] : []
      };
      paragraphs.splice(paragraphs.indexOf(paragraph) + 1, 0, created);
      return created;
    },
    removeParagraph: paragraph => {
      const index = paragraphs.indexOf(paragraph);
      if (index >= 0 && paragraphs.length > 1) paragraphs.splice(index, 1);
    },
    setParagraphFormat: (paragraph, format) => {
      if (format.align === undefined) return;
      paragraph.style.align = format.align;
      for (const run of paragraph.runs) run.style.align = format.align;
    },
    setFormat: (run, format) => {
      if (format.bold !== undefined) run.style.bold = format.bold;
      if (format.italic !== undefined) run.style.italic = format.italic;
      if (format.color !== undefined) run.style.color = format.color;
      if (format.fontSizePt !== undefined) run.style.fontSize = format.fontSizePt * PT_EMU / emuPerPx;
    }
  };
}

/** Drop stale empty runs and make sure every paragraph keeps one run for the renderer. */
export function normalizeParagraphs(paragraphs: TextParagraph[], fallback: TextRun["style"]): void {
  for (const paragraph of paragraphs) {
    const hasText = paragraph.runs.some(run => run.content !== "");
    const last = paragraph.runs[paragraph.runs.length - 1];
    if (hasText) paragraph.runs = paragraph.runs.filter(run => run.content !== "");
    else paragraph.runs = [{ content: "", style: clone(last ? last.style : fallback) }];
  }
}
