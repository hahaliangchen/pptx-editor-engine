/**
 * Text editing as a single "replace range with string" operation.
 *
 * A text body is treated as paragraphs joined by "\n"; a soft line break
 * (a:br) counts as one " " character inside its paragraph. The same
 * algorithm runs against the virtual DOM (to repaint immediately) and against
 * the OOXML tree (at export time), through the small TextModel adapter below.
 * Edits land inside existing runs wherever possible, so mixed formatting such
 * as a red word inside a black sentence survives typing.
 */

export const LINE_BREAK = " ";

/** Character formatting a toolbar can set; omitted fields stay as they are. */
export interface TextFormat {
  bold?: boolean;
  italic?: boolean;
  /** "#rrggbb" */
  color?: string;
  /** Font size in points. */
  fontSizePt?: number;
}

/** Paragraph-level formatting. */
export interface ParagraphFormat {
  align?: "left" | "center" | "right";
}

export interface TextOp {
  /** Offsets into the flat text, before the edit. */
  start: number;
  end: number;
  insert: string;
  /** When set the op only restyles [start, end) and `insert` is ignored. */
  format?: TextFormat;
  /** When set the op restyles every paragraph touched by [start, end]; `insert` is ignored. */
  paragraph?: ParagraphFormat;
}

export interface TextModel<P, R> {
  paragraphs(): P[];
  /** Runs that carry text, in order. A soft line break is a run with text LINE_BREAK. */
  runs(paragraph: P): R[];
  text(run: R): string;
  setText(run: R, text: string): void;
  removeRun(paragraph: P, run: R): void;
  /** Copy a run's formatting into a new, detached run holding `text`. */
  cloneRun(run: R, text: string): R;
  /** A fresh run for a paragraph that has no text runs to copy from. */
  newRun(paragraph: P, text: string): R;
  /** Insert so that the run becomes the `index`-th text run (or last when index is past the end). */
  insertRun(paragraph: P, index: number, run: R): void;
  /** Empty paragraph right after `paragraph`, sharing its paragraph properties. */
  newParagraphAfter(paragraph: P): P;
  removeParagraph(paragraph: P): void;
  setFormat(run: R, format: TextFormat): void;
  setParagraphFormat(paragraph: P, format: ParagraphFormat): void;
}

export function flatText<P, R>(model: TextModel<P, R>): string {
  return model.paragraphs()
    .map(paragraph => model.runs(paragraph).map(run => model.text(run)).join(""))
    .join("\n");
}

/** Single replacement covering everything between the shared prefix and suffix. */
function coarseDiff(before: string, after: string): TextOp {
  let prefix = 0;
  const max = Math.min(before.length, after.length);
  while (prefix < max && before[prefix] === after[prefix]) prefix++;
  let suffix = 0;
  while (
    suffix < max - prefix
    && before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  ) suffix++;

  const isLow = (code: number) => code >= 0xdc00 && code <= 0xdfff;
  const isHigh = (code: number) => code >= 0xd800 && code <= 0xdbff;
  // Never cut a surrogate pair in half.
  if (prefix > 0 && isHigh(before.charCodeAt(prefix - 1))) prefix--;
  if (suffix > 0 && isLow(before.charCodeAt(before.length - suffix))) suffix--;

  return {
    start: prefix,
    end: before.length - suffix,
    insert: after.slice(prefix, after.length - suffix)
  };
}

const MAX_DIFF_STEPS = 1500;

/**
 * Edits turning `before` into `after`, as separate hunks in descending
 * position order. Applying them in that order keeps every offset valid, and
 * unchanged text between two edits keeps its runs (and formatting).
 */
export function diffText(before: string, after: string): TextOp[] {
  if (before === after) return [];
  const coarse = coarseDiff(before, after);
  const A = Array.from(before.slice(coarse.start, coarse.end));
  const B = Array.from(after.slice(coarse.start, coarse.start + coarse.insert.length));
  if (A.length === 0 || B.length === 0) return [coarse];

  // Myers O(ND) diff over code points.
  const N = A.length;
  const M = B.length;
  const max = Math.min(N + M, MAX_DIFF_STEPS);
  const offset = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace: Int32Array[] = [];
  let found = -1;
  for (let d = 0; d <= max && found < 0; d++) {
    trace.push(v.slice());
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[k - 1 + offset] < v[k + 1 + offset])
        ? v[k + 1 + offset]
        : v[k - 1 + offset] + 1;
      let y = x - k;
      while (x < N && y < M && A[x] === B[y]) { x++; y++; }
      v[k + offset] = x;
      if (x >= N && y >= M) { found = d; break; }
    }
  }
  if (found < 0) return [coarse];

  // Walk back through the trace collecting moves: "=" keep, "-" delete, "+" insert.
  const moves: string[] = [];
  let x = N;
  let y = M;
  for (let d = found; d > 0; d--) {
    const previous = trace[d];
    const k = x - y;
    const down = k === -d || (k !== d && previous[k - 1 + offset] < previous[k + 1 + offset]);
    const prevK = down ? k + 1 : k - 1;
    const prevX = previous[prevK + offset];
    const prevY = prevX - prevK;
    while (x > (down ? prevX : prevX + 1) && y > (down ? prevY + 1 : prevY)) { moves.push("="); x--; y--; }
    moves.push(down ? "+" : "-");
    x = prevX;
    y = prevY;
  }
  while (x > 0 && y > 0) { moves.push("="); x--; y--; }
  moves.reverse();

  const ops: TextOp[] = [];
  let ai = 0;
  let bi = 0;
  let aChars = 0; // UTF-16 offset in the coarse region of `before`
  let current: TextOp | null = null;
  for (const move of moves) {
    if (move === "=") {
      if (current) { ops.push(current); current = null; }
      aChars += A[ai].length;
      ai++; bi++;
      continue;
    }
    if (!current) current = { start: coarse.start + aChars, end: coarse.start + aChars, insert: "" };
    if (move === "-") { aChars += A[ai].length; current.end = coarse.start + aChars; ai++; }
    else { current.insert += B[bi]; bi++; }
  }
  if (current) ops.push(current);
  return ops.reverse();
}

export function applyTextOp<P, R>(model: TextModel<P, R>, op: TextOp): void {
  if (op.format) {
    applyFormat(model, op.start, op.end, op.format);
    return;
  }
  if (op.paragraph) {
    applyParagraphFormat(model, op.start, op.end, op.paragraph);
    return;
  }
  const insert = op.insert.split(LINE_BREAK).join("");
  if (op.end > op.start) deleteRange(model, op.start, op.end);
  if (insert) insertText(model, op.start, insert);
}

interface Position<P> {
  paragraph: P;
  index: number;
  offset: number;
}

function paragraphLength<P, R>(model: TextModel<P, R>, paragraph: P): number {
  return model.runs(paragraph).reduce((sum, run) => sum + model.text(run).length, 0);
}

function locate<P, R>(model: TextModel<P, R>, position: number): Position<P> {
  const paragraphs = model.paragraphs();
  let cursor = 0;
  for (let index = 0; index < paragraphs.length; index++) {
    const length = paragraphLength(model, paragraphs[index]);
    if (position <= cursor + length || index === paragraphs.length - 1) {
      return { paragraph: paragraphs[index], index, offset: Math.min(Math.max(0, position - cursor), length) };
    }
    cursor += length + 1; // the "\n" between paragraphs
  }
  throw new Error("Text body has no paragraphs");
}

function deleteInParagraph<P, R>(model: TextModel<P, R>, paragraph: P, from: number, to: number): void {
  if (to <= from) return;
  let cursor = 0;
  for (const run of model.runs(paragraph)) {
    const text = model.text(run);
    const runStart = cursor;
    const runEnd = cursor + text.length;
    cursor = runEnd;
    const lo = Math.max(from, runStart);
    const hi = Math.min(to, runEnd);
    if (hi <= lo) continue;
    const remaining = text.slice(0, lo - runStart) + text.slice(hi - runStart);
    if (remaining) model.setText(run, remaining);
    else model.removeRun(paragraph, run);
  }
}

function deleteRange<P, R>(model: TextModel<P, R>, start: number, end: number): void {
  const from = locate(model, start);
  const to = locate(model, end);
  if (from.index === to.index) {
    deleteInParagraph(model, from.paragraph, from.offset, to.offset);
    return;
  }

  const fromLength = paragraphLength(model, from.paragraph);
  deleteInParagraph(model, from.paragraph, from.offset, fromLength);
  deleteInParagraph(model, to.paragraph, 0, to.offset);

  // Whatever is left of the last paragraph joins the first one.
  for (const run of model.runs(to.paragraph)) {
    model.removeRun(to.paragraph, run);
    model.insertRun(from.paragraph, Number.MAX_SAFE_INTEGER, run);
  }
  const paragraphs = model.paragraphs();
  for (const paragraph of paragraphs.slice(from.index + 1, to.index + 1)) {
    model.removeParagraph(paragraph);
  }
}

function insertInParagraph<P, R>(model: TextModel<P, R>, paragraph: P, offset: number, text: string): void {
  if (!text) return;
  const runs = model.runs(paragraph);
  const isBreak = (run: R) => model.text(run) === LINE_BREAK;

  let cursor = 0;
  let target: { run: R; at: number } | null = null;
  let insertIndex = runs.length;
  for (let i = 0; i < runs.length; i++) {
    const length = model.text(runs[i]).length;
    // Prefer extending the run that ends at (or contains) the caret.
    if (!isBreak(runs[i]) && offset > cursor && offset <= cursor + length) {
      target = { run: runs[i], at: offset - cursor };
      break;
    }
    if (offset <= cursor && insertIndex === runs.length) insertIndex = i;
    cursor += length;
  }
  if (!target && offset === 0) {
    const first = runs.find(run => !isBreak(run));
    if (first) target = { run: first, at: 0 };
  }

  if (target) {
    const current = model.text(target.run);
    model.setText(target.run, current.slice(0, target.at) + text + current.slice(target.at));
    return;
  }

  const template = runs.find(run => !isBreak(run));
  const fresh = template ? model.cloneRun(template, text) : model.newRun(paragraph, text);
  model.insertRun(paragraph, insertIndex, fresh);
}

function insertText<P, R>(model: TextModel<P, R>, position: number, text: string): void {
  const at = locate(model, position);
  const pieces = text.split("\n");
  if (pieces.length === 1) {
    insertInParagraph(model, at.paragraph, at.offset, text);
    return;
  }

  // Split the paragraph at the caret: everything after it moves to the last new paragraph.
  const tail: R[] = [];
  let cursor = 0;
  for (const run of model.runs(at.paragraph)) {
    const value = model.text(run);
    const runStart = cursor;
    cursor += value.length;
    if (cursor <= at.offset) continue;
    if (runStart >= at.offset) {
      model.removeRun(at.paragraph, run);
      tail.push(run);
    } else {
      const cut = at.offset - runStart;
      tail.push(model.cloneRun(run, value.slice(cut)));
      model.setText(run, value.slice(0, cut));
    }
  }

  insertInParagraph(model, at.paragraph, at.offset, pieces[0]);
  let previous = at.paragraph;
  for (let i = 1; i < pieces.length; i++) {
    const next = model.newParagraphAfter(previous);
    if (i === pieces.length - 1) {
      tail.forEach((run, index) => model.insertRun(next, index, run));
    }
    insertInParagraph(model, next, 0, pieces[i]);
    previous = next;
  }
}

/** Restyle [start, end), splitting runs at the edges so neighbours keep their look. */
function applyFormat<P, R>(model: TextModel<P, R>, start: number, end: number, format: TextFormat): void {
  if (end <= start) return;
  let cursor = 0;
  for (const paragraph of model.paragraphs()) {
    const length = paragraphLength(model, paragraph);
    const from = Math.max(start, cursor) - cursor;
    const to = Math.min(end, cursor + length) - cursor;
    cursor += length + 1;
    if (to <= from) continue;

    let runStart = 0;
    for (const run of model.runs(paragraph)) {
      const text = model.text(run);
      const runEnd = runStart + text.length;
      const lo = Math.max(from, runStart);
      const hi = Math.min(to, runEnd);
      const offset = runStart;
      runStart = runEnd;
      if (hi <= lo || text === LINE_BREAK) continue;

      const head = text.slice(0, lo - offset);
      const middle = text.slice(lo - offset, hi - offset);
      const tail = text.slice(hi - offset);
      const indexOf = () => model.runs(paragraph).indexOf(run);
      if (tail) model.insertRun(paragraph, indexOf() + 1, model.cloneRun(run, tail));
      if (head) model.insertRun(paragraph, indexOf(), model.cloneRun(run, head));
      model.setText(run, middle);
      model.setFormat(run, format);
    }
  }
}

/** Paragraph formatting reaches every paragraph the range touches; a caret touches one. */
function applyParagraphFormat<P, R>(model: TextModel<P, R>, start: number, end: number, format: ParagraphFormat): void {
  let cursor = 0;
  for (const paragraph of model.paragraphs()) {
    const length = paragraphLength(model, paragraph);
    if (start <= cursor + length && end >= cursor) model.setParagraphFormat(paragraph, format);
    cursor += length + 1;
  }
}
