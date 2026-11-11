export interface HtmlEdit { old_string: string; new_string: string }

/** Exact source edits, like a code editor. Validate every hunk before applying a page. */
export function editHtmlSource(html: string, edits: HtmlEdit[]): string {
  if (!Array.isArray(edits) || !edits.length) throw new Error('HTML_EDIT_INVALID: provide a nonempty edits array');
  let next = html;
  for (let index = 0; index < edits.length; index++) {
    const edit = edits[index];
    if (!edit || typeof edit.old_string !== 'string' || !edit.old_string || typeof edit.new_string !== 'string') {
      throw new Error(`HTML_EDIT_INVALID: edit ${index + 1} needs nonempty old_string and string new_string`);
    }
    const at = next.indexOf(edit.old_string);
    if (at < 0) throw new Error(`HTML_EDIT_NOT_FOUND: edit ${index + 1} old_string does not match the current page; copy it exactly from the source${nearestSource(next, edit.old_string)}`);
    if (next.indexOf(edit.old_string, at + 1) >= 0) throw new Error(`HTML_EDIT_AMBIGUOUS: edit ${index + 1} matches more than once; include surrounding HTML to identify one occurrence`);
    next = next.slice(0, at) + edit.new_string + next.slice(at + edit.old_string.length);
  }
  return next;
}

/** Where a mistyped anchor diverges from the real source, so the writer sees
 * the exact characters (an escaped quote, a rounded number) in one step. */
function nearestSource(source: string, wanted: string): string {
  let low = 0, high = wanted.length;
  while (low < high) { const mid = Math.ceil((low + high) / 2); if (source.includes(wanted.slice(0, mid))) low = mid; else high = mid - 1; }
  if (low < 8) return "";
  const at = source.indexOf(wanted.slice(0, low));
  return `; the source matches the first ${low} characters, then reads: ${JSON.stringify(source.slice(at + Math.max(0, low - 40), at + low + 80))}`;
}
