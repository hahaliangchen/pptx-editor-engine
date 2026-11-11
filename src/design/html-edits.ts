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
    if (at < 0) throw new Error(`HTML_EDIT_NOT_FOUND: edit ${index + 1} old_string does not match the current page; read the affected page and use its exact source`);
    if (next.indexOf(edit.old_string, at + 1) >= 0) throw new Error(`HTML_EDIT_AMBIGUOUS: edit ${index + 1} matches more than once; include surrounding HTML to identify one occurrence`);
    next = next.slice(0, at) + edit.new_string + next.slice(at + edit.old_string.length);
  }
  return next;
}
