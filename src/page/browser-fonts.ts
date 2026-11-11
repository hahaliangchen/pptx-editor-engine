/** TTC collections are accepted by the native renderer, but not by Firefox's
 * downloadable-font sanitizer. Extract the requested face as a standalone
 * SFNT, keeping its names, glyphs and metrics for both DOM and WASM rendering.
 */
export function browserFontBytes(bytes: Uint8Array, family: string, weight: number, italic: boolean): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const inside = (offset: number, size: number) => Number.isSafeInteger(offset) && Number.isSafeInteger(size) && offset >= 0 && size >= 0 && offset + size <= bytes.length;
  const tag = (offset: number) => String.fromCharCode(...bytes.subarray(offset, offset + 4));
  if (bytes.length < 4 || tag(0) !== "ttcf") return bytes;
  const invalid = () => new Error(`FONT_COLLECTION_INVALID: cannot extract ${family} from this font collection`);
  if (!inside(0, 12)) throw invalid();
  const count = view.getUint32(8);
  if (!count || count > 1024 || !inside(12, count * 4)) throw invalid();
  type Table = {tag: string; offset: number; size: number};
  let selected: {offset: number; tables: Table[]; score: number} | undefined;
  for (let face = 0; face < count; face++) {
    const offset = view.getUint32(12 + face * 4);
    if (!inside(offset, 12)) throw invalid();
    const n = view.getUint16(offset + 4);
    if (!n || n > 4095 || !inside(offset + 12, n * 16)) throw invalid();
    const tables: Table[] = [];
    for (let t = 0; t < n; t++) {
      const record = offset + 12 + t * 16;
      const entry = {tag: tag(record), offset: view.getUint32(record + 8), size: view.getUint32(record + 12)};
      if (!inside(entry.offset, entry.size)) throw invalid();
      tables.push(entry);
    }
    const names = tables.find(t => t.tag === "name");
    if (!names || names.size < 6) continue;
    const records = view.getUint16(names.offset + 2), strings = names.offset + view.getUint16(names.offset + 4);
    if (6 + records * 12 > names.size) throw invalid();
    const decoded = new Map<number, {text: string; rank: number}>();
    for (let r = 0; r < records; r++) {
      const record = names.offset + 6 + r * 12;
      const platform = view.getUint16(record), language = view.getUint16(record + 4), id = view.getUint16(record + 6);
      if (![0, 1, 3].includes(platform) || ![1, 2, 16, 17].includes(id)) continue;
      const size = view.getUint16(record + 8), start = strings + view.getUint16(record + 10);
      if (start < names.offset || start + size > names.offset + names.size || platform !== 1 && size % 2) throw invalid();
      let text = "";
      for (let i = 0; i < size; i += platform === 1 ? 1 : 2) text += String.fromCharCode(platform === 1 ? bytes[start + i] : view.getUint16(start + i));
      text = text.replace(/\0/g, "").trim();
      const rank = (platform === 3 ? 2 : platform === 0 ? 1 : 0) + (language === 0x409 ? 4 : 0);
      if (text && (!decoded.has(id) || rank > decoded.get(id)!.rank)) decoded.set(id, {text, rank});
    }
    const name = decoded.get(16)?.text || decoded.get(1)?.text;
    if (name?.toLowerCase() !== family.toLowerCase()) continue;
    const os2 = tables.find(t => t.tag === "OS/2"), head = tables.find(t => t.tag === "head");
    const faceWeight = os2 && os2.size >= 6 ? view.getUint16(os2.offset + 4) : 400;
    const faceItalic = os2 && os2.size >= 64 ? !!(view.getUint16(os2.offset + 62) & 1)
      : head && head.size >= 46 ? !!(view.getUint16(head.offset + 44) & 2) : false;
    const score = Math.abs(weight - faceWeight) + (italic === faceItalic ? 0 : 1000);
    if (!selected || score < selected.score) selected = {offset, tables, score};
  }
  if (!selected) throw new Error(`FONT_FACE_NOT_FOUND: ${family} is absent from the supplied font collection`);
  // An extracted face is no longer covered by the collection's digital signature.
  const tables = selected.tables.filter(t => t.tag !== "DSIG").sort((a, b) => a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0);
  const align = (size: number) => Math.ceil(size / 4) * 4;
  const total = 12 + tables.length * 16 + tables.reduce((sum, t) => sum + align(t.size), 0);
  if (total > 100 * 1024 * 1024) throw invalid();
  const result = new Uint8Array(total), output = new DataView(result.buffer);
  result.set(bytes.subarray(selected.offset, selected.offset + 4));
  output.setUint16(4, tables.length);
  const power = Math.floor(Math.log2(tables.length));
  output.setUint16(6, 16 * 2 ** power);output.setUint16(8, power);output.setUint16(10, tables.length * 16 - 16 * 2 ** power);
  const checksum = (start: number, size: number) => {
    let sum = 0;
    for (let i = start; i < start + size; i += 4) sum = (sum + output.getUint32(i)) >>> 0;
    return sum;
  };
  let destination = 12 + tables.length * 16, headOffset = -1;
  tables.forEach((table, index) => {
    const record = 12 + index * 16;
    for (let i = 0; i < 4; i++) result[record + i] = table.tag.charCodeAt(i);
    result.set(bytes.subarray(table.offset, table.offset + table.size), destination);
    if (table.tag === "head") {
      if (table.size < 12) throw invalid();
      headOffset = destination;output.setUint32(headOffset + 8, 0);
    }
    output.setUint32(record + 4, checksum(destination, align(table.size)));
    output.setUint32(record + 8, destination);output.setUint32(record + 12, table.size);
    destination += align(table.size);
  });
  if (headOffset < 0) throw invalid();
  output.setUint32(headOffset + 8, (0xb1b0afba - checksum(0, result.length)) >>> 0);
  return result;
}

/** CSS-connected FontFace objects cannot be added to a different document's
 * FontFaceSet in Firefox. The generic measurement fallback copies only faces
 * that the platform permits; the viewer supplies its own exact backend faces.
 */
export function copyScriptFontFaces(source: FontFaceSet, target: FontFaceSet): void {
  source.forEach(face => {
    try {(target as FontFaceSet & {add(face: FontFace): void}).add(face);}
    catch (error) {if ((error as {name?: string})?.name !== "InvalidModificationError") throw error;}
  });
}
