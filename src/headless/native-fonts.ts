import { readFileSync, readdirSync } from "node:fs";
import { delimiter, extname, join, resolve } from "node:path";
import { OPEN_FALLBACKS, OPEN_FONT_MAP, resolveOpenFamily } from "../page/font-map";
import type { FontRequest } from "./fonts";

export interface LocalFont { family: string; file: string; weight: number; italic: boolean }
const OPEN_FAMILIES = [...new Set([
  ...Object.values(OPEN_FONT_MAP), ...OPEN_FALLBACKS.eastAsian, ...OPEN_FALLBACKS.latin,
  "Noto Serif SC", "Noto Sans SC", "Noto Serif CJK SC", "Source Han Sans CN"
])];

export function nativeFontDirectories(env: NodeJS.ProcessEnv = process.env, platform = process.platform): string[] {
  const dirs = (env.PPT_FONT_DIRS || "").split(delimiter).filter(Boolean);
  if (platform === "win32") {
    if (env.WINDIR) dirs.push(join(env.WINDIR, "Fonts"));
    if (env.LOCALAPPDATA) dirs.push(join(env.LOCALAPPDATA, "Microsoft", "Windows", "Fonts"));
  }
  return dirs;
}

function styleOf(style: string): Pick<LocalFont, "weight" | "italic"> {
  const s = style.toLowerCase();
  const weight = /black|heavy/.test(s) ? 900 : /extra ?bold/.test(s) ? 800 : /semi ?bold|demi/.test(s) ? 600
    : /bold/.test(s) ? 700 : /medium/.test(s) ? 500 : /light/.test(s) ? 300 : /thin/.test(s) ? 100 : 400;
  return { weight, italic: /italic|oblique/.test(s) };
}

/** Read OpenType name tables, including collection faces; never trust filenames. */
export function fontMetadata(data: Buffer, file: string): LocalFont[] {
  const inside = (offset: number, size: number) => offset >= 0 && size >= 0 && offset + size <= data.length;
  if (!inside(0, 12)) return [];
  const collection = data.toString("ascii", 0, 4) === "ttcf";
  const count = collection ? data.readUInt32BE(8) : 1;
  if (count > 1024 || (collection && !inside(12, count * 4))) return [];
  const faces: LocalFont[] = [];
  for (let face = 0; face < count; face++) {
    const base = collection ? data.readUInt32BE(12 + face * 4) : 0;
    if (!inside(base, 12)) continue;
    const signature = data.toString("ascii", base, base + 4);
    if (data.readUInt32BE(base) !== 0x00010000 && !["OTTO", "true"].includes(signature)) continue;
    const tables = data.readUInt16BE(base + 4);
    if (!inside(base + 12, tables * 16)) continue;
    for (let table = 0; table < tables; table++) {
      const record = base + 12 + table * 16;
      if (data.toString("ascii", record, record + 4) !== "name") continue;
      const start = data.readUInt32BE(record + 8), size = data.readUInt32BE(record + 12);
      if (!inside(start, size) || size < 6) continue;
      const records = data.readUInt16BE(start + 2), storage = data.readUInt16BE(start + 4);
      if (6 + records * 12 > size) continue;
      const names = new Map<number, { text: string; rank: number }>();
      for (let n = 0; n < records; n++) {
        const entry = start + 6 + n * 12;
        const platform = data.readUInt16BE(entry), language = data.readUInt16BE(entry + 4), id = data.readUInt16BE(entry + 6);
        if (![1, 2, 16, 17].includes(id) || ![0, 1, 3].includes(platform)) continue;
        const length = data.readUInt16BE(entry + 8), offset = storage + data.readUInt16BE(entry + 10);
        if (offset + length > size || ((platform === 0 || platform === 3) && length % 2)) continue;
        const bytes = data.subarray(start + offset, start + offset + length);
        const text = (platform === 1 ? bytes.toString("latin1") : Buffer.from(bytes).swap16().toString("utf16le")).replace(/\0/g, "").trim();
        const rank = (platform === 3 ? 2 : platform === 0 ? 1 : 0) + (language === 0x409 ? 4 : 0);
        if (text && (!names.has(id) || names.get(id)!.rank < rank)) names.set(id, { text, rank });
      }
      const family = names.get(16)?.text || names.get(1)?.text;
      if (family) faces.push({ family, file, ...styleOf(names.get(17)?.text || names.get(2)?.text || "Regular") });
    }
  }
  return faces;
}

export function scanLocalFonts(dirs: string[], allowed = OPEN_FAMILIES): LocalFont[] {
  const canonical = new Map(allowed.map(family => [family.toLowerCase(), family]));
  const faces: LocalFont[] = [], visited = new Set<string>();
  const walk = (dir: string) => {
    let entries;
    try { entries = readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
      const file = resolve(dir, entry.name), key = file.toLowerCase();
      if (visited.has(key) || entry.isSymbolicLink()) continue;
      visited.add(key);
      if (entry.isDirectory()) { walk(file); continue; }
      if (!/\.(ttf|otf|ttc)$/i.test(extname(file))) continue;
      try {
        for (const face of fontMetadata(readFileSync(file), file)) {
          const family = canonical.get(face.family.toLowerCase());
          if (family) faces.push({ ...face, family });
        }
      } catch { /* An inaccessible or malformed system font must not break the catalog. */ }
    }
  };
  dirs.filter(Boolean).forEach(walk);
  return faces;
}

export function matchLocalFont(faces: LocalFont[], request: FontRequest): string {
  if (!faces.length) throw new Error("No supported open-source fonts found. Install Noto/Carlito/Liberation in Windows Fonts or set PPT_FONT_DIRS.");
  const family = resolveOpenFamily(request.family, request.eastAsian, [...new Set(faces.map(face => face.family))]);
  const weight = request.bold ? 700 : 400;
  const candidates = faces.filter(face => face.family === family);
  return candidates.reduce((best, face) => {
    const score = (f: LocalFont) => Math.abs(f.weight - weight) + (f.italic === request.italic ? 0 : 1000);
    return score(face) < score(best) ? face : best;
  }).file;
}
