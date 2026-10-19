import JSZip from "jszip";
import { OPEN_FALLBACKS, OPEN_FONT_MAP, resolveOpenFamily } from "../page/font-map";

/**
 * The fonts a .pptx asks for that are not open-source, and what draws them
 * instead. The file keeps the original names; on a computer that has them the
 * deck looks as designed, elsewhere the stand-in is used.
 */
export interface FontSubstitution {
  family: string;
  replacement: string;
}

const OPEN_FAMILIES = new Set([
  ...Object.values(OPEN_FONT_MAP).map(name => name.toLowerCase()),
  ...OPEN_FALLBACKS.eastAsian.map(name => name.toLowerCase()),
  ...OPEN_FALLBACKS.latin.map(name => name.toLowerCase()),
  "liberation sans", "liberation serif", "liberation mono"
]);

const TYPEFACE = /<a:(latin|ea|cs)\b[^>]*\btypeface="([^"]*)"/g;

export async function deckFonts(bytes: ArrayBuffer | Uint8Array, catalog?: string[]): Promise<FontSubstitution[]> {
  const zip = await JSZip.loadAsync(bytes);
  const names = new Map<string, boolean>();
  for (const file of Object.values(zip.files)) {
    if (!/^ppt\/(slides|theme)\/[^/]+\.xml$/.test(file.name) && !/^ppt\/slideMasters\/[^/]+\.xml$/.test(file.name)) continue;
    const xml = await file.async("string");
    for (const match of xml.matchAll(TYPEFACE)) {
      const name = match[2].replace(/&amp;/g, "&").trim();
      names.set(name, names.get(name) || match[1] === "ea");
    }
  }
  const out: FontSubstitution[] = [];
  for (const [family, eastAsian] of [...names].sort(([a], [b]) => a.localeCompare(b))) {
    const key = family.toLowerCase();
    // "+mj-lt" and the like point at the theme's fonts, which are listed on their own.
    if (!family || family.startsWith("+") || key === "sans-serif" || key === "serif" || key === "monospace") continue;
    if (OPEN_FAMILIES.has(key) || catalog?.some(name => name.toLowerCase() === key)) continue;
    const mapped = OPEN_FONT_MAP[key];
    out.push({ family, replacement: catalog?.length ? resolveOpenFamily(family, eastAsian, catalog) : mapped ?? (/[^\x00-\x7f]/.test(family) ? OPEN_FALLBACKS.eastAsian[0] : OPEN_FALLBACKS.latin[0]) });
  }
  return out;
}
