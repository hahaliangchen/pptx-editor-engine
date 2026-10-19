/**
 * Which open-source font draws text that asks for a common Office font. The
 * file keeps the original name (PowerPoint still shows Microsoft YaHei where it
 * is installed); only rendering swaps it. Carlito, Caladea and Liberation are
 * metric-compatible with the fonts they stand in for, so line breaks match.
 */
export const OPEN_FONT_MAP: Record<string, string> = {
  "microsoft yahei": "Noto Sans CJK SC",
  "微软雅黑": "Noto Sans CJK SC",
  "microsoft yahei ui": "Noto Sans CJK SC",
  "dengxian": "Noto Sans CJK SC",
  "等线": "Noto Sans CJK SC",
  "simhei": "Noto Sans CJK SC",
  "黑体": "Noto Sans CJK SC",
  "pingfang sc": "Noto Sans CJK SC",
  "source han sans cn": "Noto Sans CJK SC",
  "noto sans cjk sc": "Noto Sans CJK SC",
  "simsun": "Noto Serif CJK SC",
  "宋体": "Noto Serif CJK SC",
  "nsimsun": "Noto Serif CJK SC",
  "新宋体": "Noto Serif CJK SC",
  "fangsong": "Noto Serif CJK SC",
  "仿宋": "Noto Serif CJK SC",
  "kaiti": "Noto Serif CJK SC",
  "楷体": "Noto Serif CJK SC",
  "calibri": "Carlito",
  "calibri light": "Carlito",
  "aptos": "Carlito",
  "cambria": "Caladea",
  "arial": "Liberation Sans",
  "helvetica": "Liberation Sans",
  "segoe ui": "Liberation Sans",
  "times new roman": "Liberation Serif",
  "courier new": "Liberation Mono",
  "consolas": "Liberation Mono"
};

/** Used when neither the requested font nor a mapped one is available. */
export const OPEN_FALLBACKS = {
  eastAsian: ["Noto Sans CJK SC", "Noto Sans SC", "Source Han Sans CN"],
  latin: ["Carlito", "Liberation Sans", "Noto Sans", "DejaVu Sans"]
};

const GENERIC = new Set(["sans-serif", "serif", "monospace", ""]);

/**
 * The family to load from a font catalog for `requested`: the exact family
 * when the catalog has it, else its open-source stand-in, else a fallback for
 * the script, else the first family on offer.
 */
export function resolveOpenFamily(requested: string, eastAsian: boolean, catalog: string[]): string {
  const byLowerName = new Map(catalog.map(family => [family.toLowerCase(), family]));
  const wanted = requested.trim().toLowerCase();
  const pick = (name: string | undefined) => (name ? byLowerName.get(name.toLowerCase()) : undefined);

  if (!GENERIC.has(wanted)) {
    const exact = pick(wanted);
    if (exact) return exact;
    const mapped = pick(OPEN_FONT_MAP[wanted]);
    if (mapped) return mapped;
  }
  for (const candidate of eastAsian ? OPEN_FALLBACKS.eastAsian : OPEN_FALLBACKS.latin) {
    const found = pick(candidate);
    if (found) return found;
  }
  if (catalog[0]) return catalog[0];
  throw new Error("字体服务没有提供任何字体");
}
