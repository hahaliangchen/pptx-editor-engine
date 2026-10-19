/** A colour as the model keeps it: "#rrggbb", or "rgba(r,g,b,a)" when it is see-through. */

const hex2 = (value: number) => Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0");

/** The plain colour and the opacity (0..1) of a model colour string. */
export function splitColor(color: string): { hex: string; opacity: number } {
  const rgba = color.match(/^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)$/i);
  if (rgba) {
    return {
      hex: `#${hex2(Number(rgba[1]))}${hex2(Number(rgba[2]))}${hex2(Number(rgba[3]))}`,
      opacity: rgba[4] === undefined ? 1 : Math.max(0, Math.min(1, Number(rgba[4])))
    };
  }
  const short = color.match(/^#([0-9a-f])([0-9a-f])([0-9a-f])$/i);
  if (short) return { hex: `#${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`.toLowerCase(), opacity: 1 };
  return { hex: color.toLowerCase(), opacity: 1 };
}

/** The model's spelling of `color` at `opacity`: plain when fully opaque. */
export function withAlpha(color: string, opacity: number): string {
  const { hex } = splitColor(color);
  if (opacity >= 1 || !/^#[0-9a-f]{6}$/i.test(hex)) return hex;
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r},${g},${b},${Math.max(0, opacity)})`;
}
