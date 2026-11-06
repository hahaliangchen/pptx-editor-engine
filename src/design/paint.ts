import type { DesignAppearance } from "./model";

/**
 * Native DrawingML for the CSS paint the design layer accepts beyond solid
 * fills: linear gradients, one outer shadow, dashed outlines, top/bottom
 * corner pairs (round2SameRect) and polygons (custGeom). Export writes it and
 * import reads only what it can represent exactly, so a projected HTML page
 * shows the same CSS it was authored with.
 */
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const direct = (node: Element | undefined, name: string) => node ? Array.from(node.children).find(n => n.localName === name) : undefined;
const kids = (node: Element, name: string) => Array.from(node.children).filter(n => n.localName === name);
const round = (n: number, digits = 3) => Math.round(n * 10 ** digits) / 10 ** digits;

export function colorXml(color: string): string {
  let hex = color.replace(/^#/, ""), alpha = 1;
  if (hex.length === 3) hex = hex.split("").map(c => c + c).join("");
  const rgb = color.match(/^rgba?\(\s*([\d.]+)[, ]+([\d.]+)[, ]+([\d.]+)(?:\s*[,/]\s*([\d.]+))?\s*\)$/);
  if (rgb) { hex = rgb.slice(1, 4).map(n => Math.round(Number(n)).toString(16).padStart(2, "0")).join(""); alpha = rgb[4] === undefined ? 1 : Number(rgb[4]); }
  if (hex.length === 8) { alpha = parseInt(hex.slice(6), 16) / 255; hex = hex.slice(0, 6); }
  if (!/^[0-9a-f]{6}$/i.test(hex)) throw new Error(`Unsupported color: ${color}`);
  return `<a:srgbClr val="${hex.toUpperCase()}">${alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ""}</a:srgbClr>`;
}

/** The CSS spelling the browser reports for this color, so measured and imported values compare equal. */
function readColor(owner: Element | undefined): string | undefined {
  const clr = direct(owner, "srgbClr");
  const val = clr?.getAttribute("val");
  if (!clr || !val || !/^[0-9a-f]{6}$/i.test(val) || Array.from(clr.children).some(c => c.localName !== "alpha")) return undefined;
  const alpha = direct(clr, "alpha")?.getAttribute("val");
  if (alpha === undefined || alpha === null || Number(alpha) >= 100000) return `#${val.toLowerCase()}`;
  const [r, g, b] = [0, 2, 4].map(i => parseInt(val.slice(i, i + 2), 16));
  return `rgba(${r}, ${g}, ${b}, ${round(Number(alpha) / 100000, 4)})`;
}

const cssToOoxmlAngle = (deg: number) => Math.round((((deg - 90) % 360) + 360) % 360 * 60000);
const ooxmlToCssAngle = (ang: number) => round((((ang / 60000 + 90) % 360) + 360) % 360);

export function fillXml(a: Pick<DesignAppearance, "fill" | "gradient">): string {
  if (a.gradient) return `<a:gradFill rotWithShape="1"><a:gsLst>${a.gradient.stops.map(s => `<a:gs pos="${Math.round(s.position * 1000)}">${colorXml(s.color)}</a:gs>`).join("")}</a:gsLst><a:lin ang="${cssToOoxmlAngle(a.gradient.angle)}" scaled="0"/></a:gradFill>`;
  return a.fill && a.fill !== "transparent" ? `<a:solidFill>${colorXml(a.fill)}</a:solidFill>` : "<a:noFill/>";
}

export function lineXml(border: DesignAppearance["border"], emuPerCss: number): string {
  if (!border) return "<a:ln><a:noFill/></a:ln>";
  const dash = border.style === "dashed" ? `<a:prstDash val="dash"/>` : border.style === "dotted" ? `<a:prstDash val="sysDot"/>` : "";
  return `<a:ln w="${Math.round(border.width * emuPerCss)}"><a:solidFill>${colorXml(border.color)}</a:solidFill>${dash}</a:ln>`;
}

export function effectXml(shadow: DesignAppearance["shadow"], emuPerCss: number): string {
  if (!shadow) return "<a:effectLst/>";
  const dist = Math.hypot(shadow.x, shadow.y), dir = dist ? ((Math.atan2(shadow.y, shadow.x) * 180 / Math.PI) + 360) % 360 : 0;
  return `<a:effectLst><a:outerShdw blurRad="${Math.round(shadow.blur * emuPerCss)}" dist="${Math.round(dist * emuPerCss)}" dir="${Math.round(dir * 60000)}" algn="ctr" rotWithShape="0">${colorXml(shadow.color)}</a:outerShdw></a:effectLst>`;
}

export function geometryXml(a: Pick<DesignAppearance, "rect" | "radius" | "corners" | "polygon">): string {
  const side = Math.max(1, Math.min(a.rect.w, a.rect.h));
  const adj = (r: number) => Math.min(50000, Math.round(r / side * 100000));
  if (a.polygon) {
    const pt = ([x, y]: [number, number]) => `<a:pt x="${Math.round(x * 1000)}" y="${Math.round(y * 1000)}"/>`;
    const [first, ...rest] = a.polygon;
    return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100000" h="100000"><a:moveTo>${pt(first)}</a:moveTo>${rest.map(p => `<a:lnTo>${pt(p)}</a:lnTo>`).join("")}<a:close/></a:path></a:pathLst></a:custGeom>`;
  }
  if (a.corners) return `<a:prstGeom prst="round2SameRect"><a:avLst><a:gd name="adj1" fmla="val ${adj(a.corners.top)}"/><a:gd name="adj2" fmla="val ${adj(a.corners.bottom)}"/></a:avLst></a:prstGeom>`;
  const r = a.radius ? adj(a.radius) : 0;
  return `<a:prstGeom prst="${r ? "roundRect" : "rect"}"><a:avLst>${r ? `<a:gd name="adj" fmla="val ${r}"/>` : ""}</a:avLst></a:prstGeom>`;
}

/** CSS declarations matching the native paint; inverse of the browser measurement. */
export function paintCss(a: DesignAppearance, cssNumber: (n: number) => string): string[] {
  const css: string[] = [];
  if (a.fill) css.push(`background-color:${a.fill}`);
  if (a.gradient) css.push(`background-image:linear-gradient(${cssNumber(a.gradient.angle)}deg, ${a.gradient.stops.map(s => `${s.color} ${cssNumber(s.position)}%`).join(", ")})`);
  if (a.border) css.push(`border:${cssNumber(a.border.width)}px ${a.border.style ?? "solid"} ${a.border.color}`);
  if (a.radius !== undefined) css.push(`border-radius:${cssNumber(a.radius)}px`);
  if (a.corners) css.push(`border-radius:${cssNumber(a.corners.top)}px ${cssNumber(a.corners.top)}px ${cssNumber(a.corners.bottom)}px ${cssNumber(a.corners.bottom)}px`);
  if (a.polygon) css.push(`clip-path:polygon(${a.polygon.map(([x, y]) => `${cssNumber(x)}% ${cssNumber(y)}%`).join(", ")})`);
  if (a.shadow) css.push(`box-shadow:${cssNumber(a.shadow.x)}px ${cssNumber(a.shadow.y)}px ${cssNumber(a.shadow.blur)}px ${a.shadow.color}`);
  return css;
}

export type NativePaint = Pick<DesignAppearance, "fill" | "gradient" | "border" | "radius" | "corners" | "polygon" | "shadow">;

/**
 * Read explicit spPr paint written in the subset above. Returns undefined
 * when anything is outside it (theme colors, path gradients, other effects,
 * curved custom paths ...), and the caller keeps the object as a preview asset.
 */
export function readNativePaint(spPr: Element | undefined, rect: { w: number; h: number }, emuPerCss: number): NativePaint | undefined {
  if (!spPr) return undefined;
  const paint: NativePaint = {};
  const side = Math.max(1, Math.min(rect.w, rect.h));
  const gd = (geom: Element, name: string, fallback: number) => {
    const formula = kids(direct(geom, "avLst")!, "gd").find(g => g.getAttribute("name") === name)?.getAttribute("fmla");
    const value = formula?.match(/^val\s+(-?\d+)$/)?.[1];
    return value === undefined ? fallback : Number(value);
  };
  const prst = direct(spPr, "prstGeom"), cust = direct(spPr, "custGeom");
  if (prst) {
    const kind = prst.getAttribute("prst");
    if (!direct(prst, "avLst") && kind !== "rect") return undefined;
    if (kind === "roundRect") paint.radius = side * Math.min(50000, gd(prst, "adj", 16667)) / 100000;
    else if (kind === "round2SameRect") paint.corners = { top: side * Math.min(50000, gd(prst, "adj1", 16667)) / 100000, bottom: side * Math.min(50000, gd(prst, "adj2", 0)) / 100000 };
    else if (kind !== "rect") return undefined;
  } else if (cust) {
    const paths = kids(direct(cust, "pathLst") ?? cust, "path");
    if (paths.length !== 1) return undefined;
    const path = paths[0], w = Number(path.getAttribute("w")), h = Number(path.getAttribute("h"));
    if (!(w > 0 && h > 0) || path.getAttribute("fill") || path.getAttribute("stroke")) return undefined;
    const commands = Array.from(path.children), points: Array<[number, number]> = [];
    for (const [i, command] of commands.entries()) {
      const expected = i === 0 ? "moveTo" : i === commands.length - 1 ? "close" : "lnTo";
      if (command.localName !== expected) return undefined;
      if (expected === "close") continue;
      const pt = direct(command, "pt");
      const x = Number(pt?.getAttribute("x")), y = Number(pt?.getAttribute("y"));
      if (!Number.isFinite(x) || !Number.isFinite(y)) return undefined;
      points.push([round(x / w * 100), round(y / h * 100)]);
    }
    if (points.length < 3) return undefined;
    paint.polygon = points;
  } else return undefined;

  const fill = Array.from(spPr.children).find(n => ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(n.localName));
  if (fill?.localName === "solidFill") { const c = readColor(fill); if (!c) return undefined; paint.fill = c; }
  else if (fill?.localName === "gradFill") {
    const lin = direct(fill, "lin"), stops = kids(direct(fill, "gsLst") ?? fill, "gs");
    if (!lin || direct(fill, "path") || direct(fill, "tileRect") || stops.length < 2) return undefined;
    const parsed = stops.map(gs => ({ color: readColor(gs), position: round(Number(gs.getAttribute("pos")) / 1000) }));
    if (parsed.some(s => !s.color || !Number.isFinite(s.position))) return undefined;
    paint.gradient = { angle: ooxmlToCssAngle(Number(lin.getAttribute("ang") ?? 0)), stops: parsed as Array<{ color: string; position: number }> };
  } else if (fill && fill.localName !== "noFill") return undefined;

  const ln = direct(spPr, "ln");
  if (ln) {
    const lineFill = Array.from(ln.children).find(n => ["noFill", "solidFill", "gradFill", "pattFill"].includes(n.localName));
    if (lineFill?.localName === "solidFill") {
      const color = readColor(lineFill), dash = direct(ln, "prstDash")?.getAttribute("val") ?? "solid";
      const style = dash === "solid" ? undefined : ["dash", "sysDash", "lgDash"].includes(dash) ? "dashed" as const : ["sysDot", "dot"].includes(dash) ? "dotted" as const : null;
      if (!color || style === null || direct(ln, "custDash") || !ln.getAttribute("w")) return undefined;
      paint.border = { color, width: Number(ln.getAttribute("w")) / emuPerCss, ...(style ? { style } : {}) };
    } else if (lineFill && lineFill.localName !== "noFill") return undefined;
  }

  const effects = direct(spPr, "effectLst");
  if (effects?.children.length) {
    const shadow = direct(effects, "outerShdw");
    if (effects.children.length !== 1 || !shadow) return undefined;
    const color = readColor(shadow);
    if (!color || ["sx", "sy", "kx", "ky"].some(k => shadow.hasAttribute(k) && !["0", "100000"].includes(shadow.getAttribute(k)!))) return undefined;
    const dist = Number(shadow.getAttribute("dist") ?? 0) / emuPerCss, dir = Number(shadow.getAttribute("dir") ?? 0) / 60000 * Math.PI / 180;
    paint.shadow = { x: round(dist * Math.cos(dir)), y: round(dist * Math.sin(dir)), blur: round(Number(shadow.getAttribute("blurRad") ?? 0) / emuPerCss), color };
  } else if (direct(spPr, "effectDag")) return undefined;
  // HTML expresses a polygon only as a clip-path, which clips borders and shadows away; a stroked or shadowed
  // polygon (an outlined SVG icon path, say) stays a preview asset so its page still round-trips.
  if (paint.polygon && (paint.border || paint.shadow)) return undefined;
  return paint;
}

/** Whether the spPr carries anything beyond plain rect/roundRect solid paint. */
export const hasExtendedPaint = (p: NativePaint) => !!(p.gradient || p.shadow || p.corners || p.polygon || p.border?.style);

export { A as DRAWING_NS };
