/**
 * Pictures for a template's picture slots: cut to the slot's proportions (the middle of the
 * picture, like "cover"), and cut to its outline when the slot is round or rounded. Works with
 * the canvas of the browser page and of the headless renderer alike.
 */

export interface FitSpec {
  /** Width ÷ height of the slot. */
  aspect: number;
  shape?: "rect" | "ellipse" | "roundRect";
  /** Corner radius of a rounded slot as a share of its shorter side (default 0.16, the preset's). */
  radius?: number;
  /** The longest side of the result in pixels (default 1600). */
  maxSide?: number;
}

const MAX_SIDE = 1600;

interface Surface {
  canvas: any;
  context: any;
  width: number;
  height: number;
}

function surface(width: number, height: number): Surface {
  const canvas = (globalThis as any).document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  return { canvas, context: canvas.getContext("2d"), width, height };
}

function size(spec: FitSpec, longest: number): { width: number; height: number } {
  const side = Math.min(spec.maxSide ?? MAX_SIDE, Math.max(64, Math.round(longest)));
  return spec.aspect >= 1
    ? { width: side, height: Math.max(1, Math.round(side / spec.aspect)) }
    : { width: Math.max(1, Math.round(side * spec.aspect)), height: side };
}

/** Clip the next drawing to the slot's outline. */
function clipTo(context: any, shape: FitSpec["shape"], width: number, height: number, radius: number): void {
  if (!shape || shape === "rect") return;
  context.beginPath();
  if (shape === "ellipse") {
    context.ellipse(width / 2, height / 2, width / 2, height / 2, 0, 0, Math.PI * 2);
  } else {
    const r = Math.min(width, height) * Math.min(0.5, Math.max(0, radius));
    context.moveTo(r, 0);
    context.arcTo(width, 0, width, height, r);
    context.arcTo(width, height, 0, height, r);
    context.arcTo(0, height, 0, 0, r);
    context.arcTo(0, 0, width, 0, r);
    context.closePath();
  }
  context.clip();
}

/** How a picture is loaded: the browser's image by default; the headless renderer plugs in its own. */
export type ImageLoader = (url: string) => Promise<any>;

const browserLoader: ImageLoader = url => new Promise<any>((resolve, reject) => {
  const image = new (globalThis as any).Image();
  image.onload = () => resolve(image);
  image.onerror = () => reject(new Error("图片无法解码"));
  image.src = url;
});

let loadImage: ImageLoader = browserLoader;
export function setImageLoader(loader: ImageLoader | undefined): void {
  loadImage = loader ?? browserLoader;
}

/** A data URL of `url`'s picture, cut to the slot. A mask needs transparency, so the result is then a PNG. */
export async function fitImage(url: string, spec: FitSpec): Promise<string> {
  const image = await loadImage(url);
  const sourceW: number = image.width;
  const sourceH: number = image.height;
  if (!sourceW || !sourceH) throw new Error("图片没有尺寸");
  const sourceAspect = sourceW / sourceH;
  // The middle of the picture, a little high for a tall one (faces and subjects tend to sit there).
  let cropW = sourceW;
  let cropH = sourceH;
  if (sourceAspect > spec.aspect) cropW = sourceH * spec.aspect;
  else cropH = sourceW / spec.aspect;
  const cropX = (sourceW - cropW) / 2;
  const cropY = (sourceH - cropH) * 0.4;
  const out = size(spec, Math.max(cropW, cropH));
  const { canvas, context } = surface(out.width, out.height);
  context.save();
  clipTo(context, spec.shape, out.width, out.height, spec.radius ?? 0.16);
  context.drawImage(image, cropX, cropY, cropW, cropH, 0, 0, out.width, out.height);
  context.restore();
  return canvas.toDataURL(spec.shape && spec.shape !== "rect" ? "image/png" : "image/jpeg", 0.9);
}

/** A stand-in for a picture that is not there yet: a soft tinted field with a small landscape on it. */
export function placeholderImage(spec: FitSpec, tint = "#9aa7b4"): string {
  const out = size(spec, 800);
  const { canvas, context, width, height } = surface(out.width, out.height);
  context.save();
  clipTo(context, spec.shape, width, height, spec.radius ?? 0.16);
  context.fillStyle = tint;
  context.globalAlpha = 0.28;
  context.fillRect(0, 0, width, height);
  context.globalAlpha = 0.55;
  context.fillStyle = tint;
  const unit = Math.min(width, height);
  // sun and two hills
  context.beginPath();
  context.arc(width * 0.68, height * 0.34, unit * 0.09, 0, Math.PI * 2);
  context.fill();
  context.beginPath();
  context.moveTo(width * 0.18, height * 0.74);
  context.lineTo(width * 0.42, height * 0.44);
  context.lineTo(width * 0.62, height * 0.74);
  context.closePath();
  context.fill();
  context.beginPath();
  context.moveTo(width * 0.46, height * 0.74);
  context.lineTo(width * 0.64, height * 0.52);
  context.lineTo(width * 0.84, height * 0.74);
  context.closePath();
  context.fill();
  context.restore();
  return canvas.toDataURL("image/png");
}
