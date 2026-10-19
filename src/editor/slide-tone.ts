import { splitColor } from "../color";
import { FillStyle, Slide } from "../pptx-virtual-dom";

const luminance = (hex: string): number => {
  const channel = (offset: number) => parseInt(hex.slice(offset, offset + 2), 16) / 255;
  return 0.299 * channel(1) + 0.587 * channel(3) + 0.114 * channel(5);
};

const fillLuminance = (fill: FillStyle | undefined): number | null => {
  if (fill?.type === "solid") return luminance(splitColor(fill.color).hex);
  if (fill?.type === "gradient" && fill.stops.length > 0) {
    return fill.stops.reduce((sum, stop) => sum + luminance(splitColor(stop.color).hex), 0) / fill.stops.length;
  }
  return null;
};

/** Is the page's background dark, so that text put on it should be light? */
export function isDarkSlide(slide: Slide): boolean {
  const background = slide.elements.find(element => element.type === "shape" && element.id === "slide_background");
  const level = background && background.type === "shape" ? fillLuminance(background.computedStyle?.fill) : null;
  return level !== null && level < 0.45;
}

/** The colour text should have on this page when nothing says otherwise: what a deck's own text colour would be. */
export const defaultTextColor = (slide: Slide): string => (isDarkSlide(slide) ? "#ffffff" : "#000000");
