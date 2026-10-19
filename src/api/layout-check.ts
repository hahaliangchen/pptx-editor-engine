import { CM_TO_PT, INSET_CM, widthEm } from "./deck-spec";
import type { CmRect, Outline, OutlineShape } from "./outline";

/**
 * Looks for the problems a model would otherwise have to find by looking at a
 * picture: text that does not fit its box, objects off the page, objects on top
 * of each other, a page that is half empty. It works from the outline's numbers
 * alone, so it costs nothing and misses nothing it can measure.
 */

export type IssueKind = "text_overflow" | "out_of_bounds" | "overlap" | "empty_space" | "small_text" | "small_object";

export interface LayoutIssue {
  slide: number;
  kind: IssueKind;
  /** The objects concerned (ids on that slide). */
  ids: string[];
  message: string;
}

const LINE_GAP = 1.3;
const TOLERANCE = 0.05;

const flatten = (shapes: OutlineShape[]): OutlineShape[] =>
  shapes.flatMap(shape => (shape.children ? [shape, ...flatten(shape.children)] : [shape]));

const round = (value: number) => Math.round(value * 10) / 10;

/** What a text needs: its widest line, and the height of its lines when they wrap at `width`. */
function measure(shape: OutlineShape): { oneLine: number; wrappedHeight: number; lineHeight: number; lines: number } {
  const size = shape.style?.fontSizePt ?? 18;
  const lines = (shape.text ?? "").split("\n");
  const perLine = Math.max(1, ((shape.rect.w - 2 * INSET_CM) * CM_TO_PT) / size);
  const wrapped = lines.reduce((sum, line) => sum + Math.max(1, Math.ceil(widthEm(line) / perLine)), 0);
  const lineHeight = (size * LINE_GAP) / CM_TO_PT;
  return {
    oneLine: (Math.max(...lines.map(widthEm)) * size) / CM_TO_PT + 2 * INSET_CM,
    wrappedHeight: wrapped * lineHeight,
    lineHeight,
    lines: wrapped
  };
}

const overlapArea = (a: CmRect, b: CmRect) => {
  const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return w > 0 && h > 0 ? w * h : 0;
};

/** Faint text (a watermark) is decoration, and may bleed off the page or sit behind other text. */
const isContent = (shape: OutlineShape) =>
  shape.type === "chart" || shape.type === "table"
  || (!!shape.text && shape.text.trim() !== "" && (shape.style?.opacity ?? 1) >= 0.5);

/** The part of a text box its text actually fills (top anchored, as far as the box is wide). */
function inkOf(shape: OutlineShape): CmRect {
  if (!shape.text) return shape.rect;
  const need = measure(shape);
  const h = Math.min(shape.rect.h, need.wrappedHeight);
  const w = Math.min(shape.rect.w, need.oneLine);
  const align = shape.style?.align;
  const x = align === "center" ? shape.rect.x + (shape.rect.w - w) / 2 : align === "right" ? shape.rect.x + shape.rect.w - w : shape.rect.x;
  return { x, y: shape.rect.y, w, h };
}

export function checkLayout(outline: Outline, options: { slides?: number[] } = {}): LayoutIssue[] {
  const issues: LayoutIssue[] = [];
  const { w: pageW, h: pageH } = outline.slideSize;
  for (const slide of outline.slides) {
    if (options.slides && !options.slides.includes(slide.index)) continue;
    const shapes = flatten(slide.shapes);
    const content = shapes.filter(shape => isContent(shape) && shape.type !== "group");
    const add = (kind: IssueKind, ids: string[], message: string) => issues.push({ slide: slide.index, kind, ids, message });
    const label = (shape: OutlineShape) => `${shape.name}（id ${shape.id}${shape.text ? `，“${shape.text.slice(0, 12)}${shape.text.length > 12 ? "…" : ""}”` : ""}）`;

    for (const shape of content) {
      const { x, y, w, h } = shape.rect;
      // Decoration may run off the page on purpose; content may not.
      if (x < -TOLERANCE || y < -TOLERANCE || x + w > pageW + TOLERANCE || y + h > pageH + TOLERANCE) {
        add("out_of_bounds", [shape.id], `${label(shape)} 超出页面（页面 ${round(pageW)}×${round(pageH)} 厘米，它在 x ${round(x)}–${round(x + w)}，y ${round(y)}–${round(y + h)}）`);
      }
      if (shape.type === "chart" && (shape.rect.h < 5 || shape.rect.w < 8)) {
        add("small_object", [shape.id], `${label(shape)} 太小（${round(shape.rect.w)}×${round(shape.rect.h)} 厘米），坐标轴和文字会挤在一起；放大它，或去掉同页别的内容`);
      }
      // Sizes are judged against the page: 10 pt on a 33.9 cm wide page.
      const smallest = 10 * (pageW / 33.87);
      if (shape.text && (shape.style?.fontSizePt ?? 99) < smallest && shape.rect.y < pageH * 0.9) {
        add("small_text", [shape.id], `${label(shape)} 的字号只有 ${shape.style!.fontSizePt} 磅，在这个页面上太小，看不清；放大到至少 ${Math.ceil(smallest)} 磅，或精简文字`);
      }
      if (shape.text) {
        const need = measure(shape);
        // A text that fits on one line is fine, whatever the box's height; otherwise it wraps (or is cut) and must fit by height.
        if (need.oneLine > shape.rect.w * 1.1 + 0.3 && need.wrappedHeight > shape.rect.h * 1.1 + 0.15) {
          add("text_overflow", [shape.id], `${label(shape)} 的文字放不下（一行需要约 ${round(need.oneLine)} 厘米宽，折行后需要约 ${round(need.wrappedHeight)} 厘米高；框是 ${round(shape.rect.w)}×${round(shape.rect.h)}）；缩小字号、加宽加高文本框或删减文字`);
        }
      }
    }

    for (let i = 0; i < content.length; i++) {
      for (let j = i + 1; j < content.length; j++) {
        const a = inkOf(content[i]);
        const b = inkOf(content[j]);
        const common = overlapArea(a, b);
        if (common > 0.2 * Math.min(a.w * a.h, b.w * b.h) && common > 0.3) {
          add("overlap", [content[i].id, content[j].id], `${label(content[i])} 与 ${label(content[j])} 重叠`);
        }
      }
    }

    // A page with a chart, a table or several blocks that stops short of the page's lower part.
    const body = content.filter(shape => shape.rect.y > pageH * 0.14 && shape.rect.y + shape.rect.h < pageH * 0.93);
    const rich = body.length >= 3 || body.some(shape => shape.type === "chart" || shape.type === "table");
    if (rich) {
      const bottom = Math.max(...body.map(shape => inkOf(shape).y + inkOf(shape).h));
      if (bottom < pageH * 0.68) {
        add("empty_space", [], `内容只到页面的 ${Math.round((bottom / pageH) * 100)}% 处，下面大片空白；可以放大图表或文字、增加内容，或把内容垂直居中`);
      }
    }
  }
  return issues;
}

/** The issues as lines a model can act on, grouped by page. */
export function describeIssues(issues: LayoutIssue[]): string {
  const pages = [...new Set(issues.map(issue => issue.slide))];
  return pages.map(page => `第 ${page} 页：\n${issues.filter(issue => issue.slide === page).map(issue => `  - [${issue.kind}] ${issue.message}`).join("\n")}`).join("\n");
}
