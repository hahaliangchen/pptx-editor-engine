import { expandChart } from "./charts/layout";
import { expandTable } from "./tables/layout";
import { RenderElement, SlideElement } from "./pptx-virtual-dom";

/** Replace charts and tables by the shapes and text that draw them; everything else passes through. */
export function expandForRender(elements: SlideElement[]): RenderElement[] {
  const out: RenderElement[] = [];
  for (const element of elements) {
    if (element.type === "chart") out.push(...expandChart(element));
    else if (element.type === "table") out.push(...expandTable(element));
    else out.push(element);
  }
  return out;
}
