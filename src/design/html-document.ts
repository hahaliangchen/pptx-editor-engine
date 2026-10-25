import type { DesignDocument } from "./model";

/** Inspect in a truly inert template before DOMParser can load an image URL. */
export function parseDesignHtml(source: string): Document {
  const template = document.createElement("template");
  template.innerHTML = source;
  const root = template.content;
  if (root.querySelector("script,iframe,object,embed,link,base,meta[http-equiv],svg,math,video,audio,source,track,input")) {
    throw new Error("Only static HTML/CSS and embedded image assets are accepted");
  }
  for (const element of Array.from(root.querySelectorAll("*"))) {
    if (Array.from(element.attributes).some(a => /^on/i.test(a.name) || ["srcset", "poster", "background"].includes(a.name.toLowerCase()))) {
      throw new Error("Active HTML attributes are not allowed");
    }
    if (element.hasAttribute("src") && (element.tagName !== "IMG"
      || !/^(data:image\/|design-asset:|project:images\/)/.test(element.getAttribute("src") ?? ""))) {
      throw new Error("Images must use embedded project assets");
    }
  }
  return new DOMParser().parseFromString(source, "text/html");
}

/** Page identity is structural; revision and source hash are informational only. */
export function assertDesignHtmlIdentity(design: DesignDocument, body: HTMLElement, expectedPageId?: string): void {
  const pageId = body.dataset.pageId;
  if (!pageId || !design.pages.some(page => page.id === pageId) || expectedPageId && pageId !== expectedPageId) {
    throw new Error("Unknown or mismatched design page");
  }
}
