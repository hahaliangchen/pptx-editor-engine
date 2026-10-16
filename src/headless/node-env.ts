import { JSDOM } from "jsdom";
import { Canvas, DOMMatrix, Image, ImageData, createCanvas } from "@napi-rs/canvas";

/**
 * Gives Node what the editor and the renderer expect from a browser.
 *
 * The editor builds its overlay with the DOM, so jsdom's globals are installed.
 * The Rust renderer is written against `web_sys`: it looks up `window.document`,
 * makes canvases with `createElement("canvas")` and type-checks values with
 * `instanceof`, so those get Skia-backed stand-ins.
 */
export function installNodeEnv(): void {
  const scope = globalThis as any;
  if (scope.__pptHeadlessEnv) return;
  if (typeof scope.document !== "undefined") {
    throw new Error("headless 模式需要在 Node 环境运行：当前已存在浏览器 document。");
  }

  // Defined first so jsdom's own versions of these names do not win below.
  class NodeWindow {
    static [Symbol.hasInstance](value: unknown) { return value === globalThis; }
  }
  scope.Window = NodeWindow;
  scope.HTMLCanvasElement = Canvas;
  scope.HTMLImageElement = Image;
  scope.CanvasRenderingContext2D = createCanvas(1, 1).getContext("2d").constructor;
  scope.ImageData = ImageData;
  scope.DOMMatrix = DOMMatrix;

  const dom = new JSDOM("<!doctype html><html><body></body></html>", { pretendToBeVisual: true });
  const win = dom.window as any;
  for (const key of Object.getOwnPropertyNames(win)) {
    if (key in scope || key === "window" || key === "self" || key === "top" || key === "parent" || key === "frames") continue;
    Object.defineProperty(scope, key, {
      configurable: true,
      get: () => win[key],
      set: value => Object.defineProperty(scope, key, { value, writable: true, configurable: true })
    });
  }

  const document = win.document as Document;
  const createElement = document.createElement.bind(document);
  (document as any).createElement = (tag: string, options?: ElementCreationOptions) =>
    tag.toLowerCase() === "canvas" ? createCanvas(1, 1) : createElement(tag, options);
  scope.document = document;
  scope.window = globalThis;
  scope.__pptHeadlessEnv = true;
}
