import { JSDOM } from "jsdom";
import { Canvas, DOMMatrix, Image, ImageData, createCanvas } from "@napi-rs/canvas";

/**
 * The Rust renderer is written against browser classes (`web_sys`): it looks up
 * `window.document`, creates canvases through `createElement("canvas")` and
 * type-checks values with `instanceof`. This gives it Skia-backed stand-ins so
 * the same WASM runs in Node. The XML parser needs a `DOMParser`.
 */
export function installNodeEnv(): void {
  const scope = globalThis as any;
  if (scope.__pptHeadlessEnv) return;
  if (typeof scope.document !== "undefined") {
    throw new Error("headless 渲染需要在 Node 环境运行：当前已存在浏览器 document。");
  }

  if (typeof scope.DOMParser === "undefined") scope.DOMParser = new JSDOM("").window.DOMParser;

  class NodeWindow {
    static [Symbol.hasInstance](value: unknown) { return value === globalThis; }
  }
  scope.Window = NodeWindow;
  scope.window = globalThis;
  scope.document = { createElement: () => createCanvas(1, 1) };
  scope.HTMLCanvasElement = Canvas;
  scope.HTMLImageElement = Image;
  scope.CanvasRenderingContext2D = createCanvas(1, 1).getContext("2d").constructor;
  scope.ImageData = ImageData;
  scope.DOMMatrix = DOMMatrix;
  scope.__pptHeadlessEnv = true;
}
