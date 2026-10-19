/** Messages between `WorkerRenderer` and `render-worker.ts`. */
export interface RenderRequest {
  type: "render";
  id: number;
  /** `{ id, elements }` exactly as the main-thread renderer receives it. */
  slideJson: string;
  slideSize: { width: number; height: number };
  /** Output width in pixels. */
  width: number;
  /** Image URLs the slide draws; the worker fetches and decodes them. */
  images: string[];
}

export type WorkerReply =
  | { id: number; png: ArrayBuffer }
  | { id: number; error: string };
