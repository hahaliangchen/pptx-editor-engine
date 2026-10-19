import type { RenderRequest, WorkerReply } from "./render-protocol";

/**
 * The page's side of `render-worker.ts`: hands it fonts once and slides to
 * draw, and gets PNG bytes back. One worker serves the whole page.
 */
export class WorkerRenderer {
  private worker: Worker;
  private nextId = 1;
  private pending = new Map<number, { resolve: (png: Uint8Array) => void; reject: (error: Error) => void }>();
  private sentFonts = new WeakSet<Uint8Array>();

  constructor(worker?: Worker) {
    this.worker = worker ?? new Worker(new URL("./render-worker.ts", import.meta.url), { type: "module" });
    this.worker.onmessage = (event: MessageEvent<WorkerReply>) => {
      const reply = event.data;
      const waiting = this.pending.get(reply.id);
      if (!waiting) return;
      this.pending.delete(reply.id);
      if ("png" in reply) waiting.resolve(new Uint8Array(reply.png));
      else waiting.reject(new Error(reply.error));
    };
    this.worker.onerror = event => {
      const error = new Error(`渲染线程出错：${event.message}`);
      for (const waiting of this.pending.values()) waiting.reject(error);
      this.pending.clear();
    };
  }

  /** Register fonts the worker has not seen yet. The bytes are copied, not moved. */
  public addFonts(fonts: Uint8Array[]): void {
    for (const bytes of fonts) {
      if (this.sentFonts.has(bytes)) continue;
      this.sentFonts.add(bytes);
      this.worker.postMessage({ type: "font", bytes });
    }
  }

  /** A new document: image URLs of the old one are revoked. */
  public forgetImages(): void {
    this.worker.postMessage({ type: "forget-images" });
  }

  public render(request: Omit<RenderRequest, "type" | "id">): Promise<Uint8Array> {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.worker.postMessage({ type: "render", id, ...request } satisfies RenderRequest);
    });
  }

  public dispose(): void {
    this.worker.terminate();
    for (const waiting of this.pending.values()) waiting.reject(new Error("渲染线程已关闭"));
    this.pending.clear();
  }
}
