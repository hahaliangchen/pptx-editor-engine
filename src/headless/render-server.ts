import { IncomingMessage, Server, ServerResponse, createServer } from "node:http";
import { HeadlessSession } from "./session";

/**
 * A small HTTP service that draws a page of a .pptx as PNG: the backend of a
 * product posts the file it has just stored and gets a thumbnail back.
 *
 *   POST /thumbnail?width=480            body: the .pptx   → PNG of page 1
 *   POST /render?page=3&width=1280       body: the .pptx   → PNG of that page
 *   GET  /health                                           → { ok: true }
 *
 * The PNG's response carries `X-Slide-Count`. A failure is JSON:
 * { error: { code, message } } with 400 (not a deck, bad parameter), 413 (too
 * big), 503 (too many waiting) or 500.
 */
export interface RenderServerOptions {
  /** Largest accepted file, in bytes. Default 50 MB. */
  maxBodyBytes?: number;
  /** Requests allowed to wait for their turn; more get 503. Default 16. */
  maxQueue?: number;
  /** Default width of a thumbnail in pixels. Default 480. */
  thumbnailWidth?: number;
}

const WIDTH = { min: 64, max: 2560 };

class HttpError extends Error {
  constructor(public status: number, public code: string, message: string) {
    super(message);
  }
}

async function readBody(request: IncomingMessage, limit: number): Promise<Buffer> {
  const declared = Number(request.headers["content-length"] ?? 0);
  if (declared > limit) throw new HttpError(413, "TOO_LARGE", `文件超过 ${Math.round(limit / 1048576)} MB`);
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of request) {
    size += (chunk as Buffer).length;
    if (size > limit) throw new HttpError(413, "TOO_LARGE", `文件超过 ${Math.round(limit / 1048576)} MB`);
    chunks.push(chunk as Buffer);
  }
  return Buffer.concat(chunks);
}

function integer(value: string | null, fallback: number, name: string, min: number, max: number): number {
  if (value === null || value === "") return fallback;
  const number = Number(value);
  if (!Number.isInteger(number) || number < min || number > max) {
    throw new HttpError(400, "BAD_PARAMETER", `${name} 必须是 ${min} 到 ${max} 之间的整数`);
  }
  return number;
}

export function createRenderServer(options: RenderServerOptions = {}): Server {
  const maxBody = options.maxBodyBytes ?? 50 * 1048576;
  const maxQueue = options.maxQueue ?? 16;
  const thumbnailWidth = options.thumbnailWidth ?? 480;

  // A session sets up a document of its own in jsdom's globals, so decks are drawn one at a time.
  let tail: Promise<unknown> = Promise.resolve();
  let waiting = 0;
  const inTurn = <T>(job: () => Promise<T>): Promise<T> => {
    if (waiting >= maxQueue) return Promise.reject(new HttpError(503, "BUSY", "渲染队列已满，请稍后重试"));
    waiting++;
    const run = tail.then(job, job).finally(() => { waiting--; });
    tail = run.catch(() => undefined);
    return run;
  };

  async function draw(bytes: Buffer, page: number, width: number): Promise<{ png: Buffer; slides: number }> {
    let session: HeadlessSession;
    try {
      session = await HeadlessSession.open(bytes, { quiet: true });
    } catch (error) {
      throw new HttpError(400, "INVALID_PPTX", `不是可以打开的 PPTX 文件：${error instanceof Error ? error.message : String(error)}`);
    }
    try {
      const slides = session.document.slides.length;
      if (page > slides) throw new HttpError(400, "BAD_PARAMETER", `没有第 ${page} 页，这份文稿共 ${slides} 页`);
      return { png: await session.render(page, { width }), slides };
    } finally {
      session.dispose();
    }
  }

  async function handle(request: IncomingMessage, response: ServerResponse): Promise<void> {
    const url = new URL(request.url ?? "/", "http://localhost");
    if (request.method === "GET" && url.pathname === "/health") {
      response.writeHead(200, { "Content-Type": "application/json" }).end(JSON.stringify({ ok: true }));
      return;
    }
    const isThumbnail = url.pathname === "/thumbnail";
    if (request.method !== "POST" || (!isThumbnail && url.pathname !== "/render")) {
      throw new HttpError(404, "NOT_FOUND", "可用的接口：POST /thumbnail、POST /render、GET /health");
    }
    const page = isThumbnail ? 1 : integer(url.searchParams.get("page"), 1, "page", 1, 10000);
    const width = integer(url.searchParams.get("width"), isThumbnail ? thumbnailWidth : 1280, "width", WIDTH.min, WIDTH.max);
    const body = await readBody(request, maxBody);
    if (body.length === 0) throw new HttpError(400, "EMPTY_BODY", "请求体是空的：把 .pptx 文件的内容放在请求体里");
    const { png, slides } = await inTurn(() => draw(body, page, width));
    response.writeHead(200, { "Content-Type": "image/png", "Content-Length": png.length, "X-Slide-Count": String(slides) }).end(png);
  }

  return createServer((request, response) => {
    handle(request, response).catch(error => {
      const known = error instanceof HttpError;
      if (!known) console.error("渲染失败：", error);
      if (response.headersSent) return response.destroy();
      response.writeHead(known ? error.status : 500, { "Content-Type": "application/json" })
        .end(JSON.stringify({ error: { code: known ? error.code : "RENDER_FAILED", message: error instanceof Error ? error.message : String(error) } }));
    });
  });
}
