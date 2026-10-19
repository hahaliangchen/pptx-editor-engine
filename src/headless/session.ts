import { loadImage } from "@napi-rs/canvas";
import { ApplyResult, EditApi } from "../api/engine";
import { OutlineOptions } from "../api/outline";
import { DocumentSession } from "../document-session";
import { PptVirtualDocument } from "../pptx-virtual-dom";
import { HeadlessRenderOptions, HeadlessRenderer, RenderSlideOptions } from "./renderer";
import { ToolResult, callTool } from "../api/tools";
import { installNodeEnv } from "./node-env";

/**
 * A deck opened without a browser: read its structure, change it with JSON
 * commands, draw any page to PNG and write the result back as a .pptx. The same
 * `DocumentSession` the browser editor uses does the work.
 */
export class HeadlessSession {
  private constructor(
    private session: DocumentSession,
    private renderer: HeadlessRenderer
  ) {}

  public static async open(buffer: ArrayBuffer | Uint8Array, options: HeadlessRenderOptions = {}): Promise<HeadlessSession> {
    installNodeEnv();
    const renderer = new HeadlessRenderer(options);
    const container = document.createElement("div");
    document.body.appendChild(container);
    const session: DocumentSession = new DocumentSession({
      container,
      view: {
        invalidate: () => undefined,
        beginDrag: () => undefined,
        endDrag: () => undefined,
        slideEdited: () => undefined,
        beginTextEdit: () => undefined,
        endTextEdit: () => undefined,
        slidesChanged: async () => undefined,
        currentChanged: async () => undefined,
        prepareSlide: async (slide): Promise<void> => renderer.prepare({ ...session.presentation!, slides: [slide] })
      },
      probeImage: async bytes => {
        const image = await loadImage(Buffer.from(bytes));
        return { width: image.width, height: image.height };
      }
    });
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const doc = await session.open(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
    renderer.prepare(doc);
    return new HeadlessSession(session, renderer);
  }

  /** The deck as the editor sees it right now. */
  public get document(): PptVirtualDocument {
    return this.session.presentation!;
  }

  public get api(): EditApi {
    return this.session.api;
  }

  public outline(options?: OutlineOptions) {
    return this.api.outline(options);
  }

  public apply(commands: unknown): Promise<ApplyResult> {
    return this.api.apply(commands);
  }

  /** Run one of the tools (`get_outline`, `apply_edits`, `render_slide`) on this deck. */
  public callTool(name: string, args: unknown): Promise<ToolResult> {
    return callTool(this, name, args);
  }

  public undo(): void {
    this.session.editor.undo();
  }

  public redo(): void {
    this.session.editor.redo();
  }

  /** Page `page` (1-based) as PNG. */
  public async render(page: number, options?: RenderSlideOptions): Promise<Buffer> {
    const doc = this.document;
    this.renderer.prepare(doc);
    return this.renderer.renderSlide(doc, page - 1, options);
  }

  /** The deck with every edit written back into the original package. */
  public async export(): Promise<Buffer> {
    const blob = await this.session.exportPptx();
    return Buffer.from(await blob.arrayBuffer());
  }

  public dispose(): void {
    this.session.dispose();
  }
}
