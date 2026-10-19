import type { ToolResult } from "../api/tools";

/** One call the Agent host handed to this page. */
export interface PageToolCall {
  call_id: string;
  name: string;
  arguments: unknown;
}

/**
 * How the page talks to the Agent host. Plain methods so any bridge fits: an
 * HTTP API with the user's login, an iframe message bridge, a test double.
 */
export interface ToolCallTransport {
  /** Calls waiting for this page; may hold the request up to `waitMs` when none are. */
  next(waitMs: number): Promise<PageToolCall[]>;
  reply(callId: string, result: ToolResult): Promise<void>;
}

export interface PageToolLinkOptions {
  /** Run a tool against the open deck. */
  run(name: string, args: unknown): Promise<ToolResult>;
  /** Tools that change the deck; after one succeeds `persist` runs before the reply. */
  mutating: Set<string>;
  /** Store the deck and return its revision; the reply tells the Agent that revision. */
  persist?(): Promise<number>;
  /** How long one `next` may wait for work. */
  waitMs?: number;
  onError?(error: Error): void;
  onCall?(call: PageToolCall, phase: "start" | "end", result?: ToolResult): void;
}

/**
 * Takes the Agent's tool calls for this page and answers them, one at a time
 * and in order, for as long as the page is open. A closed page simply stops
 * asking; the host then gives up the task.
 */
export class PageToolLink {
  private running = false;
  private stopped: Promise<void> | null = null;

  constructor(private transport: ToolCallTransport, private options: PageToolLinkOptions) {}

  public start(): void {
    if (this.running) return;
    this.running = true;
    this.stopped = this.loop();
  }

  public async stop(): Promise<void> {
    this.running = false;
    await this.stopped;
  }

  private async loop(): Promise<void> {
    let failures = 0;
    while (this.running) {
      let calls: PageToolCall[];
      try {
        calls = await this.transport.next(this.options.waitMs ?? 20000);
        failures = 0;
      } catch (error) {
        this.options.onError?.(error instanceof Error ? error : new Error(String(error)));
        failures++;
        // Network or host restart: back off, but keep the page reachable.
        await new Promise(resolve => setTimeout(resolve, Math.min(10000, 500 * 2 ** Math.min(failures, 5))));
        continue;
      }
      for (const call of calls) {
        if (!this.running) return;
        await this.answer(call);
      }
    }
  }

  private async answer(call: PageToolCall): Promise<void> {
    this.options.onCall?.(call, "start");
    let result: ToolResult;
    try {
      result = await this.options.run(call.name, call.arguments);
      if (result.ok && this.options.mutating.has(call.name) && this.options.persist) {
        const revision = await this.options.persist();
        result = { ...result, content: [...result.content, { type: "text", text: JSON.stringify({ saved: true, revision }) }] };
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      result = { ok: false, content: [{ type: "text", text: JSON.stringify({ ok: false, error: { message } }) }] };
    }
    this.options.onCall?.(call, "end", result);
    try {
      await this.transport.reply(call.call_id, result);
    } catch (error) {
      this.options.onError?.(error instanceof Error ? error : new Error(String(error)));
    }
  }
}
