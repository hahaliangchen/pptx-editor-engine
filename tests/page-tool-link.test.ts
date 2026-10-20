import { describe, expect, it } from "vitest";
import { PageToolCall, PageToolLink, ToolCallTransport } from "../src/page/tool-link";
import { resolveOpenFamily } from "../src/page/font-map";
import { PAGE_TOOLS } from "../src/page/page-tools";
import type { ToolResult } from "../src/api/tools";

function transport(batches: PageToolCall[][]) {
  const replies: Array<{ callId: string; result: ToolResult }> = [];
  let waiting: (() => void) | null = null;
  const value: ToolCallTransport = {
    next: async () => {
      const batch = batches.shift();
      if (batch) return batch;
      await new Promise<void>(resolve => { waiting = resolve; });
      return [];
    },
    reply: async (callId, result) => { replies.push({ callId, result }); }
  };
  return { value, replies, wake: () => waiting?.() };
}

const ok = (text: string): ToolResult => ({ ok: true, content: [{ type: "text", text }] });

describe("PageToolLink", () => {
  it("answers calls in order and saves after a successful edit", async () => {
    const { value, replies, wake } = transport([[
      { call_id: "1", name: "get_outline", arguments: {} },
      { call_id: "2", name: "apply_edits", arguments: { commands: [] } }
    ]]);
    const order: string[] = [];
    const link = new PageToolLink(value, {
      run: async name => { order.push(name); return ok(name); },
      mutating: new Set(["apply_edits"]),
      persist: async () => { order.push("save"); return 9; }
    });
    link.start();
    await new Promise(resolve => setTimeout(resolve, 10));
    const stopped = link.stop();
    wake();
    await stopped;
    expect(order).toEqual(["get_outline", "apply_edits", "save"]);
    expect(replies.map(r => r.callId)).toEqual(["1", "2"]);
    expect(replies[1].result.content.at(-1)).toEqual({ type: "text", text: JSON.stringify({ saved: true, revision: 9 }) });
  });

  it("does not save after a failed edit, and turns a thrown error into a failed result", async () => {
    const { value, replies, wake } = transport([[
      { call_id: "a", name: "apply_edits", arguments: {} },
      { call_id: "b", name: "build_deck", arguments: {} }
    ]]);
    let saves = 0;
    const link = new PageToolLink(value, {
      run: async name => { if (name === "build_deck") throw new Error("boom"); return { ok: false, content: [] }; },
      mutating: new Set(["apply_edits", "build_deck"]),
      persist: async () => { saves++; return 1; }
    });
    link.start();
    await new Promise(resolve => setTimeout(resolve, 10));
    const stopped = link.stop();
    wake();
    await stopped;
    expect(saves).toBe(0);
    expect(replies[1].result.ok).toBe(false);
    expect(replies[1].result.content[0]).toMatchObject({ type: "text" });
    expect((replies[1].result.content[0] as { text: string }).text).toContain("boom");
  });
});

describe("open font policy", () => {
  const catalog = ["Noto Sans CJK SC", "Noto Serif CJK SC", "Liberation Sans", "Carlito"];

  it("draws Office fonts with their open-source stand-ins", () => {
    expect(resolveOpenFamily("Microsoft YaHei", true, catalog)).toBe("Noto Sans CJK SC");
    expect(resolveOpenFamily("宋体", true, catalog)).toBe("Noto Serif CJK SC");
    expect(resolveOpenFamily("Calibri", false, catalog)).toBe("Carlito");
    expect(resolveOpenFamily("Arial", false, catalog)).toBe("Liberation Sans");
  });

  it("falls back by script when nothing matches", () => {
    expect(resolveOpenFamily("Some Brand Font", true, catalog)).toBe("Noto Sans CJK SC");
    expect(resolveOpenFamily("Some Brand Font", false, catalog)).toBe("Carlito");
    expect(resolveOpenFamily("sans-serif", false, ["Liberation Sans"])).toBe("Liberation Sans");
  });

  it("marks every deck tool as executed by the page", () => {
    expect(PAGE_TOOLS.map(t => t.name)).toEqual(["get_outline", "apply_edits", "render_slide", "check_layout", "suggest_annotations", "check_template", "get_template", "build_deck"]);
    expect(PAGE_TOOLS.every(t => t.executor === "page")).toBe(true);
  });
});
