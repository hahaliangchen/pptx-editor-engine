import { describe, expect, it } from "vitest";
import { PageToolCall, PageToolLink, ToolCallTransport } from "../src/page/tool-link";
import { resolveOpenFamily } from "../src/page/font-map";
import { PAGE_TOOLS, MUTATING_TOOLS } from "../src/page/page-tools";
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
  it("delivers successful preflight without saving it", async () => {
    const {value,replies,wake}=transport([[{call_id:"check",name:"validate_html",arguments:{html:"<body>preview</body>"}}]]);
    let saves=0;
    const link=new PageToolLink(value,{run:async()=>ok(JSON.stringify({validated:true,applied:false,needsSave:false})),mutating:MUTATING_TOOLS,persist:async()=>{saves++;return 2;}});
    link.start();await new Promise(resolve=>setTimeout(resolve,10));const stopped=link.stop();wake();await stopped;
    expect(saves).toBe(0);expect(replies).toHaveLength(1);expect(replies[0].result.content).toHaveLength(1);
  });

  it('caps isolated page calls at three and does not persist an already acknowledged page batch again',async()=>{
    const calls=Array.from({length:11},(_,i)=>({call_id:String(i),name:'fill_layout',arguments:{page:i+1}}));
    const {value,replies,wake}=transport([calls]);let active=0,peak=0,saves=0;
    const link=new PageToolLink(value,{
      parallel:new Set(['fill_layout']),concurrency:99,mutating:new Set(['fill_layout']),
      run:async()=>{active++;peak=Math.max(peak,active);await new Promise(resolve=>setTimeout(resolve,1));active--;return ok(JSON.stringify({applied:true,saved:true,stored_revision:8}));},
      persist:async()=>{saves++;return 9;}
    });
    link.start();
    for(let i=0;i<100&&replies.length<11;i++)await new Promise(resolve=>setTimeout(resolve,1));
    const stopped=link.stop();wake();await stopped;
    expect(peak).toBe(3);expect(replies.length).toBe(11);expect(saves).toBe(0);
  });

  it('waits for concurrent page saves before a non-isolated operation reads the document',async()=>{
    const {value,replies,wake}=transport([[...Array.from({length:3},(_,i)=>({call_id:String(i),name:'fill_layout',arguments:{page:i+1}})),{call_id:'read',name:'get_design',arguments:{}}]]);
    let active=0;
    const link=new PageToolLink(value,{parallel:new Set(['fill_layout']),concurrency:3,mutating:new Set(),
      run:async name=>{if(name==='get_design')expect(active).toBe(0);else {active++;await new Promise(resolve=>setTimeout(resolve,2));active--;}return ok('{}');}});
    link.start();for(let i=0;i<100&&replies.length<4;i++)await new Promise(resolve=>setTimeout(resolve,1));
    const stopped=link.stop();wake();await stopped;expect(replies.at(-1)?.callId).toBe('read');
  });
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
    expect(replies[1].result.content.at(-1)).toEqual({ type: "text", text: JSON.stringify({ saved: true, stored_revision: 9 }) });
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
    expect(PAGE_TOOLS.map(t => t.name)).toEqual(["get_design", "validate_html", "apply_html", "get_layout", "fill_layout", "manage_design_pages", "get_outline", "render_slide", "check_layout", "suggest_annotations", "check_template", "get_template"]);
    expect(PAGE_TOOLS.every(t => t.executor === "page")).toBe(true);
  });
});
