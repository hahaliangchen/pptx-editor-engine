import { describe, expect, it } from "vitest";
import { Command, History } from "../src/editor/history";

const command = (label: string, log: string[] = [], extra: Partial<Command> = {}): Command => ({
  label,
  apply: () => { log.push(`+${label}`); },
  revert: () => { log.push(`-${label}`); },
  ...extra
});

describe("History", () => {
  it("lists steps in the order they were made, undone ones after the done ones", () => {
    const history = new History();
    history.record(command("a"));
    history.record(command("b"));
    history.record(command("c"));
    history.undo();
    expect(history.entries().map(e => e.label)).toEqual(["a", "b", "c"]);
    expect(history.position).toBe(2);
    history.undo();
    expect(history.entries().map(e => e.label)).toEqual(["a", "b", "c"]);
    expect(history.position).toBe(1);
  });

  it("jumps to any point by undoing and redoing in order", () => {
    const log: string[] = [];
    const history = new History();
    ["a", "b", "c", "d"].forEach(label => history.record(command(label, log)));
    history.goTo(1);
    expect(log).toEqual(["-d", "-c", "-b"]);
    log.length = 0;
    history.goTo(3);
    expect(log).toEqual(["+b", "+c"]);
    history.goTo(99);
    expect(history.position).toBe(4);
    history.goTo(-5);
    expect(history.position).toBe(0);
  });

  it("a new step after going back drops the undone ones", () => {
    const history = new History();
    ["a", "b", "c"].forEach(label => history.record(command(label)));
    history.goTo(1);
    history.record(command("x"));
    expect(history.entries().map(e => e.label)).toEqual(["a", "x"]);
    expect(history.canRedo).toBe(false);
  });

  it("tracks whether the document differs from the saved state", () => {
    const history = new History();
    expect(history.isDirty).toBe(false);
    history.record(command("a"));
    expect(history.isDirty).toBe(true);
    history.markSaved();
    expect(history.isDirty).toBe(false);
    history.record(command("b"));
    expect(history.isDirty).toBe(true);
    history.undo();
    expect(history.isDirty).toBe(false); // back at the saved point
    history.undo();
    expect(history.isDirty).toBe(true); // before the saved point
    history.redo();
    expect(history.isDirty).toBe(false);
  });

  it("a step merged into the saved one makes the document dirty again", () => {
    const history = new History();
    const merging = (): Command => command("move", [], { mergeKey: "k", merge: () => undefined });
    history.record(merging());
    history.markSaved();
    expect(history.isDirty).toBe(false);
    history.record(merging());
    expect(history.entries()).toHaveLength(1);
    expect(history.isDirty).toBe(true);
  });

  it("is clean again after clear, and undo-to-start after a save is dirty", () => {
    const history = new History();
    history.record(command("a"));
    history.markSaved();
    history.undo();
    expect(history.isDirty).toBe(true);
    history.clear();
    expect(history.isDirty).toBe(false);
  });
});

describe("relativeTime", () => {
  it("speaks in the coarsest unit that fits", async () => {
    const { relativeTime } = await import("../src/editor/history-panel");
    const now = 1_000_000_000_000;
    expect(relativeTime(now - 1000, now)).toBe("刚刚");
    expect(relativeTime(now - 30_000, now)).toBe("30 秒前");
    expect(relativeTime(now - 5 * 60_000, now)).toBe("5 分钟前");
    expect(relativeTime(now - 3 * 3_600_000, now)).toBe("3 小时前");
    expect(relativeTime(now - 2 * 86_400_000, now)).toBe("2 天前");
  });
});
