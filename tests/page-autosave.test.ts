import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AutoSaver, DeckStore, RevisionConflict, SaveStatus } from "../src/page/autosave";

class FakeDeck {
  dirty = false;
  exports = 0;
  async exportPptx(): Promise<Blob> {
    this.exports++;
    this.dirty = false; // the editor marks itself saved on export
    return new Blob([`deck ${this.exports}`]);
  }
  hasUnsavedEdits(): boolean { return this.dirty; }
}

function store(behaviour: (expected: number) => Promise<number>) {
  const calls: Array<{ expected: number; source: string }> = [];
  const value: DeckStore = {
    save: (_pptx, expected, source) => { calls.push({ expected, source }); return behaviour(expected); }
  };
  return { value, calls };
}

describe("AutoSaver", () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  it("saves user edits once they pause, building on the last revision", async () => {
    const deck = new FakeDeck();
    const { value, calls } = store(async expected => expected + 1);
    const statuses: SaveStatus[] = [];
    const saver = new AutoSaver(deck, value, { revision: 3, idleMs: 1000, onStatus: s => statuses.push(s) });

    deck.dirty = true;
    saver.notifyEdited();
    await vi.advanceTimersByTimeAsync(500);
    saver.notifyEdited(); // still typing
    await vi.advanceTimersByTimeAsync(900);
    expect(calls).toHaveLength(0);
    await vi.advanceTimersByTimeAsync(200);
    expect(calls).toEqual([{ expected: 3, source: "user" }]);
    expect(saver.currentRevision).toBe(4);
    expect(statuses.at(-1)).toBe("saved");
  });

  it("saves Agent edits at once and returns the new revision", async () => {
    const deck = new FakeDeck();
    const { value, calls } = store(async expected => expected + 1);
    const saver = new AutoSaver(deck, value, { revision: 1 });
    deck.dirty = true;
    await expect(saver.saveNow("agent")).resolves.toBe(2);
    expect(calls).toEqual([{ expected: 1, source: "agent" }]);
    // Nothing changed since: no upload, same revision.
    await expect(saver.saveNow("agent")).resolves.toBe(2);
    expect(calls).toHaveLength(1);
  });

  it("runs one save at a time and saves again for edits made meanwhile", async () => {
    const deck = new FakeDeck();
    let release!: () => void;
    const { value, calls } = store(expected => new Promise(resolve => { release = () => resolve(expected + 1); }));
    const saver = new AutoSaver(deck, value, { revision: 0 });
    deck.dirty = true;
    const first = saver.saveNow("user");
    await vi.advanceTimersByTimeAsync(0);
    deck.dirty = true;
    const second = saver.saveNow("agent");
    const third = saver.saveNow("user");
    expect(calls).toHaveLength(1);
    release();
    await expect(first).resolves.toBe(1);
    await vi.advanceTimersByTimeAsync(0);
    expect(calls).toEqual([{ expected: 0, source: "user" }, { expected: 1, source: "agent" }]);
    release();
    await expect(second).resolves.toBe(2);
    await expect(third).resolves.toBe(2);
  });

  it("keeps the change after a failed upload and retries it", async () => {
    const deck = new FakeDeck();
    let fail = true;
    const { value, calls } = store(async expected => { if (fail) throw new Error("network"); return expected + 1; });
    const saver = new AutoSaver(deck, value, { revision: 5, retryMs: 1000 });
    deck.dirty = true;
    await expect(saver.saveNow("user")).rejects.toThrow("network");
    expect(saver.currentStatus).toBe("error");
    expect(saver.hasPendingChanges()).toBe(true);
    fail = false;
    await vi.advanceTimersByTimeAsync(1000);
    expect(calls.map(c => c.expected)).toEqual([5, 5]);
    expect(saver.currentRevision).toBe(6);
    expect(saver.hasPendingChanges()).toBe(false);
  });

  it("stops on a conflict until the page decides", async () => {
    const deck = new FakeDeck();
    const { value, calls } = store(async expected => { if (expected === 2) throw new RevisionConflict(7); return expected + 1; });
    const saver = new AutoSaver(deck, value, { revision: 2, idleMs: 100 });
    deck.dirty = true;
    await expect(saver.saveNow("user")).rejects.toBeInstanceOf(RevisionConflict);
    expect(saver.currentStatus).toBe("conflict");
    deck.dirty = true;
    saver.notifyEdited();
    await vi.advanceTimersByTimeAsync(500);
    expect(calls).toHaveLength(1);
    await expect(saver.resolveConflict(7, true)).resolves.toBe(8);
    expect(calls.at(-1)).toEqual({ expected: 7, source: "user" });
  });
});
