/** Where the page keeps the deck: the plugin's own backend. */
export interface DeckStore {
  /**
   * Store the submitted PPTX and return its saved sequence number.
   * expectedRevision is retained only for older store implementations.
   */
  save(pptx: Blob, expectedRevision: number, source: SaveSource): Promise<number>;
}

export type SaveSource = "agent" | "user";

export type SaveStatus = "saved" | "dirty" | "saving" | "error";

/** What the saver needs from the editor. `PptViewer` has both. */
export interface SavableDeck {
  exportPptx(): Promise<Blob>;
  hasUnsavedEdits(): boolean;
}

export interface AutoSaverOptions {
  /** The revision the page opened. */
  revision: number;
  /** Quiet time after the user's last edit before saving. */
  idleMs?: number;
  /** Wait before retrying a failed save. */
  retryMs?: number;
  onStatus?: (status: SaveStatus, detail: { revision: number; error?: Error }) => void;
}

/**
 * Saves the deck to its store without the user asking: right away after an
 * Agent edit (so the tool result can carry the new revision), and after the
 * user has paused for `idleMs`. One save runs at a time; a request made while
 * one is running waits for it and then saves what has changed since. Every
 * save accepts the submitted contents without version conflict detection.
 */
export class AutoSaver {
  private revision: number;
  private status: SaveStatus = "saved";
  private timer: ReturnType<typeof setTimeout> | undefined;
  private running: Promise<number> | null = null;
  private queued: { source: SaveSource; promise: Promise<number> } | null = null;
  /** The editor marks itself clean when it exports, even if the upload then fails. */
  private unsaved = false;
  private disposed = false;
  private readonly idleMs: number;
  private readonly retryMs: number;

  constructor(private deck: SavableDeck, private store: DeckStore, private options: AutoSaverOptions) {
    this.revision = options.revision;
    this.idleMs = options.idleMs ?? 2500;
    this.retryMs = options.retryMs ?? 5000;
  }

  public get currentRevision(): number {
    return this.revision;
  }

  public get currentStatus(): SaveStatus {
    return this.status;
  }

  /** True while something is not stored yet (for a "leave page?" prompt). */
  public hasPendingChanges(): boolean {
    return this.unsaved || this.deck.hasUnsavedEdits() || this.status === "saving" || this.timer !== undefined;
  }

  /** The user changed something; save once they pause. */
  public notifyEdited(): void {
    if (this.disposed) return;
    if (!this.deck.hasUnsavedEdits() && !this.unsaved) return;
    if (this.status !== "saving") this.setStatus("dirty");
    this.schedule(this.idleMs);
  }

  /** Save now (Agent edits, the "save now" button) and return the stored revision. */
  public saveNow(source: SaveSource): Promise<number> {
    this.cancelTimer();
    if (!this.running) {
      this.running = this.run(source).finally(() => { this.running = null; });
      return this.running;
    }
    // One is running: save again after it, once, for everyone who asks meanwhile.
    if (!this.queued) {
      const after = this.running.catch(() => undefined).then(() => {
        const queued = this.queued!;
        this.queued = null;
        return this.saveNow(queued.source);
      });
      this.queued = { source, promise: after };
    } else if (source === "agent") {
      this.queued.source = "agent";
    }
    return this.queued.promise;
  }

  public dispose(): void {
    this.disposed = true;
    this.cancelTimer();
  }

  private async run(source: SaveSource): Promise<number> {
    if (!this.deck.hasUnsavedEdits() && !this.unsaved) {
      this.setStatus("saved");
      return this.revision;
    }
    this.setStatus("saving");
    let pptx: Blob;
    try {
      pptx = await this.deck.exportPptx();
      this.unsaved = true;
      const revision = await this.store.save(pptx, this.revision, source);
      this.revision = revision;
      this.unsaved = false;
      // Edits made while uploading still need a save.
      if (this.deck.hasUnsavedEdits()) {
        this.setStatus("dirty");
        this.schedule(this.idleMs);
      } else {
        this.setStatus("saved");
      }
      return revision;
    } catch (error) {
      const failure = error instanceof Error ? error : new Error(String(error));
      this.setStatus("error", failure);
      if (!this.disposed) this.schedule(this.retryMs);
      throw failure;
    }
  }

  private schedule(ms: number): void {
    this.cancelTimer();
    this.timer = setTimeout(() => {
      this.timer = undefined;
      this.saveNow("user").catch(() => undefined);
    }, ms);
  }

  private cancelTimer(): void {
    if (this.timer !== undefined) clearTimeout(this.timer);
    this.timer = undefined;
  }

  private setStatus(status: SaveStatus, error?: Error): void {
    this.status = status;
    this.options.onStatus?.(status, { revision: this.revision, error });
  }
}
