/** A reversible edit. `apply` must be safe to call again after `revert`. */
export interface Command {
  label: string;
  apply(): void;
  revert(): void;
  /** Commands with the same key pushed in quick succession collapse into one undo step. */
  mergeKey?: string;
  /** Fold a later command into this one; called only when mergeKeys match. */
  merge?(next: Command): void;
}

const MERGE_WINDOW_MS = 600;

export class History {
  private undoStack: Command[] = [];
  private redoStack: Command[] = [];
  private lastPushAt = 0;
  private listeners = new Set<() => void>();

  /** Record a command whose effect has already been applied. */
  public record(command: Command): void {
    const top = this.undoStack[this.undoStack.length - 1];
    const now = performance.now();
    if (
      top && command.mergeKey && top.mergeKey === command.mergeKey && top.merge
      && now - this.lastPushAt < MERGE_WINDOW_MS
    ) {
      top.merge(command);
    } else {
      this.undoStack.push(command);
    }
    this.lastPushAt = now;
    this.redoStack = [];
    this.emit();
  }

  public undo(): Command | null {
    const command = this.undoStack.pop();
    if (!command) return null;
    command.revert();
    this.redoStack.push(command);
    this.lastPushAt = 0;
    this.emit();
    return command;
  }

  public redo(): Command | null {
    const command = this.redoStack.pop();
    if (!command) return null;
    command.apply();
    this.undoStack.push(command);
    this.lastPushAt = 0;
    this.emit();
    return command;
  }

  public get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  public get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  public clear(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.emit();
  }

  public subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    this.listeners.forEach(listener => listener());
  }
}
