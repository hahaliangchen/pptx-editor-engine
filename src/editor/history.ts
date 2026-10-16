/** A reversible edit. `apply` must be safe to call again after `revert`. */
export interface Command {
  /** Short name shown in the history list. */
  label: string;
  apply(): void;
  revert(): void;
  /** Commands with the same key pushed in quick succession collapse into one undo step. */
  mergeKey?: string;
  /** Fold a later command into this one; called only when mergeKeys match. */
  merge?(next: Command): void;
}

export interface HistoryEntry {
  label: string;
  /** Wall-clock time the step was made. */
  at: number;
}

const MERGE_WINDOW_MS = 600;

interface Step {
  command: Command;
  at: number;
  /** Changes whenever the step changes (a merge counts), so "saved" can be told apart. */
  token: number;
}

export class History {
  private undoStack: Step[] = [];
  private redoStack: Step[] = [];
  private lastPushAt = 0;
  private nextToken = 1;
  /** Token of the top step when the document was last saved; 0 for the untouched start. */
  private savedToken = 0;
  private listeners = new Set<() => void>();

  /** Record a command whose effect has already been applied. */
  public record(command: Command): void {
    const top = this.undoStack[this.undoStack.length - 1];
    const now = performance.now();
    if (
      top && command.mergeKey && top.command.mergeKey === command.mergeKey && top.command.merge
      && now - this.lastPushAt < MERGE_WINDOW_MS
    ) {
      top.command.merge!(command);
      top.token = this.nextToken++;
    } else {
      this.undoStack.push({ command, at: Date.now(), token: this.nextToken++ });
    }
    this.lastPushAt = now;
    this.redoStack = [];
    this.emit();
  }

  public undo(): Command | null {
    const step = this.undoStack.pop();
    if (!step) return null;
    step.command.revert();
    this.redoStack.push(step);
    this.lastPushAt = 0;
    this.emit();
    return step.command;
  }

  public redo(): Command | null {
    const step = this.redoStack.pop();
    if (!step) return null;
    step.command.apply();
    this.undoStack.push(step);
    this.lastPushAt = 0;
    this.emit();
    return step.command;
  }

  public get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  public get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** Every step in the order it was made: the done ones, then the undone ones. */
  public entries(): HistoryEntry[] {
    return [...this.undoStack, ...[...this.redoStack].reverse()].map(step => ({
      label: step.command.label,
      at: step.at
    }));
  }

  /** How many steps are currently applied. */
  public get position(): number {
    return this.undoStack.length;
  }

  /** Undo or redo until exactly `position` steps are applied. */
  public goTo(position: number): void {
    const target = Math.max(0, Math.min(position, this.undoStack.length + this.redoStack.length));
    while (this.undoStack.length > target) this.undo();
    while (this.undoStack.length < target) this.redo();
  }

  /** The document is saved as it is now; undoing back to this point makes it clean again. */
  public markSaved(): void {
    const top = this.undoStack[this.undoStack.length - 1];
    this.savedToken = top ? top.token : 0;
    this.emit();
  }

  public get isDirty(): boolean {
    const top = this.undoStack[this.undoStack.length - 1];
    return (top ? top.token : 0) !== this.savedToken;
  }

  public clear(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.savedToken = 0;
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
