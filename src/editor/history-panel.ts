import { HistoryEntry } from "./history";

const CLOCK = '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 3-6.7"/><path d="M3 4v4h4"/><path d="M12 8v4l3 2"/></svg>';

/** "just now", "12 s ago" ... in Chinese, for the list. */
export function relativeTime(at: number, now = Date.now()): string {
  const seconds = Math.max(0, Math.round((now - at) / 1000));
  if (seconds < 5) return "刚刚";
  if (seconds < 60) return `${seconds} 秒前`;
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) return `${minutes} 分钟前`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours} 小时前`;
  return `${Math.round(hours / 24)} 天前`;
}

/**
 * Popover listing every step, newest first. The current point is marked, steps
 * after it (undone) are dimmed, and clicking any row jumps there.
 */
export class HistoryPanel {
  private panel = document.createElement("div");
  private list = document.createElement("ol");
  private count = document.createElement("span");
  private entries: HistoryEntry[] = [];
  private position = 0;
  private timer = 0;

  constructor(private button: HTMLButtonElement, private onJump: (position: number) => void) {
    this.button.innerHTML = CLOCK;

    this.panel.className = "history-panel";
    this.panel.hidden = true;
    const header = document.createElement("div");
    header.className = "history-header";
    const title = document.createElement("strong");
    title.textContent = "历史记录";
    this.count.className = "history-count";
    header.append(title, this.count);
    this.list.className = "history-list";
    this.panel.append(header, this.list);
    document.body.appendChild(this.panel);

    this.button.addEventListener("click", event => {
      event.stopPropagation();
      this.panel.hidden ? this.open() : this.close();
    });
    document.addEventListener("pointerdown", event => {
      if (this.panel.hidden) return;
      const target = event.target as Node;
      if (!this.panel.contains(target) && !this.button.contains(target)) this.close();
    }, true);
    document.addEventListener("keydown", event => {
      if (event.key === "Escape" && !this.panel.hidden) this.close();
    });
  }

  public update(entries: HistoryEntry[], position: number): void {
    this.entries = entries;
    this.position = position;
    this.button.disabled = false;
    this.button.classList.toggle("has-history", entries.length > 0);
    if (!this.panel.hidden) this.render();
  }

  public open(): void {
    this.panel.hidden = false;
    const box = this.button.getBoundingClientRect();
    this.panel.style.top = `${box.bottom + 8}px`;
    this.panel.style.left = `${Math.max(8, Math.min(box.left + box.width / 2 - 140, window.innerWidth - 288))}px`;
    this.render();
    // Keep the "x s ago" labels honest while the list is open.
    this.timer = window.setInterval(() => this.render(), 15000);
  }

  public close(): void {
    this.panel.hidden = true;
    window.clearInterval(this.timer);
  }

  private render(): void {
    this.count.textContent = this.entries.length ? `${this.position} / ${this.entries.length} 步` : "";
    const now = Date.now();
    const rows: HTMLElement[] = [];

    // Newest on top; the start of the document is the last row.
    for (let index = this.entries.length - 1; index >= 0; index--) {
      const entry = this.entries[index];
      const row = this.row(entry.label, relativeTime(entry.at, now), index + 1);
      rows.push(row);
    }
    rows.push(this.row("初始状态", "", 0));
    if (this.entries.length === 0) {
      const empty = document.createElement("li");
      empty.className = "history-empty";
      empty.textContent = "还没有任何操作";
      rows.unshift(empty);
    }
    this.list.replaceChildren(...rows);
    this.list.querySelector(".is-current")?.scrollIntoView({ block: "nearest" });
  }

  private row(label: string, time: string, target: number): HTMLElement {
    const row = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.className = "history-row";
    if (target === this.position) button.classList.add("is-current");
    else if (target > this.position) button.classList.add("is-undone");
    const dot = document.createElement("span");
    dot.className = "history-dot";
    const name = document.createElement("span");
    name.className = "history-label";
    name.textContent = label;
    const when = document.createElement("span");
    when.className = "history-time";
    when.textContent = time;
    button.append(dot, name, when);
    button.addEventListener("click", () => this.onJump(target));
    row.appendChild(button);
    return row;
  }
}
