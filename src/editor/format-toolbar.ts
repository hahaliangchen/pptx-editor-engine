import { TextFormat } from "./text-ops";

/** What the current selection looks like; null means "mixed". */
export interface FormatState {
  bold: boolean | null;
  italic: boolean | null;
  color: string | null;
  fontSizePt: number | null;
}

export interface ToolbarAnchor {
  x: number;
  y: number;
  w: number;
  h: number;
}

const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 80, 96];
const PALETTE = [
  "#000000", "#434343", "#7f7f7f", "#bfbfbf", "#ffffff", "#e60012", "#c00000", "#ff6a00",
  "#ffc000", "#ffe600", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0"
];

const ICONS = {
  bold: '<svg viewBox="0 0 24 24"><path d="M7 5h6a3.5 3.5 0 0 1 0 7H7zM7 12h7a3.5 3.5 0 0 1 0 7H7z"/></svg>',
  italic: '<svg viewBox="0 0 24 24"><path d="M10 5h8M6 19h8M14.5 5 9.5 19"/></svg>',
  minus: '<svg viewBox="0 0 24 24"><path d="M6 12h12"/></svg>',
  plus: '<svg viewBox="0 0 24 24"><path d="M12 6v12M6 12h12"/></svg>'
};

/**
 * Floating pill with bold / italic / size / colour, positioned over the
 * current text selection. Every control swallows mousedown so pressing it never
 * steals focus (and therefore the caret/selection) from the text editor.
 */
export class FormatToolbar {
  private root = document.createElement("div");
  private bold: HTMLButtonElement;
  private italic: HTMLButtonElement;
  private size: HTMLSpanElement;
  private swatch: HTMLSpanElement;
  private popover = document.createElement("div");
  private state: FormatState = { bold: false, italic: false, color: null, fontSizePt: null };

  constructor(parent: HTMLElement, private onFormat: (format: TextFormat) => void) {
    this.root.className = "format-toolbar";
    this.root.addEventListener("mousedown", event => event.preventDefault());
    this.root.addEventListener("pointerdown", event => event.stopPropagation());

    this.bold = this.button("bold", "加粗 (Ctrl+B)", ICONS.bold, () =>
      this.onFormat({ bold: this.state.bold !== true }));
    this.italic = this.button("italic", "斜体 (Ctrl+I)", ICONS.italic, () =>
      this.onFormat({ italic: this.state.italic !== true }));

    const smaller = this.button("step", "减小字号", ICONS.minus, () => this.step(-1));
    this.size = document.createElement("span");
    this.size.className = "format-size";
    const larger = this.button("step", "增大字号", ICONS.plus, () => this.step(1));

    const colorButton = this.button("color", "文字颜色", "", () => this.popover.classList.toggle("is-open"));
    this.swatch = document.createElement("span");
    this.swatch.className = "format-swatch";
    colorButton.innerHTML = '<span class="format-a">A</span>';
    colorButton.appendChild(this.swatch);

    this.popover.className = "format-palette";
    for (const color of PALETTE) {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "format-chip";
      chip.style.background = color;
      chip.title = color.toUpperCase();
      chip.addEventListener("click", () => {
        this.popover.classList.remove("is-open");
        this.onFormat({ color });
      });
      this.popover.appendChild(chip);
    }
    colorButton.appendChild(this.popover);

    const divider = () => {
      const node = document.createElement("span");
      node.className = "format-divider";
      return node;
    };
    this.root.append(this.bold, this.italic, divider(), smaller, this.size, larger, divider(), colorButton);
    parent.appendChild(this.root);
  }

  public get element(): HTMLElement {
    return this.root;
  }

  public hide(): void {
    this.root.classList.remove("is-visible");
    this.popover.classList.remove("is-open");
  }

  /** `anchor` and `bounds` are in the same CSS-pixel space as the parent. */
  public show(anchor: ToolbarAnchor, state: FormatState, bounds: { w: number; h: number }): void {
    this.state = state;
    this.bold.classList.toggle("is-active", state.bold === true);
    this.italic.classList.toggle("is-active", state.italic === true);
    this.size.textContent = state.fontSizePt === null ? "—" : String(Math.round(state.fontSizePt * 10) / 10);
    this.swatch.style.background = state.color || "transparent";
    this.swatch.classList.toggle("is-mixed", state.color === null);

    this.root.classList.add("is-visible");
    const width = this.root.offsetWidth || 190;
    const height = this.root.offsetHeight || 38;
    let top = anchor.y - height - 12;
    this.root.classList.toggle("is-below", top < -height);
    if (top < -height) top = anchor.y + anchor.h + 12;
    const left = Math.min(Math.max(anchor.x + anchor.w / 2 - width / 2, 4 - 0), Math.max(4, bounds.w - width - 4));
    this.root.style.left = `${left}px`;
    this.root.style.top = `${Math.min(top, bounds.h + 40)}px`;
  }

  private step(direction: 1 | -1): void {
    const current = this.state.fontSizePt ?? 18;
    const candidates = direction > 0 ? SIZES.filter(size => size > current + 0.01) : SIZES.filter(size => size < current - 0.01).reverse();
    const next = candidates[0];
    if (next !== undefined) this.onFormat({ fontSizePt: next });
  }

  private button(kind: string, title: string, icon: string, onClick: () => void): HTMLButtonElement {
    const node = document.createElement("button");
    node.type = "button";
    node.className = `format-button format-${kind}`;
    node.title = title;
    node.setAttribute("aria-label", title);
    node.innerHTML = icon;
    node.addEventListener("click", event => {
      event.stopPropagation();
      onClick();
    });
    return node;
  }
}
