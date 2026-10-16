import { NewShapeKind } from "./new-shapes";

const svg = (body: string) =>
  `<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`;

const TOOLS: Array<{ kind: NewShapeKind | "image"; title: string; icon: string }> = [
  { kind: "textbox", title: "文本框", icon: svg('<path d="M5 6V5h14v1M12 5v14M9.5 19h5"/>') },
  { kind: "rect", title: "矩形", icon: svg('<rect x="4" y="6" width="16" height="12" rx="1.5"/>') },
  { kind: "roundRect", title: "圆角矩形", icon: svg('<rect x="4" y="6" width="16" height="12" rx="5"/>') },
  { kind: "ellipse", title: "椭圆", icon: svg('<ellipse cx="12" cy="12" rx="8" ry="6"/>') },
  { kind: "triangle", title: "三角形", icon: svg('<path d="M12 5 20 19H4z"/>') },
  { kind: "rightArrow", title: "右箭头", icon: svg('<path d="M4 9h9V5l7 7-7 7v-4H4z"/>') },
  { kind: "line", title: "直线", icon: svg('<path d="M5 19 19 5"/>') },
  { kind: "arrow", title: "箭头", icon: svg('<path d="M5 19 19 5M10 5h9v9"/>') },
  { kind: "image", title: "图片", icon: svg('<rect x="4" y="5" width="16" height="14" rx="2"/><circle cx="9" cy="10" r="1.5"/><path d="m5 17 5-4.5 3.5 3L16 13l3 3"/>') }
];

export interface InsertBarActions {
  onTool(kind: NewShapeKind | null): void;
  onImage(): void;
}

/** Floating strip of things to draw; clicking a tool arms it for the next drag on the slide. */
export class InsertBar {
  private buttons = new Map<string, HTMLButtonElement>();
  private active: NewShapeKind | null = null;

  constructor(private root: HTMLElement, private actions: InsertBarActions) {
    root.replaceChildren();
    TOOLS.forEach((tool, index) => {
      if (index === 1 || tool.kind === "line" || tool.kind === "image") {
        const divider = document.createElement("span");
        divider.className = "insert-divider";
        root.appendChild(divider);
      }
      const button = document.createElement("button");
      button.type = "button";
      button.className = "insert-tool";
      button.title = tool.kind === "image" ? "插入图片" : `${tool.title}（拖动绘制）`;
      button.setAttribute("aria-label", tool.title);
      button.innerHTML = tool.icon;
      button.addEventListener("click", () => {
        if (tool.kind === "image") this.actions.onImage();
        else this.actions.onTool(this.active === tool.kind ? null : tool.kind);
      });
      this.buttons.set(tool.kind, button);
      root.appendChild(button);
    });
  }

  public update(active: NewShapeKind | null, enabled: boolean): void {
    this.active = active;
    this.root.classList.toggle("is-disabled", !enabled);
    for (const [kind, button] of this.buttons) button.classList.toggle("is-active", kind === active);
  }
}
