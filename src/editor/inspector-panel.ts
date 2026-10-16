import { Paint, SelectionProperties } from "./properties";
import { LineDash, ShapeStyleOp } from "./shape-ops";

export interface InspectorActions {
  onShapeStyle(op: ShapeStyleOp): void;
  onAlignText(align: "left" | "center" | "right"): void;
  onAlignSlide(edge: "left" | "hcenter" | "right" | "top" | "vcenter" | "bottom"): void;
  onArrange(mode: "front" | "forward" | "backward" | "back"): void;
  onDuplicate(): void;
  /** Centimetres; `keepRatio` when the lock is on. */
  onBounds(change: Partial<{ x: number; y: number; w: number; h: number }>, keepRatio: boolean): void;
  onRotate(degrees: number): void;
}

const PALETTE = [
  "#000000", "#434343", "#7f7f7f", "#bfbfbf", "#ffffff", "#e60012", "#c00000", "#ff6a00",
  "#ffc000", "#ffe600", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0"
];
const WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

const svg = (body: string) =>
  `<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`;

const ICONS = {
  textLeft: svg('<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>'),
  textCenter: svg('<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>'),
  textRight: svg('<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>'),
  anchorTop: svg('<path d="M4 4h16"/><rect x="8" y="8" width="8" height="10" rx="1.5"/>'),
  anchorMiddle: svg('<path d="M4 12h3M17 12h3"/><rect x="8" y="6" width="8" height="12" rx="1.5"/>'),
  anchorBottom: svg('<path d="M4 20h16"/><rect x="8" y="6" width="8" height="10" rx="1.5"/>'),
  alignLeft: svg('<path d="M4 4v16"/><rect x="8" y="6" width="12" height="4" rx="1"/><rect x="8" y="14" width="7" height="4" rx="1"/>'),
  alignHCenter: svg('<path d="M12 4v16"/><rect x="5" y="6" width="14" height="4" rx="1"/><rect x="8" y="14" width="8" height="4" rx="1"/>'),
  alignRight: svg('<path d="M20 4v16"/><rect x="4" y="6" width="12" height="4" rx="1"/><rect x="9" y="14" width="7" height="4" rx="1"/>'),
  alignTop: svg('<path d="M4 4h16"/><rect x="6" y="8" width="4" height="12" rx="1"/><rect x="14" y="8" width="4" height="7" rx="1"/>'),
  alignVCenter: svg('<path d="M4 12h16"/><rect x="6" y="5" width="4" height="14" rx="1"/><rect x="14" y="8" width="4" height="8" rx="1"/>'),
  alignBottom: svg('<path d="M4 20h16"/><rect x="6" y="4" width="4" height="12" rx="1"/><rect x="14" y="9" width="4" height="7" rx="1"/>'),
  front: svg('<rect x="3" y="3" width="11" height="11" rx="1.5" fill="currentColor" fill-opacity=".18"/><rect x="10" y="10" width="11" height="11" rx="1.5" fill="currentColor"/>'),
  forward: svg('<rect x="3" y="9" width="10" height="10" rx="1.5"/><rect x="9" y="4" width="10" height="10" rx="1.5" fill="currentColor" fill-opacity=".18"/>'),
  backward: svg('<rect x="9" y="4" width="10" height="10" rx="1.5"/><rect x="3" y="9" width="10" height="10" rx="1.5" fill="currentColor" fill-opacity=".18"/>'),
  back: svg('<rect x="10" y="10" width="11" height="11" rx="1.5" fill-opacity=".18"/><rect x="3" y="3" width="11" height="11" rx="1.5" fill="currentColor"/>'),
  duplicate: svg('<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2"/>'),
  lock: svg('<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>'),
  unlock: svg('<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 7.5-2"/>'),
  solid: svg('<path d="M4 12h16"/>'),
  dash: svg('<path d="M3 12h4M10 12h4M17 12h4"/>'),
  dot: svg('<path d="M4 12h.01M9 12h.01M14 12h.01M19 12h.01" stroke-width="3"/>'),
  minus: svg('<path d="M6 12h12"/>'),
  plus: svg('<path d="M12 6v12M6 12h12"/>')
};

interface Segment<T> {
  value: T;
  icon: string;
  title: string;
}

/**
 * Docked "format" panel: fill, border, text alignment and alignment to the
 * page. It only reports intent through `actions` and repaints from the
 * selection's properties, so it holds no document state of its own.
 */
export class InspectorPanel {
  private sections = new Map<string, HTMLElement>();
  private fillChips: HTMLButtonElement[] = [];
  private lineChips: HTMLButtonElement[] = [];
  private fillNone!: HTMLButtonElement;
  private lineNone!: HTMLButtonElement;
  private fillCustom!: HTMLInputElement;
  private lineCustom!: HTMLInputElement;
  private widthLabel!: HTMLElement;
  private dashButtons = new Map<LineDash, HTMLButtonElement>();
  private alignButtons = new Map<string, HTMLButtonElement>();
  private anchorButtons = new Map<string, HTMLButtonElement>();
  private subtitle: HTMLElement;
  private fields = new Map<"x" | "y" | "w" | "h" | "rotation", HTMLInputElement>();
  private ratioLock!: HTMLButtonElement;
  private props: SelectionProperties | null = null;

  constructor(private root: HTMLElement, private actions: InspectorActions) {
    root.classList.add("inspector");
    root.replaceChildren();

    const header = document.createElement("div");
    header.className = "inspector-header";
    const title = document.createElement("h2");
    title.textContent = "格式";
    this.subtitle = document.createElement("span");
    this.subtitle.className = "inspector-subtitle";
    header.append(title, this.subtitle);
    root.appendChild(header);

    this.buildGeometry();
    this.buildFill();
    this.buildLine();
    this.buildText();
    this.buildArrange();
    this.buildLayers();
    this.update(null);
  }

  public update(props: SelectionProperties | null, selectionCount = 0): void {
    this.props = props;
    this.subtitle.textContent = props ? `已选中 ${selectionCount} 个对象` : "未选择对象";
    const has = (name: string, enabled: boolean) => this.sections.get(name)?.classList.toggle("is-disabled", !enabled);
    has("geometry", !!props && !!props.bounds);
    has("fill", !!props && props.shapeCount > 0);
    has("line", !!props && props.shapeCount > 0);
    has("text", !!props && props.textCount > 0);
    has("arrange", !!props);
    has("layers", !!props);

    this.paintGeometry(props);
    this.paintPalette(this.fillChips, this.fillNone, this.fillCustom, props?.fill ?? { kind: "mixed" });
    this.paintPalette(this.lineChips, this.lineNone, this.lineCustom, props?.lineColor ?? { kind: "mixed" });
    this.widthLabel.textContent = props?.lineWidthPt != null ? `${props.lineWidthPt} 磅` : "—";
    for (const [value, button] of this.dashButtons) button.classList.toggle("is-active", props?.lineDash === value);
    for (const [value, button] of this.alignButtons) button.classList.toggle("is-active", props?.align === value);
    for (const [value, button] of this.anchorButtons) button.classList.toggle("is-active", props?.anchor === value);
  }

  // ---------------------------------------------------------------- sections

  private section(name: string, title: string): HTMLElement {
    const card = document.createElement("section");
    card.className = "inspector-section";
    const heading = document.createElement("h3");
    heading.textContent = title;
    card.appendChild(heading);
    this.root.appendChild(card);
    this.sections.set(name, card);
    return card;
  }

  private buildGeometry(): void {
    const card = this.section("geometry", "位置与大小");
    const grid = document.createElement("div");
    grid.className = "inspector-fields";
    const field = (key: "x" | "y" | "w" | "h" | "rotation", label: string, unit: string) => {
      const wrap = document.createElement("label");
      wrap.className = "inspector-field";
      const name = document.createElement("span");
      name.textContent = label;
      const input = document.createElement("input");
      input.type = "number";
      input.step = key === "rotation" ? "1" : "0.1";
      input.inputMode = "decimal";
      input.className = "inspector-number";
      input.addEventListener("change", () => this.commitField(key, input));
      input.addEventListener("keydown", event => {
        if (event.key === "Enter") input.blur();
        if (event.key === "Escape") {
          input.value = input.dataset.committed || "";
          input.blur();
        }
      });
      const suffix = document.createElement("em");
      suffix.textContent = unit;
      wrap.append(name, input, suffix);
      this.fields.set(key, input);
      return wrap;
    };
    this.ratioLock = this.iconButton(ICONS.unlock, "锁定宽高比", () => {
      const locked = this.ratioLock.classList.toggle("is-active");
      this.ratioLock.innerHTML = locked ? ICONS.lock : ICONS.unlock;
    });
    this.ratioLock.classList.add("inspector-lock");
    grid.append(
      field("x", "X", "cm"), field("y", "Y", "cm"),
      field("w", "宽", "cm"), field("h", "高", "cm"),
      field("rotation", "旋转", "°"), this.ratioLock
    );
    card.appendChild(grid);

    // A number typed in a box must land on the selection it was typed for, so
    // leave the box (committing it) before a click can change the selection.
    document.addEventListener("pointerdown", () => {
      const active = document.activeElement;
      if (active instanceof HTMLInputElement && this.root.contains(active) && active.type === "number") active.blur();
    }, true);
  }

  private commitField(key: "x" | "y" | "w" | "h" | "rotation", input: HTMLInputElement): void {
    const value = parseFloat(input.value);
    if (!Number.isFinite(value)) {
      input.value = input.dataset.committed || "";
      return;
    }
    if (key === "rotation") this.actions.onRotate(value);
    else this.actions.onBounds({ [key]: value }, this.ratioLock.classList.contains("is-active"));
  }

  private paintGeometry(props: SelectionProperties | null): void {
    const values: Record<string, number | null> = {
      x: props?.bounds?.x ?? null,
      y: props?.bounds?.y ?? null,
      w: props?.bounds?.w ?? null,
      h: props?.bounds?.h ?? null,
      rotation: props?.rotation ?? null
    };
    for (const [key, input] of this.fields) {
      const value = values[key];
      input.disabled = value === null;
      const text = value === null ? "" : String(key === "rotation" ? Math.round(value * 10) / 10 : Math.round(value * 100) / 100);
      input.dataset.committed = text;
      // Never rewrite a box the user is typing in.
      if (document.activeElement !== input) input.value = text;
    }
  }

  private buildFill(): void {
    const card = this.section("fill", "填充");
    const palette = this.palette(
      color => this.actions.onShapeStyle({ kind: "fill", color }),
      color => this.actions.onShapeStyle({ kind: "fill", color })
    );
    this.fillChips = palette.chips;
    this.fillNone = palette.none;
    this.fillCustom = palette.custom;
    card.appendChild(palette.root);
  }

  private buildLine(): void {
    const card = this.section("line", "边框");
    const palette = this.palette(
      color => this.actions.onShapeStyle({ kind: "line", color }),
      color => this.actions.onShapeStyle({ kind: "line", color })
    );
    this.lineChips = palette.chips;
    this.lineNone = palette.none;
    this.lineCustom = palette.custom;
    card.appendChild(palette.root);

    const row = document.createElement("div");
    row.className = "inspector-row";
    const stepper = document.createElement("div");
    stepper.className = "inspector-stepper";
    const minus = this.iconButton(ICONS.minus, "变细", () => this.stepWidth(-1));
    this.widthLabel = document.createElement("span");
    this.widthLabel.className = "inspector-value";
    const plus = this.iconButton(ICONS.plus, "加粗", () => this.stepWidth(1));
    stepper.append(minus, this.widthLabel, plus);

    const dashes = this.segmented<LineDash>(
      [
        { value: "solid", icon: ICONS.solid, title: "实线" },
        { value: "dash", icon: ICONS.dash, title: "虚线" },
        { value: "sysDot", icon: ICONS.dot, title: "点线" }
      ],
      value => this.actions.onShapeStyle({ kind: "line", dash: value }),
      this.dashButtons
    );
    row.append(stepper, dashes);
    card.appendChild(row);
  }

  private buildText(): void {
    const card = this.section("text", "文字对齐");
    const row = document.createElement("div");
    row.className = "inspector-row";
    row.append(
      this.segmented<"left" | "center" | "right">(
        [
          { value: "left", icon: ICONS.textLeft, title: "左对齐" },
          { value: "center", icon: ICONS.textCenter, title: "居中" },
          { value: "right", icon: ICONS.textRight, title: "右对齐" }
        ],
        value => this.actions.onAlignText(value),
        this.alignButtons
      ),
      this.segmented<"top" | "middle" | "bottom">(
        [
          { value: "top", icon: ICONS.anchorTop, title: "顶端对齐" },
          { value: "middle", icon: ICONS.anchorMiddle, title: "垂直居中" },
          { value: "bottom", icon: ICONS.anchorBottom, title: "底端对齐" }
        ],
        value => this.actions.onShapeStyle({ kind: "anchor", value }),
        this.anchorButtons
      )
    );
    card.appendChild(row);
  }

  private buildArrange(): void {
    const card = this.section("arrange", "对齐页面");
    const row = document.createElement("div");
    row.className = "inspector-grid";
    const edges: Array<[Parameters<InspectorActions["onAlignSlide"]>[0], string, string]> = [
      ["left", ICONS.alignLeft, "靠左"],
      ["hcenter", ICONS.alignHCenter, "水平居中"],
      ["right", ICONS.alignRight, "靠右"],
      ["top", ICONS.alignTop, "靠上"],
      ["vcenter", ICONS.alignVCenter, "垂直居中"],
      ["bottom", ICONS.alignBottom, "靠下"]
    ];
    for (const [edge, icon, title] of edges) {
      row.appendChild(this.iconButton(icon, title, () => this.actions.onAlignSlide(edge)));
    }
    card.appendChild(row);
  }

  private buildLayers(): void {
    const card = this.section("layers", "层级与副本");
    const row = document.createElement("div");
    row.className = "inspector-grid inspector-grid-five";
    const layers: Array<["front" | "forward" | "backward" | "back", string, string]> = [
      ["front", ICONS.front, "置于顶层 (Ctrl+Shift+])"],
      ["forward", ICONS.forward, "上移一层 (Ctrl+])"],
      ["backward", ICONS.backward, "下移一层 (Ctrl+[)"],
      ["back", ICONS.back, "置于底层 (Ctrl+Shift+[)"]
    ];
    for (const [mode, icon, title] of layers) {
      row.appendChild(this.iconButton(icon, title, () => this.actions.onArrange(mode)));
    }
    row.appendChild(this.iconButton(ICONS.duplicate, "创建副本 (Ctrl+D)", () => this.actions.onDuplicate()));
    card.appendChild(row);
  }

  // ---------------------------------------------------------------- controls

  private palette(onPick: (color: string) => void, onNone: (color: null) => void) {
    const root = document.createElement("div");
    root.className = "inspector-palette";

    const none = document.createElement("button");
    none.type = "button";
    none.className = "inspector-chip is-none";
    none.title = "无";
    none.addEventListener("click", () => onNone(null));
    root.appendChild(none);

    const chips = PALETTE.map(color => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "inspector-chip";
      chip.style.background = color;
      chip.dataset.color = color;
      chip.title = color.toUpperCase();
      chip.addEventListener("click", () => onPick(color));
      root.appendChild(chip);
      return chip;
    });

    const custom = document.createElement("input");
    custom.type = "color";
    custom.className = "inspector-chip is-custom";
    custom.title = "自定义颜色";
    custom.value = "#e60012";
    custom.addEventListener("change", () => onPick(custom.value));
    root.appendChild(custom);

    return { root, chips, none, custom };
  }

  private paintPalette(chips: HTMLButtonElement[], none: HTMLButtonElement, custom: HTMLInputElement, paint: Paint): void {
    const color = paint.kind === "solid" ? paint.color : null;
    none.classList.toggle("is-active", paint.kind === "none");
    let matched = false;
    for (const chip of chips) {
      const active = color !== null && chip.dataset.color === color;
      chip.classList.toggle("is-active", active);
      matched ||= active;
    }
    custom.classList.toggle("is-active", color !== null && !matched);
    if (color !== null) custom.value = color;
  }

  private segmented<T extends string>(
    segments: Array<Segment<T>>,
    onPick: (value: T) => void,
    registry: Map<string, HTMLButtonElement>
  ): HTMLElement {
    const root = document.createElement("div");
    root.className = "inspector-segmented";
    for (const segment of segments) {
      const button = this.iconButton(segment.icon, segment.title, () => onPick(segment.value));
      registry.set(segment.value, button);
      root.appendChild(button);
    }
    return root;
  }

  private iconButton(icon: string, title: string, onClick: () => void): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "inspector-icon";
    button.title = title;
    button.setAttribute("aria-label", title);
    button.innerHTML = icon;
    button.addEventListener("click", onClick);
    return button;
  }

  private stepWidth(direction: 1 | -1): void {
    const current = this.props?.lineWidthPt ?? 1;
    const next = direction > 0
      ? WIDTHS.find(width => width > current + 0.001)
      : [...WIDTHS].reverse().find(width => width < current - 0.001);
    if (next !== undefined) this.actions.onShapeStyle({ kind: "line", widthPt: next });
  }
}
