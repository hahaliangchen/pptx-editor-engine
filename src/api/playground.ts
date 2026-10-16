import { EditApi } from "./engine";
import { batchSchema } from "./schema";

interface Example {
  title: string;
  commands: unknown;
}

const EXAMPLES: Example[] = [
  {
    title: "改标题文字（保留格式）",
    commands: [{ op: "setText", target: { slide: 1, placeholder: "title" }, text: "新的标题" }]
  },
  {
    title: "加一个色块并写字",
    commands: {
      label: "加色块",
      commands: [
        { op: "addShape", slide: 1, shape: "roundRect", rect: { x: 3, y: 10, w: 8, h: 2.5 }, fill: "#E60012", text: "新增内容", as: "badge" },
        { op: "formatText", target: "$badge", bold: true, color: "#FFFFFF", fontSizePt: 20 },
        { op: "setLine", target: "$badge", color: null }
      ]
    }
  },
  {
    title: "加文本框（两段，各自对齐）",
    commands: [{
      op: "addText", slide: 1, rect: { x: 3, y: 2, w: 14, h: 3 },
      paragraphs: [{ text: "第一段，居中", align: "center" }, { text: "第二段，靠左" }],
      style: { fontSizePt: 24, bold: true, color: "#C00000" }
    }]
  },
  {
    title: "加图表（柱形 + 饼图）",
    commands: {
      label: "加图表",
      commands: [
        {
          op: "addChart", slide: 1, type: "column", title: "季度营收（万元）", dataLabels: true,
          categories: ["一季度", "二季度", "三季度", "四季度"],
          series: [{ name: "华东", values: [120, 150, 90, 180] }, { name: "华北", values: [80, 110, 130, 95] }],
          rect: { x: 1, y: 8, w: 15, h: 9 }, name: "营收图"
        },
        {
          op: "addChart", slide: 1, type: "pie", title: "市场份额",
          categories: ["自营", "渠道", "线上", "其他"], series: [{ name: "份额", values: [45, 25, 20, 10] }],
          rect: { x: 17.5, y: 8, w: 15, h: 9 }
        }
      ]
    }
  },
  {
    title: "改图表数据",
    commands: [{
      op: "setChart", target: { slide: 1, name: "营收图" }, title: "更新后的营收",
      categories: ["Q1", "Q2", "Q3"], series: [{ name: "线上", values: [10, 30, 60] }, { name: "线下", values: [40, 35, 30] }]
    }]
  },
  {
    title: "复制本页并在新页改字",
    commands: [{ op: "duplicateSlide", slide: 1 }, { op: "setText", target: { slide: 2, placeholder: "title" }, text: "第二页标题" }]
  },
  {
    title: "故意写错：找不到形状（整批回滚）",
    commands: [
      { op: "addShape", slide: 1, shape: "rect", rect: { x: 1, y: 1, w: 3, h: 2 } },
      { op: "setFill", target: { slide: 1, name: "不存在的形状" }, color: "#00B050" }
    ]
  },
  {
    title: "故意写错：字段不合格",
    commands: [{ op: "setFill", target: { slide: 1, id: "2" }, color: "red" }]
  }
];

/** A drawer for trying the JSON interface by hand: commands in, result and outline out. */
export class ApiPlayground {
  private panel = document.createElement("aside");
  private input = document.createElement("textarea");
  private summary = document.createElement("div");
  private output = document.createElement("pre");
  private busy = false;

  constructor(private api: EditApi, toggle: HTMLElement, private currentSlide: () => number) {
    this.panel.className = "ai-panel";
    this.panel.hidden = true;

    const header = document.createElement("div");
    header.className = "ai-header";
    const title = document.createElement("strong");
    title.textContent = "AI 接口测试";
    const close = document.createElement("button");
    close.type = "button";
    close.className = "ai-close";
    close.textContent = "×";
    close.setAttribute("aria-label", "关闭");
    close.onclick = () => this.setOpen(false);
    header.append(title, close);

    const examples = document.createElement("select");
    examples.className = "ai-select";
    examples.innerHTML = `<option value="">示例指令…</option>${EXAMPLES.map((example, index) => `<option value="${index}">${example.title}</option>`).join("")}`;
    examples.onchange = () => {
      const example = EXAMPLES[Number(examples.value)];
      if (example) this.input.value = JSON.stringify(example.commands, null, 2);
      examples.value = "";
    };

    this.input.className = "ai-input";
    this.input.spellcheck = false;
    this.input.placeholder = '[{"op": "setText", "target": {"slide": 1, "placeholder": "title"}, "text": "新标题"}]';
    this.input.value = JSON.stringify(EXAMPLES[0].commands, null, 2);
    this.input.addEventListener("keydown", event => {
      if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
        event.preventDefault();
        void this.run();
      }
    });

    const actions = document.createElement("div");
    actions.className = "ai-actions";
    actions.append(
      this.button("执行 (Ctrl+Enter)", "is-primary", () => void this.run()),
      this.button("读取全部结构", "", () => void this.readOutline(false)),
      this.button("只读当前页", "", () => void this.readOutline(true)),
      this.button("查看 Schema", "", () => this.show("Schema", batchSchema, true))
    );

    this.summary.className = "ai-summary";
    this.output.className = "ai-output";
    const copy = this.button("复制结果", "is-quiet", () => void navigator.clipboard?.writeText(this.output.textContent || ""));

    this.panel.append(header, examples, this.input, actions, this.summary, this.output, copy);
    document.body.appendChild(this.panel);
    toggle.addEventListener("click", () => this.setOpen(this.panel.hidden));
  }

  public setOpen(open: boolean): void {
    this.panel.hidden = !open;
    document.body.classList.toggle("has-ai-panel", open);
  }

  private button(label: string, className: string, onClick: () => void): HTMLButtonElement {
    const node = document.createElement("button");
    node.type = "button";
    node.className = `ai-button ${className}`.trim();
    node.textContent = label;
    node.onclick = onClick;
    return node;
  }

  private show(label: string, value: unknown, ok: boolean, detail = ""): void {
    this.summary.textContent = `${label}${detail ? ` · ${detail}` : ""}`;
    this.summary.className = `ai-summary ${ok ? "is-ok" : "is-error"}`;
    this.output.textContent = JSON.stringify(value, null, 2);
  }

  private async timed<T>(work: () => Promise<T>): Promise<[T, number]> {
    const started = performance.now();
    const value = await work();
    return [value, Math.round(performance.now() - started)];
  }

  private async run(): Promise<void> {
    if (this.busy) return;
    let commands: unknown;
    try {
      commands = JSON.parse(this.input.value);
    } catch (error) {
      this.show("✗ 不是合法的 JSON", { message: error instanceof Error ? error.message : String(error) }, false);
      return;
    }
    this.busy = true;
    try {
      const [result, ms] = await this.timed(() => this.api.apply(commands));
      const count = result.results.length;
      this.show(
        result.ok ? "✓ 已执行" : `✗ 第 ${(result.failedAt ?? 0) + 1} 条失败：${result.error?.code}`,
        result,
        result.ok,
        `${count} 条结果 · ${ms} ms`
      );
    } catch (error) {
      this.show("✗ 出错", { message: error instanceof Error ? error.message : String(error) }, false);
    } finally {
      this.busy = false;
    }
  }

  private async readOutline(currentOnly: boolean): Promise<void> {
    try {
      const [outline, ms] = await this.timed(() => this.api.outline(currentOnly ? { slides: [this.currentSlide()] } : {}));
      this.show("结构", outline, true, `${outline.slides.length} 页 · ${ms} ms`);
    } catch (error) {
      this.show("✗ 读取失败", { message: error instanceof Error ? error.message : String(error) }, false);
    }
  }
}
