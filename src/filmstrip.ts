export interface FilmstripActions {
  select(index: number): void;
  add(index: number): void;
  duplicate(index: number): void;
  remove(index: number): void;
  move(from: number, to: number): void;
  /** Render thumbnails for the given slides; the result is indexed by slide. */
  renderThumbnails(indices: number[]): Promise<string[]>;
}

interface SlideRef {
  id: string;
}

/**
 * The thumbnail column: select, drag to reorder, right-click for slide
 * actions, Delete to remove. Thumbnails are cached per slide so changing the
 * list only renders the ones that are new or edited.
 */
export class Filmstrip {
  private thumbs = new Map<string, string>();
  private slides: SlideRef[] = [];
  private active = 0;
  private menu: HTMLDivElement | null = null;
  private dragFrom = -1;

  constructor(private list: HTMLElement, private count: HTMLElement, private actions: FilmstripActions) {}

  public setActive(index: number): void {
    this.active = index;
    this.items().forEach((item, i) => item.classList.toggle("active", i === index));
    this.items()[index]?.scrollIntoView({ block: "nearest" });
  }

  public async rebuild(slides: SlideRef[], active: number): Promise<void> {
    this.slides = slides;
    this.active = active;
    this.count.textContent = String(slides.length);
    this.list.replaceChildren(...slides.map((slide, index) => this.item(slide, index)));
    await this.fillMissing();
  }

  /** A slide was edited: its cached thumbnail is stale. */
  public async refresh(index: number): Promise<void> {
    const slide = this.slides[index];
    if (!slide) return;
    this.thumbs.delete(slide.id);
    await this.fillMissing([index]);
  }

  private items(): HTMLElement[] {
    return Array.from(this.list.querySelectorAll<HTMLElement>(".slide-item"));
  }

  private item(slide: SlideRef, index: number): HTMLElement {
    const li = document.createElement("li");
    li.className = `slide-item${index === this.active ? " active" : ""}`;
    li.tabIndex = 0;
    li.draggable = true;
    li.setAttribute("aria-label", `幻灯片 ${index + 1}`);

    const thumb = document.createElement("div");
    const url = this.thumbs.get(slide.id);
    thumb.className = `slide-thumb${url ? "" : " is-loading"}`;
    const img = document.createElement("img");
    img.alt = `幻灯片 ${index + 1} 缩略图`;
    img.draggable = false;
    if (url) img.src = url;
    else img.style.display = "none";
    thumb.appendChild(img);
    li.appendChild(thumb);

    li.addEventListener("click", () => {
      this.actions.select(index);
      li.focus({ preventScroll: true });
    });
    li.addEventListener("contextmenu", event => {
      event.preventDefault();
      this.actions.select(index);
      this.openMenu(event, index);
    });
    li.addEventListener("keydown", event => this.onKey(event, index));

    li.addEventListener("dragstart", event => {
      this.dragFrom = index;
      event.dataTransfer?.setData("text/plain", String(index));
      if (event.dataTransfer) event.dataTransfer.effectAllowed = "move";
      li.classList.add("is-dragging");
    });
    li.addEventListener("dragend", () => {
      this.dragFrom = -1;
      this.clearDropMarks();
      li.classList.remove("is-dragging");
    });
    li.addEventListener("dragover", event => {
      if (this.dragFrom < 0) return;
      event.preventDefault();
      const after = event.offsetY > li.clientHeight / 2;
      this.clearDropMarks();
      li.classList.add(after ? "drop-after" : "drop-before");
    });
    li.addEventListener("dragleave", () => li.classList.remove("drop-before", "drop-after"));
    li.addEventListener("drop", event => {
      if (this.dragFrom < 0) return;
      event.preventDefault();
      const after = li.classList.contains("drop-after");
      this.clearDropMarks();
      // Index the slide will have once it has been taken out of its old place.
      let to = index + (after ? 1 : 0);
      if (this.dragFrom < to) to -= 1;
      const from = this.dragFrom;
      this.dragFrom = -1;
      if (to !== from) this.actions.move(from, to);
    });
    return li;
  }

  private clearDropMarks(): void {
    for (const item of this.items()) item.classList.remove("drop-before", "drop-after");
  }

  private async fillMissing(only?: number[]): Promise<void> {
    const wanted = (only ?? this.slides.map((_, index) => index)).filter(index => {
      const slide = this.slides[index];
      return slide && !this.thumbs.has(slide.id);
    });
    if (wanted.length === 0) return;
    const ids = wanted.map(index => this.slides[index].id);
    let urls: string[] = [];
    try {
      urls = await this.actions.renderThumbnails(wanted);
    } catch (err) {
      console.error("Failed to render thumbnails:", err);
    }
    wanted.forEach((index, order) => {
      const url = urls[index];
      if (url) this.thumbs.set(ids[order], url);
      // The list may have been rebuilt meanwhile; match by slide id.
      const at = this.slides.findIndex(slide => slide.id === ids[order]);
      const item = this.items()[at];
      if (!item) return;
      item.querySelector(".slide-thumb")?.classList.remove("is-loading");
      const img = item.querySelector("img");
      if (img && url) {
        img.src = url;
        img.style.display = "block";
      }
    });
  }

  private onKey(event: KeyboardEvent, index: number): void {
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      this.actions.remove(index);
    } else if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      const next = index + (event.key === "ArrowDown" ? 1 : -1);
      if (next >= 0 && next < this.slides.length) {
        this.actions.select(next);
        this.items()[next]?.focus({ preventScroll: true });
      }
    }
  }

  private openMenu(event: MouseEvent, index: number): void {
    this.closeMenu();
    const last = this.slides.length - 1;
    const entries: Array<{ label: string; hint?: string; enabled: boolean; run: () => void } | "-"> = [
      { label: "新建幻灯片", enabled: true, run: () => this.actions.add(index) },
      { label: "复制幻灯片", enabled: true, run: () => this.actions.duplicate(index) },
      "-",
      { label: "上移一页", enabled: index > 0, run: () => this.actions.move(index, index - 1) },
      { label: "下移一页", enabled: index < last, run: () => this.actions.move(index, index + 1) },
      "-",
      { label: "删除幻灯片", hint: "Delete", enabled: this.slides.length > 1, run: () => this.actions.remove(index) }
    ];
    const menu = document.createElement("div");
    menu.className = "edit-menu is-fixed";
    for (const entry of entries) {
      if (entry === "-") {
        const divider = document.createElement("div");
        divider.className = "edit-menu-divider";
        menu.appendChild(divider);
        continue;
      }
      const row = document.createElement("button");
      row.type = "button";
      row.className = "edit-menu-item";
      row.disabled = !entry.enabled;
      const label = document.createElement("span");
      label.textContent = entry.label;
      row.appendChild(label);
      if (entry.hint) {
        const hint = document.createElement("kbd");
        hint.textContent = entry.hint;
        row.appendChild(hint);
      }
      row.addEventListener("click", () => {
        this.closeMenu();
        entry.run();
      });
      menu.appendChild(row);
    }
    document.body.appendChild(menu);
    this.menu = menu;
    const x = Math.min(event.clientX, window.innerWidth - menu.offsetWidth - 6);
    const y = Math.min(event.clientY, window.innerHeight - menu.offsetHeight - 6);
    menu.style.left = `${Math.max(6, x)}px`;
    menu.style.top = `${Math.max(6, y)}px`;

    const dismiss = (e: Event) => {
      if (e instanceof KeyboardEvent && e.key !== "Escape") return;
      if (e.target instanceof Node && menu.contains(e.target)) return;
      this.closeMenu();
      document.removeEventListener("pointerdown", dismiss, true);
      document.removeEventListener("keydown", dismiss, true);
    };
    document.addEventListener("pointerdown", dismiss, true);
    document.addEventListener("keydown", dismiss, true);
  }

  private closeMenu(): void {
    this.menu?.remove();
    this.menu = null;
  }
}
