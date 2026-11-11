import type { MeasuredPage } from "./model";
/** Browser DOM measurement shared by the product iframe and the Node CLI. No Node dependencies. */
export async function measureDesignDOM({ nonce, collectIssues = false }: { nonce: string; collectIssues?: boolean }, document: Document = globalThis.document): Promise<MeasuredPage> {
  const computedStyle = document.defaultView!.getComputedStyle.bind(document.defaultView!);
  const issues: NonNullable<MeasuredPage["issues"]> = [];
  const issue = (message: string, nodeId?: string) => { if (!collectIssues) throw new Error(message); issues.push({ nodeId, message }); };
      await document.fonts.ready;
      await Promise.all(Array.from(document.images).map(async img => { if (!img.src.startsWith("data:image/")) throw new Error("Images must use embedded project assets"); await img.decode(); }));
      if (document.querySelector("script,iframe,object,embed,link,base,meta[http-equiv]")) throw new Error("Only static HTML/CSS and embedded assets are accepted");
      const body = document.body, br = body.getBoundingClientRect();
      if (Math.abs(br.width - 1280) > 0.1 || Math.abs(br.height - 720) > 0.1 || Math.abs(br.x) > 0.1 || Math.abs(br.y) > 0.1) throw new Error("HTML body must be exactly 1280x720 at (0,0)");
      if (body.scrollWidth > 1281 || body.scrollHeight > 721) issue("HTML page overflows its 1280x720 canvas");
      const revision = Number(body.dataset.designRevision) || 0, pageId = body.dataset.pageId;
      if (!pageId) throw new Error("Missing data-page-id");
      const px = (value: string) => Number.parseFloat(value) || 0;
      const color = (value: string): string => {
        const rgba = value.match(/^rgba?\(\s*([\d.]+)[, ]+([\d.]+)[, ]+([\d.]+)(?:\s*[,/]\s*([\d.]+))?\s*\)$/);
        if (!rgba) return value.toLowerCase();
        if (rgba[4] !== undefined && Number(rgba[4]) < 1) return value;
        return `#${rgba.slice(1, 4).map(n => Math.round(Number(n)).toString(16).padStart(2, "0")).join("")}`;
      };
      const transparent = (value: string) => value === "transparent" || value === "rgba(0, 0, 0, 0)";
      // Computed CSS values only; the browser has already normalized colors to rgb()/rgba().
      const splitTop = (value: string) => {
        const parts: string[] = []; let depth = 0, start = 0;
        for (let i = 0; i < value.length; i++) {
          if (value[i] === "(") depth++; else if (value[i] === ")") depth--;
          else if (value[i] === "," && !depth) { parts.push(value.slice(start, i).trim()); start = i + 1; }
        }
        return [...parts, value.slice(start).trim()];
      };
      const linearGradient = (value: string, id: string) => {
        const match = splitTop(value).length === 1 ? value.match(/^linear-gradient\(([\s\S]*)\)$/) : null;
        if (!match) throw new Error(`Only one linear-gradient background is supported: ${id}`);
        const args = splitTop(match[1]);
        let angle = 180;
        const sides: Record<string, number> = { "to top": 0, "to right": 90, "to bottom": 180, "to left": 270 };
        const unit = args[0].match(/^(-?[\d.]+)(deg|turn|rad|grad)$/);
        if (unit) { angle = Number(unit[1]) * ({ deg: 1, turn: 360, rad: 180 / Math.PI, grad: 0.9 } as Record<string, number>)[unit[2]]; args.shift(); }
        else if (args[0].startsWith("to ")) {
          if (sides[args[0]] === undefined) throw new Error(`Gradient corners are not supported, use an angle: ${id}`);
          angle = sides[args[0]]; args.shift();
        }
        const stops: Array<{ color: string; position?: number }> = [];
        for (const arg of args) {
          const stop = arg.match(/^(rgba?\([^)]*\)|#[0-9a-f]{3,8}|[a-z]+)\s*(.*)$/i);
          if (!stop) throw new Error(`Unsupported gradient stop "${arg}": ${id}`);
          const positions = stop[2].trim() ? stop[2].trim().split(/\s+/) : [];
          if (positions.some(p => !/^-?[\d.]+%$/.test(p))) throw new Error(`Gradient stop positions must be percentages: ${id}`);
          if (!positions.length) stops.push({ color: color(stop[1]) });
          for (const p of positions) stops.push({ color: color(stop[1]), position: Number.parseFloat(p) });
        }
        if (stops.length < 2) throw new Error(`A gradient needs at least two stops: ${id}`);
        // CSS fix-up: missing ends are 0%/100%, missing middles are spaced evenly, positions never decrease.
        stops[0].position ??= 0; stops[stops.length - 1].position ??= 100;
        for (let i = 1; i < stops.length; i++) {
          if (stops[i].position === undefined) {
            let j = i; while (stops[j].position === undefined) j++;
            const from = stops[i - 1].position!, to = stops[j].position!;
            for (let k = i; k < j; k++) stops[k].position = from + (to - from) * (k - i + 1) / (j - i + 1);
          }
          stops[i].position = Math.max(stops[i].position!, stops[i - 1].position!);
        }
        if (stops.some(s => s.position! < 0 || s.position! > 100)) throw new Error(`Gradient stops must lie within 0%-100%: ${id}`);
        return { angle: ((angle % 360) + 360) % 360, stops: stops.map(s => ({ color: s.color, position: Math.round(s.position! * 1000) / 1000 })) };
      };
      const boxShadow = (value: string, id: string) => {
        if (splitTop(value).length !== 1 || /\binset\b/.test(value)) throw new Error(`Only one outer box-shadow is supported: ${id}`);
        const tint = value.match(/rgba?\([^)]*\)/)?.[0];
        const lengths = value.replace(tint ?? "", "").trim().split(/\s+/).map(px);
        if (!tint || lengths.length < 2 || lengths[3]) throw new Error(`box-shadow needs a color and no spread: ${id}`);
        return { x: lengths[0], y: lengths[1], blur: lengths[2] ?? 0, color: color(tint) };
      };
      const polygon = (value: string, box: DOMRect, id: string) => {
        const match = value.match(/^polygon\((?:(?:nonzero|evenodd)\s*,\s*)?([\s\S]*)\)$/);
        if (!match) throw new Error(`Only polygon() clip-path is supported: ${id}`);
        const points = splitTop(match[1]).map(point => {
          const [x, y] = point.split(/\s+/);
          const coordinate = (v: string | undefined, size: number) => {
            if (v?.endsWith("%")) return Number.parseFloat(v);
            if (v?.endsWith("px") || v === "0") return Number.parseFloat(v) / size * 100;
            throw new Error(`Polygon points must use % or px: ${id}`);
          };
          return [Math.round(coordinate(x, box.width) * 1000) / 1000, Math.round(coordinate(y, box.height) * 1000) / 1000] as [number, number];
        });
        if (points.length < 3) throw new Error(`A polygon needs at least three points: ${id}`);
        return points;
      };
      // The HTML projection of an existing freeform can be a text node with a
      // polygon that traces its entire rectangular box. It clips nothing and
      // may safely be measured as ordinary text without changing its geometry.
      const rectangularClip = (value: string, el: HTMLElement) => {
        if (!value.startsWith("polygon(")) return false;
        try {
          const points = polygon(value, el.getBoundingClientRect(), el.dataset.id ?? "");
          if (points.length === 5 && points[0][0] === points[4][0] && points[0][1] === points[4][1]) points.pop();
          if (points.length !== 4 || new Set(points.map(p => p.join(","))).size !== 4) return false;
          return points.every((p, i) => (p[0] === 0 || p[0] === 100) && (p[1] === 0 || p[1] === 100)
            && ((p[0] === points[(i + 1) % 4][0]) !== (p[1] === points[(i + 1) % 4][1])));
        } catch { return false; }
      };
      const markerElements = Array.from(document.querySelectorAll<HTMLElement>("[data-kind]"));
      const elementPath = (el: HTMLElement): string => {
        const parts: string[] = [];
        for (let node: Element | null = el; node && node !== body; node = node.parentElement) {
          const index = node.parentElement ? Array.from(node.parentElement.children).indexOf(node) + 1 : 1;
          parts.unshift(`${node.tagName.toLowerCase()}:nth-child(${index})`);
        }
        return "body" + (parts.length ? " > " + parts.join(" > ") : "");
      };
      const invalidNode = (code: string, message: string, el: HTMLElement, recovery: string): never => {
        const owner = el.closest<HTMLElement>("[data-kind]");
        throw new Error(`${message}. ${JSON.stringify({
          code, phase: "html_validation", applied: false, pageId, revision, tag: el.tagName.toLowerCase(), nodeId: el.dataset.id ?? null,
          ownerId: owner?.dataset.id ?? null, ownerKind: owner?.dataset.kind ?? null,
          ownerIndex: owner ? markerElements.indexOf(owner) : null,
          selector: elementPath(el), htmlSnippet: el.outerHTML.slice(0, 700),
          textSnippet: (el.textContent ?? "").replace(/\s+/g, " ").trim().slice(0, 160),
          recovery: `${recovery} Preserve the latest get_design data-* attributes and retry apply_html. No page content was applied.`
        })}`);
      };
      // Validate before assigning IDs or walking text: browsers keep only the
      // first duplicate attribute, so an ID accidentally written as data-kind
      // would otherwise surface as an unrelated text ownership error.
      for (const el of markerElements) {
        if (!["text", "shape", "image", "asset", "group"].includes(el.dataset.kind ?? "")) invalidNode(
          "INVALID_DESIGN_KIND", `Unsupported data-kind: ${el.dataset.kind}`, el,
          "Use data-kind= text, shape, image, asset or group. Put the stable node identity in data-id, not data-kind; remove duplicate data-kind attributes."
        );
        if (["image", "asset"].includes(el.dataset.kind!) && el.tagName !== "IMG") invalidNode(
          "IMAGE_TAG_REQUIRED", "Image assets must be img elements", el,
          'Use a direct <img data-kind="image" src="..."> (or data-kind="asset" for the existing asset), not a div wrapper marked as image/asset. Keep its existing identity and metadata on the img; use data-kind="group" only for a separate container.'
        );
      }
      markerElements.forEach((el, i) => { if (!el.dataset.id) el.dataset.id = `new:${nonce}:${i}`; });
      for (const el of [document.documentElement, body, ...Array.from(body.querySelectorAll<HTMLElement>("*"))]) {
        if (["STYLE", "META", "TITLE"].includes(el.tagName)) continue;
        if (el.tagName === "IMG" && !["image", "asset"].includes(el.dataset.kind ?? "")) throw new Error("Every visible image must be a marked image/asset design node");
        const style = computedStyle(el);
        if (style.display === "none" || style.visibility !== "visible") throw new Error("Hidden HTML nodes must be removed explicitly, not silently omitted");
        for (const pseudo of ["::before", "::after"]) {
          const content = computedStyle(el, pseudo).content;
          // Firefox exposes its built-in image alternative text as ::before.
          // Images have already decoded successfully above; this is browser
          // fallback metadata, not authored CSS-generated page content.
          const imageAlt = el.tagName === "IMG" && pseudo === "::before" && content === "-moz-alt-content";
          if (!imageAlt && content && content !== "none" && content !== "normal" && content !== '""') {
            throw new Error(`Pseudo-element content is not supported; use a marked node: ${el.dataset.id ?? el.tagName} ${pseudo} content=${content}`);
          }
        }
        // Gradients and shadows are native shape paint; a polygon clip is native shape geometry.
        const paintable = ["shape", "text"].includes(el.dataset.kind ?? "");
        if (style.transform !== "none" || style.scale && style.scale !== "none" || style.rotate && style.rotate !== "none" || style.translate && style.translate !== "none" || style.filter !== "none" || style.backdropFilter && style.backdropFilter !== "none" || style.mixBlendMode !== "normal" || Number(style.opacity) !== 1 || !paintable && (style.boxShadow !== "none" || style.backgroundImage !== "none") || style.zIndex !== "auto" || style.clipPath !== "none" && el.dataset.kind !== "shape" && !(el.dataset.kind === "text" && rectangularClip(style.clipPath, el)) || style.textShadow !== "none" || style.clip !== "auto" || style.maskImage && style.maskImage !== "none" || style.contain.includes("paint") || style.contentVisibility !== "visible" || el.tagName !== "IMG" && [style.overflowX,style.overflowY].some(v=>v!=="visible")) throw new Error(`Unsupported CSS on ${el.dataset.id ?? el.tagName}; use a separate image/SVG asset`);
        const owner = el.closest<HTMLElement>("[data-kind]");
        if (el !== body && !el.dataset.kind && (!transparent(style.backgroundColor) || px(style.borderTopWidth) || px(style.borderLeftWidth))) invalidNode(
          "UNMARKED_PAINTED_CONTAINER", "Painted container is missing data-kind", el,
          'For a card containing marked objects, use a data-kind="group" container with a separate data-kind="shape" child for its background and sibling text nodes. A top-only border must be a separate thin shape; native borders must be uniform on all sides. Pure layout containers must be transparent and borderless. New nodes may omit data-id; IDs are assigned automatically. Use validate_html to check the repair without changing the deck.'
        );
        if (Array.from(el.childNodes).some(n => n.nodeType === Node.TEXT_NODE && !!n.textContent?.trim()) && (!owner || !["text", "shape"].includes(owner.dataset.kind!))) invalidNode(
          "TEXT_OWNER_REQUIRED", "Text must belong to a data-kind=text or shape node", el,
          'Place this text in a p, heading or li inside a data-kind="text" or data-kind="shape" node. A group may contain marked text nodes but cannot directly own text. Do not change an existing node identity or remove fixed assets.'
        );
      }
      const nodes: MeasuredPage["nodes"] = [];
      const textPaint: NonNullable<MeasuredPage["textPaint"]> = {};
      const textContext = document.createElement("canvas").getContext("2d")!;
      const metricsCache = new Map<string,TextMetrics>();
      for (const el of markerElements) {
        const kind = el.dataset.kind as MeasuredPage["nodes"][number]["kind"];
        const id = el.dataset.id!, style = computedStyle(el), bounds = el.getBoundingClientRect();
        if (bounds.x < -0.5 || bounds.y < -0.5 || bounds.right > 1280.5 || bounds.bottom > 720.5) issue(`Element is outside the canvas: ${id}`, id);
        const parent = el.parentElement?.closest<HTMLElement>("[data-kind]");
        if (parent && parent.dataset.kind !== "group") invalidNode("INVALID_DESIGN_PARENT", "Only group nodes may contain design nodes", el, 'Use a data-kind="group" container for sibling shape/text objects; a shape or text node cannot contain other marked design nodes.');
        if (kind === "group" && (!transparent(style.backgroundColor) || px(style.borderTopWidth))) invalidNode("PAINTED_GROUP", "Group paint must be a separate shape node", el, 'Keep this group transparent and borderless; put its visible background in a separate data-kind="shape" child.');
        const appearance: MeasuredPage["nodes"][number]["appearance"] = { rect: { x: bounds.x, y: bounds.y, w: bounds.width, h: bounds.height } };
        if (!transparent(style.backgroundColor)) appearance.fill = color(style.backgroundColor);
        if (style.backgroundImage !== "none") {
          if (appearance.fill) throw new Error(`Use either background-color or one linear-gradient, not both: ${id}`);
          appearance.gradient = linearGradient(style.backgroundImage, id);
        }
        const borders = [style.borderTopWidth, style.borderRightWidth, style.borderBottomWidth, style.borderLeftWidth].map(px);
        const edges: MeasuredPage["nodes"] = [];
        if (borders.some(n => n > 0)) {
          const styles = [style.borderTopStyle, style.borderRightStyle, style.borderBottomStyle, style.borderLeftStyle];
          const colors = [style.borderTopColor, style.borderRightColor, style.borderBottomColor, style.borderLeftColor];
          const supported = (s: string) => ["solid", "dashed", "dotted"].includes(s);
          if (borders.every(n => n === borders[0]) && colors.every(c => c === colors[0]) && styles.every(s => s === styles[0])) {
            if (!supported(styles[0])) invalidNode("UNSUPPORTED_BORDER_STYLE", "Only solid, dashed or dotted borders are supported", el, "Use a solid, dashed or dotted border; draw other line effects as separate data-kind=\"shape\" nodes.");
            appearance.border = { color: color(colors[0]), width: borders[0], ...(styles[0] !== "solid" ? { style: styles[0] as "dashed" | "dotted" } : {}) };
          } else {
            // A border on some sides only (an accent bar, a divider, an L frame)
            // has no native single-shape form; each painted side becomes its own
            // editable shape drawn over the node, so the design is kept, not refused.
            (["top", "right", "bottom", "left"] as const).forEach((side, i) => {
              const width = borders[i];
              if (width <= 0 || ["none", "hidden"].includes(styles[i])) return;
              if (!supported(styles[i])) invalidNode("UNSUPPORTED_BORDER_STYLE", "Only solid, dashed or dotted borders are supported", el, "Use a solid, dashed or dotted border; draw other line effects as separate data-kind=\"shape\" nodes.");
              const horizontal = side === "top" || side === "bottom", far = side === "bottom" || side === "right";
              const solid = styles[i] === "solid";
              // Solid: a filled bar of the border's thickness. Dashed/dotted: a
              // zero-thickness edge outlined with that dash, centred on the border.
              const offset = solid ? (far ? (horizontal ? bounds.height : bounds.width) - width : 0) : (far ? (horizontal ? bounds.height : bounds.width) - width / 2 : width / 2);
              const rect = horizontal
                ? { x: bounds.x, y: bounds.y + offset, w: bounds.width, h: solid ? width : 0 }
                : { x: bounds.x + offset, y: bounds.y, w: solid ? width : 0, h: bounds.height };
              let edgeId = `${id.startsWith("new:") ? "" : "new:"}${id}:${side}-border`;
              for (let n = 2; document.querySelector(`[data-id="${CSS.escape(edgeId)}"]`); n++) edgeId = `${id.startsWith("new:") ? "" : "new:"}${id}:${side}-border-${n}`;
              const edgeParent = el.parentElement?.closest<HTMLElement>("[data-kind]");
              edges.push({ id: edgeId, kind: "shape", ...(edgeParent ? { parentId: edgeParent.dataset.id } : {}),
                appearance: solid ? { rect, fill: color(colors[i]) } : { rect, border: { color: color(colors[i]), width, style: styles[i] as "dashed" | "dotted" } } });
            });
          }
        }
        const radii = [style.borderTopLeftRadius, style.borderTopRightRadius, style.borderBottomLeftRadius, style.borderBottomRightRadius];
        if (radii.some(r => px(r))) {
          if (radii.some(r => r.includes(" "))) throw new Error(`Only circular corners are supported: ${id}`);
          const [tl, tr, bl, br] = radii.map(r => r.endsWith("%") ? Math.min(bounds.width, bounds.height) * px(r) / 100 : px(r));
          if (radii.every(r => r === radii[0])) appearance.radius = tl;
          else if (tl === tr && bl === br) appearance.corners = { top: tl, bottom: bl };
          else throw new Error(`Corners must be all equal, or equal along the top and along the bottom: ${id}`);
        }
        if (style.clipPath !== "none") {
          appearance.polygon = polygon(style.clipPath, bounds, id);
          if (appearance.radius !== undefined || appearance.corners || appearance.border || style.boxShadow !== "none") throw new Error(`A polygon clip-path cannot be combined with border, border-radius or box-shadow: ${id}`);
        }
        if (style.boxShadow !== "none") appearance.shadow = boxShadow(style.boxShadow, id);
        const glyphs: NonNullable<MeasuredPage["textPaint"]>[string] = [];
        let clippedText = false;
        let firstOverflow: {character: string; font: string; overflow: {left: number; top: number; right: number; bottom: number}} | undefined;
        if (kind === "image" || kind === "asset") {
          if (style.objectFit !== "fill") throw new Error(`Pre-crop the asset before using object-fit: ${id}`);
          appearance.src = (el as unknown as HTMLImageElement).src;
        } else if (kind !== "group") {
          const found = el.matches("p,h1,h2,h3,h4,h5,h6,li") ? [el] : Array.from(el.querySelectorAll<HTMLElement>("[data-paragraph],p,h1,h2,h3,h4,h5,h6,li"));
          // A paragraph marker inside another paragraph (e.g. <span data-paragraph>
          // in <p data-paragraph>) is already that paragraph's text; reading it
          // again would duplicate the words.
          const blocks = found.filter(block => !found.some(outer => outer !== block && outer.contains(block)));
          if (!blocks.length && el.textContent?.trim()) throw new Error(`Text needs a paragraph element: ${id}`);
          if (blocks.length) {
            const paragraphs = blocks.map(block => {
              const bs = computedStyle(block), box = block.getBoundingClientRect();
              const runs: NonNullable<typeof appearance.text>["paragraphs"][number]["runs"] = [];
              let previous: DOMRect | undefined;
              let previousRunParent: Element | undefined;
              const formatFor = (rs: CSSStyleDeclaration) => ({ fontFamily: rs.fontFamily.split(",")[0].trim().replace(/^["']|["']$/g, ""), fontSize: px(rs.fontSize), color: color(rs.color), bold: Number(rs.fontWeight) >= 600, italic: rs.fontStyle === "italic", letterSpacing: px(rs.letterSpacing) });
              const appendRun = (text: string, parent: Element, format: ReturnType<typeof formatFor>) => {
                // Projection gives each native run its own span, including an
                // explicit break run. Preserve those boundaries for text-only diffs.
                if (runs.length && parent === previousRunParent) runs[runs.length - 1].text += text;
                else runs.push({text, ...format});
                previousRunParent = parent;
              };
              const walk = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
              while (walk.nextNode()) {
                const child = walk.currentNode;
                if (child.nodeType === Node.ELEMENT_NODE && (child as Element).tagName === "BR") { const parent=child.parentElement!; appendRun("\u2028",parent,formatFor(computedStyle(parent))); previous = undefined; continue; }
                if (child.nodeType !== Node.TEXT_NODE) continue;
                const rs = computedStyle(child.parentElement!);
                const format = formatFor(rs);
                let content = "", offset = 0;
                for (const character of child.textContent ?? "") {
                  const range = document.createRange(); range.setStart(child, offset); offset += character.length; range.setEnd(child, offset);
                  const glyph = range.getBoundingClientRect();
                  if (glyph.width === 0 && /\s/.test(character)) continue;
                  if (previous && glyph.top > previous.top + Math.max(2, previous.height * 0.5)) content += "\u2028";
                  if (glyph.width > 0) {
                    if (!/\s/.test(character) && !transparent(rs.color)) {
                      // Range bounds include the font's em ascent/descent, which can extend
                      // outside a short line box even when every visible glyph fits. Use
                      // browser ink metrics at the DOM baseline, never model-supplied boxes.
                      const font = `${rs.fontStyle} ${rs.fontWeight} ${rs.fontSize} ${rs.fontFamily}`, key = font+"\u0000"+character;
                      let metrics = metricsCache.get(key);
                      if (!metrics) { textContext.font=font; metrics=textContext.measureText(character); metricsCache.set(key,metrics); }
                      const hasInkMetrics = [metrics.fontBoundingBoxAscent,metrics.actualBoundingBoxAscent,metrics.actualBoundingBoxDescent,metrics.actualBoundingBoxLeft,metrics.actualBoundingBoxRight].every(Number.isFinite);
                      const ink = hasInkMetrics ? {x:glyph.x-metrics.actualBoundingBoxLeft,y:glyph.y+metrics.fontBoundingBoxAscent-metrics.actualBoundingBoxAscent,w:metrics.actualBoundingBoxLeft+metrics.actualBoundingBoxRight,h:metrics.actualBoundingBoxAscent+metrics.actualBoundingBoxDescent} : {x:glyph.x,y:glyph.y,w:glyph.width,h:glyph.height};
                      if (ink.w > 0 && ink.h > 0) {
                        if (ink.x < bounds.left - 1 || ink.x+ink.w > bounds.right + 1 || ink.y < bounds.top - 1 || ink.y+ink.h > bounds.bottom + 1) {
                          clippedText = true;
                          firstOverflow ??= {character,font,overflow:{left:Math.max(0,bounds.left-ink.x),top:Math.max(0,bounds.top-ink.y),right:Math.max(0,ink.x+ink.w-bounds.right),bottom:Math.max(0,ink.y+ink.h-bounds.bottom)}};
                        }
                        glyphs.push(ink);
                      }
                    }
                    previous = glyph;
                  }
                  content += character;
                }
                if (content) appendRun(content,child.parentElement!,format);
              }
              if (!runs.length) runs.push({ text: "", fontFamily: bs.fontFamily.split(",")[0].replace(/"/g, ""), fontSize: px(bs.fontSize), color: color(bs.color), bold: Number(bs.fontWeight) >= 600, italic: bs.fontStyle === "italic", letterSpacing: px(bs.letterSpacing) });
              const lineHeight = bs.lineHeight === "normal" ? px(bs.fontSize) * 1.2 : px(bs.lineHeight);
              return { runs, align: (["left", "center", "right"].includes(bs.textAlign) ? bs.textAlign : "left") as "left" | "center" | "right", lineHeight, before: px(bs.marginTop), after: px(bs.marginBottom) };
            });
            appearance.text = { paragraphs, padding: [px(style.paddingTop), px(style.paddingRight), px(style.paddingBottom), px(style.paddingLeft)], verticalAlign: style.justifyContent === "center" ? "middle" : style.justifyContent === "flex-end" ? "bottom" : "top" };
            const boxes = blocks.map(block => block.getBoundingClientRect());
            const left = bounds.left + px(style.borderLeftWidth) + px(style.paddingLeft);
            const width = bounds.width - px(style.borderLeftWidth) - px(style.borderRightWidth) - px(style.paddingLeft) - px(style.paddingRight);
            const height = bounds.height - px(style.borderTopWidth) - px(style.borderBottomWidth) - px(style.paddingTop) - px(style.paddingBottom);
            const flowHeight = boxes.reduce((sum, box, i) => sum + box.height + paragraphs[i].before + paragraphs[i].after, 0);
            let y = bounds.top + px(style.borderTopWidth) + px(style.paddingTop)
              + (appearance.text.verticalAlign === "middle" ? (height - flowHeight) / 2 : appearance.text.verticalAlign === "bottom" ? height - flowHeight : 0);
            const separateFrames = boxes.some((box, i) => {
              y += paragraphs[i].before;
              const differs = Math.abs(box.x - left) > 1 || Math.abs(box.width - width) > 1 || Math.abs(box.y - y) > 1;
              y += box.height + paragraphs[i].after;
              return differs;
            });
            if (separateFrames) appearance.text.frames = boxes.map((box, i) => {
              const block = computedStyle(blocks[i]);
              return {rect:{x:box.x,y:box.y,w:box.width,h:box.height},padding:[px(block.paddingTop),px(block.paddingRight),px(block.paddingBottom),px(block.paddingLeft)]};
            });
          }
        }
        if (appearance.text) {
          textPaint[id] = glyphs;
          if (clippedText && glyphs.length) {
            const left=Math.min(...glyphs.map(r=>r.x)), top=Math.min(...glyphs.map(r=>r.y)), right=Math.max(...glyphs.map(r=>r.x+r.w)), bottom=Math.max(...glyphs.map(r=>r.y+r.h));
            issue(`Text is clipped or outside its box: ${id}. ${JSON.stringify({nodeId:id,nodeBounds:appearance.rect,textBounds:{x:left,y:top,w:right-left,h:bottom-top},requiredTextHeight:bottom-top,overflow:{left:Math.max(0,bounds.left-left),top:Math.max(0,bounds.top-top),right:Math.max(0,right-bounds.right),bottom:Math.max(0,bottom-bounds.bottom)},firstOverflow,lineHeights:appearance.text.paragraphs.map(p=>p.lineHeight),recovery:"Layout warning only; applying HTML is allowed. Overflow is relative to this text box's own edges; moving the entire box does not change it. Inspect the actual rendering, then adjust padding, width, height, line height or wording if visibly clipped. Preserve fixed assets and data-* metadata."})}`,id);
          }
        }
        nodes.push({ id, kind, ...(parent ? { parentId: parent.dataset.id } : {}), appearance }, ...edges);
      }
      return { pageId, revision, background: transparent(computedStyle(body).backgroundColor) ? "#ffffff" : color(computedStyle(body).backgroundColor), sourceHash: body.dataset.sourceHash, nodes, textPaint, ...(collectIssues ? { issues } : {}) };
}
