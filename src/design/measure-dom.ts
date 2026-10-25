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
      const markerElements = Array.from(document.querySelectorAll<HTMLElement>("[data-kind]"));
      const invalidNode = (code: string, message: string, el: HTMLElement, recovery: string): never => {
        const owner = el.closest<HTMLElement>("[data-kind]");
        throw new Error(`${message}. ${JSON.stringify({
          code, pageId, revision, tag: el.tagName.toLowerCase(), nodeId: el.dataset.id ?? null,
          ownerId: owner?.dataset.id ?? null, ownerKind: owner?.dataset.kind ?? null,
          ownerIndex: owner ? markerElements.indexOf(owner) : null,
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
        for (const pseudo of ["::before", "::after"]) { const content = computedStyle(el, pseudo).content; if (content && content !== "none" && content !== "normal" && content !== '""') throw new Error("Pseudo-element content is not supported; use a marked node"); }
        if (style.transform !== "none" || style.scale && style.scale !== "none" || style.rotate && style.rotate !== "none" || style.translate && style.translate !== "none" || style.filter !== "none" || style.backdropFilter && style.backdropFilter !== "none" || style.mixBlendMode !== "normal" || Number(style.opacity) !== 1 || style.boxShadow !== "none" || style.backgroundImage !== "none" || style.zIndex !== "auto" || style.clipPath !== "none" || style.textShadow !== "none" || style.clip !== "auto" || style.maskImage && style.maskImage !== "none" || style.contain.includes("paint") || style.contentVisibility !== "visible" || el.tagName !== "IMG" && [style.overflowX,style.overflowY].some(v=>v!=="visible")) throw new Error(`Unsupported CSS on ${el.dataset.id ?? el.tagName}; use a separate image/SVG asset`);
        const owner = el.closest<HTMLElement>("[data-kind]");
        if (el !== body && !el.dataset.kind && (!transparent(style.backgroundColor) || px(style.borderTopWidth) || px(style.borderLeftWidth))) throw new Error("Painted containers need data-kind and data-id");
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
        if (parent && parent.dataset.kind !== "group") throw new Error(`Only group nodes may contain design nodes: ${id}`);
        if (kind === "group" && (!transparent(style.backgroundColor) || px(style.borderTopWidth))) throw new Error(`Group paint must be a separate shape node: ${id}`);
        const appearance: MeasuredPage["nodes"][number]["appearance"] = { rect: { x: bounds.x, y: bounds.y, w: bounds.width, h: bounds.height } };
        if (!transparent(style.backgroundColor)) appearance.fill = color(style.backgroundColor);
        const borders = [style.borderTopWidth, style.borderRightWidth, style.borderBottomWidth, style.borderLeftWidth].map(px);
        if (borders.some(n => n > 0)) {
          if (!borders.every(n => n === borders[0]) || ![style.borderTopColor, style.borderRightColor, style.borderBottomColor, style.borderLeftColor].every(c => c === style.borderTopColor) || style.borderTopStyle !== "solid") throw new Error(`Only uniform solid borders are supported: ${id}`);
          appearance.border = { color: color(style.borderTopColor), width: borders[0] };
        }
        const radii = [style.borderTopLeftRadius, style.borderTopRightRadius, style.borderBottomLeftRadius, style.borderBottomRightRadius];
        if (radii.some(r => px(r))) {
          if (!radii.every(r => r === radii[0]) || radii[0].includes(" ")) throw new Error(`Only equal circular corners are supported: ${id}`);
          appearance.radius = radii[0].endsWith("%") ? Math.min(bounds.width, bounds.height) * px(radii[0]) / 100 : px(radii[0]);
        }
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
        nodes.push({ id, kind, ...(parent ? { parentId: parent.dataset.id } : {}), appearance });
      }
      return { pageId, revision, background: transparent(computedStyle(body).backgroundColor) ? "#ffffff" : color(computedStyle(body).backgroundColor), sourceHash: body.dataset.sourceHash, nodes, textPaint, ...(collectIssues ? { issues } : {}) };
}
