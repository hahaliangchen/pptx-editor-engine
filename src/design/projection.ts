import { effectiveLock, effectivePurpose } from "./template-policy";
import { assertDesign, DesignDocument, DesignNode, DesignRun } from "./model";

export const escapeHtml = (value: string): string => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const cssNumber = (n: number) => String(Math.round(n * 1000) / 1000);
// The paragraph's invisible CSS strut must use its native run font too.
// A browser-default 16px serif strut can enlarge explicit CJK line spacing.
const paragraphFontStyle = (run: DesignRun) => `font-family:${JSON.stringify(run.fontFamily)};font-size:${cssNumber(run.fontSize)}px;font-weight:${run.bold ? 700 : 400};font-style:${run.italic ? "italic" : "normal"};`;
const runStyle = (run: DesignRun) => `font-family:${JSON.stringify(run.fontFamily)};font-size:${cssNumber(run.fontSize)}px;color:${run.color};font-weight:${run.bold ? 700 : 400};font-style:${run.italic ? "italic" : "normal"};letter-spacing:${cssNumber(run.letterSpacing)}px`;

/** Full absolute projection for editing. Redesign can replace its layout with flex/grid, keeping IDs. */
export function projectPage(document: DesignDocument, pageId: string): string {
  assertDesign(document);
  const page = document.pages.find(p => p.id === pageId);
  if (!page) throw new Error("Unknown design page");
  const byId = new Map(page.nodes.map(n => [n.id, n]));
  const project = (node: DesignNode): string => {
    if (node.hidden) return "";
    const a = node.appearance, r = a.rect;
    const parentRect = node.parentId ? byId.get(node.parentId)!.appearance.rect : { x: 0, y: 0 };
    const properties = [`position:absolute`, `left:${cssNumber(r.x - parentRect.x)}px`, `top:${cssNumber(r.y - parentRect.y)}px`, `width:${cssNumber(r.w)}px`, `height:${cssNumber(r.h)}px`];
    if (a.fill) properties.push(`background-color:${a.fill}`);
    if (a.border) properties.push(`border:${cssNumber(a.border.width)}px solid ${a.border.color}`);
    if (a.radius !== undefined) properties.push(`border-radius:${cssNumber(a.radius)}px`);
    const lock = effectiveLock(page, node);
    const data = `data-id="${escapeHtml(node.id)}" data-kind="${node.kind}" data-name="${escapeHtml(node.name)}"`
      + (lock ? ` data-locked="${lock}"` : "")
      + ` data-purpose="${effectivePurpose(page, node)}"`
      + (node.note ? ` data-note="${escapeHtml(node.note)}"` : "")
      + (node.role ? ` data-role="${escapeHtml(node.role)}"` : "")
      + (node.binding !== undefined ? ` data-bind="${escapeHtml(JSON.stringify(node.binding))}"` : "")
      + (node.required ? ` data-required="true"` : "")
      + (node.constraints?.maxChars !== undefined ? ` data-max-chars="${node.constraints.maxChars}"` : "");
    if (node.kind === "image" || node.kind === "asset") return `<img ${data} style="${escapeHtml(properties.join(";"))}" src="${escapeHtml(a.src ?? "")}" alt="${escapeHtml(node.name)}">`;
    let content = "";
    if (node.kind === "group") content = page.nodes.filter(n => n.parentId === node.id).map(project).join("\n");
    else if (a.text) {
      properties.push("display:flex", "flex-direction:column", `justify-content:${({ top: "flex-start", middle: "center", bottom: "flex-end" })[a.text.verticalAlign]}`, `padding:${a.text.padding.map(n => `${cssNumber(n)}px`).join(" ")}`);
      content = a.text.paragraphs.map((p, i) => `<p data-paragraph="true" style="${p.runs[0] ? escapeHtml(paragraphFontStyle(p.runs[0])) : ""}${a.text!.frames ? `position:absolute;left:${cssNumber(a.text!.frames[i].rect.x-r.x-(a.border?.width??0))}px;top:${cssNumber(a.text!.frames[i].rect.y-r.y-(a.border?.width??0))}px;width:${cssNumber(a.text!.frames[i].rect.w)}px;height:${cssNumber(a.text!.frames[i].rect.h)}px;padding:${a.text!.frames[i].padding.map(n=>`${cssNumber(n)}px`).join(" ")};margin:0;` : `margin:${cssNumber(p.before)}px 0 ${cssNumber(p.after)}px;`}text-align:${p.align};line-height:${cssNumber(p.lineHeight)}px;white-space:pre-wrap;overflow-wrap:normal">${p.runs.map(run => `<span style="${escapeHtml(runStyle(run))}">${escapeHtml(run.text).replace(/\u2028/g, "<br>")}</span>`).join("")}</p>`).join("");
    }
    return `<div ${data} style="${escapeHtml(properties.join(";"))}">${content}</div>`;
  };
  return `<!doctype html><html><head><meta charset="utf-8"><style>*{box-sizing:border-box}html,body{margin:0;width:1280px;height:720px}body{position:relative;background:${escapeHtml(page.background ?? "#ffffff")}}p{margin:0;flex-shrink:0}img{object-fit:fill}</style></head><body data-page-id="${escapeHtml(page.id)}" data-source-hash="${document.source.sha256}" data-design-revision="${document.revision}"${page.semantics?.role ? ` data-page-role="${escapeHtml(page.semantics.role)}"` : ""}${page.semantics?.layoutName ? ` data-layout-name="${escapeHtml(page.semantics.layoutName)}"` : ""}${page.semantics?.note ? ` data-page-note="${escapeHtml(page.semantics.note)}"` : ""}>${page.nodes.filter(n => !n.parentId).map(project).join("\n")}</body></html>`;
}
