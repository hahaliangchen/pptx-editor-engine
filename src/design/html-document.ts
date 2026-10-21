import type { DesignDocument } from "./model";

/** Inspect in a truly inert template before DOMParser can load an image URL. */
export function parseDesignHtml(source: string): Document {
  const template = document.createElement("template");
  template.innerHTML = source;
  const root = template.content;
  if (root.querySelector("script,iframe,object,embed,link,base,meta[http-equiv],svg,math,video,audio,source,track,input")) {
    throw new Error("Only static HTML/CSS and embedded image assets are accepted");
  }
  for (const element of Array.from(root.querySelectorAll("*"))) {
    if (Array.from(element.attributes).some(a => /^on/i.test(a.name) || ["srcset", "poster", "background"].includes(a.name.toLowerCase()))) {
      throw new Error("Active HTML attributes are not allowed");
    }
    if (element.hasAttribute("src") && (element.tagName !== "IMG"
      || !/^(data:image\/|design-asset:|project:images\/)/.test(element.getAttribute("src") ?? ""))) {
      throw new Error("Images must use embedded project assets");
    }
  }
  return new DOMParser().parseFromString(source, "text/html");
}

/** Validate host-issued identity exactly; never repair or abbreviate a model-supplied hash. */
export function assertDesignHtmlIdentity(design: DesignDocument, body: HTMLElement, expectedPageId?: string): void {
  const pageId=body.dataset.pageId, hash=body.dataset.sourceHash, revision=Number(body.dataset.designRevision);
  const reason=!pageId || !design.pages.some(p=>p.id===pageId) || expectedPageId && expectedPageId!==pageId ? "pageId"
    : !body.hasAttribute("data-design-revision") || !Number.isInteger(revision) || revision!==design.revision ? "revision"
    : !/^[a-f0-9]{64}$/.test(hash ?? "") ? "sourceHashFormat"
    : hash!==design.source.sha256 ? "sourceHashMismatch" : undefined;
  if (!reason) return;
  const stale=reason==="pageId" || reason==="revision", code=stale?"STALE_DESIGN":"INVALID_DESIGN_IDENTITY";
  throw new Error(`${code}: ${stale?"page or revision is no longer current":"source hash was omitted, abbreviated or differs from the current design"}. ${JSON.stringify({code,reason,expected:{pageId:expectedPageId ?? (design.pages.some(p=>p.id===pageId)?pageId:undefined),revision:design.revision,sourceHash:design.source.sha256},received:{pageId:pageId??null,revision:body.dataset.designRevision??null,sourceHash:hash??null,sourceHashLength:hash?.length??0},recovery:stale?"Call get_design for the current page and redo the requested changes on that latest HTML. Do not just relabel old HTML with a new revision. No content was applied.":"Copy all identity attributes verbatim from the latest get_design HTML; the source hash must be all 64 characters, never shortened or rewritten. If the page/revision has since changed, reread and redo your changes on the latest HTML. Then retry apply_html; no content was applied."})}`);
}
