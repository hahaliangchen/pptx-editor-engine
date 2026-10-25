import { cloneDesign, type DesignDocument } from "./model";
import { assertDesignHtmlIdentity, parseDesignHtml } from "./html-document";
import type { TemplatePurpose } from "./template-policy";

export interface TemplateSemanticChange {
  nodeId?: string;
  name: string;
  before: Record<string, string | undefined>;
  after: Record<string, string | undefined>;
}
export interface TemplateSemanticProposal { sourceHash: string; revision: number; pageId: string; changes: TemplateSemanticChange[] }
export interface TemplateSemanticAuthorization extends TemplateSemanticProposal { mode: "template-authoring" }
export interface TemplateSemanticEdit { nodeId?: string; purpose?: TemplatePurpose; note?: string; role?: string; layoutName?: string }
const nodeFields = { purpose: "data-purpose", note: "data-note" };
const pageFields = { role: "data-page-role", layoutName: "data-layout-name", note: "data-page-note" };
const values = (el: Element, fields: Record<string,string>): Record<string,string|undefined> => Object.fromEntries(Object.entries(fields).map(([key,attr])=>[key,el.getAttribute(attr) || undefined]));
const equal = (a: unknown,b: unknown) => JSON.stringify(a) === JSON.stringify(b);
interface ReadonlyHtmlChange { nodeId: string; name: string; attribute: string; original: string | null; submitted: string | null }
const metadataRecovery = "No page content was applied. Use the latest get_design HTML as the baseline, preserve every existing data-* attribute exactly, edit only the allowed content/style, and retry apply_html. Do not recreate the deck or stop after this recoverable validation error.";
/** Include every changed permission field, so a model can fix one rejected submission without guessing. */
export function templateSemanticPermissionError(proposal: TemplateSemanticProposal, reason: string): Error {
  return new Error(`TEMPLATE_PERMISSION: ${reason}. ${JSON.stringify({code:"TEMPLATE_PERMISSION",pageId:proposal.pageId,revision:proposal.revision,semanticChanges:proposal.changes,recovery:metadataRecovery})}`);
}
/** Compare only host-issued projection metadata, never trust an HTML unlock flag. */
export function templateSemanticProposal(design: DesignDocument, baselineHtml: string, html: string): TemplateSemanticProposal {
  const baseline = parseDesignHtml(baselineHtml), edited = parseDesignHtml(html), pageId = baseline.body.dataset.pageId!;
  assertDesignHtmlIdentity(design,baseline.body);
  assertDesignHtmlIdentity(design,edited.body,pageId);
  const page = design.pages.find(p=>p.id===pageId)!;
  const changes: TemplateSemanticChange[] = [];
  const readonlyChanges: ReadonlyHtmlChange[] = [];
  const before = values(baseline.body,pageFields), after = values(edited.body,pageFields);
  if (!equal(before,after)) changes.push({name:pageId,before,after});
  const byId = new Map(Array.from(baseline.querySelectorAll<HTMLElement>("[data-id]")).map(n=>[n.dataset.id!,n]));
  for (const el of Array.from(edited.querySelectorAll<HTMLElement>("[data-kind]"))) {
    const old = byId.get(el.dataset.id!);
    if (!old) { if (el.dataset.purpose && el.dataset.purpose !== "content" || el.dataset.note) throw new Error("New AI nodes use content purpose; annotate them after creation"); continue; }
    for (const attr of ["data-locked","data-bind","data-required","data-max-chars","data-role","data-name"]) if (old.getAttribute(attr) !== el.getAttribute(attr)) readonlyChanges.push({nodeId:el.dataset.id!,name:page.nodes.find(n=>n.id===el.dataset.id)?.name ?? el.dataset.id!,attribute:attr,original:old.getAttribute(attr),submitted:el.getAttribute(attr)});
    const before = values(old,nodeFields), after = values(el,nodeFields);
    if (!equal(before,after)) {
      if (!["fixed","content","reference"].includes(after.purpose ?? "")) throw new Error(`Invalid or missing data-purpose: ${el.dataset.id}`);
      changes.push({nodeId:el.dataset.id,name:page.nodes.find(n=>n.id===el.dataset.id)?.name ?? el.dataset.id!,before,after});
    }
  }
  if (readonlyChanges.length) throw new Error(`Read-only HTML attributes cannot change. ${JSON.stringify({code:"TEMPLATE_HTML_METADATA",pageId,revision:design.revision,readonlyChanges,semanticChanges:changes,recovery:metadataRecovery})}`);
  for (const change of changes) for (const value of Object.values(change.after)) if (value && value.length > 4096) throw new Error("Template annotations must be at most 4096 characters");
  return {sourceHash:design.source.sha256,revision:design.revision,pageId,changes};
}
export function prepareTemplateSemantics(design: DesignDocument, baselineHtml: string, html: string, authorization?: TemplateSemanticAuthorization): {design:DesignDocument; proposal:TemplateSemanticProposal} {
  const proposal = templateSemanticProposal(design,baselineHtml,html);
  if (!proposal.changes.length) return {design,proposal};
  if (!authorization || authorization.mode !== "template-authoring" || !equal({pageId:proposal.pageId,changes:proposal.changes},{pageId:authorization.pageId,changes:authorization.changes})) throw templateSemanticPermissionError(proposal,"template purpose changes need the author's confirmation for these exact objects and changes");
  const next = cloneDesign(design), page = next.pages.find(p=>p.id===proposal.pageId)!;
  page.semanticsChanged = true;
  for (const change of proposal.changes) {
    if (change.nodeId) {
      const node = page.nodes.find(n=>n.id===change.nodeId)!;
      if (node.metadata?.backdrop || node.hidden) throw new Error(`Unsupported/master object cannot be unlocked here: ${node.id}`);
      node.purpose = change.after.purpose as TemplatePurpose; node.note = change.after.note;
    } else page.semantics = {...change.after};
  }
  return {design:next,proposal};
}
export function finishTemplateSemantics(original:DesignDocument, merged:DesignDocument, proposal:TemplateSemanticProposal):DesignDocument {
  if (!proposal.changes.length) return merged;
  if (merged.revision === original.revision) merged.revision++;
  merged.changes = [...(merged.changes ?? []),...proposal.changes.map(change=>({pageId:proposal.pageId,nodeId:change.nodeId,action:"update" as const,fields:Object.keys(change.after).filter(k=>change.before[k]!==change.after[k]).map(k=>`semantics.${k}`)}))];
  return merged;
}
