import { cloneDesign, type DesignDocument } from "./model";
import { assertDesignHtmlIdentity, parseDesignHtml } from "./html-document";
import type { TemplatePageFill, TemplatePurpose } from "./template-policy";

export interface TemplateSemanticChange {
  nodeId?: string;
  name: string;
  before: Record<string, string | undefined>;
  after: Record<string, string | undefined>;
}
export interface TemplateSemanticProposal { sourceHash: string; revision: number; pageId: string; changes: TemplateSemanticChange[] }
export interface TemplateSemanticAuthorization extends TemplateSemanticProposal { mode: "template-authoring" }
export interface TemplateSemanticEdit { nodeId?: string; purpose?: TemplatePurpose; note?: string; slot?: string; role?: string; layoutName?: string; fill?: TemplatePageFill }
const nodeFields = { purpose: "data-purpose", note: "data-note", slot: "data-slot" };
const pageFields = { role: "data-page-role", layoutName: "data-layout-name", note: "data-page-note", fill: "data-page-fill" };
const values = (el: Element, fields: Record<string,string>): Record<string,string|undefined> => Object.fromEntries(Object.entries(fields).map(([key,attr])=>[key,el.getAttribute(attr) || undefined]));
const equal = (a: unknown,b: unknown) => JSON.stringify(a) === JSON.stringify(b);
interface ReadonlyHtmlChange { nodeId: string; name: string; attribute: string; original: string | null; submitted: string | null }
export interface NewNodeAnnotation {
  nodeId?: string; nodeIndex: number; name: string; tag: string; kind: string;
  textSnippet: string; attributes: Array<{attribute: string; submitted: string}>;
}
const newNodeAttributes = ["data-purpose", "data-note", "data-slot", "data-locked", "data-bind", "data-required", "data-role", "data-max-chars"];
function newNodeAnnotation(el: HTMLElement, nodeIndex: number): NewNodeAnnotation {
  return {nodeId:el.dataset.id,nodeIndex,name:el.dataset.name || el.dataset.id || `new node ${nodeIndex + 1}`,
    tag:el.tagName.toLowerCase(),kind:el.dataset.kind || "",textSnippet:(el.textContent || "").replace(/\s+/g," ").trim().slice(0,160),
    attributes:newNodeAttributes.filter(attr=>el.hasAttribute(attr) && !(attr === "data-purpose" && el.getAttribute(attr) === "content"))
      .map(attribute=>({attribute,submitted:el.getAttribute(attribute)!}))};
}
/** AI-created objects are ordinary editable content; copied template annotations
 * are not new author permissions. Existing objects and page annotations stay intact. */
export function normalizeNewContentAnnotations(baselineHtml: string, html: string): {html: string; cleaned: NewNodeAnnotation[]} {
  const baseline = parseDesignHtml(baselineHtml), edited = parseDesignHtml(html);
  const ids = new Set(Array.from(baseline.querySelectorAll<HTMLElement>("[data-kind][data-id]")).map(el=>el.dataset.id!));
  const cleaned: NewNodeAnnotation[] = [];
  Array.from(edited.querySelectorAll<HTMLElement>("[data-kind]")).forEach((el,index)=>{
    if (el.dataset.id && ids.has(el.dataset.id)) return;
    const issue = newNodeAnnotation(el,index);
    if (!issue.attributes.length) return;
    for (const {attribute} of issue.attributes) {
      if (attribute === "data-purpose") el.dataset.purpose = "content";
      else el.removeAttribute(attribute);
    }
    cleaned.push(issue);
  });
  return {html:cleaned.length ? "<!doctype html>" + edited.documentElement.outerHTML : html,cleaned};
}
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
  const newNodeAnnotations: NewNodeAnnotation[] = [];
  const before = values(baseline.body,pageFields), after = values(edited.body,pageFields);
  if (after.fill && !["layout","free"].includes(after.fill)) throw new Error(`Invalid data-page-fill: ${after.fill}; use layout or free`);
  if (!equal(before,after)) changes.push({name:pageId,before,after});
  const byId = new Map(Array.from(baseline.querySelectorAll<HTMLElement>("[data-id]")).map(n=>[n.dataset.id!,n]));
  for (const [index,el] of Array.from(edited.querySelectorAll<HTMLElement>("[data-kind]")).entries()) {
    const old = byId.get(el.dataset.id!);
    if (!old) {
      const issue = newNodeAnnotation(el,index);
      if (issue.attributes.length) newNodeAnnotations.push(issue);
      continue;
    }
    for (const attr of ["data-locked","data-bind","data-required","data-max-chars","data-role","data-name"]) if (old.getAttribute(attr) !== el.getAttribute(attr)) readonlyChanges.push({nodeId:el.dataset.id!,name:page.nodes.find(n=>n.id===el.dataset.id)?.name ?? el.dataset.id!,attribute:attr,original:old.getAttribute(attr),submitted:el.getAttribute(attr)});
    const before = values(old,nodeFields), after = values(el,nodeFields);
    if (!equal(before,after)) {
      if (!["fixed","content","reference"].includes(after.purpose ?? "")) throw new Error(`Invalid or missing data-purpose: ${el.dataset.id}`);
      changes.push({nodeId:el.dataset.id,name:page.nodes.find(n=>n.id===el.dataset.id)?.name ?? el.dataset.id!,before,after});
    }
  }
  if (newNodeAnnotations.length) {
    const first = newNodeAnnotations[0];
    throw new Error(`New node "${first.name}" (${first.tag}, index ${first.nodeIndex}) carries template attributes: ${first.attributes.map(a=>a.attribute).join(", ")}. ${JSON.stringify({
      code:"NEW_NODE_TEMPLATE_ANNOTATIONS",phase:"html_validation",applied:false,pageId,revision:design.revision,nodes:newNodeAnnotations,
      recovery:"On newly created nodes omit data-note, data-slot and copied permission/binding attributes; purpose may only be content. Keep all existing node IDs and metadata unchanged. No page content was applied; correct the listed nodes and retry edit_html."
    })}`);
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
      node.purpose = change.after.purpose as TemplatePurpose; node.note = change.after.note; node.slot = change.after.slot;
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
