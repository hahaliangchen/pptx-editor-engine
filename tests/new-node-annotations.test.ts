// @vitest-environment node
import {readFileSync} from "node:fs";
import {beforeAll,expect,it} from "vitest";
import {syncDesign,convertDesign} from "../src/design/service";
import {projectPage} from "../src/design/projection";
import {normalizeNewContentAnnotations,templateSemanticProposal} from "../src/design/template-semantics";
import {DesignBridge} from "../src/page/design-bridge";
import {nodeText,type DesignDocument} from "../src/design/model";
let design:DesignDocument, baseline:string;
const extra='<div data-kind="text" data-name="TextBox 28" data-purpose="reference" data-note="目录章节标题" data-slot="章节#2" data-locked="all" data-bind="title" data-required="true" data-role="title" data-max-chars="24" style="position:absolute;left:100px;top:300px;width:400px;height:60px"><p style="margin:0;font-size:24px">新增章节</p></div>';
const payload=(r:any)=>{expect(r.ok,JSON.stringify(r)).toBe(true);return JSON.parse(r.content[0].text);};
beforeAll(async()=>{design=await syncDesign(readFileSync("tests/fixtures/unicom-uniai-template.pptx"));baseline=projectPage(design,design.pages[0].id);},30000);
it("cleans copied annotations only on newly created nodes and identifies all offending attributes",()=>{
  const inserted=baseline.replace("</body>",extra+"</body>");
  const result=normalizeNewContentAnnotations(baseline,inserted);
  const doc=new DOMParser().parseFromString(result.html,"text/html");
  const added=doc.querySelector('[data-name="TextBox 28"]')!;
  expect(added.getAttribute("data-purpose")).toBe("content");
  for(const a of ["data-note","data-slot","data-locked","data-bind","data-required","data-role","data-max-chars"])expect(added.hasAttribute(a)).toBe(false);
  expect(result.cleaned[0]).toMatchObject({name:"TextBox 28",kind:"text",textSnippet:"新增章节"});
  expect(result.cleaned[0].attributes).toContainEqual({attribute:"data-note",submitted:"目录章节标题"});
  expect(result.cleaned[0].attributes).toHaveLength(8);
  expect(templateSemanticProposal(design,baseline,result.html).changes).toEqual([]);
  const original=new DOMParser().parseFromString(baseline,"text/html");
  for(const el of Array.from(original.querySelectorAll('[data-id]')))expect(doc.querySelector(`[data-id="${el.getAttribute('data-id')}"]`)?.outerHTML).toBe(el.outerHTML);
  expect(normalizeNewContentAnnotations(baseline,baseline)).toEqual({html:baseline,cleaned:[]});
});
it("reports the actual new node, text and attribute values; existing permissions still reject changes",()=>{
  let message="";try{templateSemanticProposal(design,baseline,baseline.replace("</body>",extra+"</body>"));}catch(e){message=(e as Error).message;}
  const issue=JSON.parse(message.slice(message.indexOf("{")));
  expect(issue.code).toBe("NEW_NODE_TEMPLATE_ANNOTATIONS");expect(issue.nodes[0].name).toBe("TextBox 28");
  expect(issue.nodes[0].attributes).toContainEqual({attribute:"data-slot",submitted:"章节#2"});
  const doc=new DOMParser().parseFromString(baseline,"text/html");
  const fixed=doc.querySelector('[data-locked="all"]')!;expect(fixed).toBeTruthy();fixed.setAttribute("data-locked","content");
  const normalized=normalizeNewContentAnnotations(baseline,"<!doctype html>"+doc.documentElement.outerHTML);
  expect(()=>templateSemanticProposal(design,baseline,normalized.html)).toThrow('"attribute":"data-locked"');
});
it("saves and reopens a source patch with copied annotations as ordinary editable content",async()=>{
  let bytes:Uint8Array=readFileSync("tests/fixtures/unicom-uniai-template.pptx"),saves=0;
  const bridge=new DesignBridge({snapshotPptx:async()=>new Blob([bytes as BlobPart]),loadPptx:async b=>{bytes=new Uint8Array(b);},getCurrentSlideIndex:()=>0,getSlidesCount:()=>4,selectSlide:async()=>{}},convertDesign,undefined,{persist:async()=>++saves});
  try{
    const before=payload(await bridge.run("get_design",{page:1,html:true}));
    const result=payload(await bridge.run("edit_html",{page:1,edits:[{old_string:"</body>",new_string:extra+"</body>"}]}));
    expect(result.saved).toBe(true);expect(result.annotationCleanups[0].name).toBe("TextBox 28");expect(saves).toBe(1);
    const reopened=await syncDesign(bytes),added=reopened.pages[0].nodes.find(n=>nodeText(n)==="新增章节")!;
    expect(added).toBeTruthy();expect(added.purpose).toBe("content");expect(added.note).toBeUndefined();expect(added.slot).toBeUndefined();
    const current=projectPage(reopened,reopened.pages[0].id);
    const oldDoc=new DOMParser().parseFromString(before.html,"text/html"),newDoc=new DOMParser().parseFromString(current,"text/html");
    for(const old of Array.from(oldDoc.querySelectorAll('[data-id]'))){const kept=newDoc.querySelector(`[data-id="${old.getAttribute('data-id')}"]`)!;expect(kept).toBeTruthy();for(const attr of ["data-note","data-purpose","data-slot","data-locked"])expect(kept.getAttribute(attr)).toBe(old.getAttribute(attr));}
  }finally{bridge.dispose();}
},30000);
