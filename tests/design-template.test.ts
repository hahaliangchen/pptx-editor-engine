// @vitest-environment node
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { syncDesign, convertDesign, materializeDesign } from "../src/design/service";
import { projectPage } from "../src/design/projection";
import { mergeHtml } from "../src/design/measure";
import { HeadlessSession } from "../src/headless/session";
import { DesignBridge } from "../src/page/design-bridge";
import { cloneDesign, mergeMeasuredPage, nodeText, type DesignDocument } from "../src/design/model";
import { templateSemanticProposal } from "../src/design/template-semantics";
import { readTemplateRecord, validateTemplateRecord, effectivePurpose } from "../src/design/template-policy";
const source = readFileSync("tests/fixtures/unicom-uniai-template.pptx");
const payload = (r:any) => { expect(r.ok,JSON.stringify(r)).toBe(true); return JSON.parse(r.content[0].text); };
const change = (html:string,id:string,fn:(e:HTMLElement)=>void) => {const d = new DOMParser().parseFromString(html,"text/html");fn(Array.from(d.querySelectorAll<HTMLElement>("[data-id]")).find(n=>n.dataset.id===id)!);return "<!doctype html>"+d.documentElement.outerHTML;};
const logo = (d:DesignDocument) => d.pages[0].nodes.find(n=>n.kind==="image")!;

describe("template HTML policy and closed-loop persistence",()=>{
  it("preserves the logo and reports foreground cover as a layout warning",async()=>{
    const d=await syncDesign(source), brand=logo(d), html=projectPage(d,d.pages[0].id);
    expect(brand.purpose).toBe("fixed");
    for(const mutate of [(e:HTMLElement)=>e.style.left="700px",(e:HTMLElement)=>e.remove(),(e:HTMLElement)=>e.setAttribute("src",d.pages[0].nodes.find(n=>n.kind==="asset")!.appearance.src!)]) await expect(mergeHtml(d,change(html,brand.id,mutate))).rejects.toThrow(/Locked|locked|Fixed/);
    const covered=await mergeHtml(d,html.replace("</body>",'<div data-kind="shape" data-id="new:cover" style="position:absolute;left:780px;top:5px;width:200px;height:80px;background:white"></div></body>'));
    expect(covered.diagnostics.some(d=>d.severity==="warning" && d.message.includes("FIXED_OCCLUSION"))).toBe(true);
    const saved=await materializeDesign(covered);
    expect(saved.diagnostics.some(d=>d.severity==="warning" && d.message.includes("FIXED_OCCLUSION"))).toBe(true);
    await expect(mergeHtml(d,change(html,brand.id,e=>{const wrap=e.ownerDocument.createElement("div");wrap.style.cssText="position:absolute;left:755px;top:0;width:20px;height:20px;overflow:hidden";e.before(wrap);wrap.appendChild(e);}))).rejects.toThrow(/Unsupported CSS/);
    // Existing edge overflow is kept; editing an unrelated title is safe.
    const next=await convertDesign({operation:"apply",design:d,expectedRevision:d.revision,html:html.replace("{{TITLE}}","共同成长")});
    expect(logo(next).appearance).toEqual(brand.appearance);
  },30000);

  it("allows title edits and reports actual glyph or fill occlusion for visual review",async()=>{
    const d=await syncDesign(source), page=d.pages.find(p=>p.part==="ppt/slides/slide3.xml")!;
    const title=page.nodes.find(n=>n.origin?.shapeId==="8")!, baseline=projectPage(d,page.id);
    const edited=(width:number,text:string,size:number,height=60.8,left=50.2,fill="")=>change(baseline,title.id,e=>{
      e.style.cssText=`position:absolute;left:${left}px;top:21.7px;width:${width}px;height:${height}px;display:flex;flex-direction:column;padding:0;${fill}`;
      e.innerHTML=`<p style="margin:0;line-height:45.6px;white-space:pre-wrap"><span style="font-family:'Noto Sans CJK SC';font-size:${size}px;color:#000;font-weight:700">${text}</span></p>`;
    });
    const safe=await mergeHtml(d,edited(740,"联通产互为什么总加班？",34));
    expect(safe.pages.find(p=>p.id===page.id)!.nodes.find(n=>n.id===title.id)!.appearance.rect.w).toBe(740);
    const painted=await mergeHtml(d,edited(740,"联通产互为什么总加班？",34,60.8,50.2,"background:white"));
    expect(painted.diagnostics.some(d=>d.message.includes("FIXED_OCCLUSION"))).toBe(true);
    const occluded=await mergeHtml(d,edited(400,"覆盖品牌",34,60.8,760));
    expect(occluded.diagnostics.some(d=>d.message.includes('"protectedBounds"'))).toBe(true);
    const clipped=await mergeHtml(d,edited(349,"加班，常是交付的末端缓冲区",30,61,50));
    const failure=clipped.diagnostics.find(d=>d.message.includes("Text is clipped or outside its box"))!.message;
    expect(failure).toContain("Text is clipped or outside its box");
    const details=JSON.parse(failure.slice(failure.indexOf("{")));
    expect(details.nodeBounds.w).toBe(349);expect(details.requiredTextHeight).toBeGreaterThan(61);expect(details.overflow.bottom).toBeGreaterThan(0);
    expect(details.firstOverflow.character).toBeTruthy();
    expect(details.recovery).toContain("own edges");
  },30000);

  it("requires an exact author grant for semantic-only edits and persists them in PPTX",async()=>{
    const d=await syncDesign(source), brand=logo(d), baseline=projectPage(d,d.pages[0].id);
    const html=change(baseline,brand.id,e=>{e.dataset.purpose="reference";e.dataset.note="按主题重做示例图";});
    await expect(convertDesign({operation:"apply",design:d,expectedRevision:d.revision,html})).rejects.toThrow("TEMPLATE_PERMISSION");
    const proposal=templateSemanticProposal(d,baseline,html);
    await expect(convertDesign({operation:"apply",design:d,expectedRevision:d.revision,html,semanticAuthorization:{...proposal,revision:proposal.revision+1,mode:"template-authoring"}})).resolves.toBeDefined();
    const changed=await convertDesign({operation:"apply",design:d,expectedRevision:d.revision,html,semanticAuthorization:{...proposal,mode:"template-authoring"}});
    expect(changed.revision).toBeGreaterThan(d.revision);
    expect(changed.changes?.[0].fields).toEqual(["semantics.purpose","semantics.note"]);
    const reopened=await syncDesign(Buffer.from(changed.source.pptxBase64,"base64"));
    expect(logo(reopened).purpose).toBe("reference");expect(logo(reopened).note).toBe("按主题重做示例图");
    const zip=await JSZip.loadAsync(Buffer.from(reopened.source.pptxBase64,"base64"));
    const record=readTemplateRecord(new DOMParser().parseFromString(await zip.file(reopened.pages[0].part)!.async("string"),"application/xml"))!;
    expect(record.nodes[brand.origin!.shapeId].purpose).toBe("reference");
    expect(Object.keys(record)).toEqual(["version","nodes"]);
    await expect(mergeHtml(reopened,projectPage(reopened,reopened.pages[0].id).replace('data-locked="all"','data-locked="content"'))).rejects.toThrow("Read-only HTML attribute");
  },30000);

  it("reports all omitted metadata with original values and allows a repaired content-only submission",async()=>{
    const d=await syncDesign(source), page=d.pages[0], brand=logo(d);
    page.semantics={note:"固定品牌不可移动、变形或遮挡"}; brand.note="品牌标识，原样保留";
    const baseline=projectPage(d,page.id);
    const broken=change(baseline,brand.id,e=>{e.removeAttribute("data-name");e.removeAttribute("data-note");e.ownerDocument.body.dataset.pageNote="保留品牌";});
    let message="";
    try {templateSemanticProposal(d,baseline,broken);} catch(error) {message=(error as Error).message;}
    expect(message).toContain("Read-only HTML attribute");
    const details=JSON.parse(message.slice(message.indexOf("{")));
    expect(details.readonlyChanges).toEqual([{nodeId:brand.id,name:brand.name,attribute:"data-name",original:brand.name,submitted:null}]);
    expect(details.semanticChanges).toHaveLength(2);
    expect(details.semanticChanges.find((c:any)=>c.nodeId===brand.id).before.note).toBe(brand.note);
    expect(details.semanticChanges.find((c:any)=>!c.nodeId).before.note).toBe(page.semantics.note);
    expect(details.recovery).toContain("retry apply_html");
    expect(d.revision).toBe(0); expect(brand.purpose).toBe("fixed");
    const noteOnly=change(baseline,brand.id,e=>e.removeAttribute("data-note"));
    await expect(convertDesign({operation:"apply",design:d,expectedRevision:d.revision,html:noteOnly})).rejects.toThrow('"semanticChanges"');
    const repaired=await convertDesign({operation:"apply",design:d,expectedRevision:d.revision,html:baseline.replace("{{TITLE}}","共同成长")});
    expect(logo(repaired).appearance).toEqual(brand.appearance);
    expect(repaired.pages[0].nodes.some(n=>nodeText(n)==="共同成长")).toBe(true);
  },30000);

  it("does not gate edits or template proposals on version or hash metadata",async()=>{
    const d=await syncDesign(source), baseline=projectPage(d,d.pages[0].id);
    for (const html of [baseline.replace(d.source.sha256,"short"),baseline.replace(d.source.sha256,"0".repeat(64)),baseline.replace('data-design-revision="0"','data-design-revision="99"'),baseline.replace(/ data-source-hash="[^"]*"/,"").replace(/ data-design-revision="[^"]*"/,"")]) {
      expect(()=>templateSemanticProposal(d,baseline,html)).not.toThrow();
      await expect(convertDesign({operation:"apply",design:d,expectedRevision:999,html})).resolves.toBeDefined();
    }
    expect(()=>templateSemanticProposal(d,baseline,baseline.replaceAll(d.pages[0].id,"ppt/slides/missing.xml"))).toThrow("design page");
  });

  it("preserves logo policies, assets and distinct identities while duplicating to twenty pages and reopening",async()=>{
    let d=await syncDesign(source);
    const count=d.pages.length;
    const beforeHash=d.source.sha256;
    d=await convertDesign({operation:"pages",design:d,expectedRevision:d.revision,commands:Array.from({length:20-count},()=>({op:"duplicateSlide" as const,slide:1}))});
    expect(d.pages).toHaveLength(20);
    expect(new Set(d.pages.map(p=>p.id)).size).toBe(20);
    expect(new Set(d.pages.flatMap(p=>p.nodes.map(n=>n.id))).size).toBe(d.pages.reduce((n,p)=>n+p.nodes.length,0));
    const reopened=await syncDesign(Buffer.from(d.source.pptxBase64,"base64"));
    for(const page of reopened.pages){expect(page.nodes.some(n=>n.kind==="image"&&n.purpose==="fixed")).toBe(true);}
    expect(d.source.sha256).not.toBe(beforeHash);
    expect((await syncDesign(source)).source.sha256).toBe(beforeHash);
    const bytes=Buffer.from(d.source.pptxBase64,"base64");
    const bridge=new DesignBridge({snapshotPptx:async()=>new Blob([bytes]),loadPptx:async()=>{},getCurrentSlideIndex:()=>0,getSlidesCount:()=>20,selectSlide:async()=>{}},convertDesign);
    try {
      const overview=payload(await bridge.run("get_design",{}));
      expect(overview.pages).toHaveLength(20);expect(overview.totalPages).toBe(20);
      const single=payload(await bridge.run("get_design",{page:20}));
      expect(single.pages).toBeUndefined();expect(single.page).toBe(20);expect(single.totalPages).toBe(20);
      expect(single.pageId).toBe(overview.pages[19].id);expect(single.semantics).toEqual(overview.pages[19].semantics);
      expect(single.design_revision).toBe(overview.design_revision);expect(single.sourceHash).toBe(overview.sourceHash);
      expect(single.nodes.length).toBeGreaterThan(0);expect(single.assets.length).toBeGreaterThan(0);
      expect(single.html).toContain(`data-page-id="${single.pageId}"`);
    } finally {bridge.dispose();}
  },60000);

  it("exposes manual annotation via the same HTML bridge and keeps author confirmation without version gating",async()=>{
    let bytes:Uint8Array=source;let confirmResolve!:(ok:boolean)=>void; let started!:()=>void; const confirming=new Promise<void>(resolve=>{started=resolve;});
    const viewer={snapshotPptx:async()=>new Blob([bytes as BlobPart]),loadPptx:async(buffer:ArrayBuffer)=>{bytes=new Uint8Array(buffer)},getCurrentSlideIndex:()=>0,getSlidesCount:()=>4,selectSlide:async()=>{}};
    const bridge=new DesignBridge(viewer,convertDesign,undefined,{mode:"template-authoring",confirmSemanticChanges:()=>new Promise(resolve=>{confirmResolve=resolve;started();})});
    expect(payload(await bridge.run("get_template",{})).pages).toHaveLength(4);
    expect(payload(await bridge.run("suggest_annotations",{slide:1})).pages[0].objects.length).toBeGreaterThan(0);
    let info=payload(await bridge.run("get_design",{page:1}));
    const brand=info.nodes.find((n:any)=>n.shapeId==="3");
    payload(await bridge.updateTemplateSemantics(1,[{nodeId:brand.id,note:"品牌标识，原样保留"},{role:"cover",layoutName:"品牌封面",note:"红白品牌风格"}]));
    info=payload(await bridge.run("get_design",{page:1}));
    expect(info.semantics.layoutName).toBe("品牌封面");expect(info.nodes.find((n:any)=>n.id===brand.id).note).toBe("品牌标识，原样保留");
    payload(await bridge.updateTemplateSemantics(1,[{role:""}]));
    info=payload(await bridge.run("get_design",{page:1}));expect(info.semanticsDefined).toBe(true);expect(info.semantics.role).toBeUndefined();
    expect((await syncDesign(bytes)).pages[0].semantics?.role).toBeUndefined();
    const pending=bridge.run("apply_html",{expected_design_revision:info.design_revision,html:change(info.html,brand.id,e=>e.dataset.purpose="content")});
    await confirming;bridge.edited();confirmResolve(true);
    expect((await pending).ok).toBe(true);
    expect(logo(await syncDesign(bytes)).purpose).toBe("content");bridge.dispose();
    const generation=new DesignBridge(viewer,convertDesign);
    expect((await generation.updateTemplateSemantics(1,[{nodeId:brand.id,purpose:"content"}])).ok).toBe(false);generation.dispose();
  },30000);

  it("protects fixed group children and permits replacement of a reference group as a whole",async()=>{
    const d=await syncDesign(readFileSync("tools/templates/blank.pptx")),p=d.pages[0];
    p.nodes.push({id:"new:group",name:"@参考素材",kind:"group",purpose:"reference",appearance:{rect:{x:100,y:100,w:200,h:200}}},{id:"new:child",name:"visual",kind:"shape",purpose:"reference",parentId:"new:group",appearance:{rect:{x:120,y:120,w:100,h:100},fill:"#000000"}});
    const measured={pageId:p.id,revision:d.revision,sourceHash:d.source.sha256,nodes:[]};
    expect(mergeMeasuredPage(d,measured).pages[0].nodes).toHaveLength(0);
    const group=p.nodes.find(n=>n.id==="new:group")!, child=p.nodes.find(n=>n.id==="new:child")!;
    group.purpose="fixed";
    expect(effectivePurpose(p,child)).toBe("fixed");
    expect(()=>mergeMeasuredPage(d,{...measured,nodes:[group]})).toThrow(/locked|Fixed/);
    expect(()=>mergeMeasuredPage(d,{...measured,nodes:p.nodes.map(n=>({...n,appearance:{...n.appearance,rect:{...n.appearance.rect,x:n.appearance.rect.x+10}}}))})).toThrow(/Locked/);
  });

  it("reconciles native delete, duplicate and undo without dropping surviving semantic policies",async()=>{
    let d=await syncDesign(source), baseline=projectPage(d,d.pages[0].id);
    const title=d.pages[0].nodes.find(n=>nodeText(n)==="{{TITLE}}")!;
    const html=change(baseline,title.id,e=>e.dataset.note="标题示例，可改写");
    d=await convertDesign({operation:"apply",design:d,expectedRevision:d.revision,html,semanticAuthorization:{...templateSemanticProposal(d,baseline,html),mode:"template-authoring"}});
    const session=await HeadlessSession.open(Buffer.from(d.source.pptxBase64,"base64"),{quiet:true});
    try{
      const deleted=await session.apply({commands:[{op:"delete",target:{slide:1,id:title.origin!.shapeId}}]});expect(deleted.ok,JSON.stringify(deleted)).toBe(true);
      const without=await syncDesign(await session.export());expect(without.pages[0].nodes.some(n=>nodeText(n)==="{{TITLE}}")).toBe(false);
      expect(logo(without).purpose).toBe("fixed");
      session.undo();const undone=await syncDesign(await session.export());expect(undone.pages[0].nodes.find(n=>nodeText(n)==="{{TITLE}}")?.note).toBe("标题示例，可改写");
      const duplicate=await session.apply({commands:[{op:"duplicate",target:{slide:1,id:title.origin!.shapeId}}]});expect(duplicate.ok,JSON.stringify(duplicate)).toBe(true);
      const copied=await syncDesign(await session.export());const twins=copied.pages[0].nodes.filter(n=>nodeText(n)==="{{TITLE}}");expect(twins).toHaveLength(2);expect(twins.every(n=>n.purpose==="content"&&n.note==="标题示例，可改写")).toBe(true);
      session.undo();expect((await syncDesign(await session.export())).pages[0].nodes.filter(n=>nodeText(n)==="{{TITLE}}")).toHaveLength(1);
      session.redo();expect((await syncDesign(await session.export())).pages[0].nodes.filter(n=>nodeText(n)==="{{TITLE}}")).toHaveLength(2);
    }finally{session.dispose();}
  },30000);

  it("rejects malformed policy instead of treating corrupted data as permission",()=>{
    expect(()=>validateTemplateRecord({version:2,nodes:{}})).toThrow();
    expect(()=>validateTemplateRecord({version:1,nodes:{"2":{purpose:"anything"}}})).toThrow();
    expect(()=>validateTemplateRecord({version:1,nodes:{"2":{purpose:"content",unlocked:true}}})).toThrow();
  });
});
