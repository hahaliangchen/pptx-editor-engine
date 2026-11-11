// @vitest-environment node
import {readFileSync} from 'node:fs';
import {beforeAll,expect,it} from 'vitest';
import {convertDesign,materializeDesign,syncDesign} from '../src/design/service';
import {cloneDesign,nodeText,type DesignDocument} from '../src/design/model';
import {DesignBridge,type DesignConverter} from '../src/page/design-bridge';

let fixture:DesignDocument;
beforeAll(async()=>{
  const original=await syncDesign(readFileSync('tests/fixtures/unicom-uniai-template.pptx'));
  for(const page of original.pages) {page.semantics={...page.semantics,fill:'layout'};page.semanticsChanged=true;}
  fixture=await materializeDesign(original);
},30000);
function setup(converter:DesignConverter=convertDesign,failSave=false,template=fixture) {
  let bytes=Uint8Array.from(Buffer.from(template.source.pptxBase64,'base64')),stored=bytes.slice(),saves=0;
  const bridge=new DesignBridge({snapshotPptx:async()=>new Blob([bytes]),loadPptx:async data=>{bytes=new Uint8Array(data);},getCurrentSlideIndex:()=>0,getSlidesCount:()=>fixture.pages.length,selectSlide:async()=>{}},converter,undefined,{
    persist:async()=>{if(failSave)throw new Error('storage unavailable');saves++;stored=bytes.slice();bridge.saved();return 10+saves;}
  });
  const fills=[1,2,3].map(page=>{
    const node=fixture.pages[page-1].nodes.find(n=>n.kind==='text'&&!n.lock)!;
    if(!node)throw Error(`fixture text missing ${page}`);
    return {page,texts:{[node.id]:`Filled page ${page}`}};
  });
  return {bridge,fills,stored:()=>stored,saves:()=>saves};
}
it('prepares three source patches concurrently and preserves fixed objects through save and reopen',async()=>{
  let active=0,peak=0;
  const context=setup(async request=>{
    if(request.operation!=='prepare')return convertDesign(request);
    active++;peak=Math.max(peak,active);
    try{return await convertDesign(request);}finally{active--;}
  });
  const jobs=[];
  for(const fill of context.fills){
    const read=await context.bridge.run('get_design',{page:fill.page,html:true});
    const html=JSON.parse((read.content[0] as {text:string}).text).html;
    const doc=new DOMParser().parseFromString(html,'text/html');
    const element=Array.from(doc.querySelectorAll<HTMLElement>('[data-id]')).find(el=>el.dataset.id===Object.keys(fill.texts)[0])!;
    const old=element.outerHTML;expect(html).toContain(old);
    element.querySelector('p')!.textContent=`Patched page ${fill.page}`;
    jobs.push({page:fill.page,edits:[{old_string:old,new_string:element.outerHTML}]});
  }
  const replies=await Promise.all(jobs.map(job=>context.bridge.run('edit_html',job)));
  expect(replies.every(r=>r.ok),JSON.stringify(replies)).toBe(true);
  expect(peak).toBe(3);expect(context.saves()).toBe(1);
  const reopened=await syncDesign(context.stored());
  for(const page of [1,2,3])expect(reopened.pages[page-1].nodes.some(n=>nodeText(n).replace(/\u2028/g,'').includes(`Patched page ${page}`))).toBe(true);
  expect(reopened.pages.flatMap(p=>p.nodes.filter(n=>n.lock).map(n=>[n.id,n.appearance])))
    .toEqual(fixture.pages.flatMap(p=>p.nodes.filter(n=>n.lock).map(n=>[n.id,n.appearance])));
  context.bridge.dispose();
},30000);
it('rejects a bad source patch atomically and keeps page identity and fixed-object protection',async()=>{
  const context=setup();
  const read=await context.bridge.run('get_design',{page:1,html:true});
  const html=JSON.parse((read.content[0] as {text:string}).text).html;
  const originalBytes=context.stored().slice();
  const doc=new DOMParser().parseFromString(html,'text/html');
  const fixed=doc.querySelector<HTMLElement>('[data-locked="all"]')!;
  const fixedSource=fixed.outerHTML;fixed.style.left=`${(parseFloat(fixed.style.left)||0)+2}px`;
  for(const edits of [
    [{old_string:'</body>',new_string:'<div></div></body>'},{old_string:'MISSING_UNIQUE_ANCHOR',new_string:'x'}],
    [{old_string:`data-page-id="${fixture.pages[0].id}"`,new_string:`data-page-id="${fixture.pages[1].id}"`}],
    [{old_string:fixedSource,new_string:fixed.outerHTML}],
  ]){
    const result=await context.bridge.run('edit_html',{page:1,edits});
    expect(result.ok,JSON.stringify(result)).toBe(false);expect(context.saves()).toBe(0);
  }
  expect(context.stored()).toEqual(originalBytes);
  const after=await context.bridge.run('get_design',{page:1,html:true});
  expect(JSON.parse((after.content[0] as {text:string}).text).html).toBe(html);
  context.bridge.dispose();
},30000);
it('prepares three real pages concurrently, merges only their pages, and acknowledges their shared saved PPTX',async()=>{
  let active=0,peak=0;
  const context=setup(async request=>{
    if(request.operation!=='prepare')return convertDesign(request);
    active++;peak=Math.max(peak,active);
    try{return await convertDesign(request);}finally{active--;}
  });
  const replies=await Promise.all(context.fills.map(fill=>context.bridge.run('fill_layout',fill)));
  const data=replies.map(result=>{expect(result.ok,JSON.stringify(result)).toBe(true);return JSON.parse((result.content[0] as {text:string}).text);});
  expect(peak).toBe(3);expect(context.saves()).toBe(1);
  expect(data.every(d=>d.saved&&d.applied&&d.stored_revision===11)).toBe(true);
  expect(data[0].pageConcurrency.savedPages).toEqual([1,2,3]);
  const reopened=await syncDesign(context.stored());
  for(const page of [1,2,3])expect(reopened.pages[page-1].nodes.some(n=>nodeText(n).replace(/\u2028/g,'')===`Filled page ${page}`),JSON.stringify({page,texts:reopened.pages[page-1].nodes.map(nodeText)})).toBe(true);
  const fixedBefore=fixture.pages.flatMap(p=>p.nodes.filter(n=>n.lock).map(n=>[n.id,n.appearance]));
  const fixedAfter=reopened.pages.flatMap(p=>p.nodes.filter(n=>n.lock).map(n=>[n.id,n.appearance]));
  expect(fixedAfter).toEqual(fixedBefore);context.bridge.dispose();
},30000);
it('preserves the original page failure and saves successful peers without replacing the failed page',async()=>{
  const context=setup(async request=>{
    if(request.operation==='prepare'&&request.html?.includes('Filled page 2'))throw Error('Locked element changed: protected-brand');
    return convertDesign(request);
  });
  const replies=await Promise.all(context.fills.map(fill=>context.bridge.run('fill_layout',fill)));
  expect(replies.map(r=>r.ok)).toEqual([true,false,true]);
  expect((replies[1].content[0] as {text:string}).text).toContain('Locked element changed: protected-brand');
  const reopened=await syncDesign(context.stored());
  expect(reopened.pages[0].nodes.some(n=>nodeText(n)==='Filled page 1')).toBe(true);
  expect(reopened.pages[1].nodes.some(n=>nodeText(n)==='Filled page 2')).toBe(false);
  expect(reopened.pages[2].nodes.some(n=>nodeText(n)==='Filled page 3')).toBe(true);
  expect(context.saves()).toBe(1);context.bridge.dispose();
},30000);
it('returns failure to every page when the shared save fails and never emits a saved receipt',async()=>{
  const context=setup(convertDesign,true);
  const replies=await Promise.all(context.fills.map(fill=>context.bridge.run('fill_layout',fill)));
  expect(replies.every(r=>!r.ok&&JSON.stringify(r).includes('storage unavailable'))).toBe(true);
  expect(replies.some(r=>JSON.stringify(r).includes('"saved":true'))).toBe(false);
  expect(context.saves()).toBe(0);context.bridge.dispose();
},30000);
it('rejects a prepared page after manual editing invalidates its baseline',async()=>{
  let finish:()=>void;const blocker=new Promise<void>(resolve=>{finish=resolve;});
  let entered:()=>void;const started=new Promise<void>(resolve=>{entered=resolve;});
  const context=setup(async request=>{if(request.operation==='prepare'){entered!();await blocker;}return convertDesign(request);});
  const job=context.bridge.run('fill_layout',context.fills[0]);
  await started;context.bridge.edited();finish!();
  const reply=await job;expect(reply.ok).toBe(false);expect(JSON.stringify(reply)).toContain('EDITOR_CHANGED');
  expect(context.saves()).toBe(0);context.bridge.dispose();
},30000);

it('isolates three full HTML page edits and reads their actual saved text back',async()=>{
  let active=0,peak=0;
  const context=setup(async request=>{
    if(request.operation!=='prepare')return convertDesign(request);
    active++;peak=Math.max(peak,active);
    try{return await convertDesign(request);}finally{active--;}
  });
  const html=[];
  for(const fill of context.fills){
    const response=await context.bridge.run('get_design',{page:fill.page,html:true});
    const page=JSON.parse((response.content[0] as {text:string}).text);
    const document=new DOMParser().parseFromString(page.html,'text/html');
    const [id,text]=Object.entries(fill.texts)[0];
    document.querySelector(`[data-id="${id}"] p`)!.textContent=text;
    html.push('<!doctype html>'+document.documentElement.outerHTML);
  }
  const replies=await Promise.all(html.map(html=>context.bridge.run('apply_html',{html})));
  expect(replies.every(result=>result.ok),JSON.stringify(replies)).toBe(true);
  expect(peak).toBe(3);expect(context.saves()).toBe(1);
  for(const fill of context.fills){
    const response=await context.bridge.run('get_design',{page:fill.page,html:true});
    const page=JSON.parse((response.content[0] as {text:string}).text);
    expect(page.nodes.find((node:{id:string})=>node.id===Object.keys(fill.texts)[0]).text.replace(/\u2028/g,''))
      .toBe(`Filled page ${fill.page}`);
  }
  const reopened=await syncDesign(context.stored());
  expect(reopened.pages.flatMap(p=>p.nodes.filter(n=>n.lock).map(n=>[n.id,n.appearance])))
    .toEqual(fixture.pages.flatMap(p=>p.nodes.filter(n=>n.lock).map(n=>[n.id,n.appearance])));
  context.bridge.dispose();
},30000);

it('preserves added block and field roles through actual save/reopen, including further growth',async()=>{
  const design=cloneDesign(fixture),page=design.pages[2];
  const field=page.nodes.find(n=>n.kind==='text'&&!n.lock)!;
  field.slot='body';field.parentId='test-block';
  page.nodes.push({id:'test-block',name:'block',kind:'group',purpose:'content',slot:'内容#1',appearance:{rect:{...field.appearance.rect}}});
  page.semanticsChanged=true;
  const context=setup(convertDesign,false,await materializeDesign(design));
  const read=async()=>JSON.parse((await context.bridge.run('get_layout',{page:3})).content.filter(c=>c.type==='text').map(c=>(c as {text:string}).text).join(''));
  for(const count of [3,5]){
    const reply=await context.bridge.run('fill_layout',{page:3,sets:[{set:'内容#1',items:Array.from({length:count},(_,i)=>({fields:{body:`第${i+1}条真实内容`}}))}]});
    expect(reply.ok,JSON.stringify(reply)).toBe(true);
    const actual=await read();expect(actual.sets[0].items).toHaveLength(count);
    expect(actual.sets[0].items.map((i:{fields:{body:string[]}})=>i.fields.body[0])).toEqual(Array.from({length:count},(_,i)=>`第${i+1}条真实内容`));
    const reopened=await syncDesign(context.stored());
    expect(reopened.pages[2].nodes.filter(n=>n.slot==='内容#1')).toHaveLength(count);
    expect(reopened.pages[2].nodes.filter(n=>n.slot==='body')).toHaveLength(count);
  }
  context.bridge.dispose();
},30000);

it('reads a committed page while peers prepare, without blocking their writes or allowing structure changes',async()=>{
  let release!:()=>void,started!:()=>void;
  let hold=false;
  const held=new Promise<void>(resolve=>{release=resolve;}),preparing=new Promise<void>(resolve=>{started=resolve;});
  const context=setup(async request=>{
    if(request.operation==='prepare'&&hold){started();await held;}
    return convertDesign(request);
  });
  const first=await context.bridge.run('fill_layout',context.fills[0]);expect(first.ok,JSON.stringify(first)).toBe(true);
  hold=true;
  const peers=context.fills.slice(1).map(fill=>context.bridge.run('fill_layout',fill));
  await preparing;
  try {
    for(const name of ['get_design','get_layout','get_template']) {
      const read=await context.bridge.run(name,{page:1,html:true,slide:1});
      expect(read.ok,JSON.stringify(read)).toBe(true);
      expect(JSON.stringify(read)).toContain('Filled page 1');
    }
    const structural=await context.bridge.run('manage_design_pages',{commands:[{op:'deleteSlide',slide:1}]});
    expect(structural.ok).toBe(false);expect(JSON.stringify(structural)).toContain('LAYOUT_FILL_BUSY');
  } finally {release();}
  expect((await Promise.all(peers)).every(r=>r.ok)).toBe(true);
  const reopened=await syncDesign(context.stored());
  for(const fill of context.fills)expect(reopened.pages[fill.page-1].nodes.some(n=>nodeText(n).replace(/\u2028/g,'')===`Filled page ${fill.page}`)).toBe(true);
  expect(context.saves()).toBe(2);
  context.bridge.dispose();
},30000);

it('shares a cold sync between a read and page writes without treating the read as a conversion lock',async()=>{
  let release!:()=>void,started!:()=>void,heldSync=false;
  const held=new Promise<void>(resolve=>{release=resolve;}),syncing=new Promise<void>(resolve=>{started=resolve;});
  const context=setup(async request=>{
    if(request.operation==='sync'&&!heldSync){heldSync=true;started();await held;}
    return convertDesign(request);
  });
  const read=context.bridge.run('get_design',{page:1,html:true});await syncing;
  const writes=context.fills.map(fill=>context.bridge.run('fill_layout',fill));
  release();
  const replies=await Promise.all([read,...writes]);
  expect(replies.every(r=>r.ok),JSON.stringify(replies)).toBe(true);
  const old=JSON.parse((replies[0].content[0] as {text:string}).text);
  expect(old.html).not.toContain('Filled page 1');
  const latest=await context.bridge.run('get_design',{page:1,html:true});
  expect(latest.ok).toBe(true);expect(JSON.stringify(latest)).toContain('Filled page 1');
  context.bridge.dispose();
},30000);

it('reports an unchanged saved page honestly and retains its actual content through cold reload',async()=>{
  const context=setup();
  const first=await context.bridge.run('fill_layout',context.fills[0]);expect(first.ok).toBe(true);
  const current=await context.bridge.run('get_design',{page:1,html:true});expect(current.ok).toBe(true);
  const html=JSON.parse((current.content[0] as {text:string}).text).html;
  const second=await context.bridge.run('apply_html',{html});expect(second.ok,JSON.stringify(second)).toBe(true);
  const receipt=JSON.parse((second.content[0] as {text:string}).text);
  expect(receipt.applied).toBe(false);expect(receipt.saved).toBe(true);expect(receipt.stored_revision).toBe(12);
  const reopened=await syncDesign(context.stored());
  expect(reopened.pages[0].nodes.some(n=>nodeText(n).replace(/\u2028/g,'')==='Filled page 1')).toBe(true);
  context.bridge.dispose();
},30000);
