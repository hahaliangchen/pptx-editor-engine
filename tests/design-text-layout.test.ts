// @vitest-environment node
import {readFileSync} from 'node:fs';
import {describe,expect,it} from 'vitest';
import {syncDesign,convertDesign} from '../src/design/service';
import {projectPage} from '../src/design/projection';
import {nodeText} from '../src/design/model';
import {DesignBridge} from '../src/page/design-bridge';
import {PAGE_TOOLS} from '../src/page/page-tools';

const blank=readFileSync('tools/templates/blank.pptx');
const paragraph=(text:string)=>`<p style="font:24px Arial;line-height:32px;margin:0">${text}</p>`;
describe('native text layout mapping',()=>{
  it.each(['display:flex;gap:24px','display:grid;grid-template-columns:180px 1fr;gap:24px'])('keeps independently positioned paragraphs editable through PPTX reopen: %s',async layout=>{
    const before=await syncDesign(blank);
    const html=projectPage(before,before.pages[0].id).replace('</body>',`<div data-id="new:outer" data-kind="group" style="position:absolute;left:70px;top:200px;width:1100px;height:160px">
      <div data-id="new:row" data-kind="text" style="width:1100px;padding:14px 12px;background:#eeeeee;border:1px solid #e60012;${layout}">
      <p style="width:180px;font:24px Arial;line-height:32px;margin:0"><strong>Left title</strong></p>${paragraph('Right explanation stays beside the title')}</div></div></body>`);
    const converted=await convertDesign({operation:'apply',design:before,expectedRevision:before.revision,html});
    const row=converted.pages[0].nodes.find(n=>n.id==='new:row')!;
    expect(row.kind).toBe('group');expect(row.parentId).toBe('new:outer');
    expect(row.appearance.text).toBeUndefined();
    const reopened=await syncDesign(Buffer.from(converted.source.pptxBase64,'base64'));
    const left=reopened.pages[0].nodes.find(n=>nodeText(n)==='Left title')!;
    const right=reopened.pages[0].nodes.find(n=>nodeText(n).startsWith('Right explanation'))!;
    expect(left.parentId).toBe(right.parentId);expect(left.parentId).toBeDefined();
    expect(left.appearance.rect.y).toBeCloseTo(right.appearance.rect.y,1);
    expect(right.appearance.rect.x).toBeGreaterThan(left.appearance.rect.x+left.appearance.rect.w+23);
    expect(left.appearance.text!.paragraphs).toHaveLength(1);
    expect(left.appearance.text!.paragraphs[0].runs[0].bold).toBe(true);
    expect(reopened.pages[0].nodes.some(n=>n.appearance.fill==='#eeeeee' && n.appearance.border?.color==='#e60012')).toBe(true);
    // Reopening and making a subsequent edit must use the materialized native
    // groups, rather than resurrect the original unrepresentable text flow.
    const next=await convertDesign({operation:'apply',design:converted,expectedRevision:converted.revision,
      html:projectPage(converted,converted.pages[0].id).replace('Left title','Revised title')});
    expect(next.pages[0].nodes.some(n=>nodeText(n)==='Revised title')).toBe(true);
    expect(next.pages[0].nodes.some(n=>nodeText(n).startsWith('Right explanation'))).toBe(true);
  },30000);

  it('keeps ordinary consecutive paragraphs in one native text frame',async()=>{
    const before=await syncDesign(blank);
    const html=projectPage(before,before.pages[0].id).replace('</body>',`<div data-id="new:flow" data-kind="text" style="position:absolute;left:80px;top:80px;width:700px;height:120px">${paragraph('First paragraph')}${paragraph('Second paragraph')}</div></body>`);
    const converted=await convertDesign({operation:'apply',design:before,expectedRevision:before.revision,html});
    const flow=converted.pages[0].nodes.find(n=>n.id==='new:flow')!;
    expect(flow.kind).toBe('text');expect(flow.appearance.text!.paragraphs).toHaveLength(2);
  },30000);

  it('replaces an existing native text flow with a group while keeping its source identity',async()=>{
    const before=await syncDesign(blank);
    const initial=await convertDesign({operation:'apply',design:before,expectedRevision:before.revision,
      html:projectPage(before,before.pages[0].id).replace('</body>',`<div data-kind="text" data-id="new:existing" style="position:absolute;left:80px;top:180px;width:1000px;height:120px">${paragraph('Existing left')}${paragraph('Existing right')}</div></body>`)});
    const original=initial.pages[0].nodes.find(n=>n.id==='new:existing')!;
    const dom=new DOMParser().parseFromString(projectPage(initial,initial.pages[0].id),'text/html');
    const owner=Array.from(dom.querySelectorAll<HTMLElement>('[data-id]')).find(n=>n.dataset.id===original.id)!;
    owner.style.cssText+=';display:flex;flex-direction:row;gap:32px';
    (owner.querySelector('p') as HTMLElement).style.width='240px';
    const updated=await convertDesign({operation:'apply',design:initial,expectedRevision:initial.revision,html:'<!doctype html>'+dom.documentElement.outerHTML});
    const group=updated.pages[0].nodes.find(n=>n.id===original.id)!;
    expect(group.kind).toBe('group');expect(group.origin!.shapeId).toBe(original.origin!.shapeId);
    const children=updated.pages[0].nodes.filter(n=>n.parentId===group.id && n.kind==='text');
    expect(children).toHaveLength(2);expect(children[0].appearance.rect.y).toBeCloseTo(children[1].appearance.rect.y,1);
    expect(children[1].appearance.rect.x).toBeGreaterThan(children[0].appearance.rect.x+240);
  },30000);

  it('keeps version counters informational and accepts missing or mismatched expected versions',async()=>{
    let bytes:Uint8Array=blank;
    const bridge=new DesignBridge({snapshotPptx:async()=>new Blob([bytes as BlobPart]),loadPptx:async buffer=>{bytes=new Uint8Array(buffer)},getCurrentSlideIndex:()=>0,getSlidesCount:()=>1,selectSlide:async()=>{}},convertDesign);
    const read=await bridge.run('get_design',{page:1});
    const info=JSON.parse((read.content[0] as {text:string}).text);
    expect(info.design_revision).toBe(0);expect(info.revision).toBeUndefined();
    const wrong=await bridge.run('apply_html',{expected_design_revision:8,html:info.html});
    expect(wrong.ok).toBe(true);
    const good=await bridge.run('apply_html',{expected_design_revision:info.design_revision,html:info.html});
    expect(good.ok).toBe(true);
    const schema=PAGE_TOOLS.find(t=>t.name==='apply_html')!.inputSchema as any;
    expect(schema.required).not.toContain('expected_design_revision');expect(schema.properties.expected_revision).toBeUndefined();
    bridge.dispose();
  },30000);
});
