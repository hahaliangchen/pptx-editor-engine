// @vitest-environment node
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { createCanvas, loadImage } from '@napi-rs/canvas';
import { syncDesign, convertDesign } from '../src/design/service';
import { projectPage } from '../src/design/projection';
import { effectiveLock, effectivePurpose } from '../src/design/template-policy';
import { HeadlessSession } from '../src/headless/session';

it('materializes SVG icons as editable shapes and keeps them across subsequent edits', async () => {
  const native = JSON.parse(readFileSync('tests/fixtures/native-svg-clock.json', 'utf8'));
  const design = await syncDesign(readFileSync('tools/templates/blank.pptx'));
  const preview='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/aN8AAAAASUVORK5CYII=';
  const html = projectPage(design, design.pages[0].id).replace('</body>', `<img data-id="new:clock" data-kind="image" src="${preview}" style="position:absolute;left:100px;top:120px;width:80px;height:80px"/></body>`);
  const saved = await convertDesign({operation:'apply', design, html, nativeAssets:{[preview]:native}});
  const zip = await JSZip.loadAsync(Buffer.from(saved.source.pptxBase64,'base64'));
  const slide = await zip.file(saved.pages[0].part)!.async('string');
  const doc = new DOMParser().parseFromString(slide,'application/xml');
  expect(doc.getElementsByTagNameNS('*','grpSp').length).toBeGreaterThan(0);
  expect(doc.getElementsByTagNameNS('*','custGeom').length).toBeGreaterThan(0);
  expect(doc.getElementsByTagNameNS('*','pic').length).toBe(0);
  const session = await HeadlessSession.open(Buffer.from(saved.source.pptxBase64,'base64'), {quiet:true});
  const rendered = await loadImage(await session.render(1, {width:1280}));
  const canvas = createCanvas(rendered.width,rendered.height), context=canvas.getContext('2d');
  context.drawImage(rendered,0,0);
  const pixels=context.getImageData(100,120,80,80).data;
  let ink=0; for(let i=0;i<pixels.length;i+=4) if(pixels[i] < 120 && pixels[i+1] < 190 && pixels[i+2] < 210) ink++;
  expect(ink).toBeGreaterThan(100);
  const update = await session.apply([{op:'setLine', target:{slide:1,name:'Ellipse 2'}, color:'#ff0000', widthPt:2}]);
  expect(update.ok, JSON.stringify(update)).toBe(true);
  const modified = await session.export();
  session.dispose();
  const modifiedZip = await JSZip.loadAsync(modified);
  expect(await modifiedZip.file(saved.pages[0].part)!.async('string')).toContain('FF0000');
  const reopened = await syncDesign(Buffer.from(saved.source.pptxBase64,'base64'));
  const icon = reopened.pages[0].nodes.find(n => n.kind === 'group' && n.name === 'new:clock');
  expect(icon).toBeDefined();
  const parts = reopened.pages[0].nodes.filter(n => n.parentId === icon!.id);
  expect(parts.length).toBe(2);
  // What the AI drew stays its content after a save and reload: only the template's skeleton is locked.
  expect(parts.map(n => [effectivePurpose(reopened.pages[0], n), effectiveLock(reopened.pages[0], n)])).toEqual([["content", undefined], ["content", undefined]]);
  const html2 = projectPage(reopened,reopened.pages[0].id).replace('</body>', '<div data-kind="shape" data-id="new:box" style="position:absolute;left:500px;top:300px;width:40px;height:40px;background:#ff0000"></div></body>');
  const edited = await convertDesign({operation:'apply',design:reopened,html:html2});
  const zip2 = await JSZip.loadAsync(Buffer.from(edited.source.pptxBase64,'base64'));
  expect(await zip2.file(edited.pages[0].part)!.async('string')).toContain('custGeom');
}, 30000);

it('accepts a model-authored inline SVG through the page bridge without an asset tool', async () => {
  const { DesignBridge } = await import('../src/page/design-bridge');
  const native = JSON.parse(readFileSync('tests/fixtures/native-svg-clock.json', 'utf8'));
  let current = readFileSync('tools/templates/blank.pptx');
  let conversions = 0;
  const viewer = {snapshotPptx:async()=>new Blob([current]), loadPptx:async(buffer:ArrayBuffer)=>{current=Buffer.from(buffer);},getCurrentSlideIndex:()=>0,getSlidesCount:()=>1,selectSlide:async()=>{}};
  const bridge = new DesignBridge(viewer, convertDesign, undefined, {convertInlineSvg:async(svg)=>{conversions++; expect(svg).toContain('<circle'); expect(svg).not.toContain('data-kind'); return native;}});
  const before = await bridge.run('get_design',{page:1});
  expect(before.ok).toBe(true);
  const data=JSON.parse(before.content[0].text!);
  const html=data.html.replace('</body>', '<svg viewBox="0 0 64 64" width="80" height="80" style="position:absolute;left:100px;top:120px"><defs><linearGradient id="tone"><stop offset="0%" stop-color="#147d92"/><stop offset="100%" stop-color="#1a3c92"/></linearGradient></defs><circle cx="32" cy="32" r="25" fill="url(#tone)"/><path d="M32 16v17l12 7" fill="none" stroke="#147d92" stroke-width="4"/></svg></body>');
  const unsafe=await bridge.run('apply_html',{html:data.html.replace('</body>', '<svg viewBox="0 0 24 24"><path onload="alert(1)" d="M0 0 L1 1"/></svg></body>')});
  expect(unsafe.ok).toBe(false); expect(conversions).toBe(0);
  const result=await bridge.run('apply_html',{html});
  expect(result.ok,JSON.stringify(result)).toBe(true); expect(conversions).toBe(1);
  const zip=await JSZip.loadAsync(current); const xml=await zip.file('ppt/slides/slide1.xml')!.async('string');
  expect(xml).toContain('custGeom'); expect(xml).not.toContain('<p:pic>');
  bridge.dispose();
},30000);

it('accepts inline SVG on the persisted apply_html path used by the product page', async () => {
  const { DesignBridge } = await import('../src/page/design-bridge');
  const native = JSON.parse(readFileSync('tests/fixtures/native-svg-clock.json', 'utf8'));
  let current = readFileSync('tools/templates/blank.pptx'), saved = 0;
  const viewer = {snapshotPptx:async()=>new Blob([current]), loadPptx:async(buffer:ArrayBuffer)=>{current=Buffer.from(buffer);},getCurrentSlideIndex:()=>0,getSlidesCount:()=>1,selectSlide:async()=>{}};
  const bridge = new DesignBridge(viewer, convertDesign, undefined, {convertInlineSvg:async()=>native, persist:async()=>++saved});
  const data = JSON.parse((await bridge.run('get_design',{page:1})).content[0].text!);
  const html = data.html.replace('</body>', '<svg viewBox="0 0 64 64" width="80" height="80" style="position:absolute;left:100px;top:120px"><circle cx="32" cy="32" r="25" fill="#147d92"/></svg></body>');
  const result = await bridge.run('apply_html',{html});
  expect(result.ok, JSON.stringify(result)).toBe(true);
  expect(saved).toBe(1);
  bridge.dispose();
},30000);
