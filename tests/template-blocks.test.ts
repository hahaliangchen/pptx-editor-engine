// @vitest-environment node
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { convertDesign } from '../src/design/service';
import { DesignBridge } from '../src/page/design-bridge';

const json = (result: {ok: boolean; content: Array<{text?: string}>}) => { expect(result.ok, result.content[0].text).toBe(true); return JSON.parse(result.content[0].text!); };

it('lets the template author group a card into a block, name its fields and fill it like a template-built page', async () => {
  let current = readFileSync('tools/templates/blank.pptx');
  const viewer = {snapshotPptx: async () => new Blob([current]), loadPptx: async (buffer: ArrayBuffer) => { current = Buffer.from(buffer); }, getCurrentSlideIndex: () => 0, getSlidesCount: () => 1, selectSlide: async () => {}};
  const bridge = new DesignBridge(viewer, convertDesign, undefined, {mode: 'template-authoring'});
  json(await bridge.updateTemplateSemantics(1, [{fill: 'layout'}]));
  let page = json(await bridge.run('get_design', {page: 1}));
  expect(page.layoutIssues.map((i: {message: string}) => i.message).join()).toContain('没有块');
  const card = '<div data-kind="shape" data-id="new:tile" style="position:absolute;left:100px;top:200px;width:300px;height:200px;background:#f4f4f4;border-radius:12px"></div>'
    + '<div data-kind="text" data-id="new:head" style="position:absolute;left:120px;top:220px;width:260px;height:40px"><p><span style="font-size:24px;color:#c00000">示例标题</span></p></div>';
  json(await bridge.run('apply_html', {html: page.html.replace('</body>', `${card}</body>`)}));
  page = json(await bridge.run('get_design', {page: 1}));
  const ids = page.nodes.filter((n: {name: string}) => /tile|head/.test(n.name)).map((n: {id: string}) => n.id);
  expect(ids).toHaveLength(2);
  json(await bridge.groupTemplateNodes(1, ids, '卡片#1'));
  page = json(await bridge.run('get_design', {page: 1}));
  const block = page.nodes.find((n: {slot?: string}) => n.slot === '卡片#1');
  expect(block.kind).toBe('group');
  // The rebuilt members keep the names the author saw; these originals still hold the ids "new:tile" and "new:head".
  expect(page.nodes.filter((n: {parentId?: string}) => n.parentId === block.id).map((n: {name: string}) => n.name)).toEqual(['new:tile-2', 'new:head-2']);
  expect(page.layoutIssues.map((i: {message: string}) => i.message).join()).toContain('没有字段名');
  const text = page.nodes.find((n: {parentId?: string; kind: string}) => n.parentId === block.id && n.kind === 'text');
  json(await bridge.updateTemplateSemantics(1, [{nodeId: text.id, purpose: 'content', slot: 'title'}]));
  expect(json(await bridge.run('get_design', {page: 1})).layoutIssues).toEqual([]);
  const layout = json(await bridge.run('get_layout', {page: 1}));
  expect(layout.sets[0]).toMatchObject({set: '卡片#1', items: [{fields: {title: ['示例标题']}}]});
  json(await bridge.run('fill_layout', {page: 1, sets: [{set: '卡片#1', items: [{fields: {title: '一'}}, {fields: {title: '二'}}, {fields: {title: '三'}}]}]}));
  const filled = json(await bridge.run('get_design', {page: 1, html: true}));
  for (const word of ['一', '二', '三']) expect(filled.html).toContain(`>${word}</span>`);
  bridge.dispose();
}, 60000);
