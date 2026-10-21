import {readFileSync,writeFileSync,mkdirSync} from 'node:fs';
import {join,dirname} from 'node:path';
import assert from 'node:assert/strict';
import {HeadlessSession} from '../src/headless/session';
import {inferTemplate,describeTemplate} from '../src/api/template';
import {checkTemplate} from '../src/api/template-authoring';

const [input,output,prepare] = process.argv.slice(2);
if (!input || !output) throw new Error('usage: verify-purpose-template.ts template.pptx output-directory [--prepare]');
mkdirSync(output,{recursive:true});
let session=await HeadlessSession.open(readFileSync(input),{quiet:true});
if(prepare==='--prepare') {
 const bounds:Record<number,Record<string,number[]>>={
  1:{'29':[1.5,2.3,12,.8],'30':[4,6.5,26,2.3],'31':[5,9,24,1.2],'33':[6,11.4,22,1.2],'35':[12,15,10,.7],'36':[23,17.6,9,.6]},
  2:{'5':[1.5,1.8,29,1.4]},3:{'14':[13,4,9,.8],'16':[4,8,26,2],'17':[5,10.8,24,1]},
  4:{'4':[1.5,1.1,12,.6],'5':[1.5,1.8,29,1.2],'6':[1.5,3,29,.8]},5:{'26':[4,6.5,26,2],'28':[5,9.8,24,1]}
 };
 const commands:any[]=[];
 for(const [page,items]of Object.entries(bounds))for(const[id,[x,y,w,h]]of Object.entries(items))commands.push({op:'setBounds',target:{slide:Number(page),id},x,y,w,h});
 for(const [id,text]of Object.entries({'12':'处理明确任务','25':'按规则编排步骤','38':'按反馈动态决策','13':'输入清晰问题','14':'输出可以验证','26':'分解目标步骤','27':'关键节点校验','39':'根据环境决策','40':'调用工具执行'}))commands.push({op:'setText',target:{slide:4,id},text});
 commands.push({op:'formatText',target:{slide:4,id:'5'},fontSizePt:25},{op:'formatText',target:{slide:4,id:'6'},fontSizePt:18});
 const changed=await session.apply({commands});assert(changed.ok,JSON.stringify(changed));
 writeFileSync(input,await session.export());session.dispose();session=await HeadlessSession.open(readFileSync(input),{quiet:true});
}
const outline=await session.outline({maxTextLength:5000});
const issues=checkTemplate(outline);assert.deepEqual(issues.filter(i=>i.severity==='error'),[]);
assert(!JSON.stringify(outline).includes('{{'));
const definition=describeTemplate(inferTemplate(outline)) as any;
assert.deepEqual(Object.keys(definition.pages),['cover','toc','section','content','end']);
for(const page of Object.values(definition.pages) as any[])assert.equal(page.readyForGeneration,true);
assert.equal(definition.pages.toc.referenceVisuals[0].id,'52');
assert.equal(definition.pages.content.referenceVisuals[0].id,'48');
assert(definition.pages.content.contentAreaMarkerIds.includes('46'));
assert(!definition.pages.content.preserveShapeIds.includes('48'));
assert.equal(definition.pages.toc.entries.length,6);
assert(!JSON.stringify(definition.pages.content.annotations).match(/💡|⚡|⚙/));
writeFileSync(join(output,'anthropic-definition.json'),JSON.stringify(definition,null,2));
writeFileSync(join(output,'template-outline.json'),JSON.stringify(outline,null,2));
writeFileSync(join(output,'template-summary-reference.png'),await session.render(2,{width:1280}));
writeFileSync(join(output,'template-content-reference.png'),await session.render(4,{width:1280}));
writeFileSync(join(dirname(input),'thumbnail.png'),await session.render(1,{width:512}));
const commands:any[]=[];
const toc=definition.pages.toc;
for(const entry of toc.entries.slice(3))for(const id of entry.shapeIds)commands.push({op:'delete',target:{slide:2,id}});
commands.push({op:'delete',target:{slide:2,id:toc.referenceVisuals[0].id}},{op:'setText',target:{slide:2,id:toc.fieldInfo.title.id},text:'用途定义如何驱动原生设计'});
for(const [index,[title,desc]]of [['标注用途','明确哪些填写、重画或保留'],['说明关系','让 AI 理解图形的表达意图'],['验证结果','检查真实内容、编辑性与渲染']].entries()) {
 const entry=toc.entries[index];commands.push({op:'setText',target:{slide:2,id:entry.titleId},text:title},{op:'setText',target:{slide:2,id:entry.descId},text:desc});
}
function box(name:string,text:string,x:number,y:number,w:number,h:number,color:string){commands.push({op:'addShape',slide:2,name,shape:'roundRect',rect:{x,y,w,h},fill:color,line:null},{op:'addText',slide:2,name:name+'文字',text,rect:{x:x+.15,y:y+.25,w:w-.3,h:h-.4},style:{fontSizePt:16,bold:true,color:'#1A1A2E',align:'center'}});}
box('用途输入','用途标注',20.5,4.8,3.7,1.5,'#E9F2FC');box('示例输入','真实示例',25.5,4.8,3.7,1.5,'#E9F2FC');box('说明输入','页面说明',23,7.5,3.7,1.5,'#E9F2FC');
box('理解节点','理解表达意图',21.5,10.5,6.7,1.7,'#F8EAE4');box('设计输出','生成新的原生结构',21.5,14.2,6.7,1.7,'#E6F6EF');
for(const [x,y,w,h,shape]of [[22,6.4,.5,3.6,'arrow'],[27.3,6.4,.5,3.6,'arrow'],[24.6,9.2,.5,1,'arrow'],[24.6,12.5,.5,1.2,'arrow']] as const)commands.push({op:'addShape',slide:2,shape,rect:{x:x+w/2,y,w:0.01,h},fill:null,line:{color:'#64748B',widthPt:1.5}});
const changed=await session.apply({commands});assert(changed.ok,JSON.stringify(changed));
writeFileSync(join(output,'anthropic-purpose-acceptance.pptx'),await session.export());
session.dispose();session=await HeadlessSession.open(readFileSync(join(output,'anthropic-purpose-acceptance.pptx')),{quiet:true});
const result=await session.outline({maxTextLength:5000});const words=JSON.stringify(result.slides[1]);
assert(!words.includes('Single LLM Call'));assert(!words.includes('@目录.6.标题'));assert(words.includes('理解表达意图'));assert(result.slides[1].shapes.some(s=>s.name==='理解节点'));
writeFileSync(join(output,'summary-redraw-acceptance.png'),await session.render(2,{width:1280}));
writeFileSync(join(output,'verification.json'),JSON.stringify({status:'passed',template:'anthropic',definitionVersion:definition.definitionVersion,templateErrors:[],referenceGroups:[52,48],unusedEntriesRemoved:true,summaryRebuiltWithNativeObjects:true,exportReopenPassed:true,otherTemplatesConverted:0},null,2));
session.dispose();console.log('Anthropic purpose template: structure, explicit native redraw, export/reopen and renders passed.');
