// @vitest-environment node
import {Blob} from "node:buffer";
import {readFileSync} from "node:fs";
import {describe, expect, it} from "vitest";
import {convertDesign, materializeDesign} from "../src/design/service";
import {DesignBridge} from "../src/page/design-bridge";

const blank = readFileSync("tools/templates/blank.pptx");
const value = (result: any) => {
  expect(result.ok, JSON.stringify(result)).toBe(true);
  return JSON.parse(result.content[0].text);
};
async function editor(brand = false) {
  let bytes = new Uint8Array(blank);
  if (brand) {
    const doc = await convertDesign({operation:"sync",pptxBase64:Buffer.from(bytes).toString("base64")});
    doc.pages[0].nodes.push({id:"new:brand",kind:"shape",name:"Fixed brand",purpose:"fixed",appearance:{rect:{x:20,y:20,w:80,h:40},fill:"#d40000"}});
    bytes = new Uint8Array(Buffer.from((await materializeDesign(doc)).source.pptxBase64,"base64"));
  }
  let loads = 0, selections = 0;
  const bridge = new DesignBridge({
    snapshotPptx:async()=>new Blob([bytes]),
    loadPptx:async buffer=>{loads++;bytes=new Uint8Array(buffer);},
    getCurrentSlideIndex:()=>0, getSlidesCount:()=>1,
    selectSlide:async()=>{selections++;},
  }, convertDesign);
  const baseline = value(await bridge.run("get_design",{page:1}));
  return {bridge,baseline,bytes:()=>Buffer.from(bytes).toString("base64"),loads:()=>loads,selections:()=>selections};
}

describe("read-only page preflight",()=>{
  it("locates an unmarked painted card and leaves the deck intact; the repaired HTML can be checked then committed",async()=>{
    const e=await editor(),before=e.bytes();
    const bad=e.baseline.html.replace("</body>",'<div style="position:absolute;left:100px;top:100px;width:500px;height:200px;background:#fafafa;border-top:3px solid #d40000"><div data-kind="text"><p>Required actual content</p></div></div></body>');
    const failed=await e.bridge.run("validate_html",{html:bad});
    expect(failed.ok).toBe(false);
    const text=(failed.content[0] as {text:string}).text;
    const details=JSON.parse(text.slice(text.indexOf("{")).split("\n")[0]);
    expect(details).toMatchObject({code:"UNMARKED_PAINTED_CONTAINER",pageId:e.baseline.pageId,tag:"div"});
    expect(details.selector).toMatch(/^body > div:nth-child\(/);
    expect(details.htmlSnippet).toContain("background:#fafafa");
    expect(details.recovery).toContain("validate_html");
    const good=e.baseline.html.replace("</body>",'<div data-kind="text" style="position:absolute;left:100px;top:100px;width:500px;height:200px;background:#fafafa;border:3px solid #d40000"><p>Required actual content</p></div></body>');
    const checked=value(await e.bridge.run("validate_html",{html:good}));
    expect(checked).toMatchObject({validated:true,applied:false,needsSave:false,content_complete:false});
    expect(checked.changes.some((change:any)=>change.action==="add")).toBe(true);
    expect(e.bytes()).toBe(before);expect(e.loads()).toBe(0);expect(e.selections()).toBe(0);expect(e.bridge.pending).toBe(false);
    expect(e.bridge.peekHtml(1)).toBe(e.baseline.html);
    const applied=value(await e.bridge.run("apply_html",{html:good,content_complete:true}));
    expect(applied).toMatchObject({applied:true,content_complete:true,needsSave:true});
    expect(e.loads()).toBe(1);expect(e.bridge.pending).toBe(true);expect(e.bytes()).not.toBe(before);
  },30000);
  it("does not weaken fixed protection or turn ordinary edits into content completion",async()=>{
    const e=await editor(true),before=e.bytes();
    const bad=e.baseline.html.replace("left:20px;top:20px","left:40px;top:20px");
    expect((await e.bridge.run("validate_html",{html:bad})).ok).toBe(false);
    expect(e.bytes()).toBe(before);expect(e.bridge.pending).toBe(false);
    const edit=e.baseline.html.replace("</body>",'<div data-kind="text" style="position:absolute;left:100px;top:120px;width:500px;height:100px"><p>Partial draft</p></div></body>');
    expect(value(await e.bridge.run("apply_html",{html:edit})).content_complete).toBe(false);
    const current=value(await e.bridge.run("get_design",{page:1}));
    expect(value(await e.bridge.run("apply_html",{html:current.html,content_complete:true}))).toMatchObject({applied:false,content_complete:true});
  },30000);
});
