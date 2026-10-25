import {expect,it} from "vitest";
import {cloneDesign,type DesignDocument,type MeasuredPage} from "../src/design/model";
import {mergeHtmlDifference} from "../src/design/html-diff";

const fixture=()=>{
  const design:DesignDocument={schemaVersion:1,revision:2,canvas:{width:1280,height:720,unit:"css-px"},source:{pptxBase64:"",sha256:"a".repeat(64),widthEMU:1,heightEMU:1},diagnostics:[],pages:[{id:"page",part:"ppt/slides/slide1.xml",nodes:[{id:"text",kind:"text",name:"Title",appearance:{rect:{x:20,y:40,w:200,h:40},text:{padding:[0,0,0,0],verticalAlign:"top",paragraphs:[{align:"left",before:0,after:0,lineHeight:30,runs:[{text:"Before",fontFamily:"Arial",fontSize:24,color:"#000000",bold:false,italic:false,letterSpacing:0}]}]}}}]}]};
  const baseline:MeasuredPage={pageId:"page",revision:2,sourceHash:design.source.sha256,nodes:cloneDesign(design.pages[0].nodes)};
  const edited=cloneDesign(baseline);edited.nodes[0].appearance.text!.paragraphs[0].runs[0].text="After";
  return{design,baseline,edited};
};
it("applies overflowing content while returning its diagnostic",()=>{
  const{design,baseline,edited}=fixture();edited.issues=[{nodeId:"text",message:"Text overflows"}];
  const next=mergeHtmlDifference(design,baseline,edited);
  expect(next.revision).toBe(3);
  expect(next.pages[0].nodes[0].appearance.text!.paragraphs[0].runs[0].text).toBe("After");
  expect(next.diagnostics).toContainEqual({pageId:"page",nodeId:"text",message:"Text overflows",severity:"warning"});
});
it("tolerates one pixel of horizontal glyph overhang and warns for larger extents",()=>{
  const{design,baseline,edited}=fixture();edited.textPaint={text:[{x:19,y:45,w:100,h:20}]};
  expect(mergeHtmlDifference(design,baseline,edited).diagnostics).toEqual([]);
  edited.textPaint.text[0].x=18;
  expect(mergeHtmlDifference(design,baseline,edited).diagnostics.some(d=>d.message.includes("own edges"))).toBe(true);
});
it("accepts old revision metadata while preserving structural checks",()=>{
  const{design,baseline,edited}=fixture();edited.revision=1;
  expect(()=>mergeHtmlDifference(design,baseline,edited)).not.toThrow();
});
it("replaces prior layout warnings only for the page being edited",()=>{
  const{design,baseline,edited}=fixture();design.diagnostics=[{pageId:"page",message:"old warning",severity:"warning"},{pageId:"other",message:"keep",severity:"warning"},{pageId:"page",message:"source note"}];
  expect(mergeHtmlDifference(design,baseline,edited).diagnostics).toEqual([{pageId:"other",message:"keep",severity:"warning"},{pageId:"page",message:"source note"}]);
});
