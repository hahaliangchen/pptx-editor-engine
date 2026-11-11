// @vitest-environment node
import {readFileSync, existsSync} from "node:fs";
import {describe,expect,it} from "vitest";
import {browserFontBytes,copyScriptFontFaces} from "../src/page/browser-fonts";
import {fontMetadata} from "../src/headless/native-fonts";

describe("browser font compatibility",()=>{
  it("preserves standalone fonts and rejects truncated collections",()=>{
    const bytes=new Uint8Array([0,1,0,0,1,2,3]);expect(browserFontBytes(bytes,"Test",400,false)).toBe(bytes);
    expect(()=>browserFontBytes(new Uint8Array([116,116,99,102]),"Test",400,false)).toThrow("FONT_COLLECTION_INVALID");
    const broken=new Uint8Array(16);broken.set([116,116,99,102]);new DataView(broken.buffer).setUint32(8,1025);
    expect(()=>browserFontBytes(broken,"Test",400,false)).toThrow("FONT_COLLECTION_INVALID");
  });
  const file="/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc";
  it.skipIf(!existsSync(file))("extracts the requested CJK face with valid offsets and SFNT checksum, preserving source bytes",()=>{
    const bytes=new Uint8Array(readFileSync(file)),before=bytes.slice();
    expect(fontMetadata(Buffer.from(bytes),file).length).toBeGreaterThan(1);
    for(const family of ["Noto Sans CJK SC","Noto Sans CJK JP"]){
      const single=browserFontBytes(bytes,family,400,false);
      expect(Buffer.from(single.subarray(0,4)).toString()).toBe("OTTO");
      expect(fontMetadata(Buffer.from(single),"standalone.otf").map(face=>face.family)).toEqual([family]);
      const view=new DataView(single.buffer);let sum=0;
      for(let i=0;i<single.length;i+=4)sum=(sum+view.getUint32(i))>>>0;
      expect(sum).toBe(0xb1b0afba);
      for(let i=0;i<view.getUint16(4);i++){
        const start=view.getUint32(12+16*i+8),length=view.getUint32(12+16*i+12);
        expect(start%4).toBe(0);expect(start+length).toBeLessThanOrEqual(single.length);
      }
    }
    expect(Buffer.from(bytes).equals(Buffer.from(before))).toBe(true);
    expect(()=>browserFontBytes(bytes,"Absent family",400,false)).toThrow("FONT_FACE_NOT_FOUND");
  });
  it("ignores CSS-connected faces in generic copying, but propagates unrelated failures",()=>{
    const css={} as FontFace,script={} as FontFace,copied:FontFace[]=[];
    const source={forEach:(fn:(face:FontFace)=>void)=>[css,script].forEach(fn)} as FontFaceSet;
    const target={add:(face:FontFace)=>{if(face===css)throw Object.assign(Error("CSS-connected"),{name:"InvalidModificationError"});copied.push(face);}} as unknown as FontFaceSet;
    copyScriptFontFaces(source,target);expect(copied).toEqual([script]);
    expect(()=>copyScriptFontFaces(source,{add:()=>{throw Error("Other failure");}} as unknown as FontFaceSet)).toThrow("Other failure");
  });
});
