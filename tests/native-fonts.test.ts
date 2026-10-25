import { describe, expect, it } from "vitest";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fontMetadata, matchLocalFont, nativeFontDirectories, scanLocalFonts, type LocalFont } from "../src/headless/native-fonts";

function font(family: string, style: string): Buffer {
  const strings = [family, style].map(text => Buffer.from(text, "utf16le").swap16());
  const name = Buffer.alloc(30 + strings.reduce((n, s) => n + s.length, 0));
  name.writeUInt16BE(2, 2); name.writeUInt16BE(30, 4);
  let offset = 0;
  strings.forEach((bytes, index) => {
    const record = 6 + index * 12;
    name.writeUInt16BE(3, record); name.writeUInt16BE(1, record + 2); name.writeUInt16BE(0x409, record + 4);
    name.writeUInt16BE(index + 1, record + 6); name.writeUInt16BE(bytes.length, record + 8); name.writeUInt16BE(offset, record + 10);
    bytes.copy(name, 30 + offset); offset += bytes.length;
  });
  const data = Buffer.alloc(28 + name.length);
  data.writeUInt32BE(0x10000); data.writeUInt16BE(1, 4); data.write("name", 12);
  data.writeUInt32BE(28, 20); data.writeUInt32BE(name.length, 24); name.copy(data, 28);
  return data;
}

describe("native open-source fonts", () => {
  it("reads family/style from OpenType metadata, including a collection", () => {
    const bytes = font("Carlito", "Bold Italic");
    expect(fontMetadata(bytes, "random.ttf")).toEqual([{family:"Carlito", file:"random.ttf", weight:700, italic:true}]);
    const ttc = Buffer.alloc(16 + bytes.length); ttc.write("ttcf"); ttc.writeUInt32BE(1, 8); ttc.writeUInt32BE(16, 12);
    bytes.copy(ttc,16); ttc.writeUInt32BE(44,36);
    expect(fontMetadata(ttc,"font.ttc")[0].family).toBe("Carlito");
    expect(fontMetadata(Buffer.from("broken"), "broken.ttf")).toEqual([]);
  });
  it("skips corrupt files and rejects unapproved families despite their filenames", () => {
    const dir = mkdtempSync(join(tmpdir(), "ppt-fonts-"));
    try {
      writeFileSync(join(dir,"arbitrary.ttf"),font("Carlito","Regular"));
      writeFileSync(join(dir,"Carlito-Bold.ttf"),font("Microsoft YaHei","Bold"));
      writeFileSync(join(dir,"broken.ttf"),Buffer.from("broken"));
      expect(scanLocalFonts([join(dir,"missing"),dir]).map(f=>f.family)).toEqual(["Carlito"]);
    } finally { rmSync(dir,{recursive:true,force:true}); }
  });
  it("matches bold/italic and uses the same Office stand-ins as the browser", () => {
    const faces: LocalFont[] = [
      {family:"Carlito",file:"regular",weight:400,italic:false}, {family:"Carlito",file:"bold",weight:700,italic:false},
      {family:"Carlito",file:"italic",weight:400,italic:true}, {family:"Carlito",file:"bold-italic",weight:700,italic:true},
      {family:"Noto Sans CJK SC",file:"chinese",weight:400,italic:false}
    ];
    for (const [bold,italic,file] of [[false,false,"regular"],[true,false,"bold"],[false,true,"italic"],[true,true,"bold-italic"]] as const) {
      expect(matchLocalFont(faces,{family:"Calibri",bold,italic,eastAsian:false})).toBe(file);
    }
    expect(matchLocalFont(faces,{family:"微软雅黑",bold:false,italic:false,eastAsian:true})).toBe("chinese");
    expect(()=>matchLocalFont([],{family:"Calibri",bold:false,italic:false,eastAsian:false})).toThrow(/PPT_FONT_DIRS/);
  });
  it("includes Windows system and per-user font directories", () => {
    expect(nativeFontDirectories({WINDIR:"C:/Windows",LOCALAPPDATA:"C:/Users/test/AppData/Local"},"win32")).toHaveLength(2);
  });
});
