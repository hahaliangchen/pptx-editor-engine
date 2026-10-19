// @vitest-environment node
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { deckFonts } from "../src/api/deck-fonts";

describe("deckFonts", () => {
  it("reports the actual catalog fallback and skips available open families", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", '<a:latin typeface="Calibri"/><a:ea typeface="Brand CJK"/><a:latin typeface="Inter"/>');
    const fonts = await deckFonts(await zip.generateAsync({ type: "uint8array" }), ["Liberation Sans", "Noto Sans CJK SC", "Inter"]);
    expect(fonts).toEqual([
      { family: "Brand CJK", replacement: "Noto Sans CJK SC" },
      { family: "Calibri", replacement: "Liberation Sans" }
    ]);
  });
  it("lists the fonts that are not open-source, with their stand-ins", async () => {
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", '<a:latin typeface="Microsoft YaHei"/><a:ea typeface="阿里巴巴普惠体"/><a:latin typeface="Brand Sans"/><a:latin typeface="+mn-lt"/><a:latin typeface="Carlito"/>');
    zip.file("ppt/theme/theme1.xml", '<a:latin typeface="Arial"/><a:latin typeface="Brand Sans"/>');
    const fonts = await deckFonts(await zip.generateAsync({ type: "uint8array" }));
    // Microsoft YaHei and Arial have a mapped open stand-in and are listed; open families and theme links are not.
    expect(fonts.map(font => font.family)).toEqual(["Arial", "Brand Sans", "Microsoft YaHei", "阿里巴巴普惠体"]);
    expect(fonts.find(font => font.family === "Microsoft YaHei")!.replacement).toBe("Noto Sans CJK SC");
    expect(fonts.find(font => font.family === "阿里巴巴普惠体")!.replacement).toBe("Noto Sans CJK SC");
    expect(fonts.find(font => font.family === "Brand Sans")!.replacement).toBe("Carlito");
  });
});
