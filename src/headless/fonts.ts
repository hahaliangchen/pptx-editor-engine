import { execFileSync } from "node:child_process";
import { expandForRender } from "../render-elements";
import { PptVirtualDocument } from "../pptx-virtual-dom";
import { OPEN_FONT_MAP } from "../page/font-map";
import { scanLocalFonts, nativeFontDirectories, matchLocalFont, type LocalFont } from "./native-fonts";

export interface FontRequest {
  family: string;
  bold: boolean;
  italic: boolean;
  eastAsian: boolean;
}

type TextStyle = { fontFamily?: string; eastAsianFontFamily?: string; bold?: boolean; italic?: boolean };

/** Every family/weight/slant the deck's text asks for, split into Latin and East Asian. */
export function collectFontRequests(doc: PptVirtualDocument): FontRequest[] {
  const requests = new Map<string, FontRequest>();
  const add = (family: string | undefined, eastAsian: boolean, bold = false, italic = false) => {
    if (!family) return;
    requests.set(`${family.toLowerCase()}|${bold}|${italic}|${eastAsian}`, { family, bold, italic, eastAsian });
  };
  const addStyle = (style: TextStyle) => {
    add(style.fontFamily, false, !!style.bold, !!style.italic);
    add(style.eastAsianFontFamily, true, !!style.bold, !!style.italic);
  };
  for (const slide of doc.slides) {
    for (const element of expandForRender(slide.elements)) {
      if (element.type !== "text") continue;
      addStyle(element.style);
      for (const paragraph of element.paragraphs || []) {
        for (const run of paragraph.runs) addStyle(run.style);
        add(paragraph.bullet?.fontFamily, false);
      }
    }
  }
  if (requests.size === 0) add("sans-serif", false);
  return [...requests.values()];
}

/**
 * Maps a requested font through native OpenType discovery on Windows or when
 * PPT_FONT_DIRS is configured, otherwise through fontconfig. Both supply the
 * open-source stand-ins used by the browser viewer.
 */
export class FontResolver {
  private cache = new Map<string, string>();
  private localFonts?: LocalFont[];

  public resolve(request: FontRequest): string {
    const key = `${request.family.toLowerCase()}|${request.bold}|${request.italic}|${request.eastAsian}`;
    const cached = this.cache.get(key);
    if (cached) return cached;
    if (process.platform === "win32" || process.env.PPT_FONT_DIRS) {
      this.localFonts ??= scanLocalFonts(nativeFontDirectories());
      const file = matchLocalFont(this.localFonts, request);
      this.cache.set(key, file);
      return file;
    }
    // Same open-source stand-ins as the browser, so both draw the same glyphs.
    const family = OPEN_FONT_MAP[request.family.trim().toLowerCase()] ?? request.family;
    const pattern = [
      family,
      `weight=${request.bold ? "bold" : "regular"}`,
      `slant=${request.italic ? "italic" : "roman"}`,
      ...(request.eastAsian ? ["lang=zh-cn"] : [])
    ].join(":");
    let file: string;
    try {
      file = execFileSync("fc-match", ["-f", "%{file}", pattern], { encoding: "utf8" }).trim();
    } catch (error) {
      throw new Error(`无法通过 fontconfig 查找字体“${request.family}”，请安装 fontconfig 与中文字体（如 fonts-noto-cjk）：${error}`);
    }
    if (!file) throw new Error(`fontconfig 没有返回字体“${request.family}”的文件。`);
    this.cache.set(key, file);
    return file;
  }
}
