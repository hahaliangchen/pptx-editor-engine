import { batchSchema } from "./schema";
import type { ApplyResult } from "./engine";
import type { Outline, OutlineOptions } from "./outline";

/**
 * The tools a model is given to read and change a deck. Nothing here belongs to
 * one vendor: a tool is a name, a description and a JSON Schema for its input,
 * and `callTool` runs one against a deck. `adapters.ts` turns these into the
 * shapes Anthropic, OpenAI and MCP each expect.
 */
export interface ToolDefinition {
  name: string;
  description: string;
  /** JSON Schema of the arguments; always an object at the top. */
  inputSchema: Record<string, unknown>;
}

/** What `callTool` needs from a deck. A `HeadlessSession` is one; so is a browser viewer's `EditApi` with a renderer. */
export interface ToolHost {
  outline(options?: OutlineOptions): Promise<Outline>;
  apply(input: unknown): Promise<ApplyResult>;
  /** Page `page` (1-based) as a PNG. Absent when the host cannot draw. */
  render?(page: number, options?: { width?: number }): Promise<Uint8Array>;
}

export type ToolContent =
  | { type: "text"; text: string }
  | { type: "image"; mimeType: string; /** base64 */ data: string };

export interface ToolResult {
  /** False when the call failed in a way the model should react to. */
  ok: boolean;
  content: ToolContent[];
}

export const RENDER_WIDTH = { min: 160, max: 2560, default: 1024 };

export const TOOLS: ToolDefinition[] = [
  {
    name: "get_outline",
    description:
      "读取演示文稿现在的结构：版式名称列表，以及每一页上每个对象的 id、名称、类型"
      + "（text、shape、image、line、table、chart、group）、位置和大小（厘米）、文字、样式；"
      + "表格给出各行各列的文字，图表给出类型、分类和各系列的数据。"
      + "修改之前先调用它，才知道要改的对象叫什么、在哪里。文稿很大时用 slides 只读需要的页。",
    inputSchema: {
      type: "object",
      properties: {
        slides: {
          type: "array", items: { type: "integer", minimum: 1 }, minItems: 1,
          description: "只返回这些页（页码从 1 开始）；不写就返回全部"
        },
        maxTextLength: {
          type: "integer", minimum: 20,
          description: "每个对象的文字超过这个长度就截断（默认 500），被截断的对象带 textTruncated: true"
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "apply_edits",
    description:
      "用一批指令修改演示文稿：新增页面、文字、形状、图片、表格、图表，修改文字与样式，移动缩放，复制，删除等。"
      + "整批默认是原子的：任何一条失败，这一批的所有改动都会撤销，并返回失败的那一条的位置和原因"
      + "（找不到或不唯一的对象会附上候选）。页码从 1 开始，长度单位是厘米，颜色写 #RRGGBB。"
      + "用 target 指向对象：{slide, id | name | placeholder | text}，或 \"$别名\"（本批前面用 as 创建的对象）。"
      + "一批里可以放同一页的全部改动。改完后请用 render_slide 检查排版。",
    inputSchema: batchSchema as unknown as Record<string, unknown>
  },
  {
    name: "render_slide",
    description:
      "把一页渲染成 PNG 图片，用来检查：文字有没有溢出文本框、对象有没有重叠或超出页面、"
      + "对齐是否整齐、图表和表格是否清楚。每次较大的修改之后都应该看一眼。",
    inputSchema: {
      type: "object",
      properties: {
        slide: { type: "integer", minimum: 1, description: "页码，从 1 开始" },
        width: {
          type: "integer", minimum: RENDER_WIDTH.min, maximum: RENDER_WIDTH.max,
          description: `图片宽度（像素），默认 ${RENDER_WIDTH.default}；高度按页面比例`
        }
      },
      required: ["slide"],
      additionalProperties: false
    }
  }
];

/** Text to add to a model's system prompt: how to work with these tools. */
export const TOOL_GUIDE = [
  "你可以用三个工具读取和修改一份 PPTX 演示文稿：get_outline、apply_edits、render_slide。",
  "",
  "工作方式：",
  "1. 先 get_outline 看清现有内容，再动手；不要凭空猜对象的名称。",
  "2. 用 apply_edits 一次提交一页（或几页）的全部改动，而不是一条一条地调用。",
  "3. 需要引用刚创建的对象时，给它写 as 别名，后面的指令用 \"$别名\" 指向它。",
  "4. 改完用 render_slide 看效果：文字是否溢出、对象是否重叠或超出页面、颜色是否清楚；有问题就再改。",
  "5. apply_edits 失败时，读返回的 error.message、error.path 和 candidates，改正那一条后重新提交；失败的批次不会留下任何改动。",
  "",
  "约定：页码从 1 开始；长度单位是厘米；颜色是 #RRGGBB；字号是磅（fontSizePt）。",
  "数据用 addChart（柱形 column、条形 bar、折线 line、饼图 pie）；成行成列的内容用 addTable；",
  "已有的图表和表格用 setChart、setTable、setCell 改数据，它们会保留原来的样式。",
  "修改已有对象时尽量保留它原来的格式：改文字用 setText 或 replaceText，不要删掉重建。"
].join("\n");

const bytesToBase64 = (bytes: Uint8Array): string => {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(binary);
};

const text = (value: unknown): ToolContent => ({
  type: "text",
  text: typeof value === "string" ? value : JSON.stringify(value)
});

const failure = (message: string, extra: Record<string, unknown> = {}): ToolResult => ({
  ok: false,
  content: [text({ ok: false, error: { message, ...extra } })]
});

function integerList(value: unknown, field: string): number[] | string {
  if (!Array.isArray(value) || value.length === 0 || value.some(item => !Number.isInteger(item) || item < 1)) {
    return `${field} 必须是页码（从 1 开始的整数）组成的非空数组`;
  }
  return value as number[];
}

/**
 * Run one tool. It never throws: a bad argument or a failed edit comes back as
 * a result with `ok: false` whose text says what to fix, so a model can retry.
 */
export async function callTool(host: ToolHost, name: string, args: unknown): Promise<ToolResult> {
  const input = (args ?? {}) as Record<string, unknown>;
  try {
    switch (name) {
      case "get_outline": {
        const options: OutlineOptions = {};
        if (input.slides !== undefined) {
          const slides = integerList(input.slides, "slides");
          if (typeof slides === "string") return failure(slides);
          options.slides = slides;
        }
        if (input.maxTextLength !== undefined) {
          if (!Number.isInteger(input.maxTextLength) || (input.maxTextLength as number) < 20) {
            return failure("maxTextLength 必须是不小于 20 的整数");
          }
          options.maxTextLength = input.maxTextLength as number;
        }
        return { ok: true, content: [text(await host.outline(options))] };
      }

      case "apply_edits": {
        const result = await host.apply(input);
        return { ok: result.ok, content: [text(result)] };
      }

      case "render_slide": {
        if (!host.render) return failure("这个环境不能渲染图片");
        if (!Number.isInteger(input.slide) || (input.slide as number) < 1) return failure("slide 必须是从 1 开始的页码");
        let width = RENDER_WIDTH.default;
        if (input.width !== undefined) {
          if (!Number.isInteger(input.width) || (input.width as number) < RENDER_WIDTH.min || (input.width as number) > RENDER_WIDTH.max) {
            return failure(`width 必须是 ${RENDER_WIDTH.min} 到 ${RENDER_WIDTH.max} 之间的整数`);
          }
          width = input.width as number;
        }
        const png = await host.render(input.slide as number, { width });
        return {
          ok: true,
          content: [
            { type: "image", mimeType: "image/png", data: bytesToBase64(png) },
            text(`第 ${input.slide} 页，宽 ${width} 像素`)
          ]
        };
      }

      default:
        return failure(`没有名为 "${name}" 的工具`, { available: TOOLS.map(tool => tool.name) });
    }
  } catch (error) {
    return failure(error instanceof Error ? error.message : String(error));
  }
}
