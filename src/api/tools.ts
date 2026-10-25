import { ApiError } from "./errors";
import { batchSchema } from "./schema";
import type { ApplyResult } from "./engine";
import type { Outline, OutlineOptions } from "./outline";
import { compileDeck } from "./deck-spec";
import { describeTemplate, inferTemplate, type PageKind } from "./template";
import { suggestAnnotations, suggestRoles } from "./annotation-suggest";
import { checkLayout } from "./layout-check";
import { checkTemplate, slideRoles } from "./template-authoring";

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
  /** `meta` says what an image shows (render_slide: `{ slide }`); hosts keep it with the image. */
  | { type: "image"; mimeType: string; /** base64 */ data: string; meta?: Record<string, unknown> };

export interface ToolResult {
  /** False when the call failed in a way the model should react to. */
  ok: boolean;
  content: ToolContent[];
  /** Saved document version inspected by the hosting application, when known. */
  document?: { project_id: number; stored_revision: number; stable: boolean };
}

export const RENDER_WIDTH = { min: 160, max: 2560, default: 1024 };

/** Pages one render_slide call may draw: a 3×3 review sheet. */
export const MAX_RENDER_SLIDES = 9;

/** A deck with more pages than this gets an overview first, unless the model asks for pages or for the full outline. */
export const BRIEF_ABOVE = 8;

interface OutlineNode { type: string; text?: string; placeholder?: string; children?: OutlineNode[] }

/** One line per page: layout, title and what is on it. The detail is a `slides` call away. */
function briefOf(outline: Outline): unknown {
  const count = (nodes: OutlineNode[], into: Record<string, number>) => {
    for (const node of nodes) {
      into[node.type] = (into[node.type] ?? 0) + 1;
      if (node.children) count(node.children, into);
    }
    return into;
  };
  return {
    slideSize: outline.slideSize,
    layouts: outline.layouts,
    note: `文稿共 ${outline.slides.length} 页，这里只给概要；要看某几页的对象、位置、文字和数据，请用 slides 指定页码再调用一次。`,
    slides: outline.slides.map(slide => {
      const shapes = slide.shapes as unknown as OutlineNode[];
      const title = shapes.find(shape => shape.placeholder === "title" || shape.placeholder === "ctrTitle")
        ?? shapes.find(shape => shape.type === "text" && shape.text);
      return {
        index: slide.index,
        id: slide.id,
        appearance: slide.appearance,
        ...(slide.layout ? { layout: slide.layout } : {}),
        ...(title?.text ? { title: title.text.length > 40 ? `${title.text.slice(0, 40)}…` : title.text } : {}),
        contents: count(shapes, {})
      };
    })
  };
}

export const TOOLS: ToolDefinition[] = [
  {
    name: "get_outline",
    description:
      "读取演示文稿现在的结构：版式名称列表，以及每一页上每个对象的 id、名称、类型"
      + "（text、shape、image、line、table、chart、group）、位置和大小（厘米）、文字、样式；"
      + "表格给出各行各列的文字，图表给出类型、分类和各系列的数据。"
      + "appearance 告诉你实际铺底后的深浅主题，可能不同于页面 background 属性；textBox 给出边距、换行和溢出策略。修改之前先调用它，才知道要改的对象叫什么、在哪里。页数很多的文稿先给概要，再用 slides 读需要的页。",
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
        },
        detail: {
          type: "string", enum: ["brief", "full"],
          description: `brief 只给每页的版式、标题和对象数量；full 给全部细节。不写时：指定了 slides 或不超过 ${BRIEF_ABOVE} 页给 full，页数更多给 brief`
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
      + "连接关系用 addConnector 明确 from/to 端点（厘米），原生箭头从 from 指向 to；支持上斜、竖直与反向，不需要负尺寸。"
      + "用 target 指向对象：{slide, id | name | placeholder | text}，或 \"$别名\"（本批前面用 as 创建的对象）。"
      + "先单独完成增删复制与排序，再读取 get_template/get_outline 确认最终页码和对象；之后逐页提交绘制批次，避免混用旧页码。"
      + "原子失败返回 rolledBack:true、appliedCommands:0，前面看似成功的编辑也没有保留。一批里可以放同一页的全部改动。删除组合会连同子对象删除，不要再重复删除子 id；失败后先重新读 outline 再修正。文字换行用真实换行符，不要写字面的反斜杠 n。改完后请用 render_slide 检查排版。",
    inputSchema: batchSchema as unknown as Record<string, unknown>
  },
  {
    name: "render_slide",
    description:
      "把页面渲染成 PNG 图片，每页一张，用来检查：文字有没有溢出文本框、对象有没有重叠或超出页面、"
      + `对齐是否整齐、图表和表格是否清楚。一页用 slide；多页用 slides，一次最多 ${MAX_RENDER_SLIDES} 页，图片按给出的页码顺序返回。`,
    inputSchema: {
      type: "object",
      properties: {
        slide: { type: "integer", minimum: 1, description: "页码，从 1 开始" },
        slides: {
          type: "array", minItems: 1, maxItems: MAX_RENDER_SLIDES, uniqueItems: true,
          items: { type: "integer", minimum: 1 }, description: `要渲染的多个页码，从 1 开始，最多 ${MAX_RENDER_SLIDES} 页`
        },
        width: {
          type: "integer", minimum: RENDER_WIDTH.min, maximum: RENDER_WIDTH.max,
          description: `图片宽度（像素），默认 ${RENDER_WIDTH.default}；高度按页面比例`
        }
      },
      additionalProperties: false
    }
  },
  {
    name: "check_layout",
    description:
      "用测量的办法检查版面，不需要看图：文字放不下文本框、对象超出页面、对象互相重叠、内容只占页面上半部分。"
      + "也提示文字与实际不透明底色对比过低、文字被后绘制面板覆盖的风险；这是本地保守提示，不要求模型视觉检查或验收，不拦截保存或导出。只报告，不自动修改布局。"
      + "每个问题带页码、对象 id 和说明，根据实际需要通过 apply_html 修改；没有问题时 issues 是空数组。",
    inputSchema: {
      type: "object",
      properties: { slides: { type: "array", items: { type: "integer", minimum: 1 }, minItems: 1, description: "只检查这些页；不写就检查全部" } },
      additionalProperties: false
    }
  },
  {
    name: "suggest_annotations",
    description:
      "给一份模板草稿的标注出主意（只读，不改文稿）。不带 slide：返回每一页大概是什么角色（封面、目录、章节、内容、结尾）；"
      + "带 slide：返回这一页哪些是模板自己的装饰（在别的页同一位置重复出现）、哪些形状排成等距的一行或一列（重复单元：卡片、步骤、条目，"
      + "按各自的零件配好字段名）、哪些文字像示例文字，并给出可以直接交给 apply_edits 的 setName 命令。"
      + "结果只是猜测：先看 render_slide 的图再用，没归进单元的连线、图标自己补，装饰不要命名。role 指定这页的角色，variant 给内容页的版式起名，unit 指定重复单元的名字（默认“项”）。",
    inputSchema: {
      type: "object",
      properties: {
        slide: { type: "integer", minimum: 1, description: "页码；不写就看所有页的角色" },
        role: { type: "string", enum: ["cover", "toc", "section", "content", "end"] },
        variant: { type: "string", minLength: 1, description: "内容页的版式名，例如 三图文" },
        unit: { type: "string", minLength: 1, description: "重复单元的名字，例如 卡片" }
      },
      additionalProperties: false
    }
  },
  {
    name: "check_template",
    description:
      "检查一份正在制作的模板：封面、目录、内容页、结束页各有至少一页（章节页可选；同一种角色有多页时，除第一页外每页要在形状名字里写版式名），"
      + "每页有角色必须的标记：形状名字以 @ 开头绑定（例如 @封面.标题、@内容.区域、@目录.1.标题），或写 {{关键字}}；"
      + "重复单元（@内容.卡片1.标题 …）要编号连续、字段一致、等距排列，目录项编号连续且有“目录”标题，内容区够大。"
      + "每个问题带页码、对象 id 和说明，severity 为 error 的必须改好才能保存为模板；roles 是引擎认出的每页角色。",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "get_template",
    description:
      "文稿是一份模板（封面、目录、章节、内容、结尾等页面，可变对象由形状名字以 @ 开头说明用途）时，"
      + "按当前形状用途定义读出页面角色、sourceSlide、版式及适用说明、annotations、字段形状 id、真实示例、选项、目录成员、图片位和重复结构。"
      + "referenceVisuals 是按主题重画的参考图，preserveShapeIds 是保留装饰。fieldInfo/fitChars 描述参考文本框，不限制新内容。"
      + "保留模板视觉语言，用 apply_edits 自主设计原生可编辑页面；内容区为空或参考不足时自由编排，不自动缩字、拉伸或分栏。",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "build_deck",
    description:
      "按模板一次生成整份演示文稿：只给每页的类型和内容，版面由引擎排好（缩放文字、对齐、均分），不需要写坐标。"
      + "每页 {page: cover|toc|section|content|end, variant?: 版式名, ...字段}；字段名和版式名见 get_template。"
      + "toc 页写 items: [\"章节名\" 或 {title, desc}]，条数按实际章节数定（1 条到模板的 maxEntries 条），不要凑满；content 页写 title 和 blocks（最多 3 块），"
      + "块有 5 种：{type:\"text\", text}、{type:\"bullets\", items:[...]}、{type:\"cards\", items:[{value, title, text}]}（2 到 4 张最好）、"
      + "{type:\"chart\", kind:column|bar|line|pie, categories, series:[{name, values}], title?}、{type:\"table\", rows:[[...]]}；"
      + "2 个块默认左右并排，也可写 layout: rows|columns。"
      + "replace 把模板里的示例文字换成别的（写空字符串就删掉），可在整份或单页上写。"
      + "只在文稿还是模板时调用一次；生成之后用 apply_edits 精修个别页。这一次是原子的：有一处不对就整份不生成，并告诉你哪里不对。",
    inputSchema: {
      type: "object",
      properties: {
        slides: {
          type: "array", minItems: 1, maxItems: 40,
          items: {
            type: "object",
            properties: { page: { type: "string", enum: ["cover", "toc", "section", "content", "end"] } },
            required: ["page"],
            additionalProperties: true
          },
          description: "按顺序，每页一个对象"
        },
        replace: { type: "object", additionalProperties: { type: "string" }, description: "模板里的示例文字 → 新文字" }
      },
      required: ["slides"],
      additionalProperties: false
    }
  }
];

/** Text to add to a model's system prompt: how to work with these tools. */
export const TOOL_GUIDE = [
  "你可以用三个工具读取和修改一份 PPTX 演示文稿：get_outline、apply_edits、render_slide。",
  "在制作模板时，用形状名字（setName，以 @ 开头）标出每一块的用途，或给页面写 {{关键字}}；先用 suggest_annotations 看每页大概是什么角色、哪些是装饰、哪些是重复结构（它给出现成的 setName 命令，只是猜测，要对照 render_slide 的图），改完用 check_template 检查，修到没有 error。",
  "check_layout 用测量检查版面（溢出、越界、重叠、大片空白），比看图快，改完可以再调用确认。",
  "模板统一使用形状名字里的 @用途 标注和真实示例文字。先读模板 design_spec.md 和 get_template，按 annotations、fieldInfo.id、目录 entries、图片位、重复结构及版式适用说明理解用途；referenceVisuals 按新主题重画，preserveShapeIds 的装饰保留。用 apply_edits 明确设计和填写，不自动布局、拉伸或缩字，不调用 build_deck。",
  "",
  "工作方式：",
  "1. 先 get_outline 看清现有内容，再动手；不要凭空猜对象的名称。页数很多时它先给概要，再用 slides 读你要改的那几页。",
  "2. 用 apply_edits 一次提交一页（或几页）的全部改动，而不是一条一条地调用。",
  "3. 需要引用刚创建的对象时，给它写 as 别名，后面的指令用 \"$别名\" 指向它。",
  "4. 改完用 render_slide 看效果：文字是否溢出、对象是否重叠或超出页面、颜色是否清楚；有问题就再改。",
  "5. 如果现有字段做不到你想要的效果，不要把同样的指令重复提交（它不会有新结果）：换一种办法，或者直接告诉用户这一点做不到。没有直接改 XML 的办法。每页检查修正最多两轮，没有明显问题就结束。",
  "6. apply_edits 失败时，读返回的 error.message、error.path 和 candidates，改正那一条后重新提交；失败的批次不会留下任何改动。",
  "",
  "约定：页码从 1 开始；长度单位是厘米；颜色是 #RRGGBB；字号是磅（fontSizePt）。",
  "新设计的页面背景用 setBackground 设置（纯色或渐变）；已有模板可能用整页形状铺底，实际底色以 get_outline/get_template 的 appearance 和渲染为准，不只看 background 属性。深色底上的文字、图表、表格要用浅色。",
  "版面：新增内容的左边与标题的左边对齐，左右留相同的边距，不要超出页面；表格和图表居中或与标题对齐，不要偏向一边。",
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
        if (input.detail !== undefined && input.detail !== "brief" && input.detail !== "full") {
          return failure('detail 必须是 "brief" 或 "full"');
        }
        const outline = await host.outline(options);
        const brief = input.detail === "brief"
          || (input.detail === undefined && options.slides === undefined && outline.slides.length > BRIEF_ABOVE);
        return { ok: true, content: [text(brief ? briefOf(outline) : outline)] };
      }

      case "apply_edits": {
        const result = await host.apply(input);
        return { ok: result.ok, content: [text(result)] };
      }

      case "render_slide": {
        if (!host.render) return failure("这个环境不能渲染图片");
        let pages: number[];
        if (input.slides !== undefined) {
          if (input.slide !== undefined) return failure("slide 和 slides 只给一个：一页用 slide，多页用 slides");
          const parsed = integerList(input.slides, "slides");
          if (typeof parsed === "string") return failure(parsed);
          if (parsed.length > MAX_RENDER_SLIDES) return failure(`slides 一次最多 ${MAX_RENDER_SLIDES} 页`);
          if (new Set(parsed).size !== parsed.length) return failure("slides 里的页码不能重复");
          pages = parsed;
        } else {
          if (!Number.isInteger(input.slide) || (input.slide as number) < 1) return failure("slide 必须是从 1 开始的页码；多页用 slides");
          pages = [input.slide as number];
        }
        let width = RENDER_WIDTH.default;
        if (input.width !== undefined) {
          if (!Number.isInteger(input.width) || (input.width as number) < RENDER_WIDTH.min || (input.width as number) > RENDER_WIDTH.max) {
            return failure(`width 必须是 ${RENDER_WIDTH.min} 到 ${RENDER_WIDTH.max} 之间的整数`);
          }
          width = input.width as number;
        }
        const content: ToolContent[] = [];
        for (const page of pages) {
          const png = await host.render(page, { width });
          content.push({ type: "image", mimeType: "image/png", data: bytesToBase64(png), meta: { slide: page } });
        }
        content.push(text(`第 ${pages.join("、")} 页，每页一张，宽 ${width} 像素`));
        return { ok: true, content };
      }

      case "check_layout": {
        let slides: number[] | undefined;
        if (input.slides !== undefined) {
          const parsed = integerList(input.slides, "slides");
          if (typeof parsed === "string") return failure(parsed);
          slides = parsed;
        }
        const issues = checkLayout(await host.outline(), { slides });
        return { ok: true, content: [text({ layout_passed: issues.length === 0, scope: "geometry_and_visibility", issues, not_checked: ["content completeness", "icon/glyph rendering", "visual design", "topic coverage"], note: "This checks geometry and conservative text visibility, not full visual or content acceptance. Empty issues do not mean the deck is ready." })] };
      }

      case "check_template": {
        const outline = await host.outline();
        const issues = checkTemplate(outline);
        return {
          ok: true,
          content: [text({ ok: !issues.some(issue => issue.severity === "error"), roles: slideRoles(outline), issues })]
        };
      }

      case "suggest_annotations": {
        const outline = await host.outline({ maxTextLength: 200 });
        const args = (input ?? {}) as { slide?: number; role?: PageKind; variant?: string; unit?: string };
        try {
          if (args.slide === undefined) return { ok: true, content: [text({ pages: suggestRoles(outline) })] };
          return { ok: true, content: [text(suggestAnnotations(outline, args.slide, args))] };
        } catch (error) {
          return failure(error instanceof Error ? error.message : String(error));
        }
      }

      case "get_template": {
        const template = inferTemplate(await host.outline());
        if (Object.keys(template.pages).length === 0) return failure("这份文稿里没有带 @用途 标注或旧 {{TOKEN}} 标记的模板页；请先确认所选模板及当前页面");
        return { ok: true, content: [text(describeTemplate(template))] };
      }

      case "build_deck": {
        const template = inferTemplate(await host.outline());
        if (Object.keys(template.pages).length === 0) {
          return failure("这份文稿里没有模板页，不能用 build_deck；请用 apply_edits 修改现有页面");
        }
        const result = await host.apply(compileDeck(input, template));
        if (!result.ok) return { ok: false, content: [text({ ok: false, failedAt: result.failedAt, error: result.error })] };
        // Say how big the deck is, not what the dozens of commands were.
        const built = await host.outline();
        const issues = checkLayout(built);
        return {
          ok: true,
          content: [text({
            ok: true, slides: built.slides.length, issues,
            note: issues.length > 0 ? "已生成，版面检查发现 issues 里的问题：只修这些页" : "已生成，版面检查没有发现问题"
          })]
        };
      }

      default:
        return failure(`没有名为 "${name}" 的工具`, { available: TOOLS.map(tool => tool.name) });
    }
  } catch (error) {
    if (error instanceof ApiError) return failure(error.message, { code: error.code, ...(error.path ? { path: error.path } : {}) });
    return failure(error instanceof Error ? error.message : String(error));
  }
}
