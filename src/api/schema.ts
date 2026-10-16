import Ajv, { ErrorObject } from "ajv";
import { ApiError } from "./errors";

const color = { type: "string", pattern: "^#[0-9a-fA-F]{6}$", description: "#RRGGBB" };
const cm = { type: "number", description: "centimetres" };
const rect = {
  type: "object",
  properties: { x: cm, y: cm, w: { type: "number", exclusiveMinimum: 0 }, h: { type: "number", minimum: 0 } },
  required: ["x", "y", "w", "h"],
  additionalProperties: false
};
const slideRef = { oneOf: [{ type: "integer", minimum: 1 }, { type: "string", minLength: 1 }], description: "page number from 1, or a slide id" };
const alias = { type: "string", pattern: "^[A-Za-z_][A-Za-z0-9_]*$", description: "name to refer to the new object as $name" };

const shapeTarget = {
  type: "object",
  properties: {
    slide: slideRef,
    id: { type: ["string", "integer"] },
    name: { type: "string", minLength: 1 },
    placeholder: { type: "string", enum: ["title", "subTitle", "body"] },
    text: { type: "string", minLength: 1 }
  },
  required: ["slide"],
  // exactly one way of pointing at the shape
  oneOf: [{ required: ["id"] }, { required: ["name"] }, { required: ["placeholder"] }, { required: ["text"] }],
  additionalProperties: false
};
const target = {
  oneOf: [
    shapeTarget,
    { type: "string", pattern: "^(\\$[A-Za-z_][A-Za-z0-9_]*|@selection)$", description: '"$alias" or "@selection"' }
  ]
};

const chartKind = { type: "string", enum: ["column", "bar", "line", "pie"], description: "column = vertical bars, bar = horizontal bars" };
const chartCategories = { type: "array", minItems: 1, maxItems: 60, items: { type: "string" }, description: "one label per group on the axis (a pie: one per slice)" };
const chartSeries = {
  type: "array", minItems: 1, maxItems: 12,
  description: "one series per bar colour or line; a pie has exactly one. Each needs one value per category.",
  items: {
    type: "object",
    properties: { name: { type: "string" }, values: { type: "array", items: { type: ["number", "null"] } }, color },
    required: ["name", "values"], additionalProperties: false
  }
};
const legendPosition = { type: "string", enum: ["top", "bottom", "left", "right", "none"] };
const align = { type: "string", enum: ["left", "center", "right"] };
const dash = { type: "string", enum: ["solid", "dash", "dot"] };
const lineStyle = {
  type: "object",
  properties: { color: { oneOf: [color, { type: "null" }] }, widthPt: { type: "number", exclusiveMinimum: 0 }, dash },
  additionalProperties: false
};
const textStyle = {
  type: "object",
  properties: {
    fontSizePt: { type: "number", exclusiveMinimum: 0 },
    bold: { type: "boolean" },
    italic: { type: "boolean" },
    color,
    align
  },
  additionalProperties: false
};

const command = (op: string, properties: Record<string, unknown>, required: string[], extra: Record<string, unknown> = {}) => ({
  type: "object",
  properties: { op: { const: op }, ...properties },
  required: ["op", ...required],
  additionalProperties: false,
  ...extra
});

const commands = [
  command("addSlide", { after: { type: "integer", minimum: 0 }, layout: { type: "string", minLength: 1 } }, []),
  command("duplicateSlide", { slide: slideRef, after: { type: "integer", minimum: 0 } }, ["slide"]),
  command("deleteSlide", { slide: slideRef }, ["slide"]),
  command("moveSlide", { slide: slideRef, to: { type: "integer", minimum: 1 } }, ["slide", "to"]),

  command("addText", {
    slide: slideRef, rect,
    text: { type: "string" },
    paragraphs: {
      type: "array", minItems: 1,
      items: { type: "object", properties: { text: { type: "string" }, align }, required: ["text"], additionalProperties: false }
    },
    style: textStyle, as: alias
  }, ["slide", "rect"], { oneOf: [{ required: ["text"] }, { required: ["paragraphs"] }] }),
  command("addShape", {
    slide: slideRef,
    shape: { type: "string", enum: ["rect", "roundRect", "ellipse", "triangle", "rightArrow", "line", "arrow"] },
    rect, fill: { oneOf: [color, { type: "null" }] }, line: lineStyle,
    text: { type: "string" }, flipH: { type: "boolean" }, flipV: { type: "boolean" }, as: alias
  }, ["slide", "shape", "rect"]),
  command("addImage", {
    slide: slideRef,
    src: { type: "string", pattern: "^data:image/(png|jpeg|gif);base64,", description: "a data: URL of a PNG, JPEG or GIF" },
    rect, name: { type: "string" }, as: alias
  }, ["slide", "src"]),

  command("addChart", {
    slide: slideRef, type: chartKind, categories: chartCategories, series: chartSeries,
    title: { type: "string" }, legend: legendPosition, dataLabels: { type: "boolean" },
    stacked: { type: "boolean", description: "stack the series (not for pie)" },
    rect, name: { type: "string" }, as: alias
  }, ["slide", "type", "categories", "series"]),
  command("setChart", {
    target, type: chartKind, categories: chartCategories, series: chartSeries,
    title: { type: ["string", "null"], description: "null removes the title" }, legend: legendPosition,
    dataLabels: { type: "boolean" }, stacked: { type: "boolean" }
  }, ["target"], {
    anyOf: ["type", "categories", "series", "title", "legend", "dataLabels", "stacked"].map(field => ({ required: [field] }))
  }),

  command("setText", { target, text: { type: "string" } }, ["target", "text"]),
  command("replaceText", {
    target, find: { type: "string", minLength: 1 }, replace: { type: "string" }, all: { type: "boolean" }
  }, ["target", "find", "replace"]),
  command("formatText", {
    target,
    range: {
      type: "object",
      properties: { from: { type: "integer", minimum: 0 }, to: { type: "integer", minimum: 0 } },
      required: ["from", "to"], additionalProperties: false
    },
    bold: { type: "boolean" }, italic: { type: "boolean" }, color, fontSizePt: { type: "number", exclusiveMinimum: 0 }
  }, ["target"], { anyOf: [{ required: ["bold"] }, { required: ["italic"] }, { required: ["color"] }, { required: ["fontSizePt"] }] }),
  command("setParagraph", { target, align }, ["target", "align"]),

  command("setBounds", {
    target, x: cm, y: cm, w: { type: "number", exclusiveMinimum: 0 }, h: { type: "number", exclusiveMinimum: 0 }, keepRatio: { type: "boolean" }
  }, ["target"], { anyOf: [{ required: ["x"] }, { required: ["y"] }, { required: ["w"] }, { required: ["h"] }] }),
  command("setRotation", { target, degrees: { type: "number" } }, ["target", "degrees"]),
  command("setFill", { target, color: { oneOf: [color, { type: "null" }] } }, ["target", "color"]),
  command("setLine", { target, color: { oneOf: [color, { type: "null" }] }, widthPt: { type: "number", exclusiveMinimum: 0 }, dash }, ["target"],
    { anyOf: [{ required: ["color"] }, { required: ["widthPt"] }, { required: ["dash"] }] }),
  command("setAnchor", { target, anchor: { type: "string", enum: ["top", "middle", "bottom"] } }, ["target", "anchor"]),
  command("align", {
    targets: { type: "array", minItems: 1, items: target },
    to: { const: "slide" },
    edge: { type: "string", enum: ["left", "hcenter", "right", "top", "vcenter", "bottom"] }
  }, ["targets", "to", "edge"]),
  command("arrange", { target, mode: { type: "string", enum: ["front", "forward", "backward", "back"] } }, ["target", "mode"]),
  command("duplicate", { target, offset: { type: "number" }, as: alias }, ["target"]),
  command("copyTo", { target, slide: slideRef, as: alias }, ["target", "slide"]),
  command("delete", { target }, ["target"])
];

export const OPS: string[] = commands.map(entry => entry.properties.op.const);

/** The schema a model is given (and every batch is checked against). */
export const batchSchema = {
  $schema: "http://json-schema.org/draft-07/schema#",
  title: "PPT edit batch",
  type: "object",
  properties: {
    label: { type: "string", description: "name of this step in the undo history" },
    atomic: { type: "boolean", description: "true (default): if any command fails nothing is changed" },
    commands: { type: "array", minItems: 1, items: { oneOf: commands } }
  },
  required: ["commands"],
  additionalProperties: false
};

const ajv = new Ajv({ allErrors: false, strict: false, discriminator: true, verbose: true });
const validateBatchSchema = ajv.compile({
  ...batchSchema,
  properties: {
    ...batchSchema.properties,
    commands: { type: "array", minItems: 1, items: { discriminator: { propertyName: "op" }, oneOf: commands } }
  }
});

export interface CommandBatch {
  label?: string;
  atomic: boolean;
  commands: Array<Record<string, unknown> & { op: string }>;
}

const dotted = (instancePath: string) =>
  instancePath.replace(/^\//, "").split("/").filter(Boolean)
    .map(part => (/^\d+$/.test(part) ? `[${part}]` : `.${part}`)).join("").replace(/^\./, "");

function describe(error: ErrorObject, value: unknown): string {
  const params = error.params as Record<string, unknown>;
  switch (error.keyword) {
    case "required": return `缺少字段 "${params.missingProperty}"`;
    case "additionalProperties": return `不认识的字段 "${params.additionalProperty}"`;
    case "enum": return `取值必须是 ${JSON.stringify(params.allowedValues)} 之一`;
    case "const": return `取值必须是 ${JSON.stringify(params.allowedValue)}`;
    case "type": return `类型必须是 ${params.type}`;
    case "pattern": return `格式不对，应为 ${params.pattern}`;
    case "minimum": case "exclusiveMinimum": return `不能小于${error.keyword === "exclusiveMinimum" ? "或等于" : ""} ${params.limit}`;
    case "oneOf": return "必须且只能写出其中一种指向方式（id、name、placeholder、text 之一）";
    case "anyOf": return "至少要写出一个要修改的字段";
    case "discriminator":
      return params.error === "mapping" || params.error === "tag"
        ? `未知的 op，可用的有：${OPS.join("、")}`
        : error.message || "指令不符合格式";
    default: return error.message || "格式不对";
  }
  void value;
}

/** Accepts a batch object or a bare array of commands; returns the normalised batch. */
export function parseBatch(input: unknown): CommandBatch {
  const raw = Array.isArray(input) ? { commands: input } : input;
  if (!validateBatchSchema(raw)) {
    const errors = validateBatchSchema.errors || [];
    // "oneOf" / "anyOf" failures wrap the failures of every branch; the wrapper
    // says it better ("pick one way of pointing at the shape"). Otherwise the
    // deepest problem is the most specific one.
    // Only "pick which of these fields" choices qualify, not unions of value types.
    const isFieldChoice = (candidate: ErrorObject) =>
      (candidate.keyword === "oneOf" || candidate.keyword === "anyOf")
      && Array.isArray(candidate.schema)
      && (candidate.schema as Array<Record<string, unknown>>).every(item => "required" in item && !("type" in item));
    const wrapper = errors.find(isFieldChoice);
    const error = wrapper
      || [...errors].sort((a, b) => b.instancePath.length - a.instancePath.length)[0];
    throw new ApiError("INVALID_SCHEMA", describe(error, raw), dotted(error.instancePath) || "(root)");
  }
  const batch = raw as { label?: string; atomic?: boolean; commands: CommandBatch["commands"] };
  return { label: batch.label, atomic: batch.atomic !== false, commands: batch.commands };
}
