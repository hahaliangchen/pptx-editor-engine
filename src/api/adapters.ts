import { sharedSchemas } from "./schema";
import { TOOLS, ToolContent, ToolDefinition, ToolHost, ToolResult, callTool } from "./tools";

/**
 * The same tools in the shapes each platform expects. The differences are only
 * in names and wrapping: all three carry a name, a description and a JSON Schema.
 */

/**
 * Replace each occurrence of a shared fragment by a `$ref` to one definition.
 * The result accepts and rejects exactly what the original does, in a fraction
 * of the size.
 */
export function shareDefinitions(schema: Record<string, unknown>, shared: Record<string, unknown>): Record<string, unknown> {
  const key = (value: unknown) => JSON.stringify(value);
  const names = new Map(Object.entries(shared).map(([name, value]) => [key(value), name]));
  const defs: Record<string, unknown> = {};

  const visit = (node: unknown): unknown => {
    if (Array.isArray(node)) return node.map(visit);
    if (!node || typeof node !== "object") return node;
    const name = names.get(key(node));
    if (name) {
      // The definition's own content may use other definitions, but never itself.
      if (!(name in defs)) defs[name] = children(shared[name]);
      return { $ref: `#/$defs/${name}` };
    }
    return children(node);
  };
  const children = (node: unknown): unknown =>
    Object.fromEntries(Object.entries(node as Record<string, unknown>).map(([field, value]) => [field, visit(value)]));

  const body = children(schema) as Record<string, unknown>;
  return Object.keys(defs).length > 0 ? { ...body, $defs: defs } : body;
}

/** A schema safe to hand to any provider: a private copy, without the `$schema` marker some of them reject. */
function portable(schema: Record<string, unknown>): Record<string, unknown> {
  const copy = shareDefinitions(JSON.parse(JSON.stringify(schema)) as Record<string, unknown>, sharedSchemas);
  delete copy.$schema;
  return copy;
}

// ---------------------------------------------------------------- Anthropic

export interface AnthropicTool {
  name: string;
  description: string;
  input_schema: Record<string, unknown>;
}

export const toAnthropic = (tools: ToolDefinition[] = TOOLS): AnthropicTool[] =>
  tools.map(tool => ({ name: tool.name, description: tool.description, input_schema: portable(tool.inputSchema) }));

export type AnthropicToolResultBlock = {
  type: "tool_result";
  tool_use_id: string;
  is_error?: boolean;
  content: Array<{ type: "text"; text: string } | { type: "image"; source: { type: "base64"; media_type: string; data: string } }>;
};

export const toAnthropicResult = (result: ToolResult, toolUseId: string): AnthropicToolResultBlock => ({
  type: "tool_result",
  tool_use_id: toolUseId,
  ...(result.ok ? {} : { is_error: true }),
  content: result.content.map(part => part.type === "text"
    ? { type: "text" as const, text: part.text }
    : { type: "image" as const, source: { type: "base64" as const, media_type: part.mimeType, data: part.data } })
});

/** Run a `tool_use` block from a response and give back the `tool_result` block for the next request. */
export async function callAnthropicTool(
  host: ToolHost,
  block: { id: string; name: string; input: unknown }
): Promise<AnthropicToolResultBlock> {
  return toAnthropicResult(await callTool(host, block.name, block.input), block.id);
}

// ------------------------------------------------------------------- OpenAI

export interface OpenAITool {
  type: "function";
  function: { name: string; description: string; parameters: Record<string, unknown> };
}

/** `strict` is not offered: the edit schema uses `oneOf`, which OpenAI's strict mode does not accept. */
export const toOpenAI = (tools: ToolDefinition[] = TOOLS): OpenAITool[] =>
  tools.map(tool => ({
    type: "function",
    function: { name: tool.name, description: tool.description, parameters: portable(tool.inputSchema) }
  }));

export type OpenAIMessage =
  | { role: "tool"; tool_call_id: string; content: string }
  | { role: "user"; content: Array<{ type: "text"; text: string } | { type: "image_url"; image_url: { url: string } }> };

/**
 * A tool message is text only, so an image comes in a user message right after
 * it. Append the returned messages to the conversation in order.
 */
export function toOpenAIMessages(result: ToolResult, toolCallId: string): OpenAIMessage[] {
  const texts = result.content.filter((part): part is Extract<ToolContent, { type: "text" }> => part.type === "text");
  const images = result.content.filter((part): part is Extract<ToolContent, { type: "image" }> => part.type === "image");
  const messages: OpenAIMessage[] = [{
    role: "tool",
    tool_call_id: toolCallId,
    content: texts.map(part => part.text).join("\n") || (images.length > 0 ? "图片见下一条消息" : "")
  }];
  if (images.length > 0) {
    messages.push({
      role: "user",
      content: [
        { type: "text", text: "上一个工具返回的图片：" },
        ...images.map(image => ({ type: "image_url" as const, image_url: { url: `data:${image.mimeType};base64,${image.data}` } }))
      ]
    });
  }
  return messages;
}

/** Run a `tool_calls` entry (its arguments are a JSON string) and give back the messages to append. */
export async function callOpenAITool(
  host: ToolHost,
  call: { id: string; function: { name: string; arguments: string } }
): Promise<OpenAIMessage[]> {
  let args: unknown;
  try {
    args = call.function.arguments ? JSON.parse(call.function.arguments) : {};
  } catch (error) {
    const failed: ToolResult = {
      ok: false,
      content: [{ type: "text", text: JSON.stringify({ ok: false, error: { message: `参数不是合法的 JSON：${error instanceof Error ? error.message : error}` } }) }]
    };
    return toOpenAIMessages(failed, call.id);
  }
  return toOpenAIMessages(await callTool(host, call.function.name, args), call.id);
}

// --------------------------------------------------------------------- MCP

export interface McpTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
}

export const toMcp = (tools: ToolDefinition[] = TOOLS): McpTool[] =>
  tools.map(tool => ({ name: tool.name, description: tool.description, inputSchema: portable(tool.inputSchema) }));

export interface McpToolResult {
  content: Array<{ type: "text"; text: string } | { type: "image"; data: string; mimeType: string }>;
  isError?: boolean;
}

export const toMcpResult = (result: ToolResult): McpToolResult => ({
  content: result.content.map(part => part.type === "text"
    ? { type: "text" as const, text: part.text }
    : { type: "image" as const, data: part.data, mimeType: part.mimeType }),
  ...(result.ok ? {} : { isError: true })
});

/** The body of an MCP server's `tools/call` handler. */
export async function callMcpTool(host: ToolHost, name: string, args: unknown): Promise<McpToolResult> {
  return toMcpResult(await callTool(host, name, args));
}
