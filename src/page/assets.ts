import { ApiError } from "../api/errors";
import type { ApplyResult } from "../api/engine";
import { batchSchema } from "../api/schema";

/**
 * `addImage.src` in a page tool call may name a picture in the plugin's
 * project (`project:images/x.png`, e.g. an attachment the host imported)
 * instead of carrying megabytes of base64. The page fetches it and puts the
 * data URL in place before the batch is checked and applied.
 */
export const PROJECT_ASSET_PREFIX = "project:";

/** Fetches a project file (`images/x.png`) and returns it as a PNG, JPEG or GIF data URL. */
export type AssetResolver = (path: string) => Promise<string>;

/** The batch schema the page declares: like the engine's, but `addImage.src` may also be a project reference. */
export function pageBatchSchema(): Record<string, unknown> {
  const schema = JSON.parse(JSON.stringify(batchSchema));
  for (const command of schema.properties.commands.items.oneOf) {
    if (command.properties.op.const !== "addImage") continue;
    command.properties.src = {
      type: "string",
      pattern: "^(data:image/(png|jpeg|gif);base64,|project:images/|placeholder:)",
      description: "a data: URL of a PNG, JPEG or GIF, or a picture in the project: \"project:\" + its path (e.g. the remote_path of an attachment, project:images/....png)"
    };
  }
  return schema;
}

function failed(index: number, error: ApiError): ApplyResult {
  return { ok: false, failedAt: index, error: error.toJSON(), results: [] };
}

/**
 * Replace every project reference in a batch (an object with `commands`, or a
 * bare array) with the data URL `resolve` returns. Each path is fetched once.
 * Returns the batch to apply, or a failed result naming the command whose
 * picture could not be read. The input is not changed.
 */
export async function resolveProjectAssets(input: unknown, resolve?: AssetResolver): Promise<{ batch: unknown } | { failure: ApplyResult }> {
  const commands = Array.isArray(input) ? input : (input as { commands?: unknown } | null)?.commands;
  if (!Array.isArray(commands)) return { batch: input };
  const isReference = (command: unknown): command is { op: "addImage"; src: string } =>
    !!command && typeof command === "object"
    && (command as { op?: unknown }).op === "addImage"
    && typeof (command as { src?: unknown }).src === "string"
    && (command as { src: string }).src.startsWith(PROJECT_ASSET_PREFIX);
  if (!commands.some(isReference)) return { batch: input };

  const fetched = new Map<string, Promise<string>>();
  const resolved = [...commands];
  for (let index = 0; index < commands.length; index++) {
    const command = commands[index];
    if (!isReference(command)) continue;
    const at = `commands[${index}].src`;
    const path = command.src.slice(PROJECT_ASSET_PREFIX.length);
    if (!resolve) return { failure: failed(index, new ApiError("UNSUPPORTED", "这个页面不能读取项目里的图片，请改用 data: URL", at)) };
    if (!path.startsWith("images/") || path.split("/").some(part => part === ".." || part === "")) {
      return { failure: failed(index, new ApiError("INVALID_SCHEMA", `项目图片的路径必须在 images/ 下：${path}`, at)) };
    }
    if (!fetched.has(path)) fetched.set(path, resolve(path));
    let src: string;
    try {
      src = await fetched.get(path)!;
    } catch (error) {
      return { failure: failed(index, new ApiError("INVALID_SCHEMA", `读取项目图片 ${path} 失败：${error instanceof Error ? error.message : error}`, at)) };
    }
    resolved[index] = { ...command, src };
  }
  return { batch: Array.isArray(input) ? resolved : { ...(input as object), commands: resolved } };
}
