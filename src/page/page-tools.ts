import { TOOLS, ToolDefinition, ToolHost, ToolResult, callTool } from "../api/tools";
import type PptViewer from "../index";
import { AssetResolver, pageBatchSchema, resolveProjectAssets } from "./assets";

/**
 * A tool the plugin declares with `executor: "page"`: the Agent host hands
 * calls to the page that started the task instead of the plugin's service.
 */
export interface PageToolDefinition extends ToolDefinition {
  executor: "page";
}

/** Every deck tool runs in the page, on the deck the user is looking at. */
export const PAGE_TOOLS: PageToolDefinition[] = TOOLS.map(tool => ({
  ...tool,
  ...(tool.name === "apply_edits" ? { inputSchema: pageBatchSchema() } : {}),
  executor: "page"
}));

/** Tools that change the deck; a successful call is saved before it is answered. */
export const MUTATING_TOOLS = new Set(["apply_edits", "build_deck"]);

export interface PageToolOptions {
  /** Fetches a project picture named by `addImage.src: "project:images/..."`; without it such a command fails. */
  resolveAsset?: AssetResolver;
}

/** The viewer as `callTool` sees it: the open deck, its edit API and the worker renderer. */
export function viewerToolHost(viewer: PptViewer, options: PageToolOptions = {}): ToolHost {
  return {
    outline: outlineOptions => viewer.api.outline(outlineOptions),
    apply: async input => {
      const resolved = await resolveProjectAssets(input, options.resolveAsset);
      return "failure" in resolved ? resolved.failure : viewer.api.apply(resolved.batch);
    },
    render: (page, renderOptions) => viewer.renderPagePng(page, renderOptions?.width)
  };
}

/** Run one page tool against the deck open in `viewer`. Never throws. */
export function runPageTool(viewer: PptViewer, name: string, args: unknown, options?: PageToolOptions): Promise<ToolResult> {
  return callTool(viewerToolHost(viewer, options), name, args);
}
