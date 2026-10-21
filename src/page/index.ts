/**
 * What a web page needs to let an Agent work on the deck the user has open:
 * the page's tools, the loop that answers the host's calls, autosave, and the
 * open-source font policy.
 */
export { PAGE_TOOLS, MUTATING_TOOLS, DESIGN_TOOLS, DESIGN_TOOL_GUIDE, runPageTool, viewerToolHost } from "./page-tools";
export type { PageToolDefinition, PageToolOptions } from "./page-tools";
export { PROJECT_ASSET_PREFIX, resolveProjectAssets } from "./assets";
export type { AssetResolver } from "./assets";
export { PageToolLink } from "./tool-link";
export type { PageToolCall, ToolCallTransport, PageToolLinkOptions } from "./tool-link";
export { AutoSaver, RevisionConflict } from "./autosave";
export type { DeckStore, SaveSource, SaveStatus, SavableDeck, AutoSaverOptions } from "./autosave";
export { OPEN_FONT_MAP, OPEN_FALLBACKS, resolveOpenFamily } from "./font-map";
export { WorkerRenderer } from "./worker-renderer";

export { DesignBridge } from "./design-bridge";
export type { DesignBridgeOptions } from "./design-bridge";
export type { TemplateSemanticProposal, TemplateSemanticChange, TemplateSemanticEdit } from "../design/template-semantics";
