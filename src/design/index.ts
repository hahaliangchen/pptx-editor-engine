/** Node conversion API. JSON is transient; callers need persist only the returned PPTX. */
export { importDesign as pptxToJson } from "./import";
export { projectPage as jsonToHtml } from "./projection";
export { mergeHtml as htmlToJson } from "./measure";
export { exportDesign as jsonToPptx } from "./export";
export { syncDesign, materializeDesign, convertDesign } from "./service";
export type { DesignRequest } from "./service";
export type { DesignDocument, DesignPage, DesignNode } from "./model";
