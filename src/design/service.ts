import { importDesign } from "./import";
import { exportDesign } from "./export";
import { mergeHtml } from "./measure";
import { installNodeEnv } from "../headless/node-env";
import { HeadlessSession } from "../headless/session";
import { createDesignConverter } from "./service-core";
export type { DesignRequest } from "./service-core";
const converter = createDesignConverter({ importDesign, exportDesign, mergeHtml: (design, html, baseline, authorization) => mergeHtml(design, html, undefined, baseline, authorization),
  async pages(bytes, commands) {
    const session = await HeadlessSession.open(bytes, { quiet: true });
    try {
      const result = await session.apply({ commands });
      if (!result.ok) throw new Error(JSON.stringify(result));
      if (!session.document.slides.length) throw new Error("Keep at least one page");
      return await session.export();
    } finally { session.dispose(); }
  }
});
export const syncDesign = converter.syncDesign;
export const materializeDesign = converter.materializeDesign;
export const convertDesign: typeof converter.convertDesign = request => {
  installNodeEnv();
  return converter.convertDesign(request);
};
