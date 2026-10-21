import { installNodeEnv } from "../headless/node-env";
import { exportDesignBytes } from "./export-core";
import type { DesignDocument } from "./model";
export async function exportDesign(design: DesignDocument, identities?: Map<string, string>): Promise<Buffer> {
  installNodeEnv();
  return Buffer.from(await exportDesignBytes(design, identities));
}
