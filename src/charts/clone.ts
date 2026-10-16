import JSZip from "jszip";
import { ChartAdd, ChartModel } from "./model";
import { CHART_CONTENT_TYPE } from "./write";
import { StructureEdits, resolveTarget } from "../pptx-exporter";

export interface ChartCloneContext {
  zip: JSZip | null;
  structure: StructureEdits;
  /** Free part paths for a chart made from scratch. */
  allocateChartParts(): { part: string; workbook: string };
}

const REL_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const dirName = (path: string) => path.slice(0, path.lastIndexOf("/"));
const baseName = (path: string) => path.slice(path.lastIndexOf("/") + 1);

/** "ppt/charts" -> "ppt/embeddings/x.xlsx" written the way a relationship in the first folder names it. */
function relativePath(fromDir: string, to: string): string {
  const from = fromDir.split("/");
  const target = to.split("/");
  while (from.length > 0 && target.length > 1 && from[0] === target[0]) {
    from.shift();
    target.shift();
  }
  return [...from.map(() => ".."), ...target].join("/");
}

const encode = (text: string) => new TextEncoder().encode(text);

/**
 * Give `model` (a copy of another chart's model) a chart part of its own, so two
 * frames never show one part. A chart made in this session just gets new part
 * names; a chart of the opened deck has all its parts copied: the chart, its
 * workbook, its style and colour parts.
 */
export async function cloneChartParts(
  ctx: ChartCloneContext,
  model: ChartModel
): Promise<{ part: string; add?: ChartAdd }> {
  if (!model.part) {
    const { part, workbook } = ctx.allocateChartParts();
    return { part, add: { part, workbook, model } };
  }

  const { zip, structure } = ctx;
  const source = model.part;
  const added = new Map(structure.media.map(item => [item.path, item.bytes]));
  const read = async (path: string): Promise<Uint8Array | null> => {
    const file = zip?.file(path);
    return file ? file.async("uint8array") : added.get(path) ?? null;
  };

  const taken = new Set<string>([
    ...Object.keys(zip?.files || {}),
    ...added.keys(),
    ...structure.charts.flatMap(add => [add.part, add.workbook])
  ]);
  const fresh = (path: string): string => {
    const dir = dirName(path);
    const file = baseName(path);
    const stem = file.replace(/\d*(\.[^.]+)$/, "");
    const extension = file.slice(file.lastIndexOf("."));
    for (let n = 1; ; n++) {
      const candidate = `${dir}/${stem.replace(/\.[^.]+$/, "")}${n}${extension}`;
      if (!taken.has(candidate)) {
        taken.add(candidate);
        return candidate;
      }
    }
  };

  // The file types of the parts being copied, as the package declares them.
  const typesXml = zip?.file("[Content_Types].xml") ? await zip.file("[Content_Types].xml")!.async("string") : "";
  const types = new DOMParser().parseFromString(typesXml || "<Types/>", "application/xml");
  const typeOf = (path: string): { type: string; override: boolean } | null => {
    // A part copied earlier in this session is only in the pending list, with the type it was given.
    const pending = structure.media.find(item => item.path === path);
    if (pending) return { type: pending.partType ?? pending.contentType, override: !!pending.partType };
    for (const node of Array.from(types.getElementsByTagName("Override"))) {
      if (node.getAttribute("PartName") === `/${path}`) return { type: node.getAttribute("ContentType") || "", override: true };
    }
    const extension = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
    for (const node of Array.from(types.getElementsByTagName("Default"))) {
      if ((node.getAttribute("Extension") || "").toLowerCase() === extension) return { type: node.getAttribute("ContentType") || "", override: false };
    }
    return null;
  };

  const chartBytes = await read(source);
  if (!chartBytes) throw new Error(`找不到图表部件 ${source}`);
  const part = fresh(source);
  structure.media.push({ path: part, bytes: chartBytes, extension: "xml", contentType: CHART_CONTENT_TYPE, partType: CHART_CONTENT_TYPE });

  // Parts the chart relates to (workbook, style, colours) are copied one by one.
  const relsPath = `${dirName(source)}/_rels/${baseName(source)}.rels`;
  const relsBytes = await read(relsPath);
  if (relsBytes) {
    const rels = new DOMParser().parseFromString(new TextDecoder().decode(relsBytes), "application/xml");
    for (const rel of Array.from(rels.getElementsByTagName("Relationship"))) {
      if (rel.getAttribute("TargetMode") === "External") continue;
      const target = resolveTarget(dirName(source), rel.getAttribute("Target") || "");
      const bytes = await read(target);
      if (!bytes) continue;
      const copy = fresh(target);
      const declared = typeOf(target);
      const extension = copy.slice(copy.lastIndexOf(".") + 1);
      structure.media.push({
        path: copy,
        bytes,
        extension,
        contentType: declared?.type || "application/octet-stream",
        ...(declared?.override ? { partType: declared.type } : {})
      });
      rel.setAttribute("Target", relativePath(dirName(part), copy));
    }
    structure.media.push({
      path: `${dirName(part)}/_rels/${baseName(part)}.rels`,
      bytes: encode(new XMLSerializer().serializeToString(rels)),
      extension: "rels",
      contentType: REL_CONTENT_TYPE
    });
  }

  // Export writes the copy from its model, so it shows this chart's data and settings.
  model.part = part;
  structure.chartEdits.set(part, model);
  return { part };
}
