import { ApiError, ShapeCandidate } from "./errors";
import { SceneNode, SlideScene, allNodes, textOf } from "./scene";

export type SlideRef = number | string;

export interface ShapeTarget {
  slide: SlideRef;
  id?: string;
  name?: string;
  placeholder?: string;
  text?: string;
}

/** A shape reference as it appears in a command. */
export type Target = ShapeTarget | string;

export interface ResolveContext {
  scenes: SlideScene[];
  /** "$alias" -> the shape that an earlier command created. */
  aliases: Map<string, { part: string; id: string }>;
  /** The user's current selection, for "@selection". */
  selection: () => Array<{ part: string; id: string }>;
}

const snippet = (node: SceneNode) => {
  const text = textOf(node).replace(/\s+/g, " ").trim();
  return text.length > 30 ? `${text.slice(0, 30)}…` : text;
};

const toCandidate = (node: SceneNode): ShapeCandidate => ({ id: node.id, name: node.name, text: snippet(node) });

export function resolveSlide(scenes: SlideScene[], ref: SlideRef, path: string): SlideScene {
  const asNumber = typeof ref === "number" ? ref : /^\d+$/.test(ref) ? parseInt(ref, 10) : null;
  const scene = asNumber !== null
    ? scenes.find(candidate => candidate.index === asNumber)
    : scenes.find(candidate => candidate.slide.id === ref);
  if (!scene) {
    throw new ApiError("SLIDE_NOT_FOUND", `没有第 ${ref} 页，当前共 ${scenes.length} 页`, path);
  }
  return scene;
}

const placeholderKind = (value: string) => (value === "ctrTitle" ? "title" : value === "obj" ? "body" : value);

function findByPart(scenes: SlideScene[], part: string, id: string): { scene: SlideScene; node: SceneNode } | null {
  for (const scene of scenes) {
    const node = allNodes(scene).find(candidate => candidate.part === part && candidate.id === id);
    if (node) return { scene, node };
  }
  return null;
}

/** Find the one shape a target means; fail with candidates rather than guess. */
export function resolveNode(context: ResolveContext, target: Target, path: string): { scene: SlideScene; node: SceneNode } {
  if (typeof target === "string") {
    if (target === "@selection") {
      const [first, ...rest] = context.selection();
      if (!first) throw new ApiError("SHAPE_NOT_FOUND", "界面里当前没有选中任何对象", path);
      if (rest.length > 0) throw new ApiError("SHAPE_AMBIGUOUS", "界面里选中了多个对象，请只选一个，或改用 id 指定", path);
      const found = findByPart(context.scenes, first.part, first.id);
      if (!found) throw new ApiError("SHAPE_NOT_FOUND", "选中的对象已经不存在", path);
      return found;
    }
    if (target.startsWith("$")) {
      const alias = context.aliases.get(target.slice(1));
      if (!alias) {
        throw new ApiError("SHAPE_NOT_FOUND", `没有别名 ${target}；别名要先在前面的指令里用 "as" 起好`, path);
      }
      const found = findByPart(context.scenes, alias.part, alias.id);
      if (!found) throw new ApiError("SHAPE_NOT_FOUND", `别名 ${target} 指向的对象已经不存在`, path);
      return found;
    }
    throw new ApiError("INVALID_SCHEMA", `无法理解的目标 "${target}"：应为 "$别名"、"@selection" 或对象 {slide, id|name|placeholder|text}`, path);
  }

  const scene = resolveSlide(context.scenes, target.slide, `${path}.slide`);
  const nodes = allNodes(scene);
  const here = `第 ${scene.index} 页`;

  if (target.id !== undefined) {
    const node = nodes.find(candidate => candidate.id === String(target.id));
    if (!node) throw new ApiError("SHAPE_NOT_FOUND", `${here}没有 id 为 ${target.id} 的形状`, `${path}.id`, nodes.slice(0, 12).map(toCandidate));
    return { scene, node };
  }

  let matches: SceneNode[];
  let what: string;
  let field: string;
  if (target.name !== undefined) {
    matches = nodes.filter(candidate => candidate.name === target.name);
    what = `名为 "${target.name}"`;
    field = "name";
  } else if (target.placeholder !== undefined) {
    matches = nodes.filter(candidate => candidate.placeholder === placeholderKind(target.placeholder!));
    what = `占位符类型为 ${target.placeholder}`;
    field = "placeholder";
  } else if (target.text !== undefined) {
    // Members of a group would match along with the group itself; prefer the members.
    const containing = nodes.filter(candidate => textOf(candidate).includes(target.text!));
    matches = containing.filter(candidate => candidate.kind !== "group" || !containing.some(other => other.parent === candidate));
    what = `文字包含 "${target.text}"`;
    field = "text";
  } else {
    throw new ApiError("INVALID_SCHEMA", "目标需要 id、name、placeholder 或 text 之一", path);
  }

  if (matches.length === 0) {
    throw new ApiError("SHAPE_NOT_FOUND", `${here}没有${what}的形状`, `${path}.${field}`, nodes.slice(0, 12).map(toCandidate));
  }
  if (matches.length > 1) {
    throw new ApiError("SHAPE_AMBIGUOUS", `${here}有 ${matches.length} 个${what}的形状，请改用 id 指定`, `${path}.${field}`, matches.map(toCandidate));
  }
  return { scene, node: matches[0] };
}
