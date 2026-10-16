import { ElementSource, GroupRef, Rect, ShapeElement, Slide, SlideElement, TextElement } from "../pptx-virtual-dom";

export type NodeKind = "text" | "shape" | "image" | "line" | "table" | "chart" | "group";

/**
 * One thing a person would pick on a slide: a shape, a picture, a table or a
 * whole group. Groups hold their members, so the tree matches PowerPoint's
 * selection pane.
 */
export interface SceneNode {
  id: string;
  name: string;
  kind: NodeKind;
  part: string;
  /** Slide logical px, as the elements carry them. */
  rect: Rect;
  rotation: number;
  /** The elements that draw this node (all descendants for a group). */
  elements: SlideElement[];
  placeholder?: string;
  children: SceneNode[];
  parent: SceneNode | null;
}

export interface SlideScene {
  index: number;
  slide: Slide;
  roots: SceneNode[];
}

const LINES = new Set(["line", "straightConnector1", "lineInv", "bentConnector2", "bentConnector3", "curvedConnector3"]);

const unionRect = (rects: Rect[]): Rect => {
  const left = Math.min(...rects.map(r => r.x));
  const top = Math.min(...rects.map(r => r.y));
  const right = Math.max(...rects.map(r => r.x + r.w));
  const bottom = Math.max(...rects.map(r => r.y + r.h));
  return { x: left, y: top, w: right - left, h: bottom - top };
};

function leafKind(elements: SlideElement[]): NodeKind {
  if (elements.some(element => element.type === "image")) return "image";
  if (elements.some(element => element.type === "chart")) return "chart";
  if (elements.some(element => element.source?.readOnlyText)) return "table";
  const shape = elements.find((element): element is ShapeElement => element.type === "shape");
  if (shape && LINES.has(shape.shapeType)) return "line";
  const painted = shape && (shape.fill !== "transparent" || !!shape.border);
  return painted ? "shape" : "text";
}

const placeholderOf = (elements: SlideElement[]): string | undefined => {
  const text = elements.find((element): element is TextElement => element.type === "text" && !!element.placeholder);
  const type = text?.placeholder?.type;
  if (!type) return undefined;
  if (type === "ctrTitle") return "title";
  if (type === "obj") return "body";
  return type;
};

/** The selection-pane tree of one slide. Master and layout decoration is not part of it. */
export function buildSlideScene(slide: Slide, index: number): SlideScene {
  const roots: SceneNode[] = [];
  const groups = new Map<string, SceneNode>();
  const leaves = new Map<string, SceneNode>();

  const attach = (node: SceneNode, parent: SceneNode | null) => {
    node.parent = parent;
    (parent ? parent.children : roots).push(node);
  };

  for (const element of slide.elements) {
    const source: ElementSource | undefined = element.source;
    if (!source) continue;

    // Make sure every group on the path exists, outermost first.
    let parent: SceneNode | null = null;
    let path = "";
    for (const ref of source.groups || ([] as GroupRef[])) {
      path += `/${ref.shapeId}`;
      let group = groups.get(path);
      if (!group) {
        group = {
          id: ref.shapeId,
          name: ref.name || `组合 ${ref.shapeId}`,
          kind: "group",
          part: source.part,
          rect: { x: 0, y: 0, w: 0, h: 0 },
          rotation: 0,
          elements: [],
          children: [],
          parent: null
        };
        groups.set(path, group);
        attach(group, parent);
      }
      group.elements.push(element);
      parent = group;
    }

    const key = `${path}|${source.shapeId}`;
    let leaf = leaves.get(key);
    if (!leaf) {
      leaf = {
        id: source.shapeId,
        name: source.name || `形状 ${source.shapeId}`,
        kind: "text",
        part: source.part,
        rect: { ...element.rect },
        rotation: 0,
        elements: [],
        children: [],
        parent: null
      };
      leaves.set(key, leaf);
      attach(leaf, parent);
    }
    leaf.elements.push(element);
  }

  for (const leaf of leaves.values()) {
    leaf.kind = leafKind(leaf.elements);
    leaf.rect = unionRect(leaf.elements.map(element => element.rect));
    leaf.rotation = leaf.elements.reduce((acc, element) => acc || element.rotation || 0, 0);
    leaf.placeholder = placeholderOf(leaf.elements);
  }
  for (const group of groups.values()) group.rect = unionRect(group.elements.map(element => element.rect));

  return { index, slide, roots };
}

export const buildScene = (slides: Slide[]): SlideScene[] =>
  slides.map((slide, index) => buildSlideScene(slide, index + 1));

/** Every node of a slide, depth first. */
export function allNodes(scene: SlideScene): SceneNode[] {
  const out: SceneNode[] = [];
  const visit = (node: SceneNode) => {
    out.push(node);
    node.children.forEach(visit);
  };
  scene.roots.forEach(visit);
  return out;
}

export const textOf = (node: SceneNode): string =>
  node.elements
    .filter((element): element is TextElement => element.type === "text" && !element.source?.readOnlyText)
    .map(element => element.content)
    .join("\n");
