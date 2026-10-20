import type { CmRect, OutlineShape } from "./outline";
import { capacityOf, parseBinding } from "./template";

/**
 * Repeating units: the designer builds a layout with, say, three cards, and names the shapes of each card
 * `@内容.卡片1.标题`, `@内容.卡片1.说明`, `@内容.卡片1` (a background or icon that belongs to the card),
 * then `卡片2…`, `卡片3…`. The deck can then ask for any number of cards: the engine clones the unit
 * and spaces the instances as the template does, or scales them down to make room.
 */

/** A picture slot: a picture, or a shape filled with one, that a deck replaces with its own picture. */
export interface ImageSlot {
  id: string;
  rect: CmRect;
  /** The slot's outline: a round or rounded slot gets a round or rounded picture. */
  shape?: "rect" | "ellipse" | "roundRect";
  /** Width ÷ height. */
  aspect: number;
  /** How many shapes of the page sit above it, so the new picture can be put at the same depth (top-level shapes only). */
  above?: number;
}

/** Whether a shape is a picture, or a shape filled with one: something a picture slot can be. */
export const isImageShape = (shape: OutlineShape) =>
  shape.type === "image" || (typeof shape.fill === "object" && shape.fill !== null && "other" in shape.fill && shape.fill.other === "picture");

export function imageSlotOf(shape: OutlineShape, above?: number): ImageSlot {
  const preset = shape.shape;
  return {
    id: shape.id, rect: shape.rect, aspect: shape.rect.w / Math.max(0.01, shape.rect.h),
    ...(preset === "ellipse" ? { shape: "ellipse" as const } : preset === "roundRect" ? { shape: "roundRect" as const } : {}),
    ...(above !== undefined ? { above } : {})
  };
}

export interface UnitMember {
  id: string;
  /** A group node moves and clones as one piece; a leaf is a single shape. */
  kind: "group" | "leaf";
  rect: CmRect;
  /** Set when the member is a picture slot (its field gives the picture). */
  image?: ImageSlot;
  /** Set for a field: its field name, sample text and type size. */
  field?: string;
  text?: string;
  fontSizePt?: number;
  align?: string;
}

export interface UnitInstance {
  index: number;
  members: UnitMember[];
  /** All the members together, in cm. */
  rect: CmRect;
}

export interface UnitGroup {
  name: string;
  instances: UnitInstance[];
  /** Fields a content item gives, from the first instance, in template order. */
  fields: string[];
  /** Fields the engine fills itself with the instance number (`序号`). */
  auto: string[];
  /** How many instances a deck may ask for. */
  min: number;
  max: number;
  /** The template's number is the only one the layout works with. */
  fixed: boolean;
  /** The step from one instance to the next, when they stand evenly in a row or a column. */
  pitch?: { dx: number; dy: number };
  /** What is wrong with the way it is built; the template check reports these. */
  problems: string[];
}

const AUTO_FIELDS = new Set(["序号", "编号", "index", "number", "no"]);
export const isAutoUnitField = (field: string) => AUTO_FIELDS.has(field.toLowerCase());

/** The smallest an instance may be scaled to when more of them are asked for than the template has room for. */
export const MIN_UNIT_SCALE = 0.6;
/** The least gap, in cm, kept between instances when they are squeezed together. */
export const MIN_UNIT_GAP = 0.4;

const union = (rects: CmRect[]): CmRect => {
  const x = Math.min(...rects.map(r => r.x));
  const y = Math.min(...rects.map(r => r.y));
  return { x, y, w: Math.max(...rects.map(r => r.x + r.w)) - x, h: Math.max(...rects.map(r => r.y + r.h)) - y };
};

const near = (a: number, b: number, tolerance: number) => Math.abs(a - b) <= tolerance;
/** How far apart, in cm, the steps between instances (and their sizes) may be and still count as even: icons and text boxes never line up exactly. */
const STEP_TOLERANCE = 0.7;
const SIZE_TOLERANCE = 0.9;

/** Finds the repeating units on a page: the names of the top-level shapes (a group node counts as one piece). */
export function findUnits(shapes: OutlineShape[]): Record<string, UnitGroup> {
  const found: Array<{ group: string; index: number; field?: string; options?: Record<string, string>; node: OutlineShape }> = [];
  const problems = new Map<string, string[]>();
  const note = (group: string, message: string) => problems.set(group, [...(problems.get(group) ?? []), message]);
  const bound = (node: OutlineShape): boolean => !!parseBinding(node.name) || (node.children ?? []).some(bound);

  const walk = (nodes: OutlineShape[]) => {
    for (const node of nodes) {
      const binding = parseBinding(node.name);
      if (binding?.group) {
        found.push({ group: binding.group, index: binding.index!, ...(binding.field ? { field: binding.field } : {}), ...(binding.options ? { options: binding.options } : {}), node });
        if (node.children?.some(bound)) note(binding.group, `${node.name}（id ${node.id}）是一个组合，里面还有以 @ 开头的形状；单元里的字段要放在组合外面，或先取消组合`);
      } else if (node.children) walk(node.children);
    }
  };
  walk(shapes);

  const groups: Record<string, UnitGroup> = {};
  const names = [...new Set(found.map(item => item.group))];
  for (const name of names) {
    const mine = found.filter(item => item.group === name);
    const indices = [...new Set(mine.map(item => item.index))].sort((a, b) => a - b);
    const instances: UnitInstance[] = indices.map(index => {
      const members: UnitMember[] = mine.filter(item => item.index === index).map(item => ({
        id: item.node.id, kind: item.node.children ? "group" as const : "leaf" as const, rect: item.node.rect,
        ...(item.field && isImageShape(item.node) ? { field: item.field, image: imageSlotOf(item.node, shapes.length - 1 - shapes.indexOf(item.node)) }
          : item.field ? { field: item.field, text: item.node.text, fontSizePt: item.node.style?.fontSizePt, align: item.node.style?.align } : {})
      }));
      return { index, members, rect: union(members.map(member => member.rect)) };
    });
    if (indices.some((index, i) => index !== i + 1)) note(name, `“${name}”的序号应该是 1、2、3…连续，现在是 ${indices.join("、")}`);

    const fieldsOf = (instance: UnitInstance) => instance.members.filter(member => member.field).map(member => member.field!).sort();
    const first = fieldsOf(instances[0]);
    instances.forEach(instance => {
      if (fieldsOf(instance).join("|") !== first.join("|")) note(name, `“${name}${instance.index}”的字段和“${name}1”不一样；每一个都要有同样的字段`);
      if (instance.members.length !== instances[0].members.length) note(name, `“${name}${instance.index}”有 ${instance.members.length} 个形状，“${name}1”有 ${instances[0].members.length} 个；每一个的组成要一样`);
    });
    const duplicates = first.filter((field, i) => first.indexOf(field) !== i);
    if (duplicates.length > 0) note(name, `“${name}1”里字段 ${[...new Set(duplicates)].join("、")} 出现了不止一次`);
    for (const field of first) if (/\d$/.test(field)) note(name, `字段名“${field}”以数字结尾，会被当成另一个实例；改个名字`);

    const options = mine.find(item => item.options)?.options ?? {};
    const declared = /^(\d+)(?:\s*[-~～]\s*(\d+))?$/.exec(options["数量"] ?? "");
    const asked = { min: declared ? Number(declared[1]) : undefined, max: declared ? Number(declared[2] ?? declared[1]) : undefined };
    let fixed = "固定" in options;

    // Evenly spaced instances can be added and removed; any other arrangement is the template's number only.
    let pitch: UnitGroup["pitch"];
    if (instances.length >= 2) {
      const steps = instances.slice(1).map((instance, i) => ({ dx: instance.rect.x - instances[i].rect.x, dy: instance.rect.y - instances[i].rect.y }));
      const even = steps.every(step => near(step.dx, steps[0].dx, STEP_TOLERANCE) && near(step.dy, steps[0].dy, STEP_TOLERANCE));
      const alike = instances.every(instance => near(instance.rect.w, instances[0].rect.w, SIZE_TOLERANCE) && near(instance.rect.h, instances[0].rect.h, SIZE_TOLERANCE));
      if (even && alike && (Math.abs(steps[0].dx) > 0.2 || Math.abs(steps[0].dy) > 0.2)) pitch = steps[0];
    }
    if (!pitch && !fixed) {
      fixed = true;
      if (instances.length >= 2) note(name, `“${name}”的各个实例排得不规则（不是等距的一排或一列），只能用模板里的 ${instances.length} 个；如果就是这么设计的，在其中一个形状名字后加 |固定`);
    }
    if (instances.length < 2 && !fixed && asked.max === undefined) fixed = true;

    let min = instances.length;
    let max = instances.length;
    if (!fixed && pitch) {
      const horizontal = Math.abs(pitch.dx) >= Math.abs(pitch.dy);
      const span = horizontal ? union(instances.map(instance => instance.rect)).w : union(instances.map(instance => instance.rect)).h;
      const size = horizontal ? instances[0].rect.w : instances[0].rect.h;
      const step = Math.abs(horizontal ? pitch.dx : pitch.dy);
      // k instances at the smallest scale and the least gap fit the span: k·size·s + (k-1)·gap ≤ span.
      const room = Math.floor((span + MIN_UNIT_GAP) / (MIN_UNIT_SCALE * size + MIN_UNIT_GAP));
      void step;
      min = Math.max(1, asked.min ?? 1);
      max = Math.max(min, Math.min(asked.max ?? room, room));
    }
    groups[name] = {
      name, instances,
      fields: first.filter(field => !isAutoUnitField(field)), auto: first.filter(isAutoUnitField),
      min, max, fixed, ...(pitch ? { pitch } : {}), problems: problems.get(name) ?? []
    };
  }
  return groups;
}

/** What a model needs to know about a page's repeating units. */
export function describeUnits(groups: Record<string, UnitGroup>): Record<string, unknown> {
  return Object.fromEntries(Object.values(groups).map(group => {
    const first = group.instances[0];
    return [group.name, {
      fields: group.fields,
      fieldInfo: Object.fromEntries(group.fields.map(field => {
        const member = first.members.find(item => item.field === field)!;
        if (member.image) return [field, { kind: "image", aspect: Math.round(member.image.aspect * 100) / 100, ...(member.image.shape ? { shape: member.image.shape } : {}), note: "写图片的路径（project:images/…，用户上传的图片），或 placeholder 先放一张占位图" }];
        return [field, { ...(member.text ? { sample: member.text } : {}), ...capacityOf(member.rect, member.fontSizePt) }];
      })),
      templateCount: group.instances.length,
      count: group.fixed ? { fixed: group.instances.length } : { min: group.min, max: group.max },
      note: group.fixed
        ? `这个版式的“${group.name}”固定是 ${group.instances.length} 个：content 里给 ${group.instances.length} 个，数量对不上就换别的版式`
        : `“${group.name}”写成数组，每项是 {${group.fields.join("、")}}，可以给 ${group.min} 到 ${group.max} 个（模板里有 ${group.instances.length} 个）；多的会自动缩小排开`
    }];
  }));
}
