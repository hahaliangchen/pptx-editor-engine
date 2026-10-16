import { OrderOp } from "../pptx-exporter";

export interface Move {
  key: string;
  op: OrderOp;
}

/**
 * New stacking order for sibling objects, given bottom-to-top `order` and the
 * selected keys. `moves` is the list of single operations, in the order they
 * must be replayed on the OOXML tree to give the same result; it is empty when
 * nothing changes.
 */
export function reorderSiblings(
  order: string[],
  selected: Set<string>,
  mode: OrderOp
): { order: string[]; moves: Move[] } {
  const moves: Move[] = [];
  const picked = order.filter(key => selected.has(key));
  const rest = order.filter(key => !selected.has(key));

  if (mode === "front") {
    const next = [...rest, ...picked];
    const changed = next.some((key, index) => key !== order[index]);
    return { order: next, moves: changed ? picked.map(key => ({ key, op: "front" as const })) : [] };
  }
  if (mode === "back") {
    const next = [...picked, ...rest];
    const changed = next.some((key, index) => key !== order[index]);
    return { order: next, moves: changed ? [...picked].reverse().map(key => ({ key, op: "back" as const })) : [] };
  }

  const next = [...order];
  if (mode === "forward") {
    for (let i = next.length - 2; i >= 0; i--) {
      if (selected.has(next[i]) && !selected.has(next[i + 1])) {
        moves.push({ key: next[i], op: "forward" });
        [next[i], next[i + 1]] = [next[i + 1], next[i]];
      }
    }
  } else {
    for (let i = 1; i < next.length; i++) {
      if (selected.has(next[i]) && !selected.has(next[i - 1])) {
        moves.push({ key: next[i], op: "backward" });
        [next[i], next[i - 1]] = [next[i - 1], next[i]];
      }
    }
  }
  return { order: next, moves };
}
