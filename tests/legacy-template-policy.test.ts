import {describe, expect, it} from "vitest";
import {applyDefaultPolicies, effectivePurpose} from "../src/design/template-policy";
import type {DesignDocument, DesignNode, DesignPage} from "../src/design/model";

function node(id: string, text: string, purpose?: DesignNode["purpose"], name = "TextBox"): DesignNode {
  return {id, name, kind: "text", purpose, appearance: {rect: {x: 20, y: 20, w: 300, h: 40}, text: {
    paragraphs: [{runs: [{text}]}]
  }}} as DesignNode;
}

describe("legacy template placeholders", () => {
  it("keeps repeated variable headings editable while protecting real brand text and explicit locks", () => {
    const pages = [1, 2].map(i => ({id: String(i), nodes: [
      node(`title${i}`, "{{PAGE_TITLE}}"), node(`number${i}`, "0{{CHAPTER_NUM}}"),
      node(`brand${i}`, "中国联通"), node(`locked${i}`, "{{PAGE_TITLE}}", "fixed"),
      node(`named${i}`, "{{TITLE}}", undefined, "brand logo")
    ]} as DesignPage));
    applyDefaultPolicies({pages} as DesignDocument);
    for (const page of pages) {
      expect(effectivePurpose(page, page.nodes[0])).toBe("content");
      expect(effectivePurpose(page, page.nodes[1])).toBe("content");
      expect(effectivePurpose(page, page.nodes[2])).toBe("fixed");
      expect(effectivePurpose(page, page.nodes[3])).toBe("fixed");
      expect(effectivePurpose(page, page.nodes[4])).toBe("fixed");
    }
  });
});
