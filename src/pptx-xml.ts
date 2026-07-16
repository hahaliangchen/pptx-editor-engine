export function querySelector(parent: Element | Document, selectors: string): Element | null {
  const parts = selectors.split(",").map(value => value.trim());
  for (const part of parts) {
    try {
      const element = parent.querySelector(part);
      if (element) return element as Element;
    } catch (_error) {
      // Namespace-aware selectors are not supported by every DOM implementation.
    }
  }

  for (const part of parts) {
    const cleanTag = part.replace(/^.*\\:/, "").replace(/^.*:/, "");
    const elements = parent.getElementsByTagName(cleanTag);
    if (elements.length > 0) return elements[0] as Element;
  }
  return null;
}

export function querySelectorAll(parent: Element | Document, selectors: string): Element[] {
  const parts = selectors.split(",").map(value => value.trim());
  for (const part of parts) {
    try {
      const elements = parent.querySelectorAll(part);
      if (elements.length > 0) return Array.from(elements) as Element[];
    } catch (_error) {
      // Namespace-aware selectors are not supported by every DOM implementation.
    }
  }

  for (const part of parts) {
    const cleanTag = part.replace(/^.*\\:/, "").replace(/^.*:/, "");
    const elements = parent.getElementsByTagName(cleanTag);
    if (elements.length > 0) return Array.from(elements) as Element[];
  }
  return [];
}

export function hasLocalName(node: Element, names: string[]): boolean {
  const localName = node.localName || node.nodeName.replace(/^.*:/, "");
  return names.includes(localName);
}

export function getDirectChildren(parent: Element | Document, ...names: string[]): Element[] {
  return Array.from(parent.childNodes)
    .filter((node): node is Element => node.nodeType === 1)
    .filter(node => hasLocalName(node, names));
}

export function getDirectChild(parent: Element | Document, ...names: string[]): Element | null {
  return getDirectChildren(parent, ...names)[0] || null;
}
