import { expect, it } from 'vitest';
import { brandProtectedRegions } from '../src/design/template-policy';

// Geometry recorded from the UniAI red/white template in the 2026-11-12 run.
const coverBand = { x: 0, y: 169, w: 1280, h: 364 };
const logos = { x: 1007, y: 38, w: 236, h: 31 };
const footerMark = { x: 979, y: 667, w: 288, h: 38 };
const fullPage = { x: 0, y: 0, w: 1280, h: 720 };

it('keeps logos protected but treats a band under the slide\'s own title as its background', () => {
  const coverTitle = { x: 0, y: 269, w: 1280, h: 154 }, coverDate = { x: 520, y: 455, w: 241, h: 23 };
  expect(brandProtectedRegions([coverBand, logos, fullPage], [coverTitle, coverDate])).toEqual([logos]);
  const contentTitle = { x: 0, y: 32, w: 149, h: 42 };
  expect(brandProtectedRegions([logos, footerMark], [contentTitle])).toEqual([logos, footerMark]);
});
