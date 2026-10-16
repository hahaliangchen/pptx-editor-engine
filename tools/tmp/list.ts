import { readFileSync } from "node:fs";
import { HeadlessSession } from "../../src/headless/session";
const s = await HeadlessSession.open(readFileSync(process.argv[2]), { quiet: true });
s.document.slides.forEach((slide, i) => {
  for (const el of slide.elements) if (el.type === "chart") {
    const c = el.chart;
    console.log(`slide ${i + 1} ${el.id} kind=${c.kind} unsupported=${c.unsupported ?? "-"} stacked=${c.stacked} title=${JSON.stringify(c.title)} cats=${c.categories.length} series=${c.series.map(x => x.name + ":" + x.values.length).join("|")} legend=${c.legend} labels=${c.dataLabels} pt=${c.labelPt}/${c.titlePt} font=${c.font.family}`);
  }
});
console.log("slides", s.document.slides.length);
s.dispose();
