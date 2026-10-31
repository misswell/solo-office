import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync } from "fflate";
import { Presentation } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";

/**
 * 授权语料照子：35 份模板母版都写了 `p:otherStyle`，一级字号 33 份是 1800（18pt）、
 * 2 份是 1035（10.35pt）；幻灯片里 65 段形状文字有 49 段自己不带 `a:rPr@sz`。
 * 非占位符的形状文字按规范取母版 otherStyle，而不是页面样式表那 14px 兜底。
 */
const CORPUS = ".scratch/wps-corpus";

function decks(): string[] {
  if (!existsSync(CORPUS)) return [];
  return readdirSync(CORPUS)
    .sort()
    .filter((f) => /\.(pptx|potx|ppsx)$/i.test(f))
    .map((f) => join(CORPUS, f));
}

describe("形状文字的字号继承母版 p:otherStyle", () => {
  const files = decks();

  it.skipIf(!files.length)("每份语料的一级字号都从母版 otherStyle 读出来", () => {
    const seen = new Map<string, number>();
    let checked = 0;
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const zip = unzipSync(bytes);
      const master = Object.keys(zip).find((k) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(k));
      const xml = master ? new TextDecoder().decode(zip[master]) : "";
      const other = /<p:otherStyle>([\s\S]*?)<\/p:otherStyle>/.exec(xml)?.[1] ?? "";
      const sz = /<a:lvl1pPr\b[\s\S]*?<a:defRPr\b[^>]*\bsz="(\d+)"/.exec(other)?.[1];
      if (!sz) continue;
      expect(Presentation.openBytes(bytes).otherStylePt, file).toBeCloseTo(Number(sz) / 100, 5);
      seen.set(sz, (seen.get(sz) ?? 0) + 1);
      checked++;
    }
    expect(checked).toBe(files.length);
    expect(seen.get("1800")).toBeGreaterThan(0);
  });

  it.skipIf(!files.length)("不带字号的形状文字按 otherStyle 画，不再落回样式的 14px", () => {
    let verified = 0;
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      pres.slides.forEach((slide, index) => {
        const html = renderSlideHtml(pres, index);
        slide.shapes.forEach((shape, shapeIndex) => {
          if (shape.placeholder !== undefined || shape.kind !== "shape") return;
          const runs = (shape.paragraphs ?? []).flatMap((p) => p.runs).filter((r) => r.text.trim());
          if (!runs.length || runs.some((r) => r.sizePt !== undefined)) return;
          const start = html.indexOf(`data-idx="${shapeIndex}"`);
          expect(start, `${file}#${index} ${shape.name} 应出现在渲染结果里`).toBeGreaterThanOrEqual(0);
          const next = html.indexOf("data-idx=", start + 1);
          const own = html.slice(start, next < 0 ? undefined : next);
          expect(own, `${file}#${index} ${shape.name} 的字号`).toContain(`font-size:${pres.otherStylePt}pt`);
          verified++;
        });
      });
    }
    expect(verified, "语料里应有不带字号的形状文字").toBeGreaterThan(40);
  });
});
