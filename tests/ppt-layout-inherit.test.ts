import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { Presentation } from "../src/core/pptx.js";

/**
 * 授权语料照子：WPS 自己写出的幻灯片里，占位符一枚 `a:xfrm` 都不带，几何全在版式上
 * （14 份 28 枚）。没带变换的占位符必须按 `p:ph@idx`、再按类型从版式取位置和大小，
 * 否则打开真件时标题/正文会挤成左上角 0×0 的盒子。
 */
const CORPUS = ".scratch/wps-corpus";

function decks(): string[] {
  if (!existsSync(CORPUS)) return [];
  return readdirSync(CORPUS)
    .sort()
    .filter((f) => /\.(pptx|potx|ppsx)$/i.test(f))
    .map((f) => join(CORPUS, f));
}

describe("幻灯片占位符的几何继承版式", () => {
  const files = decks();

  it.skipIf(!files.length)("语料里每个不带变换的占位符都取到了版式上的非零几何", () => {
    let inherited = 0;
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        const layout = slide.layoutPlaceholders;
        if (!layout?.length) continue;
        for (const shape of slide.shapes) {
          if (shape.placeholder === undefined) continue;
          expect(shape.width, `${file}#${slide.index} ${shape.name} 的宽`).toBeGreaterThan(0);
          expect(shape.height, `${file}#${slide.index} ${shape.name} 的高`).toBeGreaterThan(0);
          const hit = layout.some(
            (p) =>
              (shape.placeholderIdx !== undefined && p.idx === shape.placeholderIdx) ||
              (p.type ?? "obj") === (shape.placeholder || "obj"),
          );
          expect(hit, `${file}#${slide.index} ${shape.name} 在版式上应有对应占位符`).toBe(true);
          inherited++;
        }
      }
    }
    expect(inherited).toBeGreaterThan(0);
  });
});
