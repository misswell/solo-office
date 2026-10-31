import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync } from "fflate";
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

  it.skipIf(!files.length)("版式自己不带变换的内容占位符，几何从母版补齐", () => {
    // WPS 的 330 份版式里 1,101 枚占位符不带 a:xfrm：其中 dt/ftr/sldNum 各 330 枚、
    // 标题 83 枚、正文 7 枚、obj 21 枚；母版的 175 枚全部带几何，所以它们得从母版取。
    let declared = 0;
    let resolved = 0;
    for (const file of files) {
      const zip = unzipSync(new Uint8Array(readFileSync(file)));
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const choice of pres.layoutChoices()) {
        const xml = zip[choice.path] ? new TextDecoder().decode(zip[choice.path]) : "";
        const raw = [...xml.matchAll(/(<p:sp\b[\s\S]*?<\/p:sp>)/g)]
          .map((m) => /<p:ph\b([^>]*?)\/?>/.exec(m[1])?.[1])
          .filter((a): a is string => a !== undefined)
          .map((a) => /\btype="([^"]+)"/.exec(a)?.[1] ?? "obj")
          .filter((t) => ["title", "ctrTitle", "subTitle", "body", "obj"].includes(t));
        const parsed = pres.addSlide(choice.path).layoutPlaceholders ?? [];
        for (const type of raw) {
          declared++;
          const hit = parsed.filter((p) => (p.type ?? "obj") === type);
          expect(hit.length, `${file} ${choice.path} 的 ${type} 占位符少了一枚`).toBeGreaterThan(0);
          for (const p of hit) {
            expect(p.width, `${file} ${choice.path} 的 ${type} 宽`).toBeGreaterThan(0);
            expect(p.height, `${file} ${choice.path} 的 ${type} 高`).toBeGreaterThan(0);
          }
          resolved++;
        }
      }
    }
    expect(declared).toBeGreaterThan(0);
    expect(resolved).toBe(declared);
  });
});
