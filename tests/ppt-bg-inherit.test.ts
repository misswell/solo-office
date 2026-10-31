import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { Presentation } from "../src/core/pptx.js";

/**
 * 授权语料照子：35 份演示模板共 31 页，其中 30 页自己不带 `p:bg`，背景只写在母版上
 * （29 份是 `p:bgRef idx=1001` 取 bg1，黑模板 2 份实底 tx2，灰模板 2 份 bg2 渐变，
 * 另有 1 页自己写 `p:bg`）。没写的页必须按 版式 → 母版 取到 CSS 背景，
 * 否则黑色/灰色模板会开成白页。颜色按母版 `p:clrMap` 落到主题色（bg1=lt1、tx2=dk2…），
 * 主题只认幻灯片母版关系指向的那份——备注/讲义主题（44546A）不能盖掉母版主题。
 */
const CORPUS = ".scratch/wps-corpus";

function decks(): string[] {
  if (!existsSync(CORPUS)) return [];
  return readdirSync(CORPUS)
    .sort()
    .filter((f) => /\.(pptx|potx|ppsx)$/i.test(f))
    .map((f) => join(CORPUS, f));
}

function findBySuffix(files: string[], suffix: string): string | undefined {
  return files.find((f) => f.endsWith(suffix));
}

describe("幻灯片背景继承版式与母版", () => {
  const files = decks();

  it.skipIf(!files.length)("每页都解析出背景，黑色与灰色模板不再是白页", () => {
    let slides = 0;
    const distinct = new Set<string>();
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        slides++;
        const bg = slide.background;
        expect(bg, `${file}#${slide.index} 应解析出背景`).toBeTruthy();
        expect(bg, `${file}#${slide.index} 的背景应是 CSS 色值或线性渐变`).toMatch(
          /^(#[0-9A-F]{6}|linear-gradient\(.+\))$/,
        );
        distinct.add(bg!);
      }
    }
    expect(slides).toBeGreaterThan(0);
    // White templates say nothing visible; the dark and grey ones are the audience.
    expect(distinct.size, "语料里应既有白底也有深色/渐变底").toBeGreaterThan(1);
  });

  it.skipIf(!files.length)("黑色模板取母版的 tx2，不被备注主题色替换", () => {
    for (const suffix of [
      "blanktemplate__normal_black.pptx",
      "blanktemplate__normal_black_mac.pptx",
    ]) {
      const file = findBySuffix(files, suffix);
      if (!file) continue;
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        expect(slide.background, `${suffix}#${slide.index}`).toBe("#0F1423");
      }
    }
  });

  it.skipIf(!files.length)("灰色模板把母版的 bg2 渐变画成 CSS 线性渐变", () => {
    for (const suffix of [
      "blanktemplate__normal_gray.pptx",
      "blanktemplate__normal_gray_mac.pptx",
    ]) {
      const file = findBySuffix(files, suffix);
      if (!file) continue;
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        expect(slide.background, `${suffix}#${slide.index}`).toBe(
          "linear-gradient(180deg, #FFFFFF 0%, #D9D9D9 100%)",
        );
      }
    }
  });
});
