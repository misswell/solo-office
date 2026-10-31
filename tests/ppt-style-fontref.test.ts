import { describe, expect, it } from "vitest";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { Presentation, type Shape } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";

/**
 * 授权语料里 4 份模板的 18 个形状带 `p:style`，其中 1 份的 3 枚 run 自己不写颜色：
 * 它们的墨色来自 `p:style/a:fontRef`（主题色引用），页面此前退回默认黑字。
 * 同批形状都在 `a:spPr` 里自己答了填充与描边，`a:fillRef`/`a:lnRef` 两路 0 受众，不建模。
 */
const CORPUS = ".scratch/wps-corpus";
const dec = new TextDecoder();

/** makePptx 的椭圆带上 p:style：主题 accent1 作默认字色，另加一枚自带颜色的 run。 */
function styled(): Presentation {
  const zip = unzipSync(makePptx());
  const slide = dec
    .decode(zip["ppt/slides/slide1.xml"])
    // fontRef 只报主题引用（accent1 → 4472C4），颜色要经主题解出来。
    .replace(
      '<p:sp><p:nvSpPr><p:cNvPr id="4" name="椭圆"/><p:nvPr/></p:nvSpPr>\n<p:spPr>',
      '<p:sp><p:nvSpPr><p:cNvPr id="4" name="椭圆"/><p:nvPr/></p:nvSpPr><p:style><a:lnRef idx="2"/><a:fillRef idx="1"/><a:effectRef idx="0"/><a:fontRef idx="minor"><a:schemeClr val="accent1"/></a:fontRef></p:style>\n<p:spPr>',
    )
    // 一枚 run 自带颜色：它自己的回答比 p:style 优先。
    .replace(
      '<a:r><a:rPr sz="2000" b="1"/><a:t>+18.75%</a:t></a:r>',
      '<a:r><a:rPr sz="2000" b="1"/><a:t>+18.75%</a:t></a:r><a:r><a:rPr sz="2000"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr><a:t>红字</a:t></a:r>',
    )
    // 只写 idx、不写颜色的 fontRef 不该凭空造出字色。
    .replace(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="标题"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>\n<p:spPr>',
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="标题"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:style><a:fontRef idx="minor"/></p:style>\n<p:spPr>',
    );
  expect(slide).not.toBe(dec.decode(zip["ppt/slides/slide1.xml"]));
  zip["ppt/slides/slide1.xml"] = strToU8(slide);
  return Presentation.openBytes(zipSync(zip, { level: 6 }));
}

function flatten(shapes: Shape[]): Shape[] {
  const out: Shape[] = [];
  for (const s of shapes) {
    out.push(s);
    if (s.children) out.push(...flatten(s.children));
  }
  return out;
}

describe("pptx p:style/a:fontRef 的默认字色", () => {
  it("fontRef 的主题色解成真色并被无色的 run 继承，run 自己的颜色优先，空 ref 不造色", () => {
    const pres = styled();
    const shapes = flatten(pres.slides[0].shapes);
    const ellipse = shapes.find((s) => s.name === "椭圆");
    expect(ellipse?.styleTextColor).toBe("#4472C4");
    const title = shapes.find((s) => s.name === "标题");
    expect(title?.styleTextColor).toBeUndefined();

    const html = renderSlideHtml(pres, 0);
    // 椭圆那枚没有自己颜色的 run 用 fontRef 的 accent1。
    expect(html).toMatch(/color:#4472C4/);
    // 自带 srgbClr 的 run 不被 fontRef 盖掉；标题仍按 run 的 1F3864。
    expect(html).toMatch(/color:#C00000/);
    expect(html).toMatch(/color:#1F3864/);
    // 标题 run 上的 1F3864 只该出现一次（没有被 fontRef 追加一层）。
    expect(html.match(/color:#1F3864/g)?.length).toBe(1);
  });

  it.skipIf(!existsSync(CORPUS))("语料里 18 个带 p:style 的形状、3 枚无色 run 的墨色来自 fontRef", () => {
    const dir = CORPUS;
    const files = readdirSync(dir)
      .sort()
      .filter((f) => /\.pptx$/i.test(f))
      .map((f) => join(dir, f));
    let styledShapes = 0;
    let inheritedRuns = 0;
    let htmlColored = 0;
    const decks = new Set<string>();
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        for (const s of flatten(slide.shapes)) {
          if (!s.styleTextColor) continue;
          styledShapes++;
          decks.add(file);
          for (const p of s.paragraphs ?? []) {
            for (const r of p.runs) {
              if (r.color) continue;
              inheritedRuns++;
              expect(s.styleTextColor).toMatch(/^#[0-9A-F]{6}$/);
            }
          }
        }
        const html = renderSlideHtml(pres, slide.index);
        if (/color:#(?:FFFFFF|F2F2F2)/i.test(html)) htmlColored++;
      }
    }
    expect(files.length).toBeGreaterThanOrEqual(35);
    expect(styledShapes).toBe(18);
    expect(decks.size).toBe(4);
    expect(inheritedRuns).toBe(3);
    expect(htmlColored).toBeGreaterThanOrEqual(1);
  });
});
