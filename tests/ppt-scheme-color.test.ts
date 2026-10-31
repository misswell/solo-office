import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { OpcPackage, children as kids, rawChildren, tagLocalOf } from "../src/core/opc.js";
import { Presentation, type Shape } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";
import { writePptx } from "../src/core/write-pptx.js";

/**
 * 授权语料照子：35 份演示模板里 8 份用主题色（`a:schemeClr`）上色——**页面的填充 9 处、
 * 线色 35 处**，run 色与单元格填充各 0 处。主题色经母版 `p:clrMap` 落成具体色值
 * （bg1→lt1、tx1→dk1…），而填充/描边此前只认 `a:srgbClr`，这 44 处全数画不出来。
 * 颜色变换（lumMod/lumOff/tint/shade）按 ECMA-376 在 HSL 的亮度上做，不是逐通道相乘。
 */
const CORPUS = ".scratch/wps-corpus";
const dec = new TextDecoder();

function decks(): string[] {
  if (!existsSync(CORPUS)) return [];
  return readdirSync(CORPUS)
    .sort()
    .filter((f) => /\.pptx$/i.test(f))
    .map((f) => join(CORPUS, f));
}

function flatten(shapes: Shape[]): Shape[] {
  const out: Shape[] = [];
  for (const s of shapes) {
    out.push(s);
    if (s.children) out.push(...flatten(s.children));
  }
  return out;
}

function direct(node: any, tag: string): any {
  return node ? rawChildren(node).find((c) => tagLocalOf(c) === tag) : undefined;
}

function schemeOf(fill: any): any {
  return direct(fill, "schemeClr");
}

function deckWith(extra: string): Presentation {
  const zip = unzipSync(makePptx());
  const slide = dec.decode(zip["ppt/slides/slide1.xml"]).replace("</p:spTree>", `${extra}</p:spTree>`);
  zip["ppt/slides/slide1.xml"] = strToU8(slide);
  return Presentation.openBytes(zipSync(zip));
}

function schemeShape(name: string, spPr: string): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="20" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="9000000"/><a:ext cx="1000000" cy="500000"/></a:xfrm>${spPr}</p:spPr>
<p:txBody><a:bodyPr/><a:p><a:r><a:t>${name}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

describe("主题色（a:schemeClr）落成具体色值", () => {
  const files = decks();

  it.skipIf(!files.length)("语料里 9 处主题填充、35 处主题描边都解出了色值并进了渲染", () => {
    let fillCount = 0;
    let lineCount = 0;
    const schemeDecks = new Set<string>();
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        const html = renderSlideHtml(pres, slide.index);
        for (const s of flatten(slide.shapes)) {
          if (s.kind !== "shape" || !s.raw) continue;
          const spPr = direct(s.raw, "spPr");
          const fill = direct(spPr, "solidFill");
          if (schemeOf(fill)) {
            schemeDecks.add(file);
            fillCount++;
            expect(s.fill, `${file}#${slide.index} ${s.name} 的填充`).toMatch(/^#[0-9A-F]{6}$/);
            expect(html, `${file}#${slide.index} ${s.name} 的填充进渲染`).toContain(`background:${s.fill}`);
          }
          const lnFill = direct(direct(spPr, "ln"), "solidFill");
          if (schemeOf(lnFill)) {
            schemeDecks.add(file);
            lineCount++;
            expect(s.lineColor, `${file}#${slide.index} ${s.name} 的描边`).toMatch(/^#[0-9A-F]{6}$/);
            expect(html, `${file}#${slide.index} ${s.name} 的描边进渲染`).toContain(`solid ${s.lineColor}`);
          }
        }
      }
    }
    expect(schemeDecks.size, "语料里应有 8 份用主题色上色的模板").toBeGreaterThanOrEqual(8);
    expect(fillCount, "主题填充的受众").toBeGreaterThanOrEqual(9);
    expect(lineCount, "主题描边的受众").toBeGreaterThanOrEqual(35);
  });

  it("母版 clrMap 把 bg1/tx1 落到主题槽，identity 与 tint/shade/lumMod/lumOff 各按规范算", () => {
    const extra = [
      schemeShape("主题实心", `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`),
      schemeShape("主题恒等", `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="100000"/></a:schemeClr></a:solidFill>`),
      schemeShape("主题提亮", `<a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill>`),
      schemeShape("主题压暗", `<a:solidFill><a:schemeClr val="bg1"><a:shade val="25000"/></a:schemeClr></a:solidFill>`),
      schemeShape("主题调亮", `<a:solidFill><a:schemeClr val="bg1"><a:lumMod val="50000"/><a:lumOff val="20000"/></a:schemeClr></a:solidFill>`),
      schemeShape("主题描边", `<a:ln w="12700"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>`),
    ].join("");
    const pres = deckWith(extra);
    const byName = (n: string): Shape => pres.slides[0].shapes.find((s) => s.name === n)!;
    expect(byName("主题实心").fill).toBe("#4472C4");
    expect(byName("主题恒等").fill).toBe("#4472C4");
    expect(byName("主题提亮").fill).toBe("#404040");
    expect(byName("主题压暗").fill).toBe("#404040");
    expect(byName("主题调亮").fill).toBe("#B3B3B3");
    expect(byName("主题描边").lineColor).toBe("#000000");
    expect(byName("主题描边").lineWidth).toBeCloseTo(12700 / 9525, 5);
    const html = renderSlideHtml(pres, 0);
    const idx = pres.slides[0].shapes.findIndex((s) => s.name === "主题调亮");
    expect(html.slice(html.indexOf(`data-idx="${idx}"`), html.indexOf("data-idx=", html.indexOf(`data-idx="${idx}"`) + 1))).toContain(
      "background:#B3B3B3",
    );
  });

  it("改一次文字再保存：主题色那两句 a:schemeClr 按原件带回，不被改写成 sRGB", () => {
    const pres = deckWith(schemeShape("主题实心", `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="100000"/></a:schemeClr></a:solidFill>`));
    const shape = pres.slides[0].shapes.find((s) => s.name === "主题实心")!;
    shape.paragraphs = [{ level: 0, runs: [{ text: "已改" }] }];
    shape.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const bytes = writePptx(pres);
    const xml = OpcPackage.from(bytes).text("ppt/slides/slide1.xml")!;
    expect(xml).toContain('<a:schemeClr val="accent1">');
    expect(xml).toContain('<a:lumMod val="100000"/>');
    const back = Presentation.openBytes(bytes).slides[0].shapes.find((s) => s.name === "主题实心")!;
    expect(back.fill).toBe("#4472C4");
    expect(back.paragraphs?.[0].runs[0].text).toBe("已改");
  });
});
