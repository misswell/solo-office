import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { Presentation, type Shape } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";

/**
 * 授权语料照子：35 份演示模板的主题都带 `a:fontScheme`（major/minor 各一组 latin/ea/cs）。
 * 16 份的幻灯片在形状 `a:lstStyle/lvlNpPr/a:defRPr` 里写 `+mn-lt`/`+mj-lt`（327 处），
 * 另有 1 份的 3 枚 run 直接写 `+mj-lt`。此前只把 `a:latin@typeface` 原样当字体名交给页面，
 * 主题引用既不是字体，继承它的文字则退回页面默认脸。这里钉住「主题引用按主题的
 * fontScheme 解成真名、形状级 style 的脸被段落/run 继承、空 ea 不冒充字体」。
 */
const CORPUS = ".scratch/wps-corpus";
const dec = new TextDecoder();

const FONT_SCHEME = `<a:fontScheme name="Office"><a:majorFont>
<a:latin typeface="Trebuchet MS"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Georgia"/><a:ea typeface="SimSun-ExtB"/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>`;

/** A copy of the fixture deck whose theme names fonts and whose slide inherits them. */
function themed(): Presentation {
  const zip = unzipSync(makePptx());
  zip["ppt/theme/theme1.xml"] = strToU8(dec.decode(zip["ppt/theme/theme1.xml"]).replace("</a:themeElements>", `${FONT_SCHEME}</a:themeElements>`));
  const slide = dec
    .decode(zip["ppt/slides/slide1.xml"])
    // The first body paragraph states its face at the paragraph level, as WPS's templates do.
    .replace(
      '<a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>收入达成 108%</a:t></a:r></a:p>',
      '<a:p><a:pPr lvl="0"><a:buChar char="•"/><a:defRPr sz="1800"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:pPr><a:r><a:rPr sz="1800"/><a:t>收入达成 108%</a:t></a:r></a:p>',
    )
    // One run names the major theme face directly.
    .replace('<a:rPr sz="2000" b="1"/>', '<a:rPr sz="2000" b="1"><a:latin typeface="+mj-lt"/></a:rPr>');
  zip["ppt/slides/slide1.xml"] = strToU8(slide);
  return Presentation.openBytes(zipSync(zip, { level: 6 }));
}

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

describe("pptx 主题字体（a:fontScheme）", () => {
  it("+mn-/+mj- 引用解成主题里的脸，段落 defRPr 被 run 继承，空 ea 不冒充字体", () => {
    const pres = themed();
    const shapes = flatten(pres.slides[0].shapes);
    const body = shapes.flatMap((s) => s.paragraphs ?? []).find((p) => p.runs.some((r) => r.text === "收入达成 108%"));
    expect(body?.font).toBe("Georgia");
    expect(body?.fontEastAsia).toBe("SimSun-ExtB");
    const quoted = shapes.flatMap((s) => s.paragraphs ?? []).flatMap((p) => p.runs).find((r) => r.text === "+18.75%");
    expect(quoted?.font).toBe("Trebuchet MS");

    const html = renderSlideHtml(pres, 0);
    expect(html).toContain("font-family:&quot;solo-ea-SimSun-ExtB&quot;,&quot;Georgia&quot;");
    expect(html).toContain("font-family:&quot;Trebuchet MS&quot;");
    expect(html).not.toMatch(/\+m[jn]-(?:lt|ea|cs)/);
  });

  it.skipIf(!existsSync(CORPUS))("语料里的主题引用全部解成主题声明的脸，页面上不再出现 +mj/+mn", () => {
    const files = decks();
    let levelFonts = 0;
    let runFonts = 0;
    const deckFonts = new Set<string>();
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const zip = unzipSync(bytes);
      const faces = new Set<string>();
      for (const part of Object.keys(zip).filter((n) => /^ppt\/theme\/theme\d+\.xml$/.test(n))) {
        for (const m of dec.decode(zip[part]).matchAll(/<a:(?:latin|ea|cs) typeface="([^"]+)"/g)) faces.add(m[1]);
      }
      const pres = Presentation.openBytes(bytes);
      for (const slide of pres.slides) {
        for (const shape of flatten(slide.shapes)) {
          for (const level of shape.textLevels ?? []) {
            if (!level.font) continue;
            levelFonts++;
            deckFonts.add(file);
            expect(faces.has(level.font), `${file} → ${level.font}`).toBe(true);
          }
          for (const run of [...(shape.paragraphs ?? []).flatMap((p) => p.runs), ...(shape.rows ?? []).flat().flatMap((c) => c.runs)]) {
            if (!run.font) continue;
            runFonts++;
            expect(faces.has(run.font), `${file} → ${run.font}`).toBe(true);
          }
        }
      }
      const html = pres.slides.map((_, i) => renderSlideHtml(pres, i)).join("");
      expect(html, file).not.toMatch(/\+m[jn]-(?:lt|ea|cs)/);
    }
    expect(files.length).toBeGreaterThanOrEqual(35);
    expect(levelFonts).toBeGreaterThanOrEqual(300);
    expect(deckFonts.size).toBeGreaterThanOrEqual(15);
    expect(runFonts).toBeGreaterThanOrEqual(3);
  });
});
