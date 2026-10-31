import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { Presentation } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";

/**
 * 授权语料照子：35 份模板母版都写了 `p:otherStyle`，一级字号 33 份是 1800（18pt）、
 * 2 份是 1035（10.35pt）；幻灯片里 49 段形状文字自己不带 `a:rPr@sz`，但它们都写了
 * `a:txBody/a:lstStyle` 的一级字号 —— 按规范先取形状自己的，母版 otherStyle 是兜底。
 */
const CORPUS = ".scratch/wps-corpus";

function decks(): string[] {
  if (!existsSync(CORPUS)) return [];
  return readdirSync(CORPUS)
    .sort()
    .filter((f) => /\.(pptx|potx|ppsx)$/i.test(f))
    .map((f) => join(CORPUS, f));
}

const dec = new TextDecoder();

/** A copy of the fixture deck with `extra` shapes appended to slide 1 and `txStyles` in its master. */
function deckWith(extra: string, txStyles: string): Presentation {
  const zip = unzipSync(makePptx());
  const slide = dec.decode(zip["ppt/slides/slide1.xml"]).replace("</p:spTree>", `${extra}</p:spTree>`);
  zip["ppt/slides/slide1.xml"] = strToU8(slide);
  const master = dec.decode(zip["ppt/slideMasters/slideMaster1.xml"]).replace("</p:sldMaster>", `${txStyles}</p:sldMaster>`);
  zip["ppt/slideMasters/slideMaster1.xml"] = strToU8(master);
  return Presentation.openBytes(zipSync(zip));
}

const shape = (name: string, body: string, lstStyle = ""): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="9" name="${name}"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="9000000"/><a:ext cx="2000000" cy="600000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/>${lstStyle}<a:p><a:r><a:t>${body}</a:t></a:r></a:p></p:txBody></p:sp>`;

describe("形状文字的字号继承母版 p:otherStyle", () => {
  const files = decks();

  it.skipIf(!files.length)("每份语料的一级字号都从母版 otherStyle 读出来", () => {
    const seen = new Map<string, number>();
    let checked = 0;
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const zip = unzipSync(bytes);
      const master = Object.keys(zip).find((k) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(k));
      const xml = master ? dec.decode(zip[master]) : "";
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

  it.skipIf(!files.length)("不带字号的形状文字先取自己 lstStyle 的字号，没有才落回 otherStyle", () => {
    let verified = 0;
    let ownStated = 0;
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      pres.slides.forEach((slide, index) => {
        const html = renderSlideHtml(pres, index);
        slide.shapes.forEach((s, shapeIndex) => {
          if (s.placeholder !== undefined || s.kind !== "shape") return;
          const runs = (s.paragraphs ?? []).flatMap((p) => p.runs).filter((r) => r.text.trim());
          if (!runs.length || runs.some((r) => r.sizePt !== undefined)) return;
          const start = html.indexOf(`data-idx="${shapeIndex}"`);
          expect(start, `${file}#${index} ${s.name} 应出现在渲染结果里`).toBeGreaterThanOrEqual(0);
          const next = html.indexOf("data-idx=", start + 1);
          const slice = html.slice(start, next < 0 ? undefined : next);
          const stated = s.textLevels?.[0]?.sizePt;
          if (stated !== undefined) ownStated++;
          expect(slice, `${file}#${index} ${s.name} 的字号`).toContain(`font-size:${stated ?? pres.otherStylePt}pt`);
          verified++;
        });
      });
    }
    expect(verified, "语料里应有不带字号的形状文字").toBeGreaterThan(40);
    expect(ownStated, "这些形状文字自己写了 lstStyle 一级字号").toBeGreaterThan(40);
  });

  it("没有 lstStyle 的文本框落回母版 otherStyle，有 lstStyle 的按自己写", () => {
    const bare = shape("裸框", "裸框文字");
    const styled = shape(
      "自带样式",
      "自带样式的文字",
      `<a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:defRPr sz="2800"/></a:lvl1pPr></a:lstStyle>`,
    );
    const pres = deckWith(
      bare + styled,
      `<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:otherStyle></p:txStyles>`,
    );
    expect(pres.otherStylePt).toBe(24);
    const html = renderSlideHtml(pres, 0);
    const at = (name: string): string => {
      const idx = pres.slides[0].shapes.findIndex((s) => s.name === name);
      expect(idx, `夹具里应有形状 ${name}`).toBeGreaterThanOrEqual(0);
      const start = html.indexOf(`data-idx="${idx}"`);
      const next = html.indexOf("data-idx=", start + 1);
      return html.slice(start, next < 0 ? undefined : next);
    };
    expect(at("裸框")).toContain("font-size:24pt");
    expect(at("自带样式")).toContain("font-size:28pt");
    expect(at("自带样式")).toContain("line-height:1.95");
    expect(at("自带样式")).toContain("margin-top:10pt");
  });
});
