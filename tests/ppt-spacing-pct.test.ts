import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { Presentation } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { writePptx } from "../src/core/write-pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { makePptx } from "./helpers/make.js";

/**
 * 授权语料照子：段前/段后距除了 `a:spcPts`（绝对磅数）还有 `a:spcPct`（行高的百分比）。
 * 35 份 pptx 里只有两份**主题模板**写了非零百分比——`default.pptx` 与 `newfile.pptx`
 * 的 `p:bodyStyle` 一级到九级全是 `spcBef` 20%，其余母版/幻灯片全是明说 `val="0"`。
 * 这两份正是从 WPS 新建演示文稿时的默认底稿，所以 20% 就是新件的默认段前距。
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
const masterXml = (zip: Record<string, Uint8Array>): string => {
  const path = Object.keys(zip).find((k) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(k));
  return path ? dec.decode(zip[path]) : "";
};
const bodyStyle = (xml: string): string => /<p:bodyStyle>([\s\S]*?)<\/p:bodyStyle>/.exec(xml)?.[1] ?? "";
const levelOf = (style: string, n: number): string => new RegExp(`<a:lvl${n}pPr\\b[^>]*>[\\s\\S]*?</a:lvl${n}pPr>`).exec(style)?.[0] ?? "";

const beforePct = (lvl: string): string | undefined => /<a:spcBef\b[^>]*>\s*<a:spcPct val="(\d+)"/.exec(lvl)?.[1];
const beforePts = (lvl: string): string | undefined => /<a:spcBef\b[^>]*>\s*<a:spcPts val="(\d+)"/.exec(lvl)?.[1];

describe("母版 bodyStyle 的百分比段前距", () => {
  const files = decks();

  it.skipIf(!files.length)("只有默认主题的九级写了 20%，模型只留百分比、不留磅数", () => {
    let levels = 0;
    const named: string[] = [];
    for (const file of files) {
      const zip = unzipSync(new Uint8Array(readFileSync(file)));
      const style = bodyStyle(masterXml(zip));
      if (!style) continue;
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (let n = 1; n <= 9; n++) {
        const lvl = levelOf(style, n);
        if (!lvl) continue;
        const pct = beforePct(lvl);
        if (pct === undefined) continue;
        // 一个 spcBef 只能选一种单位，不会磅数与百分比并存。
        expect(beforePts(lvl), file).toBeUndefined();
        const model = pres.bodyLevels[n - 1];
        if (Number(pct) === 0) {
          // 明说 0 与不写同画：模型里是「没有」，不是 0 个百分点。
          expect(model.spaceBeforePct, `${file} lvl${n}`).toBeUndefined();
          expect(model.spaceBeforePt, `${file} lvl${n}`).toBeUndefined();
          continue;
        }
        expect(model.spaceBeforePct, `${file} lvl${n}`).toBeCloseTo(Number(pct) / 1000, 5);
        expect(model.spaceBeforePt, `${file} lvl${n}`).toBeUndefined();
        levels++;
        named.push(`${file.slice(-24)} lvl${n}`);
      }
    }
    expect(levels, "两份默认主题各九级").toBe(18);
    expect(named.every((n) => /default\.pptx|newfile\.pptx/.test(n))).toBe(true);
  });
});

/** 空白夹具的母版 bodyStyle 写 20% 段前，一级行距 100%、二级 90%。 */
function pctDeck(): Presentation {
  const zip = unzipSync(makePptx());
  const txStyles = `<p:txStyles><p:titleStyle/><p:bodyStyle>
<a:lvl1pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:defRPr sz="1800"/></a:lvl1pPr>
<a:lvl2pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="20000"/></a:spcBef></a:lvl2pPr>
</p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>`;
  const master = dec.decode(zip["ppt/slideMasters/slideMaster1.xml"]).replace("</p:sldMaster>", `${txStyles}</p:sldMaster>`);
  zip["ppt/slideMasters/slideMaster1.xml"] = strToU8(master);
  const slide = dec
    .decode(zip["ppt/slides/slide1.xml"])
    .replace("<a:t>费用需关注</a:t></a:r></a:p>", '<a:t>费用需关注</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:t>二级正文</a:t></a:r></a:p>');
  zip["ppt/slides/slide1.xml"] = strToU8(slide);
  return Presentation.openBytes(zipSync(zip));
}

describe("百分比段前距画到段落上", () => {
  it("按该段落自己的行高换算：一级 100% 行高留 0.26em，二级 90% 留 0.234em", () => {
    const pres = pctDeck();
    expect(pres.bodyLevels[0]).toMatchObject({ lineSpacingPct: 100, spaceBeforePct: 20 });
    expect(pres.bodyLevels[1]).toMatchObject({ lineSpacingPct: 90, spaceBeforePct: 20 });
    const html = renderSlideHtml(pres, 0);
    const body = pres.slides[0].shapes.findIndex((s) => s.name === "正文");
    const start = html.indexOf(`data-idx="${body}"`);
    const slice = html.slice(start, html.indexOf("data-idx=", start + 1));
    expect(slice).toContain("line-height:1.3");
    expect(slice).toContain("margin-top:0.26em");
    expect(slice).toContain("line-height:1.17");
    expect(slice).toContain("margin-top:0.234em");
  });

  it("更近的一级写了磅数就压过母版的百分比，两种单位不并存", () => {
    const zip = unzipSync(makePptx());
    const txStyles = `<p:txStyles><p:titleStyle/><p:bodyStyle>
<a:lvl1pPr><a:spcBef><a:spcPct val="20000"/></a:spcBef></a:lvl1pPr>
</p:bodyStyle><p:otherStyle/></p:txStyles>`;
    const master = dec.decode(zip["ppt/slideMasters/slideMaster1.xml"]).replace("</p:sldMaster>", `${txStyles}</p:sldMaster>`);
    zip["ppt/slideMasters/slideMaster1.xml"] = strToU8(master);
    const slide = dec
      .decode(zip["ppt/slides/slide1.xml"])
      .replace(
        '<p:txBody><a:bodyPr/>\n<a:p><a:pPr lvl="0"><a:buChar',
        '<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:spcBef><a:spcPts val="1000"/></a:spcBef></a:lvl1pPr></a:lstStyle>\n<a:p><a:pPr lvl="0"><a:buChar',
      );
    zip["ppt/slides/slide1.xml"] = strToU8(slide);
    const pres = Presentation.openBytes(zipSync(zip));
    const html = renderSlideHtml(pres, 0);
    const body = pres.slides[0].shapes.findIndex((s) => s.name === "正文");
    const slice = html.slice(html.indexOf(`data-idx="${body}"`), html.indexOf("data-idx=", html.indexOf(`data-idx="${body}"`) + 1));
    expect(slice).toContain("margin-top:10pt");
    expect(slice).not.toContain("margin-top:0.26em");
  });

  it("段自己写的百分比读得出，改字保存后仍写回 spcPct 而不是改成磅数", () => {
    const zip = unzipSync(makePptx());
    const slide = dec
      .decode(zip["ppt/slides/slide1.xml"])
      .replace(
        '<a:p><a:pPr lvl="0" algn="r"/>',
        '<a:p><a:pPr lvl="0" algn="r"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr>',
      );
    zip["ppt/slides/slide1.xml"] = strToU8(slide);
    const pres = Presentation.openBytes(zipSync(zip));
    const body = pres.slides[0].shapes.find((s) => s.name === "正文")!;
    const at = body.paragraphs!.findIndex((p) => p.runs.some((r) => r.text === "费用需关注"));
    expect(body.paragraphs?.[at]).toMatchObject({ spaceBeforePct: 20, spaceAfterPct: 50 });
    body.paragraphs = body.paragraphs!.map((p, i) => (i === at ? { ...p, runs: p.runs.map((r) => ({ ...r, text: "费用需关注 109%" })) } : p));
    body.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const xml = OpcPackage.from(writePptx(pres)).text("ppt/slides/slide1.xml")!;
    expect(xml).toMatch(/<a:spcBef>\s*<a:spcPct val="20000"\s*\/?>/);
    expect(xml).toMatch(/<a:spcAft>\s*<a:spcPct val="50000"\s*\/?>/);
    expect(xml).not.toContain("<a:spcPts");
    expect(Presentation.openBytes(writePptx(pres)).slides[0].shapes.find((s) => s.name === "正文")!.paragraphs?.[at]).toMatchObject({
      spaceBeforePct: 20,
      spaceAfterPct: 50,
    });
  });
});
