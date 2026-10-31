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
 * 授权语料照子：幻灯片侧的行距/段前距只写在 `a:txBody/a:lstStyle` 的级里
 * （35 份 462 枚 lnSpc、462 枚 spcBef，`a:pPr` 一级 0 枚）；母版 `p:bodyStyle`
 * 一级 33 份写了 lnSpc（23 份 90%、10 份 130%）、33 份写了 spcBef（21 份 10pt、
 * 2 份 5.75pt、10 份明说 0）。读取、阶梯与画面上这一层此前整条缺失。
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
const level1 = (style: string): string => /<a:lvl1pPr\b[\s\S]*?<\/a:lvl1pPr>/.exec(style)?.[0] ?? "";

describe("幻灯片段落的行距与段前段后", () => {
  const files = decks();

  it.skipIf(!files.length)("母版 bodyStyle 一级的 lnSpc / spcBef 读得准", () => {
    let withLn = 0;
    let withBefore = 0;
    for (const file of files) {
      const zip = unzipSync(new Uint8Array(readFileSync(file)));
      const xml = bodyStyle(masterXml(zip));
      if (!xml) continue;
      const lvl = level1(xml);
      if (!lvl) continue;
      const pct = /<a:lnSpc>\s*<a:spcPct val="(\d+)"/.exec(lvl)?.[1];
      const before = /<a:spcBef>\s*<a:spcPts val="(\d+)"/.exec(lvl)?.[1];
      const level0 = Presentation.openBytes(new Uint8Array(readFileSync(file))).bodyLevels[0];
      if (pct) {
        expect(level0.lineSpacingPct, file).toBeCloseTo(Number(pct) / 1000, 5);
        withLn++;
      }
      if (before !== undefined) {
        // 明说 0 就是 0，不要当成「没写」退回别的段距。
        expect(level0.spaceBeforePt, file).toBe(Number(before) > 0 ? Number(before) / 100 : undefined);
        withBefore++;
      }
    }
    expect(withLn).toBeGreaterThan(25);
    expect(withBefore).toBeGreaterThan(25);
  });

  it.skipIf(!files.length)("级里写的行距/段前距真的画到段落上", () => {
    let ln = 0;
    let before = 0;
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      pres.slides.forEach((slide, index) => {
        const html = renderSlideHtml(pres, index);
        slide.shapes.forEach((s, shapeIndex) => {
          const start = html.indexOf(`data-idx="${shapeIndex}"`);
          if (start < 0) return;
          const next = html.indexOf("data-idx=", start + 1);
          const slice = html.slice(start, next < 0 ? undefined : next);
          for (const p of s.paragraphs ?? []) {
            if (!p.runs.some((r) => r.text.trim())) continue;
            const own = s.textLevels?.[Math.min(p.level, (s.textLevels?.length ?? 1) - 1)];
            const pct = p.lineSpacingPct ?? own?.lineSpacingPct;
            const mb = p.spaceBeforePt ?? own?.spaceBeforePt;
            if (pct) {
              expect(slice, `${file}#${index} ${s.name} 的行距`).toContain(`line-height:${Math.round(1.3 * pct) / 100}`);
              ln++;
            }
            if (mb) {
              expect(slice, `${file}#${index} ${s.name} 的段前距`).toContain(`margin-top:${mb}pt`);
              before++;
            }
          }
        });
      });
    }
    expect(ln, "语料里应有带行距的形状段落").toBeGreaterThan(50);
    expect(before, "语料里应有带段前距的形状段落").toBeGreaterThan(40);
  });
});

/** 空白夹具外挂一段 master txStyles，正文占位符的段落才谈得上继承。 */
function deckWithMasterSpacing(): Presentation {
  const zip = unzipSync(makePptx());
  const txStyles = `<p:txStyles><p:titleStyle/><p:bodyStyle>
<a:lvl1pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:defRPr sz="1800"/></a:lvl1pPr>
<a:lvl2pPr><a:lnSpc><a:spcPct val="130000"/></a:lnSpc></a:lvl2pPr>
</p:bodyStyle><p:otherStyle>${'<a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr>'}</p:otherStyle></p:txStyles>`;
  const master = dec.decode(zip["ppt/slideMasters/slideMaster1.xml"]).replace("</p:sldMaster>", `${txStyles}</p:sldMaster>`);
  zip["ppt/slideMasters/slideMaster1.xml"] = strToU8(master);
  const extra = `<a:p><a:pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc></a:pPr><a:r><a:t>自定行距</a:t></a:r></a:p>`;
  const styled = `<p:sp><p:nvSpPr><p:cNvPr id="9" name="带样式"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="9000000"/><a:ext cx="2000000" cy="600000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>
<a:p><a:r><a:t>形状自带的间距</a:t></a:r></a:p></p:txBody></p:sp>`;
  const slide = dec
    .decode(zip["ppt/slides/slide1.xml"])
    .replace("<a:t>费用需关注</a:t></a:r></a:p>", `<a:t>费用需关注</a:t></a:r></a:p>${extra}`)
    .replace("</p:spTree>", `${styled}</p:spTree>`);
  zip["ppt/slides/slide1.xml"] = strToU8(slide);
  return Presentation.openBytes(zipSync(zip));
}

describe("段落间距的阶梯", () => {
  it("正文占位符的一级按母版 bodyStyle 画，二级按二级画，段自己写的压倒一切", () => {
    const pres = deckWithMasterSpacing();
    expect(pres.bodyLevels[0]).toMatchObject({ lineSpacingPct: 90, spaceBeforePt: 5 });
    expect(pres.bodyLevels[1]).toMatchObject({ lineSpacingPct: 130 });
    const html = renderSlideHtml(pres, 0);
    const body = pres.slides[0].shapes.findIndex((s) => s.name === "正文");
    const start = html.indexOf(`data-idx="${body}"`);
    const slice = html.slice(start, html.indexOf("data-idx=", start + 1));
    expect(slice).toContain("line-height:1.17");
    expect(slice).toContain("margin-top:5pt");
    expect(slice).toContain("line-height:1.69");
    expect(slice).toContain("line-height:2.6");
  });

  it("改一段文字后保存，段落级与形状级那句间距都还在", () => {
    const pres = deckWithMasterSpacing();
    const body = pres.slides[0].shapes.find((s) => s.name === "正文")!;
    body.paragraphs = (body.paragraphs ?? []).map((p, i) => (i === 0 ? { ...p, runs: p.runs.map((r) => ({ ...r, text: "收入达成 109%" })) } : p));
    body.textDirty = true;
    const styled = pres.slides[0].shapes.find((s) => s.name === "带样式")!;
    styled.paragraphs = [{ level: 0, runs: [{ text: "形状自带的间距（已改）" }] }];
    styled.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const out = OpcPackage.from(writePptx(pres));
    const xml = out.text("ppt/slides/slide1.xml")!;
    // 段自己写的那句随段落重建原样写出；形状 lstStyle 是 txBody 的旁支，重建段落时留下。
    expect(xml).toContain('val="200000"');
    expect(xml).toContain('val="150000"');
    expect(xml).toContain('val="1000"');
    const back = Presentation.openBytes(writePptx(pres));
    const shapes = back.slides[0].shapes;
    const paragraphs = shapes.find((s) => s.name === "正文")!.paragraphs!;
    expect(paragraphs[0].runs[0].text).toBe("收入达成 109%");
    const custom = paragraphs.find((p) => p.runs.some((r) => r.text === "自定行距"))!;
    expect(custom.lineSpacingPct).toBe(200);
    expect(shapes.find((s) => s.name === "带样式")!.textLevels?.[0]).toMatchObject({ lineSpacingPct: 150, spaceBeforePt: 10 });
  });
});
