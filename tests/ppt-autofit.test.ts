import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { OpcPackage } from "../src/core/opc.js";
import { Presentation, type Shape } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";
import { writePptx } from "../src/core/write-pptx.js";

/**
 * 授权语料照子：35 份演示模板的幻灯片里 normAutofit 说了 73 处，其中 69 处不带 `fontScale`
 * ——「溢出时缩字号」是给渲染器现场量的，没有落在文件里；只有 2 份模板的 4 个形状把
 * 挑好的比例写进了档：一张版式的 3 个对角圆角矩形 `fontScale="90000"`（run 自己写 36 磅，
 * 页面得按 32.4 磅画），另一张版式组里的「空心弧 10」是 `82500`（空段落，无字可缩）。
 * `lnSpcReduction` 在语料里 0 处，下面只拿合成件把这条腿钉住。版式与母版的 normAutofit
 * 各 278/102 处，全部不带 `fontScale`（没有可继承的尺寸），所以阶梯上不加一条。
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

function shapeSlice(html: string, idx: number): string {
  const start = html.indexOf(`data-idx="${idx}"`);
  return html.slice(start, html.indexOf("data-idx=", start + 1));
}

function deckWith(extra: string): Presentation {
  const zip = unzipSync(makePptx());
  const slide = dec.decode(zip["ppt/slides/slide1.xml"]).replace("</p:spTree>", `${extra}</p:spTree>`);
  zip["ppt/slides/slide1.xml"] = strToU8(slide);
  return Presentation.openBytes(zipSync(zip));
}

function fitShape(name: string, bodyPr: string, paragraph: string): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="21" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="8000000"/><a:ext cx="3000000" cy="900000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr>${bodyPr}</a:bodyPr>${paragraph}</p:txBody></p:sp>`;
}

describe("溢出缩字号（a:normAutofit）", () => {
  const files = decks();

  it.skipIf(!files.length)("语料里 4 个写了 fontScale 的形状：有字的 3 个按 90% 画，run 那句 36 磅跟着缩", () => {
    const scaled: { deck: string; slide: number; idx: number; shape: Shape }[] = [];
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        slide.shapes.forEach((s, idx) => {
          for (const inner of flatten([s])) if (inner.fontScalePct !== undefined) scaled.push({ deck: file, slide: slide.index, idx, shape: inner });
        });
      }
    }
    expect(scaled.length).toBe(4);
    const withRuns = scaled.filter((hit) => hit.shape.paragraphs?.some((p) => p.runs.length));
    expect(withRuns.length).toBe(3);
    for (const hit of withRuns) {
      expect(hit.shape.fontScalePct).toBe(90);
      expect(hit.shape.lineSpacingReductionPct).toBeUndefined();
      expect(hit.shape.paragraphs?.[0]?.runs[0]?.sizePt).toBe(36);
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(hit.deck)));
      const html = renderSlideHtml(pres, hit.slide);
      expect(shapeSlice(html, hit.idx), hit.shape.name).toContain("font-size:32.4pt");
    }
    // The fourth is a member of a group and has no runs; the group's members must still carry
    // their part nodes, or the parser cannot see the `bodyPr` it hides.
    const empty = scaled.find((hit) => !hit.shape.paragraphs?.some((p) => p.runs.length))!;
    expect(empty.shape.name).toBe("空心弧 10");
    expect(empty.shape.fontScalePct).toBe(82.5);
    expect(empty.shape.lineSpacingReductionPct).toBeUndefined();
  });

  it.skipIf(!files.length)("改字保存后 fontScale 那句还在，读回也还是 90%", () => {
    const file = files.find((f) => f.includes("presetLayout_12"));
    if (!file) return;
    const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
    const slide = pres.slides[0];
    const shape = slide.shapes.find((s) => s.fontScalePct === 90)!;
    shape.paragraphs = (shape.paragraphs ?? []).map((p) => ({ ...p, runs: p.runs.map((r) => ({ ...r, text: "改过的字" })) }));
    shape.textDirty = true;
    pres.markDirty(slide);
    const out = OpcPackage.from(writePptx(pres));
    const xml = out.text("ppt/slides/slide1.xml")!;
    expect(xml).toContain('fontScale="90000"');
    const back = Presentation.openBytes(writePptx(pres));
    const again = back.slides[0].shapes.find((s) => s.paragraphs?.some((p) => p.runs.some((r) => r.text === "改过的字")));
    expect(again?.fontScalePct).toBe(90);
  });

  it.skipIf(!files.length)("组里的形状也带着自己的 raw：改字落进组内那个 sp", () => {
    const file = files.find((f) => f.includes("presetLayout_14"));
    if (!file) return;
    const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
    const group = pres.slides[0].shapes.find((s) => s.kind === "group")!;
    const member = group.children!.find((c) => c.name === "空心弧 10")!;
    // `p:grpSp` 的成员直接挂在它下面（没有自己的 `p:spTree`）；拿不到成员 raw 时
    // patchShape 会静默跳过，改字落不进盘上那枚 sp。
    expect(member.raw).toBeTruthy();
    member.paragraphs = [{ level: 0, runs: [{ text: "组内改字" }] }];
    member.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const out = OpcPackage.from(writePptx(pres));
    expect(out.text("ppt/slides/slide1.xml")).toContain("组内改字");
    const back = Presentation.openBytes(writePptx(pres));
    const again = back.slides[0].shapes.find((s) => s.kind === "group")!.children!.find((c) => c.name === "空心弧 10")!;
    expect(again.paragraphs?.[0]?.runs[0]?.text).toBe("组内改字");
  });

  it("fontScale 与 lnSpcReduction 都进渲染，模型里那句 200% 行距不被改写", () => {
    const pres = deckWith(
      fitShape(
        "缩放框",
        '<a:normAutofit fontScale="80000" lnSpcReduction="20000"/>',
        '<a:p><a:pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="2000"/><a:t>缩放文字</a:t></a:r></a:p>',
      ),
    );
    const shape = pres.slides[0].shapes.find((s) => s.name === "缩放框")!;
    expect(shape.fontScalePct).toBe(80);
    expect(shape.lineSpacingReductionPct).toBe(20);
    expect(shape.paragraphs?.[0]?.lineSpacingPct).toBe(200);
    const html = renderSlideHtml(pres, 0);
    const slice = shapeSlice(html, pres.slides[0].shapes.indexOf(shape));
    expect(slice).toContain("font-size:16pt");
    expect(slice).not.toContain("data-shrink");
    // 1.3（单行）× 200% × (1 − 20%) = 2.08
    expect(slice).toContain("line-height:2.08");
    expect(shape.paragraphs?.[0]?.lineSpacingPct).toBe(200);
  });

  it("只写了 normAutofit 没写 fontScale：文件没说比例，标记文本体交给页面现场量", () => {
    const pres = deckWith(
      fitShape("待缩框", "<a:normAutofit/>", '<a:p><a:r><a:rPr sz="2000"/><a:t>原字号</a:t></a:r></a:p>'),
    );
    const shape = pres.slides[0].shapes.find((s) => s.name === "待缩框")!;
    expect(shape.fontScalePct).toBeUndefined();
    // 渲染层不改字号，只留标记；量高与二分在 DOM 里做（src/ui/text-fit.ts）。
    expect(shape.autoFit).toBe("shrink");
    const html = renderSlideHtml(pres, 0);
    const slice = shapeSlice(html, pres.slides[0].shapes.indexOf(shape));
    expect(slice).toContain("font-size:20pt");
    expect(slice).toContain('class="shape-text" data-shrink="1"');

    // 改字只重写段落，bodyPr 原样留盘：裸写的 normAutofit 不能因为编辑就丢。
    shape.paragraphs = [{ level: 0, runs: [{ text: "已改" }] }];
    shape.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const xml = OpcPackage.from(writePptx(pres)).text("ppt/slides/slide1.xml")!;
    expect(xml).toContain("<a:normAutofit/>");
    const back = Presentation.openBytes(writePptx(pres));
    expect(back.slides[0].shapes.find((s) => s.name === "待缩框")?.autoFit).toBe("shrink");
  });

  it("改字保存：bodyPr 里的两个属性原样带回", () => {
    const pres = deckWith(
      fitShape(
        "缩放框",
        '<a:normAutofit fontScale="80000" lnSpcReduction="20000"/>',
        '<a:p><a:r><a:rPr sz="2000"/><a:t>缩放文字</a:t></a:r></a:p>',
      ),
    );
    const shape = pres.slides[0].shapes.find((s) => s.name === "缩放框")!;
    shape.paragraphs = [{ level: 0, runs: [{ text: "已改" }] }];
    shape.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const out = OpcPackage.from(writePptx(pres));
    const xml = out.text("ppt/slides/slide1.xml")!;
    expect(xml).toContain('fontScale="80000"');
    expect(xml).toContain('lnSpcReduction="20000"');
    const back = Presentation.openBytes(writePptx(pres));
    const again = back.slides[0].shapes.find((s) => s.name === "缩放框")!;
    expect(again.fontScalePct).toBe(80);
    expect(again.lineSpacingReductionPct).toBe(20);
    expect(again.paragraphs?.[0]?.runs[0]?.text).toBe("已改");
  });
});
