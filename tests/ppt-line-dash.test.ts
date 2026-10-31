import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { parseOrderedXml, rawAttrs, rawChildren, rawFind, tagLocalOf } from "../src/core/opc.js";
import { Presentation } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";

/**
 * 形状的 `a:ln/a:prstDash` 与 `a:xfrm@flipH/@flipV`。语料里幻灯片上有 16 枚形状写的是
 * `sysDash`（5 份件，含 4 条直线连接符），4 枚形状带 `flipV`；4 条直线连接符里 3 条外框
 * 宽为 0（竖线）、1 条高为 0。此前描边一律画实线，直线一律画成盒子顶边——竖线的盒子宽
 * 为 0，画出来是看不见的。这里按原始 XML 逐枚对表，再用合成件钉住画面。
 */
const CORPUS = ".scratch/wps-corpus";
const dec = new TextDecoder();

function decks(): string[] {
  if (!existsSync(CORPUS)) return [];
  return readdirSync(CORPUS)
    .sort()
    .filter((f) => /\.(pptx|potx|ppsx)$/i.test(f))
    .map((f) => join(CORPUS, f));
}

const nameOf = (entry: any): string | undefined => {
  const nv = rawFind(rawChildren(entry), "nvSpPr") ?? rawFind(rawChildren(entry), "nvCxnSpPr");
  const id = nv ? rawFind(rawChildren(nv), "cNvPr") : undefined;
  return id ? rawAttrs(id)["@_name"] : undefined;
};

describe("形状描边的虚线样式与直线方向", () => {
  const files = decks();

  it.skipIf(!files.length)("prstDash 值与 flipH/flipV 逐枚读得准", () => {
    let dashed = 0;
    const dashValues: Record<string, number> = {};
    const dashFiles = new Set<string>();
    let flippedV = 0;
    let flippedH = 0;
    let lineShapes = 0;
    let drawn = 0;
    const orientations: Record<string, number> = {};
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const zip = unzipSync(bytes);
      const pres = Presentation.openBytes(bytes);
      const slidePaths = Object.keys(zip)
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort();
      pres.slides.forEach((slide, i) => {
        const tree = rawFind(parseOrderedXml(dec.decode(zip[slidePaths[i]]!)), "spTree");
        const byName = new Map(slide.shapes.map((s) => [s.name, s]));
        const html = renderSlideHtml(pres, i);
        for (const entry of tree ? rawChildren(tree) : []) {
          const local = tagLocalOf(entry);
          if (local !== "sp" && local !== "cxnSp") continue;
          const name = nameOf(entry)!;
          const shape = byName.get(name)!;
          const spPr = rawFind(rawChildren(entry), "spPr");
          if (!spPr) continue;
          const xfrm = rawFind(rawChildren(spPr), "xfrm");
          if (xfrm) {
            const a = rawAttrs(xfrm);
            const wantH = a["@_flipH"] === "1";
            const wantV = a["@_flipV"] === "1";
            expect(shape.flipH ?? false, `${name} flipH`).toBe(wantH);
            expect(shape.flipV ?? false, `${name} flipV`).toBe(wantV);
            if (wantV) flippedV++;
            if (wantH) flippedH++;
          }
          const ln = rawFind(rawChildren(spPr), "ln");
          const dash = ln ? rawFind(rawChildren(ln), "prstDash") : undefined;
          if (!dash) continue;
          const value = rawAttrs(dash)["@_val"]!;
          expect(shape.dash, `${name} dash`).toBe(value);
          dashed++;
          dashValues[value] = (dashValues[value] ?? 0) + 1;
          dashFiles.add(file);
          // 3 枚版式占位符的 `a:ln` 没写颜色（线色要另行继承），没有描边可画。
          if (shape.lineColor) {
            expect(html, `${name} dash`).toContain("dashed");
            drawn++;
          }
          const geom = rawFind(rawChildren(spPr), "prstGeom");
          if (!geom || rawAttrs(geom)["@_prst"] !== "line") continue;
          lineShapes++;
          const orientation = shape.height === 0 ? "horizontal" : shape.width === 0 ? "vertical" : "diagonal";
          orientations[orientation] = (orientations[orientation] ?? 0) + 1;
          const slope = (shape.flipV ? -1 : 1) * (shape.flipH ? -1 : 1);
          const angle = Math.round((Math.atan2(shape.height, shape.width) * 180 * slope * 100) / Math.PI) / 100;
          expect(html, `${name} angle`).toContain(`rotate(${angle}deg)`);
        }
      });
    }
    expect(dashed).toBe(16);
    expect(dashValues).toEqual({ sysDash: 16 });
    expect(dashFiles.size).toBe(5);
    expect(flippedV).toBe(4);
    expect(flippedH).toBe(0);
    expect(lineShapes).toBe(4);
    expect(drawn).toBe(13);
    // 竖线（宽 0）此前画成盒子顶边、长度 0，画面上根本看不见；现在按 90° 摆。
    expect(orientations).toEqual({ horizontal: 1, vertical: 3 });
  });
});

describe("直线与描边样式落到画面", () => {
  it("flipV 的斜线按外框对角线旋转，虚线画成 dashed", () => {
    const html = renderSlideHtml(LineDeck(), 0);
    // 2000000×1000000 EMU = 209.97×104.99px，对角线 234.76px，仰角 26.57°。
    expect(html).toContain("width:234.75");
    expect(html).toMatch(/rotate\(-26\.5\d*deg\)/);
    expect(html).toContain("dashed #FF0000");
  });

  it("没有 flip 的水平线不旋转，sysDot 画成 dotted，solid 仍画实线", () => {
    const html = renderSlideHtml(LineDeck(), 0);
    expect(html).toContain("rotate(0deg)");
    expect(html).toContain("dotted #00B050");
    expect(html).toContain("solid #0070C0");
  });
});

/** 在 makePptx 的空白页上补三条线（斜线/点线/实线）。 */
function LineDeck(): Presentation {
  const zip = unzipSync(makePptx());
  const path = "ppt/slides/slide1.xml";
  const diagonal = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="30" name="直线连接符 1"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm flipV="1"><a:off x="1000000" y="2000000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>
<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></p:spPr></p:cxnSp>`;
  const horizontal = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="31" name="直线连接符 2"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="4000000"/><a:ext cx="3000000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>
<a:ln w="12700"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill><a:prstDash val="sysDot"/></a:ln></p:spPr></p:cxnSp>`;
  const box = `<p:sp><p:nvSpPr><p:cNvPr id="32" name="矩形 1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="5000000" y="4000000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:ln w="12700"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill><a:prstDash val="solid"/></a:ln></p:spPr></p:sp>`;
  zip[path] = strToU8(dec.decode(zip[path]!).replace("</p:spTree>", `${diagonal}${horizontal}${box}</p:spTree>`));
  return Presentation.openBytes(zipSync(zip, { level: 6 }));
}
