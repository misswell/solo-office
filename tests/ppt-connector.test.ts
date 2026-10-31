import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync, zipSync, strToU8 } from "fflate";
import { parseOrderedXml, rawAttrs, rawChildren, rawFind, tagLocalOf } from "../src/core/opc.js";
import { Presentation } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";

/**
 * 直线连接符 `a:prstGeom@prst="straightConnector1"` 与它的 `a:ln` 端饰。语料里 3 枚全在
 * presetLayout_13：两条斜的直接箭头（盒 102×129、94×128），一条几乎水平的（盒 164×0.18）；
 * 端饰是 stealth 箭头（`w="lg" len="lg"`），另有一枚 `type="none"`。此前它们按普通形状画成
 * 一个 4px 实线盒子——斜线成了方框、箭头没了。这里按原始 XML 逐枚对表，再用合成件钉住折线。
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

const files = decks();
const num = (v: number): string => String(Math.round(v * 100) / 100);

function connectors(nodes: any[]): any[] {
  const out: any[] = [];
  for (const entry of nodes) {
    if (tagLocalOf(entry) === "cxnSp") out.push(entry);
    out.push(...connectors(rawChildren(entry)));
  }
  return out;
}

function endOf(node: any): { type?: string; width?: string; length?: string } | undefined {
  if (!node) return undefined;
  const a = rawAttrs(node);
  const end = { type: a["@_type"], width: a["@_w"], length: a["@_len"] };
  return end.type !== undefined || end.width !== undefined || end.length !== undefined ? end : undefined;
}

describe("直线连接符", () => {
  it.skipIf(!files.length)("语料里 3 枚 straightConnector1 的线、颜色、端饰与原始 XML 一致", () => {
    let total = 0;
    let flipV = 0;
    let stealthTails = 0;
    let noneTails = 0;
    const widths: number[] = [];
    const deckNames = new Set<string>();
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const zip = unzipSync(bytes);
      const pres = Presentation.openBytes(bytes);
      const slidePaths = Object.keys(zip)
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort();
      pres.slides.forEach((slide, i) => {
        const tree = rawFind(parseOrderedXml(dec.decode(zip[slidePaths[i]!]!)), "spTree");
        const byName = new Map(slide.shapes.map((s) => [s.name, s]));
        const html = renderSlideHtml(pres, i);
        for (const entry of tree ? connectors([tree]) : []) {
          const geom = rawFind(rawChildren(entry), "prstGeom");
          if (rawAttrs(geom ?? {})["@_prst"] !== "straightConnector1") continue;
          const id = rawFind(rawChildren(rawFind(rawChildren(entry), "nvCxnSpPr")!), "cNvPr");
          const name = rawAttrs(id ?? {})["@_name"]!;
          const shape = byName.get(name)!;
          const ln = rawFind(rawChildren(rawFind(rawChildren(entry), "spPr")!), "ln")!;
          expect(shape.geometry).toBe("straightConnector1");
          expect(shape.headEnd, `${name} headEnd`).toEqual(endOf(rawFind(rawChildren(ln), "headEnd")));
          expect(shape.tailEnd, `${name} tailEnd`).toEqual(endOf(rawFind(rawChildren(ln), "tailEnd")));
          total++;
          deckNames.add(file);
          widths.push(shape.lineWidth!);
          if (shape.flipV) flipV++;
          if (shape.tailEnd?.type === "stealth") stealthTails++;
          if (shape.tailEnd?.type === "none") noneTails++;

          const box = html.slice(html.indexOf(`data-idx="${slide.shapes.indexOf(shape)}"`));
          const div = box.slice(0, box.indexOf("</div>"));
          expect(div, `${name} svg`).toContain("<svg");
          expect(div, `${name} box border`).not.toContain("border:4px");
          const x1 = shape.flipH ? shape.width : 0;
          const y1 = shape.flipV ? shape.height : 0;
          const x2 = shape.flipH ? 0 : shape.width;
          const y2 = shape.flipV ? 0 : shape.height;
          expect(div, `${name} line`).toContain(`<line x1="${num(x1)}" y1="${num(y1)}" x2="${num(x2)}" y2="${num(y2)}"`);
          expect(div, `${name} stroke`).toContain(`stroke="${shape.lineColor}" stroke-width="4"`);
          const ends = [shape.headEnd, shape.tailEnd].filter((e) => e?.type && e.type !== "none").length;
          expect(div.split("<polygon").length - 1, `${name} arrowheads`).toBe(ends);
        }
      });
    }
    expect(total).toBe(3);
    expect([...deckNames].map((f) => f.split("__").pop())).toEqual(["presetLayout_13.pptx"]);
    expect(widths).toEqual([4, 4, 4]);
    expect(flipV).toBe(2);
    expect(stealthTails).toBe(2);
    expect(noneTails).toBe(1);
  });
});

describe("连接符落到画面", () => {
  it("斜线按盒对角线画、flipV 换对角，stealth 箭头画成折向箭尖的四边形", () => {
    const html = renderSlideHtml(ConnectorDeck(), 0);
    // 952500×952500 EMU = 100×100px；flipV 于是从 (0,100) 走到 (100,0)。
    expect(html).toContain('<line x1="0" y1="100" x2="100" y2="0" stroke="#FF0000" stroke-width="2"');
    // sysDot 落到 2px 线上是 2/4 的点划。
    expect(html).toContain('stroke-dasharray="2 4"');
    // head 是 triangle（sm: 2× 线宽 = 4px 长、4px 宽），箭底在箭尖之后；tail 是 stealth（lg: 5× = 10px）。
    expect(html).toContain('<polygon points="0,100 -1.41,104.24 -4.24,101.41" fill="#FF0000"');
    expect(html).toContain('<polygon points="100,0 96.46,10.61 95.76,4.24 89.39,3.54" fill="#FF0000"');
    expect(html).not.toContain("<polyline");
  });

  it("没有端饰的实线连接符只有一条线，不画箭头也不画盒子边", () => {
    const html = renderSlideHtml(ConnectorDeck(), 0);
    // 空白页原有 5 个形状，两条连接符追加成 idx 5（箭头）与 idx 6（无箭头）。
    const div = html.slice(html.indexOf('data-idx="6"'), html.indexOf("</div>", html.indexOf('data-idx="6"')));
    expect(div).toContain('<line x1="0" y1="0" x2="200" y2="0" stroke="#00B050" stroke-width="2"');
    expect(div).not.toContain("<polygon");
    expect(div).not.toContain("<circle");
    expect(div).not.toContain("stroke-dasharray");
    expect(div).not.toContain("border:");
  });
});

/** 在 makePptx 的空白页上补两条直线连接符：一枚带 triangle/stealth 端饰的点线，一枚无端饰实线。 */
function ConnectorDeck(): Presentation {
  const zip = unzipSync(makePptx());
  const path = "ppt/slides/slide1.xml";
  const arrow = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="40" name="箭头 A"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDot"/><a:headEnd type="triangle" w="sm" len="sm"/><a:tailEnd type="stealth" w="lg" len="lg"/></a:ln></p:spPr></p:cxnSp>`;
  const plain = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="41" name="无箭头"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="0" y="2000000"/><a:ext cx="1905000" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
<a:ln w="19050"><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`;
  zip[path] = strToU8(dec.decode(zip[path]!).replace("</p:spTree>", `${arrow}${plain}</p:spTree>`));
  return Presentation.openBytes(zipSync(zip, { level: 6 }));
}
