import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { parseOrderedXml, rawAttrs, rawChildren, rawFind, tagLocalOf } from "../src/core/opc.js";
import { Presentation, type Shape } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";

/**
 * 预设几何的调整值 `a:prstGeom/a:avLst`。语料里幻灯片上只有 13 枚形状带调整值：
 * 空心弧 8（都在一份模板的分组里）、对角圆角矩形 3、梯形 2；此前这三种几何落到渲染时
 * 都是空的 `clip-path:`（等于没画），梯形与圆角矩形还被画成方块。这里逐枚对表，再用
 * 合成件钉住三种几何的画面数字。
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

function flatten(shapes: Shape[]): Shape[] {
  return shapes.flatMap((s) => (s.kind === "group" ? [s, ...flatten(s.children ?? [])] : [s]));
}

const nameOf = (entry: any): string | undefined => {
  const nv =
    rawFind(rawChildren(entry), "nvSpPr") ??
    rawFind(rawChildren(entry), "nvCxnSpPr") ??
    rawFind(rawChildren(entry), "nvGrpSpPr");
  const id = nv ? rawFind(rawChildren(nv), "cNvPr") : undefined;
  return id ? rawAttrs(id)["@_name"] : undefined;
};

/** Every `sp`/`cxnSp` under a group or slide tree, with the `a:gd` values it states. */
function adjusted(entries: any[], out: { name: string; geom: string; adj: Record<string, number> }[] = []): typeof out {
  for (const entry of entries) {
    const local = tagLocalOf(entry);
    if (local === "grpSp") {
      adjusted(rawChildren(entry), out);
      continue;
    }
    if (local !== "sp" && local !== "cxnSp") continue;
    const spPr = rawFind(rawChildren(entry), "spPr");
    const geom = spPr ? rawFind(rawChildren(spPr), "prstGeom") : undefined;
    const avLst = geom ? rawFind(rawChildren(geom), "avLst") : undefined;
    if (!geom || !avLst) continue;
    const adj: Record<string, number> = {};
    for (const gd of rawChildren(avLst).filter((c) => tagLocalOf(c) === "gd")) {
      const a = rawAttrs(gd);
      const m = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(a["@_fmla"] ?? "");
      if (m) adj[a["@_name"]!] = Number(m[1]);
    }
    if (!Object.keys(adj).length) continue;
    out.push({ name: nameOf(entry)!, geom: rawAttrs(geom)["@_prst"]!, adj });
  }
  return out;
}

const pct = (value: number, total: number): string =>
  total > 0 ? `${((value / total) * 100).toFixed(2)}%` : "0%";

describe("预设几何调整值（a:avLst）", () => {
  const files = decks();

  it.skipIf(!files.length)("13 枚调整值逐枚读得准，且三种几何都画出了形状", () => {
    let shapes = 0;
    const presets: Record<string, number> = {};
    const filesWith = new Set<string>();
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const zip = unzipSync(bytes);
      const pres = Presentation.openBytes(bytes);
      const slidePaths = Object.keys(zip)
        .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort();
      pres.slides.forEach((slide, i) => {
        const tree = rawFind(parseOrderedXml(dec.decode(zip[slidePaths[i]]!)), "spTree");
        if (!tree) return;
        const byName = new Map(flatten(slide.shapes).map((s) => [s.name, s]));
        const html = renderSlideHtml(pres, i);
        for (const entry of adjusted(rawChildren(tree))) {
          const shape = byName.get(entry.name);
          expect(shape, `${entry.name} 未建模`).toBeDefined();
          expect(shape!.geometry, `${entry.name} 几何名`).toBe(entry.geom);
          expect(shape!.adjustments, `${entry.name} 调整值`).toEqual(entry.adj);
          shapes++;
          presets[entry.geom] = (presets[entry.geom] ?? 0) + 1;
          filesWith.add(file);
          const ss = Math.min(shape!.width, shape!.height);
          if (entry.geom === "round2DiagRect") {
            const r1 = (ss * Math.min(entry.adj.adj1 ?? 16667, 50000)) / 100000;
            const r2 = (ss * Math.min(entry.adj.adj2 ?? 0, 50000)) / 100000;
            expect(html, `${entry.name} 圆角`).toContain(
              `border-radius:${r1.toFixed(2)}px 0 ${r2.toFixed(2)}px 0`,
            );
          } else if (entry.geom === "trapezoid") {
            const cap = (50000 * shape!.height) / ss;
            const inset = (ss * Math.min(entry.adj.adj ?? 25000, cap)) / 100000;
            expect(html, `${entry.name} 梯形`).toContain(
              `clip-path:polygon(${pct(inset, shape!.width)} 0,${pct(shape!.width - inset, shape!.width)} 0,100% 100%,0 100%)`,
            );
          } else if (entry.geom === "blockArc") {
            const start = (entry.adj.adj1 ?? 16200000) / 60000;
            const rx = shape!.width / 2;
            const ry = shape!.height / 2;
            const first = `${pct(rx + rx * Math.cos((start * Math.PI) / 180), shape!.width)} ${pct(ry + ry * Math.sin((start * Math.PI) / 180), shape!.height)}`;
            expect(html, `${entry.name} 弧起点`).toContain(`clip-path:polygon(${first},`);
          }
        }
        // 空值 `clip-path:` 会让整条声明作废，形状退回方块——这正是改之前的画面。
        expect(html).not.toContain("clip-path:;");
      });
    }
    expect(shapes).toBe(13);
    expect(presets).toEqual({ blockArc: 8, round2DiagRect: 3, trapezoid: 2 });
    expect(filesWith.size).toBe(3);
  });
});

describe("三种预设几何落到画面的数字", () => {
  it("梯形按 maxAdj 截到一半宽，对角圆角矩形两只角各按自己的半径", () => {
    const html = renderSlideHtml(PresetDeck(), 0);
    // 400×100 的梯形，adj 57534 超过 maxAdj=50000·h/ss，截到 50000 → 顶边内缩 50px。
    expect(html).toContain("clip-path:polygon(12.50% 0,87.50% 0,100% 100%,0 100%)");
    // 200×100 的 square 边 = 100：adj1=16667 → 16.67px，adj2=1826 → 1.83px。
    expect(html).toContain("border-radius:16.67px 0 1.83px 0");
  });

  it("空心弧是 0°→90° 的半环，内半径 = 外半径 − 短边 25%", () => {
    const html = renderSlideHtml(PresetDeck(), 0);
    const polygon = /clip-path:polygon\((100\.00% 50\.00%,[^)]*)\)/.exec(html)?.[1];
    expect(polygon, "空心弧没有画出来").toBeDefined();
    // 90° 处外沿落在盒子底边中点；内沿（半径 50）回到 75% / 50%。
    expect(polygon).toContain("50.00% 100.00%");
    expect(polygon).toContain("75.00% 50.00%");
  });

  it("认不出的预设不写空的 clip-path", () => {
    const html = renderSlideHtml(PresetDeck(), 0);
    expect(html).not.toContain("clip-path:;");
    expect(html).not.toContain("clip-path:polygon()");
    expect(html).toContain("background:#00B0F0");
  });
});

/** 在 makePptx 的空白页上补梯形/对角圆角矩形/空心弧/未知预设四枚形状。 */
function PresetDeck(): Presentation {
  const zip = unzipSync(makePptx());
  const path = "ppt/slides/slide1.xml";
  const shape = (id: number, name: string, prst: string, cx: number, cy: number, gds: string): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
<a:prstGeom prst="${prst}"><a:avLst>${gds}</a:avLst></a:prstGeom>
<a:solidFill><a:srgbClr val="00B0F0"/></a:solidFill></p:spPr></p:sp>`;
  const xml = [
    shape(40, "梯形 1", "trapezoid", 3810000, 952500, `<a:gd name="adj" fmla="val 57534"/>`),
    shape(
      41,
      "对角圆角矩形 1",
      "round2DiagRect",
      1905000,
      952500,
      `<a:gd name="adj1" fmla="val 16667"/><a:gd name="adj2" fmla="val 1826"/>`,
    ),
    shape(
      42,
      "空心弧 1",
      "blockArc",
      1905000,
      1905000,
      `<a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 5400000"/><a:gd name="adj3" fmla="val 25000"/>`,
    ),
    shape(43, "云形 1", "cloud", 1905000, 1905000, ""),
  ].join("");
  zip[path] = strToU8(dec.decode(zip[path]!).replace("</p:spTree>", `${xml}</p:spTree>`));
  return Presentation.openBytes(zipSync(zip, { level: 6 }));
}
