import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { parseOrderedXml, rawAttrs, rawChildren, rawFind, tagLocalOf } from "../src/core/opc.js";
import { Presentation } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { writePptx } from "../src/core/write-pptx.js";
import { makePptx } from "./helpers/make.js";

/**
 * `a:bodyPr` 是文字盒那一圈：四边内缩（EMU）、竖向锚点、是否折行。语料里 51 枚
 * 幻灯片形状写了自己的 bodyPr（ctr 49、b 2，1 枚 wrap="none"），28 枚版式占位符
 * 又写了 94 个属性值；此前这些一个都没读，画面按发明出来的 4px/8px 内边距把字
 * 一律居中。读取、逐属性阶梯（形状 → 版式 → 母版）与画面这一整层都补在此处。
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

const topShape = (entry: any): string | undefined => {
  const nv = rawFind(rawChildren(entry), "nvSpPr");
  const id = nv ? rawFind(rawChildren(nv), "cNvPr") : undefined;
  return id ? rawAttrs(id)["@_name"] : undefined;
};

const bodyPrOf = (entry: any): Record<string, string> | undefined => {
  const txBody = rawFind(rawChildren(entry), "txBody");
  const bodyPr = txBody ? rawFind(rawChildren(txBody), "bodyPr") : undefined;
  return bodyPr ? rawAttrs(bodyPr) : undefined;
};

const shapeTextBody = (attrs: Record<string, string>) => {
  const emu = (name: string): number | undefined => {
    const v = attrs[`@_${name}`];
    return v !== undefined && Number.isFinite(Number(v)) ? Number(v) : undefined;
  };
  return {
    lInsEmu: emu("lIns"),
    tInsEmu: emu("tIns"),
    rInsEmu: emu("rIns"),
    bInsEmu: emu("bIns"),
    anchor: attrs["@_anchor"],
    noWrap: attrs["@_wrap"] === "none" ? true : undefined,
  };
};

describe("形状与占位符的文字盒 a:bodyPr", () => {
  const files = decks();

  it.skipIf(!files.length)("幻灯片形状自己的内缩/锚点/折行读得准", () => {
    let compared = 0;
    const anchors: Record<string, number> = {};
    const insets: Record<string, number> = {};
    let wrapped = 0;
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
        for (const entry of tree ? rawChildren(tree) : []) {
          if (tagLocalOf(entry) !== "sp") continue;
          const attrs = bodyPrOf(entry);
          if (!attrs) continue;
          const stated = ["@_anchor", "@_lIns", "@_tIns", "@_rIns", "@_bIns", "@_wrap"].filter((k) => attrs[k] !== undefined);
          if (!stated.length) continue;
          const name = topShape(entry)!;
          const shape = byName.get(name);
          expect(shape, `${file} / ${name}`).toBeDefined();
          expect(shape!.textBody, `${file} / ${name}`).toEqual(shapeTextBody(attrs));
          compared++;
          if (attrs["@_anchor"] !== undefined) {
            const a = attrs["@_anchor"]!;
            anchors[a] = (anchors[a] ?? 0) + 1;
          }
          if (attrs["@_wrap"] === "none") wrapped++;
          for (const k of ["lIns", "tIns", "rIns", "bIns"]) {
            if (attrs[`@_${k}`] === undefined) continue;
            const key = `${k}=${attrs[`@_${k}`]}`;
            insets[key] = (insets[key] ?? 0) + 1;
          }
        }
      });
    }
    // 实测：51 枚写了属性，其中 anchor=ctr 49、anchor=b 2、wrap=none 1；
    // 四边内缩 39422/19711 一对、68580/34290 一对、91440/45720 共 22 枚。
    expect(compared).toBe(51);
    expect(anchors).toEqual({ ctr: 49, b: 2 });
    expect(wrapped).toBe(1);
    expect(insets).toEqual({
      "lIns=39422": 1,
      "tIns=19711": 1,
      "rIns=39422": 1,
      "bIns=19711": 1,
      "lIns=68580": 1,
      "tIns=34290": 1,
      "rIns=68580": 1,
      "bIns=34290": 1,
      "lIns=91440": 22,
      "tIns=45720": 22,
      "rIns=91440": 22,
      "bIns=45720": 22,
    });
  });

  it.skipIf(!files.length)("版式占位符的 bodyPr 进模型，母版只补它没写的那些属性", () => {
    const seen = new Set<string>();
    let compared = 0;
    const anchors: Record<string, number> = {};
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const zip = unzipSync(bytes);
      const pres = Presentation.openBytes(bytes);
      for (const slide of pres.slides) {
        if (!slide.layoutPath) continue;
        const xml = zip[slide.layoutPath] ? dec.decode(zip[slide.layoutPath]!) : "";
        const tree = rawFind(parseOrderedXml(xml), "spTree");
        for (const ph of slide.layoutPlaceholders ?? []) {
          const key = `${file}|${slide.layoutPath}|${ph.idx ?? "-"}|${ph.type ?? "obj"}`;
          if (seen.has(key)) continue;
          seen.add(key);
          const entry = (tree ? rawChildren(tree) : []).find((e) => {
            if (tagLocalOf(e) !== "sp") return false;
            const nv = rawFind(rawChildren(e), "nvSpPr");
            const marker = nv ? rawFind(rawChildren(nv), "ph") : undefined;
            if (!marker) return false;
            const a = rawAttrs(marker);
            const idx = a["@_idx"] !== undefined ? Number(a["@_idx"]) : undefined;
            return idx === ph.idx && (a["@_type"] ?? "obj") === (ph.type ?? "obj");
          });
          if (!entry) continue;
          const attrs = bodyPrOf(entry);
          if (!attrs) continue;
          const stated = shapeTextBody(attrs);
          const got = ph.textBody ?? {};
          for (const [k, v] of Object.entries(stated)) {
            if (v === undefined) continue;
            expect((got as Record<string, unknown>)[k], `${file} ${slide.layoutPath} ${ph.idx}/${ph.type} ${k}`).toBe(v);
            if (k === "anchor") {
              const a = String(v);
              anchors[a] = (anchors[a] ?? 0) + 1;
            }
            compared++;
          }
        }
      }
    }
    // 实测：28 枚版式占位符写了 94 个值，anchor=b 13、anchor=t 1。
    expect(compared).toBe(94);
    expect(anchors).toEqual({ b: 13, t: 1 });
  });

  it("形状 → 版式 → 母版逐属性落成画面上的内边距与竖向锚点", () => {
    const pres = PreviewDeck();
    const html = renderSlideHtml(pres, 0);
    // 标题：版式/母版都没有它的占位符 → WPS 默认 0.1in/0.05in 内缩、顶对齐。
    expect(html).toContain("padding:4.8px 9.6px 4.8px 9.6px;justify-content:flex-start");
    // 正文：自己的 bodyPr 是空的，版式给 b + tIns，母版补其余三边。
    expect(html).toContain("padding:4px 9.6px 4.8px 9.6px;justify-content:flex-end");
    // 自选图形：自己的四边 + wrap="none" 全部生效。
    expect(html).toContain("padding:4px 2.7px 4px 10.7px;justify-content:center;white-space:nowrap");
  });

  it("改字保存后形状自己的 bodyPr 原样带回", () => {
    const pres = PreviewDeck();
    const custom = pres.slides[0].shapes.find((s) => s.name === "自定义")!;
    custom.paragraphs = [{ level: 0, runs: [{ text: "改过的字" }] }];
    custom.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const back = Presentation.openBytes(writePptx(pres));
    const reread = back.slides[0].shapes.find((s) => s.name === "自定义")!;
    expect(reread.textBody).toEqual({
      lInsEmu: 101600,
      tInsEmu: 38100,
      rInsEmu: 25400,
      bInsEmu: 38100,
      anchor: "ctr",
      noWrap: true,
    });
    expect(reread.paragraphs?.[0].runs[0].text).toBe("改过的字");
  });
});

/** makePptx 的版式原本没有占位符；给正文占位符补一条 b/tIns，母版补四边。 */
function PreviewDeck(): Presentation {
  const zip = unzipSync(makePptx());
  const text = (path: string): string => dec.decode(zip[path]!);
  const replace = (path: string, from: string, to: string): void => {
    zip[path] = strToU8(text(path).replace(from, to));
  };
  const layoutPh = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="正文占位符 1"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="635000" y="1800000"/><a:ext cx="5200000" cy="3600000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr anchor="b" tIns="38100"/><a:p/></p:txBody></p:sp>`;
  const masterPh = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="正文占位符 1"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="635000" y="1800000"/><a:ext cx="5200000" cy="3600000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/><a:p/></p:txBody></p:sp>`;
  const custom = `<p:sp><p:nvSpPr><p:cNvPr id="20" name="自定义"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="8000000" y="300000"/><a:ext cx="2000000" cy="800000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr anchor="ctr" lIns="101600" tIns="38100" rIns="25400" bIns="38100" wrap="none"/><a:p><a:r><a:t>自定义文本</a:t></a:r></a:p></p:txBody></p:sp>`;
  replace("ppt/slideLayouts/slideLayout1.xml", "</p:spTree>", `${layoutPh}</p:spTree>`);
  replace("ppt/slideMasters/slideMaster1.xml", "</p:spTree>", `${masterPh}</p:spTree>`);
  replace("ppt/slides/slide1.xml", "</p:spTree>", `${custom}</p:spTree>`);
  return Presentation.openBytes(zipSync(zip, { level: 6 }));
}
