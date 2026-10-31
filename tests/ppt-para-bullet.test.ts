import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync, zipSync, strToU8 } from "fflate";
import { OpcPackage, parseOrderedXml, rawAttrs, rawChildren, rawFind, tagLocalOf } from "../src/core/opc.js";
import { Presentation } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { writePptx } from "../src/core/write-pptx.js";
import { makePptx } from "./helpers/make.js";

/**
 * 段落自己的 `a:pPr@marL/@indent` 与 `<a:buNone/>`。语料里只有一份件（presetLayout_68）在
 * 段落上写了这两样：3 枚文本占位符的段落是 `marL="0" indent="0" algn="ctr"` + `a:buNone`，
 * 而这一页的正文一级样式带 `•` 与 228600/-228600 EMU —— 不看段落这一句，画面会多出一颗
 * 符号并把文字右推 24px。这里按原始 XML 对表，再验“拿掉 buNone 符号就回来”的差分与写回。
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
const deck68 = files.find((f) => f.includes("presetLayout_68"));
const open68 = (bytes?: Uint8Array): Presentation =>
  Presentation.openBytes(bytes ?? new Uint8Array(readFileSync(deck68!)));

describe("段落的 buNone 与 marL/indent", () => {
  it.skipIf(!deck68)("段落级 buNone / marL / indent 只在那一份件里，读数与原始 XML 一致", () => {
    let withNone = 0;
    let withMarL = 0;
    const seen = new Set<string>();
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        for (const shape of slide.shapes) {
          for (const p of shape.paragraphs ?? []) {
            if (p.noBullet) {
              withNone++;
              seen.add(file);
            }
            if (p.marLEmu !== undefined) withMarL++;
          }
        }
      }
    }
    expect(withNone).toBe(3);
    expect(withMarL).toBe(3);
    expect([...seen]).toEqual([deck68]);

    const pres = open68();
    const shape = pres.slides[0].shapes.find((s) => s.name === "文本占位符 8")!;
    expect(shape.paragraphs![0]).toMatchObject({ level: 0, align: "center", noBullet: true, marLEmu: 0, indentEmu: 0 });
    const zip = unzipSync(new Uint8Array(readFileSync(deck68!)));
    const tree = parseOrderedXml(dec.decode(zip["ppt/slides/slide1.xml"]!));
    const pPr = (function walk(nodes: any[]): any[] {
      return nodes.flatMap((entry) => {
        if (tagLocalOf(entry) === "pPr") return [entry];
        return walk(rawChildren(entry));
      });
    })(tree);
    const own = pPr.filter((n) => rawFind(rawChildren(n), "buNone"));
    expect(own).toHaveLength(3);
    for (const n of own) {
      expect(rawAttrs(n)).toMatchObject({ "@_marL": "0", "@_indent": "0", "@_algn": "ctr" });
    }
  });

  it.skipIf(!deck68)("画面上：正文一级样式带 •，这三段照样不画符号、不缩进", () => {
    const pres = open68();
    // 这一页继承来的正文一级样式确实是会画符号的那一套，否则本闸门什么也没验。
    expect(pres.bodyLevels![0]).toMatchObject({ bullet: "•", marLEmu: 228600, indentEmu: -228600 });
    const html = renderSlideHtml(pres, 0);
    const paras = [...html.matchAll(/<div class="sp-p"[^>]*>.*?<\/div>/g)].map((m) => m[0]);
    const body = paras.filter((p) => p.includes("请在此处添加文本"));
    expect(body).toHaveLength(3);
    for (const p of body) {
      expect(p).toContain("text-align:center");
      expect(p).toContain("padding-left:0px");
      expect(p).toContain("text-indent:0px");
      expect(p).not.toContain('class="bullet"');
    }
  });

  it.skipIf(!deck68)("拿掉那三枚 buNone，符号与 24px 缩进就回到画面（差分）", () => {
    const zip = unzipSync(new Uint8Array(readFileSync(deck68!)));
    const path = "ppt/slides/slide1.xml";
    zip[path] = strToU8(
      dec
        .decode(zip[path]!)
        .replace(/<a:pPr marL="0" indent="0" algn="ctr"><a:buNone\/><\/a:pPr>/g, '<a:pPr algn="ctr"/>')
        .replace(/<a:pPr marL="0" indent="0" algn="ctr"><a:buNone\/>/g, '<a:pPr algn="ctr">')
        .replace(/<a:pPr marL="0" indent="0" algn="ctr">/g, '<a:pPr algn="ctr">'),
    );
    const html = renderSlideHtml(Presentation.openBytes(zipSync(zip, { level: 6 })), 0);
    const body = [...html.matchAll(/<div class="sp-p"[^>]*>.*?<\/div>/g)].map((m) => m[0]).filter((p) => p.includes("请在此处添加文本"));
    expect(body).toHaveLength(3);
    for (const p of body) {
      expect(p).toContain("padding-left:24px");
      expect(p).toContain('class="bullet">•</span>');
    }
  });

  it.skipIf(!deck68)("改字写回后 buNone 与 marL/indent 仍在段落上", () => {
    const pres = open68();
    const shape = pres.slides[0].shapes.find((s) => s.name === "文本占位符 8")!;
    shape.paragraphs![0].runs = [{ text: "改写过的正文" }];
    shape.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const out = writePptx(pres);
    const xml = OpcPackage.from(out).text("ppt/slides/slide1.xml")!;
    expect(xml).toContain('<a:pPr marL="0" indent="0" algn="ctr"><a:buNone/></a:pPr>');
    const back = Presentation.openBytes(out).slides[0].shapes.find((s) => s.name === "文本占位符 8")!;
    expect(back.paragraphs![0]).toMatchObject({ noBullet: true, marLEmu: 0, indentEmu: 0, align: "center" });
    expect(back.paragraphs![0].runs.map((r) => r.text).join("")).toBe("改写过的正文");
  });

  it("合成件：没有 buNone 的同一段仍按样式画符号，写了 buNone 就压掉", () => {
    // makePptx 的正文一级段落自带 buChar，换成继承样式的那句才有意思：第一段只写 buNone，
    // 第二段什么也不写 —— 两段都落在会画 • 的正文样式上。
    const zip = unzipSync(makePptx());
    const path = "ppt/slides/slide1.xml";
    zip[path] = strToU8(
      dec.decode(zip[path]!).replace(
        '<a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>收入达成 108%</a:t></a:r></a:p>',
        '<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr sz="1800"/><a:t>压掉符号</a:t></a:r></a:p>' +
          '<a:p><a:pPr/><a:r><a:rPr sz="1800"/><a:t>保留符号</a:t></a:r></a:p>',
      ),
    );
    const pres = Presentation.openBytes(zipSync(zip, { level: 6 }));
    expect(pres.bodyLevels![0].bullet).toBeTruthy();
    const paras = [...renderSlideHtml(pres, 0).matchAll(/<div class="sp-p"[^>]*>.*?<\/div>/g)].map((m) => m[0]);
    const plain = paras.find((p) => p.includes("压掉符号"))!;
    const bulleted = paras.find((p) => p.includes("保留符号"))!;
    expect(plain).not.toContain('class="bullet"');
    expect(bulleted).toContain(`class="bullet">${pres.bodyLevels![0].bullet}</span>`);
  });
});
