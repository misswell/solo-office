import { describe, expect, it } from "vitest";
import { WordDocument, twipsToPx } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { PLACE_FOOTER, makeBandPlaceDocx } from "./helpers/make-band-place.js";

/**
 * 带子里那一段落在哪（#211）。
 *
 * 受众是量出来的，不是想出来的：授权 46 份的 `word/header*.xml` / `word/footer*.xml` 共 76 段，`w:jc` 自己说
 * `center` 的 12 枚（6 个全路径件 / 3 份模板名）、从样式链借来的 52 枚（`left` 48 / `both` 2 / `center` 2）、
 * 谁都没说的 12 段；`w:pPr/w:framePr` 6 枚（4 个全路径件 / 2 份模板名），枚枚都是 `w:wrap="around"
 * w:vAnchor="text" w:hAnchor="margin" w:y="1"` 加 `w:xAlign` right ×4 / center ×2，坐在带 PAGE 域那一段上
 * （`.scratch/band-para-align-audience.py`）。借来的那一档与自家页眉样式写明的 `w:jc w:val="left"` 相撞，
 * 所以「居中」不是重述，是纸面与文件分歧的地方。
 */

const openBand = () => {
  const doc = WordDocument.openBytes(makeBandPlaceDocx());
  const band = doc.sectionBand("footer", 0);
  expect(band, "夹具的那一节要显式绑上页脚").toBeDefined();
  return band!;
};

/** The `<div class="hf-p" …>` the page puts a paragraph in, whole. */
const bandLine = (html: string, text: string): string => {
  const m = new RegExp(`<div class="hf-p"[^>]*>[^<]*${text}</div>`).exec(html);
  return m?.[0] ?? "";
};

/** The same by index — the line holding a field is not text, so its own words cannot find it. */
const bandLineAt = (html: string, i: number): string => (html.match(/<div class="hf-p"[^>]*>[\s\S]*?<\/div>/g) ?? [])[i] ?? "";

describe("docx 页眉页脚：段自己说的落点读进模型", () => {
  it("每段一格，格数与段数一样", () => {
    const band = openBand();
    expect(band.paragraphs.length).toBe(8);
    expect(band.place?.length).toBe(band.paragraphs.length);
    expect(band.editable).toBe(true);
  });

  it("自己说的居中、链上借的居中、自己说的盖住样式说的", () => {
    const band = openBand();
    expect(band.place?.[0]?.align).toBe("center");
    expect(band.place?.[0]?.frame).toBeUndefined();
    expect(band.place?.[1]?.align).toBe("center");
    expect(band.place?.[2]?.align).toBe("right");
  });

  it("框的六句话各留各的原话，锚在纸上的也照原话记下", () => {
    const band = openBand();
    expect(band.place?.[3]?.frame).toEqual({ hAnchor: "margin", vAnchor: "text", xAlign: "right", yTwips: 1 });
    expect(band.place?.[3]?.align).toBe("left");
    expect(band.place?.[4]?.frame).toEqual({ hAnchor: "margin", vAnchor: "text", xAlign: "center", yTwips: 1 });
    expect(band.place?.[5]?.frame).toEqual({ hAnchor: "page", vAnchor: "page", xAlign: "right", yTwips: 1 });
    expect(band.place?.[6]?.frame).toEqual({ hAnchor: "text", xTwips: 108, yTwips: 1 });
  });

  it("谁都没说的那一段不替它编一句", () => {
    const band = openBand();
    expect(band.place?.[7]).toBeUndefined();
  });
});

describe("docx 页眉页脚：段自己说的落点画上纸", () => {
  const html = renderWord(WordDocument.openBytes(makeBandPlaceDocx()));

  it("把居中与靠右写成这一格的对齐", () => {
    expect(bandLine(html, "自己说居中")).toBe(`<div class="hf-p" style="text-align:center">自己说居中</div>`);
    expect(bandLine(html, "链上借来居中")).toBe(`<div class="hf-p" style="text-align:center">链上借来居中</div>`);
    expect(bandLine(html, "自己说的盖住样式说的")).toBe(`<div class="hf-p" style="text-align:right">自己说的盖住样式说的</div>`);
  });

  it("框摆到右沿：收成墨宽再推过去，偏移有落脚处", () => {
    const right = bandLineAt(html, 3);
    expect(right).toContain(`style="width:fit-content;margin-left:auto;position:relative;top:${twipsToPx(1)}px;`);
    expect(right).toContain('data-field="PAGE"');
    // 段自己没说对齐，框把它推到右沿不等于替框里那几行说话：借来的 `left` 仍是这一格的答案。
    expect(right).toContain("text-align:left");
  });

  it("框摆到中沿：两侧都留 auto", () => {
    expect(bandLine(html, "摆到中沿")).toBe(
      `<div class="hf-p" style="width:fit-content;margin-left:auto;margin-right:auto;position:relative;top:${twipsToPx(1)}px">摆到中沿</div>`,
    );
  });

  it("锚在纸上的一格横竖都不落，画成常规流的那一段", () => {
    // 授权 46 份的 `hAnchor="page"` / `vAnchor="page"` 各写 0 枚，纸沿那只尺比这一栏宽出一整个页边距，不猜。
    expect(bandLine(html, "锚在纸上")).toBe(`<div class="hf-p">锚在纸上</div>`);
  });

  it("只说偏移不说对齐的，按声明的缇偏", () => {
    expect(bandLine(html, "只说偏移不说对齐")).toBe(
      `<div class="hf-p" style="position:relative;left:${twipsToPx(108)}px;top:${twipsToPx(1)}px">只说偏移不说对齐</div>`,
    );
  });

  it("没说过落点的那一段一个属性都不多", () => {
    expect(bandLine(html, "什么都没说")).toBe(`<div class="hf-p">什么都没说</div>`);
  });
});

describe("docx 页眉页脚：落点只上行，不动字节", () => {
  it("不开任何编辑，保存把页脚件原样交回", () => {
    const out = writeDocx(WordDocument.openBytes(makeBandPlaceDocx()));
    expect(text(OpcPackage.from(out).bytes("word/footer1.xml")!)).toBe(PLACE_FOOTER);
  });

  it("不开任何编辑，带子不报脏", () => {
    // 落点是**读出来画上去**的，不该让一份没人动过的件在保存清单里排成待写（#180 同一条规矩）。
    const doc = WordDocument.openBytes(makeBandPlaceDocx());
    expect(doc.headerFooterParts().map((b) => b.dirty)).toEqual([false]);
    writeDocx(doc);
    expect(doc.headerFooterParts().map((b) => b.dirty)).toEqual([false]);
  });

  it("保存后再读，落点一格都不挪", () => {
    const doc = WordDocument.openBytes(makeBandPlaceDocx());
    const again = WordDocument.openBytes(writeDocx(doc)).sectionBand("footer", 0);
    expect(again?.place?.map((p) => [p?.align, p?.frame?.xAlign, p?.frame?.hAnchor])).toEqual(
      doc.sectionBand("footer", 0)?.place?.map((p) => [p?.align, p?.frame?.xAlign, p?.frame?.hAnchor]),
    );
  });
});
