import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { TABS_FOOTER, TABS_STYLES, makeBandTabsDocx } from "./helpers/make-band-tabs.js";

/**
 * 带子里制表符的落点（#235）。
 *
 * 受众是量出来的：授权 46 份的 `word/header*.xml` / `word/footer*.xml` 里真正的 `w:r/w:tab` 只有 **2 枚**，
 * 去重后 **1 份件、1 枚**——`event plan.dotm` 页脚段#0：裸制表符、无文字，段自己不写位，位从 `footer` 样式
 * （styleId 4）借 `center 4153` + `right 8306`，两枚都不带 `w:leader`（`.scratch/band-tab-chars-audience.py`；
 * §153 只报过总数，这份普查报位置）。此前带子按 `pre-wrap` 的浏览器默认画制表符，位整个丢了。
 *
 * 夹具五格各有出处（`tests/helpers/make-band-tabs.ts` 的头注）：借位裸 tab、自位带前导符、自己的顶掉链上的、
 * `clear` 花掉一格、无制表符的对照。前导符那格钉的是机制——语料里带子侧前导符 0 受众。
 */

const openDoc = () => WordDocument.openBytes(makeBandTabsDocx());

const bandLineAt = (html: string, i: number): string =>
  (html.match(/<div class="hf-p"[^>]*>[\s\S]*?<\/div>/g) ?? [])[i] ?? "";

describe("docx 页眉页脚：制表位读进带子的模型", () => {
  it("链上借来的那对位整份进 place，段文字里制表符还是那个字面", () => {
    const doc = openDoc();
    const band = doc.sectionBand("footer", 0)!;
    expect(band.editable).toBe(true);
    expect(band.paragraphs[0]).toBe("\t");
    expect(band.place?.[0]?.tabStops).toEqual([
      { posTwips: 4153, val: "center" },
      { posTwips: 8306, val: "right" },
    ]);
  });

  it("段自己的位顶掉链上同位置的，另一档照借；clear 把那一格花掉", () => {
    const doc = openDoc();
    const band = doc.sectionBand("footer", 0)!;
    // ③：段自己说 left 4153 —— 链上 center 4153 被顶掉，right 8306 照旧。
    expect(band.place?.[2]?.tabStops).toEqual([
      { posTwips: 4153, val: "left" },
      { posTwips: 8306, val: "right" },
    ]);
    // ④：clear 只花 4153，剩下的只有 8306。
    expect(band.place?.[3]?.tabStops).toEqual([{ posTwips: 8306, val: "right" }]);
  });

  it("链上无位、段上也无位的段不挂这一句", () => {
    const doc = openDoc();
    const band = doc.sectionBand("footer", 0)!;
    // ② 只有段自己的位；⑤ 什么位都没有。
    expect(band.place?.[1]?.tabStops).toEqual([{ posTwips: 2880, val: "center", leader: "dot" }]);
    expect(band.place?.[4]?.tabStops).toBeUndefined();
  });
});

describe("docx 页眉页脚：制表符包成 span.tab、位写上 data-tab-stops", () => {
  const html = renderWord(openDoc());

  it("唯一真实受众那一格：位上墙、制表符进 span，字面还在", () => {
    const line = bandLineAt(html, 0);
    expect(line).toContain('data-tab-stops="4153:c,8306:r"');
    expect(line).toContain('<span class="tab">&#9;</span>');
    // span 装着字面 \t（写成实体 &#9;），段的 textContent 才还是模型里那一串（读回走 textContent）。
    const inner = line.replace(/^<div class="hf-p"[^>]*>/, "").replace(/<\/div>$/, "");
    expect(inner.replace(/<[^>]+>/g, "").replace(/&#9;/g, "\t")).toBe("\t");
  });

  it("自位那格带前导符 token；两侧的字原样留在 span 外", () => {
    const line = bandLineAt(html, 1);
    expect(line).toContain('data-tab-stops="2880:c:dot"');
    expect(line).toBe(
      '<div class="hf-p" data-tab-stops="2880:c:dot">前<span class="tab">&#9;</span>后</div>',
    );
  });

  it("没有制表符的段一个属性都不多、一枚 span 都不挂", () => {
    const line = bandLineAt(html, 4);
    expect(line).toBe('<div class="hf-p" style="text-align:left">没有制表符</div>');
  });

  it("借位那格的位不在了（clear 花掉）就只写剩下的那枚", () => {
    const line = bandLineAt(html, 3);
    expect(line).toContain('data-tab-stops="8306:r"');
    expect(line).not.toContain("4153");
  });
});

describe("docx 页眉页脚：夹具本身说的与读者认得的同一句话", () => {
  it("footer 部件逐字面仍是写进去的那一份（位定义不该漏成文字）", () => {
    const doc = openDoc();
    const band = doc.sectionBand("footer", 0)!;
    // 段文字里不许出现「4153」——`hfParagraphText` 跳过 `w:pPr`，位定义不是文字。
    for (const p of band.paragraphs) expect(p).not.toContain("4153");
    expect(band.paragraphs.join("|")).toBe(
      ["\t", "前\t后", "\t盖", "\t清", "没有制表符"].join("|"),
    );
    // 件里那句 w:tabs 的字面（夹具源码里那一串）原样可寻——写回按字节走，读的这头别改字。
    expect(TABS_STYLES).toContain('<w:tab w:val="center" w:pos="4153"/>');
  });
});
