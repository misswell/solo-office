import { describe, expect, it } from "vitest";
import type { Paragraph } from "../src/core/docx.js";
import { WordDocument, carriedShape } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { text, OpcPackage } from "../src/core/opc.js";
import { makeChainDocx } from "./helpers/make-paragraph.js";
import { renderWord } from "../src/core/render.js";
import { parseTabStops } from "../src/ui/doc-layout.js";

/**
 * `w:pPr/w:tabs` (§17.4.19). The shapes these tests transcribe are the ones the authorized 46 份 write
 * (`.scratch/tabs-audience.py`, `.scratch/tabs-inside-table.py`): 370 blocks, every one of them under a `w:pPr`
 * (230 in `word/document.xml`, 90 in `word/styles.xml`, 12 in the bands), 106 of them naming two stops or more;
 * 520 stops, **all with `w:pos`**, spelled in four ways only — `left` 304, `right` 152, `center` 58, `clear` 6 —
 * with `w:leader` on 94 of them and every leader saying `dot`; `decimal` and `bar` 0, so the model has no field
 * for them. 0 of the 370 sit in a table cell. `w:defaultTabStop` is written by all 46 (43 × 420, 3 × 720).
 */

const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const stopsOf = (bytes: Uint8Array, at = 0) => paragraphs(open(bytes))[at]!.tabStops;

describe("docx 制表位读取", () => {
  it("段落自己写的每一枚都读出来，按缇排序，w:leader 留作文件自己的token", () => {
    const docx = makeChainDocx([], [
      { pPr: '<w:tabs><w:tab w:val="right" w:pos="8306"/><w:tab w:val="left" w:pos="720" w:leader="dot"/></w:tabs>', text: "两枚" },
    ]);
    expect(stopsOf(docx)).toEqual([
      { posTwips: 720, val: "left", leader: "dot" },
      { posTwips: 8306, val: "right", leader: undefined },
    ]);
  });

  it("空的 w:tabs 是「什么都没说」，不是「一处不落」", () => {
    expect(stopsOf(makeChainDocx([], [{ pPr: "<w:tabs/>", text: "空块" }]))).toBeUndefined();
  });

  it("w:val 漏写按 ST_TabVal 的默认值 left；w:pos 漏写的那一枚不收", () => {
    const docx = makeChainDocx([], [{ pPr: '<w:tabs><w:tab w:pos="4153"/><w:tab w:val="right"/></w:tabs>', text: "两种漏写" }]);
    expect(stopsOf(docx)).toEqual([{ posTwips: 4153, val: "left", leader: undefined }]);
  });

  it("样式链与段落逐**位置**合：段落那一枚只顶掉同位置的，样式其余的原样留着", () => {
    // nearest-wins over the whole element would lose 4153 here — and 106 of the 370 blocks write two or more at once.
    const docx = makeChainDocx(
      [{ id: "T", pPr: '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="center" w:pos="4153"/></w:tabs>' }],
      [{ pStyle: "T", pPr: '<w:tabs><w:tab w:val="right" w:pos="4153"/></w:tabs>', text: "顶掉一枚" }],
    );
    expect(stopsOf(docx)).toEqual([
      { posTwips: 720, val: "left", leader: undefined },
      { posTwips: 4153, val: "right", leader: undefined },
    ]);
  });

  it("两级样式各写各的：往上的那一跳不把下面那跳的清单整批抹掉", () => {
    const docx = makeChainDocx(
      [
        { id: "Base", pPr: '<w:tabs><w:tab w:val="left" w:pos="2160"/></w:tabs>' },
        { id: "Child", basedOn: "Base", pPr: '<w:tabs><w:tab w:val="right" w:pos="9000"/></w:tabs>' },
      ],
      [{ pStyle: "Child", text: "链上两级" }],
    );
    expect(stopsOf(docx)).toEqual([
      { posTwips: 2160, val: "left", leader: undefined },
      { posTwips: 9000, val: "right", leader: undefined },
    ]);
  });

  it("段落写 w:val=\"clear\" 就只办事不占位：抹掉链上同位置那一枚，别的位置不动", () => {
    // 授权 46 份 6 枚 `clear`，枚枚带 `w:pos`（`.scratch/tabs-audience.py`）。
    const docx = makeChainDocx(
      [{ id: "T", pPr: '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="8306"/></w:tabs>' }],
      [{ pStyle: "T", pPr: '<w:tabs><w:tab w:val="clear" w:pos="8306"/></w:tabs>', text: "取消那枚" }],
    );
    expect(stopsOf(docx)).toEqual([{ posTwips: 720, val: "left", leader: undefined }]);
  });

  it("拆段的新那一段带走上一段算好的制表位（`MODEL_ONLY_PARA_KEYS` 那本账）", () => {
    const p = paragraphs(open(makeChainDocx([], [{ pPr: '<w:tabs><w:tab w:val="right" w:pos="8306"/></w:tabs>', text: "有位的段" }])))[0]!;
    expect(carriedShape(p).tabStops).toEqual([{ posTwips: 8306, val: "right", leader: undefined }]);
  });
});

describe("docx 制表位的写端与画端", () => {
  /** One keystroke, then save: the path a document is actually put through. */
  const edited = (bytes: Uint8Array): string => {
    const doc = open(bytes);
    const p = paragraphs(doc)[0]!;
    p.inlines[0] = { ...p.inlines[0]!, text: "打过字的" };
    return part(writeDocx(doc));
  };

  it("改一次正文，w:tabs 按原件那一块带回来，不多写一枚", () => {
    const bytes = makeChainDocx([], [{ pPr: '<w:tabs><w:tab w:val="right" w:pos="8306"/></w:tabs>', text: "整段自述" }]);
    const saved = edited(bytes);
    expect(saved.match(/<w:tabs>/g)?.length).toBe(1);
    expect(saved).toContain('<w:tab w:val="right" w:pos="8306"/>');
    expect(saved.match(/<w:tab /g)?.length).toBe(1);
  });

  it("`data-tab-stops` 把算好的位置交给量的人；没有位的段落不带这一句", () => {
    const html = renderWord(open(makeChainDocx([], [
      { pPr: '<w:tabs><w:tab w:val="right" w:pos="8306"/><w:tab w:val="center" w:pos="4153"/></w:tabs>', text: "有位" },
      { text: "没有位", rPr: "" },
    ])));
    expect(html).toContain('data-tab-stops="4153:c,8306:r"');
    expect(html.match(/data-tab-stops/g)?.length).toBe(1);
    // 画端只认这一句字面：读回来就是同一张位置表。
    expect(parseTabStops("4153:c,8306:r")).toEqual([
      { posTwips: 4153, val: "center" },
      { posTwips: 8306, val: "right" },
    ]);
    expect(parseTabStops("8306,,:x")).toEqual([]);
  });
});
