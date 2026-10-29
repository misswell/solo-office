import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import type { Paragraph, Table } from "../src/core/docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { pack } from "./helpers/make-containers.js";
import { FLOAT_PR, SPEC_PR, TEXT_PR } from "./helpers/make-table-float.js";

/**
 * 浮动表（`w:tblPr/w:tblpPr`，§17.4.58）：表坐在哪由文件写在**纸**上，不由流排出来。
 *
 * 受众（授权 46 件，`.scratch/cantsplit-audience.py`）：**2 枚 / 2 件**——`shareholding agreement.dotm` 的
 * zh_CN 与 zh_HK 各一张签章表，两句逐字相同：`w:vertAnchor="page" w:horzAnchor="page" w:tblpX="1117"
 * w:tblpY="10912"`（左右各留 180 缇给文字）。纸 A4 11906×16838、左 1080、上 1440，所以那张表的左上角落在
 * 离纸左沿 74.47px、离纸上沿 727.47px 的地方；它比栏宽（9746 缇）还宽 568 缇，右边探出页边距。
 *
 * 这一页判的是**纸上有没有那一句**：`render.ts` 把缇原样盖成 `data-tblp-*`，落点由铺排换成像素
 * （`doc-layout.ts` 的 `applySheetGaps`，真 Chrome 里量在 `.scratch/tblppr-dom.ts`）。不能画的两档
 * （`w:tblpXSpec`/`w:tblpYSpec`、`text` 锚）留一句话都不盖：那一档授权 0 枚，按猜出来的坐标摆才是替文件说话。
 */
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const tagOf = (bytes: Uint8Array): string => /<table\b[^>]*>/.exec(renderWord(open(bytes)))?.[0] ?? "";
const attr = (tag: string, name: string): string | undefined => new RegExp(`${name}="([^"]*)"`).exec(tag)?.[1];
const partOf = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);

/** 一张 2×2 的表；`pr` 就是那枚 `w:tblPr` 的内层（`w:tblpPr` 该住在它最前面，紧跟着 `w:tblStyle`）。 */
const oneTable = (pr: string): Uint8Array =>
  pack(
    `<w:tbl><w:tblPr><w:tblStyle w:val="8"/>${pr}<w:tblW w:w="10314" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="5157"/><w:gridCol w:w="5157"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="5157" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>甲方（签章）</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="5157" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>____年__月__日</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="5157" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>公司其他股东签章</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="5157" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>乙方（签章）</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
  );

/** 授权那 2 枚的字面（与夹具共用一份，见 `tests/helpers/make-table-float.ts`）。 */
const REAL = FLOAT_PR;

describe("w:tblPr/w:tblpPr 浮动表", () => {
  it("carries the file's twips onto the page, anchor and all", () => {
    const tag = tagOf(oneTable(REAL));
    expect(attr(tag, "data-tblp-x"), "tblpX in twips, as the file spells it").toBe("1117");
    expect(attr(tag, "data-tblp-y")).toBe("10912");
    expect(attr(tag, "data-tblp-vert")).toBe("page");
    expect(attr(tag, "data-tblp-horz")).toBe("page");
  });

  it("leaves the drawing to the layout pass: no position is invented on the page", () => {
    // 缇换成像素要等铺排知道这一枚在哪张纸上（`doc-layout.ts` 的 `applySheetGaps`），渲染器只在纸上写下
    // 「文件说 1117 缇」这句话；谁在这一层写死 `left:`，换一张纸、换一节边距就都错了。
    const tag = tagOf(oneTable(REAL));
    expect(tag).not.toContain("position");
    expect(tag).not.toContain("left:");
  });

  it("does not draw an anchor it cannot name: w:tblpXSpec/w:tblpYSpec", () => {
    // 「居中/靠右」那几句说的是相对谁居中，没有一处数可以量；授权 46 件里 0 枚。
    const tag = tagOf(oneTable(SPEC_PR));
    expect(tag, "the table stays a table, just not a placed one").toContain("doc-table");
    expect(tag).not.toContain("data-tblp");
  });

  it("does not draw a text anchor either", () => {
    // `vertAnchor="text"` 量的是一枚段落标记，表已经离开流、那枚标记在哪儿我们自己说了不算。
    const tag = tagOf(oneTable(TEXT_PR));
    expect(tag).not.toContain("data-tblp");
  });

  it("a table the file did not float gets no such words", () => {
    const tag = tagOf(oneTable(""));
    expect(tag).toContain("doc-table");
    expect(tag).not.toContain("data-tblp");
  });

  it("writes the element back once, as the file spelled it", () => {
    const out = partOf(open(oneTable(REAL)));
    expect((out.match(/<w:tblpPr\b/g) ?? []).length, "one sentence, not the source plus a copy").toBe(1);
    expect(out).toContain(REAL);
    // 孩子序也照 `CT_TblPrBase`：`w:tblStyle` 之后、`w:tblW` 之前。
    expect(out).toMatch(/<w:tblStyle w:val="8"\/><w:tblpPr [^/]*\/><w:tblW /);
  });

  it("keeps the float through a keystroke inside one of its cells", () => {
    const doc = open(oneTable(REAL));
    const table = doc.blocks.find((b): b is Table => b.kind === "table")!;
    const run = table.rows[0]![0]!.paragraphs.find((b): b is Paragraph => b.kind === "paragraph")!.inlines[0]!;
    run.text += "打字";
    const out = partOf(doc);
    expect(out).toContain("甲方（签章）打字");
    expect(out, "the anchor is not a property a keystroke may drop").toContain(REAL);
    expect((out.match(/<w:tblpPr\b/g) ?? []).length).toBe(1);
  });

  it("does not invent one on a table that had none", () => {
    const doc = open(oneTable(""));
    const table = doc.blocks.find((b): b is Table => b.kind === "table")!;
    table.rows[0]![1]!.paragraphs.find((b): b is Paragraph => b.kind === "paragraph")!.inlines[0]!.text = "改过";
    const out = partOf(doc);
    expect(out).toContain("改过");
    expect(out).not.toContain("w:tblpPr");
  });
});
