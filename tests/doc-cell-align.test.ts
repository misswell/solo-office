import { describe, expect, it } from "vitest";
import type { Paragraph, Table, TableCell } from "../src/core/docx.js";
import { WordDocument, syncCellVAlign, syncParaSlots } from "../src/core/docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { makeCellAlignDocx } from "./helpers/make-tables.js";

/**
 * 「对齐」那一组的两层：`w:tcPr/w:vAlign` 归格，`w:pPr/w:jc` 归格里的每一段。这一件量的是**件里那两层**——
 * 读得出、画得出、落笔时只动自己那一句、邻居按 `CT_PPr` / `CT_TcPr` 的孩子序原样带回。
 * 命令逐格逐段怎么走、回显什么时候灭，在 `.scratch/table-tools-dom.ts` 的 I 节按 ribbon 量。
 */
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;
const cellOf = (doc: WordDocument, row: number, col: number): TableCell => tableOf(doc).rows[row]![col]!;
const parasOf = (cell: TableCell): Paragraph[] => cell.paragraphs.filter((b): b is Paragraph => b.kind === "paragraph");
const rowOf = (doc: WordDocument, needle: string): string =>
  [...saved(doc).matchAll(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g)].map((m) => m[0]).find((r) => r.includes(needle)) ?? "";
const cellXmlOf = (doc: WordDocument, needle: string): string =>
  [...rowOf(doc, needle).matchAll(/<w:tc[\s>][\s\S]*?<\/w:tc>/g)].map((m) => m[0]).find((c) => c.includes(needle)) ?? "";

describe("docx 表格对齐两层：格说 w:vAlign，格里的段说 w:pPr/w:jc", () => {
  it("reads each layer where the file stated it and where it said nothing", () => {
    const doc = open(makeCellAlignDocx());
    const t = tableOf(doc);
    expect(t.rows[0]!.map((c) => c.valign)).toEqual([undefined, "center", "bottom"]);
    expect(parasOf(t.rows[0]![0]!).map((p) => p.align)).toEqual([undefined]);
    expect(parasOf(t.rows[0]![1]!).map((p) => p.align)).toEqual(["center"]);
    expect(parasOf(t.rows[0]![2]!).map((p) => p.align)).toEqual(["right"]);
    // 一格两段是这一族的常态：两句各说各的，「这一格怎么摆」就答不出唯一一句。
    expect(parasOf(t.rows[1]![0]!).map((p) => p.align)).toEqual(["center", undefined]);
    // 「说了 left」与「没说」是两句话，读回来要分得开——合并与回显都按这一条判。
    expect(parasOf(t.rows[1]![1]!).map((p) => p.align)).toEqual(["left"]);
    expect(parasOf(t.rows[1]![2]!).map((p) => p.align)).toEqual([undefined]);
  });

  it("draws both layers on the page", () => {
    const html = renderWord(open(makeCellAlignDocx()));
    expect(html).toContain("vertical-align:middle");
    expect(html).toContain("vertical-align:bottom");
    expect(html).toContain("text-align:center");
    expect(html).toContain("text-align:right");
    // 戊二那句 `left` 与甲一的「没说」画在同一个位置，但件里那句得是自己那句（写回时才知道谁说过）。
    expect(html).toContain('class="para" style="text-align:left"');
  });

  it("moves only the alignment child and keeps the neighbour in CT_PPr's order", () => {
    const doc = open(makeCellAlignDocx());
    const p = parasOf(cellOf(doc, 0, 2))[0]!; // 丙一：`w:spacing` 在前，`w:jc` 在后
    p.align = "center";
    syncParaSlots(p);
    expect(cellXmlOf(doc, "丙一")).toContain('<w:pPr><w:spacing w:after="240"/><w:jc w:val="center"/></w:pPr>');
    const q = parasOf(cellOf(doc, 1, 2))[0]!; // 己二：说过 spacing，没说 jc —— 新的一句排在它之后
    q.align = "right";
    syncParaSlots(q);
    expect(cellXmlOf(doc, "己二")).toContain('<w:pPr><w:spacing w:before="120"/><w:jc w:val="right"/></w:pPr>');
    const r = parasOf(cellOf(doc, 0, 0))[0]!; // 甲一：整段没有 `w:pPr`，落笔才现造那一枚
    r.align = "center";
    syncParaSlots(r);
    expect(cellXmlOf(doc, "甲一")).toContain('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>');
    // 三处落笔没有一处补出一句 `left`：那一档在件里的说法是「什么都不写」。戊二原来那句  left 一个字不动。
    expect((saved(doc).match(/<w:jc w:val="left"\/>/g) ?? []).length).toBe(1);
  });

  it("says 左 by taking the sentence back out, not by writing left", () => {
    const doc = open(makeCellAlignDocx());
    const p = parasOf(cellOf(doc, 1, 1))[0]!; // 戊二说过 `left`
    p.align = undefined;
    syncParaSlots(p);
    const xml = saved(doc);
    expect(xml).not.toContain('<w:jc w:val="left"');
    expect(cellXmlOf(doc, "戊二")).toContain("<w:p><w:r>");
    // 戊二那句收回之后与甲一（从来没说）在纸上落在同一个位置：左是默认档，两种说法画的是一件事。
    const html = renderWord(doc);
    expect((html.match(/text-align:left/g) ?? []).length).toBe(0);
    // 没说的那一段按「左」再问一次：一句也不写，盘上还是原样（回显亮着的左不是要补的一句新话）。
    const quiet = open(makeCellAlignDocx());
    const q = parasOf(cellOf(quiet, 0, 0))[0]!;
    q.align = undefined;
    syncParaSlots(q);
    expect(saved(quiet)).toBe(saved(open(makeCellAlignDocx())));
  });

  it("moves each paragraph of a cell on its own", () => {
    const doc = open(makeCellAlignDocx());
    const [first, second] = parasOf(cellOf(doc, 1, 0));
    second!.align = "center";
    syncParaSlots(second!);
    const cell = cellXmlOf(doc, "丁二");
    expect(first!.align).toBe("center");
    // 两段都说 center：件里是两句 `w:jc`，一句属于第一段——合并不是一句写在段标记上就够。
    expect((cell.match(/<w:jc w:val="center"\/>/g) ?? []).length).toBe(2);
    expect(cell).toContain('<w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>丁二第二段');
  });

  it("writes the cell's own vAlign after the width, leaving the row's word alone", () => {
    const doc = open(makeCellAlignDocx());
    syncCellVAlign(cellOf(doc, 0, 0), "bottom");
    const cell = cellXmlOf(doc, "甲一");
    expect(cell).toContain('<w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>');
    // 行那句 `w:trHeight` 与竖向对齐是两件事：改格不动行。
    expect(rowOf(doc, "甲一")).toContain('<w:trPr><w:trHeight w:val="1440" w:hRule="atLeast"/></w:trPr>');
    expect((saved(doc).match(/<w:vAlign\b/g) ?? []).length).toBe(3);
  });
});
