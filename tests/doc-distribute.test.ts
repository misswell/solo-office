import { describe, expect, it } from "vitest";
import type { Table, TableCell } from "../src/core/docx.js";
import { WordDocument, cellShapeOf, signCellRevision, twipsToPx } from "../src/core/docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { columnCells, distributeColumns, gridWidths, setColumnWidth } from "../src/core/doc-table-ops.js";
import { makeCellSizeDocx, makeDistributeDocx } from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
/** 三张：`odd` 除不尽、`even` 本来就等宽、`spanned` 中间那一列被跨列格盖住。 */
const trio = (): { doc: WordDocument; odd: Table; even: Table; spanned: Table } => {
  const doc = open(makeDistributeDocx());
  const [odd, even, spanned] = doc.blocks.filter((b) => b.kind === "table") as Table[];
  return { doc, odd: odd!, even: even!, spanned: spanned! };
};
const every = (t: Table): number[] => [0, 1, 2].slice(0, gridWidths(t).length).map((at) => at);

describe("docx「分布列」：等宽而和不变（w:gridCol 换数、那一列每格那句 w:tcW 跟着换、w:tblW 一个字不动）", () => {
  it("reads the three shapes the strike has to answer for", () => {
    const { odd, even, spanned } = trio();
    expect(gridWidths(odd)).toEqual([1000, 1100, 1102]);
    expect(gridWidths(even)).toEqual([1200, 1200]);
    expect(gridWidths(spanned)).toEqual([1500, 2000, 500]);
    // 中间那一列上下两行都被跨列格盖着：纸面上没有这一条边，所以这一列数不出可动的格。
    expect(spanned.rows[0]!.length).toBe(2);
    expect(spanned.rows[1]!.length).toBe(2);
  });

  it("splits the remainder to the leading column and keeps the grid's sum exactly", () => {
    const { doc, odd } = trio();
    const out = distributeColumns(odd, every(odd));
    expect(out.ok).toBe(true);
    expect(out.ok && out.cells.length).toBe(6);
    // 3,202 ÷ 3 = 1,067 余 1：最左那一列多拿一缇，于是和还是 3,202——授权 46 份里 44/58 张走的就是这一条。
    expect(gridWidths(odd)).toEqual([1068, 1067, 1067]);
    const xml = saved(doc);
    expect(xml).toContain('<w:gridCol w:w="1068"/><w:gridCol w:w="1067"/><w:gridCol w:w="1067"/>');
    expect(xml).toContain('<w:tblW w:w="3202" w:type="dxa"/>');
    expect((xml.match(/<w:tcW w:w="1067" w:type="dxa"\/>/g) ?? []).length).toBe(4);
    expect((xml.match(/<w:tcW w:w="1068" w:type="dxa"\/>/g) ?? []).length).toBe(2);
    // 那一枚 `w:tblLayout` 是表自己说的，均分只换宽：把它挤掉或挪位就是替件编形状。
    expect(xml).toContain('<w:tblW w:w="3202" w:type="dxa"/><w:tblLayout w:type="fixed"/>');
  });

  it("moves only the columns the selection touched, leaving the others' bytes alone", () => {
    const { doc, odd } = trio();
    const out = distributeColumns(odd, [0, 1]);
    expect(out.ok && out.cells.length).toBe(4);
    expect(gridWidths(odd)).toEqual([1050, 1050, 1102]);
    const xml = saved(doc);
    expect(xml).toContain('<w:gridCol w:w="1050"/><w:gridCol w:w="1050"/><w:gridCol w:w="1102"/>');
    expect(xml).toContain('<w:tblW w:w="3202" w:type="dxa"/>');
    expect((xml.match(/<w:tcW w:w="1102" w:type="dxa"\/>/g) ?? []).length).toBe(2);
  });

  it("refuses columns that already say one number, without spending a byte or an undo step", () => {
    const { doc, even } = trio();
    expect(distributeColumns(even, every(even))).toMatchObject({ ok: false, reason: "这几列已经是等宽" });
    expect(distributeColumns(even, [9])).toMatchObject({ ok: false, reason: "这一列不在表格里" });
    expect(gridWidths(even)).toEqual([1200, 1200]);
    expect(saved(doc)).toBe(saved(open(makeDistributeDocx())));
  });

  it("distributes between the edges the paper shows and never rewrites a spanning cell", () => {
    const { doc, spanned } = trio();
    const out = distributeColumns(spanned, every(spanned));
    expect(out.ok && out.cells.length).toBe(2);
    // 中间那一列一格都数不出（两行都从它头上跨过去）：摘出均分，它的网格数原样留着，和仍等于 4,000。
    expect(gridWidths(spanned)).toEqual([1000, 2000, 1000]);
    const xml = saved(doc);
    expect(xml).toContain('<w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="1000"/>');
    expect(xml).toContain('<w:tblW w:w="4000" w:type="dxa"/>');
    expect(xml).toContain('<w:tcW w:w="2500" w:type="dxa"/><w:gridSpan w:val="2"/>');
    expect(xml).toContain('<w:tcW w:w="3500" w:type="dxa"/><w:gridSpan w:val="2"/>');
  });

  it("refuses when the touched columns leave one visible edge, because a single column cannot be divided", () => {
    const doc = open(makeCellSizeDocx());
    const wide = doc.blocks.filter((b) => b.kind === "table")[0] as Table;
    // 那四列里只有最左那一列数得出宽度（其余三列被一枚跨三列的格盖着，见 #194 那份夹具）。
    expect(distributeColumns(wide, [0, 1, 2, 3])).toMatchObject({ ok: false, reason: "碰到的列里数得出宽度的只剩一列，均分要有两条边" });
    expect(gridWidths(wide)).toEqual([2000, 1500, 1500, 1500]);
    expect(saved(doc)).toBe(saved(open(makeCellSizeDocx())));
  });

  it("draws the equalised columns on the page at the new widths, with the table box unmoved", () => {
    const { doc, odd } = trio();
    const before = renderWord(doc);
    distributeColumns(odd, every(odd));
    const after = renderWord(doc);
    expect(before).toContain(`width:${twipsToPx(1102)}px`);
    expect(after).not.toContain(`width:${twipsToPx(1102)}px`);
    expect((after.match(new RegExp(`width:${twipsToPx(1068)}px`, "g")) ?? []).length).toBe(2);
    expect((after.match(new RegExp(`width:${twipsToPx(1067)}px`, "g")) ?? []).length).toBe(4);
    // 逐格画宽（`render.ts:918`）而表盒按 `w:tblW`：和没变，那只盒的宽度必须还是同一个数。
    expect(after).toContain(`width:${twipsToPx(3202)}px`);
    expect(before).toContain(`width:${twipsToPx(3202)}px`);
  });

  it("signs the whole strike as one revision and lets the refusal put every column back", () => {
    const press = (id: number): { doc: WordDocument; t: Table } => {
      const doc = open(makeDistributeDocx());
      const t = doc.blocks.filter((b) => b.kind === "table")[0] as Table;
      const backs = new Map<TableCell, ReturnType<typeof cellShapeOf>>();
      for (const at of every(t)) for (const c of columnCells(t, at)) backs.set(c, cellShapeOf(c));
      distributeColumns(t, every(t));
      for (const at of every(t)) for (const c of columnCells(t, at)) signCellRevision(c, { id, author: "顾峰" }, backs.get(c)!, undefined);
      return { doc, t };
    };
    const refused = press(501);
    expect(gridWidths(refused.t)).toEqual([1068, 1067, 1067]);
    // 一次落笔 = 一枚 `w:id`：三列六格各挂一枚 `w:tcPrChange`，写出的字节里那枚 id 只有一种。
    const signed = saved(refused.doc);
    expect((signed.match(/<w:tcPrChange\b/g) ?? []).length).toBe(6);
    expect(new Set(Array.from(signed.matchAll(/<w:tcPrChange w:id="(\d+)"/g), (m) => m[1]))).toEqual(new Set(["501"]));
    expect(refused.doc.resolveRevisions("reject", 501)).toBe(6);
    // 六枚 `w:tcPrChange` 共用一枚 `w:id` = 一次落笔；OOXML 没有网格的记录族，所以那一列的数按格摆回来。
    expect(gridWidths(refused.t)).toEqual([1000, 1100, 1102]);
    const xml = saved(refused.doc);
    expect(xml).toContain('<w:gridCol w:w="1000"/><w:gridCol w:w="1100"/><w:gridCol w:w="1102"/>');
    expect(xml).toContain('<w:tblW w:w="3202" w:type="dxa"/>');
    expect(xml).not.toContain("tcPrChange");

    const kept = press(502);
    expect(kept.doc.resolveRevisions("accept", 502)).toBe(6);
    expect(gridWidths(kept.t)).toEqual([1068, 1067, 1067]);
    expect(saved(kept.doc)).toContain('<w:tblW w:w="3202" w:type="dxa"/>');
  });

  it("keeps the earlier single-column strike answering for one column only", () => {
    // 「分布列」接在同一枚 `w:gridCol` 的 machinery 上，不能把「单元格宽度」那条按列的口径挤掉。
    const { doc, odd } = trio();
    expect(setColumnWidth(odd, 2, 900).ok).toBe(true);
    expect(gridWidths(odd)).toEqual([1000, 1100, 900]);
    expect(saved(doc)).toContain('<w:tblW w:w="3000" w:type="dxa"/>');
  });
});
