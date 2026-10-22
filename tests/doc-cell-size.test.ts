import { describe, expect, it } from "vitest";
import type { RowInfo, Table, TableCell } from "../src/core/docx.js";
import { WordDocument, rowShapeOf, signRowRevision, syncRowHeight } from "../src/core/docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { columnCells, gridWidths, restateColumnFromCells, setColumnWidth } from "../src/core/doc-table-ops.js";
import { makeCellSizeDocx, makeColumnWidthRevisionDocx } from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
/** 两件：`wide` 说「跨列格盖住整列」，`mixed` 说「一列里既有单格又有跨列/按百分比摆的格」。 */
const pair = (): { doc: WordDocument; wide: Table; mixed: Table } => {
  const doc = open(makeCellSizeDocx());
  const [wide, mixed] = doc.blocks.filter((b) => b.kind === "table") as Table[];
  return { doc, wide: wide!, mixed: mixed! };
};
const rowInfo = (t: Table, at: number): RowInfo => t.rows[at]![0]!.row!;
const widthsOf = (cells: TableCell[]): (number | undefined)[] => cells.map((c) => c.widthTwips);

describe("docx 单元格尺寸：高度归行（w:trHeight）、宽度归列（w:gridCol + w:tcW）", () => {
  it("reads both layers, and says nothing where the file says nothing", () => {
    const { wide, mixed } = pair();
    expect(rowInfo(wide, 0)).toMatchObject({ heightTwips: 1440, heightRule: "atLeast" });
    expect(rowInfo(wide, 1)).toMatchObject({ heightTwips: 600, heightRule: "atLeast" });
    expect(wide.rows[2]![0]!.row).toBeUndefined();
    expect(rowInfo(wide, 3)).toMatchObject({ heightTwips: 800, heightRule: "exact" });
    expect(gridWidths(wide)).toEqual([2000, 1500, 1500, 1500]);
    // 只占这一列的格才算这一列的数：跨三列那枚说的是三列之和。
    expect(widthsOf(columnCells(wide, 0))).toEqual([2000, 2000, 2000, 2000]);
    expect(columnCells(wide, 1)).toEqual([]);
    expect(widthsOf(columnCells(mixed, 1))).toEqual([1000]);
    expect(widthsOf(columnCells(mixed, 2))).toEqual([20]);
  });

  it("moves the whole column: every single-span cell, the grid, and the table's own width", () => {
    const { doc, wide } = pair();
    const out = setColumnWidth(wide, 0, 2400);
    expect(out.ok).toBe(true);
    expect(out.ok && out.cells.length).toBe(4);
    expect(gridWidths(wide)).toEqual([2400, 1500, 1500, 1500]);
    const xml = saved(doc);
    expect(widthsOf(columnCells(wide, 0))).toEqual([2400, 2400, 2400, 2400]);
    expect(xml).toContain('<w:gridCol w:w="2400"/><w:gridCol w:w="1500"/>');
    expect(xml).toContain('<w:tblW w:w="6900" w:type="dxa"/>');
    // 跨三列那一句一个字不动——把它的 4500 换成 2400 说的是「这一列加剩下三列」，那是编出来的数。
    expect(xml).toContain('<w:tcW w:w="4500" w:type="dxa"/><w:gridSpan w:val="3"/>');
    expect((xml.match(/<w:tcW w:w="2400" w:type="dxa"\/>/g) ?? []).length).toBe(4);
  });

  it("refuses a column that cannot state one width, and one that already says this number", () => {
    const { doc, wide, mixed } = pair();
    expect(setColumnWidth(wide, 4, 100)).toMatchObject({ ok: false, reason: "这一列不在表格里" });
    expect(setColumnWidth(wide, -1, 100).ok).toBe(false);
    // 整列被跨列格盖住 / 整列按百分比摆：这一击没有格可改，只挪网格会让表变宽而纸面一格没变。
    expect(setColumnWidth(wide, 2, 100)).toMatchObject({ ok: false });
    expect(setColumnWidth(mixed, 2, 100)).toMatchObject({ ok: false });
    expect(setColumnWidth(wide, 0, 2000)).toMatchObject({ ok: false, reason: "这一列已经是这个宽度" });
    expect(gridWidths(wide)).toEqual([2000, 1500, 1500, 1500]);
    expect(gridWidths(mixed)).toEqual([2000, 1000, 1000]);
    expect(saved(doc)).toBe(saved(open(makeCellSizeDocx())));
  });

  it("leaves a spanning and a percent-stated cell alone when the column moves", () => {
    const { doc, mixed } = pair();
    const out = setColumnWidth(mixed, 1, 1600);
    expect(out.ok && out.cells.length).toBe(1);
    expect(gridWidths(mixed)).toEqual([2000, 1600, 1000]);
    const xml = saved(doc);
    expect(xml).toContain('<w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/>');
    expect(xml).toContain('<w:tcW w:w="20" w:type="pct"/>');
    expect(xml).toContain('<w:tblW w:w="4600" w:type="dxa"/>');
  });

  it("writes the height into the row's carried children, keeping its neighbours and their order", () => {
    const { doc, wide } = pair();
    syncRowHeight(rowInfo(wide, 1), 1000, "atLeast");
    const xml = saved(doc);
    expect(xml).toContain('<w:trPr><w:cantSplit/><w:trHeight w:val="1000" w:hRule="atLeast"/></w:trPr>');
    expect(xml).not.toContain('w:val="600"');
    // 一件里同一枚孩子只有这一份：字段换了数、孩子没换，存盘会发现那一句还是件里的数（双写则读回来打架）。
    expect((xml.match(/<w:trHeight\b/g) ?? []).length).toBe(3);
  });

  it("spells the row's word when the row carried nothing, and keeps a rule the file stated", () => {
    const { doc, wide } = pair();
    const quiet: RowInfo = {};
    for (const cell of wide.rows[2]!) cell.row = quiet;
    syncRowHeight(quiet, 500, "atLeast");
    syncRowHeight(rowInfo(wide, 3), 900, "exact");
    const xml = saved(doc);
    expect(xml).toContain('<w:trPr><w:trHeight w:val="500" w:hRule="atLeast"/></w:trPr>');
    expect(xml).toContain('<w:trHeight w:val="900" w:hRule="exact"/>');
    expect((xml.match(/<w:trHeight\b/g) ?? []).length).toBe(4);
    // 行高是画出来的那一行的地板（#128）：落笔之后纸上真的差出这么多（500 缇 = 33.33px）。
    expect(renderWord(doc)).toContain('<tr style="height:33.3');
  });

  it("signs a row-height change once and lets the decision put the earlier number back", () => {
    const press = (id: number): { doc: WordDocument; t: Table } => {
      const doc = open(makeCellSizeDocx());
      const t = doc.blocks.filter((b) => b.kind === "table")[0] as Table;
      const info = rowInfo(t, 1);
      const was = rowShapeOf(info);
      const wasSlots = info.slots ?? [];
      syncRowHeight(info, 1000, "atLeast");
      signRowRevision(info, { id, author: "顾峰" }, was, wasSlots);
      return { doc, t };
    };
    const signed = press(301);
    const xml = saved(signed.doc);
    expect(xml).toContain('<w:trPr><w:cantSplit/><w:trHeight w:val="1000" w:hRule="atLeast"/>');
    expect(xml).toContain('<w:trPrChange w:id="301" w:author="顾峰"><w:trPr><w:cantSplit/><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr></w:trPrChange>');
    // 记录挂进带着孩子的那一行 = 一枚孩子；写出器认同一枚记录，不会再按字段补一份。
    expect((xml.match(/<w:trPrChange\b/g) ?? []).length).toBe(1);
    expect((xml.match(/<w:trHeight\b/g) ?? []).length).toBe(4);

    const refused = press(302);
    expect(refused.doc.resolveRevisions("reject", 302)).toBe(1);
    expect(rowInfo(refused.t, 1)).toMatchObject({ heightTwips: 600, heightRule: "atLeast", keepTogether: true });
    const refusedXml = saved(refused.doc);
    expect(refusedXml).toContain('<w:trPr><w:cantSplit/><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr>');
    expect(refusedXml).not.toContain("trPrChange");

    const kept = press(303);
    expect(kept.doc.resolveRevisions("accept", 303)).toBe(1);
    expect(rowInfo(kept.t, 1).heightTwips).toBe(1000);
    const keptXml = saved(kept.doc);
    expect(keptXml).toContain('<w:trPr><w:cantSplit/><w:trHeight w:val="1000" w:hRule="atLeast"/></w:trPr>');
    expect(keptXml).not.toContain("trPrChange");

    // 同一枚记录只裁一次：同一个人接着改第二击复用同一枚 `w:id`，两击在审核者眼里是一件事、一起退。
    const twice = open(makeCellSizeDocx());
    const t2 = twice.blocks.filter((b) => b.kind === "table")[0] as Table;
    const info2 = rowInfo(t2, 1);
    const first = signRowRevision(info2, { id: 304, author: "顾峰" }, rowShapeOf(info2), info2.slots ?? []);
    syncRowHeight(info2, 900, "atLeast");
    expect(signRowRevision(info2, { id: 305, author: "顾峰" }, rowShapeOf(info2), info2.slots ?? [])).toBe(first);
    expect((saved(twice).match(/<w:trPrChange\b/g) ?? []).length).toBe(1);
  });

  it("restates the grid column the cells sign, when a revised column is refused or shown as original", () => {
    const doc = open(makeColumnWidthRevisionDocx());
    const t = doc.blocks.find((b) => b.kind === "table") as Table;
    expect(gridWidths(t)).toEqual([1500, 2000]);
    expect((t.rows[0]![0]!.revs ?? []).map((r) => [r.tag, r.id, r.prev.widthTwips])).toEqual([["tcPrChange", 90, 2000]]);
    expect(doc.resolveRevisions("reject", 90)).toBe(2);
    // 两枚 `w:tcPrChange` 共用一个 id = 「改过的一列」；网格那一句没有记录族，所以按格摆回来。
    expect(gridWidths(t)).toEqual([2000, 2000]);
    const xml = saved(doc);
    expect(xml).toContain('<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>');
    expect(xml).toContain('<w:tblW w:w="4000" w:type="dxa"/>');
    expect(xml).not.toContain("tcPrChange");

    const original = open(makeColumnWidthRevisionDocx());
    const view = original.originalView(original.blocks).find((b) => b.kind === "table") as Table;
    expect(gridWidths(view)).toEqual([2000, 2000]);
    expect(gridWidths(original.blocks.find((b) => b.kind === "table") as Table)).toEqual([1500, 2000]);
    // 2000 缇 = 133.333…px、1500 缇 = 100px：原文态那一列的四格按格摆回来的数画，最终态两格。
    const originalHtml = renderWord(original, "original");
    const finalHtml = renderWord(original);
    expect((originalHtml.match(/width:133\.33333333333334px/g) ?? []).length).toBe(4);
    expect((finalHtml.match(/width:133\.33333333333334px/g) ?? []).length).toBe(2);
    expect(finalHtml).toContain("width:100px");
    expect(renderWord(original)).toBe(renderWord(open(makeColumnWidthRevisionDocx())));

    const accepted = open(makeColumnWidthRevisionDocx());
    expect(accepted.resolveRevisions("accept", 90)).toBe(2);
    expect(gridWidths(accepted.blocks.find((b) => b.kind === "table") as Table)).toEqual([1500, 2000]);
    expect(saved(accepted)).toContain('<w:tblW w:w="3500" w:type="dxa"/>');
  });

  it("does not invent a grid number the column cannot state", () => {
    const { doc, wide, mixed } = pair();
    // 网格与格本来就同一个数：不动，一个字节也不换。
    expect(restateColumnFromCells(mixed, 0)).toBe(false);
    // 那一列的格只被一枚跨列格盖着：说不出这一列多宽。
    expect(restateColumnFromCells(wide, 1)).toBe(false);
    // 同列各说各的（授权 46 份 0 张这样，32 张跨列表量不出这一条）也不猜。
    const cells = columnCells(wide, 0);
    cells[1]!.widthTwips = 1800;
    expect(restateColumnFromCells(wide, 0)).toBe(false);
    expect(gridWidths(wide)).toEqual([2000, 1500, 1500, 1500]);
    expect(saved(doc)).toBe(saved(open(makeCellSizeDocx())));
    // 全列说同一个数而网格说的不是它：这一列的网格跟着格回去，表盒跟着网格的和回去。
    for (const cell of cells) cell.widthTwips = 1200;
    expect(restateColumnFromCells(wide, 0)).toBe(true);
    expect(gridWidths(wide)).toEqual([1200, 1500, 1500, 1500]);
    const xml = saved(doc);
    expect(xml).toContain('<w:gridCol w:w="1200"/>');
    expect(xml).toContain('<w:tblW w:w="5700" w:type="dxa"/>');
  });
});
