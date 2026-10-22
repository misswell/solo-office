import { describe, expect, it } from "vitest";
import type { Block, Table, TableCell } from "../src/core/docx.js";
import { WordDocument, signCellPlaceRevision } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
import {
  cellRect,
  columnStarts,
  deleteCellHere,
  deleteColumnHere,
  deleteRowHere,
  deleteTableBlock,
  gridColumns,
  insertCellHere,
  insertColumnHere,
  insertRowHere,
  mergeHere,
  splitHere,
} from "../src/core/doc-table-ops.js";
import { makeMergedTableDocx, makeNestedTableDocx, makeOpsGridDocx, makeTableFidelityDocx } from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;

/** The grid every command here is checked against: three columns of 1000/2000/3000, last row spans two. */
const grid = (): Table => tableOf(open(makeOpsGridDocx()));
const withDoc = (t: Table): WordDocument => {
  const doc = WordDocument.blank();
  doc.blocks.push(t);
  return doc;
};
/** A table written out and read back, so a command's shape is proved against the file and not only the model. */
const reloaded = (t: Table): Table => tableOf(open(writeDocx(withDoc(t))));

/** Each row as `row:start,start|width x colspanVmerge` — the geometry a structural command has to leave right. */
const shape = (t: Table): string[] =>
  t.rows.map(
    (row, r) =>
      `${r}:${columnStarts(row).join(",")}=${row
        .map((c) => `${c.widthTwips ?? 0}x${c.spanCols ?? 1}${c.spanRows === undefined ? "" : `v${c.spanRows}`}`)
        .join(" ")}`,
  );
const body = (cell: TableCell): string =>
  cell.paragraphs
    .map((p: Block) => (p.kind === "table" ? "[表]" : p.kind === "carried" ? "[容器]" : p.inlines.map((i) => i.text).join("")))
    .join("/");

/** Every grid column of every row is covered, by something, exactly once — the rule a merge cannot break. */
const tilesTheGrid = (t: Table): boolean =>
  t.rows.every((row) => row.reduce((n, c) => n + Math.max(1, c.spanCols ?? 1), 0) === gridColumns(t));

describe("docx table structure commands", () => {
  it("inserts a row that copies the grid it lands in and leaves the file's grid untouched", () => {
    const t = grid();
    const carried = t.gridXml;
    expect(insertRowHere(t, { row: 0, col: 2 }, "below")).toEqual({ ok: true, at: { row: 1, col: 0 } });
    expect(t.rows).toHaveLength(4);
    expect(t.rows[1].map(body)).toEqual(["", "", ""]);
    expect(t.rows[1].map((c) => c.widthTwips)).toEqual([1000, 2000, 3000]);
    expect(t.gridXml).toBe(carried);
    expect(saved(withDoc(t)).match(/<w:tr[ >]/g)).toHaveLength(4);
    expect(shape(reloaded(t))).toEqual(shape(t));
    expect(tilesTheGrid(t)).toBe(true);
  });

  it("inserts a row inside a merged band and makes that band taller", () => {
    const t = tableOf(open(makeMergedTableDocx()));
    expect(insertRowHere(t, { row: 0, col: 1 }, "below")).toEqual({ ok: true, at: { row: 1, col: 0 } });
    expect(t.rows[1][0]).toMatchObject({ spanRows: 0 });
    expect(t.rows[1][1]!.spanRows).toBeUndefined();
    expect(renderWord(withDoc(t))).toContain('rowspan="3"');
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("splits the column a new one lands in front of, and pays for it out of the cells that own it", () => {
    const t = grid();
    expect(insertColumnHere(t, { row: 0, col: 1 }, "left")).toEqual({ ok: true, at: { row: 0, col: 1 } });
    expect(gridColumns(t)).toBe(4);
    expect(t.columnWidths).toEqual([1000, 1000, 1000, 3000]);
    expect(t.gridXml).toBe('<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="3000"/></w:tblGrid>');
    // The cell the seam starts inside keeps the host column's remainder; the one running through gains span only.
    expect(t.rows[0].map((c) => [c.widthTwips, c.spanCols ?? 1])).toEqual([[1000, 1], [1000, 1], [1000, 1], [3000, 1]]);
    expect(t.rows[2].map((c) => [c.widthTwips, c.spanCols ?? 1])).toEqual([[3000, 3], [3000, 1]]);
    expect(t.rows[2][0]!.tcPrXml).toContain('<w:gridSpan w:val="3"/>');
    expect(t.rows[2][0]!.tcPrXml).not.toContain('w:val="2"');
    expect(tilesTheGrid(t)).toBe(true);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("keeps a cell's own borders, shading and margins byte for byte while the grid moves", () => {
    const t = grid();
    const untouched = t.rows[0]![0]!.tcPrXml!;
    insertColumnHere(t, { row: 0, col: 1 }, "left");
    expect(t.rows[0]![0]!.tcPrXml).toBe(untouched);

    const other = grid();
    const before = other.rows[0]![0]!.tcPrXml!;
    insertColumnHere(other, { row: 0, col: 0 }, "left");
    expect(other.rows[0]![1]!.tcPrXml).toBe(before.replace('w:w="1000"', 'w:w="500"'));
    expect(other.rows[0]![1]!.tcPrXml).toContain('<w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="0070C0"/>');
    expect(other.rows[0]![1]!.tcPrXml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>');
    expect(other.rows[0]![1]!.tcPrXml).toContain('<w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar>');
    expect(other.columnWidths).toEqual([500, 500, 2000, 3000]);
  });

  it("takes a column out and one column of span off the cell that covered it", () => {
    const t = grid();
    expect(deleteColumnHere(t, { row: 0, col: 0 })).toEqual({ ok: true, at: { row: 0, col: 0 } });
    expect(gridColumns(t)).toBe(2);
    expect(t.columnWidths).toEqual([2000, 3000]);
    expect(t.rows.map((row) => row.length)).toEqual([2, 2, 2]);
    expect(t.rows[2][0]).toMatchObject({ spanCols: undefined, widthTwips: 2000 });
    expect(t.rows[2][0]!.tcPrXml).not.toContain("gridSpan");
    expect(t.rows.map((row) => row.reduce((n, c) => n + (c.widthTwips ?? 0), 0))).toEqual([5000, 5000, 5000]);
    expect(tilesTheGrid(t)).toBe(true);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("keeps the table's own width statement equal to the grid, the way the files write it", () => {
    // 授权口径 46 件里 58 枚 `w:tblW` 全拼 `dxa` 且等于本表网格和（`.scratch/table-width-audience.py`）：
    // 「整表多宽」这一句是照网格算出来的，所以命令动了网格的和，这一句要跟着动。
    const grows = grid();
    expect(grows.tblPrXml).toContain('<w:tblW w:w="6000" w:type="dxa"/>');
    expect(insertColumnHere(grows, { row: 0, col: 2 }, "right").ok).toBe(true);
    expect(grows.columnWidths).toEqual([1000, 2000, 3000, 3000]);
    expect(grows.widthTwips).toBe(9000);
    expect(saved(withDoc(grows))).toContain('<w:tblW w:w="9000" w:type="dxa"/>');
    expect(saved(withDoc(grows))).not.toContain('<w:tblW w:w="6000"');

    const shrinks = grid();
    expect(deleteColumnHere(shrinks, { row: 0, col: 0 }).ok).toBe(true);
    expect(shrinks.widthTwips).toBe(5000);
    expect(saved(withDoc(shrinks))).toContain('<w:tblW w:w="5000" w:type="dxa"/>');

    // 插在两条列中间是把那一条一分为二：和没变，那一枚孩子就一个字也不许多动。
    const halves = grid();
    const stated = halves.tblPrXml;
    expect(insertColumnHere(halves, { row: 0, col: 1 }, "left").ok).toBe(true);
    expect(halves.tblPrXml).toBe(stated);
    expect(halves.widthTwips).toBe(6000);

    // `pct` 量的是纸面可用宽度的百分比、不是网格的和；没说过宽度的表也不给它造一枚。
    const percent = grid();
    percent.widthType = "pct";
    percent.widthTwips = 5000;
    percent.tblPrXml = '<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/></w:tblPr>';
    expect(insertColumnHere(percent, { row: 0, col: 2 }, "right").ok).toBe(true);
    expect(percent.widthTwips).toBe(5000);
    expect(saved(withDoc(percent))).toContain('<w:tblW w:w="5000" w:type="pct"/>');

    const silent = grid();
    silent.widthTwips = undefined;
    silent.tblPrXml = '<w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>';
    expect(insertColumnHere(silent, { row: 0, col: 2 }, "right").ok).toBe(true);
    expect(silent.widthTwips).toBeUndefined();
    expect(saved(withDoc(silent))).not.toContain("<w:tblW");
  });

  it("deletes a row and hands a merge that lost its start to the row below", () => {
    const t = tableOf(open(makeMergedTableDocx()));
    insertRowHere(t, { row: 0, col: 0 }, "below");
    expect(deleteRowHere(t, { row: 0, col: 1 }).ok).toBe(true);
    expect(t.rows).toHaveLength(3);
    expect(t.rows[0][0]!.spanRows).toBe(1);
    expect(t.rows[0][0]!.spanRows).toBe(1);
    expect(body(t.rows[0]![0]!)).toBe("");
    expect(saved(withDoc(t))).toContain('<w:vMerge w:val="restart"/>');
    expect(renderWord(withDoc(t))).toContain('rowspan="2"');
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("merges two cells in a row into one and keeps both bodies in reading order", () => {
    const t = grid();
    expect(mergeHere(t, { row: 0, col: 1 }, { row: 0, col: 2 })).toEqual({ ok: true, at: { row: 0, col: 1 } });
    expect(t.rows[0]).toHaveLength(2);
    expect(t.rows[0][1]).toMatchObject({ spanCols: 2, widthTwips: 5000 });
    expect(t.rows[0][1]!.spanRows).toBeUndefined();
    expect(body(t.rows[0][1]!)).toBe("乙一/丙一");
    expect(t.rows[0][1]!.tcPrXml).toContain('<w:gridSpan w:val="2"/>');
    expect(t.rows[0][1]!.tcPrXml).not.toContain("vMerge");
    expect(saved(withDoc(t)).match(/<w:tc>/g)).toHaveLength(7);
    expect(tilesTheGrid(t)).toBe(true);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("merges a rectangle down as a start with a continuation under it", () => {
    const t = grid();
    expect(mergeHere(t, { row: 0, col: 2 }, { row: 1, col: 2 })).toEqual({ ok: true, at: { row: 0, col: 2 } });
    expect(t.rows[0][2]).toMatchObject({ spanRows: 1, spanCols: undefined, widthTwips: 3000 });
    expect(t.rows[1][2]).toMatchObject({ spanRows: 0 });
    expect(body(t.rows[0][2]!)).toBe("丙一/丙二");
    expect(body(t.rows[1][2]!)).toBe("");
    expect(t.rows[1][2]!.tcPrXml).toBe('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr>');
    const html = renderWord(withDoc(t));
    expect(html).toContain('rowspan="2"');
    expect(html.match(/<td/g)).toHaveLength(8);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("refuses a merge that would cut through a merge nobody selected", () => {
    const t = grid();
    const before = shape(t);
    expect(mergeHere(t, { row: 1, col: 1 }, { row: 2, col: 1 })).toEqual({
      ok: false,
      reason: "选区压到了已有合并的边界",
    });
    expect(shape(t)).toEqual(before);
    expect(mergeHere(t, { row: 0, col: 0 })).toEqual({ ok: false, reason: "已经是一个单元格了" });
    expect(splitHere(t, { row: 0, col: 0 })).toEqual({ ok: false, reason: "这个单元格没有合并" });
    expect(insertRowHere(t, { row: 9, col: 0 }, "above")).toEqual({ ok: false, reason: "这一行不在表格里" });
  });

  it("splits a cell that spans columns back into the grid it came from", () => {
    const t = grid();
    expect(splitHere(t, { row: 2, col: 0 })).toEqual({ ok: true, at: { row: 2, col: 0 } });
    expect(t.rows[2].map((c) => [body(c), c.widthTwips])).toEqual([["跨两列", 1500], ["", 1500], ["丙三", 3000]]);
    expect(t.rows[2].every((c) => c.spanCols === undefined)).toBe(true);
    expect(t.rows[2][0]!.tcPrXml).not.toContain("gridSpan");
    expect(t.columnWidths).toEqual([1000, 2000, 3000]);
    expect(tilesTheGrid(t)).toBe(true);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("splits a vertical merge back into ordinary cells, in every row it covered", () => {
    const t = tableOf(open(makeMergedTableDocx()));
    expect(splitHere(t, { row: 0, col: 0 })).toEqual({ ok: true, at: { row: 0, col: 0 } });
    expect(t.rows.map((row) => row.map((c) => c.spanRows))).toEqual([
      [undefined, undefined],
      [undefined, undefined],
      [1, undefined],
    ]);
    // The third row started a merge of its own; splitting the one above has no business with it.
    expect(saved(withDoc(t)).match(/<w:vMerge/g)).toHaveLength(1);
    expect(renderWord(withDoc(t))).not.toContain("rowspan");
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("keeps a table that would be emptied, and a percent-wide cell the file sized", () => {
    const single = reloaded({ kind: "table", columns: 1, rows: [[{ paragraphs: [], widthTwips: 4513 }]] });
    expect(deleteColumnHere(single, { row: 0, col: 0 })).toEqual({ ok: false, reason: "表格至少得有一列" });
    expect(deleteRowHere(single, { row: 0, col: 0 })).toEqual({ ok: false, reason: "表格至少得有一行" });

    // A `pct` width is measured against the text column, not against `w:tblGrid`, so twips arithmetic has no
    // business touching it: the cell keeps saying percent and the seam simply does not pay it any twips.
    const t = grid();
    const pct = t.rows[0]![1]!;
    pct.tcPrXml = '<w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr>';
    pct.widthType = "pct";
    const before = pct.tcPrXml;
    expect(insertColumnHere(t, { row: 0, col: 1 }, "left").ok).toBe(true);
    expect(pct.tcPrXml).toBe(before);
    expect(saved(withDoc(t))).toContain('<w:tcW w:w="2500" w:type="pct"/>');
  });

  it("writes a row's own properties in the order the schema gives them", () => {
    const t = grid();
    const row = { keepTogether: true, heightTwips: 600, heightRule: "atLeast", header: true };
    t.rows[1].forEach((c) => (c.row = row));
    expect(saved(withDoc(t))).toContain('<w:trPr><w:cantSplit/><w:trHeight w:val="600" w:hRule="atLeast"/><w:tblHeader/></w:trPr>');
  });
});

describe("删除单元格（DeleteCells）：一格离开，旁边的格补进来", () => {
  it("左移：这一行少一格，末格接管腾出来的列，别的行一格不动", () => {
    const t = grid();
    expect(deleteCellHere(t, { row: 0, col: 1 }, "left")).toEqual({ ok: true, at: { row: 0, col: 1 } });
    expect(gridColumns(t)).toBe(3);
    expect(t.rows[0].map((c) => c.widthTwips)).toEqual([1000, 5000]);
    expect(t.rows[0].map((c) => c.spanCols ?? 1)).toEqual([1, 2]);
    expect(columnStarts(t.rows[0])).toEqual([0, 1]);
    expect(tilesTheGrid(t)).toBe(true);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("左移不动别的行，也不腾网格：一格要走的是一行的位子，不是一列", () => {
    const t = grid();
    const others = t.rows.slice(1).map((_, r) => shape(t)[r + 1]);
    const widths = [...(t.columnWidths ?? [])];
    deleteCellHere(t, { row: 0, col: 0 }, "left");
    expect(t.rows.slice(1).map((_, r) => shape(t)[r + 1])).toEqual(others);
    expect(t.columnWidths).toEqual(widths);
  });

  it("上移：下面那格的字补进光标那一格，末行少一格并自己接管", () => {
    const t = grid();
    const last = t.rows.length - 1;
    const band = (r: number): number => columnStarts(t.rows[r]).indexOf(2);
    const below = body(t.rows[1][band(1)]!);
    const tailBody = body(t.rows[last][band(last)]!);
    const widths = t.rows[0].map((c) => c.widthTwips);
    expect(deleteCellHere(t, { row: 0, col: 2 }, "up")).toEqual({ ok: true, at: { row: 0, col: 2 } });
    // 一格走了，这一条带上的字一路往上补；格子自己不动，动的只有字——与「下方插入」同一条规矩
    expect(body(t.rows[0][2]!)).toBe(below);
    expect(body(t.rows[1][band(1)]!)).toBe(tailBody);
    expect(t.rows[0].map((c) => c.widthTwips)).toEqual(widths);
    expect(gridColumns(t)).toBe(3);
    expect(t.rows[last]).toHaveLength(1);
    expect(t.rows[last][0]!.spanCols ?? 1).toBe(3);
    expect(tilesTheGrid(t)).toBe(true);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("形状说不出来就回绝，不硬画", () => {
    expect(deleteCellHere(grid(), { row: 9, col: 0 }, "left")).toEqual({ ok: false, reason: "这里没有单元格" });
    expect(deleteCellHere(grid(), { row: 9, col: 0 }, "up")).toEqual({ ok: false, reason: "这里没有单元格" });
    // 最后一行下面没有可补上来的
    const last = grid();
    expect(deleteCellHere(last, { row: last.rows.length - 1, col: 0 }, "up")).toEqual({
      ok: false,
      reason: "下面没有可补上来的单元格",
    });
    // 那一列上有横向合并的格：这一带没有格子接得住，回绝而不是把它压扁
    expect(deleteCellHere(grid(), { row: 0, col: 1 }, "up")).toEqual({
      ok: false,
      reason: "这一列上有横向合并的单元格，上移会让它装不进来",
    });
  });

  it("一句 w:cellDel 落到件上、读回来还在，接受后与不跟踪那一击同一个形状", () => {
    const t = grid();
    const doc = withDoc(t);
    signCellPlaceRevision(t.rows[0][1]!, { id: 7, author: "甲", date: "2026-10-22T00:00:00Z" }, "cellDel");
    const xml = saved(doc);
    expect(xml).toContain('<w:cellDel w:id="7" w:author="甲"');
    const back = tableOf(open(writeDocx(doc)));
    expect(back.rows[0][1]!.revs?.map((r) => r.tag)).toEqual(["cellDel"]);
    expect(back.rows[0][1]!.revs?.[0]?.prev).toEqual({});

    const plain = grid();
    deleteCellHere(plain, { row: 0, col: 1 }, "left");
    doc.resolveRevisions("accept");
    expect(shape(tableOf(doc))).toEqual(shape(plain));

    const refusedTwice = grid();
    const again = withDoc(refusedTwice);
    const cell = refusedTwice.rows[0][0]!;
    signCellPlaceRevision(cell, { id: 8, author: "甲", date: "2026-10-22T00:00:00Z" }, "cellDel");
    expect(signCellPlaceRevision(cell, { id: 9, author: "甲", date: "2026-10-22T00:00:00Z" }, "cellDel").id).toBe(8);
    expect(saved(again).match(/<w:cellDel/g)).toHaveLength(1);
  });

  it("拒绝这一句：格子留在原处，那一句离开件", () => {
    const t = grid();
    const doc = withDoc(t);
    const kept = shape(t);
    signCellPlaceRevision(t.rows[0][1]!, { id: 7, author: "甲", date: "2026-10-22T00:00:00Z" }, "cellDel");
    doc.resolveRevisions("reject");
    expect(shape(tableOf(doc))).toEqual(kept);
    expect(saved(doc)).not.toContain("<w:cellDel");
  });
});

describe("插入单元格（InsertCells）：一个格，旁边的格让位", () => {
  it("右移：这一行多一个空格，其余的行只多盖一列", () => {
    const t = grid();
    expect(insertCellHere(t, { row: 0, col: 1 }, "right")).toEqual({ ok: true, at: { row: 0, col: 1 } });
    expect(gridColumns(t)).toBe(4);
    expect(t.columnWidths).toEqual([1000, 1000, 1000, 3000]);
    expect(t.gridXml).toBe(
      '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="3000"/></w:tblGrid>',
    );
    // The row that asked for a cell pays for it; the rows that did not gain a column of span and no width,
    // because the column they now cover is one they already owned.
    expect(shape(t)).toEqual([
      "0:0,1,2,3=1000x1 1000x1 1000x1 3000x1",
      "1:0,1,3=1000x1 2000x2 3000x1",
      "2:0,3=3000x3 3000x1",
    ]);
    expect(t.rows[0].map(body)).toEqual(["甲一", "", "乙一", "丙一"]);
    expect(t.rows.map((row) => row.reduce((n, c) => n + (c.widthTwips ?? 0), 0))).toEqual([6000, 6000, 6000]);
    expect(tilesTheGrid(t)).toBe(true);
    expect(shape(reloaded(t))).toEqual(shape(t));
    expect(saved(withDoc(t)).match(/<w:tc>/g)).toHaveLength(9);
  });

  it("右移落在跨两列的格上：它让出宽度，不让出跨列", () => {
    const t = grid();
    expect(insertCellHere(t, { row: 2, col: 0 }, "right")).toEqual({ ok: true, at: { row: 2, col: 0 } });
    expect(t.columnWidths).toEqual([500, 500, 2000, 3000]);
    expect(shape(t)).toEqual([
      "0:0,2,3=1000x2 2000x1 3000x1",
      "1:0,2,3=1000x2 2000x1 3000x1",
      "2:0,1,3=500x1 2500x2 3000x1",
    ]);
    expect(t.rows[2][0]!.paragraphs).toEqual([{ kind: "paragraph", inlines: [] }]);
    expect(t.rows[2][1]!.tcPrXml).toContain('<w:gridSpan w:val="2"/>');
    expect(t.rows[2][1]!.tcPrXml).toContain('w:w="2500"');
    expect(tilesTheGrid(t)).toBe(true);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("右移不动按百分比定宽的格，那一条缝不该用 twips 付", () => {
    const t = grid();
    const pct = t.rows[0]![1]!;
    pct.tcPrXml = '<w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr>';
    pct.widthType = "pct";
    const before = pct.tcPrXml;
    expect(insertCellHere(t, { row: 0, col: 1 }, "right").ok).toBe(true);
    expect(pct.tcPrXml).toBe(before);
    expect(saved(withDoc(t))).toContain('<w:tcW w:w="2500" w:type="pct"/>');
  });

  it("下移：一格的内容往下走一格，最后一行把没有去处的那格收下", () => {
    const t = tableOf(open(makeMergedTableDocx()));
    const carried = t.gridXml;
    expect(insertCellHere(t, { row: 0, col: 1 }, "down")).toEqual({ ok: true, at: { row: 0, col: 1 } });
    expect(t.rows.map((row) => row.map(body))).toEqual([
      ["占两行", ""],
      ["", "第一行右"],
      ["新的一格", "第二行右/第三行右"],
    ]);
    // A cell is added to the column, not a row to the table: the grid and every cell's own properties stay.
    expect(t.gridXml).toBe(carried);
    expect(t.columnWidths).toEqual([2000, 3000]);
    expect(t.rows[0][1]!.tcPrXml).toBe('<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>');
    expect(saved(withDoc(t)).match(/<w:vMerge/g)).toHaveLength(3);
    expect(renderWord(withDoc(t))).toContain('rowspan="2"');
    const back = reloaded(t);
    expect(shape(back)).toEqual(shape(t));
    expect(back.rows.map((row) => row.map(body))).toEqual(t.rows.map((row) => row.map(body)));
  });

  it("下移把格子里的表一起带走，写出的件里格子仍以段落收尾", () => {
    const t = tableOf(open(makeNestedTableDocx()));
    expect(insertCellHere(t, { row: 0, col: 1 }, "down")).toEqual({ ok: true, at: { row: 0, col: 1 } });
    expect(gridColumns(t)).toBe(2);
    expect(t.rows[0][1]!.paragraphs.map((b) => b.kind)).toEqual(["paragraph"]);
    expect(t.rows[1][1]!.paragraphs.map((b) => b.kind)).toEqual(["paragraph", "table", "paragraph"]);
    const xml = saved(withDoc(t));
    expect(xml).not.toContain("</w:tbl></w:tc>");
    const back = tableOf(open(writeDocx(withDoc(t))));
    expect(back.rows[1]![1]!.paragraphs.map((b) => b.kind)).toEqual(["paragraph", "table", "paragraph"]);
    expect((back.rows[1]![1]!.paragraphs[1] as Table).rows.map((row) => row.map(body).join("|"))).toEqual([
      "内一|内二",
      "内三|内四",
    ]);
  });

  it("下移在装不下的地方停下，并且什么都不动", () => {
    const merged = tableOf(open(makeMergedTableDocx()));
    const before = shape(merged);
    expect(insertCellHere(merged, { row: 2, col: 1 }, "down")).toEqual({ ok: false, reason: "下面没有可让位的单元格" });
    expect(insertCellHere(merged, { row: 0, col: 0 }, "down")).toEqual({
      ok: false,
      reason: "这一列上有纵向合并的单元格，下移会把它挪位",
    });
    expect(insertCellHere(merged, { row: 1, col: 0 }, "down")).toEqual({
      ok: false,
      reason: "这一列上有纵向合并的单元格，下移会把它挪位",
    });
    expect(shape(merged)).toEqual(before);

    const t = grid();
    const gridBefore = shape(t);
    expect(insertCellHere(t, { row: 0, col: 1 }, "down")).toEqual({
      ok: false,
      reason: "这一列上有横向合并的单元格，下移会让它装不进来",
    });
    expect(shape(t)).toEqual(gridBefore);
    expect(insertCellHere(t, { row: 9, col: 0 }, "down")).toEqual({ ok: false, reason: "这里没有单元格" });
    expect(insertCellHere(t, { row: 9, col: 0 }, "right")).toEqual({ ok: false, reason: "这里没有单元格" });
  });
});

describe("cell geometry", () => {
  it("finds the grid rectangle each cell covers", () => {
    const t = grid();
    expect(cellRect(t, { row: 0, col: 0 })).toEqual({ c0: 0, c1: 1, r0: 0, r1: 1 });
    expect(cellRect(t, { row: 2, col: 0 })).toEqual({ c0: 0, c1: 2, r0: 2, r1: 3 });
    const merged = tableOf(open(makeMergedTableDocx()));
    expect(cellRect(merged, { row: 0, col: 0 })).toEqual({ c0: 0, c1: 1, r0: 0, r1: 2 });
    expect(cellRect(merged, { row: 1, col: 0 })).toEqual({ c0: 0, c1: 1, r0: 1, r1: 2 });
  });
});

describe("docx 删除表格（块表里的那一张离开）", () => {
  /** 一张表换成一段空行：表没了，块序还在原来的位子上。 */
  const kinds = (blocks: Block[]): string[] => blocks.map((b) => b.kind);
  const para = (text: string): Block => ({ kind: "paragraph", inlines: [{ text }] });

  it("leaves one empty paragraph in the slot and reports where it landed", () => {
    const blocks: Block[] = [para("表前"), grid(), para("表后")];
    expect(deleteTableBlock(blocks, 1)).toEqual({ ok: true, at: 1 });
    expect(kinds(blocks)).toEqual(["paragraph", "paragraph", "paragraph"]);
    expect(blocks[1]).toEqual({ kind: "paragraph", inlines: [] });
    // 那一段两边的正文各归各，删表不碰别人的字。
    expect(blocks.map((b) => (b.kind === "paragraph" ? b.inlines.map((i) => i.text).join("") : "")).join("/")).toBe("表前//表后");
  });

  it("keeps the list something the caret can sit on: one table alone leaves a paragraph, not nothing", () => {
    const blocks: Block[] = [grid()];
    expect(deleteTableBlock(blocks, 0)).toEqual({ ok: true, at: 0 });
    expect(kinds(blocks)).toEqual(["paragraph"]);
  });

  it("never leaves two tables side by side, which is the shape a reader counts as one table", () => {
    const blocks: Block[] = [grid(), grid()];
    expect(deleteTableBlock(blocks, 0).ok).toBe(true);
    expect(kinds(blocks)).toEqual(["paragraph", "table"]);
  });

  it("refuses a slot that holds no table and changes nothing when it does", () => {
    const blocks: Block[] = [{ kind: "paragraph", inlines: [{ text: "正文" }] }];
    expect(deleteTableBlock(blocks, 0)).toEqual({ ok: false, reason: "这里没有表格" });
    expect(deleteTableBlock(blocks, 7)).toEqual({ ok: false, reason: "这里没有表格" });
    expect(blocks).toEqual([{ kind: "paragraph", inlines: [{ text: "正文" }] }]);
  });

  it("deletes a nested table out of the cell that holds it and leaves the outer table standing", () => {
    const doc = open(makeNestedTableDocx());
    const outer = tableOf(doc);
    const holder = outer.rows.flatMap((row) => row).find((cell) => cell.paragraphs.some((b) => b.kind === "table"));
    expect(holder).toBeDefined();
    const at = holder!.paragraphs.findIndex((b) => b.kind === "table");
    expect(deleteTableBlock(holder!.paragraphs, at)).toEqual({ ok: true, at });
    expect(holder!.paragraphs[at]).toEqual({ kind: "paragraph", inlines: [] });
    expect(outer.rows.length).toBeGreaterThan(0);
    expect(holder!.paragraphs.some((b) => b.kind === "table")).toBe(false);
  });

  it("writes the deletion out and reads it back with the paragraph already in the model", () => {
    const doc = WordDocument.blank();
    const t = grid();
    doc.blocks.push(t);
    // blank() 起头就留了一段可落的光标行，删表再把表自己的那位子也换成一段，读回来是两段。
    const at = doc.blocks.indexOf(t);
    expect(deleteTableBlock(doc.blocks, at)).toEqual({ ok: true, at });
    expect(saved(doc)).not.toContain("<w:tbl>");
    expect(kinds(open(writeDocx(doc)).blocks)).toEqual(["paragraph", "paragraph"]);
  });
});
