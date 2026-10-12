import { describe, expect, it } from "vitest";
import type { Table } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { cellRangeFor, cellsInRect, columnStarts, gridColumns } from "../src/core/doc-table-ops.js";
import { makeMergedTableDocx, makeNestedTableDocx, makeOpsGridDocx } from "./helpers/make-tables.js";

/**
 * 选中整格 / 整行 / 整列 / 整个表格 answers a rectangle, so these tests are about geometry only: which two cells
 * hold the corners, and which cells the page then has to light up. `tests/doc-table-paste.test.ts` and the DOM
 * probe cover what the page does with them.
 */
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;
const grid = (): Table => tableOf(open(makeOpsGridDocx()));
const merged = (): Table => tableOf(open(makeMergedTableDocx()));
const nested = (): Table => tableOf(open(makeNestedTableDocx()));

type Kind = "cell" | "row" | "column" | "table";

/** "row,col" for each corner, so a失败 reads as the cells on the page rather than as four numbers. */
const corners = (t: Table, row: number, col: number, kind: Kind): string => {
  const ran = cellRangeFor(t, { row, col }, kind);
  return ran.ok ? `${ran.range.from.row},${ran.range.from.col}–${ran.range.to.row},${ran.range.to.col}` : ran.reason;
};
/** The cells a range paints, in the order the page walks them — one entry per cell, bands at their top. */
const painted = (t: Table, row: number, col: number, kind: Kind): string => {
  const ran = cellRangeFor(t, { row, col }, kind);
  return ran.ok ? cellsInRect(t, ran.range).map((c) => `${c.row},${c.col}`).join(" ") : ran.reason;
};
const refusal = (t: Table, row: number, col: number, kind: Kind): string => {
  const ran = cellRangeFor(t, { row, col }, kind);
  return ran.ok ? "" : ran.reason;
};
const shape = (t: Table): string[] =>
  t.rows.map((row, r) => `${r}:${columnStarts(row).join(",")}=${row.map((c) => `${c.widthTwips ?? 0}x${c.spanCols ?? 1}v${c.spanRows ?? "-"}`).join(" ")}`);

describe("选中整格/整行/整列/整个表格（SelectCell / SelectRow / SelectColumn / SelectTable）", () => {
  it("lights up the caret's own cell, a wide one taken as the whole span it covers", () => {
    const t = grid();
    // 甲一/乙一/丙一 各占一列：选中一格就是那一格，行和列都不动。
    expect(corners(t, 0, 1, "cell")).toBe("0,1–0,1");
    // 第三行第一格是 `w:gridSpan` 2：点亮的是那一个格子（两个角是同一格），不是它盖住的两列各一格。
    expect(corners(t, 2, 0, "cell")).toBe("2,0–2,0");
    expect(refusal(t, 0, 0, "cell")).toBe("");
  });

  it("takes a row's cells from its first to its last, whatever the row's own shape", () => {
    const t = grid();
    expect(corners(t, 1, 2, "row")).toBe("1,0–1,2");
    // 横向合并的那一行只有两个 `w:tc`：整行仍是从首格到末格，末格是丙三（网格第三列）。
    expect(corners(t, 2, 0, "row")).toBe("2,0–2,1");
    expect(corners(t, 0, 0, "row")).toBe("0,0–0,2");
    // 整行永远是整行：从第 0 列铺到网格末列，与光标落在哪一格无关。
    expect(corners(t, 0, 2, "row")).toBe(corners(t, 0, 0, "row"));
  });

  it("grows a row selection to the vertical band that pokes through it", () => {
    const t = merged();
    // 第二行第一格是续格（`w:vMerge` 无值）：整行把它上面那半截也带上，画面不会点亮半个格子。
    expect(corners(t, 1, 0, "row")).toBe("0,0–1,1");
    expect(corners(t, 0, 1, "row")).toBe("0,0–1,1");
    // 第三行自己起的那条带没被续过，所以整行就是它自己一行。
    expect(corners(t, 2, 0, "row")).toBe("2,0–2,1");
  });

  it("names a vertical band by its first row, from whichever half the caret sits in", () => {
    const t = merged();
    expect(corners(t, 0, 0, "cell")).toBe("0,0–1,0");
    // 光标在续格上：带还是那条带，两个角是它的顶与底。
    expect(corners(t, 1, 0, "cell")).toBe("0,0–1,0");
    expect(corners(t, 2, 0, "cell")).toBe("2,0–2,0");
    expect(corners(t, 0, 1, "cell")).toBe("0,1–0,1");
  });

  it("runs a column down the grid column the caret's cell starts at", () => {
    const t = merged();
    expect(corners(t, 0, 1, "column")).toBe("0,1–2,1");
    // 左列上有两条竖合并的带：整列照走，每格展开成自己的带，顶与底都在这一列上。
    expect(corners(t, 1, 0, "column")).toBe("0,0–2,0");
    const g = grid();
    // 第三列没有任何格子越过它。
    expect(corners(g, 0, 2, "column")).toBe("0,2–2,1");
  });

  it("takes the whole table when 整个表格 is asked for, bands and spans folded to one cell each", () => {
    const g = grid();
    // 三行网格 + 第三行跨两列：八格，从左上到右下，落点在哪一格都一样。
    expect(corners(g, 1, 2, "table")).toBe("0,0–2,1");
    expect(painted(g, 1, 2, "table")).toBe("0,0 0,1 0,2 1,0 1,1 1,2 2,0 2,1");
    expect(corners(g, 0, 0, "table")).toBe(corners(g, 2, 1, "table"));
    const t = merged();
    // 竖合并的续格不再列一遍：两条带 + 右列三格 = 五格。
    expect(corners(t, 0, 1, "table")).toBe("0,0–2,1");
    expect(painted(t, 2, 0, "table")).toBe("0,0 0,1 1,1 2,0 2,1");
    // 整张表永远选得出矩形：只有整列才要求那一列上下一般宽。
    expect(refusal(g, 1, 1, "table")).toBe("");
    expect(refusal(t, 1, 0, "table")).toBe("");
  });

  it("lists one entry per painted cell, so the highlight and the count are the same set", () => {
    // 整列要一格一段地画：一个 Range 只能是从前到后的一段，往下走的路上会把左邻格一起扫进来。
    // 这张表就是那一段路：第三列的三格点亮，靠右的那格在网格 0–2，没有越过边界。
    const g = grid();
    expect(painted(g, 0, 2, "column")).toBe("0,2 1,2 2,1");
    // 横向合并的那一行两个 `w:tc`、三列网格：点亮的是那两格，不是三列。
    expect(painted(g, 2, 0, "row")).toBe("2,0 2,1");
    const t = merged();
    // 竖合并的带在它盖住的每一行都留了一格 `w:tc`：列表里只出现带顶那一格，续格不再重复一遍。
    expect(painted(t, 1, 1, "row")).toBe("0,0 0,1 1,1");
    expect(painted(t, 0, 0, "column")).toBe("0,0 2,0");
    expect(painted(t, 0, 1, "column")).toBe("0,1 1,1 2,1");
    expect(painted(t, 0, 0, "cell")).toBe("0,0");
    // 续格自己当光标：展开的还是那条带，顶与底都在这张列表里。
    expect(painted(t, 1, 0, "cell")).toBe("0,0");
  });

  it("refuses a column whose strip is not a rectangle instead of over-lighting the row beside it", () => {
    const t = grid();
    // 第二列的下端是那个跨两列的格：它的网格是 0–2 而不是 1–2，矩形选法会把甲二那一格一起算进来，所以拒绝。
    expect(refusal(t, 0, 1, "column")).toContain("横向合并");
    // 拒绝的理由给得出来，但同一张表的第一列也一样越不过来：跨两列的格起点在 0，光标从丙一那列才走得通。
    expect(refusal(t, 0, 0, "column")).toContain("横向合并");
    expect(refusal(t, 2, 0, "column")).toContain("横向合并");
    // 竖合并不是拒绝的理由：合并件的两条带都在同一列里。
    expect(refusal(merged(), 0, 0, "column")).toBe("");
  });

  it("refuses a cell that is not there, and changes nothing when it answers", () => {
    const t = grid();
    const before = shape(t);
    expect(refusal(t, 9, 0, "cell")).toContain("不在表格");
    expect(refusal(t, 0, 9, "row")).toContain("不在表格");
    expect(refusal(t, 0, 9, "column")).toContain("不在表格");
    for (const kind of ["cell", "row", "column", "table"] as const) for (const [r, c] of [[0, 0], [1, 2], [2, 1]]) corners(t, r, c, kind);
    expect(refusal(t, 9, 0, "table")).toContain("不在表格");
    expect(shape(t)).toEqual(before);
    expect(t.rows.map((row) => row.length)).toEqual([3, 3, 2]);
    expect(gridColumns(t)).toBe(3);
  });

  it("reads a nested table's own cells, not the table around them", () => {
    const outer = nested();
    expect(corners(outer, 0, 1, "cell")).toBe("0,1–0,1");
    const inner = outer.rows[0][1]!.paragraphs.find((b) => b.kind === "table") as Table;
    expect(inner.columns).toBe(2);
    expect(corners(inner, 0, 0, "row")).toBe("0,0–0,1");
    expect(corners(inner, 0, 0, "column")).toBe("0,0–1,0");
    expect(corners(inner, 1, 1, "cell")).toBe("1,1–1,1");
    // 外层的那一格装着内层表，但选中外层不会把内层的格子算进来：两张表各有各的网格。
    expect(corners(outer, 0, 1, "row")).toBe("0,0–0,1");
    // 整张表也一样：外层四格（其中一格里塞着一张表，那颗不算进去），内层四格各自成一张。
    expect(painted(outer, 1, 1, "table")).toBe("0,0 0,1 1,0 1,1");
    expect(painted(inner, 0, 0, "table")).toBe("0,0 0,1 1,0 1,1");
  });
});
