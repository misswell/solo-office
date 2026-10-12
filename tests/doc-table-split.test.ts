import { describe, expect, it } from "vitest";
import type { Paragraph, Revision, RowInfo, Table, TableCell } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";

import { splitTableAt } from "../src/core/doc-table-ops.js";
import { writeDocx } from "../src/core/write-docx.js";
import { makeDeepNestedTableDocx, makeTablePropertyRevisionDocx } from "./helpers/make-tables.js";

const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;
const nestedIn = (t: Table, row: number, col: number): Table | undefined =>
  t.rows[row]?.[col]?.paragraphs.find((b) => b.kind === "table") as Table | undefined;
const text = (t: Table): string => t.rows.map((r) => r.map((c) => c.paragraphs.map(blockText).join("")).join("|")).join("/");
const blockText = (b: Table["rows"][number][number]["paragraphs"][number]): string =>
  b.kind === "paragraph" ? b.inlines.map((i) => i.text).join("") : b.kind === "table" ? `[${text(b)}]` : "";

const para = (value: string): Paragraph => ({ kind: "paragraph", inlines: [{ text: value }] });
const cell = (width: number, value: string): TableCell => ({ paragraphs: [para(value)], widthTwips: width });
/** A rows-by-columns grid whose cells all say which row and column they sit in. */
const grid = (rows: number, cols: number, width = 1000): Table => ({
  kind: "table",
  columns: cols,
  rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_2, c) => cell(width, `${r}${c}`))),
  widthTwips: width * cols,
  layout: "fixed",
  columnWidths: Array.from({ length: cols }, () => width),
});

describe("docx 表格：拆分表格", () => {
  it("moves the caret's row and everything under it into a second table", () => {
    const t = grid(4, 2);
    const split = splitTableAt(t, 2);
    expect(split.ok).toBe(true);
    if (!split.ok) return;
    expect(text(t)).toBe("00|01/10|11");
    expect(text(split.tail)).toBe("20|21/30|31");
  });

  it("gives the lower half the grid, the widths and the properties the file wrote for the table", () => {
    const t = grid(3, 2, 800);
    t.styleId = "a3";
    t.align = "center";
    t.indentTwips = 108;
    t.gridXml = '<w:tblGrid><w:gridCol w:w="800"/><w:gridCol w:w="800"/></w:tblGrid>';
    t.tblPrXml = '<w:tblPr><w:tblW w:w="1600" w:type="dxa"/><w:tblLook w:val="04A0"/></w:tblPr>';
    const split = splitTableAt(t, 1);
    if (!split.ok) throw new Error(split.reason);
    expect(split.tail.styleId).toBe("a3");
    expect(split.tail.align).toBe("center");
    expect(split.tail.indentTwips).toBe(108);
    expect(split.tail.gridXml).toBe(t.gridXml);
    expect(split.tail.tblPrXml).toBe(t.tblPrXml);
    expect(split.tail.columnWidths).toEqual([800, 800]);
    expect(split.tail.columns).toBe(2);
  });

  it("refuses where a half would come out empty, and changes nothing when it does", () => {
    const t = grid(2, 2);
    expect(splitTableAt(t, 0)).toEqual({ ok: false, reason: "第一行上面没有可拆的半张表" });
    expect(splitTableAt(t, t.rows.length)).toEqual({ ok: false, reason: "最后一行下面没有可拆的半张表" });
    expect(splitTableAt(t, 9)).toEqual({ ok: false, reason: "最后一行下面没有可拆的半张表" });
    expect(text(t)).toBe("00|01/10|11");
  });

  it("refuses a table that carries its own undecided property record", () => {
    const doc = WordDocument.openBytes(makeTablePropertyRevisionDocx());
    const t = tableOf(doc);
    expect(t.revFormat?.id).toBe(61);
    const split = splitTableAt(t, 1);
    // One `w:tblPrChange` with one `w:id` cannot describe two tables, so the command says no rather than
    // signing the same change onto both halves and listing it twice in the pane.
    expect(split.ok).toBe(false);
    expect(t.rows.length).toBe(2);
    expect(t.revFormat?.id).toBe(61);
  });

  it("lets a vertical merge cut in half start again in the half it landed in", () => {
    const t = grid(4, 1);
    t.rows[0][0].spanRows = 1;
    t.rows[1][0].spanRows = 0;
    t.rows[2][0].spanRows = 0;
    const split = splitTableAt(t, 2);
    if (!split.ok) throw new Error(split.reason);
    // The head keeps its merge; the tail's first cell was a continuation with no anchor above it.
    expect(t.rows.map((r) => r[0].spanRows)).toEqual([1, 0]);
    expect(split.tail.rows.map((r) => r[0].spanRows)).toEqual([1, undefined]);
  });

  it("carries a row's own record down with the row it belongs to", () => {
    const t = grid(3, 2);
    const rev: Revision = { kind: "ins", id: 71, author: "顾峰" };
    const row: RowInfo = { rev };
    t.rows[2].forEach((c) => (c.row = row));
    const split = splitTableAt(t, 2);
    if (!split.ok) throw new Error(split.reason);
    expect(split.tail.rows[0][0].row?.rev?.id).toBe(71);
    expect(t.rows.some((r) => r.some((c) => c.row?.rev))).toBe(false);
  });

  it("writes the two tables and the paragraph between them, and reads them back the same way", () => {
    const doc = WordDocument.openBytes(makeDeepNestedTableDocx());
    const outer = tableOf(doc);
    const split = splitTableAt(outer, 1);
    if (!split.ok) throw new Error(split.reason);
    const at = doc.blocks.indexOf(outer);
    doc.blocks.splice(at + 1, 0, { kind: "paragraph", inlines: [] }, split.tail);

    const bytes = writeDocx(doc);
    const reopened = WordDocument.openBytes(bytes);
    const blocks = reopened.blocks.map((b) => (b.kind === "table" ? `table(${text(b)})` : b.kind === "paragraph" ? `paragraph(${b.inlines.map((i) => i.text).join("")})` : b.kind));
    expect(blocks).toEqual([
      "paragraph(表前)",
      "table(外一|外二[中一[深一|深二/深三|深四]|中二/中三|中四])",
      "paragraph()",
      "table(外三|外四)",
      "paragraph()",
    ]);
    // The nested table travelled with the row it sat in, and the split did not merge the two halves back.
    const first = reopened.blocks[1] as Table;
    expect(nestedIn(first, 0, 1) && text(nestedIn(first, 0, 1)!)).toBe("中一[深一|深二/深三|深四]|中二/中三|中四");
  });
});
