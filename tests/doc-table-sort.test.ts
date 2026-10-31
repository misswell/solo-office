import { describe, expect, it } from "vitest";
import type { Inline, Paragraph, Table, TableCell } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { sortRows } from "../src/core/doc-table-ops.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const para = (value: string): Paragraph => ({ kind: "paragraph", inlines: [{ text: value }] as Inline[] });

/** 一张 2 列的表：每行第一格是排序列，第二格是这一行的名字（命令挪行时名字跟着走）。 */
const table = (header: boolean, rows: [string, string][]): Table => {
  const made: TableCell[][] = rows.map(([key, name]) => [
    { paragraphs: [para(key)], widthTwips: 1000 },
    { paragraphs: [para(name)], widthTwips: 1000 },
  ]);
  if (header) made.unshift([{ paragraphs: [para("键")], widthTwips: 1000 }, { paragraphs: [para("名")], widthTwips: 1000 }]);
  const body: Table = { kind: "table", columns: 2, rows: made, widthTwips: 2000, layout: "fixed", columnWidths: [1000, 1000] };
  if (header) made[0]!.forEach((cell) => (cell.row = { header: true }));
  return body;
};

const cellText = (cell: TableCell): string =>
  cell.paragraphs.map((p) => (p.kind === "paragraph" ? p.inlines.map((i: Inline) => i.text ?? "").join("") : "")).join("");
const keys = (t: Table): string[] => t.rows.map((row) => cellText(row[0]!));
const names = (t: Table): string[] => t.rows.map((row) => cellText(row[1]!));

const sort = (t: Table, spec: Partial<Parameters<typeof sortRows>[1]> = {}) =>
  sortRows(t, { column: 0, order: "asc", kind: "auto", headerRows: 0, ...spec });

const docWith = (t: Table): WordDocument => {
  const doc = WordDocument.blank();
  doc.blocks.push(t);
  return doc;
};

/**
 * 表格「排序」（`CT_TableTools.kuip:90` 的 `Sort`）。模型层钉四件事：排谁（标题行钉在顶上）、按什么比
 * （数字/文字/自动、升降序、并列保持原次序）、整行走（名字跟着键走、写盘读回同一序）、什么形状拒排
 * （竖合并、行间标记、待判的表属性改动、表体未裁决的行增删）。
 */
describe("docx 表格排序", () => {
  it("升序：标题行钉住，表体按第一列的数字从小到大", () => {
    const t = table(true, [["10", "十"], ["9", "九"], ["2", "二"]]);
    const out = sort(t, { headerRows: 1, kind: "number" });
    expect(out.ok).toBe(true);
    expect(keys(t)).toEqual(["键", "2", "9", "10"]);
    // 第二格跟着行一起走：挪的是整行，不是排序那一格。
    expect(names(t)).toEqual(["名", "二", "九", "十"]);
  });

  it("降序与「全表都排」：headerRows 0 时第一行也参与", () => {
    const t = table(false, [["a", "一"], ["c", "三"], ["b", "二"]]);
    const out = sort(t, { order: "desc" });
    expect(out.ok).toBe(true);
    expect(keys(t)).toEqual(["c", "b", "a"]);
  });

  it("数字档：不是数的排在他们后面，彼此按文字；文字档：一列数字按字面比", () => {
    const numeric = table(false, [["abc", "甲"], ["10", "乙"], ["9", "丙"]]);
    expect(sort(numeric, { kind: "number" }).ok).toBe(true);
    expect(keys(numeric)).toEqual(["9", "10", "abc"]);

    const textual = table(false, [["10", "甲"], ["9", "乙"], ["2", "丙"]]);
    expect(sort(textual, { kind: "text" }).ok).toBe(true);
    // 文字次序里 "10" 在 "2" 之前（逐字比）；数字档里才不会。
    expect(keys(textual)).toEqual(["10", "2", "9"]);
  });

  it("并列保持原次序（稳定），不会把同键的行翻过来", () => {
    const t = table(true, [["5", "先"], ["1", "一"], ["5", "后"]]);
    expect(sort(t, { headerRows: 1, kind: "number" }).ok).toBe(true);
    expect(names(t)).toEqual(["名", "一", "先", "后"]);
  });

  it("排完写盘再读回：行次序落在文件里，读出同一个序", () => {
    const t = table(true, [["3", "三"], ["1", "一"], ["2", "二"]]);
    expect(sort(t, { headerRows: 1, kind: "number" }).ok).toBe(true);
    const back = open(writeDocx(docWith(t))).blocks.find((b) => b.kind === "table") as Table;
    expect(names(back)).toEqual(["名", "一", "二", "三"]);
  });

  it("光标落在表体第一行的排序列：跨列格之后用的是格序，不是网格列", () => {
    const t: Table = {
      kind: "table",
      columns: 3,
      rows: [
        [{ paragraphs: [para("横跨两列")], spanCols: 2, widthTwips: 2000 }, { paragraphs: [para("c")], widthTwips: 1000 }],
        [{ paragraphs: [para("甲")], widthTwips: 1000 }, { paragraphs: [para("1")], widthTwips: 1000 }, { paragraphs: [para("x")], widthTwips: 1000 }],
        [{ paragraphs: [para("乙")], widthTwips: 1000 }, { paragraphs: [para("0")], widthTwips: 1000 }, { paragraphs: [para("y")], widthTwips: 1000 }],
      ],
      widthTwips: 3000,
      layout: "fixed",
      columnWidths: [1000, 1000, 1000],
    };
    const out = sort(t, { column: 1, kind: "number", headerRows: 1 });
    expect(out).toEqual({ ok: true, at: { row: 1, col: 1 } });
    expect(keys(t).slice(1)).toEqual(["乙", "甲"]);
  });

  it("拒绝竖合并：行一挪，按相邻行认领的合并就断", () => {
    const t = table(true, [["3", "三"], ["1", "一"]]);
    t.rows[1]![1]!.spanRows = 1;
    t.rows[2]![1]!.spanRows = 0;
    const out = sort(t, { headerRows: 1 });
    expect(out.ok).toBe(false);
    if (!out.ok) expect(out.reason).toContain("竖合并");
    expect(keys(t)).toEqual(["键", "3", "1"]);
  });

  it("拒绝行间标记与待判的表属性改动：它们的落点按行号记住", () => {
    const between = table(true, [["3", "三"], ["1", "一"]]);
    between.betweenRows = [{ afterRow: 0, tag: "w:bookmarkEnd", xml: '<w:bookmarkEnd w:id="1"/>' }];
    expect(sort(between, { headerRows: 1 }).ok).toBe(false);

    const pending = table(true, [["3", "三"], ["1", "一"]]);
    pending.revFormat = { kind: "format", tag: "tblPrChange", id: 1, author: "甲", date: "2026-01-01T00:00:00Z", prev: {}, prevSlots: [] };
    expect(sort(pending, { headerRows: 1 }).ok).toBe(false);
  });

  it("表体不足两行时不说「排好了」", () => {
    const one = table(true, [["1", "一"]]);
    const out = sort(one, { headerRows: 1 });
    expect(out.ok).toBe(false);
    if (!out.ok) expect(out.reason).toContain("只有一行");
  });

  it("标题行占满全表时不越界、也不排那一行：回绝并原样留下", () => {
    const t = table(true, [["3", "三"], ["1", "一"]]);
    t.rows.forEach((row) => row.forEach((cell) => (cell.row = { header: true })));
    const out = sort(t, { headerRows: 3, kind: "number" });
    expect(out.ok).toBe(false);
    if (!out.ok) expect(out.reason).toContain("只有一行");
    expect(keys(t)).toEqual(["键", "3", "1"]);
  });

  it("保存后表头那一句 w:tblHeader 不被排序写动", () => {
    const t = table(true, [["2", "二"], ["1", "一"]]);
    expect(sort(t, { headerRows: 1, kind: "number" }).ok).toBe(true);
    expect(saved(docWith(t))).toContain("<w:tblHeader/>");
  });
});
