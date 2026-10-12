import { describe, expect, it } from "vitest";
import type { Block, Table, TableCell } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
import {
  carriesUndecidedRevision,
  columnStarts,
  gridColumns,
  pasteAppendTableHere,
  pasteAsNestedTableHere,
} from "../src/core/doc-table-ops.js";
import {
  makeCellRevisionDocx,
  makeElementAttrsDocx,
  makeMergedTableDocx,
  makeNestedRevisedTableDocx,
  makeNestedTableDocx,
  makeOpsGridDocx,
  makeTableFidelityDocx,
  makeWideRootDocx,
} from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;

/** Two tables: the one being pasted into, and the one sitting on the clipboard. Each is read fresh. */
const grid = (): Table => tableOf(open(makeOpsGridDocx()));
const nested = (): Table => tableOf(open(makeNestedTableDocx()));
const merged = (): Table => tableOf(open(makeMergedTableDocx()));
const withDoc = (t: Table): WordDocument => {
  const doc = WordDocument.blank();
  doc.blocks.push(t);
  return doc;
};
/** A pasted shape has to survive the file, not only the model, so half of these checks read the bytes back. */
const reloaded = (t: Table): Table => tableOf(open(writeDocx(withDoc(t))));

const shape = (t: Table): string[] =>
  t.rows.map(
    (row, r) =>
      `${r}:${columnStarts(row).join(",")}=${row
        .map((c) => `${c.widthTwips ?? 0}x${c.spanCols ?? 1}${c.spanRows === undefined ? "" : `v${c.spanRows}`}`)
        .join(" ")}`,
  );
const kinds = (cell: TableCell): string => cell.paragraphs.map((b: Block) => b.kind).join(",");
const textOf = (t: Table): string =>
  t.rows
    .map((row) =>
      row
        .map((c) =>
          c.paragraphs
            .map((b) =>
              b.kind === "table" ? `[${textOf(b)}]` : b.kind === "paragraph" ? b.inlines.map((i) => i.text).join("") : `[${b.kind}]`,
            )
            .join(""),
        )
        .join("|"),
    )
    .join("/");
/** Every grid column of every row is covered exactly once — the rule a paste cannot break. */
const tilesTheGrid = (t: Table): boolean =>
  t.rows.every((row) => row.reduce((n, c) => n + Math.max(1, c.spanCols ?? 1), 0) === gridColumns(t));
const cellText = (c: TableCell): string =>
  c.paragraphs.map((b) => (b.kind === "paragraph" ? b.inlines.map((i) => i.text).join("") : "")).join("");

describe("粘贴为嵌套表（PasteAsNestedTable）", () => {
  it("puts the copied table inside the cell the caret is in and keeps its own grid", () => {
    const t = nested();
    const pasted = pasteAsNestedTableHere(t, { row: 0, col: 0 }, grid(), 1);
    expect(pasted.ok).toBe(true);
    const host = t.rows[0]![0]!.paragraphs;
    expect(host.map((b) => b.kind)).toEqual(["paragraph", "table", "paragraph"]);
    const inner = host[1] as Table;
    expect(inner.rows).toHaveLength(3);
    expect(inner.columnWidths).toEqual([1000, 2000, 3000]);
    expect(inner.gridXml).toBe(grid().gridXml);
    expect(textOf(inner)).toBe(textOf(grid()));
    // 落进去的那张表之外，一张表都没被碰过。
    expect(t.rows).toHaveLength(2);
    expect(t.columnWidths).toEqual([2000, 2000]);
    expect(kinds(t.rows[0]![1]!)).toBe(kinds(nested().rows[0]![1]!));
  });

  it("lands where the insertion point was, not at the bottom of the cell", () => {
    const t = nested();
    const blocks = t.rows[0]![1]!.paragraphs;
    expect(blocks.map((b) => b.kind)).toEqual(["paragraph", "table"]);
    // 光标在第一段上（块号 0）：新表跟着这一段，而不是排到既有那张表后面。
    expect(pasteAsNestedTableHere(t, { row: 0, col: 1 }, grid(), 1).ok).toBe(true);
    expect(blocks.map((b) => b.kind)).toEqual(["paragraph", "table", "paragraph", "table"]);
    expect((blocks[1] as Table).rows).toHaveLength(3);
    expect((blocks[3] as Table).rows).toHaveLength(2);
  });

  it("separates two tables in one cell and closes the cell with a paragraph, in the model", () => {
    const t = nested();
    expect(pasteAsNestedTableHere(t, { row: 0, col: 0 }, grid(), 1).ok).toBe(true);
    // 同一格再来一张，落在新表之后：两张表之间补一段，末尾也补一段。
    expect(pasteAsNestedTableHere(t, { row: 0, col: 0 }, merged(), 3).ok).toBe(true);
    const before = kinds(t.rows[0]![0]!);
    expect(before).toBe("paragraph,table,paragraph,table,paragraph");
    // 画面与存完重读回来的是同一份：写出器不再往格子里补人看不见的段落。
    expect(saved(withDoc(t))).not.toContain("</w:tbl></w:tc>");
    expect(kinds(tableOf(open(writeDocx(withDoc(t)))).rows[0]![0]!)).toBe(before);
  });

  it("saves and reads the nested paste back whole", () => {
    const t = nested();
    pasteAsNestedTableHere(t, { row: 0, col: 0 }, grid(), 1);
    expect((saved(withDoc(t)).match(/<w:tbl>/g) ?? []).length).toBe(3);
    const pasted = tableOf(open(writeDocx(withDoc(t)))).rows[0]![0]!.paragraphs[1] as Table;
    expect(textOf(pasted)).toBe(textOf(grid()));
    expect(pasted.rows[0]![0]!.tcPrXml).toContain('w:val="dotted"');
    expect(pasted.rows[2]![0]!.spanCols).toBe(2);
    expect(shape(pasted)).toEqual(shape(reloaded(grid())));
  });

  it("refuses a caret that is not in a cell, an empty copy, and a copy still under judgement", () => {
    const t = nested();
    const before = kinds(t.rows[0]![0]!);
    expect(pasteAsNestedTableHere(t, { row: 9, col: 0 }, grid())).toEqual({ ok: false, reason: "这里没有单元格" });

    const empty = grid();
    empty.rows = [];
    expect(pasteAsNestedTableHere(t, { row: 0, col: 0 }, empty)).toEqual({ ok: false, reason: "复制的表里没有行" });

    const revised = tableOf(open(makeCellRevisionDocx()));
    expect(carriesUndecidedRevision(revised)).toBe(true);
    const refused = pasteAsNestedTableHere(t, { row: 0, col: 0 }, revised);
    expect(refused.ok).toBe(false);
    expect((refused as { reason: string }).reason).toContain("w:id");
    expect(kinds(t.rows[0]![0]!)).toBe(before);

    // 待判记录藏在第二层表里也得拦：审阅的编号空间整份件共用。
    expect(carriesUndecidedRevision(tableOf(open(makeNestedRevisedTableDocx())))).toBe(true);
    expect(carriesUndecidedRevision(grid())).toBe(false);
  });

  it("does not carry the original paragraph's identity into the copy", () => {
    const withIds = tableOf(open(makeElementAttrsDocx()));
    const original = withIds.rows[0]![0]!.paragraphs[0];
    if (original.kind !== "paragraph") throw new Error("语料里的这格不是以段落开头");
    expect(original.attrsXml).toContain("paraId");

    const t = nested();
    pasteAsNestedTableHere(t, { row: 0, col: 0 }, withIds, 1);
    const pasted = t.rows[0]![0]!.paragraphs[1] as Table;
    const copy = pasted.rows[0]![0]!.paragraphs[0];
    if (copy.kind === "paragraph") expect(copy.attrsXml).toBeUndefined();
    // 原件的身份一个字都不动。
    expect(original.attrsXml).toContain("paraId");
  });
});

describe("粘贴为追加行（PasteAppendTable）", () => {
  it("adds the copied rows below and sizes them to the table they joined", () => {
    const t = nested();
    const pasted = pasteAppendTableHere(t, { row: 1, col: 0 }, merged());
    expect(pasted.ok).toBe(true);
    expect(t.rows).toHaveLength(5);
    // 等列宽追加：新的两格按目标的 `w:tblGrid` 走，源表那句 3000 不再是依据。
    expect(t.rows[2]!.map((c) => c.widthTwips)).toEqual([2000, 2000]);
    expect(t.rows[2]![1]!.tcPrXml).toContain('<w:tcW w:w="2000" w:type="dxa"/>');
    expect(t.rows[2]![1]!.tcPrXml).not.toContain('w:w="3000"');
    expect(t.gridXml).toBe(nested().gridXml);
    expect(t.rows[2]!.map(cellText)).toEqual(["占两行", "第一行右"]);
    expect(t.rows[4]!.map(cellText)).toEqual(["新的一格", "第三行右"]);
    expect(tilesTheGrid(t)).toBe(true);
    expect(renderWord(withDoc(t))).toContain('rowspan="2"');
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("hands a short row the leftover grid columns instead of leaving a hole", () => {
    const t = grid();
    const src = tableOf(open(makeWideRootDocx()));
    expect(gridColumns(src)).toBe(1);
    expect(pasteAppendTableHere(t, { row: 0, col: 0 }, src).ok).toBe(true);
    const row = t.rows[3]!;
    expect(row).toHaveLength(1);
    expect(row[0]).toMatchObject({ spanCols: 3, widthTwips: 6000 });
    expect(gridColumns(t)).toBe(3);
    expect(t.rows[0]!.map((c) => c.widthTwips)).toEqual([1000, 2000, 3000]);
    expect(tilesTheGrid(t)).toBe(true);
    // 没有 `w:tcPr` 的格子由写出器按模型字段补出 gridSpan，重读回来形状一致。
    expect(shape(reloaded(t))).toEqual(shape(t));
    expect(textOf(reloaded(t))).toContain("唯一的格");
  });

  it("refuses a copy wider than the table it would join, and changes nothing", () => {
    const t = nested();
    const before = shape(t);
    const refused = pasteAppendTableHere(t, { row: 0, col: 0 }, grid());
    expect(refused.ok).toBe(false);
    expect((refused as { reason: string }).reason).toContain("装不下");
    expect(shape(t)).toEqual(before);
    expect(t.rows).toHaveLength(2);
    expect(t.gridXml).toBe(nested().gridXml);
  });

  it("promotes a vertical merge that was cut in half on the way in", () => {
    const src = merged();
    src.rows.shift(); // 只抄到合并的下半截：续格在新表里没有上头
    const t = nested();
    expect(pasteAppendTableHere(t, { row: 0, col: 0 }, src).ok).toBe(true);
    expect(t.rows[2]![0]!.spanRows).toBe(1);
    expect(t.rows[3]![0]!.spanRows).toBe(1); // 源表第二个起点照旧
    const xml = saved(withDoc(t));
    expect((xml.match(/<w:vMerge/g) ?? [])).toHaveLength(2);
    expect(xml).not.toContain("<w:vMerge/>");
    expect(shape(reloaded(t))).toEqual(shape(t));
  });

  it("gives the joined row its own record, not a share of the original's", () => {
    const t = grid();
    const src = grid();
    const info = { heightTwips: 600, keepTogether: true, header: true };
    src.rows[0]!.forEach((c) => (c.row = info));
    expect(pasteAppendTableHere(t, { row: 0, col: 0 }, src).ok).toBe(true);
    const joined = t.rows[3]!;
    expect(joined[0]!.row).toEqual(info);
    expect(joined[0]!.row).toBe(joined[1]!.row);
    expect(joined[0]!.row).not.toBe(info);
    // 改新行不动源表那一行：两张表各有各的 `w:trPr`。
    joined[0]!.row!.heightTwips = 900;
    expect(info.heightTwips).toBe(600);
    expect(saved(withDoc(t))).toContain('<w:trHeight w:val="900"');
  });

  it("keeps a percent-wide copy measured against the text column, not the grid", () => {
    const t = grid();
    const src = grid();
    const pct = src.rows[0]![1]!;
    pct.widthType = "pct";
    pct.tcPrXml = '<w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr>';
    const before = pct.tcPrXml;
    expect(pasteAppendTableHere(t, { row: 0, col: 0 }, src).ok).toBe(true);
    const joined = t.rows[3]!;
    expect(joined[1]!.tcPrXml).toBe(before);
    expect(joined[1]!.widthType).toBe("pct");
    expect(joined[0]!.widthTwips).toBe(1000);
    expect(joined[2]!.widthTwips).toBe(3000);
    expect(saved(withDoc(t))).toContain('<w:tcW w:w="2500" w:type="pct"/>');
  });

  it("refuses a copy that carries an undecided change, for the same reason as a nested paste", () => {
    const t = grid();
    const before = shape(t);
    const refused = pasteAppendTableHere(t, { row: 0, col: 0 }, tableOf(open(makeTableFidelityDocx())));
    expect(refused.ok).toBe(false);
    expect((refused as { reason: string }).reason).toContain("w:id");
    expect(shape(t)).toEqual(before);
    expect(t.rows).toHaveLength(3);

    const empty = merged();
    empty.rows = [];
    expect(pasteAppendTableHere(t, { row: 0, col: 0 }, empty)).toEqual({ ok: false, reason: "复制的表里没有行" });
    expect(pasteAppendTableHere(t, { row: 9, col: 0 }, merged()).ok).toBe(false);
  });

  it("joins the destination's still-open merge band instead of inventing a second one", () => {
    // 抄到的第一行是个续格（`w:vMerge` 无值），而这张表最后一行正好起了一个没被续过的带。
    // ISO 的读法只有一个：续格接在它上面那格，于是那条带从一行长成两行——`spanRows` 只记身份
    // （1=起点、0=续格），带长是往下扫出来的，所以这里钉的是 `rowspan="2"` 而不是别的数。
    // WPS 究竟这样接，还是把没有同伴可接的续格升成起点，**未取证**。
    const t = merged();
    const src = merged();
    src.rows.shift();
    src.rows.pop();
    expect(pasteAppendTableHere(t, { row: 0, col: 0 }, src).ok).toBe(true);
    expect(t.rows.map((row) => row.map((c) => c.spanRows ?? 1).join("/"))).toEqual(["1/1", "0/1", "1/1", "0/1"]);
    expect(cellText(t.rows[2]![0]!)).toBe("新的一格");
    // 追加前这张表只有一条两行带；现在第二条带也接上了，画面上的带还是两条、各跨两行。
    expect(renderWord(withDoc(t)).match(/rowspan="2"/g)).toHaveLength(2);
    expect(saved(withDoc(t)).match(/<w:vMerge\/>/g)).toHaveLength(2);
    expect(shape(reloaded(t))).toEqual(shape(t));
  });
});
