import { describe, expect, it } from "vitest";
import type { RevisionEntry, Table, TableCell } from "../src/core/docx.js";
import { WordDocument, borderDelta, cellShapeOf, signCellRevision, signTableRevision, syncCellBox, syncTableBox, tableShapeOf } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { TABLE_EDGE_SHAPE, applyBoxPatch } from "../src/ui/doc-para-border.js";
import type { BoxPatch, CellBoxSide, TableBoxSide } from "../src/ui/doc-para-border.js";
import { pack } from "./helpers/make-containers.js";
import {
  makeCellRevisionDocx,
  makePartialCellInsDocx,
  makeRowSignatureDocx,
  makeSoleColumnDocx,
  makeTablePropertyRevisionDocx,
} from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;
/** What the table's cells hold, one entry per row, its cells joined by `|`. */
const cellTexts = (t: Table): string[] =>
  t.rows.map((r) =>
    r.map((c) => c.paragraphs.map((p) => (p.kind === "paragraph" ? p.inlines.map((i) => i.text).join("") : "")).join("")).join("|"),
  );
/** The 审阅窗格's row, in the order it lists them: which change, whose, and what it is about. */
const listed = (doc: WordDocument): string =>
  doc.revisionEntries().map((e) => `${e.index}.${e.type}:${e.id}:${e.author}:${e.text}`).join(",");
const hits = (xml: string, needle: string): number => xml.split(needle).length - 1;

/** A table whose parts each carry somebody else's undecided property change. */
describe("docx 修订：表格部件上的属性改动", () => {
  it("lists the table's, the row's and the cell's record as three changes, in document order", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    expect(listed(doc)).toBe("1.table-format:61:林晓:甲 乙 丙 丁,2.row-format:62:顾峰:甲 乙,3.cell-format:63:林晓:甲");
    // A part has no characters of its own to range over, so the change is placed at its first paragraph.
    expect(doc.revisionEntries().map((e) => `${e.para}:${e.begin}-${e.end}`)).toEqual(["1:0-0", "1:0-0", "1:0-0"]);
  });

  it("reads the recorded set as the fields the page draws", () => {
    const t = tableOf(open(makeTablePropertyRevisionDocx()));
    expect(t.revFormat).toMatchObject({ id: 61, author: "林晓", prev: { widthTwips: 0, widthType: "auto" } });
    expect(t.widthTwips).toBe(5000);
    // `w:tblCellMar` is a field the page draws, so the record's copy of it is read as one and not only carried.
    expect(t.cellMargin).toEqual({ left: 108 });
    expect(t.revFormat!.prev.cellMargin).toEqual({ left: 72 });
    const row = t.rows[0][0]!.row!;
    expect(row.revFormat).toMatchObject({ id: 62, author: "顾峰", prev: { heightTwips: 300, keepTogether: true } });
    expect(row).toMatchObject({ heightTwips: 600, heightRule: "exact", header: true });
    expect(t.rows[0][0]!.revs).toMatchObject([{ tag: "tcPrChange", id: 63, prev: { widthTwips: 1500, shading: "#DDEBF7" } }]);
    expect(t.rows[0][0]).toMatchObject({ widthTwips: 2000, shading: "#FFF2CC", margin: { top: 60 } });
    expect(t.rows[0][0]!.revs![0]!.prev.margin).toEqual({ top: 120 });
  });

  it("reads a record whether the set it holds is wrapped in its own element or written straight", () => {
    const t = tableOf(open(makeTablePropertyRevisionDocx()));
    // `w:tblPrChange` and `w:trPrChange` as MS's example spells them, `w:tcPrChange` the way Word's own
    // fixtures here do. Both arrive as the same kind of list of properties.
    expect(t.revFormat!.prevSlots.map((s) => s.tag)).toEqual(["tblW", "tblCellMar"]);
    expect(t.rows[0][0]!.row!.revFormat!.prevSlots.map((s) => s.tag)).toEqual(["cantSplit", "trHeight"]);
    expect(t.rows[0][0]!.revs![0]!.prevSlots.map((s) => s.tag)).toEqual(["tcW", "shd", "tcMar"]);
  });

  it("accepting the table's change keeps what the table looks like now and takes the record out", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    expect(doc.resolveRevisions("accept", 61)).toBe(1);
    const out = saved(doc);
    expect(out).toContain('<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>');
    expect(out).not.toContain("<w:tblPrChange");
    expect(tableOf(doc).widthType).toBe("pct");
    // The two changes nobody has decided yet are untouched by this one.
    expect(listed(doc)).toBe("1.row-format:62:顾峰:甲 乙,2.cell-format:63:林晓:甲");
  });

  it("refusing it puts the complete recorded set back, properties the model has no field for included", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    doc.resolveRevisions("reject", 61);
    const out = saved(doc);
    // ISO/IEC 29500 on a `w:*PrChange`: its child holds the complete set applied before the revision. So the
    // table goes back to auto width *and* the 72-twip margin the record names, not to a merge of old and new.
    expect(out).toContain('<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellMar><w:left w:w="72" w:type="dxa"/></w:tblCellMar></w:tblPr>');
    expect(out).not.toContain("<w:tblPrChange");
    expect(tableOf(doc)).toMatchObject({ widthTwips: 0, widthType: "auto" });
    expect(listed(doc)).toBe("1.row-format:62:顾峰:甲 乙,2.cell-format:63:林晓:甲");
  });

  it("makes the page itself follow the decision, not only the bytes", () => {
    // 边距现在是画得出的一层：判定改了模型，纸上的 padding 就得跟着改，不然「记录带回」只带回一半。
    const asSaid = renderWord(open(makeTablePropertyRevisionDocx()));
    expect(asSaid).toContain("padding-top:4px;padding-left:7.2px;");
    const tableBack = open(makeTablePropertyRevisionDocx());
    tableBack.resolveRevisions("reject", 61);
    expect(renderWord(tableBack)).toContain("padding-top:4px;padding-left:4.8px;");
    const cellBack = open(makeTablePropertyRevisionDocx());
    cellBack.resolveRevisions("reject", 63);
    expect(renderWord(cellBack)).toContain("padding-top:8px;padding-left:7.2px;");
  });

  it("refusing a row's change puts its height back and leaves the row's own signature alone", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    expect(doc.resolveRevisions("reject", 62)).toBe(1);
    const out = saved(doc);
    expect(out).toContain("<w:trPr><w:cantSplit/><w:trHeight w:val=\"300\"/></w:trPr>");
    expect(out).not.toContain("<w:trPrChange");
    expect(out).toContain("<w:tblPrChange");
    const row = tableOf(doc).rows[0][0]!.row!;
    expect(row).toMatchObject({ heightTwips: 300, keepTogether: true });
    // The recorded set was the row's whole earlier shape, so what it did not name is gone rather than unset.
    expect(row.heightRule).toBeUndefined();
    expect(row.header).toBeUndefined();
  });

  it("accepting a row's change leaves every other property the row carries", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    doc.resolveRevisions("accept", 62);
    const row = tableOf(doc).rows[0][0]!.row!;
    expect(row.slots?.map((s) => s.tag)).toEqual(["trHeight", "tblHeader"]);
    expect(saved(doc)).toContain('<w:trHeight w:val="600" w:hRule="exact"/>');
    expect(listed(doc)).toBe("1.table-format:61:林晓:甲 乙 丙 丁,2.cell-format:63:林晓:甲");
  });

  it("a row's format record does not duplicate the signature the row already carries", () => {
    const doc = open(makeRowSignatureDocx());
    expect(doc.resolveRevisions("reject", 52)).toBe(1);
    const trPr = tableOf(doc).rows[0][0]!.row!;
    // The recorded set is the row's whole earlier shape: height goes, cantSplit comes back, and the row is
    // still an insertion nobody has decided on — one `w:ins`, not two.
    expect(trPr).toMatchObject({ rev: { kind: "ins", id: 51 }, keepTogether: true });
    expect(trPr.heightTwips).toBeUndefined();
    const out = saved(doc);
    expect(hits(out, "<w:ins ")).toBe(1);
    expect(out).toContain("<w:cantSplit/>");
    expect(out).not.toContain("w:val=\"700\"");
    expect(listed(doc)).toBe("1.row-insert:51:顾峰:加进来的一行");
  });

  it("refusing a cell's change restores the cell's shape without disturbing its neighbour", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    expect(doc.resolveRevisions("reject", 63)).toBe(1);
    const out = saved(doc);
    const cell = tableOf(doc).rows[0][0]!;
    expect(cell).toMatchObject({ widthTwips: 1500, shading: "#DDEBF7" });
    expect(out).toContain('<w:tcPr><w:tcW w:w="1500" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="DDEBF7"/><w:tcMar><w:top w:w="120" w:type="dxa"/></w:tcMar></w:tcPr>');
    expect(out).toContain('<w:cnfStyle w:val="000000000100"/>');
    expect(out).not.toContain("<w:tcPrChange");
  });

  it("leaves nothing pending once every record is decided, and the file reads back the same", () => {
    for (const outcome of ["accept", "reject"] as const) {
      const doc = open(makeTablePropertyRevisionDocx());
      expect(doc.resolveRevisions(outcome)).toBe(3);
      expect(doc.revisionEntries()).toEqual([]);
      const out = saved(doc);
      expect(hits(out, "PrChange")).toBe(0);
      // The band's own record — see the test below — is the one change this does not answer for.
      expect(open(writeDocx(doc)).revisionEntries()).toEqual([]);
      expect(out).toContain("<w:cnfStyle");
    }
  });

  it("carries a `w:tblPrExChange` through every decision but lists and decides neither", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    expect(doc.revisionEntries().map((e) => e.id)).not.toContain(64);
    doc.resolveRevisions("accept");
    const out = saved(doc);
    // A limit, stated as one: the band's record is not lost, and it is not silently resolved either.
    expect(out).toContain('<w:tblPrExChange w:id="64" w:author="顾峰"');
    expect(out).toContain('w:fill="FDE9D9"');
    expect(out).toContain('w:fill="E6EEF6"');
  });
});

/** A cell that arrived, is on its way out, or changed its merge — the records that move a cell. */
describe("docx 修订：单元格的位置与合并", () => {
  it("lists one change per `w:id`, however many cells carry it", () => {
    const doc = open(makeCellRevisionDocx());
    expect(listed(doc)).toBe(
      "1.cell-merge:71:林晓:甲一,2.cell-format:72:顾峰:甲一,3.cell-delete:73:林晓:乙一乙二,4.cell-insert:70:顾峰:丙一丙二",
    );
  });

  it("keeps two records that sit on one cell apart", () => {
    const doc = open(makeCellRevisionDocx());
    const cell = tableOf(doc).rows[0][0]!;
    expect(cell.revs!.map((r) => `${r.tag}:${r.id}`)).toEqual(["cellMerge:71", "tcPrChange:72"]);
    expect(cell.revs!.find((r) => r.id === 71)!.prev).toEqual({ spanRows: 0 });
  });

  it("accepting an inserted column lets the column stay and clears its marks", () => {
    const doc = open(makeCellRevisionDocx());
    expect(doc.resolveRevisions("accept", 70)).toBe(2);
    const t = tableOf(doc);
    expect(t.columns).toBe(3);
    expect(t.rows.map((r) => r.length)).toEqual([3, 3]);
    const out = saved(doc);
    expect(hits(out, "<w:cellIns")).toBe(0);
    expect(hits(out, "<w:gridCol")).toBe(3);
    expect(out).toContain("丙一");
    expect(listed(doc)).toBe("1.cell-merge:71:林晓:甲一,2.cell-format:72:顾峰:甲一,3.cell-delete:73:林晓:乙一乙二");
  });

  it("refusing an inserted column takes the column, its grid entry and its text out", () => {
    const doc = open(makeCellRevisionDocx());
    expect(doc.resolveRevisions("reject", 70)).toBe(2);
    const t = tableOf(doc);
    expect(cellTexts(t)).toEqual(["甲一|乙一", "甲二|乙二"]);
    expect(t.columns).toBe(2);
    expect(t.columnWidths).toEqual([1000, 2000]);
    const out = saved(doc);
    expect(hits(out, "<w:gridCol")).toBe(2);
    expect(out).not.toContain("丙");
    expect(hits(out, "<w:cellIns")).toBe(0);
    // The other reviewer's changes are still theirs to decide.
    expect(out).toContain("<w:cellDel");
    expect(out).toContain("<w:cellMerge");
  });

  it("accepting a deleted column takes it out and refusing it leaves the cell where it is", () => {
    const doc = open(makeCellRevisionDocx());
    expect(doc.resolveRevisions("accept", 73)).toBe(2);
    const t = tableOf(doc);
    expect(t.columns).toBe(2);
    expect(t.columnWidths).toEqual([1000, 3000]);
    expect(saved(doc)).not.toContain("乙");
    expect(listed(doc)).toBe("1.cell-merge:71:林晓:甲一,2.cell-format:72:顾峰:甲一,3.cell-insert:70:顾峰:丙一丙二");

    const kept = open(makeCellRevisionDocx());
    expect(kept.resolveRevisions("reject", 73)).toBe(2);
    expect(tableOf(kept).columns).toBe(3);
    const out = saved(kept);
    expect(out).toContain("乙一");
    expect(hits(out, "<w:cellDel")).toBe(0);
    expect(hits(out, "<w:gridCol")).toBe(3);
  });

  it("refusing a merge change puts back the one property it named, leaving the width alone", () => {
    const doc = open(makeCellRevisionDocx());
    expect(doc.resolveRevisions("reject", 71)).toBe(1);
    const cell = tableOf(doc).rows[0][0]!;
    expect(cell).toMatchObject({ widthTwips: 1000, spanRows: 0 });
    const out = saved(doc);
    expect(out).toContain('<w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:vMerge/><w:tcPrChange');
    expect(out).not.toContain('<w:tcW w:w="1000" w:type="dxa"/><w:vMerge w:val="restart"');
    expect(listed(doc)).toBe("1.cell-format:72:顾峰:甲一,2.cell-delete:73:林晓:乙一乙二,3.cell-insert:70:顾峰:丙一丙二");
  });

  it("accepting a merge change keeps the merge the file settled on", () => {
    const doc = open(makeCellRevisionDocx());
    expect(doc.resolveRevisions("accept", 71)).toBe(1);
    const cell = tableOf(doc).rows[0][0]!;
    expect(cell).toMatchObject({ widthTwips: 1000, spanRows: 1 });
    const out = saved(doc);
    expect(out).toContain('<w:vMerge w:val="restart"/><w:tcPrChange');
    expect(hits(out, "<w:cellMerge")).toBe(0);
  });

  it("refusing a cell's width change restores the whole recorded set, merge included", () => {
    const doc = open(makeCellRevisionDocx());
    expect(doc.resolveRevisions("reject", 72)).toBe(1);
    const cell = tableOf(doc).rows[0][0]!;
    expect(cell).toMatchObject({ widthTwips: 1200, spanRows: 1 });
    expect(saved(doc)).toContain('<w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:vMerge w:val="restart"/><w:cellMerge');
    expect(listed(doc)).toBe("1.cell-merge:71:林晓:甲一,2.cell-delete:73:林晓:乙一乙二,3.cell-insert:70:顾峰:丙一丙二");
  });

  it("settles a revised table end to end and reads back clean", () => {
    for (const outcome of ["accept", "reject"] as const) {
      const doc = open(makeCellRevisionDocx());
      expect(doc.resolveRevisions(outcome)).toBe(6);
      expect(doc.revisionEntries()).toEqual([]);
      const out = saved(doc);
      expect(hits(out, "<w:cellIns")).toBe(0);
      expect(hits(out, "<w:cellDel")).toBe(0);
      expect(hits(out, "<w:cellMerge")).toBe(0);
      expect(hits(out, "<w:tcPrChange")).toBe(0);
      expect(open(writeDocx(doc)).revisionEntries()).toEqual([]);
      // Whichever way it went, the grid and the cells still agree with each other.
      const t = tableOf(open(writeDocx(doc)));
      expect(t.columns).toBe(2);
      for (const row of t.rows) expect(row.length).toBe(2);
    }
  });

  it("keeps a record it would have to answer by emptying the table", () => {
    const doc = open(makeSoleColumnDocx());
    expect(listed(doc)).toBe("1.cell-delete:80:顾峰:唯一的列");
    expect(doc.resolveRevisions("accept", 80)).toBe(0);
    // A one-column table has no column left to remove, so the change is answered with nothing rather than
    // half-applied — mark gone, cell still there. The pane sees the same thing through its own entry.
    expect(doc.resolveRevisionAt("accept", 1)).toBe(false);
    expect(listed(doc)).toBe("1.cell-delete:80:顾峰:唯一的列");
    expect(tableOf(doc).rows[0]).toHaveLength(1);
    expect(saved(doc)).toContain("<w:cellDel");
  });

  it("refusing a lone cell insertion takes that cell and not the column it sits in", () => {
    const doc = open(makePartialCellInsDocx());
    expect(listed(doc)).toBe("1.cell-insert:90:顾峰:甲二");
    expect(doc.resolveRevisions("reject")).toBe(1);
    const t = tableOf(doc);
    // The cell above it was nobody's to decide, so it is still there — and so is the grid it lines up with.
    expect(cellTexts(t)).toEqual(["甲一|乙一", "乙二"]);
    expect(t.columns).toBe(2);
    expect(t.columnWidths).toEqual([2000, 2000]);
    // The row keeps tiling the grid by giving the freed column to its last cell, which is Word's
    // "delete cells, shift cells left" rather than a hole the page has to guess at.
    expect(t.rows[1][0]).toMatchObject({ spanCols: 2, widthTwips: 4000 });
    const out = saved(doc);
    expect(out).not.toContain("甲二");
    expect(hits(out, "<w:cellIns")).toBe(0);
    expect(hits(out, "<w:gridSpan")).toBe(1);
    expect(doc.revisionEntries()).toEqual([]);
    expect(open(writeDocx(doc)).revisionEntries()).toEqual([]);
  });

  it("accepting a lone cell insertion leaves every cell where the reader put it", () => {
    const doc = open(makePartialCellInsDocx());
    expect(doc.resolveRevisions("accept")).toBe(1);
    const t = tableOf(doc);
    expect(cellTexts(t)).toEqual(["甲一|乙一", "甲二|乙二"]);
    expect(t.columns).toBe(2);
    const out = saved(doc);
    expect(hits(out, "<w:cellIns")).toBe(0);
    expect(hits(out, "<w:gridSpan")).toBe(0);
    expect(open(writeDocx(doc)).revisionEntries()).toEqual([]);
  });
});

/** The pane reaches these through the same three entries as every other change. */
describe("docx 修订：窗格怎么找到表格上的改动", () => {
  it("answers with the table's own record when asked what waits at its first cell", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    const entry: RevisionEntry = doc.revisionEntries()[2]!;
    // The pane indexes paragraphs the way `revisionEntries()` walks them, table cells included.
    expect(doc.revisionIndexAt(doc.paragraphs()[1]!, 0)).toBe(1);
    expect(entry.type).toBe("cell-format");
    expect(doc.resolveRevisionAt("accept", 3)).toBe(true);
    expect(listed(doc)).toBe("1.table-format:61:林晓:甲 乙 丙 丁,2.row-format:62:顾峰:甲 乙");
    expect(doc.resolveRevisionAt("accept", 99)).toBe(false);
  });

  it("decides every record of one author at once", () => {
    const doc = open(makeTablePropertyRevisionDocx());
    expect(doc.resolveRevisionsByAuthor("reject", "林晓")).toBe(2);
    expect(listed(doc)).toBe("1.row-format:62:顾峰:甲 乙");
    const out = saved(doc);
    expect(out).toContain('<w:tblW w:w="0" w:type="auto"/>');
    expect(out).toContain('<w:tcW w:w="1500" w:type="dxa"/>');
    expect(out).not.toContain('<w:tblW w:w="5000"');
  });

  it("counts a revised column as one change for one author, wherever its cells sit", () => {
    const doc = open(makeCellRevisionDocx());
    // One row of the pane, one answer: the column is 顾峰's single change, the widening another.
    expect(doc.resolveRevisionsByAuthor("accept", "顾峰")).toBe(2);
    expect(listed(doc)).toBe("1.cell-merge:71:林晓:甲一,2.cell-delete:73:林晓:乙一乙二");
  });
});

/**
 * 上面那几族读的是别人写好的件；这一族写的是我们自己的编辑动作（#149）。开着修订时 ribbon 的「边框与底纹」
 * 落在表那一圈（`w:tblBorders`、`w:tblPr/w:shd`）或某几格那一圈（`w:tcBorders`、`w:tcPr/w:shd`），
 * `signTableRevision`/`signCellRevision` 签出一枚 §17.13.5 的记录，两份都得留：
 * **文字**挂进带着的 `w:tblPr`/`w:tcPr`（`decideSetXml` 与写出读的都是这一枚孩子，模型没有字段的 `w:tblLook`、
 * `w:tcMar` 连自己的字一起回去），**字段**是给「拒绝」把属性捡回模型用的。只留一份就是那条静默交回原字节的坑。
 *
 * 落笔形状照 ECMA-376 §17.13.5 与本仓库段落那两条（`w:pPrChange`）的同一口径：记录只说部件**自己说过**的属性，
 * 从 `w:tblStyle` 借来的不进快照，本来什么也没说时记的是一枚空套。**WPS 开着修订改表格线写出的件仍未取证**
 * （`.scratch/wps-corpus` 里 0 枚表格修订），所以既不替没动过的边补属性，也不编 WPS 的拼法。
 */
describe("docx 修订：表与格那一圈的改动由我们签出（w:tblPrChange / w:tcPrChange）", () => {
  const STAMP = "2026-10-16T00:00:00Z";

  /** 表与格各把自己说过的一套写全，孩子之间留着换行与空格；没有一枚待裁决的记录。 */
  const SAID = pack(
    `<w:tbl><w:tblPr>
  <w:tblStyle w:val="a3"/>
  <w:tblW w:w="0" w:type="auto"/>
  <w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>
  <w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar>
  <w:tblLook w:val="04A0" w:firstRow="1"/>
</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`,
  );

  /** 一张没写过 `w:tblPr`、一格没写过 `w:tcPr` 的表——与编辑器自己造的部件同一条件：没有原件文字可挂。 */
  const BARE = pack(
    `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t>独</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
  );

  /** 件里某一枚孩子的整枚原文（连它自己的孩子）：第几枚按文档顺序数。 */
  const setOf = (xml: string, tag: string, nth = 0): string =>
    [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>|<w:${tag}\\b[^>]*/>`, "g"))][nth]?.[0] ?? "";

  /** 与 `setTableBox` 同几步：先比、落模型、只把动过的边写回件里那一枚孩子，最后按 `revSeq` 那本账签一枚。 */
  const clickTable = (doc: WordDocument, rev: { n: number }, author: string, patch: BoxPatch<TableBoxSide>, ground?: string): Table => {
    const t = tableOf(doc);
    const after = applyBoxPatch(t.borders, patch, TABLE_EDGE_SHAPE);
    const changes = borderDelta(t.borders, after);
    const wasShape = tableShapeOf(t);
    const wasXml = t.tblPrXml;
    t.borders = after;
    if (ground !== undefined) t.shading = ground;
    syncTableBox(t, changes, ground);
    if (t.revFormat?.author !== author) {
      rev.n += 1;
      signTableRevision(t, { id: rev.n, author, date: STAMP }, wasShape, wasXml);
    }
    return t;
  };

  /** 格那一腿：与 `setCellBox` 同几步，逐格比、逐格写、逐格签。 */
  const clickCell = (doc: WordDocument, rev: { n: number }, author: string, patch: BoxPatch<CellBoxSide>, r = 0, c = 0, ground?: string): TableCell => {
    const cell = tableOf(doc).rows[r][c]!;
    const after = applyBoxPatch(cell.borders, patch, TABLE_EDGE_SHAPE);
    const changes = borderDelta(cell.borders, after);
    const wasShape = cellShapeOf(cell);
    const wasXml = cell.tcPrXml;
    cell.borders = after;
    if (ground !== undefined) cell.shading = ground;
    syncCellBox(cell, changes, ground);
    if (!(cell.revs ?? []).some((x) => x.tag === "tcPrChange" && x.author === author)) {
      rev.n += 1;
      signCellRevision(cell, { id: rev.n, author, date: STAMP }, wasShape, wasXml);
    }
    return cell;
  };

  it("签出的一枚表级记录：文字挂进件里那一枚孩子，记的是改之前自己说过的全部", () => {
    const doc = open(SAID);
    const t = clickTable(doc, { n: 80 }, "顾峰", { sides: { bottom: { val: "single" } } });
    expect(t.revFormat).toMatchObject({ kind: "format", tag: "tblPrChange", id: 81, author: "顾峰", date: STAMP });
    const rec = setOf(saved(doc), "tblPrChange");
    expect(rec.startsWith(`<w:tblPrChange w:id="81" w:author="顾峰" w:date="${STAMP}"><w:tblPr>`)).toBe(true);
    // 「改之前」是那一整套，不是这一击动过的那一条边：模型没有字段的 `w:tblLook`、`w:tblCellMar` 也在里面。
    expect(rec).toContain('<w:tblStyle w:val="a3"/>');
    expect(rec).toContain('<w:tblW w:w="0" w:type="auto"/>');
    expect(rec).toContain('<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    expect(rec).toContain('<w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar>');
    expect(rec).toContain('<w:tblLook w:val="04A0" w:firstRow="1"/>');
    // 这一击新说的那一条只活在现在那一套里，不在记录里。
    expect(rec).not.toContain("<w:bottom");
    const out = saved(doc);
    // 记录摆 `CT_TblPr`（§17.4.60）给它的位子：孩子的末尾，不是随手挤到 `w:tblStyle` 前面。
    expect(out.indexOf("<w:tblLook")).toBeLessThan(out.indexOf("<w:tblPrChange"));
    expect(setOf(out, "tblBorders")).toContain('<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    expect(setOf(out, "tblBorders")).toContain('<w:top w:val="single"');
  });

  it("存下去再读开，别人读到的还是同一枚待裁决：号、人、那一套都在", () => {
    const doc = open(SAID);
    clickTable(doc, { n: 80 }, "顾峰", { sides: { bottom: { val: "single" } } });
    const again = open(writeDocx(doc));
    expect(again.revisionEntries().map((e) => `${e.type}:${e.id}:${e.author}`)).toEqual(["table-format:81:顾峰"]);
    const t = tableOf(again);
    expect(t.revFormat!.prev).toEqual({ styleId: "a3", widthTwips: 0, widthType: "auto", borders: { top: { val: "single", size: 4, space: 0, color: "auto" } }, cellMargin: { left: 108 } });
    expect(t.revFormat!.prevSlots.map((s) => s.tag)).toEqual(["tblStyle", "tblW", "tblBorders", "tblCellMar", "tblLook"]);
    expect(t).toMatchObject({ widthType: "auto", shading: undefined });
  });

  it("拒绝把「改之前」那一整套放回去，含模型没有字段的 `w:tblLook`", () => {
    const doc = open(SAID);
    clickTable(doc, { n: 80 }, "顾峰", { sides: { bottom: { val: "single" } } });
    expect(doc.resolveRevisions("reject", 81)).toBe(1);
    const t = tableOf(doc);
    expect(t.borders).toEqual({ top: { val: "single", size: 4, space: 0, color: "auto" } });
    const out = saved(doc);
    expect(out).not.toContain("<w:tblPrChange");
    expect(setOf(out, "tblBorders")).not.toContain("<w:bottom");
    expect(out).toContain('<w:tblLook w:val="04A0" w:firstRow="1"/>');
    expect(out).toContain('<w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar>');
    expect(doc.revisionEntries()).toEqual([]);
    expect(open(writeDocx(doc)).revisionEntries()).toEqual([]);
  });

  it("接受留下现在这一套，摘掉的只有记录", () => {
    const doc = open(SAID);
    clickTable(doc, { n: 80 }, "顾峰", { sides: { bottom: { val: "single" } } });
    expect(doc.resolveRevisions("accept", 81)).toBe(1);
    const out = saved(doc);
    expect(out).not.toContain("<w:tblPrChange");
    expect(setOf(out, "tblBorders")).toContain('<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    expect(out).toContain('<w:tblLook w:val="04A0" w:firstRow="1"/>');
    expect(tableOf(doc).borders!.bottom).toMatchObject({ val: "single" });
  });

  it("同一个人第二击不再另发一个号，一次拒绝退掉两击", () => {
    const rev = { n: 80 };
    const doc = open(SAID);
    clickTable(doc, rev, "顾峰", { sides: { bottom: { val: "single" } } });
    clickTable(doc, rev, "顾峰", { sides: { right: { val: "single" } } });
    expect(rev.n).toBe(81);
    const out = saved(doc);
    expect(hits(out, "<w:tblPrChange")).toBe(1);
    expect(out).not.toContain('w:id="82"');
    // 记录说的是**第一击之前**那一句，所以拒绝之后两条新边一起没了。
    doc.resolveRevisions("reject", 81);
    const back = saved(doc);
    expect(setOf(back, "tblBorders")).not.toContain("<w:bottom");
    expect(setOf(back, "tblBorders")).not.toContain("<w:right");
    expect(setOf(back, "tblBorders")).toContain("<w:top");
  });

  it("另一个人记过的这一枚给新的一枚让位：`CT_TblPr` 只容一枚记录", () => {
    const rev = { n: 80 };
    const doc = open(makeTablePropertyRevisionDocx());
    clickTable(doc, rev, "顾峰", { sides: { bottom: { val: "single" } } });
    const out = saved(doc);
    expect(hits(out, "<w:tblPrChange")).toBe(1);
    expect(out).toContain('w:id="81" w:author="顾峰"');
    expect(out).not.toContain('w:id="61"');
    expect(doc.revisionEntries().filter((e) => e.type === "table-format").map((e) => `${e.id}:${e.author}`)).toEqual(["81:顾峰"]);
    // 快照问的是这一击之前表说什么，与林晓那一句无关：她的记录不会变成「改之前」的一部分——
    // 套里再套一枚记录的话，上面那句「整篇只有一枚 `w:tblPrChange`」就先数不过来了。
  });

  it("没有 `w:tblPr` 的表：字段那一份照样写出记录，记的是一枚空套", () => {
    const doc = open(BARE);
    const t = clickTable(doc, { n: 80 }, "顾峰", { sides: { bottom: { val: "single" } } });
    expect(t.tblPrXml).toBeUndefined();
    const rec = setOf(saved(doc), "tblPrChange");
    expect(rec).toContain('w:id="81"');
    expect(rec).toContain("<w:tblPr/>");
    expect(saved(doc)).toContain("<w:tblBorders>");
    // 这张表改之前什么也没说，所以拒绝之后什么也不说：回到向样式借那一边。
    doc.resolveRevisions("reject", 81);
    const out = saved(doc);
    expect(out).not.toContain("<w:tblPrChange");
    expect(out).not.toContain("<w:tblBorders");
    expect(tableOf(doc).borders).toBeUndefined();
  });

  it("填了一地的表：拒绝把地撤掉，其余孩子一个字节不动", () => {
    const doc = open(SAID);
    const t = clickTable(doc, { n: 80 }, "顾峰", {}, "#DEEBF7");
    expect(t.shading).toBe("#DEEBF7");
    expect(t.revFormat!.prev.shading).toBeUndefined();
    expect(doc.resolveRevisions("reject", 81)).toBe(1);
    const after = tableOf(doc);
    expect(after.shading).toBeUndefined();
    const out = saved(doc);
    expect(out).not.toContain("<w:shd");
    expect(out).toContain('<w:tblLook w:val="04A0" w:firstRow="1"/>');
    expect(renderWord(doc)).toContain("padding-left:7.2px;");
  });

  it("格那一腿同两分：记录挂进这一格的 `w:tcPr`，`w:tcMar` 连自己的字一起回去", () => {
    const doc = open(SAID);
    const cell = clickCell(doc, { n: 90 }, "顾峰", { sides: { bottom: { val: "single" } } });
    expect(cell.revs).toMatchObject([{ tag: "tcPrChange", id: 91, author: "顾峰" }]);
    const rec = setOf(saved(doc), "tcPrChange");
    expect(rec.startsWith('<w:tcPrChange w:id="91" w:author="顾峰" w:date="2026-10-16T00:00:00Z"><w:tcPr>')).toBe(true);
    expect(rec).toContain('<w:tcW w:w="2000" w:type="dxa"/>');
    expect(rec).toContain('<w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar>');
    expect(rec).toContain('<w:vAlign w:val="bottom"/>');
    expect(rec).not.toContain("<w:bottom");
    const out = saved(doc);
    expect(out.indexOf("<w:vAlign")).toBeLessThan(out.indexOf("<w:tcPrChange"));
    expect(setOf(out, "tcBorders")).toContain('<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    // 邻格一条也没点，就不该被牵连：它没有记录，也没有新说的边。
    expect(tableOf(doc).rows[0][1]!.revs ?? []).toEqual([]);
  });

  it("拒绝一格记录的把整套捡回来，接受则只留现在这一套", () => {
    const back = open(SAID);
    clickCell(back, { n: 90 }, "顾峰", { sides: { bottom: { val: "single" } } });
    expect(back.resolveRevisions("reject", 91)).toBe(1);
    const cell = tableOf(back).rows[0][0]!;
    expect(cell).toMatchObject({ widthTwips: 2000, valign: "bottom", margin: { top: 60 } });
    expect(cell.borders).toBeUndefined();
    const out = saved(back);
    expect(out).not.toContain("<w:tcPrChange");
    expect(out).toContain('<w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar>');
    expect(out).toContain('<w:vAlign w:val="bottom"/>');

    const stay = open(SAID);
    clickCell(stay, { n: 90 }, "顾峰", { sides: { bottom: { val: "single" } } });
    expect(stay.resolveRevisions("accept", 91)).toBe(1);
    const kept = saved(stay);
    expect(kept).not.toContain("<w:tcPrChange");
    expect(setOf(kept, "tcBorders")).toContain('<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    expect(kept).toContain('<w:vAlign w:val="bottom"/>');
  });

  it("同一个人连着点两击只留一条记录，别人记过的那一枚随新的让位", () => {
    const rev = { n: 90 };
    const doc = open(SAID);
    clickCell(doc, rev, "顾峰", { sides: { bottom: { val: "single" } } });
    clickCell(doc, rev, "顾峰", { sides: { top: { val: "dashed" } } });
    expect(rev.n).toBe(91);
    const out = saved(doc);
    expect(hits(out, "<w:tcPrChange")).toBe(1);
    doc.resolveRevisions("reject", 91);
    expect(setOf(saved(doc), "tcBorders")).toBe("");
    expect(tableOf(doc).rows[0][0]!.borders).toBeUndefined();
  });

  it("没有 `w:tcPr` 的格：记录写得出，拒绝之后这一格什么也不说", () => {
    const doc = open(BARE);
    const cell = clickCell(doc, { n: 90 }, "顾峰", { sides: { bottom: { val: "single" } } });
    expect(cell.tcPrXml).toBeUndefined();
    const rec = setOf(saved(doc), "tcPrChange");
    expect(rec).toContain('w:id="91"');
    expect(rec).toContain("<w:tcPr/>");
    expect(doc.resolveRevisions("reject", 91)).toBe(1);
    const out = saved(doc);
    expect(out).not.toContain("<w:tcPrChange");
    expect(out).not.toContain("<w:tcBorders");
    expect(tableOf(doc).rows[0][0]!.borders).toBeUndefined();
  });

  it("存下去再读开：这一格的记录还是同一枚，窗格数得出、判得掉", () => {
    const doc = open(SAID);
    clickCell(doc, { n: 90 }, "顾峰", { sides: { bottom: { val: "single" } } });
    const again = open(writeDocx(doc));
    expect(again.revisionEntries().map((e) => `${e.type}:${e.id}:${e.author}:${e.text}`)).toEqual(["cell-format:91:顾峰:甲"]);
    const cell = tableOf(again).rows[0][0]!;
    expect(cell.revs![0]!.prevSlots.map((s) => s.tag)).toEqual(["tcW", "tcMar", "vAlign"]);
    expect(cell.revs![0]!.prev).toMatchObject({ widthTwips: 2000, valign: "bottom", margin: { top: 60 } });
    expect(again.resolveRevisions("reject", 91)).toBe(1);
    expect(again.revisionEntries()).toEqual([]);
  });
});
