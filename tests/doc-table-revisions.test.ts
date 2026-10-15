import { describe, expect, it } from "vitest";
import type { RevisionEntry, Table } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
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
