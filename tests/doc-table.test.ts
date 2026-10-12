import { describe, expect, it } from "vitest";
import type { Table } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeMergedTableDocx, makeRowSignatureDocx, makeTableFidelityDocx, makeWideRootDocx } from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;

/** What the page does to a document that has everything in it, before and after one keystroke elsewhere. */
describe("docx table properties", () => {
  it("reads what the table itself says", () => {
    const t = tableOf(open(makeTableFidelityDocx()));
    expect(t.styleId).toBe("a7");
    expect(t.widthTwips).toBe(5000);
    expect(t.widthType).toBe("pct");
    expect(t.align).toBe("center");
    expect(t.indentTwips).toBe(108);
    expect(t.layout).toBe("fixed");
    expect(t.borders?.top).toEqual({ val: "single", size: 8, space: 0, color: "FF0000" });
    expect(t.borders?.insideH?.val).toBe("dashed");
    expect(t.borders?.insideV?.val).toBe("none");
    expect(t.columnWidths).toEqual([2000, 3000]);
  });

  it("reads what the row says, once for the row rather than once per cell", () => {
    const t = tableOf(open(makeTableFidelityDocx()));
    const row = t.rows[0][0]!.row!;
    expect(row).toBe(t.rows[0][1]!.row);
    expect({ height: row.heightTwips, rule: row.heightRule, header: row.header, keep: row.keepTogether }).toEqual({
      height: 600,
      rule: "atLeast",
      header: true,
      keep: true,
    });
    expect(t.rows[1][0]!.row).toMatchObject({ heightTwips: 300, heightRule: "exact", hidden: true });
    expect(t.rows[2][0]!.row).toBeUndefined();
  });

  it("reads what each cell says, including the one that merges down and the one that refuses to wrap", () => {
    const t = tableOf(open(makeTableFidelityDocx()));
    expect(t.rows[0][0]).toMatchObject({
      widthTwips: 2000,
      shading: "#FFF2CC",
      valign: "center",
      noWrap: true,
      textDirection: "tbRlV",
      borders: { bottom: { val: "dotted", size: 4, color: "0070C0" }, top: { val: "double", size: 6 } },
    });
    expect(t.rows[1][0]!.row).not.toBe(t.rows[0][0]!.row);
    expect(t.rows[1][0]).toMatchObject({ spanRows: 1 });
    expect(t.rows[2][0]).toMatchObject({ spanRows: 0 });
  });

  it("keeps a table nested in a cell as a table, in the slot it sat in", () => {
    const t = tableOf(open(makeTableFidelityDocx()));
    const cell = t.rows[0][1]!;
    expect(cell.paragraphs.map((b) => b.kind)).toEqual(["paragraph", "table"]);
    const nested = cell.paragraphs[1] as Table;
    expect(nested.rows[0][0]!.paragraphs.map((p) => (p.kind === "paragraph" ? p.inlines.map((i) => i.text).join("") : ""))).toEqual(["表中之表"]);
    expect(nested.borders?.top?.val).toBe("dotted");
  });

  it("carries the property sets it cannot explain, and writes them back without touching them", () => {
    const doc = open(makeTableFidelityDocx());
    const t = tableOf(doc);
    const out = saved(doc);
    // `w:tblLook`, `w:tblCellMar` and a pending `w:tblPrChange` are read but never rebuilt from fields.
    expect(out).toContain(t.tblPrXml!);
    expect(t.tblPrXml).toContain('<w:tblLook w:val="04A0"');
    expect(t.tblPrXml).toContain("<w:tblCellMar>");
    expect(t.tblPrXml).toContain("<w:tblPrChange ");
    expect(out).toContain(t.gridXml!);
    for (const row of t.rows) {
      for (const cell of row) expect(out).toContain(cell.tcPrXml!);
    }
    // A cell the model has no field for: `w:cnfStyle` and `w:tcMar` come back as they were written.
    expect(t.rows[2][1]!.tcPrXml).toContain('<w:cnfStyle w:val="000000000100"/>');
    expect(t.rows[0][0]!.tcPrXml).toContain("<w:tcMar>");
    // `w:gridBefore` is not a field either, and it is not the row's own record, so it is a slot.
    expect(t.rows[0][0]!.row!.slots?.map((s) => s.tag)).toEqual(["gridBefore", "trHeight", "tblHeader", "cantSplit"]);
    for (const slot of t.rows[0][0]!.row!.slots!) expect(out).toContain(slot.xml);
  });

  it("carries the band of table exceptions a row holds after its cells", () => {
    const doc = open(makeTableFidelityDocx());
    const row = tableOf(doc).rows[0][0]!.row!;
    expect(row.extras?.map((s) => s.tag)).toEqual(["tblPrEx"]);
    expect(row.extras?.[0]?.xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="E6EEF6"/>');
    const last = doc.blocks[doc.blocks.length - 1];
    if (last.kind === "paragraph") last.inlines[0].text = "改过的段落";
    const after = saved(doc);
    const band = row.extras![0]!.xml;
    expect(after.split(band).length - 1).toBe(1);
    // `CT_Row` puts the band after the row's cells, so it must not come back in front of them either.
    expect(after.indexOf(band)).toBeGreaterThan(after.indexOf("甲格"));
    expect(after.indexOf(band)).toBeLessThan(after.indexOf("<w:tr", after.indexOf(band)));
    expect(tableOf(doc).rows[1][0]!.row?.extras).toBeUndefined();
  });

  it("still has every one of them after a character is edited outside the table", () => {
    const doc = open(makeTableFidelityDocx());
    const before = tableOf(doc);
    const last = doc.blocks[doc.blocks.length - 1];
    if (last.kind === "paragraph") last.inlines[0].text = "改过的段落";
    const after = saved(doc);
    expect(after).toContain(before.tblPrXml!);
    expect(after).toContain(before.gridXml!);
    for (const row of before.rows) for (const cell of row) expect(after).toContain(cell.tcPrXml!);
    expect(after).toContain("表中之表");
    expect(after.split("<w:tbl>").length - 1).toBe(2);
    expect(after).toContain("改过的段落");
  });

  it("saves the same table the file wrote, tag for tag, and reads it back as the same model", () => {
    const doc = open(makeTableFidelityDocx());
    const again = open(writeDocx(doc));
    expect(tableOf(again)).toEqual(tableOf(doc));
    expect(again.blocks).toEqual(doc.blocks);
  });
});

/** The row's own tracked change is the one child of `w:trPr` a decision rewrites. */
describe("docx row signature and its neighbours", () => {
  it("accepting the row's insertion drops the record and leaves the rest of the row alone", () => {
    const doc = open(makeRowSignatureDocx());
    expect(tableOf(doc).rows[0][0]!.row).toMatchObject({ rev: { kind: "ins", id: 51 }, heightTwips: 700, keepTogether: true });
    doc.resolveRevisions("accept", 51);
    const out = saved(doc);
    expect(out).not.toContain("<w:ins ");
    expect(out).toContain("<w:cantSplit/>");
    expect(out).toContain('<w:trHeight w:val="700"/>');
    expect(out).toContain("<w:trPrChange ");
    expect(out).toContain("加进来的一行");
  });

  it("refusing it takes the row out with the record, and leaves the other row untouched", () => {
    const doc = open(makeRowSignatureDocx());
    doc.resolveRevisions("reject", 51);
    const out = saved(doc);
    expect(tableOf(doc).rows).toHaveLength(1);
    expect(out).not.toContain("加进来的一行");
    expect(out).toContain("普通行");
    expect(out).not.toContain("<w:ins ");
  });
});

/** Carried markup is only faithful if the part it lands in still resolves the prefixes it uses. */
describe("docx root declarations", () => {
  it("writes back the prefixes the root declared, and the list of ones a reader may skip", () => {
    const doc = open(makeWideRootDocx());
    const out = saved(doc);
    expect(doc.documentIgnorable()).toBe("w14 w15 wp14");
    expect(out).toContain('mc:Ignorable="w14 w15 wp14"');
    for (const prefix of ["w14", "w15", "wp14", "wpc", "wpsCustomData", "mc", "wne"]) {
      expect((out.match(new RegExp(`xmlns:${prefix}=`, "g")) ?? []).length).toBe(1);
    }
    expect((out.match(/xmlns:w=/g) ?? []).length).toBe(1);
    expect(out).toContain('<w:tblLook w:val="04A0"');
  });
});

/** A table that says these things to the file has to say them on the page too. */
describe("docx table rendering", () => {
  const page = (): string => renderWord(open(makeTableFidelityDocx()));

  it("draws the table's own width, alignment, layout and lines", () => {
    const html = page();
    expect(html).toContain("table-layout:fixed;");
    expect(html).toContain("width:100%;");
    expect(html).toContain("margin-left:auto;margin-right:auto;");
    expect(html).toContain("border-top:1px solid #FF0000;");
    expect(html).toContain("border-top:1px dashed currentColor;");
    expect(html).toContain("border-bottom:1px dashed currentColor;");
    // `w:insideV w:val="none"` draws nothing between the columns, so the only vertical lines on the page
    // are the two the table set for itself.
    expect((html.match(/border-left:/g) ?? []).length).toBe(1);
    expect((html.match(/border-right:/g) ?? []).length).toBe(1);
  });

  it("draws what each cell asked for: shading, vertical text, no wrapping, its own borders", () => {
    const html = page();
    expect(html).toContain("background:#FFF2CC;");
    expect(html).toContain("writing-mode:vertical-rl;");
    expect(html).toContain("white-space:nowrap;");
    expect(html).toContain("vertical-align:center;");
    expect(html).toContain("vertical-align:bottom;");
    expect(html).toContain("border-top:1px double currentColor;");
    expect(html).toContain("border-bottom:1px dotted #0070C0;");
  });

  it("gives a row its height and hides the one that is hidden", () => {
    const html = page();
    expect(html).toContain('<tr style="min-height:40px;">');
    expect(html).toContain('<tr style="display:none;height:20px;max-height:20px;">');
  });

  it("renders the nested table inside its cell, and marks an inserted row on the row", () => {
    expect(page().split('<table class="doc-table"').length - 1).toBe(2);
    const signature = renderWord(open(makeRowSignatureDocx()));
    expect(signature).toContain('<tr class="rev-row rev-ins" data-row-rev="ins" data-row-rev-id="51"');
    expect(signature).toContain('style="min-height:46.666666666666664px;"');
  });

  it("spans a vertical merge over the rows it covers, and keeps the covered cell on the page", () => {
    const html = page();
    expect(html).toContain('rowspan="2"');
    // Two things hide on this page: the row the file marked `w:hidden`, and the cell a merge above drew.
    expect((html.match(/display:none;/g) ?? []).length).toBe(2);
    // Neither the merge nor the span costs the page a cell, because the next save reads its cells back.
    expect(html.split("<td").length - 1).toBe(7);
  });

  it("reads a drawn merge as a restart and a continue, and writes both back", () => {
    const doc = open(makeMergedTableDocx());
    const t = tableOf(doc);
    expect(t.rows.map((r) => r.map((c) => c.spanRows))).toEqual([
      [1, undefined],
      [0, undefined],
      [1, undefined],
    ]);
    const out = saved(doc);
    expect(out.split('<w:vMerge w:val="restart"').length - 1).toBe(2);
    expect(out.split("<w:vMerge/>").length - 1).toBe(1);
    expect(tableOf(open(writeDocx(doc)))).toEqual(t);
    expect(renderWord(doc).split('rowspan="2"').length - 1).toBe(1);
  });
});

/** A table the editor made has no carried text at all, and still has to be a valid one. */
describe("docx tables built by the editor", () => {
  it("writes a grid, widths and the properties the model does know", () => {
    const doc = WordDocument.blank();
    doc.blocks.push({
      kind: "table",
      columns: 2,
      rows: [[{ paragraphs: [], widthTwips: 4513 }, { paragraphs: [], widthTwips: 4513 }]],
    });
    const out = saved(doc);
    expect(out).toContain("<w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>");
    expect(out).toContain('<w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>');
    expect(out).toContain('<w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr>');
    expect(tableOf(open(writeDocx(doc))).columnWidths).toEqual([4513, 4513]);
  });

  it("regrows the grid when a column is added to a table that came from a file", () => {
    const doc = open(makeTableFidelityDocx());
    const t = tableOf(doc);
    for (const row of t.rows) row.push({ paragraphs: [], widthTwips: 1000 });
    t.columns = 3;
    const out = saved(doc);
    expect(out).toContain('<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/><w:gridCol w:w="1000"/></w:tblGrid>');
    expect(out).toContain("<w:tblStyle w:val=\"a7\"/>");
  });
});
