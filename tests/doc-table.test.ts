import { describe, expect, it } from "vitest";
import type { Table } from "../src/core/docx.js";
import { syncRowHeader, WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeCellMarginDocx, makeMergedTableDocx, makeRowSignatureDocx, makeTableFidelityDocx, makeWideRootDocx } from "./helpers/make-tables.js";

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
    // 说了 pct 宽（没有缇宽可算半径）→ 退回浏览器的 auto 居中，但那枚 108 缇缩进照样上身（#213）。
    expect(html).toContain("margin-left:auto;margin-right:auto;position:relative;left:7.2px;");
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
    // `vertical-align` 认的是 `middle`，不是文件里的 `center`：写成 `center` 是无效值，浏览器丢掉它、
    // 字照旧贴着上沿——授权 46 份里 1,816 枚 center 全部没画对（#191 的 DOM 闸门量出来的）。
    expect(html).toContain("vertical-align:middle;");
    expect(html).toContain("vertical-align:bottom;");
    expect(html).toContain("border-top:1px double currentColor;");
    expect(html).toContain("border-bottom:1px dotted #0070C0;");
  });

  it("gives a row its height and hides the one that is hidden", () => {
    const html = page();
    // 地板要写成浏览器认的那一句：`min-height` 落在 `<tr>`（或 `<td>`）上整条被丢掉——实测声明 200px 的行
    // 画 23px（`.scratch/row-height-css.html`），而表里的行高是铺排量的依据（`doc-layout.ts` 一行算一枚），
    // 认不到的下限就是「表比文件说的矮」。`height` 才是内容短时按它高、内容长时自己长的下限。
    // 第一行还带着文件说的标题行那一句（`w:trPr/w:tblHeader`），纸面按同一句盖 `data-tbl-header`。
    expect(html).toContain('<tr data-tbl-header="1" data-cant-split="1" style="height:40px;">');
    expect(html).toContain('<tr style="display:none;height:20px;">');
  });

  it("atLeast 与 exact 在纸上画同一下限，各自那句仍留在模型里", () => {
    const doc = open(makeTableFidelityDocx());
    const t = tableOf(doc);
    expect([t.rows[0][0]!.row!.heightRule, t.rows[1][0]!.row!.heightRule]).toEqual(["atLeast", "exact"]);
    const html = renderWord(doc);
    // 这一版没法把表行压到内容之下，exact 那句的「裁掉」画不出来；但下限必须画出来，且两句都不许退回
    // min-height（那条在这套引擎里整条无效，见 `.scratch/row-height-css.html`）。
    expect(html).toContain('<tr data-tbl-header="1" data-cant-split="1" style="height:40px;">');
    expect(html).toContain('<tr style="display:none;height:20px;">');
    expect(html).not.toContain("min-height:");
  });

  it("标题行重复：关掉时摘掉那一枚孩子，其余孩子原位不动；写出与读回同一句", () => {
    const doc = open(makeTableFidelityDocx());
    const row = tableOf(doc).rows[0][0]!.row!;
    expect(row.header).toBe(true);
    syncRowHeader(row, false);
    expect(row.header).toBeUndefined();
    // 只摘 `tblHeader`：`w:gridBefore`、`w:trHeight`、`w:cantSplit` 都还在，次序也不动（`CT_TrPr` 的孩子序）。
    expect(row.slots?.map((s) => s.tag)).toEqual(["gridBefore", "trHeight", "cantSplit"]);
    const out = saved(doc);
    expect(out).not.toContain("<w:tblHeader/>");
    expect(out).toContain("<w:cantSplit/>");
    // 存盘再读回来：这一行不再是标题行，其余字段逐个还在。
    const again = tableOf(open(writeDocx(doc))).rows[0][0]!.row!;
    expect(again.header).toBeUndefined();
    expect({ height: again.heightTwips, rule: again.heightRule, keep: again.keepTogether }).toEqual({ height: 600, rule: "atLeast", keep: true });
  });

  it("标题行重复：编辑器自己造的行没有孩子可换，字段说了算；重复落笔不多出一枚", () => {
    const doc = open(makeTableFidelityDocx());
    const cells = tableOf(doc).rows[2]!;
    const row = {};
    for (const cell of cells) cell.row = row;
    syncRowHeader(row, true);
    syncRowHeader(row, true);
    expect((row as { header?: boolean }).header).toBe(true);
    expect((row as { slots?: unknown[] }).slots).toBeUndefined();
    // 夹具第 0 行本来就带一枚（`w:tblHeader`）；这一行落两次笔只该再添一枚。
    expect(saved(doc).split("<w:tblHeader/>").length - 1).toBe(2);
    expect(tableOf(open(writeDocx(doc))).rows[2][0]!.row?.header).toBe(true);
    syncRowHeader(row, false);
    const back = tableOf(open(writeDocx(doc))).rows;
    expect(back[2]![0]!.row?.header).toBeUndefined();
    expect(back[0]![0]!.row?.header).toBe(true);
  });

  it("renders the nested table inside its cell, and marks an inserted row on the row", () => {
    expect(page().split('<table class="doc-table"').length - 1).toBe(2);
    const signature = renderWord(open(makeRowSignatureDocx()));
    expect(signature).toContain('<tr class="rev-row rev-ins" data-row-rev="ins" data-row-rev-id="51"');
    expect(signature).toContain('style="height:46.666666666666664px;"');
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
    // 表格自己没声明宽度，就不写 `w:tblW`：`w:w="0" w:type="auto"` 说的正是"没声明"，
    // 为一列宽度都没定的表补上这一句只是写出模型自己造的属性。宽度由 `w:tblGrid` 与单元格各自的说法带着。
    expect(out).not.toContain("<w:tblPr>");
    // 边距同宽度一样：编辑器没让它说过，就不替它写一句 `w:tblCellMar`。
    expect(out).not.toContain("tblCellMar");
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

/**
 * A cell's padding is stated in three places and the page has to obey the closest one. Before this round the
 * padding was not read at all: the sheet gave every cell `5px 9px` (`src/ui/index.html`), while all 115 of the
 * corpus's WPS-written table-bearing files state margins — 204 tables in `w:tblPr`, 1,660 rows in the
 * `w:tblPrEx` band, 1,897 cells in `w:tcPr` — and the commonest table says `0` above and below against 108
 * twips at the sides. Because the page cuts its tables off measured row boxes (`doc-layout.ts`), those 5px
 * were 10px of height on every one of the corpus's 2,647 rows: paper the file never asked for.
 */
describe("docx cell margins", () => {
  const tablesOf = (doc: WordDocument): Table[] => doc.blocks.filter((b) => b.kind === "table") as Table[];
  const marginDoc = (): WordDocument => open(makeCellMarginDocx());

  it("reads the padding the table, the row's exception band and the cell each state", () => {
    const [t] = tablesOf(marginDoc());
    expect(t.cellMargin).toEqual({ top: 0, left: 108, bottom: 0, right: 108 });
    // `w:tblPrEx` is a child of `w:tr`, not of `w:trPr`, so the band is read off the row it belongs to.
    expect(t.rows[0][0]!.row!.cellMargin).toEqual({ left: 0, right: 0 });
    expect(t.rows[0][1]!.row).toBe(t.rows[0][0]!.row);
    expect(t.rows[0][0]!.margin).toEqual({ top: 15 });
    expect(t.rows[1][0]!.margin).toEqual({ top: 12, left: 170, bottom: 12, right: 170 });
    // A row with no `w:trPr` and no band holds no row record at all, so the sides fall to the table.
    expect(t.rows[1][0]!.row).toBeUndefined();
    expect(tablesOf(marginDoc())[1]!.cellMargin).toBeUndefined();
  });

  it("tells a side that says zero from a side that says nothing, and leaves a non-`dxa` side asking the next layer", () => {
    const [t] = tablesOf(marginDoc());
    // `<w:tcMar></w:tcMar>` names no side: the same fact as having no `w:tcMar` at all.
    expect(t.rows[1][1]!.margin).toBeUndefined();
    // `w:type="auto"` is not a number counted in twips, so this cell states only its top and the table's
    // 108 stays the answer at the left.
    expect(t.rows[2][0]!.margin).toEqual({ top: 57 });
  });

  it("draws the cell's own padding over the row band's, and the band's over the table's", () => {
    const html = renderWord(marginDoc());
    // 15 twips above (cell) beats nothing above from the table's 0; the band's 0 at the sides beats 108.
    expect(html).toContain("padding-top:1px;padding-left:0px;padding-bottom:0px;padding-right:0px;");
    // The cell beside it says nothing: the band answers the sides, the table the top and bottom.
    expect(html).toContain("padding-top:0px;padding-left:0px;padding-bottom:0px;padding-right:0px;");
    // 12 and 170 twips: the cell states all four, so neither the band nor the table gets a say.
    expect(html).toContain("padding-top:0.8px;padding-left:11.333333333333334px;padding-bottom:0.8px;padding-right:11.333333333333334px;");
    // 108 twips is 7.2px, and the page's own 9px is gone from a table that never asked for it.
    expect(html).toContain("padding-top:0px;padding-left:7.2px;padding-bottom:0px;padding-right:7.2px;");
  });

  it("leaves a table that says nothing of padding to the page's own spacing", () => {
    const html = renderWord(marginDoc());
    const silent = html.slice(html.lastIndexOf('<table class="doc-table"'));
    const td = /<td[^>]*>/.exec(silent)![0];
    expect(td).not.toContain("padding-");
    expect(silent).not.toContain("padding-");
  });

  it("saves the padding the file stated, and reads it back as the same numbers", () => {
    const doc = marginDoc();
    const [t] = tablesOf(doc);
    const out = saved(doc);
    // The writer still spells these sets from the carried text: reading a padding is not authority to change one.
    expect(out).toContain(t.tblPrXml!);
    expect(out).toContain(t.rows[0][0]!.tcPrXml!);
    expect(out).toContain(t.rows[2][0]!.tcPrXml!);
    const again = tablesOf(open(writeDocx(doc)))[0]!;
    expect(again).toEqual(t);
  });
});
