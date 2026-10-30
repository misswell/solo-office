import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

const pack = (document: string, extraRootAttributes = ""): Uint8Array => {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      XML_DECL +
        `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/></Types>`,
    ),
    "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"${extraRootAttributes}><w:body>${document}</w:body></w:document>`),
  };
  return zipSync(files, { level: 6 });
};

const REV_ATTRS = (id: number, author: string): string => `w:id="${id}" w:author="${author}" w:date="2026-03-01T08:00:00Z"`;

/**
 * A table that says everything about itself, in the shapes ECMA-376 defines for `w:tblPr`, `w:trPr` and
 * `w:tcPr`: percent width, alignment, indent, a full border set including the inside lines, shading, fixed
 * layout, cell margins and a look record; a header row with a height that may not be broken; a cell with
 * its own borders, margins, vertical text, wrapping refusal and a merge that spans rows; and a second table
 * inside one of its cells.
 *
 * It exists because a table is the one block this editor rebuilds from the page: before the model carried
 * these property sets, editing a character outside the table rewrote it as a bare grid and deleted the
 * nested table with no complaint. The `.scratch/tbl-fidelity.ts` probe prints the same shape as a table.
 *
 * The header row also carries a `w:tblPrEx` band after its cells — the row drawing with the table's borders
 * and shading overridden — because `CT_Row` puts that group after the cells and lets it repeat. The shape is
 * ECMA-derived: no file in the WPS corpus here holds a table at all, though WPS's own Word ML module spells
 * `w:tblPrEx` among its element names (`.scratch` scan of `wordml12w-arm64`, UTF-16 literals).
 */
export function makeTableFidelityDocx(): Uint8Array {
  const document = `<w:p><w:r><w:t>表格前的段落</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblStyle w:val="a7"/><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblInd w:w="108" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="FF0000"/><w:left w:val="single" w:sz="8" w:space="0" w:color="FF0000"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="FF0000"/><w:right w:val="single" w:sz="8" w:space="0" w:color="FF0000"/><w:insideH w:val="dashed" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/><w:tblPrChange ${REV_ATTRS(41, "林晓")}><w:tblW w:w="0" w:type="auto"/></w:tblPrChange></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr><w:trPr><w:gridBefore w:val="1"/><w:trHeight w:val="600" w:hRule="atLeast"/><w:tblHeader/><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcBorders><w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="0070C0"/><w:top w:val="double" w:sz="6" w:space="0" w:color="auto"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:noWrap/><w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar><w:textDirection w:val="tbRlV"/><w:vAlign w:val="center"/><w:tcPrChange ${REV_ATTRS(42, "顾峰")}><w:tcW w:w="1500" w:type="dxa"/></w:tcPrChange></w:tcPr><w:p><w:r><w:t>甲格</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:r><w:t>乙格</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="a3"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="dotted" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>表中之表</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc><w:tblPrEx><w:tblBorders><w:top w:val="dotted" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:shd w:val="clear" w:color="auto" w:fill="E6EEF6"/></w:tblPrEx></w:tr>
<w:tr><w:trPr><w:trHeight w:val="300" w:hRule="exact"/><w:hidden/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>竖合并起</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>第二行</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:cnfStyle w:val="000000000100"/></w:tcPr><w:p><w:r><w:t>竖合并续</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>表格后的段落</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * A row that was inserted with a tracked change, and that also says something else about itself: it may not
 * split across pages and it has a height. Deciding the row's signature must rewrite only that signature —
 * the other children of `w:trPr` belong to the row whether the change is accepted or refused.
 */
export function makeRowSignatureDocx(): Uint8Array {
  const document = `<w:p><w:r><w:t>正文</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="700"/><w:ins ${REV_ATTRS(51, "顾峰")}/><w:trPrChange ${REV_ATTRS(52, "林晓")}><w:cantSplit/></w:trPrChange></w:trPr><w:tc><w:p><w:r><w:t>加进来的一行</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>普通行</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * A table whose three parts each carry somebody else's undecided property change: the table's own
 * `w:tblPrChange` (width was switched from auto to percent, and the left cell margin from 72 to 108), the
 * first row's `w:trPrChange` (height 300 became 600 exact, the row may no longer be kept together, and the
 * header repeat came with the change), and the first cell's `w:tcPrChange` (1500 became 2000 and the shading
 * colour changed). Each record holds the *complete* set from before its revision, which is what ISO/IEC
 * 29500 says a `w:*PrChange` carries and what makes a refusal more than "put this one property back" — so the
 * sets here deliberately include properties the model has no field for, and a refusal has to bring them back
 * as text: `w:tblCellMar` in the table's, `w:tcMar` and `w:shd` in the cell's.
 *
 * Two spellings of the record are in here on purpose: the table and row records wrap the set in an element of
 * its own name (`<w:tblPrChange…><w:tblPr…/></w:tblPrChange>`, MS's own example for `w:pPrChange`), the cell
 * record writes its properties directly. Both are read, because guessing one and meeting the other would
 * quietly drop a reviewer's properties. 未取证: neither spelling was observed in a WPS-written file — the
 * corpus here has no docx with a revised table, so the shapes are ECMA-376 §17.13.5 plus MS's class pages.
 *
 * The header row also carries a `w:tblPrEx` band with a `w:tblPrExChange` of its own, which this model reads
 * as part of the band's text and writes back untouched: it survives every decision, but it is not listed and
 * cannot be decided — the one record of the family the 审阅窗格 does not cover yet.
 */
export function makeTablePropertyRevisionDocx(): Uint8Array {
  const document = `<w:p><w:r><w:t>表前</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar><w:tblPrChange ${REV_ATTRS(61, "林晓")}><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellMar><w:left w:w="72" w:type="dxa"/></w:tblCellMar></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="exact"/><w:tblHeader/><w:trPrChange ${REV_ATTRS(62, "顾峰")}><w:trPr><w:cantSplit/><w:trHeight w:val="300"/></w:trPr></w:trPrChange></w:trPr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar><w:tcPrChange ${REV_ATTRS(63, "林晓")}><w:tcW w:w="1500" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="DDEBF7"/><w:tcMar><w:top w:w="120" w:type="dxa"/></w:tcMar></w:tcPrChange></w:tcPr><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:cnfStyle w:val="000000000100"/></w:tcPr><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc><w:tblPrEx><w:shd w:val="clear" w:color="auto" w:fill="E6EEF6"/><w:tblPrExChange ${REV_ATTRS(64, "顾峰")}><w:shd w:val="clear" w:color="auto" w:fill="FDE9D9"/></w:tblPrExChange></w:tblPrEx></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>丙</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>丁</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>表后</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * A table whose padding is stated in all three places a file may state it, spelled the way the corpus spells
 * it: the table's `w:tblCellMar` (`w:w="0"` above and below against `108` at the sides — the shape 204 of the
 * 204 tables in the 115 WPS-written table-bearing files use), a row band `w:tblPrEx/w:tblCellMar` that zeroes
 * the two sides (431 of the corpus's 1,660 banded rows do exactly this against a table that said 108), and
 * cells that say their own `w:tcMar` — one side only, all four, and one side spelled `w:type="auto"`, which
 * asks for the next layer's answer instead of giving one. `w:tblCellSpacing` is left out: no file in the
 * corpus states it, and 0 of the 115 say anything about it.
 *
 * The second table says nothing anywhere, and the fourth cell of the first table carries an empty
 * `<w:tcMar></w:tcMar>`: both are the "silence" half of the distinction this fixture exists to test, since a
 * page that cannot tell `w:w="0"` from "not said" puts spacing back on a table that asked for none.
 *
 * Values are drawn from what the corpus actually uses (108, 12, 15, 57, 170 twips), not invented; every side
 * the corpus states comes with `w:type="dxa"`, so the `auto` side here is a shape no file in it has. 未取证:
 * what WPS draws for these — the page-level comparison is against the file's own numbers.
 */
export function makeCellMarginDocx(): Uint8Array {
  const cell = (text: string, mar = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${mar}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const document = `<w:p><w:r><w:t>表前</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>${cell("甲", '<w:tcMar><w:top w:w="15" w:type="dxa"/></w:tcMar>')}${cell("乙")}<w:tblPrEx><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPrEx></w:tr>
<w:tr>${cell("丙", '<w:tcMar><w:top w:w="12" w:type="dxa"/><w:left w:w="170" w:type="dxa"/><w:bottom w:w="12" w:type="dxa"/><w:right w:w="170" w:type="dxa"/></w:tcMar>')}<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcMar></w:tcMar></w:tcPr><w:p><w:r><w:t>丁</w:t></w:r></w:p></w:tc></w:tr>
<w:tr>${cell("戊", '<w:tcMar><w:top w:w="57" w:type="dxa"/><w:left w:w="144" w:type="auto"/></w:tcMar>')}${cell("己")}</w:tr>
</w:tbl>
<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>${cell("没说边距的表")}</w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * A table a decision cannot take away: one column, whose only cell carries `w:cellDel`. Accepting the deletion
 * would leave a table with no cells, which is not a table a reader gets back, so the record stays pending
 * rather than being half-applied.
 */
export function makeSoleColumnDocx(): Uint8Array {
  const document = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:cellDel ${REV_ATTRS(80, "顾峰")}/></w:tcPr><w:p><w:r><w:t>唯一的列</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * A grid with one column that arrived and one that is on its way out, plus the two records that ride on a
 * single cell: three columns, where column 2 (starting at 1000) has `w:cellDel` on both its cells under one
 * `w:id`, column 3 (starting at 3000) has `w:cellIns` on both its cells under another, and the first cell of
 * the first row — the anchor of a vertical merge it was made to start — carries both a `w:cellMerge`, whose
 * recorded `w:vMerge` says it used to continue the merge above it, and a `w:tcPrChange` of its own.
 *
 * The shared `w:id` down a column is what makes a revised column one row of the 审阅窗格 rather than one per
 * cell, the way `w:ins` across several paragraphs is one row. `CT_TcPrBase` puts the records after the
 * cell's own properties and in the order `cellIns`, `cellDel`, `cellMerge`, `tcPrChange`, and `w:cellMerge`
 * is the one record of the family that names a single property (`w:vMerge`) instead of the whole set —
 * refusing it must not touch the cell's width, while refusing the `w:tcPrChange` beside it restores the
 * recorded width *and* merge and is expected to. 未取证: same caveat as above — no WPS file with a revised
 * table exists here.
 */
export function makeCellRevisionDocx(): Uint8Array {
  const cell = (width: number, text: string, extra = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const merge = `<w:cellMerge ${REV_ATTRS(71, "林晓")}><w:vMerge/></w:cellMerge>`;
  const widened = `<w:tcPrChange ${REV_ATTRS(72, "顾峰")}><w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr></w:tcPrChange>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr>${cell(1000, "甲一", `<w:vMerge w:val="restart"/>${merge}${widened}`)}${cell(2000, "乙一", `<w:cellDel ${REV_ATTRS(73, "林晓")}/>`)}${cell(3000, "丙一", `<w:cellIns ${REV_ATTRS(70, "顾峰")}/>`)}</w:tr>
<w:tr>${cell(1000, "甲二", `<w:vMerge/>`)}${cell(2000, "乙二", `<w:cellDel ${REV_ATTRS(73, "林晓")}/>`)}${cell(3000, "丙二", `<w:cellIns ${REV_ATTRS(70, "顾峰")}/>`)}</w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * A table whose record does not reach down the whole column: two columns of 2000 twips, two rows, and only
 * the second row's first cell carries `w:cellIns` — the cell above it and the one to its right say nothing.
 * A decision must not read this as a revised column, because dropping column 1 would take 甲一 with it, and
 * nobody revised 甲一. Refusing the insertion therefore takes that one cell out of its row and hands the
 * column it freed to the row's last cell, which is the only way the row goes on tiling `w:tblGrid` — a row
 * whose cells stop short of the grid leaves a hole the page has to invent an answer for. 未取证: no
 * WPS-written file in the corpus carries a table revision at all, so how WPS itself resolves a lone
 * `w:cellIns` is unobserved; the shape here is ECMA-376 §17.13.5 read against the grid-tiling rule.
 */
export function makePartialCellInsDocx(): Uint8Array {
  const cell = (width: number, text: string, extra = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const document = `<w:p><w:r><w:t>表前</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>${cell(2000, "甲一")}${cell(2000, "乙一")}</w:tr>
<w:tr>${cell(2000, "甲二", `<w:cellIns ${REV_ATTRS(90, "顾峰")}/>`)}${cell(2000, "乙二")}</w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * A table big enough to edit inside a cell: an outer two-by-two of 2000-twip columns, whose top-right cell
 * holds a paragraph and its own two-by-two table of 700-twip columns. `makeTableFidelityDocx` nests a table
 * too, but a single cell of it — enough to ask "does the nested table survive a save", not enough to ask
 * "which table does 插入行 hit when the caret is inside it". Both tables are fixed-layout with their own
 * `w:tblGrid`, so a command that lands on the wrong one shows up as a row or column the other table never
 * asked for.
 */
export function makeNestedTableDocx(): Uint8Array {
  const cell = (width: number, ...blocks: string[]): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${blocks.join("")}</w:tc>`;
  const para = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const inner = `<w:tbl><w:tblPr><w:tblW w:w="1400" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="700"/><w:gridCol w:w="700"/></w:tblGrid>` +
    `<w:tr>${cell(700, para("内一"))}${cell(700, para("内二"))}</w:tr>` +
    `<w:tr>${cell(700, para("内三"))}${cell(700, para("内四"))}</w:tr></w:tbl>`;
  const document = `<w:p><w:r><w:t>表前</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>${cell(2000, para("外一"))}${cell(2000, para("外二"), inner)}</w:tr>
<w:tr>${cell(2000, para("外三"))}${cell(2000, para("外四"))}</w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * Three tables deep, in the shape ECMA-376 uses: a nested table is a **child of `w:tc`**, sitting beside the
 * cell's paragraphs rather than inside one of them. The outer two-by-two's top-right cell holds a middle
 * table, whose own top-left cell holds a third. Something that walks tables one level at a time loses the
 * deepest one, so every level carries its own text.
 */
export function makeDeepNestedTableDocx(): Uint8Array {
  const cell = (width: number, ...blocks: string[]): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${blocks.join("")}</w:tc>`;
  const para = (t: string): string => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const table = (w: number, rows: string[][]): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${w * 2}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${w}"/><w:gridCol w:w="${w}"/></w:tblGrid>` +
    rows.map((r) => `<w:tr>${r.map((c) => cell(w, c)).join("")}</w:tr>`).join("") +
    "</w:tbl>";
  const deepest = table(300, [[para("深一"), para("深二")], [para("深三"), para("深四")]]);
  const middle = table(600, [[para("中一") + deepest, para("中二")], [para("中三"), para("中四")]]);
  const outer = table(2000, [[para("外一"), para("外二") + middle], [para("外三"), para("外四")]]);
  return pack(`<w:p><w:r><w:t>表前</w:t></w:r></w:p>${outer}<w:p/>`);
}

/**
 * A nested table whose record was written by somebody else and is still undecided: the *inner* table's second
 * row carries `w:trPr/w:ins`, and the cell beside it an inserted run. Built as a file rather than through the
 * editor, so it asks the 审阅 machinery a question the editor cannot — does it reach a record two tables deep,
 * and can it decide it without touching the table that holds it?
 */
export function makeNestedRevisedTableDocx(): Uint8Array {
  const cell = (width: number, ...blocks: string[]): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${blocks.join("")}</w:tc>`;
  const para = (t: string): string => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const inserted = (id: number, t: string): string =>
    `<w:p><w:ins ${REV_ATTRS(id, "顾峰")}><w:r><w:t>${t}</w:t></w:r></w:ins></w:p>`;
  const inner =
    `<w:tbl><w:tblPr><w:tblW w:w="1400" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="700"/><w:gridCol w:w="700"/></w:tblGrid>` +
    `<w:tr>${cell(700, para("内一"))}${cell(700, para("内二"))}</w:tr>` +
    `<w:tr><w:trPr><w:ins ${REV_ATTRS(88, "顾峰")}/></w:trPr>${cell(700, inserted(89, "内三"))}${cell(700, para("内四"))}</w:tr>` +
    "</w:tbl>";
  const document =
    `<w:p><w:r><w:t>表前</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
    `<w:tr>${cell(2000, para("外一"))}${cell(2000, para("外二"), inner)}</w:tr>` +
    `<w:tr>${cell(2000, para("外三"))}${cell(2000, para("外四"))}</w:tr>` +
    "</w:tbl><w:p/>";
  return pack(document);
}

/**
 * A vertical merge that is drawn, not hidden: three rows of two columns where the left cell of the first two
 * rows is one merged cell (`w:vMerge w:val="restart"` then `w:vMerge`), and the third row starts its own.
 * `makeTableFidelityDocx` does merge too, but its merge begins in a row the file marked `w:hidden`, which
 * says nothing about how a page has to line the columns up afterwards.
 */
export function makeMergedTableDocx(): Uint8Array {
  const document = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>占两行</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>第一行右</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>第二行右</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>新的一格</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:gridSpan w:val="1"/></w:tcPr><w:p><w:r><w:t>第三行右</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * A grid the structure commands can be checked against: three columns of different widths (`1000`, `2000`,
 * `3000`), the first cell of the last one spanning two of them, and one cell that carries a border set, a
 * shading and a `w:tcMar` of its own. The merge gives a command a row that does *not* have one cell per column,
 * and the property sets give it something it must not disturb — an inserted column has to leave a cell's dotted
 * bottom edge exactly as the file wrote it.
 */
export function makeOpsGridDocx(): Uint8Array {
  const cell = (width: number, text: string, extra = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const span = (width: number, columns: number, text: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:gridSpan w:val="${columns}"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr>${cell(1000, "甲一", '<w:tcBorders><w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="0070C0"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:tcMar><w:top w:w="60" w:type="dxa"/></w:tcMar>')}${cell(2000, "乙一")}${cell(3000, "丙一")}</w:tr>
<w:tr>${cell(1000, "甲二")}${cell(2000, "乙二")}${cell(3000, "丙二")}</w:tr>
<w:tr>${span(3000, 2, "跨两列")}${cell(3000, "丙三")}</w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * The root a real word processor writes: seventeen prefixes and a `mc:Ignorable` list naming the ones a
 * reader may skip if it does not know them. The declarations below are copied out of the three
 * WPS-produced docx in `.scratch/wps-corpus/` (`mc:Ignorable="w14 w15 wp14"`), minus `w:` and `r:`, which
 * `pack` adds; only the table inside is ours, because that corpus holds no docx with a table in it.
 */
const WPS_ROOT =
  ` xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"` +
  ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
  ` xmlns:o="urn:schemas-microsoft-com:office:office"` +
  ` xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"` +
  ` xmlns:v="urn:schemas-microsoft-com:vml"` +
  ` xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"` +
  ` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"` +
  ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
  ` xmlns:w10="urn:schemas-microsoft-com:office:word"` +
  ` xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"` +
  ` xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"` +
  ` xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"` +
  ` xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"` +
  ` xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"` +
  ` xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData"` +
  ` mc:Ignorable="w14 w15 wp14"`;

export function makeWideRootDocx(): Uint8Array {
  const document = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="4513"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>唯一的格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
  return pack(document, WPS_ROOT);
}

/**
 * The attributes an element keeps for itself, on its own opening tag rather than in a child: `w14:paraId` on
 * `w:p`, a `w:rsidR`/`w:rsidTr` pair on `w:tr`, a `w:rsidR` on `w:tc`. None of them is a property of the text
 * and none has a field in this model, so before the model carried them a save silently deleted the identity
 * every paragraph was written under — measured at 48 of 48 `w14:paraId` values lost across the WPS word files
 * here (`.scratch/attr-loss.ts`).
 *
 * Which half of this is a transcription and which is a guess is worth stating exactly: the `w14:paraId` shape
 * (eight hex digits, alone on the tag) and the root's declarations are copied out of WPS-written files, and
 * the `w:rsidR`/`w:rsidTr`/`w:rsidP` names are MS-OI29500's, not observed here — the six WPS word files carry
 * no rsid at all, and no rsid name appears in ONLYOFFICE's word sdk either, so a Word-authored file is where
 * these show up. The carry is by source text, not by name, which is why the test can rely on neither.
 */
export function makeElementAttrsDocx(): Uint8Array {
  const document = `<w:p w14:paraId="1B2BC3C3" w:rsidR="00A12345" w:rsidRPr="00A12345" w:rsidP="003E5678"><w:r><w:t>带身份标记的段落</w:t></w:r></w:p>
<w:p w14:paraId="5859A351"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>居中且带标记</w:t></w:r></w:p>
<w:p w14:paraId="7777AA11" w:rsidR="00A12345"/>
<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
<w:tr w:rsidR="00B24C7A" w:rsidTr="00C11111"><w:tc w:rsidR="00B24C7A"><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p w14:paraId="0A0B0C0D"><w:r><w:t>格内带标记的段落</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>没有任何标记的段落</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;
  return pack(document, WPS_ROOT);
}

/**
 * Rows that say where their text sits (`w:trPr/w:jc`), which is the one alignment word the authorized
 * corpus puts inside a table: 190 枚 on 4 templates, all `center`, with 1,294 cells below them that state
 * no alignment of their own (`.scratch/cell-align-audience.py`; `w:tcPr/w:jc` is 0 枚 cohort-wide).
 *
 * Four rows: centered with one cell whose paragraph states `left` of its own (the schema's precedence,
 * paragraph over row), a row that says nothing, a right-aligned row, and a centered row whose
 * `w:trPrChange` recorded that it was `left` before — the recorded set has to come back as `left` when
 * the change is refused.
 */
export function makeRowAlignDocx(): Uint8Array {
  const cell = (text: string, ppr = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p>${ppr}<w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
<w:tr><w:trPr><w:jc w:val="center"/></w:trPr>${cell("行居中的一格")}${cell("段自己说左", '<w:pPr><w:jc w:val="left"/></w:pPr>')}</w:tr>
<w:tr>${cell("没说话的一行")}${cell("也没说")}</w:tr>
<w:tr><w:trPr><w:jc w:val="right"/></w:trPr>${cell("靠右的一行")}${cell("同右")}</w:tr>
<w:tr><w:trPr><w:jc w:val="center"/><w:trPrChange ${REV_ATTRS(101, "顾峰")}><w:trPr><w:jc w:val="left"/></w:trPr></w:trPrChange></w:trPr>${cell("改过对齐的一行")}${cell("改过")}</w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * 一格里的字往哪儿靠（`w:tcPr/w:vAlign`）：第一行高 1440 缇（1 英寸），字少，所以三档在纸上真的差出位置。
 * 三格分别「没说话」（顶部）、`bottom`、`center`；甲格还带着自己的框线与底纹，改竖向对齐时它们必须原样不动、
 * 新的 `w:vAlign` 按 `CT_TcPr` 的孩子序排在它们之后。第二行是三格没说过的样子，表外还有一段用于「光标不在表里」。
 */
export function makeCellVAlignDocx(): Uint8Array {
  const cell = (text: string, tcPr: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="1440" w:hRule="atLeast"/></w:trPr>${cell("甲格", '<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="0070C0"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>')}${cell("乙格", '<w:vAlign w:val="bottom"/>')}${cell("丙格", '<w:vAlign w:val="center"/>')}</w:tr>
<w:tr>${cell("丁格", "")}${cell("戊格", '<w:tcBorders><w:top w:val="dotted" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders>')}${cell("己格", "")}</w:tr>
</w:tbl><w:p><w:r><w:t>表外的一段</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * 「对齐」那一组问的两层：竖向那句归**格**（`w:tcPr/w:vAlign`），横向那句归格里的**每一段**（`w:pPr/w:jc`）。
 * 第一行高 1440 缇、字少，三档在纸上真差得出位置。六格摆的是这一族全部五种说法形状：
 *  · 甲一 两样都没说（顶部 + 左，§17.4.84 与 §17.3.1.13 的默认值）；
 *  · 乙一 说了 `center` + 段说了 `center`；
 *  · 丙一 说了 `bottom` + 段说了 `right`，且那句 `w:jc` 前面还带着 `w:spacing`（落笔时排在哪由 `CT_PPr` 的孩子序定）；
 *  · 丁二 **一格两段**：一段说 `center`、一段没说——一格里两种说法，横向那一行说不出唯一一句；
 *  · 戊二 段明确说了 `left`（「说了 left」与「没说」是两句话，但画出来是同一个位置）；
 *  · 己二 段带着 `w:spacing` 而没说 `w:jc`（命令落笔要按孩子序插那一枚，`w:spacing` 原样带回）。
 * 表外留一段，给「光标不在表格里」。
 */
export function makeCellAlignDocx(): Uint8Array {
  const cell = (text: string, tcPr: string, paras: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr>${paras}</w:tc>`;
  const para = (text: string, pPr = ""): string => `<w:p>${pPr}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const jc = (v: string): string => `<w:pPr><w:jc w:val="${v}"/></w:pPr>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="1440" w:hRule="atLeast"/></w:trPr>${cell("甲一", "", para("甲一"))}${cell("乙一", '<w:vAlign w:val="center"/>', para("乙一", jc("center")))}${cell("丙一", '<w:vAlign w:val="bottom"/>', para("丙一", `<w:pPr><w:spacing w:after="240"/><w:jc w:val="right"/></w:pPr>`))}</w:tr>
<w:tr>${cell("丁二", "", `${para("丁二", jc("center"))}${para("丁二第二段")}`)}${cell("戊二", "", para("戊二", jc("left")))}${cell("己二", "", para("己二", '<w:pPr><w:spacing w:before="120"/></w:pPr>'))}</w:tr>
</w:tbl><w:p><w:r><w:t>表外的一段</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * 装机的壳在**盘上**留下的那一形状：甲一说了 `vAlign center`，而那枚 `w:tcPrChange` 的改前没有这一句
 * （#149 那一腿开着修订按边框时留下的记录，`w:tcBorders`/`w:shd` 是同格邻居）。第二趟开的就是这一件，
 * 于是「回显说的 ↔ 盘上写的那句」第一次要跨一趟对账。
 *
 * 这一形状此前只在装机往返里出现过（`makeCellAlignDocx` 的格不带记录），单测与 DOM 探针都摸不到它，
 * 而装机的壳量出「甲一那格读回来的竖向是顶部、`w:jc` 那句倒还在」。这里把它钉下来：改前说了什么不该被
 * 当成格现在说的话，读回来必须是 `center`。
 */
export function makeCellAlignRevisionDocx(): Uint8Array {
  const border = (side: string): string => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
  const borders = (sides: string[]): string => `<w:tcBorders>${sides.map(border).join("")}</w:tcBorders>`;
  const shd = '<w:shd w:val="clear" w:color="auto" w:fill="C00000"/>';
  const tcPrChange = (id: number): string =>
    `<w:tcPrChange ${REV_ATTRS(id, "林晓")}><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${borders(["top", "left"])}${shd}</w:tcPr></w:tcPrChange>`;
  const cell = (tcPr: string, paras: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr>${paras}</w:tc>`;
  const para = (text: string, pPr = ""): string => `<w:p>${pPr}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const jc = (v: string): string => `<w:pPr><w:jc w:val="${v}"/></w:pPr>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="1440" w:hRule="atLeast"/></w:trPr>${cell(`${borders(["top", "left", "bottom"])}${shd}<w:vAlign w:val="center"/>${tcPrChange(2)}`, para("甲一", jc("center")))}${cell("", para("乙一"))}</w:tr>
<w:tr>${cell('<w:vAlign w:val="bottom"/>', para("丙二", jc("right")))}${cell("", para("丁二"))}</w:tr>
</w:tbl><w:p><w:r><w:t>表外的一段</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * 合并记录的两套范围：三行两列。第一列在「现在」是 1 行 + 2 行两段合并（第二行重启），而第二行那格挂着一枚
 * `w:cellMerge`，记录的改前是「继续上面」——原文态整列是一整块 3 行。原文与最终态各画哪一套范围、差几行，
 * 都从这里量。第二列每行一格，用来认「行还在、只有第一列的范围变了」。
 */
export function makeCellMergeViewDocx(): Uint8Array {
  const plain = (text: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const merged = (text: string, v: string, extra = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vMerge${v ? ` w:val="${v}"` : ""}/>${extra}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>${merged("甲零", "restart")}${plain("乙零")}</w:tr>
<w:tr>${merged("甲一", "restart", `<w:cellMerge ${REV_ATTRS(74, "林晓")}><w:vMerge/></w:cellMerge>`)}${plain("乙一")}</w:tr>
<w:tr>${merged("甲二", "")}${plain("乙二")}</w:tr>
</w:tbl>`;
  return pack(document);
}

/**
 * 「单元格尺寸」问的两层：高度归**行**（`w:trPr/w:trHeight`，§17.3.2.77），宽度归**列**（`w:gridCol/w:w` 加上
 * 那一列每一格那句 `w:tcPr/w:tcW`）。两张表各摆自家量出来的那几种形状（授权 46 份 58 张：同列全等 26 张、
 * 同列两个数 **0** 张、有跨列格或说不出数 32 张；`w:trHeight` 非 0 的 580 枚全说 `atLeast`，`exact` 0 枚）。
 *
 * 第一张：4 行 × 4 列网格（2000/1500/1500/1500，`w:tblW` 说的是网格和 6,500 缇）。
 *  · 行 0 说 `1440 atLeast`；行 1 说 `600 atLeast` 并带着 `w:cantSplit`（改高度时那一枚必须原样不动）；
 *  · 行 2 什么也没说（回显要空着，落笔才现造那一枚 `w:trPr`）；行 3 说 `800 exact`（落笔只换数，不换成 `atLeast`）；
 *  · 第 0 列每格都说 2,000 缇、只占一列（改得动）；第 1~3 列只被一枚 `w:gridSpan="3"` 的格盖着
 *    （那一列说不出单一宽度，改它必须回绝，而不是把跨列那句改成「这一列加剩下三列」）。
 *
 * 第二张：3 列网格（2000/1000/1000，`w:tblW` 4,000 缇）说的是另外三种落点。第 0 列照常改得动；第 1 列既有
 * 只占一列的格、又有一枚跨两列的格（改这一列只动前者，后者一个字节不动）；第 2 列的格按 `pct` 摆（缇数不是
 * 它的口径，整列回绝）。表外留一段，给「光标不在表格里」。
 */
export function makeCellSizeDocx(): Uint8Array {
  const cell = (text: string, w: number, type = "dxa", span = 1): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="${type}"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ""}</w:tcPr>` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const row = (trPr: string, cells: string): string => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ""}${cells}</w:tr>`;
  const height = (val: number, rule?: string): string => `<w:trHeight w:val="${val}"${rule ? ` w:hRule="${rule}"` : ""}/>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="6500" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>` +
    row(height(1440, "atLeast"), cell("甲一", 2000) + cell("甲跨三列", 4500, "dxa", 3)) +
    row(`<w:cantSplit/>${height(600, "atLeast")}`, cell("乙一", 2000) + cell("乙跨三列", 4500, "dxa", 3)) +
    row("", cell("丙一", 2000) + cell("丙跨三列", 4500, "dxa", 3)) +
    row(height(800, "exact"), cell("丁一", 2000) + cell("丁跨三列", 4500, "dxa", 3)) +
    `</w:tbl>` +
    `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>` +
    row("", cell("戊一", 2000) + cell("戊二", 1000) + cell("戊三", 20, "pct")) +
    row("", cell("己一", 2000) + cell("己跨两列", 2000, "dxa", 2)) +
    `</w:tbl><w:p><w:r><w:t>表外的一段</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * 一列被改过宽度、正等着点头的形状：网格与这一列的格都说 1,500 缇，每格挂一枚 `w:tcPrChange`，记的改前是
 * 2,000——两枚共用一个 `w:id`（`docx.ts:coversColumn` 认的就是这一条：同一枚 `w:id` 摆在整列每一格上，
 * 说的才是「改过的一列」）。OOXML 没有 `w:tblGrid` 的记录族，所以「拒绝」那一趟得把这一列的网格数从格摆回来。
 */
export function makeColumnWidthRevisionDocx(): Uint8Array {
  const cell = (text: string, w: number, rec = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${rec}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const record = `<w:tcPrChange ${REV_ATTRS(90, "顾峰")}><w:tcW w:w="2000" w:type="dxa"/></w:tcPrChange>`;
  const document = `<w:tbl><w:tblPr><w:tblW w:w="3500" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="2000"/></w:tblGrid>` +
    `<w:tr>${cell("甲", 1500, record)}${cell("乙", 2000)}</w:tr>` +
    `<w:tr>${cell("丙", 1500, record)}${cell("丁", 2000)}</w:tr>` +
    `</w:tbl>`;
  return pack(document);
}

/**
 * 一张**每一边各有各说法**的表（#196，格子边的墨从哪来）：
 * `w:tblBorders` 只说 `w:top`（单线）与 `w:insideV`（不要线）；甲格自己说左边不要线，丙格自己说底边要一条单线。
 * 于是 16 条格边凑齐四种说法：格说的要画、格说的不要、表替这条缝/这一圈说的、两人都没说。
 * 「查看网格线」的导引线只该落在**最后三种里的"没人画线"那两种**上，所以这份夹具的每一格都是判据。
 */
export function makeGridlinesDocx(): Uint8Array {
  const border = (side: string, val: string, sz = 8, color = "FF0000"): string =>
    `<w:${side} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
  const cell = (label: string, w: number, tcBorders = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${tcBorders}</w:tcPr><w:p><w:r><w:t>${label}</w:t></w:r></w:p></w:tc>`;
  const document =
    `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblBorders>` +
    `${border("top", "single")}${border("insideV", "none", 0, "auto")}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>` +
    `<w:tr>${cell("甲", 1500, `<w:tcBorders>${border("left", "none", 0, "auto")}</w:tcBorders>`)}${cell("乙", 1500)}</w:tr>` +
    `<w:tr>${cell("丙", 1500, `<w:tcBorders>${border("bottom", "single", 6, "0070C0")}</w:tcBorders>`)}${cell("丁", 1500)}</w:tr>` +
    `</w:tbl>`;
  return pack(document);
}

/**
 * 一条缝**只有一头说了要画线**的表（#196 的续）：整张表一句 `w:tblBorders` 也没有，甲格自己说底边一条单线，
 * 它下面那一格（丙）顶边什么都没说。这一条缝在纸上是**同一条线**（`border-collapse:collapse` 把两格的边合成一条），
 * 所以导引线要认对面那一句：丙的顶边不许再留一条导引，否则画得出的墨外侧又压一条。
 */
export function makeSeamInkDocx(): Uint8Array {
  const cell = (label: string, tcBorders = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/>${tcBorders}</w:tcPr><w:p><w:r><w:t>${label}</w:t></w:r></w:p></w:tc>`;
  const document =
    `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>` +
    `<w:tr>${cell("甲", `<w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="0070C0"/></w:tcBorders>`)}${cell("乙")}</w:tr>` +
    `<w:tr>${cell("丙")}${cell("丁")}</w:tr>` +
    `</w:tbl>`;
  return pack(document);
}

/**
 * 「分布列」那三张表（#206），形状按授权 46 份量到的三类各摆一张（`.scratch/col-distribute-audience.py`：
 * 58 张多列表里本来就等宽 0 张、网格和被列数除不尽 44 张、含跨列格 32 张、`w:tblW` 58/58 等于网格和）：
 *  · `odd`——3 列 1,000 / 1,100 / 1,102（和 3,202，除不尽，余 1 缇要落在最左那一列）；每行三格各说自己那一列。
 *  · `even`——2 列各 1,200，本来就等宽：那一击在这张表上没活，得回绝且不进撤销栈。
 *  · `spanned`——3 列 1,500 / 2,000 / 500（和 4,000），中间那一列**上下两行都被跨列格盖着**：纸面上没有这一条边，
 *    于是均分只在看得见边的两列间分（1,500 + 500 → 1,000 / 1,000），跨列那两句（2,500 与 3,500）一个字节不动。
 * 三张都带 `w:tblW w:type="dxa"` 且说的就是网格和——这一击保住和，所以那一句必须原样留着（`w:tblW` 变了才是编出来的）。
 */
export function makeDistributeDocx(): Uint8Array {
  const cell = (label: string, w: number, span = 1): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ""}</w:tcPr>` +
    `<w:p><w:r><w:t>${label}</w:t></w:r></w:p></w:tc>`;
  const table = (grid: number[], rows: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="${grid.reduce((n, w) => n + w, 0)}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>${rows}</w:tbl>`;
  const document =
    table([1000, 1100, 1102],
      `<w:tr>${cell("甲一", 1000)}${cell("甲二", 1100)}${cell("甲三", 1102)}</w:tr>` +
        `<w:tr>${cell("乙一", 1000)}${cell("乙二", 1100)}${cell("乙三", 1102)}</w:tr>`) +
    table([1200, 1200],
      `<w:tr>${cell("丙一", 1200)}${cell("丙二", 1200)}</w:tr>` +
        `<w:tr>${cell("丁一", 1200)}${cell("丁二", 1200)}</w:tr>`) +
    table([1500, 2000, 500],
      `<w:tr>${cell("戊一", 1500)}${cell("戊跨两列", 2500, 2)}</w:tr>` +
        `<w:tr>${cell("己跨两列", 3500, 2)}${cell("己三", 500)}</w:tr>`) +
    `<w:p><w:r><w:t>表外的一段</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * 「分布行」那四张表（#207），形状按授权 46 份量到的四类各摆一张（`.scratch/row-distribute-audience.py`：
 * 58 张多行表里 38 张逐行说了 `w:trHeight`、20 张一句没说；说了的 38 张里本来就等高 6 张、和除不尽 24 张；
 * `w:hRule` 580 枚全说 `atLeast`、`exact` 0 枚；`w:vMerge` 226 枚却没有一行因此没有自己的横边）：
 *  · `unequal`——4 行 385 / 402 / 647 / 647（和 2,081，除不尽，余 1 缇落在最上那一行），每行都带 `w:hRule="atLeast"`。
 *    这一组数就是从 `purchase contract.dotm` 那张 4 行表照搬的（385 / 402 / 647 / 647，和 2,081）。
 *  · `equal`——3 行各 600：本来就等高，那一击在这张表上没活，得回绝且不进撤销栈。
 *  · `unstated`——3 行一句 `w:trHeight` 也没有（20/58 张就是这一形状）：行高只能从纸上量。
 *  · `merged`——3 行 200 / 400 / 600（和 1,200，除得尽），左列那两行是一带竖合并：行都有自己那条横边，
 *    所以哪一行都不该被摘出均分，`w:vMerge` 那两句一个字节不动。
 */
export function makeRowDistributeDocx(): Uint8Array {
  const cell = (label: string, vmerge = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/>${vmerge}</w:tcPr>` +
    `<w:p><w:r><w:t>${label}</w:t></w:r></w:p></w:tc>`;
  const row = (h: number | undefined, a: string, b: string, vmerge = ""): string =>
    `<w:tr>${h === undefined ? "" : `<w:trPr><w:trHeight w:val="${h}" w:hRule="atLeast"/></w:trPr>`}` +
    `${cell(a, vmerge)}${cell(b)}</w:tr>`;
  const table = (rows: string): string =>
    `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>${rows}</w:tbl>`;
  const document =
    table(
      row(385, "甲一", "甲二") +
        row(402, "乙一", "乙二") +
        row(647, "丙一", "丙二") +
        row(647, "丁一", "丁二")) +
    table(row(600, "戊一", "戊二") + row(600, "己一", "己二") + row(600, "庚一", "庚二")) +
    table(row(undefined, "辛一", "辛二") + row(undefined, "壬一", "壬二") + row(undefined, "癸一", "癸二")) +
    table(
      row(200, "子一", "子二", `<w:vMerge w:val="restart"/>`) +
        row(400, "", "丑二", `<w:vMerge/>`) +
        row(600, "寅一", "寅二"));
  return pack(document);
}

/**
 * 「显示原文」那一档的表层读数（#225）：一张表把三层记录都挂上，纸面与读数框都该穿记录里的 `prev`。
 *  · 表：现值 居中 / 缩进 284 / 固定列宽 / 左边距 108 / 地 DDEEFF / 顶边单线 12 红；记录说改前
 *    右对齐 / 无缩进 / 自动列宽 / 左边距 72 / 地 F2F2F2 / 顶边点线 4 黑。
 *  · 首行：现值 高 600 exact / 标题行 / 行对齐 center；记录说改前 高 300 / 允许断开 / 行对齐 left。
 *  · 首格：现值 宽 2000 / 地 FFF2CC / 底对齐 / 顶边单线 8 绿；记录说改前 宽 1500 / 地 DDEBF7 /
 *    顶对齐 / 顶边双线 4 品红。
 * 第二行一句没说，两档同脸——原文不是把整张表都翻个面。记录按 MS 的例子裹一层同名元素
 * （`<w:tblPrChange><w:tblPr>`、`<w:trPrChange><w:trPr>`），格那条直接写属性（同
 * `makeTablePropertyRevisionDocx` 的两种拼法）。未取证：WPS 自己写这几种记录的字面仍未见过。
 *
 * 四格的词用 卯/辰/巳/午 而不是 甲/乙/丙/丁：装机烟测里 #146/#149/#152 那几腿按「甲」「末」找表落笔，
 * 用同一批词会被它们改写，读数就不再是件里原本那一句（第一趟实测：格的底纹被填成 00B050、记录被接受掉）。
 */
export function makeOriginalTableReadDocx(): Uint8Array {
  const document =
    `<w:p><w:r><w:t>表前读数</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:jc w:val="center"/><w:tblInd w:w="284" w:type="dxa"/>` +
    `<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="FF0000"/></w:tblBorders>` +
    `<w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/><w:tblLayout w:type="fixed"/>` +
    `<w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar>` +
    `<w:tblPrChange ${REV_ATTRS(71, "顾峰")}><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:jc w:val="right"/><w:tblInd w:w="0" w:type="dxa"/>` +
    `<w:tblBorders><w:top w:val="dashed" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>` +
    `<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:tblLayout w:type="autofit"/>` +
    `<w:tblCellMar><w:left w:w="72" w:type="dxa"/></w:tblCellMar></w:tblPr></w:tblPrChange></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="exact"/><w:tblHeader/><w:jc w:val="center"/>` +
    `<w:trPrChange ${REV_ATTRS(72, "顾峰")}><w:trPr><w:trHeight w:val="300"/><w:cantSplit/><w:jc w:val="left"/></w:trPr></w:trPrChange></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>` +
    `<w:tcBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="00FF00"/></w:tcBorders>` +
    `<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:vAlign w:val="bottom"/>` +
    `<w:tcPrChange ${REV_ATTRS(73, "顾峰")}><w:tcW w:w="1500" w:type="dxa"/>` +
    `<w:tcBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="FF00FF"/></w:tcBorders>` +
    `<w:shd w:val="clear" w:color="auto" w:fill="DDEBF7"/><w:vAlign w:val="top"/></w:tcPrChange></w:tcPr>` +
    `<w:p><w:r><w:t>卯</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>辰</w:t></w:r></w:p></w:tc></w:tr>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>巳</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>午</w:t></w:r></w:p></w:tc></w:tr>` +
    `</w:tbl>` +
    `<w:p><w:r><w:t>表后读数</w:t></w:r></w:p>`;
  return pack(document);
}
