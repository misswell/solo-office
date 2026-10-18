import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml";

const REV = (id: number, author: string): string => `w:id="${id}" w:author="${author}" w:date="2026-03-01T08:00:00Z"`;

const pack = (files: Record<string, string>): Uint8Array => {
  const out: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) out[k] = strToU8(XML_DECL + v);
  return zipSync(out, { level: 6 });
};

const SKELETON = (document: string): Record<string, string> => ({
  "[Content_Types].xml": `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/></Types>`,
  "_rels/.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  "word/document.xml": document,
});

/**
 * A document with tracked changes in it, as Word and WPS write them: an insertion and a deletion
 * (whose text is spelled `w:delText`) side by side in one paragraph, a revision inside a table
 * cell, an insertion that carries a hyperlink, and a formatting revision (`w:rPrChange`) this
 * model reads the text of but does not rebuild.
 */
export function makeRevisedDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:r><w:t>未修订的段落</w:t></w:r></w:p>
<w:p><w:r><w:t>前</w:t></w:r><w:ins ${REV(1, "顾峰")}><w:r><w:t>插入的字</w:t></w:r></w:ins><w:del ${REV(2, "林晓")}><w:r><w:delText xml:space="preserve">删掉的字</w:delText></w:r></w:del><w:r><w:t>后</w:t></w:r></w:p>
<w:p><w:r><w:t>加粗的</w:t></w:r><w:ins ${REV(3, "顾峰")}><w:r><w:rPr><w:b/></w:rPr><w:t>加粗插入</w:t></w:r></w:ins></w:p>
<w:p><w:ins ${REV(4, "林晓")}><w:hyperlink r:id="rId9"><w:r><w:t>链接里的插入</w:t></w:r></w:hyperlink></w:ins></w:p>
<w:p><w:r><w:rPr><w:rPrChange ${REV(5, "顾峰")}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>改过格式的字的</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="120"/><w:gridCol w:w="120"/></w:tblGrid>
<w:tr><w:tc><w:p><w:ins ${REV(6, "顾峰")}><w:r><w:t>格内插入</w:t></w:r></w:ins></w:p></w:tc><w:tc><w:p><w:r><w:t>普通格</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
</w:body>
</w:document>`;

  const out: Record<string, string> = SKELETON(document);
  out["word/_rels/document.xml.rels"] = `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId9" Type="${NS_R}/hyperlink" Target="https://example.com/revised" TargetMode="External"/></Relationships>`;
  return pack(out);
}

/**
 * The shape Word writes when revised text is revised again: an insertion that holds a second
 * `w:ins`, each signed by its own author. A reader that walks only the direct children of the outer
 * revision loses the inner text, and a writer that flattens the two makes one change of two.
 */
export function makeNestedRevisionDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:r><w:t>正文</w:t></w:r><w:ins ${REV(11, "顾峰")}><w:r><w:t>外层插入</w:t></w:r><w:ins ${REV(12, "林晓")}><w:r><w:t>内层再插入</w:t></w:r></w:ins></w:ins></w:p>
<w:p><w:del ${REV(13, "顾峰")}><w:r><w:delText xml:space="preserve">外层删除</w:delText></w:r><w:del ${REV(14, "林晓")}><w:r><w:delText xml:space="preserve">内层再删除</w:delText></w:r></w:del></w:del></w:p>
</w:body>
</w:document>`;

  return pack(SKELETON(document));
}

/**
 * A tracked move, which Word writes as `w:moveFrom` at the place the text left and `w:moveTo` at the
 * place it arrived. Both halves carry the same characters, so a reader that knows only `w:ins` and
 * `w:del` reads this document with the moved text missing entirely.
 */
export function makeMovedDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:r><w:t>开头</w:t></w:r><w:moveFrom ${REV(21, "顾峰")}><w:r><w:delText xml:space="preserve">搬走的字</w:delText></w:r></w:moveFrom><w:r><w:t>结尾</w:t></w:r></w:p>
<w:p><w:r><w:t>新位置</w:t></w:r><w:moveTo ${REV(22, "顾峰")}><w:r><w:t>搬走的字</w:t></w:r></w:moveTo></w:p>
</w:body>
</w:document>`;
  return pack(SKELETON(document));
}

/**
 * Revisions on paragraphs rather than on runs: a break somebody inserted (`w:pPr/w:rPr/w:ins` on the
 * paragraph it ends), a break somebody deleted, the same deleted break inside a table cell, and a
 * paragraph whose properties were revised (`w:pPrChange`). A reader that knows only run-level
 * `w:ins`/`w:del` sees nothing at all of the first three, and the paragraph's earlier alignment is
 * gone with the fourth.
 */
export function makeParaMarkDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:pPr><w:rPr><w:ins ${REV(31, "顾峰")}/></w:rPr></w:pPr><w:r><w:t>回车分出的前半</w:t></w:r></w:p>
<w:p><w:r><w:t>回车分出的后半</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:del ${REV(32, "林晓")}/></w:rPr></w:pPr><w:r><w:t>合并的上段</w:t></w:r></w:p>
<w:p><w:r><w:t>合并的下段</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange ${REV(33, "顾峰")}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>改过对齐的段落</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="120"/></w:tblGrid>
<w:tr><w:tc><w:p><w:pPr><w:rPr><w:del ${REV(34, "顾峰")}/></w:rPr></w:pPr><w:r><w:t>格内上段</w:t></w:r></w:p><w:p><w:r><w:t>格内下段</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
</w:body>
</w:document>`;

  return pack(SKELETON(document));
}

/**
 * Almost nothing revised, so an edit's record is the only one there is: four body paragraphs of distinct
 * text — one centred, one bold, one carrying somebody else's `w:pPrChange` — and a two-cell table. Used
 * to watch a paragraph split and join again: which half keeps the alignment, which keeps the bold, and
 * whose record survives the reconstruction.
 */
export function makeBreakableDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>居中的第一段</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>第二段文字</w:t></w:r></w:p>
<w:p><w:pPr><w:pPrChange ${REV(70, "林晓")}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>有人改过对齐的段</w:t></w:r></w:p>
<w:p><w:r><w:t>第四段文字</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="120"/><w:gridCol w:w="120"/></w:tblGrid>
<w:tr><w:tc><w:p><w:r><w:t>表格左格</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>表格右格</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
</w:body>
</w:document>`;

  return pack(SKELETON(document));
}

/**
 * The shape an editor writes when a selection reached over a paragraph break: one deletion signed with a
 * single `w:id` but split into two `w:del` elements because the struck characters sit in two different
 * paragraphs, the break those characters terminated marked deleted in the upper paragraph's `w:pPr`, and
 * the characters typed in their place in between. A reader that lists one row per element would show the
 * same change twice and let a reviewer decide half of it; a reader that lists one row per `w:id` must
 * still find both halves' text and each half's own formatting.
 */
export function makeStraddleDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:del ${REV(73, "审阅者甲")}/></w:rPr></w:pPr><w:r><w:t>居中的第一</w:t></w:r><w:del ${REV(71, "审阅者甲")}><w:r><w:delText xml:space="preserve">段</w:delText></w:r></w:del><w:ins ${REV(72, "审阅者甲")}><w:r><w:t>换</w:t></w:r></w:ins></w:p>
<w:p><w:del ${REV(71, "审阅者甲")}><w:r><w:rPr><w:b/></w:rPr><w:delText xml:space="preserve">第</w:delText></w:r></w:del><w:r><w:rPr><w:b/></w:rPr><w:t>二段文字</w:t></w:r></w:p>
<w:p><w:pPr><w:pPrChange ${REV(70, "林晓")}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>有人改过对齐的段</w:t></w:r></w:p>
</w:body>
</w:document>`;

  return pack(SKELETON(document));
}

/**
 * A whole paragraph moved, which is the case where a move signs the *break* rather than characters: the
 * source paragraph's mark carries `w:moveFrom` and the paragraph that arrived carries `w:moveTo`, both in
 * `w:pPr/w:rPr`. A reader that only knows `w:ins` and `w:del` here loses the fact that anybody moved
 * anything, and a writer that rebuilds the paragraph drops the tag and with it the reviewer's clue.
 */
export function makeMoveMarkDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:pPr><w:rPr><w:moveFrom ${REV(41, "顾峰")}/></w:rPr></w:pPr><w:r><w:t>被搬走的那段</w:t></w:r></w:p>
<w:p><w:r><w:t>留下的段落</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:moveTo ${REV(42, "顾峰")}/></w:rPr></w:pPr><w:r><w:t>搬到新处的那段</w:t></w:r></w:p>
</w:body>
</w:document>`;

  return pack(SKELETON(document));
}

/**
 * Revisions whose `w:id`s run backwards: the paragraph record carries the lowest id and the run
 * records higher ones. A 审阅窗格 that numbered changes by id would list them in the wrong order, and
 * an editor that handed out ids from 1 would sign its own change with a number the file already uses.
 */
export function makeLowIdDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:pPr><w:rPr><w:ins ${REV(3, "顾峰")}/></w:rPr></w:pPr><w:r><w:t>低号的插入</w:t></w:r><w:r><w:rPr><w:rPrChange ${REV(2, "林晓")}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>低号的格式</w:t></w:r></w:p>
<w:p><w:pPr><w:pPrChange ${REV(1, "顾峰")}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>低号的对齐段</w:t></w:r></w:p>
</w:body>
</w:document>`;

  return pack(SKELETON(document));
}

/**
 * A table whose rows were added and removed while changes were tracked. The record for a whole row sits
 * on the row itself (`w:trPr/w:ins`, `w:trPr/w:del`), and the text inside the row carries its own
 * signatures with their own ids — so one row of the file is several `w:id`s, and a reader that only
 * looks inside the cells loses the row's own change and writes back a plain table.
 *
 * The shape is from ECMA-376 (`EG_TrPrChanges`) plus this fixture: the corpus has no revised table, so
 * nothing here is WPS evidence.
 */
export function makeTableRowDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:r><w:t>表格前的段落</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblStyle w:val="a7"/></w:tblPr><w:tblGrid><w:gridCol w:w="120"/><w:gridCol w:w="120"/></w:tblGrid>
<w:tr><w:trPr><w:del ${REV(31, "顾峰")}/></w:trPr><w:tc><w:p><w:del ${REV(32, "顾峰")}><w:r><w:delText xml:space="preserve">删掉的行</w:delText></w:r></w:del></w:p></w:tc><w:tc><w:p><w:del ${REV(33, "林晓")}><w:r><w:delText xml:space="preserve">它的第二格</w:delText></w:r></w:del></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>留下的行</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>普通格</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:trPr><w:ins ${REV(34, "林晓")}/></w:trPr><w:tc><w:p><w:r><w:t>新增的行</w:t></w:r></w:p></w:tc><w:tc><w:p><w:ins ${REV(35, "林晓")}><w:r><w:t>格里的插入</w:t></w:r></w:ins></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>表格后的段落</w:t></w:r></w:p>
</w:body>
</w:document>`;

  return pack(SKELETON(document));
}

/**
 * The same record on the only row a table has: accepting that row's deletion leaves no row at all, which
 * is the case a writer has to answer for rather than emit a `w:tbl` with nothing in it.
 */
export function makeLastRowDeletedDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:r><w:t>唯一的行被删了</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="120"/></w:tblGrid>
<w:tr><w:trPr><w:del ${REV(41, "顾峰")}/></w:trPr><w:tc><w:p><w:del ${REV(42, "顾峰")}><w:r><w:delText xml:space="preserve">最后一行</w:delText></w:r></w:del></w:p></w:tc></w:tr>
</w:tbl>
</w:body>
</w:document>`;

  return pack(SKELETON(document));
}

/**
 * 五族属性改动各一枚、外加一枚「这一格不在了」的结构记录，凑在同一件里给「显示原文」那一档当靶子：
 * `w:rPrChange`（现值粗体、改前斜体）、`w:pPrChange`（现值居中、改前靠右）、`w:tblPrChange`（表宽 3000→1500 缇）、
 * `w:trPrChange`（行高 700→300 缇）、`w:tcPrChange`（格的宽与地 FF0000→00FF00），以及一枚 `w:cellDel`——
 * 它记的是格子在不在，套里没有属性可摆，视图若把它当属性套，就把这格从没说过的属性抹了。
 * 授权 46 件里这七枚记录**一枚都没有**（`.scratch/rev-format-audience.py`），所以这一族的受众只有自家写出的
 * 修订件；记录按 MS 的例子拼（`w:tblPrChange` 里裹 `w:tblPr`），不是照 WPS 的字面（那仍未取证）。
 */
export function makeFormatViewDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:r><w:rPr><w:b/><w:rPrChange ${REV(80, "顾峰")}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>族一 run</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange ${REV(81, "顾峰")}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>族二 段落</w:t></w:r></w:p>
<w:tbl>
<w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblPrChange ${REV(82, "顾峰")}><w:tblPr><w:tblW w:w="1500" w:type="dxa"/></w:tblPr></w:tblPrChange></w:tblPr>
<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="600"/></w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="700"/><w:trPrChange ${REV(83, "顾峰")}><w:trPr><w:trHeight w:val="300" w:hRule="atLeast"/></w:trPr></w:trPrChange></w:trPr>
<w:tc><w:tcPr><w:tcW w:w="1200" w:type="dxa"/><w:shd w:val="clear" w:fill="FF0000"/><w:tcPrChange ${REV(84, "顾峰")}><w:tcPr><w:tcW w:w="900" w:type="dxa"/><w:shd w:val="clear" w:fill="00FF00"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>族五 格</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:tcW w:w="600" w:type="dxa"/><w:cellDel ${REV(85, "林晓")}/></w:tcPr><w:p><w:r><w:t>删掉的一格</w:t></w:r></w:p></w:tc>
</w:tr>
</w:tbl>
</w:body>
</w:document>`;
  return pack(SKELETON(document));
}
