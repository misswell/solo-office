import { strToU8, zipSync } from "fflate";

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";

/**
 * 正文里那几张表各自坐在哪——按授权口径的字面拼，一张表只差 `w:tblPr` 里的那两枚。
 *
 * 出处（`.scratch/tbl-place-audience.py`，46 件 / 28 份模板名，逐枚问 `w:tblPr`）：58 张表里 `w:jc` 说
 * **`center` 14 枚 / 2 份模板名**（`left` / `right` / `start` / `end` 各 **0 枚**、谁都没说 44 枚）；
 * `w:tblInd` 说非 0 缇 **28 枚 / 3 份模板名**（−687…+288，其中负 **4 枚**），说 0 的 30 枚；**两句同时说的 8 枚**。
 * 节按有受众的那两份模板名摆：`bidding document.dotm` 纸 11906 缇、左 1588、右 1474——**栏宽 8844 缇**。
 * 那 14 枚居中里有 **5 枚比栏还宽**（9,231 / 9,500 / 9,637 / 9,803 / 9,806），所以过宽那一枚取 9,637。
 * 每张表 `w:tblW` 与 `w:tblGrid` 之和同字面（授权 58/58 就是这一形状），`w:tblLayout` 说 fixed（58/58）。
 *
 * 六格各有名字，断言里分得开：过宽居中（甲）、居中不说缩进（乙）、居中 +288（丙）、居中 −687（丁）、
 * 只说缩进没说居中（戊）、两句都没说（己，对照）。
 */

/** 节里那三枚缇量出来的栏宽：11906 − 1588 − 1474。 */
export const COL_TWIPS = 8844;

const p = (t: string): string => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

/** 一张表：`w:tblPr` 只带 `w:tblW` 与调用方给的那几枚，网格和 `w:tblW` 同字面、逐格说 `w:tcW`。 */
export function tableXml(tag: string, widthTwips: number, pr: string): string {
  const each = Math.floor(widthTwips / 4);
  const grid = [each, each, each, widthTwips - each * 3];
  const cells = grid.map((w, i) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${p(`${tag}-${i + 1}`)}</w:tc>`).join("");
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${widthTwips}" w:type="dxa"/>${pr}` +
    `<w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>` +
    `<w:tr>${cells}</w:tr><w:tr>${cells}</w:tr></w:tbl>`
  );
}

const CENTER = `<w:jc w:val="center"/>`;
const indent = (twips: number): string => `<w:tblInd w:w="${twips}" w:type="dxa"/>`;

/** 六张表，按甲乙丙丁戊己排；DOM 探针与装机烟测都按这一序问。 */
export const PLACE_TABLES =
  tableXml("甲", 9637, CENTER) +
  tableXml("乙", 5000, CENTER) +
  tableXml("丙", 5000, CENTER + indent(288)) +
  tableXml("丁", 5000, CENTER + indent(-687)) +
  tableXml("戊", 5000, indent(108)) +
  tableXml("己", 5000, "");

/** 对照：把六张表的 `w:jc` 与 `w:tblInd` 全剥掉——这一族只上横向，铺排（纸数与正文起点）该一字不差。 */
export const BARE_TABLES = PLACE_TABLES.replace(/<w:jc[^/]*\/>/g, "").replace(/<w:tblInd[^/]*\/>/g, "");

export function makeTablePlaceDocx(tables: string = PLACE_TABLES): Uint8Array {
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_REL}"><w:body>${p("表前一段")}${tables}${p("表后一段")}` +
          `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
          `<w:pgMar w:top="1440" w:right="1474" w:bottom="1440" w:left="1588" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
      ),
    },
    { level: 6 },
  );
}
