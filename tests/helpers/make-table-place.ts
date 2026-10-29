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

/** 把一枚 `w:trPr` 插进**最后**那一行——件里的行各有自己的说法，两张一模一样的行问不出「逐行派发」。 */
const withRowPr = (tbl: string, trPr: string): string => {
  const at = tbl.lastIndexOf("<w:tr>");
  return tbl.slice(0, at + "<w:tr>".length) + trPr + tbl.slice(at + "<w:tr>".length);
};

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

/**
 * #214 那三张「等着被改」的表——名字一律避开 甲/末：装机探针里 #146/#190 那几腿按 `甲` 认表，
 * 会被它们牵动的字节账不属于这一族。三张各占一种起始态，落笔之后盘上该是三种不同形状：
 * - 摆1 两句都没说 → 点「居中」+ 左缩进 0.5 厘米（= 283 缇）：盘上凭空多出两枚，纸上走 `calc` 那一支；
 * - 摆2 只说居中 → 原样确定：一个字节也不该动（纸上仍是 `calc(50% - 166.66666666666666px)`，没有带号的那一截）；
 * - 摆3 说居中 + 288 → 只把左缩进填成 `0`：盘上那句是**明写的 0**（`w:w="0"`），不是删掉——
 *   「说了 0」与「没说」在纸上画到同一处，在盘上是两句话（授权 46 份里 `w:tblInd` 说 0 的有 30 枚，见 §154）。
 *
 * #216 给摆3 的**末行**另加一句行的话（`w:trPr/w:jc`）：表那三句问的是「逐张表」，行那一句问的是「逐行」，
 * 两张一模一样的行问不出这条派发。那一枚的形状是数出来的（`.scratch/row-jc-shape-audience.py`：授权 46 件里
 * 190 枚说的全是 `center`，孩子序只有 `cantSplit+trHeight+jc`(18)、`trHeight+jc`(154)、`jc`(18) 三种，`jc` 永远收尾），
 * 这里取最常见那一支 `trHeight+jc`——**「左对齐」把 `jc` 摘掉时 `trHeight` 得留在原位**，所以兄弟孩子必须有。
 * 摆1 与摆2 的行一句都没说：摆1 要问「凭空多出这一枚」，摆2 要问「没点名的不许被发明成 0/center」。
 */
export const EDIT_TABLES =
  tableXml("摆1", 5000, "") +
  tableXml("摆2", 5000, CENTER) +
  withRowPr(tableXml("摆3", 5000, CENTER + indent(288)), `<w:trPr><w:trHeight w:val="600"/>${CENTER}</w:trPr>`);

/** 六张摆位表 + 三张待改表：#214 的装机件用这一份，前六张的形状与 #213 逐字相同。 */
export const PLACE_AND_EDIT = PLACE_TABLES + EDIT_TABLES;

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
