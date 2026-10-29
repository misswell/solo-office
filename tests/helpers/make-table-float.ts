import { strToU8, zipSync } from "fflate";

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";

/**
 * #218 浮动表（`w:tblPr/w:tblpPr`，ECMA-376 §17.4.58）的夹具：一张按纸坐标钉住的表，前后各有一串正文段。
 *
 * 授权 46 件里 2 枚 / 2 件——`shareholding agreement.dotm` 的 zh_CN 与 zh_HK，两句逐字节相同：
 * `<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="page" w:horzAnchor="page"
 * w:tblpX="1117" w:tblpY="10912"/>`。这一份件照那一句字面写，节也照那两份件的节：纸 11906×16838、
 * 上下 1440、左右 1080（`twipsToPx` 15.0001 缇/px → 表左上角在纸左 74.47px、纸上 727.47px）。
 *
 * 前 90 段是为了把「锚」推到第 2 张往后：锚页取表**后面第一块**所在的张（与 `applySheetGaps` 同一条判据），
 * 90 段无论行高怎么算都高过一页（正文净高 930.5px）。后 6 段保证锚之后仍有正文。
 *
 * 另有两张「说了 `w:tblpPr` 也不该离流」的对照：`w:tblpXSpec`/`w:tblpYSpec`（规格锚）与 `text` 锚。
 * 这两支这一片不画——规格档的 9 种取值要逐档取证，`text` 锚的绕排没建模；留在流里是不猜的落笔。
 */
export const TBLP_X_TWIPS = 1117;
export const TBLP_Y_TWIPS = 10912;
export const TBLP_W_TWIPS = 10314;

const p = (t: string): string => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

const cell = (t: string): string => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;

/** 一张两行两列的表；`pr` 是调用方原样写进 `w:tblPr` 的那几枚。 */
const table = (tag: string, widthTwips: number, pr: string): string => {
  const each = Math.floor(widthTwips / 2);
  const row = `<w:tr>${cell(`${tag}-1`)}${cell(`${tag}-2`)}</w:tr>`;
  return (
    `<w:tbl><w:tblPr>${pr}<w:tblW w:w="${widthTwips}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="${each}"/><w:gridCol w:w="${widthTwips - each}"/></w:tblGrid>${row}${row}</w:tbl>`
  );
};

/** 真身那一句，逐字节照抄（两件里的只差节，表这一句一字不差）。 */
export const FLOAT_PR =
  `<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="page" w:horzAnchor="page" ` +
  `w:tblpX="${TBLP_X_TWIPS}" w:tblpY="${TBLP_Y_TWIPS}"/>`;

/** 对照甲：规格锚（`w:tblpXSpec`/`w:tblpYSpec`）——这一片不画，留在流里。 */
export const SPEC_PR = `<w:tblpPr w:vertAnchor="page" w:horzAnchor="page" w:tblpXSpec="center" w:tblpYSpec="bottom"/>`;

/** 对照乙：`text` 锚——绕排没建模，留在流里。 */
export const TEXT_PR = `<w:tblpPr w:vertAnchor="text" w:horzAnchor="text" w:tblpX="${TBLP_X_TWIPS}" w:tblpY="${TBLP_Y_TWIPS}"/>`;

export const FLOAT_TABLE = table("浮动", TBLP_W_TWIPS, FLOAT_PR);
export const SPEC_TABLE = table("规格锚", 5000, SPEC_PR);
export const TEXT_TABLE = table("文字锚", 5000, TEXT_PR);
export const PLAIN_TABLE = table("普通表", 5000, "");

/** 基准段数：90 段正文 + 浮动表 + 6 段正文 + 三张对照表，每张表后各一段。 */
const fill = (from: number, to: number, prefix: string): string =>
  Array.from({ length: to - from + 1 }, (_, i) => p(`${prefix}${from + i}`)).join("");

const BODY_TAIL = `${fill(91, 96, "浮动表之后第")}${p("规格锚之前")}${SPEC_TABLE}${p("规格锚之后")}${TEXT_TABLE}${p("文字锚之后")}${PLAIN_TABLE}${p("普通表之后")}`;

/**
 * `floating=false` 造一份**没有浮动表**的同形件：用来量「浮动表不占流」——两张的纸数与浮动表后面那段
 * 正文的上沿都该一字不差。
 */
export function makeTblpPrDocx(floating = true): Uint8Array {
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
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_REL}"><w:body>${fill(1, 90, "表前第")}${floating ? FLOAT_TABLE : ""}${BODY_TAIL}` +
          `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
          `<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="708" w:footer="170" w:gutter="0"/>` +
          `<w:cols w:space="708" w:num="1"/></w:sectPr></w:body></w:document>`,
      ),
    },
    { level: 6 },
  );
}
