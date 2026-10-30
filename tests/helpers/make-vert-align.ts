import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/**
 * 授权 9 份 docx 里 `w:vertAlign` 的两种住处与两种值：6 枚**全在 `word/styles.xml`**（wpspreset01/02/03 各 2，
 * 全是 `superscript`），正文与页眉页脚 0 枚、`subscript` 0、明写 `baseline` 0（`.scratch/vertalign-audience.py`）。
 * 也就是说「上标/下标」这两条命令在家里是**新写**的一句，不是修补语料里本来就有的那一族——所以夹具要摆的正是
 * 语料没给的那几档：
 *
 * - 甲：run 什么 `w:rPr` 都没写（命令落下去只该多一句 `w:vertAlign`）。
 * - 乙：本来就上标（回显要亮，按回去要撤掉，且不许改写成 `baseline` —— WPS 真机写哪一句未取证）。
 * - 丙：本来就下标（上标按下去的「换一档」要把它换掉，不是叠两句）。
 * - 丁：明说 `baseline`（§17.3.2.24 的取消极；纸上不抬不落，回显两枚都不亮）。
 * - 戊：上标且自己写了 22 磅（渲染器把 `.75em` 的缩一圈套在内层，外层那句磅数仍要读得回来）。
 * - 尾段：留给烟测里那条无条件打字腿。
 *
 * `w:rPr` 孩子的次序照 `CT_RPr`（`sz` 在 `vertAlign` 之前），与 `write-docx.ts` 的 `RPR_ORDER` 同一条。
 */
const BODY =
  `<w:p><w:r><w:t>甲组 什么都没写的一段</w:t></w:r></w:p>` +
  `<w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>乙组 本来就上标</w:t></w:r></w:p>` +
  `<w:p><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>丙组 本来就下标</w:t></w:r></w:p>` +
  `<w:p><w:r><w:rPr><w:vertAlign w:val="baseline"/></w:rPr><w:t>丁组 明说在基线</w:t></w:r></w:p>` +
  `<w:p><w:r><w:rPr><w:sz w:val="44"/><w:szCs w:val="44"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>戊组 上标二十二磅</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>尾段 留给人打字的</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>` +
  `<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>`;

const STYLES =
  `<w:styles xmlns:w="${NS_W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
  `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/>` +
  `</w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`;

export function makeVertAlignDocx(): Uint8Array {
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="${CT_STY}"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${BODY}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + STYLES),
    },
    { level: 6 },
  );
}
