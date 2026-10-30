import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/**
 * 「更改大小写」那五档（WPS `CT_Home.kuip:121` 的折叠钮 `Font1`，`ksoCmd="ChangeCase"`）的夹具。
 * 这一族改的是**字面**，所以每一段各摆一处「改字面时最容易碰坏的东西」：
 *
 * - 甲：一句 run 里汉字与拉丁并排——汉字没有大小写，改完必须一字不少一字不多。
 * - 乙：三枚 run，中间那枚只有 `w:tab` 没有字面。它在模型偏移里占一枚 `NUL`：改写不许把它抹掉，
 *   也不许把它的位子让给别的字（`w:tab` 从一枚变成两枚、或干脆消失，都是这一族的降级）。
 * - 丙：`ß`。`toUpperCase()` 把它摊成 `SS`（长度 +1），家里按「长度不变」那一档留着原字面——
 *   WPS 真机在这一点上怎么落**未取证**，所以这一段专门钉住「没有变长」。
 * - 丁：自己写了粗与 16 磅的一句。改字面只动 `w:t`，`w:rPr` 那两句得原样站着（纸面上还是粗、还是 21.33px）。
 * - 尾段：留给烟测里那条无条件打字腿。
 */
const BODY =
  `<w:p><w:r><w:t>甲组 hello 世界 beta</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>乙组 one</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>number</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>丙组 straße 大写</w:t></w:r></w:p>` +
  `<w:p><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr><w:t>丁组 WORLD hello</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>尾段 留给人打字的</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>` +
  `<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>`;

const STYLES =
  `<w:styles xmlns:w="${NS_W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
  `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/>` +
  `</w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`;

export function makeCaseDocx(): Uint8Array {
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
