import { strToU8, zipSync } from "fflate";

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/**
 * 目录 sdt 的最小件，形状照 `event plan.dotm` 的字面（#236 的受众，46 份里带目录的 3 份之一）：
 * `w:docPartGallery val="Table of Contents"`；段0 是居中的「目 录」（宋体加粗 36 半点）；条目段带
 * pStyle（toc 1/toc 2，toc 2 从链上借 `w:ind left=420`）与自己那句 `w:tabs` 右 8306 点线，条目字
 * 与页码是 HYPERLINK/PAGEREF 复杂域的**缓存结果**（fldChar 边界与 instrText 指令不是文字）。
 * 另带一格「内容里混着表」的 sdt（`TOC_MIXED` 拼进另一份件），钉住「不是枚枚是段就不逐行」的门。
 */
const TOC_SDT_CONTENT =
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>` +
  `<w:r><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/><w:b/><w:sz w:val="36"/></w:rPr><w:t>目 录</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="TocOne"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs></w:pPr>` +
  `<w:r><w:rPr><w:rFonts w:hint="eastAsia" w:asciiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia"/><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l _Toc27457 </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r><w:rPr><w:b/></w:rPr><w:t>1、市场分析</w:t></w:r>` +
  `<w:r><w:tab/></w:r>` +
  `<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc27457 \\h </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r><w:t>1</w:t></w:r>` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="TocTwo"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs></w:pPr>` +
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l _Toc27458 </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r><w:t>2、活动主题</w:t></w:r>` +
  `<w:r><w:tab/></w:r>` +
  `<w:r><w:t>2</w:t></w:r>` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

export const TOC_SDT =
  `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr>` +
  `<w:sdtContent>${TOC_SDT_CONTENT}</w:sdtContent></w:sdt>`;

/** 对照格：sdtContent 里混着一张表——不是「枚枚是段」，逐行的门不许开，退回整块一段的旧画法。 */
export const MIXED_SDT =
  `<w:sdt><w:sdtPr><w:alias w:val="混合"/></w:sdtPr><w:sdtContent>` +
  `<w:p><w:r><w:t>盒子里一段</w:t></w:r></w:p>` +
  `<w:tbl><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>盒子里一张表</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
  `</w:sdtContent></w:sdt>`;

export const TOC_STYLES =
  `${XML_DECL}<w:styles xmlns:w="${NS_W}">` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="TocOne"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="TocTwo"><w:name w:val="toc 2"/><w:basedOn w:val="TocOne"/><w:pPr><w:ind w:left="420" w:leftChars="200"/></w:pPr></w:style>` +
  `</w:styles>`;

const pack = (extra: string): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="${CT_STY}"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_PKG}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/></Relationships>`,
      ),
      "word/styles.xml": strToU8(TOC_STYLES),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${extra}<w:p><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr></w:body></w:document>`,
      ),
    },
    { level: 6 },
  );

/** 目录那份件：sdt（gallery 型目录）+ 一段正文 + sectPr。 */
export function makeTocDocx(): Uint8Array {
  return pack(TOC_SDT);
}

/** 混合内容那份件：sdtContent 里有一段加一张表。 */
export function makeMixedSdtDocx(): Uint8Array {
  return pack(MIXED_SDT);
}
