import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/**
 * The ids the fixtures name. They look like the ones WPS writes — opaque digits with the readable name in
 * `w:name` — because the gallery shows `w:name` and writes `w:styleId`, and a fixture whose ids were the names
 * would not catch a reader that mixed the two up.
 */
export const DEFAULT_STYLE = "1";
export const HEADING_STYLE = "2";
export const QUOTE_STYLE = "3";
export const EMPHASIS_CHAR_STYLE = "7";

/**
 * One default paragraph style, a heading two steps down a chain, a body style with an indent, and one character
 * style — the four shapes the gallery has to keep apart. The chain is what the corpus' own heading styles are
 * (`.scratch/para-style-gallery-audience.py` reads 112 paragraph styles across the authorized docx, 103 with
 * `w:qFormat`, the default one named `Normal` in 6/6):
 *
 * - `Normal` is `w:default="1"` and states nothing beyond its name, so a paragraph that names it and a paragraph
 *   that names no style draw the same.
 * - `标题 1` states `w:outlineLvl`, a `w:keepNext`, a `w:spacing` and a `w:rPr` of its own (bold, 16pt, a dark
 *   blue, a Latin and a Chinese face), which is the shelf the gallery has to move a paragraph onto whole.
 * - `引用` is a second body style, distinguished only by an indent and an italic — the case where switching styles
 *   changes the paragraph's indent and the runs' slant without the paragraph stating either.
 * - `强调` is a character style: a run that names it keeps its own face and ink when the paragraph moves, because a
 *   character style sits above the paragraph's chain in the run ladder (§17.7.2).
 *
 * The `w:pPr` children appear in `CT_PPr` order (`keepNext` … `spacing` … `ind` … `outlineLvl`), as do the `w:rPr`
 * ones (`rStyle`, `rFonts`, `b`, `color`, `sz`, `szCs`), so a fixture paragraph is shaped like one the files write.
 */
const STYLES =
  `<w:styles xmlns:w="${NS_W}">` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/>` +
  `<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="${DEFAULT_STYLE}"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="${HEADING_STYLE}"><w:name w:val="标题 1"/><w:basedOn w:val="${DEFAULT_STYLE}"/>` +
  `<w:next w:val="${DEFAULT_STYLE}"/><w:qFormat/>` +
  `<w:pPr><w:keepNext/><w:spacing w:before="340" w:after="330" w:line="576" w:lineRule="auto"/><w:outlineLvl w:val="0"/></w:pPr>` +
  `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="黑体"/><w:b/><w:color w:val="1F3864"/>` +
  `<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="${QUOTE_STYLE}"><w:name w:val="引用"/><w:basedOn w:val="${DEFAULT_STYLE}"/><w:qFormat/>` +
  `<w:pPr><w:ind w:left="567"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="${EMPHASIS_CHAR_STYLE}"><w:name w:val="强调"/>` +
  `<w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:color w:val="C00000"/></w:rPr></w:style>` +
  `</w:styles>`;

/**
 * Five paragraphs, each one of the shapes the gallery meets, plus a plain tail for the smoke legs that type at the
 * end of the body:
 *
 * - 甲 names no style and states nothing: switching it to the heading must leave a `w:pStyle` and nothing else.
 * - 乙 states an indirect indent and has a bold run: both are its own, so both survive the switch, and the spacing
 *   the heading lends must not be written down as if 乙 had said it.
 * - 丙 already names the heading: switching it back to the default must take `w:pStyle` away, not name `Normal`.
 * - 丁's run names the character style 强调: its face and ink come from there and stay put while the paragraph's
 *   chain lends the italic of 引用 — the ladder's order in one paragraph.
 * - 戊 names 引用, the shape 「换成标题 1 再换回来」 starts from.
 */
const BODY =
  `<w:p><w:r><w:t>甲组 什么都没说的一段</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>乙组 自己写了缩进的一段</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="${HEADING_STYLE}"/></w:pPr><w:r><w:t>丙组 已是标题的一段</w:t></w:r></w:p>` +
  `<w:p><w:r><w:rPr><w:rStyle w:val="${EMPHASIS_CHAR_STYLE}"/></w:rPr><w:t>丁组 挂着字符样式的一段</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="${QUOTE_STYLE}"/></w:pPr><w:r><w:t>戊组 引用样式的一段</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>尾段 留给人打字的</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>` +
  `<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>`;

export function makeStyleGalleryDocx(): Uint8Array {
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
