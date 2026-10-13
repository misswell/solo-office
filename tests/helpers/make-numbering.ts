import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_NUM = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

/**
 * The list definitions, in the spellings the 86 WPS-written documents use. `.scratch/numfmt-census.ts` counts
 * every `w:lvl` in their `word/numbering.xml` (617 of them, with `w:lvlOverride` in none), and the shapes this
 * fixture gathers are the ones it found:
 *
 * - every one of the 617 levels states its own `w:lvlText`, so the symbol in front of a paragraph always comes
 *   from a template. What follows the placeholder is 、 for 91 of the 92 `chineseCounting` levels, ")" for 72 of
 *   the 75 `lowerLetter` ones, and only 84 of the 205 `decimal` ones.
 * - `w:numFmt` is `decimal` 205, `bullet` 170, `chineseCounting` 92, `lowerLetter` 75, `lowerRoman` 73,
 *   `decimalEnclosedCircleChinese` 1, `none` 1 — which is why `ordinal` had to gain the last two.
 * - `w:start` appears with 1 more often than not, and the rest of the values are small — 2 and 10 are both in
 *   the census, so a fixture that wants a sequence to start elsewhere asks for it with `w:start`.
 *
 * `w:lvlJc` and `w:pPr/w:ind` inside a level are the files' own extras; this model reads only the left indent
 * out of them (`indent`), and nothing here draws from them.
 */
export const NUMBERING = `<w:numbering xmlns:w="${NS_W}">
<w:abstractNum w:abstractNumId="0">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/><w:lvlJc w:val="left"/></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="（%2）"/><w:lvlJc w:val="left"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="1">
<w:lvl w:ilvl="0"><w:start w:val="10"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="2">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="2"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="3">
<w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="4">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
<w:lvl w:ilvl="1"><w:numFmt w:val="decimalEnclosedCircleChinese"/><w:lvlText w:val="%2"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="5">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1)"/></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="(%2)"/></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
<w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num>
<w:num w:numId="5"><w:abstractNumId w:val="4"/></w:num>
<w:num w:numId="6"><w:abstractNumId w:val="5"/></w:num>
</w:numbering>`;

/**
 * `Normal` besides `Numbered` is there because the ladder ends at the default paragraph style, and a fixture
 * about a paragraph taking its list from a style should not be one whose chain has no bottom.
 *
 * `Numbered` lends a whole `w:numPr` (§17.3.4.25), and `Unnumbered` lends a cancellation over it: `w:ilvl`
 * alone, the shape of 2 of the corpus's 58 style-stated `w:numPr` elements (`.scratch/style-numpr-census.ts`)
 * and the one 11 of its 48 chain-borrowing paragraphs hit (`.scratch/num-lent-shapes.ts`), which `numPrOf`
 * reads as numId 0. `Unnumbered` naming `Numbered` as its `w:basedOn` is the point of the
 * pair — the leaf states the property, so the ancestor's list gets no say.
 */
const STYLES = `<w:styles xmlns:w="${NS_W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Numbered"><w:name w:val="编号来自样式"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Unnumbered"><w:name w:val="样式里取消编号"/><w:basedOn w:val="Numbered"/><w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:style>
</w:styles>`;

/** One paragraph: `pPr` is the caller's own `w:pPr` children, written the way the files write them. */
export interface NumberedPara {
  pPr?: string;
  text: string;
}

/**
 * A document of the given paragraphs over `NUMBERING` (or a caller's own numbering part), for the tests about
 * which symbol a numbered paragraph draws. The list membership itself is the caller's `w:numPr`, so a test that
 * wants the corpus's cancellation shape writes `<w:numPr><w:numId w:val="0"/></w:numPr>` and one that wants the
 * borrowed kind names a style instead.
 */
export function makeNumberingDocx(paras: readonly NumberedPara[], numbering: string = NUMBERING): Uint8Array {
  const document =
    paras
      .map((e) => `<w:p>${e.pPr ? `<w:pPr>${e.pPr}</w:pPr>` : ""}<w:r><w:t>${e.text}</w:t></w:r></w:p>`)
      .join("\n") + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${CT_NUM}"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${NS_R}/numbering" Target="numbering.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${document}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + STYLES),
      "word/numbering.xml": strToU8(XML_DECL + numbering),
    },
    { level: 6 },
  );
}

/** The list a paragraph joins, spelled the way the corpus spells it: `w:ilvl` first, then `w:numId`. */
export const numPr = (numId: number, ilvl?: number): string =>
  `<w:numPr>${ilvl === undefined ? "" : `<w:ilvl w:val="${ilvl}"/>`}<w:numId w:val="${numId}"/></w:numPr>`;
