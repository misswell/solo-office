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
 *
 * The last four abstract numbers are the corpus's bullet spellings: `numId` 7 puts the mark in `w:lvlText` as a
 * private-use codepoint (164 of the corpus's 170 `bullet` levels do) and names the font that codepoint belongs
 * to in the level's own `w:rPr/w:rFonts` (Wingdings 158, Symbol 6, other 6), while `numId` 8 states no font at
 * all. `w:sym` is in none of the 86 files (`.scratch/sym-count.ts`), so this is the shape a bullet has to
 * survive here — drawn in the paragraph's own font, the codepoint is tofu.
 *
 * `numId` 9 and 10 add the size the level states for its own mark (`w:lvl/w:rPr/w:sz`). Of the 42 corpus
 * paragraphs that sit on a fonted bullet level, every one of them sits on a level that also writes a `w:sz`, and
 * in all 42 that size is *smaller* than the paragraph's own text resolves to; the attribute set of those levels
 * is `w:ascii + w:hAnsi + w:hint` (`numId` 9 spells it the same way), and 9 more levels carry a `w:sz` with no
 * font at all — the shape of `numId` 10 (`.scratch/num-level-rpr.ts`).
 *
 * `numId` 11 is the level that names a *family* but no face: `w:rFonts w:hint="eastAsia"` and nothing else, which
 * is how 358 of the 424 corpus paragraphs whose level states no font have it stated (`w:hint` on `eastAsia` for
 * 282 of them, on a value outside §17.18.39 for 76 — `.scratch/mark-font.ts`). The face such a mark ends up
 * wearing therefore comes from the paragraph's own `w:pPr/w:rPr/w:rFonts` (§17.3.1.4) or, when that names none,
 * from the chain: of those 424, 172 name the face for their own script on the mark, 234 more get it from a rung
 * like `MarkFaced` below, and 18 name nothing anywhere — those are the ones the page must leave alone.
 */
const WINGDINGS_MARK = String.fromCharCode(0xf0a7);
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
<w:abstractNum w:abstractNumId="6">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${WINGDINGS_MARK}"/><w:lvlJc w:val="left"/><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:cs="Wingdings"/></w:rPr></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="7">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${WINGDINGS_MARK}"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="8">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${WINGDINGS_MARK}"/><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="eastAsia"/><w:sz w:val="16"/></w:rPr></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="9">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr><w:sz w:val="18"/></w:rPr></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="10">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
<w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num>
<w:num w:numId="5"><w:abstractNumId w:val="4"/></w:num>
<w:num w:numId="6"><w:abstractNumId w:val="5"/></w:num>
<w:num w:numId="7"><w:abstractNumId w:val="6"/></w:num>
<w:num w:numId="8"><w:abstractNumId w:val="7"/></w:num>
<w:num w:numId="9"><w:abstractNumId w:val="8"/></w:num>
<w:num w:numId="10"><w:abstractNumId w:val="9"/></w:num>
<w:num w:numId="11"><w:abstractNumId w:val="10"/></w:num>
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
 *
 * `MarkSized` states the size on the *style's* `w:rPr`, the rung a paragraph's own mark falls through to.
 * Across the 86 WPS-written documents 455 of 466 listed paragraphs write `w:pPr/w:rPr` at all and 326 of them a
 * `w:sz` — and those 326 always pair it with a `w:szCs` (`<w:sz w:val="28"/><w:szCs w:val="28"/>` is the shape this
 * fixture copies, `.scratch/mark-size.ts`); a mark that states none takes its size from a rung like this one, which
 * is where 98 of the paragraphs no size-stating level covers get their number.
 *
 * `MarkFaced` is the same rung for the two faces: `w:rFonts` naming one for each script, the shape 234 of the
 * 424 paragraphs whose level names no face reach through the chain (`.scratch/mark-font.ts`). `numberingStyles`
 * takes the `w:rPrDefault` children as an argument because one test needs a document where *nothing* names a face
 * — the 18 paragraphs that must go on wearing whatever the page wears.
 *
 * `MarkColored` is that rung for the colour. Of the 473 listed paragraphs, 74 have a mark that states a
 * `w:rPr/w:color` — every one of them spelled `w:val` with six hex digits and nothing else on the tag, so no
 * paragraph in this corpus asks for a theme colour (`w:themeColor` appears on 0 of them, `.scratch/mark-color.ts`)
 * — and in 74/74 that colour is the same one the paragraph's own text resolves to: the files write the ink on the
 * mark as well as on the runs, and it is the *symbol* that had been left out of it. 8 more reach a rung like this
 * one, and 391 state nothing anywhere and keep whatever ink the page wears. The level side has no such paragraph:
 * 8 of the 617 levels do state a colour, and **no** listed paragraph sits on any of them — the same shape
 * §17.3.1.12's other level properties were counted in, so this fixture has no colour-stating level either.
 */
export const numberingStyles = (docDefaultRun = `<w:rFonts w:ascii="Calibri"/>`): string => `<w:styles xmlns:w="${NS_W}">
<w:docDefaults><w:rPrDefault><w:rPr>${docDefaultRun}<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Numbered"><w:name w:val="编号来自样式"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Unnumbered"><w:name w:val="样式里取消编号"/><w:basedOn w:val="Numbered"/><w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="MarkSized"><w:name w:val="段落标记自己写了字号"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="MarkFaced"><w:name w:val="段落标记从样式借脸"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="宋体" w:hAnsi="Arial"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="MarkColored"><w:name w:val="段落标记从样式借颜色"/><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="7030A0"/></w:rPr></w:style>
</w:styles>`;

const STYLES = numberingStyles();

/** One paragraph: `pPr` is the caller's own `w:pPr` children, written the way the files write them. */
export interface NumberedPara {
  pPr?: string;
  text: string;
}

/**
 * A document of the given paragraphs over `NUMBERING` (or a caller's own numbering part) and `numberingStyles()`
 * (or a caller's own styles part, when the test is about what the ladder states). The list membership itself is
 * the caller's `w:numPr`, so a test that wants the corpus's cancellation shape writes
 * `<w:numPr><w:numId w:val="0"/></w:numPr>` and one that wants the borrowed kind names a style instead.
 */
export function makeNumberingDocx(
  paras: readonly NumberedPara[],
  numbering: string = NUMBERING,
  styles: string = STYLES,
): Uint8Array {
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
      "word/styles.xml": strToU8(XML_DECL + styles),
      "word/numbering.xml": strToU8(XML_DECL + numbering),
    },
    { level: 6 },
  );
}

/** The list a paragraph joins, spelled the way the corpus spells it: `w:ilvl` first, then `w:numId`. */
export const numPr = (numId: number, ilvl?: number): string =>
  `<w:numPr>${ilvl === undefined ? "" : `<w:ilvl w:val="${ilvl}"/>`}<w:numId w:val="${numId}"/></w:numPr>`;
