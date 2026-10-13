import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

const pack = (document: string, styles: string): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${document}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + styles),
    },
    { level: 6 },
  );

const REV = (id: number, author: string): string => `w:id="${id}" w:author="${author}" w:date="2026-03-01T08:00:00Z"`;

const STYLES = `<w:styles xmlns:w="${NS_W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
</w:styles>`;

/**
 * Paragraphs that say far more about themselves than this model has a field for, in the shapes WPS's own
 * files were measured in (`.scratch/wps-ppr-names.ts` over 86 WPS-written documents, 45949 paragraphs of
 * which 45794 carry their own `w:pPr` set): the `w:pPr` child census there is `rPr` 45402, `ind` 29011,
 * `spacing` 24877, `jc` 22340, `widowControl` 8625, `shd` 7836, `textAlignment` 7696, `pStyle` 7324,
 * `keepNext`/`keepLines` 6440, `suppressLineNumbers` 5221, `wordWrap` 5045, `snapToGrid` 2531, `bidi` 2144,
 * `tabs` 2058, `pageBreakBefore`/`adjustRightInd` 1297, `kinsoku`/`overflowPunct`/`topLinePunct`/`autoSpaceDE`/
 * `autoSpaceDN` 1293, `numPr` 981, `outlineLvl` 264, `framePr` 80, `sectPr` 29, `pBdr` 4, and the paragraph
 * mark's `w:rPr` holds `rFonts` 44968, `szCs` 42429, `sz` 41083, `color` 26168, `rStyle` 9342, `lang` 8185,
 * `kern` 7688, `spacing` 5833, `highlight` 1939, `textFill` 1442 and `ligatures` 1089 among them.
 *
 * Before the paragraph carried its own `w:pPr` child by child, a save rebuilt every paragraph from its typed
 * fields and deleted the rest of them — measured at 45402 lost `w:pPr/w:rPr`, 41091 `sz`, 26688 `color`,
 * 21188 `w:ind`, 7836 `w:shd` and 1939 `w:highlight` across the same corpus (`.scratch/wps-shd-probe.ts`).
 *
 * Four of its children the model does own (`pStyle`, `pageBreakBefore`, `numPr`, `jc`), and the last of those
 * is also what a paragraph's style can lend it, which is why the third paragraph here has an alignment it
 * never spelled and the fifth has a frame and a section break no field can hold. The fifth also carries a
 * `w:pPrChange` whose previous set names a shading and a `w:snapToGrid` its current set does not, because a
 * reviewer decides what the paragraph said before, not what this model can spell.
 */
export function makeParaPropsDocx(): Uint8Array {
  // The first paragraph's children follow `CT_PPrBase`'s own order, so a test can name the rank one of the
  // four modelled properties has to land at instead of guessing where the writer put it.
  const document = `<w:p><w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:widowControl w:val="0"/><w:suppressLineNumbers/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:tabs><w:tab w:val="right" w:pos="4320"/></w:tabs><w:bidi/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/><w:spacing w:before="120" w:after="60" w:line="360" w:lineRule="auto"/><w:ind w:left="420" w:right="120" w:firstLine="0"/><w:contextualSpacing/><w:textAlignment w:val="baseline"/><w:outlineLvl w:val="1"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/><w:b/><w:color w:val="FF0000"/><w:sz w:val="28"/><w:spacing w:val="20"/><w:highlight w:val="yellow"/></w:rPr></w:pPr><w:r><w:t>整段自述</w:t></w:r></w:p>
<w:p><w:pPr><w:keepNext/><w:rPr><w:sz w:val="24"/><w:del ${REV(7, "林晓")}/></w:rPr></w:pPr><w:r><w:t>断点被删</w:t></w:r></w:p><w:p><w:r><w:t>被并进来的下段</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>靠样式的对齐</w:t></w:r></w:p>
<w:p><w:pPr><w:keepNext/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:jc w:val="both"/><w:pPrChange ${REV(9, "顾峰")}><w:pPr><w:keepNext/><w:shd w:val="clear" w:color="auto" w:fill="E2EFDA"/><w:snapToGrid w:val="0"/><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>格式改动</w:t></w:r></w:p>
<w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:x="108" w:y="1"/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>带框与分节</w:t></w:r></w:p>`;
  return pack(document, STYLES);
}

/**
 * Paragraphs whose `w:ind` states one shape each, plus a style that can lend one. The spellings are measured,
 * not invented (`.scratch/ind-census.ts`, `ind-census3.ts`, `ind-order.ts` over the 86 documents WPS itself
 * wrote, which hold 29,889 `w:ind` elements):
 *
 * - 16,235 elements state `w:firstLine` together with `w:firstLineChars`, and both are `0` in 17,856 of the
 *   pairs' occurrences — "no first-line indent" is written out loud, not left unwritten.
 * - 5,097 state a positive `w:firstLine` with no character form at all; 1,147 state both positively (their
 *   ratio is always a font size: 180/210/240/264/300/320/440 twips per character); 236 state only
 *   `w:firstLineChars`; **88 state a positive `w:firstLine` next to `w:firstLineChars="0"`**, and 4 the other
 *   way round. A file never states a positive first line and a positive hanging in the same element (0 of
 *   29,889), and a positive `w:leftChars`, `w:rightChars` or `w:hangingChars` never appears without its twips
 *   value in the same element (0 of the 29,207 in the documents' own bodies), which is why only the first line
 *   needs a second, character-counted field.
 * - 107 elements state a negative `w:right` — all of them `-107` — so the right indent is not a non-negative
 *   number here.
 * - The attributes come out in the order left, leftChars, right, rightChars, firstLine, firstLineChars, with
 *   `w:hanging` last (`left,leftChars,right,rightChars,firstLine,firstLineChars` 123 times, `left,hanging` 118).
 *
 * Not observed: `w:start`/`w:end` (the bidi pair) in any of them, and a negative `w:firstLine` (0).
 */
export function makeIndentDocx(
  entries: readonly { readonly ind?: string; readonly pStyle?: string; readonly text: string }[],
  lentPPr = "",
): Uint8Array {
  const document =
    entries
      .map((e) => {
        const kids = [e.pStyle ? `<w:pStyle w:val="${e.pStyle}"/>` : "", e.ind ? `<w:ind ${e.ind}/>` : ""].join("");
        return `<w:p>${kids ? `<w:pPr>${kids}</w:pPr>` : ""}<w:r><w:t>${e.text}</w:t></w:r></w:p>`;
      })
      .join("\n") + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  const styles = lentPPr
    ? STYLES.replace("</w:styles>", `<w:style w:type="paragraph" w:styleId="Lent"><w:name w:val="借来的"/><w:pPr>${lentPPr}</w:pPr></w:style></w:styles>`)
    : STYLES;
  return pack(document, styles);
}

/**
 * A paragraph whose current set states no first-line indent while its `w:pPrChange` records one, in the shape
 * §17.13.5.19 gives it: the record is the last child of `w:pPr` and holds a whole set of its own.
 */
export function makeIndentRevisionDocx(): Uint8Array {
  const document =
    `<w:p><w:pPr><w:ind w:firstLine="0" w:firstLineChars="0"/>` +
    `<w:pPrChange w:id="12" w:author="顾峰" w:date="2026-03-01T08:00:00Z">` +
    `<w:pPr><w:ind w:firstLine="480" w:firstLineChars="200"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>缩进被改过</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  return pack(document, STYLES);
}
