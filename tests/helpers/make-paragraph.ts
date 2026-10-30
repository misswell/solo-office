import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

const pack = (document: string, styles: string, settings = ""): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${
            settings ? `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` : ""
          }</Types>`,
      ),
      // 件级的 `.rels` 只挂正文本体；`word/settings.xml` 挂在正文那一页的关系表上——授权 46 份份份如此
      // （`.scratch/settings-rel-census.py`：件级挂 settings 的 0 份）。挂错了地方，保存时就会发出第二张身份证。
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/>${
            settings ? `<Relationship Id="rId2" Type="${NS_R}/settings" Target="settings.xml"/>` : ""
          }</Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${document}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + styles),
      ...(settings ? { "word/settings.xml": strToU8(XML_DECL + `<w:settings xmlns:w="${NS_W}">${settings}</w:settings>`) } : {}),
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
  /** The `w:settings` children, for the case where the document states its own 默认制表位. */
  settings = "",
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
  return pack(document, styles, settings);
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

/**
 * Paragraphs whose `w:spacing` states one shape each, plus a style that can lend a set. The spellings and the
 * numbers are measured, not invented (`.scratch/spacing-census.ts` over the 25,801 `w:spacing` elements that
 * state a height or a gap in the 86 documents WPS itself wrote, counting their `word/document.xml`,
 * `word/styles.xml` and header/footer parts together — 38,427 `w:spacing` elements in all, of which 12,617 are
 * the character-spacing kind inside `w:rPr`):
 *
 * - `w:lineRule` is stated as `w:auto` 13,396 times, `w:exact` 5,812 and `w:atLeast` 991, and 5,602 elements
 *   state 段前段后 with no line at all. Every element that states a positive `w:line` states its rule too, so
 *   there is no "height with no unit" to fall back on; 6,802 of the 20,198 stated heights are a fixed value or
 *   a minimum, which are twips, not 240ths of a line.
 * - The heights themselves cluster: `auto` on 240 (12,353 — single spacing written out loud), 360 (536),
 *   480 (148), 180 (80); `exact` on 240 (2,968), 578 (1,518 — 28.9pt), 280 (910); `atLeast` on 363 (448)
 *   and 270 (330).
 * - `w:beforeLines`/`w:afterLines` are stated 2,572/2,581 times but positive only 128/135 of those (2,444/2,446
 *   state `0`), and 120 of the positive pairs imply a "line" of 100 twips — a grid pitch, not a line height —
 *   so this fixture keeps the twips spellings only. `w:beforeAutospacing`/`w:afterAutospacing` state `1` 66/54
 *   times and `0` 4,850/4,852 times. Both pairs travel as the file's own bytes either way.
 *
 * `.scratch/spacing-lent.ts` counted the other half of it: 5,110 paragraphs state no `w:line` while the style
 * named by their `w:pStyle` does (4,847 `auto`, 247 `exact`, 16 `atLeast`), which is what the `Lent` style
 * here is for.
 *
 * `.scratch/wps-snap-where.mjs` counts **where** the files write the grid switch: of the corpus's 3,466
 * `<w:snapToGrid>`, 3,394 are children of a `w:pPr` — the slot §17.3.1.87 defines — and 72 are children of a
 * `w:rPr` (19 in the paragraph mark's, 53 in a run's). `markRPr` and `runRPr` spell those two shapes so the read
 * can be pinned to the slot; a section's `w:docGrid` comes in through `sectKids`.
 */
export function makeSpacingDocx(
  entries: readonly {
    readonly spacing?: string;
    readonly pStyle?: string;
    readonly snapToGrid?: string;
    /** `w:wordWrap`（「允许西文在单词中间换行」）同一条拼法：空串是裸写=说了允许，`w:val="0"` 是明着不许。 */
    readonly wordWrap?: string;
    /**
     * 「换行和分页」那四枚（`w:keepNext`/`w:keepLines`/`w:pageBreakBefore`/`w:widowControl`）的原样字节。
     * 它们按 `PPR_ORDER` 排在 `w:pStyle` 之后、`w:wordWrap` 之前，顺序由调用方拼。
     */
    readonly keep?: string;
    /** Extra children appended to this paragraph's `w:pPr`, for the shapes that are not paragraph properties. */
    readonly markRPr?: string;
    /** Extra children of this paragraph's one run's `w:rPr`. */
    readonly runRPr?: string;
    readonly text: string;
  }[],
  lentPPr = "",
  sectKids = "",
): Uint8Array {
  const document =
    entries
      .map((e) => {
        const kids = [
          e.pStyle ? `<w:pStyle w:val="${e.pStyle}"/>` : "",
          e.keep ?? "",
          // `w:wordWrap` 排在 `w:snapToGrid` 之前，与 `PPR_ORDER` 同一条（`w:kinsoku` 之后、`w:overflowPunct` 之前）。
          // 件里说了这一枚的 200 段枚枚裸写（授权口径 `.scratch/cjk-typo-audience.py`），`w:val="0"` 是本机自己造的形状。
          e.wordWrap === undefined ? "" : `<w:wordWrap${e.wordWrap ? ` ${e.wordWrap}` : ""}/>`,
          // `w:snapToGrid` is written exactly as the file spells it: `w:val="0"` is the off, and an empty string
          // here is the bare `<w:snapToGrid/>`, which is the switch stated *on* (§17.3.1.87 defaults to on).
          e.snapToGrid === undefined ? "" : `<w:snapToGrid${e.snapToGrid ? ` ${e.snapToGrid}` : ""}/>`,
          e.spacing ? `<w:spacing ${e.spacing}/>` : "",
          e.markRPr ?? "",
        ].join("");
        const run = `<w:r>${e.runRPr ? `<w:rPr>${e.runRPr}</w:rPr>` : ""}<w:t>${e.text}</w:t></w:r>`;
        return `<w:p>${kids ? `<w:pPr>${kids}</w:pPr>` : ""}${run}</w:p>`;
      })
      .join("\n") +
    // `sectKids` goes after `w:pgMar`, which is where §17.3 puts `w:docGrid`: a `w:pgMar` written after a
    // `w:docGrid` is a file Word refuses to open (the order list in `docx.ts` keeps the same place).
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>${sectKids}</w:sectPr>`;
  const styles = lentPPr
    ? STYLES.replace("</w:styles>", `<w:style w:type="paragraph" w:styleId="Lent"><w:name w:val="借来的"/><w:pPr>${lentPPr}</w:pPr></w:style></w:styles>`)
    : STYLES;
  return pack(document, styles);
}

/**
 * Two paragraphs whose `w:pPrChange` records a different line-height unit than the set it sits in, in the shape
 * §17.13.5.19 gives it: the record is the last child of `w:pPr` and holds a whole set of its own. The first
 * states a 1.5x line now and 固定值 28.9pt before; the second states 固定值 16pt now, borrows a 最小值 from its
 * style, and its record states neither — so refusing it has to hand the borrowed pair back, not leave the height
 * without the unit it is counted in.
 */
export function makeSpacingRevisionDocx(): Uint8Array {
  const document =
    `<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/>` +
    `<w:pPrChange w:id="14" w:author="顾峰" w:date="2026-03-01T08:00:00Z">` +
    `<w:pPr><w:spacing w:line="578" w:lineRule="exact"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>行距被改过</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:pStyle w:val="Lent"/><w:spacing w:line="240" w:lineRule="exact"/>` +
    `<w:pPrChange w:id="15" w:author="林晓" w:date="2026-03-01T08:00:00Z">` +
    `<w:pPr><w:pStyle w:val="Lent"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>行距靠样式也被改过</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  const styles = STYLES.replace(
    "</w:styles>",
    `<w:style w:type="paragraph" w:styleId="Lent"><w:name w:val="借来的"/><w:pPr><w:spacing w:line="363" w:lineRule="atLeast"/></w:pPr></w:style></w:styles>`,
  );
  return pack(document, styles);
}

/**
 * Two paragraphs whose `w:pPrChange` disagrees with the set it sits in about 「对齐到网格」, in the shape
 * §17.13.5.19 gives the record: the last child of `w:pPr`, holding a whole set of its own. The first states a
 * bare `<w:snapToGrid/>` now while its record states nothing about the grid, so refusing it has to take the
 * paragraph's own statement back rather than leave the switch claimed. The second states nothing now while its
 * record states `w:val="0"`, so refusing it has to hand that statement back — bytes and field both.
 */
export function makeGridRevisionDocx(): Uint8Array {
  const document =
    `<w:p><w:pPr><w:snapToGrid/><w:spacing w:line="300" w:lineRule="auto"/>` +
    `<w:pPrChange ${REV(21, "顾峰")}><w:pPr><w:spacing w:line="300" w:lineRule="auto"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>网格是刚加上的</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:spacing w:line="300" w:lineRule="auto"/>` +
    `<w:pPrChange ${REV(22, "林晓")}><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:line="300" w:lineRule="auto"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>网格是刚被删掉的</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>`;
  return pack(document, STYLES);
}

/**
 * The same two shapes for 「允许西文在单词中间换行」（§17.4.77 `w:wordWrap`）: one states a bare
 * `<w:wordWrap/>` now while its record says nothing, one says nothing now while its record states
 * `w:val="0"`. Refusing either has to move the field *and* the bytes—the switch lives outside
 * `PARA_STYLE_KEYS`, so it travels only in `w:pPr`'s slots.
 */
export function makeWrapRevisionDocx(): Uint8Array {
  const document =
    `<w:p><w:pPr><w:wordWrap/><w:spacing w:line="300" w:lineRule="auto"/>` +
    `<w:pPrChange ${REV(31, "顾峰")}><w:pPr><w:spacing w:line="300" w:lineRule="auto"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>换行是刚加上的</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:spacing w:line="300" w:lineRule="auto"/>` +
    `<w:pPrChange ${REV(32, "林晓")}><w:pPr><w:wordWrap w:val="0"/><w:spacing w:line="300" w:lineRule="auto"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>换行是刚被删掉的</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  return pack(document, STYLES);
}

/**
 * The same two shapes for 「换行和分页」's keeps: one now says `<w:keepLines/>` while its record says nothing,
 * one now says nothing while its record says `w:widowControl w:val="0"`. Neither key is in `PARA_STYLE_KEYS`'
 * write half, so refusing has to clear the field *and* drop the slot.
 */
export function makeKeepRevisionDocx(): Uint8Array {
  const document =
    `<w:p><w:pPr><w:keepLines/><w:spacing w:line="300" w:lineRule="auto"/>` +
    `<w:pPrChange ${REV(41, "顾峰")}><w:pPr><w:spacing w:line="300" w:lineRule="auto"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>不分页是刚加上的</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:spacing w:line="300" w:lineRule="auto"/>` +
    `<w:pPrChange ${REV(42, "林晓")}><w:pPr><w:widowControl w:val="0"/><w:spacing w:line="300" w:lineRule="auto"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>孤行是刚被删掉的</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  return pack(document, STYLES);
}

/**
 * 「显示原文」那一档纸面穿的是记录里的 `prev`，所以读数的框也得说同一句（#224）：第一段现值说右对齐、
 * 段前 480 缇、两倍行距，记录说改前是居中、段前 240 缇、一点五倍行距；第二段没有任何记录，两档同脸，
 * 用来证明「原文」不是把每一段都翻个面。
 */
export function makeOriginalReadDocx(): Uint8Array {
  const document =
    `<w:p><w:pPr><w:jc w:val="right"/><w:spacing w:before="480" w:line="480" w:lineRule="auto"/>` +
    `<w:pPrChange ${REV(77, "顾峰")}><w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrChange></w:pPr>` +
    `<w:r><w:t>被改过段距的一段</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:before="240" w:line="360" w:lineRule="auto"/></w:pPr>` +
    `<w:r><w:t>没被改过的一段</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  return pack(document, STYLES);
}

/** A style as `word/styles.xml` spells it, for the tests about how far up `w:basedOn` the read walks. */
export interface ChainStyle {
  id: string;
  name?: string;
  basedOn?: string;
  /** `w:next`: the style a paragraph takes when Enter splits the one naming this one (§17.7.2.24). */
  next?: string;
  /** `w:default="1"`: the style a paragraph that names none falls back on. */
  isDefault?: boolean;
  /** The style's `w:pPr` children, written the way the files write them. */
  pPr?: string;
  /** The style's `w:rPr` children — the character properties §17.7.2 hands down to the runs inside it. */
  rPr?: string;
  /** `w:type`, needed only for the case that must not leak: a character style that is the default. */
  type?: "paragraph" | "character";
}

/**
 * A styles part of the given styles and one paragraph per entry, for the shape §17.7.2 puts the model's read
 * against. The census is `.scratch/basedon-chain.ts` over the 86 WPS-written documents:
 *
 * - 1,337 paragraph styles, 1,182 of them carrying `w:basedOn` — every one of those references a style the part
 *   declares, so no chain in these files dangles or circles. Chain depth runs 1 to 4 (155 / 1,011 / 169 / 2
 *   styles).
 * - 879 paragraphs ask for a property only an ancestor states it: 501 line heights (every one `w:exact`, at hop
 *   2 in 453 of them and hop 3 in 48), 327 alignments, 254 indents, 115 段前段后 gaps.
 * - 38,475 of the 46,258 paragraphs name no style, and 3,147 of those take a property from the default
 *   paragraph style (1,956 alignments, 613 固定值 heights, 478 of each gap). All 86 files declare a default
 *   paragraph style, and 47 of them state a line height there.
 * - `w:docDefaults/w:pPrDefault` states no paragraph property in any of the 86, which is why the fixture's
 *   `w:docDefaults` carries only the run font and size the files carry. The font it carries is the shape 67 of
 *   the 86 write there — one element naming both faces, and naming a different one for `w:eastAsia`
 *   (`.scratch/run-font-eastasia.ts`); the other 19 write no `w:rFonts` at that level at all.
 *
 * The style ids are readable rather than the `a1`/`af0` WPS writes, since a test about which style in the chain
 * a value came from has to name them; every attribute spelling is the files' own.
 */
export function makeChainDocx(
  styles: readonly ChainStyle[],
  paras: readonly { readonly pStyle?: string; readonly pPr?: string; readonly text: string; readonly rPr?: string; readonly rStyle?: string }[],
  /** The floor's `w:rPr` children. The default is the shape 67 of the 86 files write there. */
  defaults = '<w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/><w:sz w:val="22"/>',
): Uint8Array {
  const sx = `<w:styles xmlns:w="${NS_W}">
<w:docDefaults><w:rPrDefault><w:rPr>${defaults}</w:rPr></w:rPrDefault></w:docDefaults>
${styles
  .map((s) => {
    const attrs = [`w:type="${s.type ?? "paragraph"}"`, s.isDefault ? 'w:default="1"' : "", `w:styleId="${s.id}"`]
      .filter(Boolean)
      .join(" ");
    return `<w:style ${attrs}><w:name w:val="${s.name ?? s.id}"/>${s.basedOn ? `<w:basedOn w:val="${s.basedOn}"/>` : ""}${s.next ? `<w:next w:val="${s.next}"/>` : ""}${
      s.type === "character" ? "" : "<w:qFormat/>"
    }${s.pPr ? `<w:pPr>${s.pPr}</w:pPr>` : ""}${s.rPr ? `<w:rPr>${s.rPr}</w:rPr>` : ""}</w:style>`;
  })
  .join("\n")}
</w:styles>`;
  const document =
    paras
      .map((e) => {
        const kids = [e.pStyle ? `<w:pStyle w:val="${e.pStyle}"/>` : "", e.pPr ?? ""].join("");
        const own = [e.rStyle ? `<w:rStyle w:val="${e.rStyle}"/>` : "", e.rPr ?? ""].join("");
        return `<w:p>${kids ? `<w:pPr>${kids}</w:pPr>` : ""}<w:r>${own ? `<w:rPr>${own}</w:rPr>` : ""}<w:t>${e.text}</w:t></w:r></w:p>`;
      })
      .join("\n") + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  return pack(document, sx);
}
