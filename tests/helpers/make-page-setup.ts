import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/**
 * A one-paragraph document whose section is exactly the bytes the caller spells. Page setup is a question
 * about `w:sectPr`, so the fixture hands that element over whole rather than assembling attributes: a test
 * that wants no `w:pgMar` says so by not writing one, and a test that wants the corpus's `w:cols` and
 * `w:docGrid` standing behind it writes those too. `at` decides which section owns the opening page — the
 * body's own child, a paragraph's section break, or neither.
 */
export function makePageSetupDocx(section: string, at: "body" | "paragraph" | "none" = "body"): Uint8Array {
  const body =
    at === "none"
      ? `<w:p><w:r><w:t>正文</w:t></w:r></w:p>`
      : at === "body"
        ? `<w:p><w:r><w:t>正文</w:t></w:r></w:p>${section}`
        : `<w:p><w:pPr>${section}</w:pPr><w:r><w:t>正文</w:t></w:r></w:p>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}</w:body></w:document>`),
    },
    { level: 6 },
  );
}

/** The opening section as the 86 files this machine's WPS wrote spell it: all seven distances, then columns and grid. */
export const WPS_SECTION =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="283"/><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>';

/** A landscape opening section, spelled the only way the corpus spells landscape: long edge first and `w:orient`. */
export const WPS_LANDSCAPE_SECTION =
  '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>';

/**
 * The same A4 page with the body margins it states, and nothing about the paper-edge distances or the binding
 * width. A `w:pgMar` may leave all three out (§17.3.2.13 makes them optional), and this is the shape where a
 * dialog that prefilled Word's own defaults would put numbers in the file that neither the file nor the user
 * named.
 */
export const WPS_SECTION_NO_DISTANCES = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

/**
 * The near-A4 sheet four of the 115 corpus sections spell — 11900×16840, six twips off A4's width and two off
 * its height. It is on no preset list, so it is the case where a dialog that offered sizes by name could pull
 * the page to the named sheet instead of keeping the numbers the file wrote.
 */
export const WPS_NEAR_A4_SECTION =
  '<w:sectPr><w:pgSz w:w="11900" w:h="16840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>';

/**
 * A section that says nothing about the page at all: no `w:pgSz`, no `w:pgMar`, and a column setting that has
 * to survive a change untouched. This is the shape a minimal producer writes, and the one where a save owes
 * both elements rather than neither.
 */
export const BARE_SECTION = '<w:sectPr><w:cols w:space="720"/></w:sectPr>';

/**
 * A section carrying a page border, in the shape the 24 of the 115 corpus sections spell one: `w:pgBorders`
 * after `w:pgMar` and before `w:cols`, its sides written `w:top`, `w:left`, `w:bottom`, `w:right`, and no
 * `w:offsetFrom` on the container — none of those 24 spelled it. What was observed in the corpus is that shape
 * and how often it appears; the stroke tokens, widths, gaps and colours below are this fixture's own choice
 * from what §17.3.2.10 admits, because a probe over the 24 found nothing to copy: all four sides of every one
 * of them state `none`, so no real file in this corpus draws a line.
 */
export const WPS_PAGE_BORDER_SECTION =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="283"/><w:pgBorders><w:top w:val="single" w:sz="8" w:space="6" w:color="auto"/><w:left w:val="dashed" w:sz="12" w:space="0" w:color="FF0000"/><w:bottom w:val="double" w:sz="4" w:space="12"/><w:right w:val="none" w:sz="4" w:space="0"/></w:pgBorders><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>';

/**
 * The page border as WPS' own 24 files write it: four sides, each stating `w:val="none"` plus a `w:sz` and a
 * `w:space`, and no `w:color`. This is a border that was cleared rather than drawn, and it is the only
 * page-border shape real files in this corpus show — so the obligation it pins is that the element is read, kept,
 * and above all draws **nothing** on the page. Observed: the four sides, the three attributes each carries, and
 * the `none` token. Not observed: the two numbers, which are this fixture's — a probe counts attribute names and
 * the closed stroke enumeration, never a file's values.
 */
export const WPS_CLEARED_BORDER_SECTION =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="283"/><w:pgBorders><w:top w:val="none" w:sz="0" w:space="0"/><w:left w:val="none" w:sz="0" w:space="0"/><w:bottom w:val="none" w:sz="0" w:space="0"/><w:right w:val="none" w:sz="0" w:space="0"/></w:pgBorders><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>';

/**
 * One section carrying every `w:sectPr` child this model has no field for that WPS' own files actually write.
 * Measured over the authorized 46 份 / 58 节 (`.scratch/sect-children-audience.py`, 件数按全路径计): nine tag names
 * account for **all** of their children — `w:pgSz`, `w:pgMar`, `w:cols`, `w:docGrid`, `w:footerReference`(32 枚),
 * `w:pgNumType`(20), `w:headerReference`(6), `w:footnotePr`(4), `w:titlePg`(2) — while `w:paperSrc`, `w:lnNumType`,
 * `w:formProt`, `w:vAlign`, `w:textDirection`, `w:bidi`, `w:endnotePr`, `w:printerSettings` **and `w:pgBorders`**
 * appear in none. Shapes differ from what this fixture used to claim: `w:pgNumType` states `w:fmt` too (fmt 4 枚,
 * fmt+start 14, start only 2 — none omits both) but never a non-default value (`w:start` is always `1`, `w:fmt`
 * always `decimal`); `w:titlePg` is self-closing with no attribute; 3 of the 58 `w:cols` state `w:space` **without**
 * `w:num`; every `w:docGrid` states `w:linePitch`, while `w:charSpace` rides 55 and `w:type` 55 of them.
 * `w:footnotePr`
 * holds one child, `<w:numFmt w:val="decimal"/>`, and is the only container besides the border set below — that is
 * why it belongs in a survival fixture: a child that carries children can die differently.
 *
 * The order is measured too, not borrowed from §17.3: the real files put band references and `w:footnotePr` *before*
 * `w:pgSz` (e.g. `headerReference > footerReference > footnotePr > pgSz > pgMar > pgNumType > cols > docGrid`), with
 * `w:pgNumType` between `w:pgMar` and `w:cols` and `w:titlePg` between `w:cols` and `w:docGrid`.
 *
 * Not observed: `w:pgBorders` has **0** audience in the authorized 46 份 (the border set this fixture and the
 * page-border support #67 carry comes from a cohort whose counts are void), so the stroke values below and their
 * pairing with `w:pgNumType` are this fixture's own choice from what §17.3 admits — a shape the walk must survive,
 * not a claim about what WPS writes. Only structure was read from the authorized files (tag names, attribute names
 * and values, counts, order); no document text is quoted here.
 */
export const WPS_SECTION_UNREAD_CHILDREN =
  '<w:sectPr><w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/><w:pgBorders><w:top w:val="none" w:sz="0" w:space="0"/><w:left w:val="none" w:sz="0" w:space="0"/><w:bottom w:val="none" w:sz="0" w:space="0"/><w:right w:val="none" w:sz="0" w:space="0"/></w:pgBorders><w:pgNumType w:start="1"/><w:cols w:space="425" w:num="1"/><w:titlePg/><w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/></w:sectPr>';
