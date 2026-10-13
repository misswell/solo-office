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
