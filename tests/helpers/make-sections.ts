import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

const STYLES = `<w:styles xmlns:w="${NS_W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

/**
 * A `w:sectPr` with A4 portrait and the margins the corpus writes: the geometry every part of the fixture
 * declares. It is the same blob for the break paragraphs and the body, because the fixtures here are about
 * **which paragraphs the boundaries hang off**, not about per-section paper (that slice has not landed).
 */
export const SECT_PR = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;

/** One paragraph. `pPr` is the caller's own `w:pPr` children; `text` omitted means the paragraph has no run at all. */
export interface SectionPara {
  pPr?: string;
  text?: string;
}

/**
 * A document of the given blocks over `SECT_PR`, for the tests about section boundaries: a paragraph whose
 * `w:pPr` carries a `w:sectPr` ends the section it sits in (§17.3.1.33: `w:sectPr` is a `CT_PPr` child), and the
 * body's own direct `w:sectPr` child closes the last section — it never appears among the blocks. The caller
 * passes the break paragraphs as `SectionPara`s with `pPr: SECT_PR` so the boundary is in the blocks, exactly
 * where the model has to count it from.
 */
export function makeSectionsDocx(blocks: readonly SectionPara[]): Uint8Array {
  const body =
    blocks
      .map((b) => `<w:p>${b.pPr ? `<w:pPr>${b.pPr}</w:pPr>` : ""}${b.text === undefined ? "" : `<w:r><w:t>${b.text}</w:t></w:r>`}</w:p>`)
      .join("\n") + SECT_PR;
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
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + STYLES),
    },
    { level: 6 },
  );
}
