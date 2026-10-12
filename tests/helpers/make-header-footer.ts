import { strToU8, zipSync } from "fflate";

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_HDR = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const CT_FTR = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

/**
 * The section's reference to the header, in the order the local WPS corpus writes it — `r:id` first on
 * every one of its 59 header and 57 footer references, self-closed, and with the id attribute prefixed.
 */
export const REFERENCE_ID_FIRST = '<w:headerReference r:id="rId5" w:type="default"/>';

/** Word's own order, which none of the 116 references in the corpus uses. Both have to read. */
export const REFERENCE_TYPE_FIRST = '<w:headerReference w:type="default" r:id="rId5"/>';

/** A section that binds only a first-page header. There is still a part here, and it is this one. */
export const REFERENCE_FIRST_PAGE_ONLY = '<w:headerReference r:id="rId5" w:type="first"/>';

/**
 * A one-paragraph body whose section points at one header or footer part. The caller owns that part whole,
 * declaration included, so what a save hands back can be compared to it byte for byte.
 */
function packHf(kind: "header" | "footer", part: string, reference: string): Uint8Array {
  const name = `word/${kind}1.xml`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/${name}" ContentType="${kind === "header" ? CT_HDR : CT_FTR}"/></Types>`,
      ),
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr>${reference}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/${kind}" Target="${kind}1.xml"/></Relationships>`,
      ),
      [name]: strToU8(part),
    },
    { level: 6 },
  );
}

export function makeHeaderFooterDocx(headerPart: string, reference: string = REFERENCE_ID_FIRST): Uint8Array {
  return packHf("header", headerPart, reference);
}

/** The same file with the part bound as the section's footer, which is where a page number usually lives. */
export function makeFooterDocx(footerPart: string): Uint8Array {
  return packHf("footer", footerPart, '<w:footerReference r:id="rId5" w:type="default"/>');
}

/**
 * A header in the shape the local WPS corpus writes: parts declare around 17 namespaces and every one of
 * them carries `mc:Ignorable`, and of the 86 parts a text edit can reach, 85 give their paragraph a
 * `w:pStyle` and 18 a `w:jc`; 2 hold a literal `w:tab`. So the first paragraph here has all of it — the
 * tab stop `w:pPr/w:tabs` defines, which is not text, beside the tab the run holds, which is — and the
 * second is the plain one an edit elsewhere must leave alone.
 *
 * The ids, the bookmark name and the texts are ours; the shapes are transcriptions.
 */
export const STYLED_HEADER =
  `${XML_DECL}<w:hdr xmlns:w="${NS_W}" xmlns:mc="${NS_MC}" xmlns:w14="${NS_W14}" xmlns:r="${NS_R}" mc:Ignorable="w14">` +
  `<w:p w14:paraId="11111111" w14:textId="22222222">` +
  `<w:pPr><w:pStyle w:val="Header"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9026"/></w:tabs><w:jc w:val="right"/><w:rPr><w:sz w:val="18"/></w:rPr></w:pPr>` +
  `<w:bookmarkStart w:id="7" w:name="HeaderMark"/>` +
  `<w:r><w:rPr><w:b/><w:sz w:val="18"/></w:rPr><w:t>机密</w:t><w:tab/><w:t>第一段</w:t></w:r>` +
  `<w:bookmarkEnd w:id="7"/>` +
  `</w:p>` +
  `<w:p w14:paraId="33333333"><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>第二段</w:t></w:r></w:p>` +
  `</w:hdr>`;

/** The first paragraph of {@link STYLED_HEADER}, whole — what an edit elsewhere must hand back byte for byte. */
export const STYLED_FIRST_PARA = STYLED_HEADER.slice(
  STYLED_HEADER.indexOf('<w:p w14:paraId="11111111"'),
  STYLED_HEADER.indexOf('<w:p w14:paraId="33333333"'),
);

export function makeStyledHeaderDocx(): Uint8Array {
  return makeHeaderFooterDocx(STYLED_HEADER);
}

/** The header the opening section binds, and the one the closing section binds. Different parts, different text. */
export const FIRST_SECTION_HEADER =
  `${XML_DECL}<w:hdr xmlns:w="${NS_W}"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:t>开篇页眉</w:t></w:r></w:p></w:hdr>`;
export const LAST_SECTION_HEADER =
  `${XML_DECL}<w:hdr xmlns:w="${NS_W}"><w:p><w:r><w:t>结尾页眉</w:t></w:r></w:p></w:hdr>`;

/**
 * A document with two sections: the break hanging off the first paragraph ends the opening one, and the
 * body's own `sectPr` is the closing one. Each carries its own header and its own page box — portrait first,
 * landscape last — and each `sectPr` wears a distinguishable `w:rsidR`, so a test can say which section an
 * edited file put a reference in.
 *
 * The corpus is the reason this shape matters: 18 of the 86 files WPS wrote here have more than one section,
 * and in 4 of them the header is cited only by a later one, in 5 the footer is.
 */
export function makeTwoSectionDocx(o: { firstRef?: string; lastRef?: string } = {}): Uint8Array {
  const firstRef = o.firstRef === undefined ? '<w:headerReference r:id="rId5" w:type="default"/>' : o.firstRef;
  const lastRef = o.lastRef === undefined ? '<w:headerReference r:id="rId6" w:type="default"/>' : o.lastRef;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/header1.xml" ContentType="${CT_HDR}"/><Override PartName="/word/header2.xml" ContentType="${CT_HDR}"/></Types>`,
      ),
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
          `<w:p><w:pPr><w:sectPr w:rsidR="00A1">${firstRef}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr><w:r><w:t>第一节</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t>第二节</w:t></w:r></w:p>` +
          `<w:sectPr w:rsidR="00B2">${lastRef}<w:pgSz w:w="16838" w:h="11906"/></w:sectPr>` +
          `</w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/header" Target="header1.xml"/><Relationship Id="rId6" Type="${NS_R}/header" Target="header2.xml"/></Relationships>`,
      ),
      "word/header1.xml": strToU8(FIRST_SECTION_HEADER),
      "word/header2.xml": strToU8(LAST_SECTION_HEADER),
    },
    { level: 6 },
  );
}

