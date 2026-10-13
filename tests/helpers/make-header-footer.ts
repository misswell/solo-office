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

/** A footer part holding whatever paragraphs the test names. */
export const footerPart = (body: string): string => `${XML_DECL}<w:ftr xmlns:w="${NS_W}" xmlns:r="${NS_R}">${body}</w:ftr>`;

/** A field marker run, spelled the way the files spell it. */
const fldChar = (stage: "begin" | "separate" | "end"): string => `<w:r><w:fldChar w:fldCharType="${stage}"/></w:r>`;

/** The instruction run, with the spaces at its edges kept by `xml:space="preserve"` as the corpus does. */
const instr = (text: string): string => `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;

/**
 * A page field as the local WPS corpus writes one — and as `insertPageNumber` writes one: `begin`, the
 * instruction, `separate`, the number it last showed, `end`. Of that corpus's 21 fields that sit at run level
 * in a header or footer part, 16 spell the instruction exactly this way, 2 spell `PAGE` with no switch at
 * all, and all 21 cache a digit.
 */
export const pageFieldRuns = (instruction = " PAGE  \\* MERGEFORMAT ", cached = "1"): string =>
  `${fldChar("begin")}${instr(instruction)}${fldChar("separate")}<w:r><w:t>${cached}</w:t></w:r>${fldChar("end")}`;

/** 「第 」, a page field, 「 页」 — the shape those 4 corpus paragraphs that are not field-only take. */
export const PAGE_FOOTER = footerPart(
  `<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t xml:space="preserve">第 </w:t></w:r>${pageFieldRuns()}<w:r><w:t xml:space="preserve"> 页</w:t></w:r></w:p><w:p><w:r><w:t>普通一行</w:t></w:r></w:p>`,
);

/** A footer whose only content is the field, which is 17 of the 21 in that corpus. */
export const PAGE_ONLY_FOOTER = footerPart(`<w:p>${pageFieldRuns()}</w:p>`);

/** The field alone in its paragraph, with the paragraph's formatting on the run that carries the digit. */
export const PAGE_ONLY_FOOTER_STYLED = footerPart(
  `<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
);

/**
 * A document with two sections: the break hanging off the first paragraph ends the opening one, and the
 * body's own `sectPr` is the closing one. Each carries its own header and its own page box — portrait first,
 * landscape last — and each `sectPr` wears a distinguishable `w:rsidR`, so a test can say which section an
 * edited file put a reference in.
 *
 * The corpus is the reason this shape matters: 18 of the 86 files WPS wrote here have more than one section,
 * and in 4 of them the header is cited only by a later one, in 5 the footer is.
 */
export function makeTwoSectionDocx(o: { firstRef?: string; lastRef?: string; firstLandscape?: boolean; firstMarginTwips?: number } = {}): Uint8Array {
  const firstRef = o.firstRef === undefined ? '<w:headerReference r:id="rId5" w:type="default"/>' : o.firstRef;
  const lastRef = o.lastRef === undefined ? '<w:headerReference r:id="rId6" w:type="default"/>' : o.lastRef;
  // The opening section's page box either way round, so a check can tell "the band follows the section that
  // owns it" from "the band matches a portrait page by coincidence" — the stylesheet's own fallback is 794px.
  const firstBox = o.firstLandscape ? '<w:pgSz w:w="16838" w:h="11906"/>' : '<w:pgSz w:w="11906" w:h="16838"/>';
  // 24 of the 24 corpus files that show a band give their opening section margins other than the 1440 default
  // (15 of them 1800), so the fixture wears that shape too — otherwise a band drawn at the stylesheet's fixed
  // 96px padding lines up by accident and the check proves nothing.
  const firstMar = o.firstMarginTwips ?? 1440;
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
          `<w:p><w:pPr><w:sectPr w:rsidR="00A1">${firstRef}${firstBox}<w:pgMar w:top="1440" w:right="${firstMar}" w:bottom="1440" w:left="${firstMar}"/></w:sectPr></w:pPr><w:r><w:t>第一节</w:t></w:r></w:p>` +
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


const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_V = "urn:schemas-microsoft-com:vml";

/** The anchored picture the corpus writes: `mc:Choice` carries the DrawingML, `mc:Fallback` the VML. */
const anchored = (inside: string, fallbackInside = `<v:shape id="Logo"/>`): string =>
  `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1000000" cy="500000"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${inside}` +
  `<wp:docPr id="7" name="信头徽标"/><a:graphic><a:graphicData uri="${NS_PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="徽标"/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>` +
  `</wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict>${fallbackInside}</w:pict></mc:Fallback></mc:AlternateContent>`;

/** The picture alone: no `w:t` anywhere in it, which is what lets the band carry it as one piece. */
export const LOGO_RUN = `<w:r><w:rPr><w:noProof/></w:rPr>${anchored("<wp:wrapNone/>")}</w:r>`;

/** The paragraph inside the corpus's page box: two words the page shows with the field between them. */
const BOX_PARA =
  `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>第 </w:t></w:r>` +
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
  `<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t> 页</w:t></w:r></w:p>`;

/** A VML copy of the same box, which is what 16 of the 16 corpus runs write: the words live in both branches. */
const vmlBox = (para: string): string =>
  `<v:shape id="PageBox"><v:textbox inset="0,0,0,0"><w:txbxContent>${para}</w:txbxContent></v:textbox></v:shape>`;

/**
 * The shape all 16 of these runs in the local corpus have: a floating box whose one inner paragraph holds the
 * page number, written twice over — `w:txbxContent` inside the Choice's anchor and again inside the Fallback's
 * VML shape — with the same words in both.
 */
export const PAGE_BOX_RUN =
  `<w:r><w:rPr><w:noProof/></w:rPr>${anchored(`<wp:wrapNone/><w:txbxContent>${BOX_PARA}</w:txbxContent>`, vmlBox(BOX_PARA))}</w:r>`;

/** The two copies disagree about what the box says: no single edit could write them both. */
export const PAGE_BOX_DISAGREEING_RUN =
  `<w:r><w:rPr><w:noProof/></w:rPr>${anchored(
    `<wp:wrapNone/><w:txbxContent>${BOX_PARA}</w:txbxContent>`,
    vmlBox(BOX_PARA.replace("第 ", "共 ").replace(" 页", " 张")),
  )}</w:r>`;

/** A box whose text is split over two paragraphs: the band has no paragraph to hang them on. */
export const PAGE_BOX_TWO_PARA_RUN =
  `<w:r><w:rPr><w:noProof/></w:rPr>${anchored(
    `<wp:wrapNone/><w:txbxContent>${BOX_PARA}<w:p><w:r><w:t>第二段落</w:t></w:r></w:p></w:txbxContent>`,
    `<v:shape id="PageBox"/>`,
  )}</w:r>`;

/** A footer whose first paragraph is one of those runs and whose second is text the user can type in. */
export function makeLogoFooterDocx(logo: string): Uint8Array {
  const part =
    `${XML_DECL}<w:ftr xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}" xmlns:v="${NS_V}" mc:Ignorable="wp">` +
    `<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>${logo}</w:p>` +
    `<w:p><w:r><w:t>联系文字</w:t></w:r></w:p></w:ftr>`;
  return makeFooterDocx(part);
}

/** A footer whose one paragraph holds typed words *beside* the box — a shape this model does not claim. */
export function makeBoxWithWordsFooterDocx(box: string): Uint8Array {
  const part =
    `${XML_DECL}<w:ftr xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}" xmlns:v="${NS_V}" mc:Ignorable="wp">` +
    `<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>页脚左边</w:t></w:r>${box}</w:p></w:ftr>`;
  return makeFooterDocx(part);
}

/**
 * A file with one header band and one footer band and a `w:pgMar` spelled exactly as the caller writes it, so
 * the three distances beyond the body margins (`w:header` / `w:footer` / `w:gutter`) can be handed over — or
 * withheld. WPS states all seven on every one of the 115 `w:pgMar` in the 86 real files, so the stated shape
 * is the ordinary one; the withheld shape is what a hand-made file looks like, and it must not be answered
 * with numbers nobody wrote.
 *
 * `at` chooses which `sectPr` owns the section: the body's, or the first paragraph's break. Both are places a
 * distance can be stated, and they travel by different routes on a save — the body's as whole `sectPr` bytes,
 * the paragraph's inside its `w:pPr`.
 */
export function makePageDistanceDocx(mar: string, at: "body" | "paragraph" = "body"): Uint8Array {
  const section = `<w:sectPr><w:headerReference r:id="rId5" w:type="default"/><w:footerReference r:id="rId6" w:type="default"/>` +
    `<w:pgSz w:w="11906" w:h="16838"/><w:pgMar ${mar}/></w:sectPr>`;
  const body =
    at === "body"
      ? `<w:p><w:r><w:t>正文</w:t></w:r></w:p>${section}`
      : `<w:p><w:pPr>${section.replace("<w:sectPr>", '<w:sectPr w:rsidR="00A1">')}</w:pPr><w:r><w:t>正文</w:t></w:r></w:p>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/header1.xml" ContentType="${CT_HDR}"/><Override PartName="/word/footer1.xml" ContentType="${CT_FTR}"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(
        XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}</w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/header" Target="header1.xml"/><Relationship Id="rId6" Type="${NS_R}/footer" Target="footer1.xml"/></Relationships>`,
      ),
      "word/header1.xml": strToU8(`${XML_DECL}<w:hdr xmlns:w="${NS_W}"><w:p><w:r><w:t>页眉文字</w:t></w:r></w:p></w:hdr>`),
      "word/footer1.xml": strToU8(`${XML_DECL}<w:ftr xmlns:w="${NS_W}"><w:p><w:r><w:t>页脚文字</w:t></w:r></w:p></w:ftr>`),
    },
    { level: 6 },
  );
}

/** The same file with all seven attributes stated, the way the local WPS corpus writes it. */
export function makeFullPageDistanceDocx(at: "body" | "paragraph" = "body"): Uint8Array {
  return makePageDistanceDocx('w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="288"', at);
}

