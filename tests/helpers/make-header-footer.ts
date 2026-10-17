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
 * The section's reference to the header, in the order the authorized WPS templates write it: 授权 46 份的
 * **38 枚 `w:headerReference`/`w:footerReference` 全部是 `r:id` 在前、`w:type` 在后、自闭合，且 38 枚全写 `w:type`**
 * （`.scratch/band-per-section.ts`；枚数按类型分：`footer:default` 26、`header:default` 6、`footer:even` 4、
 * `footer:first` 2）。
 */
export const REFERENCE_ID_FIRST = '<w:headerReference r:id="rId5" w:type="default"/>';

/** Word's own order, which 授权 46 份的 38 枚引用里没有一枚使用。两种都得读。 */
export const REFERENCE_TYPE_FIRST = '<w:headerReference w:type="default" r:id="rId5"/>';

/** A section that binds only a first-page header. Shown only when the section also states `w:titlePg`. */
export const REFERENCE_FIRST_PAGE_ONLY = '<w:headerReference r:id="rId5" w:type="first"/>';

/** A section that binds only an even-page header — a part held, and on one page never the one to show. */
export const REFERENCE_EVEN_ONLY = '<w:headerReference r:id="rId5" w:type="even"/>';

/**
 * A one-paragraph body whose section points at one header or footer part. The caller owns that part whole,
 * declaration included, so what a save hands back can be compared to it byte for byte. `titlePg` writes the
 * section's 「首页不同」 switch, which is what decides whether a `w:type="first"` reference is the page's band.
 */
function packHf(kind: "header" | "footer", part: string, reference: string, titlePg = false): Uint8Array {
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
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr>${reference}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>${titlePg ? "<w:titlePg/>" : ""}</w:sectPr></w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/${kind}" Target="${kind}1.xml"/></Relationships>`,
      ),
      [name]: strToU8(part),
    },
    { level: 6 },
  );
}

export function makeHeaderFooterDocx(headerPart: string, reference: string = REFERENCE_ID_FIRST, titlePg = false): Uint8Array {
  return packHf("header", headerPart, reference, titlePg);
}

/** The same file with the part bound as the section's footer, which is where a page number usually lives. */
export function makeFooterDocx(footerPart: string): Uint8Array {
  return packHf("footer", footerPart, '<w:footerReference r:id="rId5" w:type="default"/>');
}

/**
 * A header in the shape the authorized WPS templates write: 授权 46 份共 **38 个 `word/header*.xml` /
 * `word/footer*.xml` 部件，38/38 的根元素都带 `mc:Ignorable`，根上合计声明 646 个命名空间（平均 17 个）；
 * 部件内 **76 段**，其中 **62 段给段落写了 `w:pStyle`、12 段写了 `w:jc`、2 段带一个字面 `w:tab`**。
 * So the first paragraph here has all of it — the
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
 * A page field as the authorized templates write one — and as `insertPageNumber` writes one: `begin`, the
 * instruction, `separate`, the number it last showed, `end`. 授权 46 份的 `word/header*.xml` / `word/footer*.xml`
 * 里数到 **38 枚 `w:instrText` 与 112 枚 `w:fldChar`（begin 38 / separate 36 / end 38）**——差的那 2 枚是
 * **不带 `separate` 的域**，即没有缓存结果的写法：指令首词 **PAGE 36 枚、NUMPAGES 2 枚**；带 `\* MERGEFORMAT` 的 28 枚、
 * 不带的 10 枚；指令串两侧留空格的 34 枚、不留的 4 枚；**36 枚缓存着一串数字**、2 枚没有缓存结果。作废语料那句
 * 「21 枚 / 16 枚这样拼 / 2 枚裸写 PAGE / 21 枚都缓存」不再是可引读数。
 */
export const pageFieldRuns = (instruction = " PAGE  \\* MERGEFORMAT ", cached = "1"): string =>
  `${fldChar("begin")}${instr(instruction)}${fldChar("separate")}<w:r><w:t>${cached}</w:t></w:r>${fldChar("end")}`;

/**
 * The same field with **no `separate` and therefore no cached number**. 授权 46 份的 38 枚域里有 **2 枚**是这个
 * 形状，其余 36 枚都带缓存结果；0 枚嵌套域（`.scratch/field-stage-shapes.ts` 按深度走了一遍，与枚数口径 112/112
 * 一致）。这一支问的是「文件没说这一页是第几页」时我们别自己编一个。
 */
export const pageFieldWithoutResult = (instruction = " PAGE  \\* MERGEFORMAT "): string =>
  `${fldChar("begin")}${instr(instruction)}${fldChar("end")}`;

/** 「第 」, a page field, 「 页」 — the shape a page band takes when it is not the field alone. */
export const PAGE_FOOTER = footerPart(
  `<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t xml:space="preserve">第 </w:t></w:r>${pageFieldRuns()}<w:r><w:t xml:space="preserve"> 页</w:t></w:r></w:p><w:p><w:r><w:t>普通一行</w:t></w:r></w:p>`,
);

/** 「第 」与「 页」中间那枚域不带缓存结果 —— 授权语料那 2 枚的形状。 */
export const PAGE_NO_RESULT_FOOTER = footerPart(
  `<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t xml:space="preserve">第 </w:t></w:r>${pageFieldWithoutResult()}<w:r><w:t xml:space="preserve"> 页</w:t></w:r></w:p>`,
);

/** A footer whose only content is the field — the majority shape among the 38 fields counted above. */
export const PAGE_ONLY_FOOTER = footerPart(`<w:p>${pageFieldRuns()}</w:p>`);

/** The field alone in its paragraph, with the paragraph's formatting on the run that carries the digit. */
export const PAGE_ONLY_FOOTER_STYLED = footerPart(
  `<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
);

/** The footer the opening section binds, and the one the closing section binds. Different parts, different text. */
export const FIRST_SECTION_FOOTER = footerPart(`<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>开篇页脚</w:t></w:r></w:p>`);
export const LAST_SECTION_FOOTER = footerPart(`<w:p><w:r><w:t>结尾页脚</w:t></w:r></w:p>`);

/**
 * A document with two sections: the break hanging off the first paragraph ends the opening one, and the
 * body's own `sectPr` is the closing one. Each carries its own header and its own page box — portrait first,
 * landscape last — and each `sectPr` wears a distinguishable `w:rsidR`, so a test can say which section an
 * edited file put a reference in. `footers` gives the sections footers too: `both` one each, `first` only the
 * opening section cites one (so the closing section carries it forward), `last` only the **later** section
 * cites one (so the opening section shows none), `none` neither.
 *
 * 授权口径的受众（`.scratch/section-band-audience.mjs` 与 `.scratch/band-per-section.ts` 两条口径同为 46 份 58 节）：
 * **多节的件 8 个实例 = 4 个模板身份 × 各自的本地化副本**；这 4 个模板里 **2 个的第一节什么都不 cite、后面的节各
 * cite 一枚页脚**（labor contract 是 `[沉默, footer:default]`，project plan 是 `[沉默, footer:default,
 * footer:default]`），另外 2 个每一节都 cite。「页眉只由后面的节 cite」这条在授权口径上 **0 份**，所以 `last` 这一
 * 支只喂页脚那一半的问题；作废语料那句「18 of 86 多节 / 4 份页眉只由后面的节 cite / 5 份页脚」不再可引。
 */
export function makeTwoSectionDocx(o: { firstRef?: string; lastRef?: string; firstLandscape?: boolean; firstMarginTwips?: number; footers?: "both" | "first" | "last" | "none" } = {}): Uint8Array {
  const firstRef = o.firstRef === undefined ? '<w:headerReference r:id="rId5" w:type="default"/>' : o.firstRef;
  const lastRef = o.lastRef === undefined ? '<w:headerReference r:id="rId6" w:type="default"/>' : o.lastRef;
  // 带子按节取要看的是「末张的那一份归谁」：两节各 cite 一枚页脚时，纸盒外那份得是第二节的，第一节的那份
  // 降级成母本；只有第一节 cite 时末张接着第一节的，纸盒外不再有第二份；只有第二节 cite 时第一节什么都不显示
  // ——向前继承不回头（§17.10.3 的 linkToPrevious 只往回找，不往后找）。
  const ftr = o.footers ?? "none";
  const firstFtr = ftr === "none" || ftr === "last" ? "" : '<w:footerReference r:id="rId7" w:type="default"/>';
  const lastFtr = ftr === "both" || ftr === "last" ? '<w:footerReference r:id="rId8" w:type="default"/>' : "";
  // The opening section's page box either way round, so a check can tell "the band follows the section that
  // owns it" from "the band matches a portrait page by coincidence" — the stylesheet's own fallback is 794px.
  const firstBox = o.firstLandscape ? '<w:pgSz w:w="16838" w:h="11906"/>' : '<w:pgSz w:w="11906" w:h="16838"/>';
  // 授权 46 份里 cite 过带的 **22 份，首节全都写 `w:pgMar`，且 22/22 不是 1440 那个默认值**（全体 46 份的首节里
  // 36 份写 1800/1800，其余是 1701、1361、1080、1588/1474、1588/1418 这些取整值），所以 the fixture wears that
  // shape too — otherwise a band drawn at the stylesheet's fixed 96px padding lines up by accident and the
  // check proves nothing.
  const firstMar = o.firstMarginTwips ?? 1440;
  // 每一枚 cite 都得有它自己的件、关系与 `[Content_Types]` 声明：`footers: "last"` 引用 rId8 却没有 `word/footer2.xml`
  // 是一枚悬空引用，严格读法直接拒件，那这一支测的就不是「向前继承不回头」而是包装本身。
  const hasFirstFtr = ftr === "both" || ftr === "first";
  const hasLastFtr = ftr === "both" || ftr === "last";
  const footerTypes = (hasFirstFtr ? `<Override PartName="/word/footer1.xml" ContentType="${CT_FTR}"/>` : "") + (hasLastFtr ? `<Override PartName="/word/footer2.xml" ContentType="${CT_FTR}"/>` : "");
  const footerRels = (hasFirstFtr ? `<Relationship Id="rId7" Type="${NS_R}/footer" Target="footer1.xml"/>` : "") + (hasLastFtr ? `<Relationship Id="rId8" Type="${NS_R}/footer" Target="footer2.xml"/>` : "");
  const footerParts = { ...(hasFirstFtr ? { "word/footer1.xml": strToU8(FIRST_SECTION_FOOTER) } : {}), ...(hasLastFtr ? { "word/footer2.xml": strToU8(LAST_SECTION_FOOTER) } : {}) };
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/header1.xml" ContentType="${CT_HDR}"/><Override PartName="/word/header2.xml" ContentType="${CT_HDR}"/>${footerTypes}</Types>`,
      ),
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
          `<w:p><w:pPr><w:sectPr w:rsidR="00A1">${firstRef}${firstFtr}${firstBox}<w:pgMar w:top="1440" w:right="${firstMar}" w:bottom="1440" w:left="${firstMar}"/></w:sectPr></w:pPr><w:r><w:t>第一节</w:t></w:r></w:p>` +
          `<w:p><w:r><w:t>第二节</w:t></w:r></w:p>` +
          `<w:sectPr w:rsidR="00B2">${lastRef}${lastFtr}<w:pgSz w:w="16838" w:h="11906"/></w:sectPr>` +
          `</w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/header" Target="header1.xml"/><Relationship Id="rId6" Type="${NS_R}/header" Target="header2.xml"/>${footerRels}</Relationships>`,
      ),
      "word/header1.xml": strToU8(FIRST_SECTION_HEADER),
      "word/header2.xml": strToU8(LAST_SECTION_HEADER),
      ...footerParts,
    },
    { level: 6 },
  );
}

/** A header of one plain paragraph, distinguishable by text alone. */
const plainHeader = (t: string): string => `${XML_DECL}<w:hdr xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:hdr>`;

export const BAND_DEFAULT_TEXT = "默认页眉";
export const BAND_FIRST_TEXT = "首页页眉";
export const BAND_EVEN_TEXT = "偶数页眉";

/**
 * A section that cites the whole first-page band set the way this machine's WPS cites it, with 「首页不同」
 * either on or off — the fixture the question 「首页不同」 opened on which part is the page's is about.
 *
 * 授权口径量到的形状（`.scratch/sect-ref-order.ts`，46 份 58 节）：**58 节里 28 节什么都不 cite、24 节只 cite
 * 一枚、6 节 cite 多枚**，多枚那 6 节按 `w:type` 分档写——`footer:default > footer:even`（2 节）、
 * `header:default > footer:even`（2 节）、`footer:first > header:default > footer:default`（2 节）；
 * `w:type="first"` 在授权口径上**只出现在页脚那一侧（2 枚）**，「first 档里页眉页脚谁前谁后」因此是 **未取证**。
 * 正文里说 `w:titlePg` 的只有 **2 枚**，而 **46/46 的 `word/settings.xml` 都没说 `evenAndOddHeaders`**：这就是
 * 为什么一枚 even 引用在这里是「部件持有」而不是「这一张显示」。`w:titlePg` 排在 `w:pgMar` 与 `w:cols` 之后，
 * 按的是写了它的那两节自己的位置。这里只把页眉那一侧的交错拼出来——问的是这一张纸显示哪一份页眉，而模型的读法
 * 本来就是两类各问各的。
 *
 * The part names, the relationship ids and the three texts are ours; the shapes are transcriptions.
 */
export function makeFirstPageHeaderDocx(o: { titlePg?: boolean; refs?: ("first" | "default" | "even")[] } = {}): Uint8Array {
  const refs = o.refs ?? ["first", "default", "even"];
  const target = (t: "first" | "default" | "even") => (t === "first" ? "header2.xml" : t === "even" ? "header3.xml" : "header1.xml");
  const id = (t: "first" | "default" | "even") => (t === "first" ? "rId6" : t === "even" ? "rId7" : "rId5");
  const reference = (t: "first" | "default" | "even") => `<w:headerReference r:id="${id(t)}" w:type="${t}"/>`;
  const used = new Set(refs.map(target));
  return zipSync(    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/>` +
          [...used].map((n) => `<Override PartName="/word/${n}" ContentType="${CT_HDR}"/>`).join("") +
          `</Types>`,
      ),
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p>` +
          `<w:sectPr>${refs.map(reference).join("")}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:space="425" w:num="1"/>${o.titlePg ? "<w:titlePg/>" : ""}</w:sectPr>` +
          `</w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}">` +
          refs.map((t) => `<Relationship Id="${id(t)}" Type="${NS_R}/header" Target="${target(t)}"/>`).join("") +
          `</Relationships>`,
      ),
      // Only the cited parts, and each declared once above: a part in the package with no `Override` is a file
      // a strict reader refuses, and the point here is which cited part the page shows, not which sit unused.
      ...Object.fromEntries(refs.map((t) => [`word/${target(t)}`, strToU8(plainHeader(t === "first" ? BAND_FIRST_TEXT : t === "even" ? BAND_EVEN_TEXT : BAND_DEFAULT_TEXT))])),
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

