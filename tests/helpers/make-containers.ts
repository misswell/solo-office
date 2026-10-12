import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

const pack = (document: string): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/></Types>`,
      ),
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:my="urn:contoso:sample"><w:body>${document}</w:body></w:document>`),
    },
    { level: 6 },
  );

/**
 * The body containers this model does not explain, in the shapes Word writes for them: a content control
 * holding a two-paragraph table of contents (`w:sdt` with `w:sdtPr`/`w:sdtContent`), a `w:customXml` block
 * wrapping a paragraph with an attribute of its own, an inline `w:sdt`, a `w:fldSimple` page field, and a
 * bookmark stretched over a paragraph — each sitting between ordinary paragraphs so a round trip can be
 * asked one question: did the structure and its place in the document survive, and did its text.
 *
 * The WPS corpus in `.scratch/wps-corpus/` holds no such file (`.scratch/sdt-scan.mjs`: 0 of 3 docx), so
 * these shapes are ECMA-derived, not WPS-verified. They exist so the loss is measured rather than assumed.
 */
export function makeContainersDocx(): Uint8Array {
  const document = `<w:p><w:r><w:t>容器前的段落</w:t></w:r></w:p>
<w:sdt><w:sdtPr><w:id w:val="-1666666666"/><w:alias w:val="目录标题"/><w:tag w:val="TocEntrance"/><w:dataBinding w:xpath="/w:document[1]/w:body[1]" w:storeItemID="{12345678-ABCD-4321-DCBA-87654321FEDC}"/></w:sdtPr><w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>第一章 概览</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>第二章 细节</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:customXml w:item="&lt;my:块 xmlns:my=&quot;urn:contoso:sample&quot;/&gt;" w:itemId="A1" w:itemName="样张标记"><w:customXmlPr><w:attr w:name="owner" w:value="linxiao"/></w:customXmlPr><w:p><w:r><w:t>自定义标记包着的一段</w:t></w:r></w:p></w:customXml>
<w:p><w:r><w:t>行内控件在这里：</w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="日期"/><w:tag w:val="InlineDate"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>2026 年 10 月 11 日</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>（后面还有字）</w:t></w:r></w:p>
<w:p><w:r><w:t>页码 </w:t></w:r><w:fldSimple w:instr=" PAGE   \\* MERGEFORMAT "><w:r><w:rPr><w:noProof/></w:rPr><w:t>12</w:t></w:r></w:fldSimple><w:bookmarkStart w:id="0" w:name="_Toc_1"/><w:r><w:t>书签跨过的字</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
<w:p><w:r><w:t>容器后的段落</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * The same structures inside a table cell and inside a paragraph that sits in a `w:sdt` block, because the
 * cell is the other place the editor rebuilds its model from the page: a container that survives the body
 * but not a cell is still lost the moment somebody types in it.
 */
export function makeContainersInTableDocx(): Uint8Array {
  const document = `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:bookmarkStart w:id="3" w:name="_CellMark"/><w:r><w:t>格里的普通段</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p><w:sdt><w:sdtPr><w:alias w:val="格内控件"/><w:tag w:val="CellSdt"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>格内控件的字</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>
<w:p><w:r><w:t>表后的段落</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * A container the body's read never sees because it lives in the header part. It asks the other question:
 * a save that keeps the body's containers must not reach into a part this model cannot explain.
 */
export function makeContainerInHeaderDocx(): Uint8Array {
  const header = `<w:hdr xmlns:w="${NS_W}"><w:p><w:r><w:t>页眉的普通段</w:t></w:r></w:p><w:sdt><w:sdtPr><w:alias w:val="页眉控件"/><w:tag w:val="HeaderSdt"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>页眉控件的字</w:t></w:r></w:p></w:sdtContent></w:sdt></w:hdr>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body><w:p><w:r><w:t>正文段落</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId5"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/header" Target="header1.xml"/></Relationships>`,
      ),
      "word/header1.xml": strToU8(XML_DECL + header),
    },
    { level: 6 },
  );
}

/**
 * The one inline structure the WPS corpus really holds: a `_GoBack` pair alone in an empty first paragraph,
 * copied from what `.scratch/bookmark-scan.mjs` read out of three WPS-written files. Every other inline
 * shape in these helpers is ECMA-derived; this one is a transcription.
 */
export function makeGoBackDocx(): Uint8Array {
  const document = `<w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p>
<w:p><w:r><w:t>第二段的字</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;
  return pack(document);
}

/**
 * A container the body was written round with a comment between it and its neighbour. The comment is not a
 * child either list counts, so the pairing that decides where a container sits has to hold anyway — a
 * producer that comments its markup is not a producer that moved the body's contents.
 */
export function makeContainerWithCommentDocx(): Uint8Array {
  const document = `<w:p><w:r><w:t>注释前的段落</w:t></w:r></w:p><!-- 中间的一句注释 -->
<w:sdt><w:sdtPr><w:alias w:val="带注释的控件"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>控件里的字</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p><w:r><w:t>注释后的段落</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;
  return pack(document);
}

/**
 * The other way a document holds a field: not one `w:fldSimple` element but a run sequence — begin, the
 * instruction, separate, the cached result, end — with a `w:proofErr` pair stretched over it. The machinery
 * lives *inside* runs, so a reader that only looks for elements beside them never sees it, and a run whose
 * only content is field machinery reads as nothing at all.
 *
 * ECMA-derived: `.scratch/bookmark-scan.mjs` counted 0 `w:fldChar` and 0 `w:instrText` in the WPS corpus,
 * so this is the shape the spec gives and Word writes, not one transcribed from a WPS file.
 */
export function makeFieldCharDocx(): Uint8Array {
  const document = `<w:p><w:r><w:t>页码 </w:t></w:r><w:proofErr w:type="spellStart"/>
<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:proofErr w:type="spellEnd"/></w:p>
<w:p><w:r><w:t>域后面的段落</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;
  return pack(document);
}
