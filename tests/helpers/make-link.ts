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
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/hyperlink" Target="https://example.com/detail" TargetMode="External"/></Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${document}</w:body></w:document>`),
    },
    { level: 6 },
  );

/**
 * The two ways a `w:hyperlink` points somewhere. An outside target travels as a relationship; a place inside
 * this document has no relationship at all and is spelled `w:anchor` — the shape Word and WPS use for a
 * table-of-contents entry, a cross-reference, or "jump to bookmark".
 *
 * The element also carries attributes of its own that say nothing about the target: `w:history` (remember
 * the visited colour), `w:docProp` (take the text from a document property), `w:tm`, `w:invalid`. Word writes
 * `w:history` on every link it creates.
 *
 * The WPS corpus holds no `w:hyperlink` at all (`.scratch/wps-hyperlink-attrs.ts`: 0 of 92 files), so these
 * shapes are ECMA-376-derived rather than WPS-verified. They exist because the loss is measurable: an
 * anchor link that the model cannot name is gone the moment anyone types in its paragraph, and an attribute
 * it cannot name is gone the moment the file is saved.
 */
export function makeLinkDocx(): Uint8Array {
  return pack(
    `<w:p><w:bookmarkStart w:id="0" w:name="_Toc_1"/><w:hyperlink w:anchor="_Toc_1" w:history="1"><w:r><w:rPr><w:b/></w:rPr><w:t>跳到书签</w:t></w:r></w:hyperlink><w:bookmarkEnd w:id="0"/></w:p>
<w:p><w:r><w:t>链接前</w:t></w:r><w:hyperlink r:id="rId5" w:history="1"><w:r><w:t>外部链接</w:t></w:r></w:hyperlink><w:r><w:t>链接后</w:t></w:r></w:p>
<w:p><w:ins w:id="7" w:author="甲"><w:hyperlink w:anchor="_Toc_1"><w:r><w:t>修订里的锚点</w:t></w:r></w:hyperlink></w:ins></w:p>
<w:p><w:hyperlink w:anchor="_Toc_1" w:docProp="title" w:tm="2024-01-01T00:00:00"><w:r><w:t>按名域取文字</w:t></w:r></w:hyperlink></w:p>`,
  );
}

/** The same anchor link split over two runs, which is what Word writes when part of it was formatted apart. */
export function makeSplitAnchorLinkDocx(): Uint8Array {
  return pack(
    `<w:p><w:hyperlink w:anchor="_Toc_2"><w:r><w:t>前半</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>后半</w:t></w:r></w:hyperlink></w:p>`,
  );
}

/**
 * One link whose runs the page has pulled apart into two elements: the same anchor, but only the first
 * carries the attributes the file wrote. A writer that merges by target alone would give them to both.
 */
export function makeHalfAttributedLinkDocx(): Uint8Array {
  return pack(
    `<w:p><w:hyperlink w:anchor="_Toc_3" w:history="1"><w:r><w:t>带属性</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="_Toc_3"><w:r><w:t>不带属性</w:t></w:r></w:hyperlink></w:p>`,
  );
}

/** An attribute written with XML's other quote style: it travels as it was written, not re-spelled. */
export function makeSingleQuotedLinkDocx(): Uint8Array {
  return pack(
    `<w:p><w:hyperlink w:anchor="_Toc_4" w:history='1'><w:r><w:t>单引号</w:t></w:r></w:hyperlink></w:p>`,
  );
}

/** A link whose target this model cannot name: the element is already plain runs, so it has no attributes to carry. */
export function makeTargetlessLinkDocx(): Uint8Array {
  return pack(
    `<w:p><w:hyperlink w:anchor="" w:history="1"><w:r><w:t>无目标</w:t></w:r></w:hyperlink></w:p>`,
  );
}
