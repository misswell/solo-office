import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const NS_W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const REL_COMMENTS_EX = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const REL_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const REL_COMMENTS_EXTENSIBLE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
const REL_PEOPLE = "http://schemas.microsoft.com/office/2011/relationships/people";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/**
 * A comment anchored in the page header instead of the body. The header part carries a letterhead
 * table besides the marker, so the editor must leave its bytes alone rather than rewrite the anchor
 * it cannot represent.
 */
export function makeHeaderCommentDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:r><w:t>正文</w:t></w:r></w:p>
<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`;

  const header = `<w:hdr xmlns:w="${NS_W}"><w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>信头</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>页眉文字</w:t></w:r></w:p><w:p><w:commentRangeStart w:id="7"/><w:r><w:t>页眉文字</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p></w:hdr>`;

  const comments = `<w:comments xmlns:w="${NS_W}"><w:comment w:id="7" w:author="顾峰" w:initials="G" w:date="2026-01-04T09:00:00Z"><w:p><w:r><w:t>页眉里的批注</w:t></w:r></w:p></w:comment></w:comments>`;

  const files: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${CT_W}.header+xml"/><Override PartName="/word/comments.xml" ContentType="${CT_W}.comments+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/header" Target="header1.xml"/><Relationship Id="rId2" Type="${NS_R}/comments" Target="comments.xml"/></Relationships>`,
    "word/document.xml": document,
    "word/header1.xml": header,
    "word/comments.xml": comments,
  };
  const out: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) out[k] = strToU8(XML_DECL + v);
  return zipSync(out, { level: 6 });
}

/**
 * A docx carrying every shape a comment can take, written from scratch (not from WPS):
 * a single-paragraph range, a range that crosses paragraphs, an anchor inside a table
 * cell, a rich two-paragraph body, a body holding a table, and the `w14:paraId` keys the
 * extended parts hang off.
 */
export function makeCommentedDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>收入同比增长</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r></w:p>
<w:p><w:r><w:t>跨段落的</w:t></w:r><w:commentRangeStart w:id="1"/></w:p>
<w:p><w:commentRangeEnd w:id="1"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r><w:r><w:t>结尾</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:tc><w:p><w:r><w:t>表格里的</w:t></w:r><w:commentRangeStart w:id="2"/></w:p><w:p><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r><w:r><w:t>文字</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>没有批注的段落</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`;

  const comments = `<w:comments xmlns:w="${NS_W}" xmlns:mc="${NS_MC}" xmlns:w14="${NS_W14}" mc:Ignorable="w14">
<w:comment w:id="0" w:author="顾峰" w:initials="G" w:date="2026-01-01T09:00:00Z"><w:p w14:paraId="00000001"><w:r><w:t>请核对 &amp; 确认</w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="林晓" w:initials="L" w:date="2026-01-02T10:30:00Z"><w:p w14:paraId="00000002"><w:r><w:rPr><w:b/></w:rPr><w:t>加粗说明</w:t></w:r><w:r><w:t>与正文</w:t></w:r></w:p><w:p w14:paraId="00000003"><w:r><w:t>第二段</w:t></w:r></w:p></w:comment>
<w:comment w:id="2" w:author="林晓" w:initials="L" w:date="2026-01-03T10:00:00Z"><w:p w14:paraId="00000004"><w:r><w:t>这段带表格</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>批注里的表格</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:comment>
</w:comments>`;

  // Word keys the extension parts on the `w14:paraId` of each comment's first paragraph, but
  // writes the rows themselves under w15 (resolve state, reply link) and w16cid (durable id).
  const commentsEx = `<w15:commentsEx xmlns:w15="${NS_W15}">
<w15:commentEx w15:paraId="00000001" w15:done="0"/>
<w15:commentEx w15:paraId="00000002" w15:done="1"/>
<w15:commentEx w15:paraId="00000004" w15:done="0"/>
</w15:commentsEx>`;

  const commentsIds = `<w16cid:commentsIds xmlns:w16cid="${NS_W16CID}">
<w16cid:commentId w16cid:paraId="00000001" w16cid:durableId="11223344"/>
<w16cid:commentId w16cid:paraId="00000002" w16cid:durableId="55667788"/>
</w16cid:commentsIds>`;

  const files: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${CT_W}.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="${CT_W}.commentsExtended+xml"/><Override PartName="/word/commentsIds.xml" ContentType="${CT_W}.commentsIds+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/comments" Target="comments.xml"/><Relationship Id="rId6" Type="${REL_COMMENTS_EX}" Target="commentsExtended.xml"/><Relationship Id="rId7" Type="${REL_COMMENTS_IDS}" Target="commentsIds.xml"/></Relationships>`,
    "word/document.xml": document,
    "word/comments.xml": comments,
    "word/commentsExtended.xml": commentsEx,
    "word/commentsIds.xml": commentsIds,
  };
  const out: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) out[k] = strToU8(XML_DECL + v);
  return zipSync(out, { level: 6 });
}

/**
 * A thread as Word writes it: an anchor comment, a reply that carries no markers of its own
 * (it hangs off the parent's paragraph id), a resolved comment, and one row in each extension
 * part keyed on a paragraph this document does not contain.
 */
export function makeThreadedDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>第一段</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r></w:p>
<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>第二段</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`;

  const comments = `<w:comments xmlns:w="${NS_W}" xmlns:mc="${NS_MC}" xmlns:w14="${NS_W14}" mc:Ignorable="w14">
<w:comment w:id="0" w:author="顾峰" w:initials="G" w:date="2026-02-01T09:00:00Z"><w:p w14:paraId="10000001"><w:r><w:t>父批注</w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="林晓" w:initials="L" w:date="2026-02-01T09:30:00Z"><w:p w14:paraId="10000002"><w:r><w:t>回复父批注</w:t></w:r></w:p></w:comment>
<w:comment w:id="2" w:author="顾峰" w:initials="G" w:date="2026-02-02T09:00:00Z"><w:p w14:paraId="10000003"><w:r><w:t>已经处理完了</w:t></w:r></w:p></w:comment>
</w:comments>`;

  // The last row of each part is keyed on a paragraph id nothing here owns.
  const commentsEx = `<w15:commentsEx xmlns:w15="${NS_W15}">
<w15:commentEx w15:paraId="10000001" w15:done="0"/>
<w15:commentEx w15:paraId="10000002" w15:paraIdParent="10000001" w15:done="0"/>
<w15:commentEx w15:paraId="10000003" w15:done="1"/>
<w15:commentEx w15:paraId="0FEEDFACE" w15:done="1"/>
</w15:commentsEx>`;

  const commentsIds = `<w16cid:commentsIds xmlns:w16cid="${NS_W16CID}">
<w16cid:commentId w16cid:paraId="10000001" w16cid:durableId="20000001"/>
<w16cid:commentId w16cid:paraId="10000002" w16cid:durableId="20000002"/>
<w16cid:commentId w16cid:paraId="10000003" w16cid:durableId="20000003"/>
<w16cid:commentId w16cid:paraId="0BEEF00D" w16cid:durableId="20000009"/>
</w16cid:commentsIds>`;

  const files: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${CT_W}.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="${CT_W}.commentsExtended+xml"/><Override PartName="/word/commentsIds.xml" ContentType="${CT_W}.commentsIds+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/comments" Target="comments.xml"/><Relationship Id="rId6" Type="${REL_COMMENTS_EX}" Target="commentsExtended.xml"/><Relationship Id="rId7" Type="${REL_COMMENTS_IDS}" Target="commentsIds.xml"/></Relationships>`,
    "word/document.xml": document,
    "word/comments.xml": comments,
    "word/commentsExtended.xml": commentsEx,
    "word/commentsIds.xml": commentsIds,
  };
  const packed: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) packed[k] = strToU8(XML_DECL + v);
  return zipSync(packed, { level: 6 });
}

/**
 * Comment markup this model reads but does not rewrite: `word/commentsExtensible.xml` (per
 * durable id, the resolve timestamp Word and WPS add) and `word/people.xml` (the author list).
 * Both arrive with a document whose comments a user then edits, so the save has to carry them
 * through with their relationships and content types intact rather than lose them with the parts
 * it regenerates.
 */
export function makeReactionsDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>需要反应的段落</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`;

  const comments = `<w:comments xmlns:w="${NS_W}" xmlns:mc="${NS_MC}" xmlns:w14="${NS_W14}" mc:Ignorable="w14">
<w:comment w:id="0" w:author="顾峰" w:initials="G" w:date="2026-02-03T09:00:00Z"><w:p w14:paraId="30000001"><w:r><w:t>请反应</w:t></w:r></w:p></w:comment>
</w:comments>`;

  const files: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${CT_W}.comments+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="${CT_W}.commentsExtended+xml"/><Override PartName="/word/commentsIds.xml" ContentType="${CT_W}.commentsIds+xml"/><Override PartName="/word/commentsExtensible.xml" ContentType="${CT_W}.commentsExtensible+xml"/><Override PartName="/word/people.xml" ContentType="${CT_W}.people+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/comments" Target="comments.xml"/><Relationship Id="rId6" Type="${REL_COMMENTS_EX}" Target="commentsExtended.xml"/><Relationship Id="rId7" Type="${REL_COMMENTS_IDS}" Target="commentsIds.xml"/><Relationship Id="rId8" Type="${REL_COMMENTS_EXTENSIBLE}" Target="commentsExtensible.xml"/><Relationship Id="rId9" Type="${REL_PEOPLE}" Target="people.xml"/></Relationships>`,
    "word/document.xml": document,
    "word/comments.xml": comments,
    "word/commentsExtended.xml": `<w15:commentsEx xmlns:w15="${NS_W15}"><w15:commentEx w15:paraId="30000001" w15:done="1"/></w15:commentsEx>`,
    "word/commentsIds.xml": `<w16cid:commentsIds xmlns:w16cid="${NS_W16CID}"><w16cid:commentId w16cid:paraId="30000001" w16cid:durableId="30000001"/></w16cid:commentsIds>`,
    "word/commentsExtensible.xml": `<w16cex:commentsExtensible xmlns:w16cex="${NS_W16CEX}"><w16cex:commentExtensible w16cex:durableId="30000001" w16cex:dateUtc="2026-02-03T10:00:00Z"/></w16cex:commentsExtensible>`,
    "word/people.xml": `<w15:people xmlns:w15="${NS_W15}"><w15:person w15:author="顾峰"><w15:presenceInfo w15:providerId="None" w15:userId="顾峰"/></w15:person></w15:people>`,
  };
  const packed: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) packed[k] = strToU8(XML_DECL + v);
  return zipSync(packed, { level: 6 });
}

/**
 * 批注的锚挂在修订的字上（#189）：一条压在 `w:del` 里、一条压在 `w:ins` 里、一条压在没人动过的字上。
 * 三档之下「在场」的形状不一样——删掉的字在最终状态不画，插入的字在原文不画，普通字三档都画——所以这一份
 * 一次给齐「定位」要答的三种情形。第 2 段另带一条段落属性改动（`w:pPrChange`），让同一份件里
 * `revealRevision` 也有可指的行。授权 46 份里批注族 0 枚，这份是自己写的夹具；WPS 在这一处说什么、跳不跳，
 * **未取证**。
 */
export function makeCommentRevDocx(): Uint8Array {
  const document = `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}">
<w:body>
<w:p><w:commentRangeStart w:id="0"/><w:del w:id="51" w:author="顾峰" w:date="2026-03-01T09:00:00Z"><w:r><w:delText>被删掉的锚</w:delText></w:r></w:del><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:commentRangeStart w:id="1"/><w:ins w:id="52" w:author="林晓" w:date="2026-03-01T09:10:00Z"><w:r><w:t>新插入的锚</w:t></w:r></w:ins><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:commentRangeStart w:id="2"/><w:r><w:t>没人动过的字</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>
<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="53" w:author="顾峰" w:date="2026-03-01T09:20:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>属性被改过的第二段</w:t></w:r></w:p>
<w:p><w:r><w:t>尾巴</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`;

  const comments = `<w:comments xmlns:w="${NS_W}">
<w:comment w:id="0" w:author="顾峰" w:initials="G" w:date="2026-03-01T09:05:00Z"><w:p><w:r><w:t>删掉的这段还要不要留</w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="林晓" w:initials="L" w:date="2026-03-01T09:15:00Z"><w:p><w:r><w:t>新加的这句请核对</w:t></w:r></w:p></w:comment>
<w:comment w:id="2" w:author="顾峰" w:initials="G" w:date="2026-03-01T09:25:00Z"><w:p><w:r><w:t>普通的一条</w:t></w:r></w:p></w:comment>
</w:comments>`;

  const files: Record<string, string> = {
    "[Content_Types].xml": `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${CT_W}.comments+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_R}/comments" Target="comments.xml"/></Relationships>`,
    "word/document.xml": document,
    "word/comments.xml": comments,
  };
  const out: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) out[k] = strToU8(XML_DECL + v);
  return zipSync(out, { level: 6 });
}
