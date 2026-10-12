import { describe, expect, it } from "vitest";
import type { Block, WordDocument } from "../src/core/docx.js";
import type { CommentAnchor } from "../src/core/comments.js";
import { commentBody } from "../src/core/comments.js";
import { WordDocument as Doc } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { makeCommentedDocx, makeHeaderCommentDocx, makeReactionsDocx, makeThreadedDocx } from "./helpers/make-comments.js";
import { makeDocx } from "./helpers/make.js";

const open = (bytes: Uint8Array): WordDocument => Doc.openBytes(bytes);

/** Every anchor in body order, table cells included. */
function anchorsOf(blocks: Block[]): CommentAnchor[] {
  const out: CommentAnchor[] = [];
  const walk = (list: Block[]): void => {
    for (const b of list) {
      if (b.kind === "table") {
        b.rows.forEach((row) => row.forEach((cell) => walk(cell.paragraphs)));
        continue;
      }
      if (b.kind !== "paragraph") continue;
      for (const inline of b.inlines) if (inline.anchor) out.push(inline.anchor);
    }
  };
  walk(blocks);
  return out;
}

/** What deleting a range in the editor does to the model: its anchors go with it. */
function stripAnchors(doc: WordDocument, id: number): void {
  const walk = (list: Block[]): void => {
    for (const b of list) {
      if (b.kind === "table") {
        b.rows.forEach((row) => row.forEach((cell) => walk(cell.paragraphs)));
        continue;
      }
      if (b.kind !== "paragraph") continue;
      b.inlines = b.inlines.filter((i) => i.anchor?.id !== id);
    }
  };
  walk(doc.blocks);
}

/** The triple a new comment needs around an existing paragraph. */
function anchorParagraph(doc: WordDocument, at: number, id: number): void {
  const block = doc.blocks[at];
  if (block.kind !== "paragraph") throw new Error("fixture: block is not a paragraph");
  block.inlines = [
    { text: "", anchor: { id, kind: "start" } },
    ...block.inlines,
    { text: "", anchor: { id, kind: "end" } },
    { text: "", anchor: { id, kind: "ref" } },
  ];
}

describe("docx 批注：读取", () => {
  it("reads the list, the rich bodies and the anchored ranges", () => {
    const doc = open(makeCommentedDocx());
    expect(doc.comments.map((c) => c.id)).toEqual([0, 1, 2]);
    expect(doc.comments[0].author).toBe("顾峰");
    expect(doc.comments[0].initials).toBe("G");
    expect(doc.comments[0].date).toBe("2026-01-01T09:00:00Z");
    expect(commentBody(doc.comments[0])).toBe("请核对 & 确认");
    expect(doc.comments[1].paragraphs.map((runs) => runs.map((i) => i.text).join(""))).toEqual(["加粗说明与正文", "第二段"]);
    expect(doc.comments[1].paragraphs[0][0].bold).toBe(true);
    // A body holding a table cannot be rebuilt from runs, so it stays read-only.
    expect(doc.comments[2].editable).toBe(false);
    expect(doc.comments.map((c) => c.editable)).toEqual([true, true, false]);
  });

  it("keeps every anchor position, in table cells included", () => {
    expect(anchorsOf(open(makeCommentedDocx()).blocks)).toEqual([
      { id: 0, kind: "start" },
      { id: 0, kind: "end" },
      { id: 0, kind: "ref" },
      { id: 1, kind: "start" },
      { id: 1, kind: "end" },
      { id: 1, kind: "ref" },
      { id: 2, kind: "start" },
      { id: 2, kind: "end" },
      { id: 2, kind: "ref" },
    ]);
  });

  it("counts a document nobody edited as unedited and saves it byte-for-byte", () => {
    const doc = open(makeCommentedDocx());
    expect(doc.edited).toBe(false);
    expect(doc.commentsDirty()).toBe(false);
    const out = OpcPackage.from(writeDocx(doc));
    for (const part of doc.untouchedParts()) expect(text(out.bytes(part.path)!)).toBe(text(part.bytes));
  });
});

describe("docx 批注：锚点的保存往返", () => {
  it("writes all three anchor sites back when the body text changes", () => {
    const doc = open(makeCommentedDocx());
    const plain = doc.blocks[4];
    if (plain.kind !== "paragraph") throw new Error("fixture changed");
    plain.inlines = [{ text: "改过的段落" }];
    const bytes = writeDocx(doc);
    const body = OpcPackage.from(bytes).text("word/document.xml")!;
    for (const id of [0, 1, 2]) {
      expect(body).toContain(`<w:commentRangeStart w:id="${id}"/>`);
      expect(body).toContain(`<w:commentRangeEnd w:id="${id}"/>`);
      expect(body).toContain(`<w:commentReference w:id="${id}"/>`);
    }
    const back = open(bytes);
    expect(anchorsOf(back.blocks)).toEqual(anchorsOf(doc.blocks));
    expect(back.comments.map((c) => commentBody(c))).toEqual(doc.comments.map((c) => commentBody(c)));
  });

  it("emits the reference mark as Word's own styled run", () => {
    const doc = open(makeCommentedDocx());
    const plain = doc.blocks[4];
    if (plain.kind !== "paragraph") throw new Error("fixture changed");
    plain.inlines = [{ text: "x" }];
    const body = OpcPackage.from(writeDocx(doc)).text("word/document.xml")!;
    for (const id of [0, 1, 2]) {
      expect(body).toContain(`<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`);
    }
  });

  /**
   * Deleting the anchored text is the editor's job (Word prunes the comment then); the
   * writer's rule is narrower — it never drops a comment on its own initiative, so a
   * document that legitimately carries an unanchored comment survives a round trip.
   */
  it("keeps a comment whose anchors left the body instead of dropping it silently", () => {
    const doc = open(makeCommentedDocx());
    doc.blocks[0] = { kind: "paragraph", inlines: [{ text: "x" }] };
    const bytes = writeDocx(doc);
    const pkg = OpcPackage.from(bytes);
    expect(anchorsOf(doc.blocks).map((a) => a.id)).toEqual([1, 1, 1, 2, 2, 2]);
    expect(commentBody(open(bytes).comments.find((c) => c.id === 0)!)).toBe("请核对 & 确认");
    expect(pkg.text("word/comments.xml")).toContain('w:id="0"');
  });

  it("leaves the comments part untouched while only the body changed", () => {
    const doc = open(makeCommentedDocx());
    const source = OpcPackage.from(writeDocx(open(makeCommentedDocx()))).text("word/comments.xml")!;
    const plain = doc.blocks[4];
    if (plain.kind === "paragraph") plain.inlines = [{ text: "只是正文改了" }];
    expect(OpcPackage.from(writeDocx(doc)).text("word/comments.xml")).toBe(source);
  });
});

describe("docx 批注：新增、编辑与删除", () => {
  it("rebuilds only the comment whose text changed and keeps the rest verbatim", () => {
    const doc = open(makeCommentedDocx());
    const untouched = doc.comments[0].raw;
    expect(doc.comments[0].raw).toContain("请核对 &amp; 确认");
    expect(doc.setCommentBody(1, "改过的第一行\n改过的第二行")).toBe(true);
    expect(doc.edited).toBe(true);
    expect(doc.commentsDirty()).toBe(true);

    const xml = OpcPackage.from(writeDocx(doc)).text("word/comments.xml")!;
    expect(xml).toContain(untouched);
    expect(xml).toContain("改过的第一行");
    expect(xml).toContain("改过的第二行");
    expect(xml).not.toContain("加粗说明");
    // The rebuilt paragraphs carry the ids commentsExtended.xml already points at.
    expect(xml).toContain('w14:paraId="00000002"');
    // The extension rows are rebuilt from the model, so what has to survive is the state:
    // the same keys, the same resolved flag, and the durable ids the package carried.
    const saved = OpcPackage.from(writeDocx(doc));
    expect(saved.text("word/commentsExtended.xml")).toContain('w15:paraId="00000001" w15:done="0"');
    expect(saved.text("word/commentsExtended.xml")).toContain('w15:paraId="00000002" w15:done="1"');
    expect(saved.text("word/commentsIds.xml")).toContain('w16cid:paraId="00000002" w16cid:durableId="55667788"');
  });

  it("refuses to edit a body it cannot rebuild, so its markup survives", () => {
    const doc = open(makeCommentedDocx());
    expect(doc.setCommentBody(2, "改不动")).toBe(false);
    expect(doc.comments[2].paragraphs.length).toBe(1);
    const xml = OpcPackage.from(writeDocx(doc)).text("word/comments.xml")!;
    expect(xml).toContain("<w:tbl>");
    expect(xml).toContain("批注里的表格");
  });

  it("clears the comment, its anchors and its paragraph-id keys on delete", () => {
    const doc = open(makeCommentedDocx());
    stripAnchors(doc, 0);
    expect(doc.removeComment(0)).toBe(true);
    const bytes = writeDocx(doc);
    const pkg = OpcPackage.from(bytes);
    expect(pkg.text("word/document.xml")).not.toContain('w:id="0"');
    expect(pkg.text("word/comments.xml")).not.toContain("请核对");
    expect(pkg.text("word/comments.xml")).toContain("这段带表格");
    expect(pkg.text("word/commentsExtended.xml")).not.toContain("00000001");
    expect(pkg.text("word/commentsExtended.xml")).toContain("00000002");
    expect(pkg.text("word/commentsIds.xml")).not.toContain("00000001");
    expect(pkg.text("word/commentsIds.xml")).toContain("00000002");
    const back = open(bytes);
    expect(back.comments.map((c) => c.id)).toEqual([1, 2]);
    expect(anchorsOf(back.blocks).filter((a) => a.id === 0)).toEqual([]);
  });

  it("creates the comments part, relationship and content type for a document that had none", () => {
    const doc = open(makeDocx());
    expect(doc.comments).toEqual([]);
    const id = doc.nextCommentId();
    anchorParagraph(doc, 3, id);
    doc.addComment(id, "新写的批注\n第二行", "顾峰", "G");
    expect(doc.commentsDirty()).toBe(true);

    const bytes = writeDocx(doc);
    const pkg = OpcPackage.from(bytes);
    expect(pkg.contentType("word/comments.xml")).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml");
    const rel = [...pkg.rels("word/document.xml").values()].find((r) => r.type.endsWith("/comments"));
    expect(rel?.target).toBe("comments.xml");
    expect(pkg.text("word/document.xml")).toContain('<w:commentRangeStart w:id="0"/>');

    const back = open(bytes);
    expect(back.comments.map((c) => c.id)).toEqual([0]);
    expect(commentBody(back.comments[0])).toBe("新写的批注\n第二行");
    expect(back.comments[0].author).toBe("顾峰");
    expect(back.comments[0].initials).toBe("G");
    expect(anchorsOf(back.blocks)).toEqual([
      { id: 0, kind: "start" },
      { id: 0, kind: "end" },
      { id: 0, kind: "ref" },
    ]);
  });

  it("numbers a new comment above the ids already in use", () => {
    const doc = open(makeCommentedDocx());
    expect(doc.nextCommentId()).toBe(3);
    anchorParagraph(doc, 4, doc.nextCommentId());
    doc.addComment(3, "再来一条", "顾峰", "G");
    const back = open(writeDocx(doc));
    expect(back.comments.map((c) => c.id)).toEqual([0, 1, 2, 3]);
    expect(anchorsOf(back.blocks).filter((a) => a.id === 3)).toHaveLength(3);
  });

  it("round-trips comment text that needs XML escaping", () => {
    const doc = open(makeDocx());
    anchorParagraph(doc, 0, 0);
    doc.addComment(0, '含 <标签> 和 & 符号 "引号"', "顾峰", "G");
    const xml = OpcPackage.from(writeDocx(doc)).text("word/comments.xml")!;
    expect(xml).toContain("含 &lt;标签&gt; 和 &amp; 符号");
    expect(commentBody(open(writeDocx(doc)).comments[0])).toBe('含 <标签> 和 & 符号 "引号"');
  });
});

describe("docx 批注：渲染", () => {
  it("marks the anchor positions and highlights the anchored text", () => {
    const html = renderWord(open(makeCommentedDocx()));
    expect(html).toContain('<span class="cmt-anchor" data-cid="0" data-kind="start" contenteditable="false"></span>');
    expect(html).toContain('<span class="cmt-anchor cmt-ref" data-cid="0" data-kind="ref" contenteditable="false">1</span>');
    // `data-cids` is how the editor maps a caret inside the text back to its comment.
    expect(html).toContain('<span class="cmt-in" data-cids="0"><span>收入同比增长</span></span>');
    // The range ends before "结尾", so only the run before it is inside.
    expect(html).toContain('<span>结尾</span>');
  });

  it("keeps a document without comments exactly as it rendered before", () => {
    const html = renderWord(open(makeDocx()));
    expect(html).not.toContain("cmt-anchor");
    expect(html).not.toContain("cmt-in");
    expect(html).toContain('<div class="para">');
  });
});

describe("docx 批注：页眉里的批注", () => {
  it("reads the comment but leaves the header it lives in alone", () => {
    const bytes = makeHeaderCommentDocx();
    const doc = open(bytes);
    expect(doc.comments.map((c) => c.id)).toEqual([7]);
    expect(commentBody(doc.comments[0])).toBe("页眉里的批注");
    // The body has no marker for it, so the editor must not prune it as an orphan either.
    expect(anchorsOf(doc.blocks)).toEqual([]);
    expect(doc.header?.editable).toBe(false);

    const plain = doc.blocks[0];
    if (plain.kind !== "paragraph") throw new Error("fixture changed");
    plain.inlines = [{ text: "正文改动" }];
    const out = OpcPackage.from(writeDocx(doc));
    const source = OpcPackage.from(bytes);
    for (const name of ["word/header1.xml", "word/comments.xml"]) {
      expect(text(out.bytes(name)!)).toBe(text(source.bytes(name)!));
    }
    expect(out.text("word/document.xml")).toContain("正文改动");
  });
});

describe("docx 批注：解决状态、回复与扩展部件", () => {
  it("reads the resolved flag, the reply link and the durable ids off the extension parts", () => {
    const doc = open(makeThreadedDocx());
    expect(doc.comments.map((c) => c.id)).toEqual([0, 1, 2]);
    expect(doc.comments[0].resolved).toBe(false);
    expect(doc.comments[2].resolved).toBe(true);
    expect(doc.comments[1].parentId).toBe(0);
    expect(doc.comments[0].paraIds).toEqual(["10000001"]);
    expect(doc.comments[2].durableId).toBe("20000003");
    // The reply has no markers of its own: it shares the thread anchored by its parent.
    expect(anchorsOf(doc.blocks).map((a) => a.id)).toEqual([0, 0, 0, 2, 2, 2]);
  });

  it("saves a threaded document nobody touched byte-for-byte", () => {
    const bytes = makeThreadedDocx();
    const doc = open(bytes);
    expect(doc.edited).toBe(false);
    expect(Buffer.from(writeDocx(doc))).toEqual(Buffer.from(bytes));
  });

  it("writes the resolved flag onto the comment's own paragraph id", () => {
    const doc = open(makeThreadedDocx());
    expect(doc.setCommentResolved(0, true)).toBe(true);
    expect(doc.commentsDirty()).toBe(true);
    const saved = OpcPackage.from(writeDocx(doc));
    expect(saved.text("word/commentsExtended.xml")).toContain('w15:paraId="10000001" w15:done="1"');
    expect(open(writeDocx(doc)).comments[0].resolved).toBe(true);
    // Flipping it back is not a change to anything else: the anchors are untouched.
    expect(saved.text("word/document.xml")).toContain('<w:commentRangeStart w:id="0"/>');
  });

  it("resolving a comment resolves every reply under it", () => {
    const doc = open(makeThreadedDocx());
    doc.addComment(doc.nextCommentId(), "回复的回复", "顾峰", "G", 1);
    expect(doc.setCommentResolved(0, true)).toBe(true);
    // Word counts a thread, not a row: an unresolved reply would keep it in the list otherwise.
    expect(doc.comments.filter((c) => c.id !== 2).every((c) => c.resolved === true)).toBe(true);
    const ex = OpcPackage.from(writeDocx(doc)).text("word/commentsExtended.xml")!;
    expect(ex).toContain('w15:paraId="10000002" w15:paraIdParent="10000001" w15:done="1"');
    expect(doc.setCommentResolved(0, false)).toBe(true);
    expect(doc.comments.map((c) => `${c.id}:${c.resolved ? 1 : 0}`).join(",")).toBe("0:0,1:0,2:1,3:0");
  });

  it("keeps extension rows keyed on paragraphs this document does not own", () => {
    const doc = open(makeThreadedDocx());
    doc.setCommentResolved(0, true);
    const saved = OpcPackage.from(writeDocx(doc));
    expect(saved.text("word/commentsExtended.xml")).toContain('w15:paraId="0FEEDFACE"');
    expect(saved.text("word/commentsIds.xml")).toContain('w16cid:paraId="0BEEF00D"');
  });

  it("round-trips a reply that carries no anchors of its own", () => {
    const doc = open(makeThreadedDocx());
    const id = doc.nextCommentId();
    doc.addComment(id, "我的回复", "顾峰", "G", 0);
    const bytes = writeDocx(doc);
    const saved = OpcPackage.from(bytes);
    expect(saved.text("word/comments.xml")).toContain(`w:comment w:id="${id}"`);
    // A reply threads onto its parent's range: the body gains no marker for it.
    expect(saved.text("word/document.xml")).not.toContain(`w:id="${id}"`);
    const replyParaId = new RegExp(`<w:comment w:id="${id}"[^>]*><w:p w14:paraId="([0-9A-F]{8})"`).exec(saved.text("word/comments.xml")!)?.[1];
    expect(replyParaId).toBeDefined();
    expect(saved.text("word/commentsExtended.xml")).toContain(`w15:paraId="${replyParaId}" w15:paraIdParent="10000001"`);
    expect(open(bytes).comments.find((c) => c.id === id)?.parentId).toBe(0);
  });

  it("re-parents a reply instead of dropping it when its comment is deleted", () => {
    const doc = open(makeThreadedDocx());
    expect(doc.removeComment(0)).toBe(true);
    expect(doc.comments.map((c) => c.id)).toEqual([1, 2]);
    expect(doc.comments[0].parentId).toBeUndefined();
    const saved = OpcPackage.from(writeDocx(doc));
    const ex = saved.text("word/commentsExtended.xml")!;
    expect(ex).not.toContain('w15:paraId="10000001"');
    expect(ex).toContain('w15:paraId="10000002"');
    expect(saved.text("word/commentsIds.xml")).not.toContain('w16cid:paraId="10000001"');
    expect(saved.text("word/comments.xml")).toContain("回复父批注");
  });

  it("rewrites the extension parts once the last comment is gone", () => {
    const doc = open(makeThreadedDocx());
    for (const id of doc.comments.map((c) => c.id)) expect(doc.removeComment(id)).toBe(true);
    const saved = OpcPackage.from(writeDocx(doc));
    const ex = saved.text("word/commentsExtended.xml")!;
    const ids = saved.text("word/commentsIds.xml")!;
    // Rows keyed on a deleted comment have to go with it — Word would otherwise be pointed at
    // a paragraph that is not there — while a row nothing owns is kept.
    expect(ex).not.toContain("1000000");
    expect(ex).toContain('w15:paraId="0FEEDFACE"');
    expect(ids).not.toContain("1000000");
    expect(ids).toContain('w16cid:paraId="0BEEF00D"');
    // The part itself stays, related: dropping it would break the relationship it is under.
    expect(saved.text("word/comments.xml")).toContain("</w:comments>");
    expect(saved.text("word/_rels/document.xml.rels")).toContain("comments.xml");
  });

  it("creates the extension parts a package never had, related and typed", () => {
    const doc = open(makeHeaderCommentDocx());
    expect(doc.setCommentResolved(7, true)).toBe(true);
    const saved = OpcPackage.from(writeDocx(doc));
    const ex = saved.text("word/commentsExtended.xml")!;
    const ids = saved.text("word/commentsIds.xml")!;
    expect(ex).toContain('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"');
    expect(ex).toContain('w15:done="1"');
    // The row's key has to exist in comments.xml, so the namespace is declared there too.
    const paraId = /w15:paraId="([0-9A-F]{8})"/.exec(ex)?.[1];
    expect(paraId).toBeDefined();
    expect(saved.text("word/comments.xml")).toContain(`w14:paraId="${paraId}"`);
    expect(ids).toContain(`w16cid:paraId="${paraId}"`);
    const rels = saved.text("word/_rels/document.xml.rels")!;
    expect(rels).toContain('Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"');
    expect(rels).toContain('Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"');
    const types = saved.text("[Content_Types].xml")!;
    expect(types).toContain('<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"');
    expect(types).toContain('<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"');
    expect(open(writeDocx(doc)).comments[0].resolved).toBe(true);
  });

  it("numbers a new comment's id above every id the package already used", () => {
    const doc = open(makeThreadedDocx());
    doc.addComment(doc.nextCommentId(), "新的", "顾峰", "G");
    const ids = [...doc.comments.map((c) => c.paraIds?.[0] ?? ""), ...doc.comments.map((c) => c.durableId ?? "")];
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.every((id) => /^[0-7][0-9A-F]{7}$/.test(id))).toBe(true);
  });
});

describe("docx 批注：本模型不重写的批注部件", () => {
  it("carries the reaction and people parts through a save that rewrites comments", () => {
    const bytes = makeReactionsDocx();
    const before = OpcPackage.from(bytes);
    const doc = open(bytes);
    expect(doc.comments[0].resolved).toBe(true);
    expect(doc.setCommentBody(0, "改过的请反应")).toBe(true);
    const saved = OpcPackage.from(writeDocx(doc));
    expect(saved.text("word/comments.xml")).toContain("改过的请反应");
    for (const part of ["word/commentsExtensible.xml", "word/people.xml"]) {
      expect(saved.text(part)).toBe(before.text(part));
      expect(saved.text("[Content_Types].xml")).toContain(`PartName="/${part}"`);
    }
    const rels = saved.text("word/_rels/document.xml.rels")!;
    expect(rels).toContain('Type="http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible" Target="commentsExtensible.xml"');
    expect(rels).toContain('Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"');
    // The resolve flag lives in the part we do rebuild, so it has to survive the round trip.
    expect(open(writeDocx(doc)).comments[0].resolved).toBe(true);
  });

  it("keeps a document with reactions byte-for-byte while nobody edits it", () => {
    const bytes = makeReactionsDocx();
    expect(writeDocx(open(bytes))).toEqual(bytes);
  });
});

describe("docx 批注：富正文编辑", () => {
  it("keeps run formatting through a text edit", () => {
    const doc = open(makeCommentedDocx());
    const rich = doc.comments.find((c) => c.id === 1)!;
    expect(rich.paragraphs[0][0].bold).toBe(true);
    expect(doc.setCommentRuns(1, [[{ text: "改过的加粗", bold: true }, { text: "与正文" }], [{ text: "第二段" }]])).toBe(true);
    const body = OpcPackage.from(writeDocx(doc)).text("word/comments.xml")!;
    expect(body).toContain('<w:rPr><w:b/></w:rPr><w:t>改过的加粗</w:t>');
    expect(body).toContain('<w:t>与正文</w:t>');
    const again = open(writeDocx(doc)).comments.find((c) => c.id === 1)!;
    expect(again.paragraphs[0][0].bold).toBe(true);
    expect(again.paragraphs[0][1].bold).toBeUndefined();
    expect(commentBody(again)).toBe("改过的加粗与正文\n第二段");
  });

  it("counts a formatting-only change as an edit", () => {
    const doc = open(makeCommentedDocx());
    const rich = doc.comments.find((c) => c.id === 1)!;
    const before = structuredClone(rich.paragraphs);
    // The same words with the emphasis moved: text alone would say nothing changed.
    rich.paragraphs = [before[0].map((r, i) => ({ ...r, bold: i === 1 })), before[1]];
    expect(doc.edited).toBe(true);
    const body = OpcPackage.from(writeDocx(doc)).text("word/comments.xml")!;
    expect(body).toContain('<w:t>加粗说明</w:t>');
    expect(body).toContain('<w:rPr><w:b/></w:rPr><w:t>与正文</w:t>');
  });

  it("does not rebuild a comment whose runs only got re-split", () => {
    const bytes = makeCommentedDocx();
    const doc = open(bytes);
    const plain = doc.comments.find((c) => c.id === 0)!;
    const text = plain.paragraphs[0][0].text;
    // A DOM editor hands one sentence back as several runs. Reading the body as unchanged in
    // that case is what keeps a click in the panel from flattening a comment's markup.
    plain.paragraphs = [[{ text: text.slice(0, 3) }, { text: text.slice(3) }]];
    expect(doc.edited).toBe(false);
    expect(Buffer.from(writeDocx(doc))).toEqual(Buffer.from(bytes));
  });
});
