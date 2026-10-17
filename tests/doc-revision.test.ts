import { describe, expect, it } from "vitest";
import type { Block, Inline, Paragraph, ParagraphFormatRevision, Revision, Table, WordDocument } from "../src/core/docx.js";
import { WordDocument as Doc } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { makeBreakableDocx, makeRevisedDocx, makeNestedRevisionDocx, makeMovedDocx, makeParaMarkDocx, makeStraddleDocx, makeMoveMarkDocx, makeLowIdDocx, makeTableRowDocx, makeLastRowDeletedDocx } from "./helpers/make-revisions.js";
import { makeCommentedDocx } from "./helpers/make-comments.js";
import { unstamped } from "./helpers/unstamp.js";

const open = (bytes: Uint8Array): WordDocument => Doc.openBytes(bytes);

/** Body paragraphs, table cells included, in document order. */
function paragraphs(doc: WordDocument): Paragraph[] {
  const out: Paragraph[] = [];
  const walk = (list: Block[]): void => {
    for (const b of list) {
      if (b.kind === "table") {
        b.rows.forEach((row) => row.forEach((cell) => walk(cell.paragraphs)));
        continue;
      }
      if (b.kind !== "paragraph") continue;
      out.push(b);
    }
  };
  walk(doc.blocks);
  return out;
}

const textOf = (p: Paragraph): string => p.inlines.map((i) => i.text).join("");

const allRuns = (doc: WordDocument): Inline[] => paragraphs(doc).flatMap((p) => p.inlines);

const body = (doc: WordDocument): string => OpcPackage.from(writeDocx(doc)).text("word/document.xml") ?? "";

describe("docx 修订：读取 w:ins / w:del", () => {
  it("reads an insertion as text that knows who inserted it", () => {
    const doc = open(makeRevisedDocx());
    const run = allRuns(doc).find((r) => r.text === "插入的字");
    expect(run?.rev).toMatchObject({ kind: "ins", id: 1, author: "顾峰", date: "2026-03-01T08:00:00Z" });
  });

  it("reads deleted text, which is spelled w:delText and is not body text yet", () => {
    const doc = open(makeRevisedDocx());
    const run = allRuns(doc).find((r) => r.text === "删掉的字");
    expect(run?.rev).toMatchObject({ kind: "del", id: 2, author: "林晓" });
  });

  it("keeps a revision inside a table cell", () => {
    const doc = open(makeRevisedDocx());
    expect(paragraphs(doc).some((p) => p.inlines.some((r) => r.text === "格内插入" && r.rev?.kind === "ins" && r.rev.id === 6))).toBe(true);
  });

  it("keeps a link that sits inside an insertion", () => {
    const doc = open(makeRevisedDocx());
    const run = allRuns(doc).find((r) => r.text === "链接里的插入");
    expect(run?.rev?.kind).toBe("ins");
    expect(run?.hyperlink).toBe("https://example.com/revised");
  });

  it("carries the run's own formatting alongside the revision", () => {
    const doc = open(makeRevisedDocx());
    const run = allRuns(doc).find((r) => r.text === "加粗插入");
    expect(run?.bold).toBe(true);
    expect(run?.rev?.id).toBe(3);
  });

  it("lists the revisions still waiting, one entry per w:id", () => {
    const ids = open(makeRevisedDocx())
      .revisions()
      .map((r: Revision) => `${r.kind}:${r.id}`);
    expect(ids.join(",")).toBe("ins:1,del:2,ins:3,ins:4,ins:6");
  });

  it("reads a formatting revision's text and keeps the record as its own kind", () => {
    const doc = open(makeRevisedDocx());
    const run = allRuns(doc).find((r) => r.text === "改过格式的字的");
    // It is not an insertion or a deletion: the run shows its new formatting and its text is body text.
    expect(run?.rev).toBeUndefined();
    expect(run?.revFormat).toMatchObject({ kind: "format", id: 5, author: "顾峰", date: "2026-03-01T08:00:00Z", prev: { italic: true } });
    expect(doc.revisions().map((r) => `${r.kind}:${r.id}`).join(",")).toBe("ins:1,del:2,ins:3,ins:4,ins:6");
    expect(doc.formatRevisions().map((r) => `format:${r.id}`).join(",")).toBe("format:5");
  });
});

describe("docx 修订：写回与往返", () => {
  it("saves a document nobody edited byte-for-byte", () => {
    const bytes = makeRevisedDocx();
    expect(unstamped(writeDocx(open(bytes)))).toEqual(unstamped(bytes));
  });

  it("keeps every mark when an unrelated paragraph is edited", () => {
    const doc = open(makeRevisedDocx());
    const first = paragraphs(doc)[0]!;
    first.inlines = [{ text: "未修订的段落改过了" }];
    const xml = body(doc);
    expect(xml).toContain('<w:ins w:id="1" w:author="顾峰" w:date="2026-03-01T08:00:00Z"><w:r><w:t>插入的字</w:t></w:r></w:ins>');
    expect(xml).toContain('<w:del w:id="2" w:author="林晓" w:date="2026-03-01T08:00:00Z"><w:r><w:delText>删掉的字</w:delText></w:r></w:del>');
    expect(xml).toContain('<w:ins w:id="6"');
  });

  it("round-trips the revisions through a save and back", () => {
    const doc = open(makeRevisedDocx());
    const before = JSON.stringify(doc.revisions());
    const back = open(writeDocx(doc));
    expect(JSON.stringify(back.revisions())).toBe(before);
    expect(paragraphs(back).map(textOf).join("|")).toBe(paragraphs(doc).map(textOf).join("|"));
  });

  it("writes a hyperlink inside the insertion, not outside it", () => {
    const xml = body(open(makeRevisedDocx()));
    expect(xml).toMatch(/<w:ins [^>]*><w:hyperlink r:id="rId\d+"><w:r>/);
  });
});

describe("docx 修订：接受与拒绝", () => {
  it("accepting lets insertions stand and deletions take effect", () => {
    const doc = open(makeRevisedDocx());
    expect(doc.resolveRevisions("accept")).toBeGreaterThan(0);
    expect(doc.revisions()).toEqual([]);
    const p = paragraphs(doc)[1]!;
    expect(textOf(p)).toBe("前插入的字后");
    expect(p.inlines.every((r) => r.rev === undefined)).toBe(true);
  });

  it("refusing puts the deleted text back and takes the insertion out", () => {
    const doc = open(makeRevisedDocx());
    doc.resolveRevisions("reject");
    expect(textOf(paragraphs(doc)[1]!)).toBe("前删掉的字后");
    expect(doc.revisions()).toEqual([]);
  });

  it("decides one revision when given its id", () => {
    const doc = open(makeRevisedDocx());
    doc.resolveRevisions("accept", 2);
    expect(textOf(paragraphs(doc)[1]!)).toBe("前插入的字后");
    expect(doc.revisions().map((r) => r.id).join(",")).toBe("1,3,4,6");
  });

  it("writes the decided text as ordinary runs", () => {
    const doc = open(makeRevisedDocx());
    doc.resolveRevisions("accept");
    const xml = body(doc);
    expect(xml).not.toContain("<w:ins");
    expect(xml).not.toContain("<w:del");
    expect(xml).toContain("插入的字");
    expect(xml).not.toContain("删掉的字");
  });

  it("refusing a deletion gives the text back as text, not as a mark", () => {
    const doc = open(makeRevisedDocx());
    doc.resolveRevisions("reject", 2);
    const xml = body(doc);
    expect(xml).not.toContain("<w:delText");
    expect(xml).toContain("<w:t>删掉的字</w:t>");
  });

  it("counts as an edit, so the save regenerates the body", () => {
    const doc = open(makeRevisedDocx());
    expect(doc.edited).toBe(false);
    doc.resolveRevisions("accept", 1);
    expect(doc.edited).toBe(true);
  });
});

describe("docx 修订：页面显示", () => {
  it("marks insertions and deletions and names their authors", () => {
    const html = renderWord(open(makeRevisedDocx()));
    expect(html).toContain('data-rev="ins"');
    expect(html).toContain('data-rev-author="顾峰"');
    expect(html).toContain('class="rev rev-del"');
    expect(html).toContain("删掉的字");
  });

  it("marks them by class rather than by inline style, so an edit cannot read the mark as formatting", () => {
    const html = renderWord(open(makeRevisedDocx()));
    const tag = /<span class="rev rev-(?:ins|del)"[^>]*>/.exec(html)?.[0] ?? "";
    expect(tag).not.toContain("style=");
    expect(tag).not.toContain("text-decoration");
  });
});

describe("docx 修订：修订套在修订里", () => {
  it("reads an insertion nested in one under its own signature", () => {
    const doc = open(makeNestedRevisionDocx());
    expect(allRuns(doc).find((r) => r.text === "外层插入")?.rev).toMatchObject({ kind: "ins", id: 11, author: "顾峰" });
    expect(allRuns(doc).find((r) => r.text === "内层再插入")?.rev).toMatchObject({ kind: "ins", id: 12, author: "林晓" });
  });

  it("reads a deletion nested in one under its own signature", () => {
    const doc = open(makeNestedRevisionDocx());
    expect(allRuns(doc).find((r) => r.text === "外层删除")?.rev).toMatchObject({ kind: "del", id: 13, author: "顾峰" });
    expect(allRuns(doc).find((r) => r.text === "内层再删除")?.rev).toMatchObject({ kind: "del", id: 14, author: "林晓" });
  });

  it("lists both signatures as separate changes waiting", () => {
    expect(open(makeNestedRevisionDocx()).revisions().map((r) => `${r.kind}:${r.id}`).join()).toBe("ins:11,ins:12,del:13,del:14");
  });

  it("decides the inner change without disturbing the outer one", () => {
    const doc = open(makeNestedRevisionDocx());
    expect(doc.resolveRevisions("reject", 12)).toBe(1);
    expect(paragraphs(doc).map(textOf).join("")).toBe("正文外层插入外层删除内层再删除");
    expect(doc.revisions().map((r) => r.id).join()).toBe("11,13,14");
  });

  it("writes the two nested changes as siblings and re-reads them the same way", () => {
    const doc = open(makeNestedRevisionDocx());
    doc.resolveRevisions("accept", 14);
    const back = open(writeDocx(doc));
    expect(back.revisions().map((r) => `${r.kind}:${r.id}`).join()).toBe("ins:11,ins:12,del:13");
    expect(paragraphs(back).map(textOf).join("")).toBe("正文外层插入内层再插入外层删除");
  });

  it("keeps the comment markers a tracked deletion runs through", () => {
    const source = body(open(makeCommentedDocx()));
    const marked = open(makeCommentedDocx());
    const paragraph = paragraphs(marked).find((p) => p.inlines.some((i) => i.anchor))!;
    const rev = { kind: "del" as const, id: 99, author: "审阅者", date: "2026-03-01T08:00:00Z" };
    paragraph.inlines = paragraph.inlines.map((i) => ({ ...i, rev }));
    const out = body(marked);
    const markers = (s: string): number => (s.match(/commentRangeStart|commentRangeEnd|commentReference/g) ?? []).length;
    expect(markers(out)).toBe(markers(source));
    expect(out).toContain('<w:delText>');
    expect(markers(out)).toBeGreaterThan(0);
  });
});

describe("docx 修订：搬移（moveFrom / moveTo）", () => {
  it("reads a tracked move as the deletion and insertion it behaves like", () => {
    const doc = open(makeMovedDocx());
    expect(doc.revisions().map((r) => `${r.kind}:${r.id}`).join()).toBe("del:21,ins:22");
    const moved = allRuns(doc).filter((r) => r.text === "搬走的字");
    expect(moved.map((r) => `${r.rev?.kind}:${r.rev?.author}`)).toEqual(["del:顾峰", "ins:顾峰"]);
    expect(moved.map((r) => r.rev?.move)).toEqual(["from", "to"]);
  });

  it("keeps both copies of the moved text in the model", () => {
    expect(paragraphs(open(makeMovedDocx())).map(textOf)).toEqual(["开头搬走的字结尾", "新位置搬走的字"]);
  });

  it("accepting the move leaves the text only where it arrived", () => {
    const doc = open(makeMovedDocx());
    expect(doc.resolveRevisions("accept")).toBe(2);
    expect(doc.revisions()).toEqual([]);
    expect(paragraphs(doc).map(textOf).join("")).toBe("开头结尾新位置搬走的字");
  });

  it("refusing the move puts the text back where it started", () => {
    const doc = open(makeMovedDocx());
    doc.resolveRevisions("reject");
    expect(paragraphs(doc).map(textOf).join("")).toBe("开头搬走的字结尾新位置");
    expect(doc.revisions()).toEqual([]);
  });

  it("saves an unedited move byte-for-byte and re-reads the same two changes", () => {
    const bytes = makeMovedDocx();
    const doc = open(bytes);
    expect(unstamped(writeDocx(doc))).toEqual(unstamped(bytes));
    const back = open(writeDocx(doc));
    expect(back.revisions().map((r) => `${r.kind}:${r.id}`).join()).toBe("del:21,ins:22");
  });

  it("keeps a move's own tag through an edit elsewhere, losing no character", () => {
    const doc = open(makeMovedDocx());
    // An edit anywhere forces the body to be rebuilt, which is when the model has to say what a move became.
    paragraphs(doc)[0]!.inlines[0]!.text = "起头";
    const xml = body(doc);
    expect(doc.edited).toBe(true);
    expect(xml).toContain('<w:moveFrom w:id="21"');
    expect(xml).toContain('<w:moveTo w:id="22"');
    expect((xml.match(/搬走的字/g) ?? []).length).toBe(2);
    const back = open(writeDocx(doc));
    expect(paragraphs(back).map(textOf)).toEqual(["起头搬走的字结尾", "新位置搬走的字"]);
    expect(back.revisions().map((r) => `${r.kind}:${r.id}`).join()).toBe("del:21,ins:22");
  });

  it("names each half of the move for the pane, because they are two records with two signatures", () => {
    const doc = open(makeMovedDocx());
    expect(doc.revisionEntries().map((e) => `${e.type}:${e.id}:${e.text}`).join(",")).toBe("move-from:21:搬走的字,move-to:22:搬走的字");
  });

  it("decides one half without disturbing the other's tag", () => {
    const doc = open(makeMovedDocx());
    expect(doc.resolveRevisions("accept", 22)).toBe(1);
    expect(body(doc)).toContain('<w:moveFrom w:id="21"');
    expect(body(doc)).not.toContain("<w:moveTo");
    expect(doc.revisionEntries().map((e) => `${e.type}:${e.id}`).join(",")).toBe("move-from:21");
    // The two halves are two records with two signatures, so deciding one on its own is a real answer to
    // a real question: the characters now stand where they arrived *and* are still waiting where they left.
    expect(paragraphs(doc).map(textOf).join("")).toBe("开头搬走的字结尾新位置搬走的字");
    expect(doc.resolveRevisions("reject", 21)).toBe(1);
    expect(paragraphs(doc).map(textOf).join("")).toBe("开头搬走的字结尾新位置搬走的字");
    expect(doc.revisionEntries()).toEqual([]);
  });

  it("reads a moved paragraph's break as the move it is, on the paragraph the break ends", () => {
    const doc = open(makeMoveMarkDocx());
    const ps = paragraphs(doc);
    expect(ps[0]!.markRev).toMatchObject({ kind: "del", move: "from", id: 41, author: "顾峰" });
    expect(ps[2]!.markRev).toMatchObject({ kind: "ins", move: "to", id: 42, author: "顾峰" });
    expect(doc.markRevisions().map((r) => `${r.move}:${r.id}`).join(",")).toBe("from:41,to:42");
    expect(doc.revisionEntries().map((e) => `${e.type}:${e.id}:${e.text}`).join(",")).toBe("move-from:41:¶,move-to:42:¶");
  });

  it("keeps a moved break's own tag through a save and through an edit elsewhere", () => {
    const bytes = makeMoveMarkDocx();
    expect(unstamped(writeDocx(open(bytes)))).toEqual(unstamped(bytes));
    const doc = open(bytes);
    paragraphs(doc)[1]!.inlines[0]!.text = "改过的段落";
    const back = open(writeDocx(doc));
    expect(marksOf(back)).toBe("del:41,ins:42");
    expect(back.revisionEntries().map((e) => `${e.type}:${e.id}`).join(",")).toBe("move-from:41,move-to:42");
    expect(paragraphs(back).map(textOf)).toEqual(["被搬走的那段", "改过的段落", "搬到新处的那段"]);
  });

  /** The break records with their own tags, which is what the pane needs to call a move a move. */
  const marksOf = (doc: WordDocument): string => doc.markRevisions().map((r) => `${r.kind}:${r.id}`).join(",");
});

describe("docx 修订：格式改动（w:rPrChange）", () => {
  const formatRun = (doc: WordDocument): Inline => allRuns(doc).find((r) => r.text === "改过格式的字的")!;

  it("keeps the record when another paragraph is edited", () => {
    const doc = open(makeRevisedDocx());
    paragraphs(doc)[0]!.inlines = [{ text: "改动一处" }];
    expect(body(doc)).toContain('<w:rPrChange w:id="5" w:author="顾峰" w:date="2026-03-01T08:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>');
    expect(open(writeDocx(doc)).formatRevisions().map((r) => `format:${r.id}`)).toEqual(["format:5"]);
  });

  it("round-trips the signature and the before-state through a regenerated body", () => {
    const doc = open(makeRevisedDocx());
    paragraphs(doc)[0]!.inlines = [{ text: "改动一处" }];
    expect(JSON.stringify(open(writeDocx(doc)).formatRevisions())).toBe(JSON.stringify(doc.formatRevisions()));
  });

  it("writes the record inside the run's own properties, after them", () => {
    const doc = open(makeRevisedDocx());
    const run = formatRun(doc);
    run.bold = true;
    run.revFormat = { kind: "format", id: 77, author: "审阅者", date: "2026-04-01T00:00:00Z", prev: {} };
    expect(body(doc)).toContain('<w:rPr><w:b/><w:rPrChange w:id="77" w:author="审阅者" w:date="2026-04-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr>');
  });

  it("accepting keeps the new formatting and takes the record out", () => {
    const doc = open(makeRevisedDocx());
    expect(doc.resolveRevisions("accept", 5)).toBe(1);
    const run = formatRun(doc);
    expect(run.revFormat).toBeUndefined();
    expect(run.italic).toBeUndefined();
    expect(doc.formatRevisions()).toEqual([]);
    expect(body(doc)).not.toContain("rPrChange");
  });

  it("refusing puts the recorded properties back", () => {
    const doc = open(makeRevisedDocx());
    doc.resolveRevisions("reject", 5);
    const run = formatRun(doc);
    expect(run.italic).toBe(true);
    expect(run.revFormat).toBeUndefined();
    expect(body(doc)).not.toContain("rPrChange");
  });

  it("refusing clears a property the record does not mention, because the record holds the whole previous set", () => {
    const doc = open(makeRevisedDocx());
    formatRun(doc).bold = true;
    doc.resolveRevisions("reject", 5);
    const after = formatRun(doc);
    expect(after.italic).toBe(true);
    expect(after.bold).toBeUndefined();
  });

  it("deciding everything decides the formatting changes along with the insertions and deletions", () => {
    const doc = open(makeRevisedDocx());
    doc.resolveRevisions("accept");
    expect(doc.revisions()).toEqual([]);
    expect(doc.formatRevisions()).toEqual([]);
    expect(paragraphs(doc).map(textOf).join("|")).toContain("改过格式的字的");
  });

  it("leaves a formatting record alone while deciding somebody else's change", () => {
    const doc = open(makeRevisedDocx());
    doc.resolveRevisions("accept", 1);
    expect(doc.formatRevisions().map((r) => r.id).join()).toBe("5");
    expect(formatRun(doc).revFormat?.author).toBe("顾峰");
  });
});

describe("docx 修订：段落标记与段落格式（w:pPr/w:rPr、w:pPrChange）", () => {
  const marks = (doc: WordDocument): string => doc.markRevisions().map((r) => `${r.kind}:${r.id}`).join(",");
  /** The pane's own row: one per pending change of any kind, carrying the text it covers. */
  const entry = (doc: WordDocument, id: number) => doc.revisionEntries().find((e) => e.id === id);
  const paraFmtIds = (doc: WordDocument): string => doc.paragraphFormatRevisions().map((r) => r.id).join(",");
  const texts = (doc: WordDocument): string[] => paragraphs(doc).map(textOf);
  const paraHolding = (doc: WordDocument, text: string): Paragraph => paragraphs(doc).find((p) => textOf(p) === text)!;
  /** The break at the end of this paragraph, whoever signed it. */
  const markOf = (doc: WordDocument, text: string): Revision | undefined => paraHolding(doc, text)?.markRev;

  it("reads a paragraph mark revision as the break somebody added or removed", () => {
    const doc = open(makeParaMarkDocx());
    expect(markOf(doc, "回车分出的前半")).toMatchObject({ kind: "ins", id: 31, author: "顾峰", date: "2026-03-01T08:00:00Z" });
    expect(markOf(doc, "合并的上段")).toMatchObject({ kind: "del", id: 32, author: "林晓" });
    expect(marks(doc)).toBe("ins:31,del:32,del:34");
  });

  it("keeps a break out of the list of text insertions and deletions, which is what the pane shows for runs", () => {
    const doc = open(makeParaMarkDocx());
    expect(doc.revisions()).toEqual([]);
    expect(doc.formatRevisions()).toEqual([]);
  });

  it("reads a paragraph property change with the properties from before it", () => {
    const doc = open(makeParaMarkDocx());
    const p = paraHolding(doc, "改过对齐的段落");
    expect(p.align).toBe("center");
    expect(p.revFormat).toMatchObject({ kind: "format", id: 33, author: "顾峰", prev: { align: "left" } });
    expect(paraFmtIds(doc)).toBe("33");
  });

  it("saves a document nobody edited byte-for-byte", () => {
    const bytes = makeParaMarkDocx();
    expect(unstamped(writeDocx(open(bytes)))).toEqual(unstamped(bytes));
  });

  it("writes the mark as a property of the paragraph, not as a run of text", () => {
    const xml = body(open(makeParaMarkDocx()));
    expect(xml).toContain('<w:pPr><w:rPr><w:del w:id="32" w:author="林晓" w:date="2026-03-01T08:00:00Z"/></w:rPr></w:pPr><w:r>');
    expect(xml).not.toContain("<w:delText");
  });

  it("writes a paragraph property change after the current properties, and last", () => {
    expect(body(open(makeParaMarkDocx()))).toContain(
      '<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="33" w:author="顾峰" w:date="2026-03-01T08:00:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>',
    );
  });

  it("round-trips both paragraph records through a save and back", () => {
    const doc = open(makeParaMarkDocx());
    const back = open(writeDocx(doc));
    expect(marks(back)).toBe(marks(doc));
    expect(paraFmtIds(back)).toBe(paraFmtIds(doc));
    expect(back.paragraphFormatRevisions()[0]?.prev).toEqual(doc.paragraphFormatRevisions()[0]?.prev);
  });

  it("accepting a deleted break merges the paragraph with the one it was split from", () => {
    const doc = open(makeParaMarkDocx());
    expect(doc.resolveRevisions("accept", 32)).toBe(1);
    expect(texts(doc)).toContain("合并的上段合并的下段");
    expect(texts(doc)).not.toContain("合并的下段");
    expect(marks(doc)).toBe("ins:31,del:34");
  });

  it("refusing an inserted break takes the break back out the same way", () => {
    const doc = open(makeParaMarkDocx());
    expect(doc.resolveRevisions("reject", 31)).toBe(1);
    expect(texts(doc)).toContain("回车分出的前半回车分出的后半");
    expect(marks(doc)).toBe("del:32,del:34");
  });

  it("accepting an inserted break lets it stand, and only the record goes", () => {
    const doc = open(makeParaMarkDocx());
    expect(doc.resolveRevisions("accept", 31)).toBe(1);
    expect(texts(doc)).toContain("回车分出的前半");
    expect(texts(doc)).toContain("回车分出的后半");
    expect(markOf(doc, "回车分出的前半")).toBeUndefined();
  });

  it("refusing a deleted break keeps the break the reviewer wanted gone", () => {
    const doc = open(makeParaMarkDocx());
    expect(doc.resolveRevisions("reject", 32)).toBe(1);
    expect(texts(doc)).toContain("合并的上段");
    expect(texts(doc)).toContain("合并的下段");
    expect(markOf(doc, "合并的上段")).toBeUndefined();
  });

  it("merges paragraphs inside one cell and leaves the body paragraphs alone", () => {
    const doc = open(makeParaMarkDocx());
    expect(doc.resolveRevisions("accept", 34)).toBe(1);
    expect(texts(doc)).toContain("格内上段格内下段");
    expect(texts(doc).filter((t) => t.startsWith("回车")).length).toBe(2);
  });

  it("drops a mark record that has no break to join instead of losing the paragraph", () => {
    const doc = open(makeParaMarkDocx());
    const p = paraHolding(doc, "改过对齐的段落");
    p.markRev = { kind: "del", id: 99, author: "顾峰" };
    expect(doc.resolveRevisions("accept", 99)).toBe(1);
    expect(texts(doc)).toContain("改过对齐的段落");
    expect(marks(doc)).toBe("ins:31,del:32,del:34");
  });

  it("refusing a paragraph property change puts the recorded properties back", () => {
    const doc = open(makeParaMarkDocx());
    expect(doc.resolveRevisions("reject", 33)).toBe(1);
    const p = paraHolding(doc, "改过对齐的段落");
    expect(p.align).toBe("left");
    expect(p.revFormat).toBeUndefined();
    expect(body(doc)).not.toContain("pPrChange");
  });

  it("accepting a paragraph property change keeps the alignment the reviewer set", () => {
    const doc = open(makeParaMarkDocx());
    expect(doc.resolveRevisions("accept", 33)).toBe(1);
    expect(paraHolding(doc, "改过对齐的段落").align).toBe("center");
    expect(paraFmtIds(doc)).toBe("");
    expect(body(doc)).not.toContain("pPrChange");
  });

  it("clears a property the paragraph record does not mention, because it holds the whole previous set", () => {
    const doc = open(makeParaMarkDocx());
    const p = paraHolding(doc, "改过对齐的段落");
    p.spacingAfter = 120;
    doc.resolveRevisions("reject", 33);
    const after = paraHolding(doc, "改过对齐的段落");
    expect(after.align).toBe("left");
    expect(after.spacingAfter).toBeUndefined();
  });

  it("decides the paragraph records along with everything else, and loses no character", () => {
    const doc = open(makeParaMarkDocx());
    const everyChar = texts(doc).join("");
    expect(doc.resolveRevisions("accept")).toBe(4);
    expect(marks(doc)).toBe("");
    expect(paraFmtIds(doc)).toBe("");
    // Accepting stands both inserted breaks up and folds both deleted ones away, so two paragraphs
    // join their neighbours. No character moves anywhere else, so the text read whole is unchanged.
    expect(texts(doc).length).toBe(5);
    expect(texts(doc).join("")).toBe(everyChar);
  });

  it("leaves a paragraph record alone while deciding somebody else's change", () => {
    const doc = open(makeParaMarkDocx());
    doc.resolveRevisions("accept", 31);
    expect(paraFmtIds(doc)).toBe("33");
    expect(marks(doc)).toBe("del:32,del:34");
  });

  it("marks the revised break in the page and carries the record for the next edit", () => {
    const html = renderWord(open(makeParaMarkDocx()));
    expect(html).toContain('class="para rev-mark-del"');
    expect(html).toContain('data-rev-mark="del" data-rev-mark-id="32" data-rev-mark-author="林晓"');
    expect(html).toContain("删除段落标记");
    expect(html).toContain('class="para rev-mark-ins"');
  });

  it("marks a paragraph whose properties are revised, without inline style the editor would misread", () => {
    const html = renderWord(open(makeParaMarkDocx()));
    const tag = /<div class="para para-rev"[^>]*>/.exec(html)?.[0] ?? "";
    expect(tag).toContain('data-para-rev="33"');
    expect(tag).toContain("data-para-prev=");
    expect(tag).toContain("修改段落格式");
  });
});

describe("docx 修订：本机签在段落上的记录写出后再读仍是同一条", () => {
  const marks = (doc: WordDocument): string => doc.markRevisions().map((r) => `${r.kind}:${r.id}`).join(",");
  /** The pane's own row: one per pending change of any kind, carrying the text it covers. */
  const entry = (doc: WordDocument, id: number) => doc.revisionEntries().find((e) => e.id === id);
  const paraFmt = (doc: WordDocument): string => doc.paragraphFormatRevisions().map((r) => `pfmt:${r.id}`).join(",");
  /** The one `w:p` holding this exact text, as the writer put it in the file. */
  const paraXml = (doc: WordDocument, text: string): string =>
    (body(doc).match(/<w:p>(?:(?!<\/w:p>).)*?<\/w:p>/g) ?? []).find((p) => p.includes(`>${text}</w:t>`)) ?? "";
  const ours = { author: "审阅者甲", date: "2026-04-02T10:00:00Z" };

  it("writes a document nothing was signed on back with its paragraphs, alignment and somebody else's record", () => {
    const doc = open(makeBreakableDocx());
    const again = open(writeDocx(doc));
    expect(paragraphs(again).map(textOf).join("|")).toBe(paragraphs(doc).map(textOf).join("|"));
    expect(paragraphs(again).map((p) => p.align ?? "").join("|")).toBe(paragraphs(doc).map((p) => p.align ?? "").join("|"));
    expect(marks(again)).toBe("");
    expect(paraFmt(again)).toBe("pfmt:70");
    expect(paragraphs(again).find((p) => textOf(p) === "第二段文字")?.inlines[0]?.bold).toBe(true);
  });

  it("keeps a break the editor signed on the paragraph it ends, and gives each half the properties it was cut with", () => {
    const doc = open(makeBreakableDocx());
    const ins: Revision = { kind: "ins", id: 920, ...ours };
    const first = paragraphs(doc)[0]!;
    doc.blocks.splice(doc.blocks.indexOf(first), 1, { ...first, inlines: [{ text: "居中的" }], markRev: ins }, { ...first, inlines: [{ text: "第一段" }] });
    // `w:jc` then `w:rPr` then nothing: the order CT_PPr asks for, with the mark inside the run
    // properties of the paragraph the break terminates.
    expect(paraXml(doc, "居中的")).toBe(
      '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:ins w:id="920" w:author="审阅者甲" w:date="2026-04-02T10:00:00Z"/></w:rPr></w:pPr><w:r><w:t>居中的</w:t></w:r></w:p>',
    );
    const again = open(writeDocx(doc));
    expect(marks(again)).toBe("ins:920");
    // A break is not a change of characters, so the run pane shows nothing for it.
    expect(again.revisions()).toEqual([]);
    expect(paragraphs(again).slice(0, 2).map(textOf)).toEqual(["居中的", "第一段"]);
    expect(paragraphs(again)[0]?.align).toBe("center");
    expect(paragraphs(again)[1]?.align).toBe("center");
    expect(again.resolveRevisions("reject", 920)).toBe(1);
    expect(textOf(paragraphs(again)[0]!)).toBe("居中的第一段");
    expect(paragraphs(again).length).toBe(paragraphs(doc).length - 1);
  });

  it("holds both halves of a break whose deletion is still waiting, and decides it without moving a character", () => {
    const doc = open(makeBreakableDocx());
    const del: Revision = { kind: "del", id: 921, ...ours };
    const last = paragraphs(doc).find((p) => textOf(p) === "第四段文字")!;
    last.markRev = del;
    const again = open(writeDocx(doc));
    expect(marks(again)).toBe("del:921");
    expect(paragraphs(again).map(textOf).join("")).toBe(paragraphs(doc).map(textOf).join(""));
    // The last paragraph of a stretch has no neighbour to fold into, so deciding its break takes the
    // record and leaves every paragraph standing — the file already holds them as two.
    expect(again.resolveRevisions("accept", 921)).toBe(1);
    expect(marks(again)).toBe("");
    expect(paragraphs(again).map(textOf)).toEqual(paragraphs(doc).map(textOf));
  });

  it("decides a paragraph's break and its properties apart from one another", () => {
    const doc = open(makeBreakableDocx());
    const del: Revision = { kind: "del", id: 921, ...ours };
    const pfmt: ParagraphFormatRevision = { kind: "format", id: 922, ...ours, prev: { align: "left" } };
    const last = paragraphs(doc).find((p) => textOf(p) === "第四段文字")!;
    last.markRev = del;
    last.revFormat = pfmt;
    const again = open(writeDocx(doc));
    expect(marks(again)).toBe("del:921");
    expect(paraFmt(again)).toBe("pfmt:70,pfmt:922");
    expect(again.resolveRevisions("reject", 922)).toBe(1);
    expect(paraFmt(again)).toBe("pfmt:70");
    expect(marks(again)).toBe("del:921");
    expect(paragraphs(again).find((p) => textOf(p) === "第四段文字")?.align).toBe("left");
    expect(again.resolveRevisions("accept", 921)).toBe(1);
    expect(marks(again)).toBe("");
    expect(paraFmt(again)).toBe("pfmt:70");
  });

  it("keeps 林晓's paragraph record on its own paragraph while ours are signed beside it", () => {
    const doc = open(makeBreakableDocx());
    const ins: Revision = { kind: "ins", id: 920, ...ours };
    const held = paragraphs(doc).find((p) => textOf(p) === "有人改过对齐的段")!;
    held.markRev = ins;
    const again = open(writeDocx(doc));
    expect(marks(again)).toBe("ins:920");
    expect(paraFmt(again)).toBe("pfmt:70");
    expect(again.paragraphFormatRevisions()[0]?.author).toBe("林晓");
    expect(again.resolveRevisions("accept", 920)).toBe(1);
    expect(marks(again)).toBe("");
    expect(paraFmt(again)).toBe("pfmt:70");
  });
});

describe("docx 修订：待决项的统一有序清单（审阅窗格与 WPS Revisions 集合的接口）", () => {
  const rows = (doc: WordDocument): string => doc.revisionEntries().map((e) => `${e.index}.${e.type}:${e.id}:${e.author}:${e.text}`).join(",");
  const texts = (doc: WordDocument): string[] => paragraphs(doc).map(textOf);

  it("lists every kind of pending change once, in the order it sits in the document", () => {
    const doc = open(makeRevisedDocx());
    expect(rows(doc)).toBe(
      "1.insert:1:顾峰:插入的字,2.delete:2:林晓:删掉的字,3.insert:3:顾峰:加粗插入,4.insert:4:林晓:链接里的插入,5.run-format:5:顾峰:改过格式的字的,6.insert:6:顾峰:格内插入",
    );
    // The collection covers the same records the four separate lists do, no more and no fewer.
    expect(doc.revisionEntries().length).toBe(doc.revisions().length + doc.formatRevisions().length);
  });

  it("counts a revision that was nested in the file once for each signature it carries", () => {
    const doc = open(makeNestedRevisionDocx());
    expect(doc.revisionEntries().map((e) => `${e.index}:${e.type}:${e.id}`).join(",")).toBe("1:insert:11,2:insert:12,3:delete:13,4:delete:14");
  });

  it("shows a break as the paragraph mark it is and a property change as its paragraph's text", () => {
    const doc = open(makeParaMarkDocx());
    expect(rows(doc)).toBe(
      "1.insert-mark:31:顾峰:¶,2.delete-mark:32:林晓:¶,3.paragraph-format:33:顾峰:改过对齐的段落,4.delete-mark:34:顾峰:¶",
    );
  });

  it("decides the N-th change and renumbers what is left", () => {
    const doc = open(makeRevisedDocx());
    expect(doc.resolveRevisionAt("accept", 2)).toBe(true);
    expect(rows(doc)).toBe(
      "1.insert:1:顾峰:插入的字,2.insert:3:顾峰:加粗插入,3.insert:4:林晓:链接里的插入,4.run-format:5:顾峰:改过格式的字的,5.insert:6:顾峰:格内插入",
    );
    // Refusing the break, not the characters, is the same one-record decision.
    const brk = open(makeParaMarkDocx());
    expect(brk.resolveRevisionAt("reject", 1)).toBe(true);
    expect(brk.revisionEntries().map((e) => `${e.index}:${e.id}`).join(",")).toBe("1:32,2:33,3:34");
    expect(doc.resolveRevisionAt("accept", 99)).toBe(false);
  });

  it("decides one author's changes and leaves everybody else's waiting", () => {
    const doc = open(makeRevisedDocx());
    expect(doc.resolveRevisionsByAuthor("reject", "林晓")).toBe(2);
    expect(doc.revisionEntries().map((e) => e.author).join(",")).toBe("顾峰,顾峰,顾峰,顾峰");
    // Refusing 林晓's deletion puts her characters back and takes her insertion out with the record.
    expect(texts(doc).join("|")).toContain("删掉的字");
    expect(texts(doc).join("|")).not.toContain("链接里的插入");
    const other = open(makeRevisedDocx());
    expect(other.resolveRevisionsByAuthor("accept", "顾峰")).toBe(4);
    expect(other.revisionEntries().map((e) => `${e.id}`).join(",")).toBe("2,4");
  });

  it("numbers the rows by where the change sits, not by the id the file gave it", () => {
    // The paragraph record carries the lowest `w:id` here and the run records higher ones, so a pane
    // that counted by id would put 段落改动 first and the text ones after it.
    const doc = open(makeLowIdDocx());
    expect(rows(doc)).toBe("1.run-format:2:林晓:低号的格式,2.insert-mark:3:顾峰:¶,3.paragraph-format:1:顾峰:低号的对齐段");
    expect(doc.revisionEntries().map((e) => e.id).join(",")).toBe("2,3,1");
    expect(doc.revisionEntries().every((e) => !!e.date)).toBe(true);
  });

  it("knows the highest id in use, so a change signed by the editor cannot reuse one", () => {
    expect(open(makeLowIdDocx()).maxRevisionId()).toBe(3);
    expect(open(makeRevisedDocx()).maxRevisionId()).toBe(6);
    expect(open(makeParaMarkDocx()).maxRevisionId()).toBe(34);
    expect(open(makeBreakableDocx()).maxRevisionId()).toBe(70);
    expect(open(makeMovedDocx()).maxRevisionId()).toBe(22);
    const fresh = new Doc();
    expect(fresh.maxRevisionId()).toBe(0);
  });

  it("says which paragraph and which characters each row covers", () => {
    const doc = open(makeRevisedDocx());
    expect(doc.paragraphs().map(textOf).join("|")).toBe("未修订的段落|前插入的字删掉的字后|加粗的加粗插入|链接里的插入|改过格式的字的|格内插入|普通格");
    expect(doc.revisionEntries().map((e) => `${e.para}:${e.begin}-${e.end}`).join(",")).toBe("1:1-5,1:5-9,2:3-7,3:0-6,4:0-7,5:0-4");
  });

  it("answers which row covers a place, and the innermost one where changes overlap", () => {
    const doc = open(makeLowIdDocx());
    const ps = doc.paragraphs();
    // 低号的插入 carries no record of its own, so a caret there has nothing to decide.
    expect(doc.revisionIndexAt(ps[0]!, 0)).toBeUndefined();
    expect(doc.revisionIndexAt(ps[0]!, 7)).toBe(1);
    // At the end of the paragraph both the formatting record and the break reach the place; the break
    // is the smaller range, and it is the one a reviewer clicking on the pilcrow means.
    expect(doc.revisionIndexAt(ps[0]!, 10)).toBe(2);
    expect(doc.revisionIndexAt(ps[1]!, 0)).toBe(3);
    expect(doc.revisionIndexAt(ps[1]!, 6)).toBe(3);
    const other = open(makeParaMarkDocx());
    const os = other.paragraphs();
    expect(other.revisionIndexAt(os[0]!, 7)).toBe(1);
    expect(other.revisionIndexAt(os[1]!, 0)).toBeUndefined();
    expect(other.revisionIndexAt(os[4]!, 3)).toBe(3);
    expect(other.revisionIndexAt(os[5]!, 4)).toBe(4);
  });
});

describe("docx 修订：一条删除跨过段落边界（两个 w:del 共用一个 w:id）", () => {
  const rows = (doc: WordDocument): string => doc.revisions().map((r) => `${r.kind}:${r.id}`).join(",");
  const said = (doc: WordDocument): string[] => paragraphs(doc).map(textOf);
  const marks = (doc: WordDocument): string => doc.markRevisions().map((r) => `${r.kind}:${r.id}`).join(",");
  /** The pane's own row: one per pending change of any kind, carrying the text it covers. */
  const entry = (doc: WordDocument, id: number) => doc.revisionEntries().find((e) => e.id === id);

  it("lists it once, because one id is one thing for a reviewer to decide", () => {
    const doc = open(makeStraddleDocx());
    expect(rows(doc)).toBe("del:71,ins:72");
    expect(entry(doc, 71)?.text).toBe("段第");
  });

  it("keeps each half in the paragraph it was struck from, with that paragraph's own formatting", () => {
    const doc = open(makeStraddleDocx());
    const ps = paragraphs(doc);
    expect(ps[0]!.inlines.filter((r) => r.rev?.id === 71).map((r) => r.text).join("")).toBe("段");
    const lower = ps[1]!.inlines.find((r) => r.rev?.id === 71)!;
    expect(lower.text).toBe("第");
    expect(lower.bold).toBe(true);
  });

  it("saves a document nobody edited byte-for-byte", () => {
    const bytes = makeStraddleDocx();
    expect(unstamped(writeDocx(open(bytes)))).toEqual(unstamped(bytes));
  });

  it("writes the two elements back and reads them as the same one entry", () => {
    const doc = open(makeStraddleDocx());
    expect((body(doc).match(/<w:del w:id="71"/g) ?? []).length).toBe(2);
    const back = open(writeDocx(doc));
    expect(rows(back)).toBe(rows(doc));
    expect(entry(back, 71)?.text).toBe("段第");
    expect(paragraphs(back).map(textOf)).toEqual(paragraphs(doc).map(textOf));
  });

  it("accepting takes both halves out at once and leaves the rest of the file's records signed", () => {
    const doc = open(makeStraddleDocx());
    expect(doc.resolveRevisions("accept", 71)).toBeGreaterThan(1);
    expect(said(doc).slice(0, 2)).toEqual(["居中的第一换", "二段文字"]);
    expect(rows(doc)).toBe("ins:72");
    expect(marks(doc)).toBe("del:73");
    expect(doc.paragraphFormatRevisions().map((r) => r.id).join(",")).toBe("70");
  });

  it("refusing puts each character back in its own paragraph", () => {
    const doc = open(makeStraddleDocx());
    expect(doc.resolveRevisions("reject", 71)).toBeGreaterThan(1);
    expect(said(doc).slice(0, 2)).toEqual(["居中的第一段换", "第二段文字"]);
    expect(entry(doc, 71)).toBeUndefined();
    expect(rows(doc)).toBe("ins:72");
  });

});

describe("docx 修订：表格行自己的记录（w:trPr/w:ins、w:trPr/w:del）", () => {
  const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;
  const said = (doc: WordDocument): string[] => paragraphs(doc).map(textOf);
  const listed = (doc: WordDocument): string => doc.revisionEntries().map((e) => `${e.type}:${e.id}:${e.text}`).join(",");

  it("reads the row's own change, which sits on the row rather than in its cells", () => {
    const doc = open(makeTableRowDocx());
    const t = tableOf(doc);
    expect(t.rows.map((row) => row.map((c) => `${c.row?.rev?.kind ?? ""}:${c.row?.rev?.id ?? 0}`).join("|"))).toEqual([
      "del:31|del:31",
      ":0|:0",
      "ins:34|ins:34",
    ]);
    // Every cell of one row holds the same record: the file signs the row once, not each cell.
    expect(t.rows[0][0]!.row?.rev).toBe(t.rows[0][1]!.row?.rev);
  });

  it("lists the row before the text inside it, so a pane reads 加了行 then what the row holds", () => {
    const doc = open(makeTableRowDocx());
    expect(listed(doc)).toBe(
      "row-delete:31:删掉的行 它的第二格,delete:32:删掉的行,delete:33:它的第二格,row-insert:34:新增的行 格里的插入,insert:35:格里的插入",
    );
  });

  it("saves a table nobody edited byte-for-byte", () => {
    const bytes = makeTableRowDocx();
    expect(unstamped(writeDocx(open(bytes)))).toEqual(unstamped(bytes));
  });

  it("keeps each row's record through an edit made anywhere else in the document", () => {
    const doc = open(makeTableRowDocx());
    paragraphs(doc)[0]!.inlines[0]!.text = "改过的段落";
    const xml = body(doc);
    expect(doc.edited).toBe(true);
    expect(xml).toContain("<w:trPr><w:del w:id=\"31\"");
    expect(xml).toContain("<w:trPr><w:ins w:id=\"34\"");
    const back = open(writeDocx(doc));
    expect(listed(back)).toBe(listed(doc));
    expect(said(back)).toEqual(said(doc));
  });

  it("counting the numbers a row signs, so the editor never signs a new change with one", () => {
    expect(open(makeTableRowDocx()).maxRevisionId()).toBe(35);
  });

  it("accepting a deleted row takes the row away, its struck text with it", () => {
    const doc = open(makeTableRowDocx());
    expect(doc.resolveRevisions("accept", 31)).toBeGreaterThan(0);
    expect(tableOf(doc).rows).toHaveLength(2);
    expect(listed(doc)).toBe("row-insert:34:新增的行 格里的插入,insert:35:格里的插入");
  });

  it("refusing a deleted row keeps the row and leaves its cells' own marks pending", () => {
    const doc = open(makeTableRowDocx());
    doc.resolveRevisions("reject", 31);
    expect(tableOf(doc).rows).toHaveLength(3);
    expect(tableOf(doc).rows[0][0]!.row?.rev).toBeUndefined();
    // The cell text carries separate w:ids the file gave it, and deciding the row does not decide them.
    expect(listed(doc)).toBe("delete:32:删掉的行,delete:33:它的第二格,row-insert:34:新增的行 格里的插入,insert:35:格里的插入");
  });

  it("refusing an inserted row takes the row back out, accepting lets it stand", () => {
    const refused = open(makeTableRowDocx());
    refused.resolveRevisions("reject", 34);
    expect(refused.revisionEntries().map((e) => e.id).join(",")).toBe("31,32,33");
    const accepted = open(makeTableRowDocx());
    accepted.resolveRevisions("accept", 34);
    expect(tableOf(accepted).rows).toHaveLength(3);
    expect(tableOf(accepted).rows[2][0]!.row?.rev).toBeUndefined();
    expect(listed(accepted)).toBe("row-delete:31:删掉的行 它的第二格,delete:32:删掉的行,delete:33:它的第二格,insert:35:格里的插入");
  });

  it("accepting every change leaves the table as the document would be without them", () => {
    const doc = open(makeTableRowDocx());
    doc.resolveRevisions("accept");
    expect(doc.revisionEntries()).toEqual([]);
    expect(tableOf(doc).rows.map((r) => r.map((c) => c.paragraphs.map((p) => (p.kind === "paragraph" ? textOf(p) : "")).join("")).join(" "))).toEqual(["留下的行 普通格", "新增的行 格里的插入"]);
  });

  it("refusing every change puts the deleted row back and drops the one that was added", () => {
    const doc = open(makeTableRowDocx());
    doc.resolveRevisions("reject");
    expect(doc.revisionEntries()).toEqual([]);
    expect(tableOf(doc).rows.map((r) => r[0]!.paragraphs.map((p) => (p.kind === "paragraph" ? textOf(p) : "")).join(""))).toEqual(["删掉的行", "留下的行"]);
    expect(body(doc)).not.toContain("<w:trPr>");
  });

  it("a table whose last row is accepted away is no longer a table", () => {
    const doc = open(makeLastRowDeletedDocx());
    expect(doc.resolveRevisions("accept", 41)).toBeGreaterThan(0);
    expect(doc.blocks.some((b) => b.kind === "table")).toBe(false);
    const xml = body(doc);
    expect(xml).not.toContain("<w:tbl>");
    expect(said(open(writeDocx(doc)))).toEqual(["唯一的行被删了"]);
  });

  it("renders the row's change as a class and data attributes, never as a style", () => {
    const doc = open(makeTableRowDocx());
    const html = renderWord(doc);
    expect(html).toContain('<tr class="rev-row rev-del" data-row-rev="del" data-row-rev-id="31" data-row-rev-author="顾峰"');
    expect(html).toContain('<tr class="rev-row rev-ins" data-row-rev="ins" data-row-rev-id="34"');
    expect(html).not.toContain("<tr style");
  });
});
