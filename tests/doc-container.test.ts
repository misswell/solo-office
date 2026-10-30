import { describe, expect, it } from "vitest";
import type { Block, Carried, Inline, Table } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeContainerInHeaderDocx, makeContainersDocx, makeContainersInTableDocx, makeContainerWithCommentDocx, makeFieldCharDocx, makeGoBackDocx } from "./helpers/make-containers.js";
import { makeTableFidelityDocx } from "./helpers/make-tables.js";

const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);

/** One keystroke somewhere else: the only path that shows whether a structure was really kept. */
const edited = (bytes: Uint8Array, name = "改过的段落"): Uint8Array => {
  const doc = WordDocument.openBytes(bytes);
  const p = doc.blocks.find((b) => b.kind === "paragraph");
  if (p?.kind === "paragraph" && p.inlines[0]) p.inlines[0].text = name;
  return writeDocx(doc);
};

const carried = (doc: WordDocument): Carried[] => doc.blocks.filter((b): b is Carried => b.kind === "carried");

/** The shapes here are ECMA-derived: no file in the WPS corpus holds a container at all (`.scratch/sdt-scan.mjs`). */
describe("docx body containers", () => {
  it("reads each block container as a block of its own, where it sat", () => {
    const doc = WordDocument.openBytes(makeContainersDocx());
    expect(doc.blocks.map((b) => b.kind)).toEqual(["paragraph", "carried", "carried", "paragraph", "paragraph", "paragraph"]);
    expect(carried(doc).map((b) => b.tag)).toEqual(["sdt", "customXml"]);
  });

  it("shows what a reader would see inside, and not the field code", () => {
    const [sdt, customXml] = carried(WordDocument.openBytes(makeContainersDocx()));
    expect(sdt?.text).toBe("第一章 概览 第二章 细节");
    expect(customXml?.text).toBe("自定义标记包着的一段");
  });

  it("saves the container back byte for byte after a keystroke elsewhere", () => {
    const original = part(makeContainersDocx());
    const block = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(original)?.[0] ?? "";
    expect(block).toContain('w:alias w:val="目录标题"');
    const after = saved(WordDocument.openBytes(edited(makeContainersDocx())));
    expect(after).toContain(block);
    expect(after).toContain('w:itemName="样张标记"');
    expect(after).toContain('w:name="owner"');
  });

  it("costs nothing to save twice: the second file has the same containers as the first", () => {
    const once = edited(makeContainersDocx());
    const twice = WordDocument.openBytes(once);
    expect(part(edited(once))).toBe(saved(twice));
    expect(carried(WordDocument.openBytes(once)).map((b) => b.tag)).toEqual(["sdt", "customXml"]);
  });

  it("keeps a container in a table cell among the cell's paragraphs, and its properties as properties", () => {
    const table = WordDocument.openBytes(makeContainersInTableDocx()).blocks.find((b) => b.kind === "table") as Table;
    const cell = table.rows[0]![0]!;
    expect(cell.paragraphs.map((b) => b.kind)).toEqual(["paragraph", "carried"]);
    expect(cell.tcPrXml).toContain('<w:tcW w:w="3000"');
    const after = saved(WordDocument.openBytes(edited(makeContainersInTableDocx())));
    expect(after).toContain('<w:sdtPr><w:alias w:val="格内控件"/><w:tag w:val="CellSdt"/></w:sdtPr>');
    expect(after).toContain("格内控件的字");
  });

  it("is not confused by a comment between two of the body's children", () => {
    const doc = WordDocument.openBytes(makeContainerWithCommentDocx());
    expect(doc.blocks.map((b) => b.kind)).toEqual(["paragraph", "carried", "paragraph"]);
    const after = saved(WordDocument.openBytes(edited(makeContainerWithCommentDocx(), "改过的段落")));
    expect(after).toContain('w:alias w:val="带注释的控件"');
    expect(after).toContain("控件里的字");
    expect(after.indexOf("控件里的字")).toBeGreaterThan(after.indexOf("注释前的段落"));
    expect(after.indexOf("注释后的段落")).toBeGreaterThan(after.indexOf("控件里的字"));
  });

  it("draws the container on the page, read-only", () => {
    const html = renderWord(WordDocument.openBytes(makeContainersDocx()));
    expect(html).toContain('class="doc-carry"');
    expect(html).toContain('contenteditable="false"');
    // 盒子的孩子枚枚是段（#236），逐行画成两枚 .carry-p；字不变，只是从「一整段流文」换成了「一行一条」。
    expect(html).toContain('<div class="carry-p">第一章 概览</div>');
    expect(html).toContain('<div class="carry-p">第二章 细节</div>');
  });

  it("adds nothing to a document that has no containers in it", () => {
    const doc = WordDocument.openBytes(makeTableFidelityDocx());
    expect(doc.blocks.some((b) => b.kind === "carried")).toBe(false);
    expect(saved(doc)).not.toContain("doc-carry");
  });

  it("reads a container in the header part into the band, and leaves its bytes alone", () => {
    const bytes = makeContainerInHeaderDocx();
    const headerOf = (b: Uint8Array): string => text(OpcPackage.from(b).bytes("word/header1.xml")!);
    const doc = WordDocument.openBytes(bytes);
    // A control that only names itself is one the band now reads through, so its words are the header's own
    // and show on the page. Nothing was edited there, so the part still travels exactly as the file wrote it.
    expect(doc.header?.editable).toBe(true);
    expect(doc.header?.paragraphs).toEqual(["页眉的普通段", "页眉控件的字"]);
    const out = edited(bytes, "改过的正文");
    expect(headerOf(out)).toBe(headerOf(bytes));
    expect(headerOf(out)).toContain('w:alias w:val="页眉控件"');
    expect(headerOf(out)).toContain("页眉控件的字");
    expect(part(out)).toContain("改过的正文");
  });
});

const countOf = (haystack: string, needle: string): number => haystack.split(needle).length - 1;

/** Every inline in body order, table cells included — a bookmark inside a cell is still the document's. */
function eachInline(blocks: Block[], doIt: (inline: Inline) => void): void {
  for (const block of blocks) {
    if (block.kind === "table") {
      block.rows.forEach((row) => row.forEach((cell) => eachInline(cell.paragraphs, doIt)));
      continue;
    }
    if (block.kind === "paragraph") block.inlines.forEach(doIt);
  }
}

/** What the page holds as `tag|what a reader would see`, in body order. */
function inlineCarries(doc: WordDocument): string[] {
  const out: string[] = [];
  eachInline(doc.blocks, (inline) => {
    if (inline.carry) out.push(`${inline.carry.tag}|${inline.carry.show}`);
  });
  return out;
}

/**
 * The run-level shapes a real producer writes between its runs. None of these have a field in this model,
 * and a save that reads the paragraph's runs but not its neighbours throws away the reader's page number,
 * the date a control holds, and every place a cross-reference points to.
 */
describe("docx inline structures", () => {
  it("reads a bookmark, a field and an inline control as carries beside the runs", () => {
    expect(inlineCarries(WordDocument.openBytes(makeContainersDocx()))).toEqual([
      "sdt|2026 年 10 月 11 日",
      "fldSimple|12",
      "bookmarkStart|",
      "bookmarkEnd|",
    ]);
  });

  it("hands a bookmark start the w:name it came in with — the one thing 「定位」 addresses it by", () => {
    const startsOf = (doc: WordDocument): (string | undefined)[] => {
      const out: (string | undefined)[] = [];
      eachInline(doc.blocks, (inline) => {
        if (inline.carry?.tag === "bookmarkStart") out.push(inline.carry.name);
      });
      return out;
    };
    expect(startsOf(WordDocument.openBytes(makeContainersDocx()))).toEqual(["_Toc_1"]);
    expect(startsOf(WordDocument.openBytes(makeContainersInTableDocx()))).toEqual(["_CellMark"]);
    expect(startsOf(WordDocument.openBytes(makeGoBackDocx()))).toEqual(["_GoBack"]);
  });

  it("writes no name onto a structure that has none, and none into the saved bytes", () => {
    const doc = WordDocument.openBytes(makeContainersDocx());
    eachInline(doc.blocks, (inline) => {
      if (inline.carry && inline.carry.tag !== "bookmarkStart") expect(inline.carry.name).toBeUndefined();
    });
    // The name is a read, not a spelling: what goes back is the bytes, so the paragraph still writes its own spelling.
    const after = saved(doc);
    expect(after).toContain('<w:bookmarkStart w:id="0" w:name="_Toc_1"/>');
  });

  it("saves each inline structure back byte for byte after a keystroke", () => {
    const original = part(makeContainersDocx());
    const spans = [
      /<w:fldSimple[\s\S]*?<\/w:fldSimple>/.exec(original)![0]!,
      /<w:bookmarkStart[^>]*>/.exec(original)![0]!,
      /<w:bookmarkEnd[^>]*>/.exec(original)![0]!,
      /<w:sdt><w:sdtPr><w:alias w:val="日期"[\s\S]*?<\/w:sdt>/.exec(original)![0]!,
    ];
    const after = saved(WordDocument.openBytes(edited(makeContainersDocx())));
    for (const span of spans) expect(after).toContain(span);
  });

  it("keeps a bookmark stretched around the text it marked, not beside the paragraph", () => {
    const after = saved(WordDocument.openBytes(edited(makeContainersDocx())));
    expect(after.indexOf("bookmarkStart")).toBeLessThan(after.indexOf("书签跨过的字"));
    expect(after.indexOf("bookmarkEnd")).toBeGreaterThan(after.indexOf("书签跨过的字"));
  });

  it("writes each marker once, and the same bytes on the second save", () => {
    const once = edited(makeContainersDocx());
    const twice = edited(once);
    expect(countOf(part(twice), "<w:bookmarkStart")).toBe(1);
    expect(countOf(part(twice), "<w:fldSimple")).toBe(1);
    expect(part(twice)).toBe(saved(WordDocument.openBytes(twice)));
  });

  it("keeps a _GoBack pair alone in a paragraph that somebody types into", () => {
    const doc = WordDocument.openBytes(makeGoBackDocx());
    expect(inlineCarries(doc)).toEqual(["bookmarkStart|", "bookmarkEnd|"]);
    const first = doc.blocks[0];
    if (first?.kind !== "paragraph") throw new Error("the fixture's first block is not a paragraph");
    first.inlines.push({ text: "打进去的字" });
    const after = saved(doc);
    expect(after).toContain('<w:bookmarkStart w:id="0" w:name="_GoBack"/>');
    expect(after).toContain('<w:bookmarkEnd w:id="0"/>');
    expect(after).toContain("打进去的字");
    expect(after.indexOf("_GoBack")).toBeLessThan(after.indexOf("打进去的字"));
  });

  it("keeps a bookmark pair inside a table cell while the cell is edited", () => {
    const doc = WordDocument.openBytes(makeContainersInTableDocx());
    expect(inlineCarries(doc)).toEqual(["bookmarkStart|", "bookmarkEnd|"]);
    const after = saved(WordDocument.openBytes(edited(makeContainersInTableDocx(), "改过的段落")));
    expect(after).toContain('<w:bookmarkStart w:id="3" w:name="_CellMark"/>');
    expect(after).toContain('<w:bookmarkEnd w:id="3"/>');
    expect(after).toContain("格里的普通段");
  });

  it("reads a compound field's machinery as carries and its cached result as text", () => {
    const doc = WordDocument.openBytes(makeFieldCharDocx());
    expect(inlineCarries(doc)).toEqual(["proofErr|", "fldChar|", "instrText|", "fldChar|", "fldChar|", "proofErr|"]);
    const words: string[] = [];
    eachInline(doc.blocks, (inline) => words.push(inline.text));
    expect(words.filter((w) => w !== "").join("")).toBe("页码 7域后面的段落");
  });

  it("keeps every part of a field built from runs when its own paragraph is rewritten", () => {
    const after = saved(WordDocument.openBytes(edited(makeFieldCharDocx(), "改过的页码段")));
    expect(countOf(after, "<w:fldChar")).toBe(3);
    expect(countOf(after, "<w:instrText")).toBe(1);
    expect(countOf(after, "<w:proofErr")).toBe(2);
    expect(after).toContain("PAGE   \\* MERGEFORMAT");
    for (const piece of ['w:fldCharType="begin"', "instrText", 'w:fldCharType="separate"', ">7<", 'w:fldCharType="end"']) {
      expect(after.indexOf(piece)).toBeGreaterThan(after.indexOf("改过的页码段"));
    }
    expect(after.indexOf('w:fldCharType="begin"')).toBeLessThan(after.indexOf("instrText"));
    expect(after.indexOf("instrText")).toBeLessThan(after.indexOf('w:fldCharType="separate"'));
    expect(after.indexOf(">7<")).toBeLessThan(after.indexOf('w:fldCharType="end"'));
  });

  it("shows a field's result on the page and never its instruction", () => {
    const html = renderWord(WordDocument.openBytes(makeFieldCharDocx()));
    expect(html).toContain(">7<");
    expect(html).not.toContain("MERGEFORMAT");
    expect(countOf(html, 'class="run-carry"')).toBe(6);
  });

  it("draws them on the page read-only, showing what a reader would see", () => {
    const html = renderWord(WordDocument.openBytes(makeContainersDocx()));
    expect(countOf(html, 'class="run-carry"')).toBe(4);
    expect(html).toContain('data-tag="fldSimple"');
    expect(html).toContain(">12<");
    expect(html).toContain("2026 年 10 月 11 日");
    expect(html).toContain('contenteditable="false"');
  });

  it("adds nothing to a document that has no inline structures in it", () => {
    const doc = WordDocument.openBytes(makeTableFidelityDocx());
    expect(inlineCarries(doc)).toEqual([]);
    expect(renderWord(doc)).not.toContain("run-carry");
  });
});
