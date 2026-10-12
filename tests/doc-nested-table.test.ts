import { describe, expect, it } from "vitest";
import type { Table } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { splitTableAt } from "../src/core/doc-table-ops.js";

import { OpcPackage, text } from "../src/core/opc.js";
import { makeDeepNestedTableDocx, makeNestedRevisedTableDocx, makeNestedTableDocx } from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;
/** The table a cell holds. A nested table is a block beside the cell's paragraphs, not inside one of them. */
const nestedIn = (t: Table, row: number, col: number): Table | undefined =>
  t.rows[row]?.[col]?.paragraphs.find((b) => b.kind === "table") as Table | undefined;
const textOf = (t: Table): string => t.rows.map((r) => r.map((c) => c.paragraphs.map(blockText).join("")).join("|")).join("/");
const blockText = (b: Table["rows"][number][number]["paragraphs"][number]): string =>
  b.kind === "paragraph" ? b.inlines.map((i) => i.text).join("") : b.kind === "table" ? textOf(b) : "";

/**
 * The element path each `w:tbl` was written under, outermost first, as the file stands. A stack scan because
 * the claim is about **parenthood**: `w:tbl` belongs among a cell's block children, and a `w:tbl` that ended
 * up inside a `w:p` would be a part Word refuses to open even though it still round-trips through us.
 */
const tableAncestry = (xml: string): string[][] => {
  const out: string[][] = [];
  const stack: string[] = [];
  for (const m of xml.matchAll(/<(\/?)([\w:]+)([^>]*?)(\/?)>/g)) {
    const [, close, name, , selfClose] = m;
    if (close) {
      stack.pop();
      continue;
    }
    if (name === "w:tbl") out.push([...stack]);
    if (!selfClose && !/^<\?/.test(m[0])) stack.push(name);
  }
  return out;
};

describe("docx 表中之表：三层嵌套的读取、写出形状与往返", () => {
  it("reads three levels of tables, each level keeping its own text", () => {
    const outer = tableOf(open(makeDeepNestedTableDocx()));
    expect(outer.rows.length).toBe(2);
    const middle = nestedIn(outer, 0, 1);
    expect(middle && textOf(middle)).toBe("中一深一|深二/深三|深四|中二/中三|中四");
    const deepest = middle && nestedIn(middle, 0, 0);
    expect(deepest && textOf(deepest)).toBe("深一|深二/深三|深四");
    // Reading deeper must not smear the levels: the outer cell holds one paragraph and one table, in that order.
    expect(outer.rows[0][1]!.paragraphs.map((b) => b.kind)).toEqual(["paragraph", "table"]);
    expect(textOf(outer)).toBe("外一|外二中一深一|深二/深三|深四|中二/中三|中四/外三|外四");
  });

  it("writes every nested table as a direct child of its cell, never inside a paragraph", () => {
    const ancestry = tableAncestry(saved(open(makeDeepNestedTableDocx())));
    expect(ancestry.length).toBe(3);
    expect(ancestry[0]).toEqual(["w:document", "w:body"]);
    expect(ancestry[1]).toEqual(["w:document", "w:body", "w:tbl", "w:tr", "w:tc"]);
    expect(ancestry[2]).toEqual(["w:document", "w:body", "w:tbl", "w:tr", "w:tc", "w:tbl", "w:tr", "w:tc"]);
    // The same claim in one line: no `w:tbl` may have a `w:p` among its ancestors.
    expect(ancestry.some((chain) => chain.includes("w:p"))).toBe(false);
  });

  it("keeps the three levels, their widths and their block order through a save", () => {
    const doc = open(makeDeepNestedTableDocx());
    const again = open(writeDocx(doc));
    const middle = nestedIn(tableOf(again), 0, 1);
    const deepest = middle && nestedIn(middle, 0, 0);
    expect(tableOf(again).rows.map((r) => r.length)).toEqual([2, 2]);
    expect(middle && middle.rows[0][0]!.widthTwips).toBe(600);
    expect(deepest && textOf(deepest)).toBe("深一|深二/深三|深四");
    expect(deepest && deepest.rows[0][0]!.widthTwips).toBe(300);
    expect(deepest && deepest.rows[0][0]!.paragraphs.map((b) => b.kind)).toEqual(["paragraph"]);
  });

  it("saves the same bytes twice: the second write has nothing new to invent", () => {
    const once = saved(open(makeDeepNestedTableDocx()));
    const twice = saved(open(writeDocx(open(makeDeepNestedTableDocx()))));
    expect(twice).toBe(once);
  });

  it("holds a two-by-two inside a two-by-two without the outer one noticing", () => {
    const outer = tableOf(open(makeNestedTableDocx()));
    expect(outer.rows.length).toBe(2);
    expect(outer.rows.map((r) => r.length)).toEqual([2, 2]);
    const inner = nestedIn(outer, 0, 1);
    expect(inner && textOf(inner)).toBe("内一|内二/内三|内四");
    expect(inner && inner.columns).toBe(2);
    // Nothing the outer table holds is a copy of what the inner one holds.
    expect(outer.rows[0][0]!.paragraphs.length).toBe(1);
  });
});

describe("docx 表中之表：别人写在内层表上的修订", () => {
  it("lists a row record and an inserted run from two tables deep, placed at the inner paragraph", () => {
    const doc = open(makeNestedRevisedTableDocx());
    expect(doc.revisionEntries().map((e) => `${e.index}.${e.type}:${e.id}:${e.author}:${e.text}`)).toEqual([
      "1.row-insert:88:顾峰:内三 内四",
      "2.insert:89:顾峰:内三",
    ]);
    // `para` counts the paragraphs the way the pane walks them, cells and nested tables included:
    // 表前,外一,外二,内一,内二 then the inner row's first paragraph.
    const entry = doc.revisionEntries()[0]!;
    expect(doc.paragraphs()[entry.para]).toBe(doc.paragraphs()[5]);
    expect(entry.para).toBe(5);
    const inner = doc.paragraphs()[5]!;
    expect(doc.revisionIndexAt(inner, 1)).toBe(2);
  });

  it("refuses to decide a record it was not given: an unknown id changes nothing", () => {
    const doc = open(makeNestedRevisedTableDocx());
    expect(doc.resolveRevisions("reject", 9999)).toBe(0);
    expect(nestedIn(tableOf(doc), 0, 1)!.rows.length).toBe(2);
  });

  it("rejecting the inner row record takes the inner row out and leaves the outer table alone", () => {
    const doc = open(makeNestedRevisedTableDocx());
    expect(doc.resolveRevisions("reject", 88)).toBe(1);
    const inner = nestedIn(tableOf(doc), 0, 1)!;
    expect(textOf(inner)).toBe("内一|内二");
    expect(tableOf(doc).rows.length).toBe(2);
    const xml = saved(doc);
    expect(xml.split("<w:trPr><w:ins ").length - 1).toBe(0);
    // The inserted run still waiting in that row went with it: its text was the row's, so the pane has nothing
    // left to offer and the file has nothing left to carry.
    expect(doc.revisionEntries().map((e) => e.id)).toEqual([]);
    expect(xml.includes("<w:ins ")).toBe(false);
  });

  it("accepting the inner row record keeps the row and drops only the record", () => {
    const doc = open(makeNestedRevisedTableDocx());
    expect(doc.resolveRevisions("accept", 88)).toBe(1);
    expect(textOf(nestedIn(tableOf(doc), 0, 1)!)).toBe("内一|内二/内三|内四");
    expect(saved(doc).split("<w:trPr><w:ins ").length - 1).toBe(0);
    // The text record in the same row is somebody else's decision and stays waiting.
    expect(doc.revisionEntries().map((e) => e.id)).toEqual([89]);
  });

  it("carries both records out to the file when nobody decides", () => {
    const reopened = open(writeDocx(open(makeNestedRevisedTableDocx())));
    expect(reopened.revisionEntries().map((e) => `${e.type}:${e.id}`)).toEqual(["row-insert:88", "insert:89"]);
    expect(reopened.paragraphs()[5]!.inlines.map((i) => i.text).join("")).toBe("内三");
  });
});

describe("docx 表中之表：我们写出的格子不能以表收尾", () => {
  /** Split the inner table in two and put the halves in the cell the way 拆分表格 does. */
  const splitInner = (doc: WordDocument): void => {
    const cell = tableOf(doc).rows[0][1]!;
    const inner = cell.paragraphs[1] as Table;
    const split = splitTableAt(inner, 1);
    if (!split.ok) throw new Error(split.reason);
    cell.paragraphs.splice(2, 0, { kind: "paragraph", inlines: [] }, split.tail);
  };

  it("拆过内层表再存：那一格以段落收尾，件里没有 `</w:tbl></w:tc>`", () => {
    const doc = open(makeNestedTableDocx());
    splitInner(doc);
    const xml = saved(doc);
    // Word 一族的读法把"格子以表收尾"当损坏，两个段落夹一张表才是一张正常的格子。补段落在写出这一步做，
    // 模型里那格原本有什么不变；一份没人改过的件仍按原字节存，补的是我们重算出来的那部分。
    expect(xml.split("</w:tbl></w:tc>").length - 1).toBe(0);
    expect(xml.split("</w:tbl><w:p/></w:tc>").length - 1).toBe(1);
  });

  it("补上的段落不会越补越多：读回来改点文字再存，还是那一个", () => {
    const doc = open(makeNestedTableDocx());
    splitInner(doc);
    const back = open(writeDocx(doc));
    expect(back.blocks.length).toBe(doc.blocks.length);
    // 段落进了件子，再读回来时它已在模型里，写出就不再补第二个（改动只挑一处正文文字，为的是让它重算）。
    const touched = back.paragraphs()[0]!;
    touched.inlines.push({ text: "改过" });
    const xml = saved(back);
    // 这次那格自己带着收尾的那个段落（是上一轮写进件里的），所以补不出第二个来。
    expect(xml.split("</w:tbl></w:tc>").length - 1).toBe(0);
    expect(xml.includes("</w:tbl><w:p></w:p></w:tc>")).toBe(true);
    expect(xml.includes("<w:p/><w:p/></w:tc>")).toBe(false);
    expect(tableOf(back).rows[0][1]!.paragraphs.map((b) => b.kind).join(",")).toBe("paragraph,table,paragraph,table,paragraph");
  });

  it("两张表并排坐在同一格里时，中间补一个段落", () => {
    const doc = open(makeNestedTableDocx());
    const cell = tableOf(doc).rows[0][1]!;
    const inner = cell.paragraphs[1] as Table;
    // 模型里真能出现并排的两张表 —— 外来件就是这么写的，我们没有段落夹着它。
    cell.paragraphs.splice(2, 0, structuredClone(inner));
    expect(saved(doc).includes("</w:tbl><w:p/><w:tbl>")).toBe(true);
  });
});
