import { describe, expect, it } from "vitest";
import type { Inline, Paragraph, Table, TableCell, TableStyle } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { tableToText } from "../src/core/doc-table-ops.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);

const para = (...inlines: Inline[]): Paragraph => ({ kind: "paragraph", inlines });
const cell = (width: number, ...paragraphs: Paragraph[]): TableCell => ({ paragraphs, widthTwips: width });
const cellText = (p: Paragraph): string => p.inlines.map((i) => (i.tab ? "\\t" : i.text)).join("");
const lineTexts = (paras: Paragraph[]): string[] => paras.map(cellText);

/** A rows-by-columns grid whose cells all say which row and column they sit in, one paragraph each. */
const grid = (rows: number, cols: number, width = 1000): Table => ({
  kind: "table",
  columns: cols,
  rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_2, c) => cell(width, para({ text: `${r}${c}` })))),
  widthTwips: width * cols,
  layout: "fixed",
  columnWidths: Array.from({ length: cols }, () => width),
});

const ok = (t: Table, sep: "tab" | "comma" | "paragraph"): Paragraph[] => {
  const made = tableToText(t, sep);
  if (!made.ok) throw new Error(made.reason);
  return made.paragraphs;
};

/**
 * 表格转文本（`CT_TableTools.kuip:89` 的 `ConvertTabletoText`）。三档分隔各写出什么、多段格怎么摊开、
 * 格式跟谁走、什么形状拒转，都在这几层上钉住；页面上那一趟点击在真 Chrome 探针与装机烟测里。
 */
describe("docx 表格转文本", () => {
  it("制表符档：一行并成一段，格与格之间是 w:tab", () => {
    const out = ok(grid(2, 2), "tab");
    expect(out).toHaveLength(2);
    expect(lineTexts(out)).toEqual(["00\\t01", "10\\t11"]);
    expect(out[0]!.inlines.filter((i) => i.tab)).toHaveLength(1);
    expect(out[0]!.inlines.every((i) => !i.tab || i.text === "")).toBe(true);
  });

  it("逗号档：同一行用一枚逗号文字接起来", () => {
    const out = ok(grid(2, 3), "comma");
    expect(lineTexts(out)).toEqual(["00,01,02", "10,11,12"]);
    expect(out[0]!.inlines.some((i) => i.tab)).toBe(false);
  });

  it("段落标记档：一格一段，谁也不并", () => {
    const out = ok(grid(2, 2), "paragraph");
    expect(out).toHaveLength(4);
    expect(lineTexts(out)).toEqual(["00", "01", "10", "11"]);
  });

  it("格里的第二段起另起一行，不挤进第一行", () => {
    const t: Table = {
      kind: "table",
      columns: 2,
      rows: [[cell(1000, para({ text: "a" }), para({ text: "a2" })), cell(1000, para({ text: "b" }))]],
    };
    const out = ok(t, "tab");
    expect(lineTexts(out)).toEqual(["a\\tb", "a2"]);
  });

  it("一行的段格式跟行首那一格走，每枚 run 的格式各自留着", () => {
    const bold: Inline = { text: "甲", rPrSlots: [{ tag: "b", xml: "<w:b/>" }] };
    const t: Table = {
      kind: "table",
      columns: 2,
      rows: [[{ paragraphs: [{ kind: "paragraph", styleId: "a3", inlines: [bold] }], widthTwips: 1000 }, cell(1000, para({ text: "乙" }))]],
    };
    const out = ok(t, "tab");
    expect(out[0]!.styleId).toBe("a3");
    expect(out[0]!.inlines[0]!.rPrSlots).toEqual([{ tag: "b", xml: "<w:b/>" }]);
    // The source table is left alone: conversion reads, it does not move the table out from under the caller.
    expect(t.rows[0]![0]!.paragraphs).toHaveLength(1);
    expect(lineTexts(out)).toEqual(["甲\\t乙"]);
  });

  it("没行的表也留一行：槽位不能空着", () => {
    const out = ok({ kind: "table", columns: 1, rows: [] }, "tab");
    expect(out).toHaveLength(1);
    expect(lineTexts(out)).toEqual([""]);
  });

  it("格里有嵌套表、表行之间夹着东西、或表上有待判属性改动时拒转并说明理由", () => {
    const nested: Table = {
      kind: "table",
      columns: 1,
      rows: [[{ paragraphs: [{ kind: "table", columns: 1, rows: [] }], widthTwips: 1000 }]],
    };
    expect(tableToText(nested, "tab")).toMatchObject({ ok: false, reason: "格子里有嵌套表或内容控件，转成文本会把它整个丢掉" });
    const between = { ...grid(1, 1), betweenRows: [{ afterRow: 0, tag: "bookmarkEnd", xml: '<w:bookmarkEnd w:id="0"/>' }] };
    expect(tableToText(between, "tab")).toMatchObject({ ok: false, reason: "表行之间夹着书签一类的东西，转成文本会把它丢掉" });
    const revised = { ...grid(1, 1), revFormat: { kind: "format" as const, tag: "tblPrChange" as const, id: 1, author: "甲", prevSlots: [], prev: {} as TableStyle } };
    expect(tableToText(revised, "tab")).toMatchObject({ ok: false, reason: "这张表上有一条待判的属性改动，转成文本会把它丢掉" });
  });

  it("写回一份件：表没了，制表符是 w:tab，文字顺序不变，读回来还是这些段", () => {
    const doc = WordDocument.blank();
    doc.blocks = [{ kind: "paragraph", inlines: [{ text: "表前" }] }, grid(2, 2), { kind: "paragraph", inlines: [{ text: "表后" }] }];
    const made = ok(doc.blocks[1] as Table, "tab");
    doc.blocks.splice(1, 1, ...made);
    const out = saved(doc);
    expect(out).not.toContain("<w:tbl>");
    expect(out.split("<w:tab/>").length - 1).toBe(2);
    expect([...out.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1])).toEqual(["表前", "00", "01", "10", "11", "表后"]);
    const again = open(writeDocx(doc));
    const texts = again.blocks.filter((b): b is Paragraph => b.kind === "paragraph").map((p) => lineTexts([p])[0]);
    expect(texts).toEqual(["表前", "00\\t01", "10\\t11", "表后"]);
    expect(again.blocks.some((b) => b.kind === "table")).toBe(false);
  });
});
