import { describe, expect, it } from "vitest";
import type { Block, Paragraph, Table } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { insertRowHere } from "../src/core/doc-table-ops.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeElementAttrsDocx } from "./helpers/make-tables.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const saved = (doc: WordDocument): string => part(writeDocx(doc));

/** One keystroke in the first body paragraph — the path a document is really put through. */
const edited = (bytes: Uint8Array): Uint8Array => {
  const doc = WordDocument.openBytes(bytes);
  const p = doc.blocks.find((b): b is Paragraph => b.kind === "paragraph");
  if (p?.inlines[0]) p.inlines[0].text = "改过的段落";
  return writeDocx(doc);
};

const ids = (xml: string, attr: string): string[] =>
  [...xml.matchAll(new RegExp(`${attr}="([^"]*)"`, "g"))].map((m) => m[1]!);
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");

/**
 * `w14:paraId` is transcribed from WPS-written files; the rsid names are MS-OI29500's and no file here shows
 * one (`.scratch/attr-loss.ts`: 48 `w14:paraId`, 0 rsid across six WPS word files). Both ride the same rule,
 * which is that a save writes back the attribute text it was given rather than a list of names it knows.
 */
describe("docx element-own attributes", () => {
  it("reads an element's own attributes onto the element, not into its text", () => {
    const doc = WordDocument.openBytes(makeElementAttrsDocx());
    const paras = paragraphs(doc);
    expect(paras).toHaveLength(4);
    expect(paras[0]?.attrsXml).toContain('w14:paraId="1B2BC3C3"');
    expect(paras[0]?.attrsXml).toContain('w:rsidP="003E5678"');
    // A paragraph written self-closing still has a tag, and the tag had the id on it.
    expect(paras[2]?.attrsXml).toBe(' w14:paraId="7777AA11" w:rsidR="00A12345"');
    expect(paras[3]?.attrsXml).toBeUndefined();

    const table = doc.blocks.find((b): b is Table => b.kind === "table")!;
    expect(table.rows[0]?.[0]?.attrsXml).toBe(' w:rsidR="00B24C7A"');
    expect(table.rows[0]?.[0]?.row?.attrsXml).toBe(' w:rsidR="00B24C7A" w:rsidTr="00C11111"');
  });

  it("writes each attribute back on the tag it came off, in the order it was given", () => {
    const out = saved(WordDocument.openBytes(makeElementAttrsDocx()));
    expect(out).toContain('<w:p w14:paraId="1B2BC3C3" w:rsidR="00A12345" w:rsidRPr="00A12345" w:rsidP="003E5678">');
    expect(out).toContain('<w:p w14:paraId="5859A351"><w:pPr><w:jc w:val="center"/></w:pPr>');
    expect(out).toContain('<w:tr w:rsidR="00B24C7A" w:rsidTr="00C11111">');
    expect(out).toContain('<w:tc w:rsidR="00B24C7A"><w:tcPr>');
  });

  it("costs a document no identity when one character changes somewhere in it", () => {
    const before = part(makeElementAttrsDocx());
    const after = part(edited(makeElementAttrsDocx()));
    expect(ids(after, "w14:paraId").sort()).toEqual(ids(before, "w14:paraId").sort());
    expect(ids(after, "w:rsidTr")).toEqual(["00C11111"]);
    expect(ids(after, "w:rsidP")).toEqual(["003E5678"]);
    expect(after).toContain("<w:p w14:paraId=\"1B2BC3C3\"");
  });

  it("gives a paragraph it invents no identity, so no two paragraphs share one", () => {
    const after = part(edited(makeElementAttrsDocx()));
    for (const id of ids(after, "w14:paraId")) expect(ids(after, "w14:paraId").filter((x) => x === id)).toHaveLength(1);
    expect(saved(WordDocument.blank())).not.toContain("w14:paraId");
  });

  it("binds the prefix an attribute uses: the root still declares w14 and what may be ignored", () => {
    const after = part(edited(makeElementAttrsDocx()));
    expect(after).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
    expect(after).toContain('mc:Ignorable="w14 w15 wp14"');
  });

  it("costs nothing to save twice: the second file repeats no attribute", () => {
    const once = edited(makeElementAttrsDocx());
    expect(part(edited(once))).toBe(saved(WordDocument.openBytes(once)));
  });

  it("writes a row's own attributes once, though every cell of the row shares the record", () => {
    const doc = WordDocument.openBytes(makeElementAttrsDocx());
    const table = doc.blocks.find((b): b is Table => b.kind === "table")!;
    insertRowHere(table, { row: 0, col: 0 }, "below");
    const after = saved(doc);
    expect(ids(after, "w:rsidTr")).toEqual(["00C11111"]);
    expect(after.match(/<w:tr>/g)?.length).toBe(1);
  });

  it("keeps a cell's own attributes with the cell, and a new cell without any", () => {
    const doc = WordDocument.openBytes(makeElementAttrsDocx());
    const table = doc.blocks.find((b): b is Table => b.kind === "table")!;
    insertRowHere(table, { row: 0, col: 0 }, "below");
    const after = saved(doc);
    // Two on paragraphs, one on the row, one on the cell — and the row that was inserted carries none.
    expect(ids(after, "w:rsidR").sort()).toEqual(["00A12345", "00A12345", "00B24C7A", "00B24C7A"]);
    expect(table.rows[1]?.[0]?.attrsXml).toBeUndefined();
  });

  it("leaves the paragraph's place and words alone while doing it", () => {
    const doc = WordDocument.openBytes(edited(makeElementAttrsDocx()));
    expect(doc.blocks.map((b) => b.kind)).toEqual(["paragraph", "paragraph", "paragraph", "table", "paragraph"]);
    expect(paragraphs(doc).map((p) => p.inlines.map((i) => i.text).join(""))).toEqual([
      "改过的段落",
      "居中且带标记",
      "",
      "没有任何标记的段落",
    ]);
    const empty = doc.blocks[2] as Paragraph;
    expect(empty.attrsXml).toBe(' w14:paraId="7777AA11" w:rsidR="00A12345"');
    expect(empty.inlines).toEqual([]);
    const table = doc.blocks.find((b): b is Table => b.kind === "table")!;
    expect((table.rows[0]![0]!.paragraphs[0] as Paragraph).inlines.map((i) => i.text).join("")).toBe("格内带标记的段落");
    expect((table.rows[0]![0]!.paragraphs[0] as Paragraph).attrsXml).toBe(' w14:paraId="0A0B0C0D"');
  });
});
