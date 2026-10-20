import { describe, expect, it } from "vitest";
import type { Table } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { makeRowAlignDocx } from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const tableOf = (doc: WordDocument): Table => doc.blocks.find((b) => b.kind === "table") as Table;
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
const rowJc = (t: Table): (string | undefined)[] => t.rows.map((cells) => cells[0]?.row?.jc);
const rowOf = (t: Table, index: number) => t.rows[index][0]!.row!;

describe("docx 行对齐：w:trPr/w:jc", () => {
  it("reads the row's word, and the one its trPrChange recorded", () => {
    const t = tableOf(open(makeRowAlignDocx()));
    expect(rowJc(t)).toEqual(["center", undefined, "right", "center"]);
    expect(rowOf(t, 3).revFormat).toMatchObject({ id: 101, author: "顾峰", prev: { jc: "left" } });
  });

  it("draws the row's word on the row, and a paragraph that states its own still wins", () => {
    const html = renderWord(open(makeRowAlignDocx()));
    expect(html).toContain('<tr style="text-align:center;">');
    expect(html).toContain('<tr style="text-align:right;">');
    // The cell that says nothing takes the row's word by inheritance; the one whose paragraph says `left`
    // keeps it. `left` had no entry in the renderer's map, so an explicit left inside a centered row would
    // otherwise have inherited the row's center.
    expect(html).toContain('class="para" style="text-align:left"');
    expect((html.match(/<tr style="text-align:center;">/g) ?? []).length).toBe(2);
  });

  it("writes the row's word back untouched through a save", () => {
    const doc = open(makeRowAlignDocx());
    const out = saved(doc);
    expect(out).toContain('<w:trPr><w:jc w:val="center"/></w:trPr>');
    expect(out).toContain('<w:trPr><w:jc w:val="right"/></w:trPr>');
    expect(out).toContain('<w:jc w:val="center"/><w:trPrChange');
    expect(out).toContain('<w:pPr><w:jc w:val="left"/></w:pPr>');
    expect((out.match(/<w:jc\b/g) ?? []).length).toBe(5);
  });

  it("refusing the recorded change puts the earlier word back; accepting keeps today's", () => {
    const refused = open(makeRowAlignDocx());
    expect(refused.resolveRevisions("reject", 101)).toBe(1);
    expect(rowOf(tableOf(refused), 3).jc).toBe("left");
    const refusedXml = saved(refused);
    expect(refusedXml).toContain('<w:trPr><w:jc w:val="left"/></w:trPr>');
    expect(refusedXml).not.toContain("trPrChange");

    const kept = open(makeRowAlignDocx());
    expect(kept.resolveRevisions("accept", 101)).toBe(1);
    expect(rowOf(tableOf(kept), 3).jc).toBe("center");
    expect(saved(kept)).toContain('<w:trPr><w:jc w:val="center"/></w:trPr>');
  });

  it("draws the row as it stood in the original state, without rewriting the document", () => {
    const doc = open(makeRowAlignDocx());
    const original = doc.originalView(doc.blocks);
    expect(rowOf(original.find((b) => b.kind === "table") as Table, 3).jc).toBe("left");
    expect(renderWord(doc, "original")).toContain('<tr style="text-align:left;">');
    expect(rowOf(tableOf(doc), 3).jc).toBe("center");
  });
});
