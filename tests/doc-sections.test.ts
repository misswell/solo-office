import { describe, expect, it } from "vitest";
import type { Block } from "../src/core/docx.js";
import { WordDocument, sectionEndsOf } from "../src/core/docx.js";
import { makeSectionsDocx, SECT_PR } from "./helpers/make-sections.js";

/** Text as a reader sees it, one entry per top-level block. */
const shape = (blocks: readonly Block[]): string[] =>
  blocks.map((b) => (b.kind === "paragraph" ? b.inlines.map((i) => i.text ?? "").join("") : b.kind));

describe("sectionEndsOf", () => {
  it("边界是带 w:pPr/w:sectPr 的段落，断点段落自己属于它结束的那一节", () => {
    const doc = WordDocument.openBytes(
      makeSectionsDocx([
        { text: "甲一" },
        { text: "甲二" },
        { pPr: SECT_PR },
        { text: "乙一" },
        { text: "乙二" },
        { pPr: SECT_PR },
        { text: "丙一" },
        { text: "丙二" },
      ]),
    );
    expect(shape(doc.blocks)).toEqual(["甲一", "甲二", "", "乙一", "乙二", "", "丙一", "丙二"]);
    expect(sectionEndsOf(doc.blocks)).toEqual([2, 5]);
    // 正文自己那个 w:sectPr 不进 blocks（blockChildren 把它挡在外面），它只是最后一节的收尾。
    expect(doc.blocks.length).toBe(8);
    expect(sectionEndsOf(doc.blocks).length + 1).toBe(doc.sections.length);
  });

  it("没有断点的文档：ends 空、一节", () => {
    const doc = WordDocument.openBytes(makeSectionsDocx([{ text: "只有一段" }, { text: "第二段" }]));
    expect(sectionEndsOf(doc.blocks)).toEqual([]);
    expect(doc.sections.length).toBe(1);
  });

  it("断点段落自己进 blocks（不是被吃掉的空段），pPrSlots 里带着那个 sectPr", () => {
    const doc = WordDocument.openBytes(makeSectionsDocx([{ text: "甲" }, { pPr: SECT_PR }, { text: "乙" }]));
    const brk = doc.blocks[1]!;
    expect(brk.kind).toBe("paragraph");
    if (brk.kind === "paragraph") expect(brk.pPrSlots?.some((s) => s.tag === "sectPr")).toBe(true);
    expect(sectionEndsOf(doc.blocks)).toEqual([1]);
  });
});
