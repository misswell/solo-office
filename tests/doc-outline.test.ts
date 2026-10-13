import { describe, expect, it } from "vitest";
import { headingLevel, WordDocument, type Paragraph } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { jcBefore, makeOutlineDocx, outlineLvl, pStyle, type OutlineBlock } from "./helpers/make-outline.js";

/**
 * What an outline level does. Measured over the 86 WPS-written documents before anything here changed
 * (`.scratch/outline-census.ts`):
 *
 * - the level a paragraph is drawn under comes mostly from its style, not from the paragraph: 264 of their
 *   46,417 paragraphs state `w:outlineLvl` themselves, while 582 more take it from the chain, and 433 paragraph
 *   styles state one. A model that reads only the paragraph sees a quarter of the headings.
 * - `w:val="9"` is a real shape and not a typo for "no level": 56 of the 264 self-stating paragraphs write it,
 *   and 61 of the 46,417 resolve to it — §17.3.2.26's body text, which is a heading level its style lends being
 *   taken back. Keeping the number rather than folding it into "a level or none" is what cancels.
 * - the renderer's old way of spotting a heading — `w:pStyle` named `Heading1`…`Heading9` — matches **0** of the
 *   433 styleIds that state a level here, so no heading in these files was ever drawn as one, and no pane had
 *   anything to list. 785 paragraphs resolve to a level of 0…3, which is the whole shape this covers.
 */

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
/**
 * A regenerated `word/document.xml`: an unedited document saves back as the bytes it read, so a test about what
 * the writer spells has to make it edit something first. The paragraph pushed here is not the subject.
 */
const saved = (doc: WordDocument): string => {
  doc.blocks.push({ kind: "paragraph", inlines: [{ text: "补一行" }] });
  expect(doc.edited).toBe(true);
  return part(writeDocx(doc));
};

const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
const pPrOf = (xml: string, needle: string): string =>
  /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>|<w:pPr\/>/.exec(paraXml(xml, needle))?.[0] ?? "";

/** Every paragraph in the model, table cells included, in the order the page walks them. */
function allParas(doc: WordDocument): Paragraph[] {
  const out: Paragraph[] = [];
  const walk = (blocks: WordDocument["blocks"]): void => {
    for (const b of blocks) {
      if (b.kind === "paragraph") out.push(b);
      else if (b.kind === "table") for (const row of b.rows) for (const cell of row) walk(cell.paragraphs);
    }
  };
  walk(doc.blocks);
  return out;
}
const byText = (doc: WordDocument, needle: string): Paragraph =>
  allParas(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;

/**
 * The page's heading marks in the order they are written: level and text. This is what the 导航窗格 counts to
 * turn a row number into a place in the document, so its agreement with `doc.outline()` is the whole addressing
 * scheme and needs to be tested as one thing.
 */
function marks(html: string): { level: number; text: string }[] {
  const out: { level: number; text: string }[] = [];
  for (const m of html.matchAll(/<(div|h[1-6])\s[^>]*\bdata-outline="(\d)"[^>]*>([\s\S]*?)<\/\1>/g)) {
    out.push({ level: Number(m[2]), text: m[3].replace(/<[^>]*>/g, "").replace(/&nbsp;/g, "").trim() });
  }
  return out;
}

/** The element the page draws for the paragraph that says `needle`. */
function drawn(html: string, needle: string): string {
  for (const m of html.matchAll(/<(div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>(?:(?!<\/(?:div|h[1-6])>)[\s\S])*<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return m[0];
  }
  return "";
}
/** The tag the paragraph is drawn under, and the level the mark names it by ("" when it carries none). */
const drawnTag = (html: string, needle: string): string => /<(div|h[1-6])\b/.exec(drawn(html, needle))?.[1] ?? "";
const drawnOutline = (html: string, needle: string): string => /data-outline="(\d)"/.exec(drawn(html, needle))?.[1] ?? "";

const BODY: OutlineBlock[] = [
  { pPr: pStyle("T1"), text: "样式给的一级" },
  { pPr: outlineLvl(2), text: "自己写的三级" },
  { pPr: outlineLvl(9), text: "自己写9回正文" },
  { pPr: pStyle("Body"), text: "样式链给9回正文" },
  { pPr: pStyle("Deep"), text: "借两级到的二级" },
  { pPr: `${pStyle("T1")}${outlineLvl(3)}`, text: "自己盖过样式" },
  { pPr: jcBefore("center", 1), text: "对齐里带级别" },
  { pPr: `${pStyle("T1")}${outlineLvl(9)}`, text: "自己收回样式给的级" },
  { pPr: pStyle("Heading2"), text: "只有名字像标题" },
  { text: "普通正文" },
  { table: [[{ pPr: pStyle("T2"), text: "格子里的标题" }, { text: "格子里的正文" }]] },
  { pPr: outlineLvl(8), text: "第九级" },
  { pPr: outlineLvl(0), text: "" },
];

describe("docx 大纲级别：从样式链读全、9 是正文、页面按解析出的级画标题", () => {
  it("reads the level the paragraph states for itself", () => {
    const doc = open(makeOutlineDocx(BODY));
    expect(byText(doc, "自己写的三级").outlineLevel).toBe(2);
    expect(headingLevel(byText(doc, "自己写的三级"))).toBe(3);
  });

  it("reads the level its own style states, and the one two styles down the chain lends", () => {
    const doc = open(makeOutlineDocx(BODY));
    expect(byText(doc, "样式给的一级").outlineLevel).toBe(0);
    expect(headingLevel(byText(doc, "样式给的一级"))).toBe(1);
    // `Deep` states nothing and is based on `T2`, which states 1: the level travels the whole §17.7.2 chain.
    expect(byText(doc, "借两级到的二级").outlineLevel).toBe(1);
    expect(headingLevel(byText(doc, "借两级到的二级"))).toBe(2);
  });

  it("keeps w:val=\"9\" as body text, whether the paragraph or its style writes it", () => {
    const doc = open(makeOutlineDocx(BODY));
    for (const needle of ["自己写9回正文", "样式链给9回正文"]) {
      expect(byText(doc, needle).outlineLevel).toBe(9);
      expect(headingLevel(byText(doc, needle))).toBeUndefined();
    }
    // The paragraph that writes 9 over a style that lends 0 has taken a level back — 61 of the corpus's
    // paragraphs resolve to exactly that.
    const taken = byText(doc, "自己收回样式给的级");
    expect(taken.outlineLevel).toBe(9);
    expect(headingLevel(taken)).toBeUndefined();
  });

  it("lets the paragraph's own statement beat the one its style lends", () => {
    const doc = open(makeOutlineDocx(BODY));
    expect(byText(doc, "自己盖过样式").outlineLevel).toBe(3);
    expect(headingLevel(byText(doc, "自己盖过样式"))).toBe(4);
  });

  it("draws the heading the resolved level names, and marks it for the pane", () => {
    const html = renderWord(open(makeOutlineDocx(BODY)));
    expect(drawnTag(html, "样式给的一级")).toBe("h1");
    expect(drawnOutline(html, "样式给的一级")).toBe("1");
    expect(drawnTag(html, "自己写的三级")).toBe("h3");
    expect(drawnOutline(html, "自己写的三级")).toBe("3");
    expect(drawnTag(html, "借两级到的二级")).toBe("h2");
    expect(drawnOutline(html, "格子里的标题")).toBe("2");
    // Body text, the 9 that cancels a level and the paragraph that merely has a heading-like name get no mark:
    // a row in the pane has to stand for a paragraph the file called a heading, not for one that looks like it.
    for (const needle of ["普通正文", "自己写9回正文", "样式链给9回正文", "自己收回样式给的级", "格子里的正文"]) {
      expect(drawnOutline(html, needle), needle).toBe("");
    }
  });

  it("draws a Word-style styleId as a heading without inventing an outline row for it", () => {
    const doc = open(makeOutlineDocx(BODY));
    const html = renderWord(doc);
    expect(drawnTag(html, "只有名字像标题")).toBe("h2");
    expect(drawnOutline(html, "只有名字像标题")).toBe("");
    expect(doc.outline().some((e) => e.text === "只有名字像标题")).toBe(false);
  });

  it("holds the level of a deep one as a mark while leaving the tag a paragraph", () => {
    const html = renderWord(open(makeOutlineDocx(BODY)));
    // Every level the corpus resolves to is 0…3, so the page's own h1…h4 rules cover them. A browser draws h5
    // and below *smaller* than body text, so a deeper level is marked for the pane and drawn as a paragraph.
    expect(drawnTag(html, "第九级")).toBe("div");
    expect(drawnOutline(html, "第九级")).toBe("9");
  });

  it("lists the headings in the order the page writes its marks, table cells inside that order", () => {
    const doc = open(makeOutlineDocx(BODY));
    expect(doc.outline()).toEqual(marks(renderWord(doc)));
    expect(doc.outline()).toEqual([
      { level: 1, text: "样式给的一级" },
      { level: 3, text: "自己写的三级" },
      { level: 2, text: "借两级到的二级" },
      { level: 4, text: "自己盖过样式" },
      { level: 2, text: "对齐里带级别" },
      { level: 2, text: "格子里的标题" },
      { level: 9, text: "第九级" },
      { level: 1, text: "" },
    ]);
  });

  it("writes the level of a paragraph the editor made, in the order CT_PPr gives it", () => {
    const doc = open(makeOutlineDocx([{ text: "先有一段" }]));
    // A paragraph split with Enter arrives with fields and no `w:pPr` slots, so the level has to be spelled
    // from the field or a heading cut in two saves as one heading and one body paragraph.
    doc.blocks.push({ kind: "paragraph", inlines: [{ text: "新起的标题" }], outlineLevel: 0, align: "center" });
    expect(doc.edited).toBe(true);
    const xml = part(writeDocx(doc));
    expect(pPrOf(xml, "新起的标题")).toBe(`<w:pPr><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr>`);
  });

  it("does not claim a level the paragraph only borrows from its style", () => {
    const xml = saved(open(makeOutlineDocx(BODY)));
    expect(pPrOf(xml, "样式给的一级")).toBe(`<w:pPr>${pStyle("T1")}</w:pPr>`);
    expect(pPrOf(xml, "借两级到的二级")).toBe(`<w:pPr>${pStyle("Deep")}</w:pPr>`);
  });

  it("keeps the level the file wrote, including the 9 that cancels one", () => {
    const xml = saved(open(makeOutlineDocx(BODY)));
    expect(pPrOf(xml, "自己写的三级")).toContain(outlineLvl(2));
    expect(pPrOf(xml, "自己写9回正文")).toContain(outlineLvl(9));
    expect(pPrOf(xml, "自己盖过样式")).toBe(`<w:pPr>${pStyle("T1")}${outlineLvl(3)}</w:pPr>`);
    // The children travel as the part wrote them, `w:jc` before `w:outlineLvl` the way `CT_PPr` ranks them —
    // re-spelling the set from the fields is how a save takes the order down with it.
    expect(pPrOf(xml, "对齐里带级别")).toBe(`<w:pPr>${jcBefore("center", 1)}</w:pPr>`);
  });
});
