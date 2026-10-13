import { describe, expect, it } from "vitest";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { childSpans, elementContent, OpcPackage, text } from "../src/core/opc.js";
import { makeRunPropsDocx, makeUnderlineDocx } from "./helpers/make-run.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));

const PARA_RE = /<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g;
const RUN_RE = /<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*<\/w:r>/g;
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const paraAt = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text).join("").includes(needle))!;
const runAt = (doc: WordDocument, needle: string): Inline => paraAt(doc, needle).inlines[0]!;

/** A keystroke in the paragraph under test — the path a run's own property takes. */
const retyped = (doc: WordDocument, needle: string, next: Inline): string => {
  paraAt(doc, needle).inlines[0] = next;
  return save(doc);
};

/** One added paragraph: the cheapest edit that rebuilds the whole body from the model. */
const elsewhere = (bytes: Uint8Array): string => {
  const doc = open(bytes);
  doc.blocks.push({ kind: "paragraph", inlines: [{ text: "回归" }] });
  return part(writeDocx(doc));
};

const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(PARA_RE)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
const runsOf = (xml: string, needle: string): string[] => [...paraXml(xml, needle).matchAll(RUN_RE)].map((m) => m[0]);
/** One run's own `w:rPr`, nesting honoured: the record inside it is part of what the file wrote. */
const rPrOf = (xml: string, needle: string, at = 0): string => {
  const run = runsOf(xml, needle)[at];
  const slot = run ? childSpans(elementContent(run)).find((k) => k.tag === "rPr") : undefined;
  return slot?.xml ?? "";
};

/**
 * A run says more about itself than this model has fields for, and the page is rebuilt from the model. The
 * shapes here were measured, not invented: `.scratch/wps-rpr-names.ts` over 86 WPS-written documents counted
 * 101825 run property sets holding `rFonts` 101175, `szCs` 96165, `sz` 93946, `color` 61554, `lang` 26334,
 * `kern` 18399, `rStyle` 17306, `bCs` 12233, `b` 8889, `spacing` 6666, `vertAlign` 2701, `highlight` 2666,
 * `textFill` 2214, `shd` 1454, `u` 1256, `i` 971, `iCs` 960, `caps` 274, `strike` 38.
 */
describe("docx run property fidelity", () => {
  it("reads a run's whole w:rPr child by child and keeps the fields the page draws", () => {
    const run = runAt(open(makeRunPropsDocx()), "整段自述");
    expect(run.rPrSlots?.map((s) => s.tag)).toEqual([
      "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "color", "spacing", "w", "kern",
      "position", "sz", "szCs", "highlight", "u", "shd", "vertAlign", "rtl", "lang", "textFill",
    ]);
    expect(run.rPrSlots?.find((s) => s.tag === "color")?.xml).toContain('w:themeColor="text1"');
    expect(run.rPrSlots?.find((s) => s.tag === "rFonts")?.xml).toContain('w:eastAsia="宋体"');
    expect(run.rPrSlots?.find((s) => s.tag === "textFill")?.xml).toContain('w14:val="C00000"');
    // The fields the page draws still read what they read before, plus the second face of this element.
    expect(run.bold).toBe(true);
    expect(run.sizePt).toBe(12);
    expect(run.color).toBe("#808080");
    expect(run.font).toBe("Times New Roman");
    expect(run.fontEastAsia).toBe("宋体");
    expect(run.asRead).toEqual({
      bold: true, italic: false, underline: true, strike: true, sizePt: 12, color: "#808080",
      highlight: "yellow", font: "Times New Roman", fontEastAsia: "宋体",
    });
  });

  it("costs a run nothing when a paragraph is added somewhere else in the document", () => {
    const before = part(makeRunPropsDocx());
    const after = elsewhere(makeRunPropsDocx());
    for (const needle of ["整段自述", "字号成对", "留痕的格式", "链接里的字"]) {
      expect(rPrOf(after, needle)).toBe(rPrOf(before, needle));
    }
    expect(rPrOf(after, "两截同式", 1)).toBe(rPrOf(before, "两截同式", 1));
  });

  it("re-spells only the child whose field moved and leaves the rest of the set where it sat", () => {
    const doc = open(makeRunPropsDocx());
    const out = rPrOf(retyped(doc, "整段自述", { ...runAt(doc, "整段自述"), bold: false }), "整段自述");
    expect(out).toContain('<w:b w:val="0"/>');
    // `CT_RPr` puts `w:b` between `w:rFonts` and `w:bCs`, and the complex-script face is not the field's to take.
    expect(out).toMatch(/<w:rFonts[^>]*\/><w:b w:val="0"\/><w:bCs\/>/);
    expect(out).toContain('<w:kern w:val="28"/>');
    expect(out).toContain('w:themeColor="text1"');
    expect(out).toContain("<w14:textFill>");
  });

  it("keeps the faces a font pick does not own when it rewrites the faces it does", () => {
    const doc = open(makeRunPropsDocx());
    const out = rPrOf(retyped(doc, "整段自述", { ...runAt(doc, "整段自述"), font: "黑体" }), "整段自述");
    const fonts = /<w:rFonts[^>]*>/.exec(out)?.[0] ?? "";
    expect(fonts).toContain('w:ascii="黑体"');
    expect(fonts).toContain('w:hAnsi="黑体"');
    expect(fonts.match(/w:ascii=/g)).toHaveLength(1);
    expect(fonts).toContain('w:eastAsia="宋体"');
    expect(fonts).toContain('w:cs="Times New Roman"');
    expect(fonts).toContain('w:hint="eastAsia"');
    expect(fonts).not.toContain('w:ascii="Times New Roman"');
  });

  it("re-spells the east-asian face on the run that moved it and leaves the Latin one where it sat", () => {
    // The other half of the same rule: the two faces live on one element, so each moves only with its field.
    const doc = open(makeRunPropsDocx());
    const out = rPrOf(retyped(doc, "整段自述", { ...runAt(doc, "整段自述"), fontEastAsia: "仿宋" }), "整段自述");
    const fonts = /<w:rFonts[^>]*>/.exec(out)?.[0] ?? "";
    expect(fonts).toContain('w:eastAsia="仿宋"');
    expect(fonts.match(/w:eastAsia=/g)).toHaveLength(1);
    expect(fonts).toContain('w:ascii="Times New Roman"');
    expect(fonts).toContain('w:cs="Times New Roman"');
    expect(fonts).toContain('w:hint="eastAsia"');
  });

  it("moves the complex-script size with the Latin one, and only when the size moved", () => {
    const untouched = rPrOf(elsewhere(makeRunPropsDocx()), "字号成对");
    expect(untouched).toContain('<w:sz w:val="21"/>');
    expect(untouched).toContain('<w:szCs w:val="24"/>');

    const doc = open(makeRunPropsDocx());
    const out = rPrOf(retyped(doc, "字号成对", { ...runAt(doc, "字号成对"), sizePt: 10 }), "字号成对");
    expect(out).toContain('<w:sz w:val="20"/><w:szCs w:val="20"/>');
    expect(out).toContain('w:eastAsia="仿宋"');
    expect(out).not.toContain('w:val="21"');
  });

  it("keeps a run whose every child is machinery the model has no case for", () => {
    const doc = open(makeRunPropsDocx());
    const shape = paraAt(doc, "带形状的段");
    expect(shape.inlines.filter((i) => i.carry)).toHaveLength(1);
    const after = save(doc);
    const carried = paraXml(after, "带形状的段");
    expect(carried).toContain("<mc:AlternateContent>");
    expect(carried).toContain('name="直接连接符 1"');
    expect(rPrOf(after, "带形状的段")).toBe(rPrOf(part(makeRunPropsDocx()), "带形状的段"));
  });

  it("puts back a run's empty w:rPr instead of taking it away", () => {
    expect(rPrOf(elsewhere(makeRunPropsDocx()), "空的性质")).toBe("<w:rPr/>");
  });

  it("keeps the unmodelled properties when the run's own record is decided either way", () => {
    const accepted = open(makeRunPropsDocx());
    expect(accepted.resolveRevisions("accept", 11)).toBe(1);
    const kept = rPrOf(save(accepted), "留痕的格式");
    expect(kept).not.toContain("rPrChange");
    expect(kept).toContain('<w:sz w:val="28"/>');
    expect(kept).toContain('<w:kern w:val="28"/>');
    expect(kept).toContain('w:eastAsia="zh-CN"');

    const refused = open(makeRunPropsDocx());
    expect(refused.resolveRevisions("reject", 11)).toBe(1);
    const back = rPrOf(save(refused), "留痕的格式");
    expect(back).not.toContain("rPrChange");
    // The record holds the whole set from before, and a property it does not name was not set: the size pair
    // moves, the kerning and the faces come back as the record wrote them, and the language goes with them.
    expect(back).toContain('<w:sz w:val="22"/><w:szCs w:val="22"/>');
    expect(back).toContain('<w:kern w:val="28"/>');
    expect(back).toContain('<w:highlight w:val="green"/>');
    expect(back).toContain('w:eastAsia="宋体"');
    expect(back).not.toContain('w:eastAsia="zh-CN"');
  });

  it("keeps two runs that read alike but were written apart", () => {
    const before = part(makeRunPropsDocx());
    const after = save(open(makeRunPropsDocx()));
    expect(rPrOf(after, "两截同式")).toBe('<w:rPr><w:b/><w:lang w:val="en-US"/></w:rPr>');
    expect(rPrOf(after, "两截同式", 1)).toBe('<w:rPr><w:b/><w:kern w:val="28"/></w:rPr>');
    expect(rPrOf(after, "两截同式", 1)).toBe(rPrOf(before, "两截同式", 1));
  });

  it("writes an underline only when the field moved, not because a link reads as underlined", () => {
    const doc = open(makeRunPropsDocx());
    expect(runAt(doc, "链接里的字").underline).toBe(true);
    // The `w:Hyperlink` style draws it and the file never wrote a `w:u`, so an untouched link comes back as it sat.
    expect(rPrOf(elsewhere(makeRunPropsDocx()), "链接里的字")).toBe('<w:rPr><w:rStyle w:val="Hyperlink"/><w:kern w:val="28"/></w:rPr>');

    // Moved, it is spelled into the carried set at the rank `CT_RPr` gives it: after `w:szCs`, before `w:shd`.
    const on = open(makeRunPropsDocx());
    expect(rPrOf(retyped(on, "字号成对", { ...runAt(on, "字号成对"), underline: true }), "字号成对")).toBe(
      '<w:rPr><w:rFonts w:eastAsia="仿宋" w:cs="仿宋"/><w:sz w:val="21"/><w:szCs w:val="24"/><w:u w:val="single"/></w:rPr>',
    );
    // Off, the cancellation is a value the run writes rather than an element it deletes, exactly as `w:b` above
    // becomes `<w:b w:val="0"/>`. That is the shape WPS itself uses: `<w:u w:val="none"/>` appears 952 times in
    // these bodies and 254 of them sit on runs this reader resolves into fields
    // (`.scratch/underline-counted.ts`, `.scratch/underline-audience.ts`).
    const off = open(makeRunPropsDocx());
    expect(rPrOf(retyped(off, "整段自述", { ...runAt(off, "整段自述"), underline: false }), "整段自述")).toContain(
      '<w:u w:val="none"/>',
    );
  });

  it("carries a run's properties through every round trip, not just the first", () => {
    const once = writeDocx(open(makeRunPropsDocx()));
    const twice = writeDocx(open(once));
    expect(part(twice)).toBe(part(once));
    expect(runAt(open(twice), "整段自述").rPrSlots).toHaveLength(23);
  });

  it("still spells a run the editor invented straight from its fields", () => {
    const doc = WordDocument.blank();
    const p = paragraphs(doc)[0]!;
    expect(p.inlines[0]?.rPrSlots).toBeUndefined();
    p.inlines.push({ text: "新写的", bold: true, sizePt: 14 });
    expect(paraXml(save(doc), "新写的")).toContain('<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>');
  });
});

/**
 * `w:u` counted over the same 86 WPS documents, twice and reconciled (`.scratch/underline-counted.ts` walks the
 * bytes by tag stack, `.scratch/underline-audience.ts` asks the production model). The bodies write 1586 `w:u`:
 * 1256 on runs and 330 on paragraph marks. Of the runs, 488 reach this reader as fields (254 `none`, 194 `single`,
 * 40 `thick`, and not one leaves `w:val` off) while the other 768 travel as bytes inside table cells.
 * `word/styles.xml` is not silent: it states `w:u` 35 times, 31 of them on character styles with `single`, and
 * those styles do lend a line to 8 runs that say nothing of their own — but all 254 cancelling runs sit on chains
 * that say nothing about `w:u`, so no run today has a line to beat. `w:rPrDefault` states none, and the bodies
 * contain 0 `w:hyperlink` elements, so `linkLook`'s forced line has no audience in hand either.
 * The rung that lends a line (`LendU` below) is therefore the shape the corpus points at but does not itself
 * carry: built to the 31 character styles' pattern so the 254 sentences have something they must beat. Nothing
 * here claims a pixel the corpus has not shown — what the change buys is that a stated cancellation survives.
 */
describe("docx underline read as a statement", () => {
  it("reads a stated no-underline as a value the run gave, not a missing element", () => {
    const run = runAt(open(makeUnderlineDocx()), "说明没有");
    expect(run.underline).toBe(false);
    expect(run.asRead?.underline).toBe(false);
  });

  it("stops the ladder per property at the rung that stated one", () => {
    const doc = open(makeUnderlineDocx());
    // §17.7.2 ranks the run's own `w:rPr` above the character style it names, so this run's `none` beats the
    // line the style lends. The other three are the controls: a style that is not contradicted lends a line,
    // and a run that names no style still says which way it goes for itself.
    expect(runAt(doc, "这句话盖掉借来的线").underline).toBe(false);
    expect(runAt(doc, "链上借来的线").underline).toBe(true);
    expect(runAt(doc, "自己点了线").underline).toBe(true);
    expect(runAt(doc, "粗线原样回去").underline).toBe(true);
  });

  it("hands back a stated none and a stated thick exactly as they sat", () => {
    const before = part(makeUnderlineDocx());
    const after = save(open(makeUnderlineDocx()));
    expect(rPrOf(after, "说明没有")).toBe(rPrOf(before, "说明没有"));
    expect(rPrOf(after, "说明没有")).toBe('<w:rPr><w:u w:val="none"/></w:rPr>');
    // The model holds whether a line is on, not what shape it has, so a shape it cannot name travels as bytes.
    expect(rPrOf(after, "粗线原样回去")).toBe('<w:rPr><w:u w:val="thick"/></w:rPr>');
  });

  it("writes the cancellation when the page turns off a line the style lent", () => {
    const doc = open(makeUnderlineDocx());
    const out = rPrOf(
      retyped(doc, "链上借来的线", { ...runAt(doc, "链上借来的线"), underline: false }),
      "链上借来的线",
    );
    expect(out).toBe('<w:rPr><w:rStyle w:val="LendU"/><w:u w:val="none"/></w:rPr>');
  });

  it("writes a line back when the page turns a cancelled run on", () => {
    const doc = open(makeUnderlineDocx());
    const out = rPrOf(retyped(doc, "说明没有", { ...runAt(doc, "说明没有"), underline: true }), "说明没有");
    expect(out).toBe('<w:rPr><w:u w:val="single"/></w:rPr>');
  });
});
