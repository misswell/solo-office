import { describe, expect, it } from "vitest";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { childSpans, elementContent, OpcPackage, text } from "../src/core/opc.js";
import { renderInline } from "../src/core/render.js";
import { makeCharSpacingDocx } from "./helpers/make-run.js";

/**
 * `w:rPr/w:spacing` — 字符间距. 授权 46 件里 32 枚、住在 8 件里、枚枚非缺省：18 枚在字符样式上写 `-6`（紧缩 0.3 磅），
 * 14 枚在 run 上写 `82`（加宽 4.1 磅），段落标记那一侧 0 枚（`.scratch/layout-elements-audience.py`、
 * `.scratch/layout-second-gate.py`，件按全路径计）。Before this round it travelled as bytes only: the page drew
 * the run at its unstated advance, so a 加宽 4.1 磅 title sat narrower than the file said and broke its line
 * somewhere else.
 */

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const paraAt = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text).join("").includes(needle))!;
const runAt = (doc: WordDocument, needle: string): Inline => paraAt(doc, needle).inlines[0]!;

const PARA_RE = /<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g;
const RUN_RE = /<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*<\/w:r>/g;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(PARA_RE)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
/** One run's own `w:rPr`, nesting honoured — the `w:rPrChange` inside it is part of what the file wrote. */
const rPrOf = (xml: string, needle: string, at = 0): string => {
  const run = [...paraXml(xml, needle).matchAll(RUN_RE)].map((m) => m[0])[at];
  const slot = run ? childSpans(elementContent(run)).find((k) => k.tag === "rPr") : undefined;
  return slot?.xml ?? "";
};
const spacingCount = (xml: string): number => (xml.match(/<w:spacing\b/g) ?? []).length;
const stylesPart = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/styles.xml")!);

describe("w:rPr/w:spacing 字符间距", () => {
  const bytes = makeCharSpacingDocx();

  it("holds the three states the element has", () => {
    const doc = open(bytes);
    expect(runAt(doc, "加宽的字").charSpacingTwips, "the run's own 82 缇").toBe(82);
    expect(runAt(doc, "紧缩的字").charSpacingTwips, "a negative is a legal advance").toBe(-6);
    expect(runAt(doc, "说了不加不减").charSpacingTwips, "0 is said, not silence").toBe(0);
    expect(runAt(doc, "什么都没说").charSpacingTwips, "no element answers nothing").toBeUndefined();
  });

  it("draws the advance on the paper, not only in the model", () => {
    expect(renderInline(runAt(open(bytes), "加宽的字"))).toContain("letter-spacing:4.1pt");
    expect(renderInline(runAt(open(bytes), "紧缩的字"))).toContain("letter-spacing:-0.3pt");
    expect(renderInline(runAt(open(bytes), "说了不加不减"))).toContain("letter-spacing:0pt");
    const silent = renderInline(runAt(open(bytes), "什么都没说"));
    expect(silent, "a run that said nothing gains no spacing").not.toContain("letter-spacing");
    // 借来的那一格也得上纸：梯级读出来的值与 run 自己写的是同一条 CSS。
    expect(renderInline(runAt(open(bytes), "借来的字距"))).toContain("letter-spacing:-0.3pt");
  });

  it("lends it down the style ladder and lets the run's own sentence win", () => {
    const doc = open(bytes);
    expect(runAt(doc, "借来的字距").charSpacingTwips, "the character style's -6 reaches the run").toBe(-6);
    expect(runAt(doc, "自己说过借来的").charSpacingTwips, "§17.7.2 stops at the run that answered").toBe(82);
  });

  it("writes the value back once, in the file's own spelling", () => {
    const out = save(open(bytes));
    expect(rPrOf(out, "加宽的字")).toContain('<w:spacing w:val="82"/>');
    expect(rPrOf(out, "紧缩的字")).toContain('<w:spacing w:val="-6"/>');
    expect(rPrOf(out, "说了不加不减")).toContain('<w:spacing w:val="0"/>');
    expect(rPrOf(out, "自己说过借来的")).toContain('<w:spacing w:val="82"/>');
    expect(rPrOf(out, "自己说过借来的"), "one element per run, not a field copy plus the source bytes").not.toContain(
      '<w:spacing w:val="82"/><w:spacing',
    );
    expect(spacingCount(out), "the fixture's 6 own-side statements, no doubles").toBe(6);
  });

  it("invents nothing: no element for a silent run, none for a borrowed one", () => {
    const out = save(open(bytes));
    expect(rPrOf(out, "什么都没说"), "the run never stated spacing").not.toContain("w:spacing");
    expect(rPrOf(out, "借来的字距"), "a value borrowed from the style is not the run's own sentence").not.toContain(
      "w:spacing",
    );
    expect(rPrOf(out, "借来的字距"), "the pointer it did state stays").toContain('<w:rStyle w:val="WideTitle"/>');
    expect(stylesPart(bytes), "the character style keeps its own -6").toContain('<w:spacing w:val="-6"/>');
  });

  it("keeps the spacing through a keystroke inside the run", () => {
    const doc = open(bytes);
    // The cheapest true keystroke this model has: the editor re-reads the paragraph's spans and hands the run
    // back with new text — the same object the CSS read-back produces. Drop the advance here and the saved file
    // loses a sentence the run is still standing on.
    paraAt(doc, "加宽的字").inlines[0] = { ...runAt(doc, "加宽的字"), text: "加宽的字打过的字" };
    expect(rPrOf(save(doc), "打过的字"), "the re-read run carries the advance back").toContain('<w:spacing w:val="82"/>');
  });

  it("unsays the spacing only when the field is cleared, and says 0 when told to", () => {
    const cleared = open(bytes);
    runAt(cleared, "加宽的字").charSpacingTwips = undefined;
    expect(rPrOf(save(cleared), "加宽的字"), "the run took its own sentence back").not.toContain("w:spacing");

    const zeroed = open(bytes);
    runAt(zeroed, "紧缩的字").charSpacingTwips = 0;
    expect(rPrOf(save(zeroed), "紧缩的字"), "0 is a value, spelled as one").toContain('<w:spacing w:val="0"/>');
  });

  it("keeps a formatting record's own spacing as the record wrote it", () => {
    const doc = open(bytes);
    const run = runAt(doc, "改过字距的字");
    expect(run.charSpacingTwips, "the run now says -6").toBe(-6);
    expect(run.revFormat?.prev.charSpacingTwips, "the record says what was there before").toBe(82);
    const out = save(doc);
    expect(rPrOf(out, "改过字距的字"), "the record's bytes travel whole, both sentences in one element set").toContain(
      '<w:rPrChange w:id="31"',
    );
    expect(rPrOf(out, "改过字距的字")).toContain("<w:spacing w:val=\"82\"/><w:sz");
    expect(spacingCount(out), "own -6 plus the record's 82, once each").toBe(6);
  });

  it("spells a record that has no bytes of its own from the same list the runs are spelled from", () => {
    // A record this editor signs carries no `prevSlots`, so its `w:rPrChange` is spelled from fields alone. That
    // set used to be written out a second time beside `RUN_PROPS`, and the second list did not know `w:spacing`:
    // the same sentence survived on a run and vanished inside its own record (`.scratch/char-spacing-dom.ts`).
    const doc = open(bytes);
    const run = runAt(doc, "改过字距的字");
    delete run.revFormat!.prevSlots;
    const out = save(doc);
    expect(rPrOf(out, "改过字距的字"), "the record still names the advance it was signed with").toContain(
      '<w:spacing w:val="82"/>',
    );
    expect(spacingCount(out), "the run's -6 and the record's 82, once each, and nothing else gained").toBe(6);
  });

  it("survives an open-save-open with every value where it was", () => {
    const before = open(bytes);
    const again = open(writeDocx(before));
    for (const needle of ["加宽的字", "紧缩的字", "说了不加不减", "什么都没说", "借来的字距", "自己说过借来的", "改过字距的字"]) {
      expect(runAt(again, needle).charSpacingTwips, needle).toBe(runAt(before, needle).charSpacingTwips);
    }
    const savedStyles = stylesPart(writeDocx(before));
    expect((savedStyles.match(/<w:spacing\b/g) ?? []).length, "the style's -6 is neither doubled nor dropped").toBe(1);
    expect(savedStyles).toContain('<w:spacing w:val="-6"/>');
  });
});
