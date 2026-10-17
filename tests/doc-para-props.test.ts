import { describe, expect, it } from "vitest";
import type { Paragraph } from "../src/core/docx.js";
import { syncParaSlots, WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { childSpans, elementContent, OpcPackage, text } from "../src/core/opc.js";
import { makeChainDocx, makeParaPropsDocx } from "./helpers/make-paragraph.js";
import { renderWord } from "../src/core/render.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));

/** One keystroke in one paragraph: the path a document is actually put through before its properties move. */
const edited = (bytes: Uint8Array, needle = "被并进来的下段"): string => {
  const doc = open(bytes);
  const p = paragraphs(doc).find((x) => textOf(x).includes(needle))!;
  p.inlines[0] = { ...p.inlines[0]!, text: "改过的" };
  return part(writeDocx(doc));
};

const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const textOf = (p: Paragraph): string => p.inlines.map((r) => r.text ?? "").join("");
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
/** The paragraph's own `w:pPr` content, nesting honoured, minus nothing: a record inside it is part of it. */
const pPrOf = (xml: string, needle: string): string => {
  const slot = childSpans(elementContent(paraXml(xml, needle))).find((k) => k.tag === "pPr");
  return slot ? elementContent(slot.xml) : "";
};

/**
 * A paragraph says more about itself than this model has fields for, and the page is rebuilt from the model.
 * The shapes these tests transcribe were measured, not invented — on the authorized 46 份
 * (`/tmp/wps-authorized.txt`, 9 `.docx` + 34 `.dotm` + 3 `.dotx`; `.scratch/ppr-children-authorized.ts` walks a tag
 * stack, so a `w:pPr` counts only when its parent really is a `w:p`: 7,357 of them in `word/document.xml`, and no
 * `w:pPr` is self-closed). Its direct children run `rPr` 7199, `spacing` 3924, `jc` 2548, `ind` 2278, `pStyle` 1518,
 * `snapToGrid` 406, `adjustRightInd` 334, `numPr` 298, `bidi` 294, `widowControl` 246, `tabs` 230, `outlineLvl` 202,
 * then `keepNext`, `keepLines`, `pageBreakBefore`, `kinsoku`, `wordWrap`, `overflowPunct`, `topLinePunct`,
 * `autoSpaceDE`, `autoSpaceDN` and `textAlignment` 200 each, `shd` 30 and `sectPr` 12 — and that list is exhaustive,
 * which is what the two halves below are for: `suppressLineNumbers`, `framePr` and `pBdr` are written 0 times by a
 * body paragraph (the bands write `framePr` 6 and `pBdr` 2).
 * The paragraph mark's own `w:rPr` children are `rFonts` 7177, `sz` 6648, `szCs` 2066, `b` 1422, `color` 1324,
 * `bCs` 412, `u` 335, `kern` 324, `lang` 231, `i` 114, `iCs` 78, and nothing else — no `rStyle`, no `highlight`,
 * no `w:rStyle`, no `highlight`, no `textFill`, no `ligatures`, no `vertAlign`, and no character `spacing` on a
 * paragraph mark (the 14 `w:rPr/w:spacing` the part writes are all on runs). `contextualSpacing` is 0 in every
 * `w:pPr` this walk reaches. The walk is proven not to lose any: counting the part loosely gives `word/document.xml`
 * 3,938 `w:spacing` = the 3,924 above plus those 14, `word/styles.xml` 264, `word/footnotes.xml` 12 and
 * `word/endnotes.xml` 12 (`.scratch/spacing-loose.ts`). That is also the reconciliation of the 3,948 the §105-era
 * comments cite: it is the same `w:pPr` count over **three** body-side parts — 3,924 + 12 + 12 — not this part's,
 * so name the parts whenever the number is quoted.
 * These are the authorized 46 份 only: the 86-file cohort these numbers used to come from was built from the
 * owner's personal documents and is void — see `.scratch/HANDOFF.md`.
 */
describe("docx paragraph property fidelity", () => {
  it("reads a paragraph's whole w:pPr child by child and keeps the record as a decision", () => {
    const p = paragraphs(open(makeParaPropsDocx()))[0]!;
    expect(p.pPrSlots?.map((s) => s.tag)).toEqual([
      "keepNext", "keepLines", "pageBreakBefore", "widowControl", "suppressLineNumbers", "pBdr", "shd", "tabs",
      "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "textAlignment", "outlineLvl", "rPr",
    ]);
    expect(p.pPrSlots?.find((s) => s.tag === "rPr")?.xml).toContain('<w:highlight w:val="yellow"/>');
    expect(p.pPrSlots?.find((s) => s.tag === "ind")?.xml).toContain('w:firstLine="0"');
    // The fields still read what they always read, so the page draws the same paragraph.
    expect(p.pageBreakBefore).toBe(true);
    expect(p.spacingBefore).toBe(120);
    expect(p.indentLeft).toBe(420);
  });

  it("reads a written 不断页 as the switch it is and gives the page no break", () => {
    // 授权 46 份的正文写 200 枚 `<w:pageBreakBefore w:val="0"/>`，0 枚说「要」(`.scratch/pbb-spellings.ts`)，且 200 枚
    // 全是 `w:val="0"` 一种拼法（`.scratch/ppr-ladder-referee.ts` 按偏移数，18 份 `.dotm`＝9 个模板名各挂 zh_CN/zh_HK
    // 两份，`.scratch/pbb-ext-split.ts`；`.docx`/`.dotx` 一枚都不写），而样式部件一枚都不写，所以这一格没有「从链上
    // 借来」的一段（`.scratch/toggle-lent-chain.ts`）。把元素的存在读成要求，就等于给说了「不断页」的段各插一页——
    // 200 枚里 180 枚是模型段落，修之前 `page-break-before:always` 画的正是这 180 枚（另 20 枚在 `w:sdt` 里按字节
    // 带走，本来就不由这一格摆页）；装好的壳对 `office rental contract.dotm`（一件 21 枚）现在数到 0
    // （`.scratch/pbb-dom-count.js`）。
    const bytes = makeChainDocx(
      [{ id: "Normal", isDefault: true }],
      [
        { pPr: '<w:pageBreakBefore w:val="0"/>', text: "说了不断页" },
        { pPr: "<w:pageBreakBefore/>", text: "说了要断页" },
        { text: "什么都没说" },
      ],
    );
    const doc = open(bytes);
    const byText = (needle: string): Paragraph => paragraphs(doc).find((x) => textOf(x) === needle)!;
    expect(byText("说了不断页").pageBreakBefore).toBe(false);
    expect(byText("说了要断页").pageBreakBefore).toBe(true);
    expect(byText("什么都没说").pageBreakBefore).toBe(false);
    expect(renderWord(doc)).toContain("page-break-before:always");
    expect([...renderWord(doc).matchAll(/page-break-before:always/g)]).toHaveLength(1);
    // 认领一次不改动它：那一句是文件自己说的，既不升级成裸元素，也不整枚消失。
    const off = byText("说了不断页");
    syncParaSlots(off);
    expect(off.pPrSlots?.find((s) => s.tag === "pageBreakBefore")?.xml).toBe('<w:pageBreakBefore w:val="0"/>');
    const after = part(writeDocx(doc));
    expect(paraXml(after, "说了不断页")).toContain('<w:pageBreakBefore w:val="0"/>');
    expect(paraXml(after, "说了要断页")).toContain("<w:pageBreakBefore/>");
    expect(paraXml(after, "什么都没说")).not.toContain("pageBreakBefore");
  });

  it("costs a paragraph nothing when a character changes somewhere else in the document", () => {
    const before = part(makeParaPropsDocx());
    const after = edited(makeParaPropsDocx());
    for (const needle of ["整段自述", "断点被删", "格式改动", "带框与分节"]) {
      expect(pPrOf(after, needle)).toBe(pPrOf(before, needle));
    }
    // Nothing at all was dropped, including the frame and the section break no field can hold.
    expect(after).toContain('<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:x="108" w:y="1"/>');
    expect(after).toContain("<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>");
  });

  it("does not write a paragraph a property only its style lends it", () => {
    const doc = open(makeParaPropsDocx());
    const styled = paragraphs(doc).find((p) => textOf(p).includes("靠样式的对齐"))!;
    expect(styled.align).toBe("center");
    expect(styled.asRead?.align).toBe("center");
    const after = edited(makeParaPropsDocx());
    expect(pPrOf(after, "靠样式的对齐")).toBe('<w:pStyle w:val="Heading1"/>');
    expect(paraXml(after, "靠样式的对齐")).not.toContain("<w:jc");
  });

  it("moves the alignment the toolbar owns into the carried set and leaves the rest where it sat", () => {
    const doc = open(makeParaPropsDocx());
    const p = paragraphs(doc)[0]!;
    p.align = "center";
    syncParaSlots(p);
    const out = paraXml(save(doc), "整段自述");
    expect(out).toContain('<w:jc w:val="center"/>');
    // `CT_PPr` puts `w:jc` after `w:contextualSpacing` and before `w:textDirection`/`w:textAlignment`.
    expect(out).toMatch(/<w:contextualSpacing\/><w:jc w:val="center"\/><w:textAlignment w:val="baseline"\/><w:outlineLvl/);
    expect(out).toContain('<w:highlight w:val="yellow"/>');
    expect(out).toContain('<w:ind w:left="420" w:right="120" w:firstLine="0"/>');
  });

  it("keeps a mark's character properties when its own record is decided either way", () => {
    // Accepting folds the break away: the two paragraphs become one, and the surviving paragraph still says
    // 24pt and keep-with-next, because those belong to it and not to the record being decided.
    const accepted = open(makeParaPropsDocx());
    expect(accepted.resolveRevisions("accept", 7)).toBe(1);
    expect(paragraphs(accepted).map(textOf)).toContain("断点被删被并进来的下段");
    const out = paraXml(save(accepted), "断点被删");
    expect(out).not.toContain("<w:del");
    expect(out).toContain('<w:sz w:val="24"/>');
    expect(out).toContain("<w:keepNext/>");

    // Refusing puts the break back and takes the record with it — the mark's font and the paragraph's
    // `w:keepNext` are not part of that record, and rebuilding `w:pPr` from fields used to take them down too.
    const refused = open(makeParaPropsDocx());
    expect(refused.resolveRevisions("reject", 7)).toBe(1);
    const kept = paraXml(save(refused), "断点被删");
    expect(kept).not.toContain("<w:del");
    expect(kept).toContain('<w:sz w:val="24"/>');
    expect(kept).toContain("<w:keepNext/>");
    expect(paragraphs(open(writeDocx(refused))).map(textOf)).toContain("被并进来的下段");
  });

  it("puts back the whole set a refused record names, properties no field can hold included", () => {
    const doc = open(makeParaPropsDocx());
    expect(doc.resolveRevisions("reject", 9)).toBe(1);
    const out = save(doc);
    // A record holds the previous set as it stood, so refusing puts that set back as it was written: the
    // alignment it spells, the shading only it names, and the `w:snapToGrid` no field can hold.
    expect(pPrOf(out, "格式改动")).toBe(
      '<w:keepNext/><w:shd w:val="clear" w:color="auto" w:fill="E2EFDA"/><w:snapToGrid w:val="0"/><w:jc w:val="left"/>',
    );
    expect(paraXml(out, "格式改动")).not.toContain("pPrChange");

    // Accepting keeps the set as it stands and takes only the record with it.
    const yes = open(makeParaPropsDocx());
    expect(yes.resolveRevisions("accept", 9)).toBe(1);
    expect(pPrOf(save(yes), "格式改动")).toBe(
      '<w:keepNext/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:jc w:val="both"/>',
    );
  });

  it("keeps a record's previous set through a round trip so a later refusal still has it", () => {
    const doc = open(writeDocx(open(makeParaPropsDocx())));
    const p = paragraphs(doc).find((x) => textOf(x).includes("格式改动"))!;
    expect(p.revFormat?.prevSlots?.map((s) => s.tag)).toEqual(["keepNext", "shd", "snapToGrid", "jc"]);
    expect(doc.resolveRevisions("reject", 9)).toBe(1);
    expect(pPrOf(save(doc), "格式改动")).toContain('w:fill="E2EFDA"');
  });

  it("carries a paragraph's properties through every round trip, not just the first", () => {
    const once = writeDocx(open(makeParaPropsDocx()));
    const twice = writeDocx(open(once));
    expect(part(twice)).toBe(part(once));
    expect(paragraphs(open(twice))[0]!.pPrSlots?.map((s) => s.tag)).toHaveLength(17);
  });

  it("still spells a paragraph the editor invented straight from its fields", () => {
    const doc = WordDocument.blank();
    const p = paragraphs(doc)[0]!;
    expect(p.pPrSlots).toBeUndefined();
    p.align = "center";
    p.markRev = { kind: "ins", id: 5, author: "顾峰", date: "2026-04-02T10:00:00Z" };
    expect(part(writeDocx(doc))).toContain(
      '<w:pPr><w:jc w:val="center"/><w:rPr><w:ins w:id="5" w:author="顾峰" w:date="2026-04-02T10:00:00Z"/></w:rPr></w:pPr>',
    );
  });
});
