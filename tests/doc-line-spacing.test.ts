import { describe, expect, it } from "vitest";
import type { Paragraph } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeSpacingDocx, makeSpacingRevisionDocx } from "./helpers/make-paragraph.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));

const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const byText = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
/** The paragraph's own `w:spacing`, as written. Every fixture here spells it self-closing, so the tag is the element. */
const spacingXml = (xml: string, needle: string): string => /<w:spacing[^>]*\/>/.exec(paraXml(xml, needle))?.[0] ?? "";
/**
 * The `style` the page puts on the paragraph carrying `needle`, or "" when it puts none. The line height can
 * only show up on the paragraph's own element, so this is what "the page draws it" means here.
 */
const drawnStyle = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<(?:div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>(?:(?!<\/(?:div|h[1-6])>)[\s\S])*<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return / style="([^"]*)"/.exec(m[0])?.[1] ?? "";
  }
  return "";
};
/** The height as the page drew it, attribute and all. */
const drawnLineHeight = (doc: WordDocument, needle: string): string =>
  /(?:^|;)line-height:([^;]*)/.exec(drawnStyle(doc, needle))?.[1] ?? "";

/**
 * `w:spacing@w:line` and its `w:lineRule`, as WPS writes them, with the numbers it writes. Counted over the
 * 25,801 `w:spacing` elements that state a height or a gap in the 86 documents this repository's docx shapes
 * come from (`.scratch/spacing-census.ts`; 38,427 `w:spacing` elements in all, 12,617 of them the
 * character-spacing kind inside `w:rPr`):
 *
 * - `w:lineRule` comes out as `w:auto` 13,396 times, `w:exact` 5,812, `w:atLeast` 991, and 5,602 elements
 *   state 段前段后 with no line at all. Nothing else appears — §17.3.1.27's four-value list is three values here.
 * - **Every element that states a positive `w:line` states its rule** (there is no "height with no unit" to
 *   resolve against the spec's default), so 6,802 of the 20,198 stated heights are a fixed value or a minimum,
 *   counted in twips rather than in the 240ths of a line the page used for every one of them.
 * - The heights cluster on `auto` 240 (12,353 — single spacing written out loud), 360 (536), 480 (148),
 *   180 (80); `exact` 240 (2,968), 578 (1,518), 280 (910); `atLeast` 363 (448), 270 (330).
 *
 * `.scratch/spacing-lent.ts` counted where the height comes *from*: 5,110 paragraphs state no `w:line` while
 * the style named by their `w:pStyle` does (4,847 auto, 247 exact, 16 atLeast), and the styles themselves
 * state a positive height 473 times (234 auto, 200 exact, 39 atLeast).
 *
 * Not modelled: `w:beforeLines`/`w:afterLines` (stated 2,572 and 2,581 times, but positive in only 128 and 135
 * of those and `0` in 2,444 and 2,446; 120 of the positive pairs work out to a "line" of 100 twips, which is a
 * grid pitch and not a line height — `w:docGrid@w:linePitch` is 312 twips in these files) and
 * `w:beforeAutospacing`/`w:afterAutospacing` (66 and 54 state `1`, 4,850 and 4,852 state `0`). Those attributes
 * still travel as the file's own bytes.
 */
describe("docx 行距（w:spacing@w:line 与 w:lineRule）：按文件写的单位读、画、带回", () => {
  it("reads the height together with the unit the file put on it", () => {
    const doc = open(
      makeSpacingDocx([
        { spacing: 'w:line="578" w:lineRule="exact"', text: "固定值 28.9 磅" },
        { spacing: 'w:line="363" w:lineRule="atLeast"', text: "最小值 18.15 磅" },
        { spacing: 'w:line="360" w:lineRule="auto"', text: "1.5 倍" },
      ]),
    );
    expect(byText(doc, "固定值").lineSpacing).toBe(578);
    expect(byText(doc, "固定值").lineSpacingRule).toBe("exact");
    expect(byText(doc, "最小值").lineSpacingRule).toBe("atLeast");
    expect(byText(doc, "1.5 倍").lineSpacingRule).toBe("auto");
  });

  it("leaves a set that states no line without a height or a unit", () => {
    // 5,602 of the elements state 段前段后 only. Their absence has to stay absence: a paragraph holding
    // `w:line` and one holding none are different statements, and the second is what lets a style's through.
    const doc = open(makeSpacingDocx([{ spacing: 'w:before="120" w:after="60"', text: "只有空段" }]));
    const p = byText(doc, "只有空段");
    expect(p.spacingBefore).toBe(120);
    expect(p.spacingAfter).toBe(60);
    expect(p.lineSpacing).toBeUndefined();
    expect(p.lineSpacingRule).toBeUndefined();
    expect(drawnLineHeight(doc, "只有空段")).toBe("");
  });

  it("draws a rule it does not know as no rule at all, and keeps the bytes", () => {
    // Unrecognised literals stay unread rather than guessed at; the element itself still travels verbatim, so
    // a save does not quietly rewrite someone else's unit.
    const doc = open(makeSpacingDocx([{ spacing: 'w:line="360" w:lineRule="wordCount"', text: "没见过的规则" }]));
    const p = byText(doc, "没见过");
    expect(p.lineSpacing).toBe(360);
    expect(p.lineSpacingRule).toBeUndefined();
    expect(drawnLineHeight(doc, "没见过的规则")).toBe("1.50");
    expect(spacingXml(save(doc), "没见过的规则")).toBe('<w:spacing w:line="360" w:lineRule="wordCount"/>');
  });

  it("draws `w:auto` as the multiple of a line it is", () => {
    const doc = open(
      makeSpacingDocx([
        { spacing: 'w:line="240" w:lineRule="auto"', text: "单倍行距 12353 例里的一档" },
        { spacing: 'w:line="360" w:lineRule="auto"', text: "1.50 倍" },
        { spacing: 'w:line="480" w:lineRule="auto"', text: "2.00 倍" },
        { spacing: 'w:line="180" w:lineRule="auto"', text: "0.75 倍" },
      ]),
    );
    expect(drawnLineHeight(doc, "单倍行距")).toBe("1.00");
    expect(drawnLineHeight(doc, "1.50 倍")).toBe("1.50");
    expect(drawnLineHeight(doc, "2.00 倍")).toBe("2.00");
    expect(drawnLineHeight(doc, "0.75 倍")).toBe("0.75");
  });

  it("draws 固定值 as that height instead of a multiple of a line", () => {
    // The defect this is here for: 28.9pt was drawn as 2.41 倍行距, and 16pt as exactly single, because the
    // twips were divided by 240. 15 twips to the pixel: 578 → 38.53px, 240 → 16px, 280 → 18.67px.
    const doc = open(
      makeSpacingDocx([
        { spacing: 'w:line="578" w:lineRule="exact"', text: "固定值 28.9 磅" },
        { spacing: 'w:line="240" w:lineRule="exact"', text: "固定值 12 磅" },
        { spacing: 'w:line="280" w:lineRule="exact"', text: "固定值 14 磅" },
      ]),
    );
    expect(drawnLineHeight(doc, "固定值 28.9")).toBe("38.53333333333333px");
    expect(drawnLineHeight(doc, "固定值 12 磅")).toBe("16px");
    expect(drawnLineHeight(doc, "固定值 14 磅")).toBe("18.666666666666668px");
  });

  it("draws 最小值 as the greater of the stated height and one line of glyphs", () => {
    // CSS has no line height that says "this much, or whatever the text needs". `1em` is the least height that
    // cannot clip the paragraph's own glyphs; 未取证 how much taller WPS lets such a line get, since that is a
    // question about font metrics and this page has none.
    const doc = open(
      makeSpacingDocx([
        { spacing: 'w:line="363" w:lineRule="atLeast"', text: "最小值 18.15 磅" },
        { spacing: 'w:line="270" w:lineRule="atLeast"', text: "最小值 13.5 磅" },
      ]),
    );
    expect(drawnLineHeight(doc, "最小值 18.15")).toBe("max(24.2px, 1em)");
    expect(drawnLineHeight(doc, "最小值 13.5")).toBe("max(18px, 1em)");
  });

  it("takes the height and its unit from the paragraph's style when the paragraph states neither", () => {
    // 5,110 paragraphs in the corpus ask for a line height this way and used to get none at all.
    const auto = open(makeSpacingDocx([{ pStyle: "Lent", text: "行距靠样式" }], '<w:spacing w:line="360" w:lineRule="auto"/>'));
    expect(byText(auto, "行距靠样式").lineSpacing).toBe(360);
    expect(drawnLineHeight(auto, "行距靠样式")).toBe("1.50");

    const exact = open(makeSpacingDocx([{ pStyle: "Lent", text: "行距靠样式" }], '<w:spacing w:line="578" w:lineRule="exact"/>'));
    expect(byText(exact, "行距靠样式").lineSpacingRule).toBe("exact");
    expect(drawnLineHeight(exact, "行距靠样式")).toBe("38.53333333333333px");
  });

  it("keeps the height the paragraph states itself, and the zero it states, against the style's", () => {
    const doc = open(
      makeSpacingDocx(
        [
          { pStyle: "Lent", spacing: 'w:line="240" w:lineRule="exact"', text: "自己写死" },
          { pStyle: "Lent", spacing: 'w:line="0" w:lineRule="atLeast"', text: "自己写了个零" },
        ],
        '<w:spacing w:line="363" w:lineRule="atLeast"/>',
      ),
    );
    expect(drawnLineHeight(doc, "自己写死")).toBe("16px");
    // A stated `w:line="0"` is a stated value, not an absence: it beats the style (as the indent fields do) and
    // draws no height. The corpus states one exactly once, beside `w:atLeast`.
    expect(byText(doc, "自己写了个零").lineSpacing).toBe(0);
    expect(drawnLineHeight(doc, "自己写了个零")).toBe("");
  });

  it("carries every attribute of the set the file wrote through a save", () => {
    // Including the ones this model does not read: 按行计的段前段后 (positive in 128/135 of the elements that
    // state them) and `w:beforeAutospacing` (66 elements state `1`) are the file's own bytes, and a save that
    // re-spelled `w:spacing` from the fields would drop them along with the paragraph's shading and outline
    // level.
    const stated =
      'w:beforeAutospacing="1" w:beforeLines="50" w:before="120" w:afterLines="50" w:after="60" w:line="578" w:lineRule="exact"';
    const doc = open(makeSpacingDocx([{ spacing: stated, text: "整份自述" }]));
    const p = byText(doc, "整份自述");
    expect(p.lineSpacing).toBe(578);
    expect(p.lineSpacingRule).toBe("exact");
    expect(p.spacingBefore).toBe(120);
    expect(spacingXml(save(doc), "整份自述")).toBe(`<w:spacing ${stated}/>`);
  });

  it("spells a paragraph the editor invented from its fields, unit included", () => {
    const doc = WordDocument.blank();
    const p = paragraphs(doc)[0]!;
    p.lineSpacing = 578;
    p.lineSpacingRule = "exact";
    p.spacingBefore = 120;
    expect(part(writeDocx(doc))).toContain('<w:spacing w:before="120" w:line="578" w:lineRule="exact"/>');

    // A height with no rule of its own is the common case this writer signs, and §17.3.1.27's default is
    // `w:auto`, which is also what the 13,396 auto elements say.
    const plain = WordDocument.blank();
    const q = paragraphs(plain)[0]!;
    q.lineSpacing = 360;
    expect(part(writeDocx(plain))).toContain('<w:spacing w:line="360" w:lineRule="auto"/>');
  });

  it("puts back the unit a refused record named, and hands the borrowed pair back when the record omits both", () => {
    const doc = open(makeSpacingRevisionDocx());
    expect(doc.resolveRevisions("reject", 14)).toBe(1);
    // Refusing puts 28.9pt 固定值 back: the height comes with the unit the record wrote it in, or the page would
    // draw the same number as 2.41 倍行距 again.
    expect(drawnLineHeight(doc, "行距被改过")).toBe("38.53333333333333px");
    expect(spacingXml(save(doc), "行距被改过")).toBe('<w:spacing w:line="578" w:lineRule="exact"/>');

    // The second paragraph's record states no spacing at all, so the height goes back the way it was read: from
    // the style, unit and all. A record that dropped the height but left the rule behind would be a lie.
    const other = open(makeSpacingRevisionDocx());
    expect(other.resolveRevisions("reject", 15)).toBe(1);
    expect(byText(other, "行距靠样式也被改过").lineSpacing).toBe(363);
    expect(byText(other, "行距靠样式也被改过").lineSpacingRule).toBe("atLeast");
    expect(drawnLineHeight(other, "行距靠样式也被改过")).toBe("max(24.2px, 1em)");
    expect(spacingXml(save(other), "行距靠样式也被改过")).toBe("");

    // Accepting keeps the set as it stands, unit included.
    const yes = open(makeSpacingRevisionDocx());
    expect(yes.resolveRevisions("accept", 14)).toBe(1);
    expect(drawnLineHeight(yes, "行距被改过")).toBe("1.50");
    expect(spacingXml(save(yes), "行距被改过")).toBe('<w:spacing w:line="360" w:lineRule="auto"/>');
  });
});
