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

/**
 * `w:docGrid` (§17.3.2.9) and `w:snapToGrid` (§17.3.1.87), the line grid the files ask for and the paragraphs
 * they take back off it. Counted over the 201 WPS-written documents now on disk (list sha256 `25bcf9f82306`;
 * 235 `w:sectPr`, 27,852 `w:p` of which 159 are self-closed, 27,511 opened `w:pPr`): the sections come from
 * `.scratch/wps-grid-audience.mjs`, the paragraphs from `.scratch/wps-snap-where.mjs`, which walks the ancestor
 * stack and attributes every element to the **property slot** it sits in instead of slicing blocks — the slot is
 * what decides whether it is a paragraph's say at all.
 *
 * - **234 of the 235 sections state a grid, and every one of those says `w:type="lines"`.** `w:linePitch` comes
 *   with all of them: 312 缇 (20.8px) 211 times, 326 (21.733px) 16, 319 (21.267px) 7. No section says
 *   `default`, `linesAndChars` or `snapToColumns`, and every `w:charSpace` in the corpus says `0`.
 * - 0 of the 201 files state two different grids across their own sections, so one page draws one grid.
 * - 3,394 of the 27,511 `w:pPr` blocks state `w:snapToGrid` at the paragraph's own slot and **1,994 of them state
 *   it off**; the other 24,117 say nothing, which is the switch being on (§17.3.1.87 defaults to on). 72 more of
 *   the corpus's 3,466 elements sit in a run-property slot (`w:pPr>w:rPr` 19, `w:p>w:r>w:rPr` 53), which is not a
 *   paragraph's say and which this reader cannot see. `word/styles.xml` states it at a lendable slot 275 times
 *   and **267 of those say off** (56 more in a `w:rPr`), and no table style states it at all.
 * - 6,546 of the snapping paragraphs state their own `w:line`: `exact` 5,492 (360 缇 1,928 times, 400 1,546,
 *   560 724), `auto` 1,051, `atLeast` 3. Of the 1,101 that opt out and state a height, 989 say `auto` and 112
 *   `exact`. No paragraph's own `w:pPr` states `w:line="0"` in this cohort.
 *
 * The page drew none of this before: 234/235 of those sections were read for their paper and their margins and
 * their grid was left in the bytes it travels in, so every line on every page came out at the App's own
 * `line-height: 1.62` on 12pt = 25.92px — 24.6% taller than the 20.8px the files ask 211 of them to measure in.
 */
describe("docx 行网格（w:docGrid 与 w:snapToGrid）：按节的间距量每一行", () => {
  const GRID = '<w:docGrid w:type="lines" w:linePitch="312"/>';
  /** 312 缇 = 20.8px，语料里 211/234 枚间距都是这个数。 */
  const PITCH = 20.8;

  it("reads the grid the section states, and none at all when it states nothing", () => {
    const doc = open(makeSpacingDocx([{ text: "正文" }], "", GRID));
    expect(doc.sections[0]?.docGrid).toEqual({ type: "lines", linePitch: 312 });
    // A section that never wrote the element is owed no grid: the page then draws the heights the file states.
    expect(open(makeSpacingDocx([{ text: "正文" }])).sections[0]?.docGrid).toBeUndefined();
  });

  it("measures a paragraph that states no height in whole grid lines", () => {
    // The corpus's dominant case: 24,117 of its 27,511 `w:pPr` blocks say nothing about the grid, which is the
    // switch being on.
    const doc = open(makeSpacingDocx([{ text: "没说行高" }], "", GRID));
    expect(drawnLineHeight(doc, "没说行高")).toBe(`${PITCH}px`);
  });

  it("rounds a 固定值 up to the next whole grid line rather than drawing the number alone", () => {
    // 1,928 of the 5,492 snapped 固定值 are 360 缇 = 24px on a 20.8px grid: two lines, not one and eleven
    // twentieths. This is what 对齐到网格 does to a fixed height, and the file's own opt-out below is the knob.
    const doc = open(
      makeSpacingDocx(
        [
          { spacing: 'w:line="360" w:lineRule="exact"', text: "固定值 18 磅" },
          { spacing: 'w:line="240" w:lineRule="exact"', text: "固定值 12 磅" },
          { spacing: 'w:line="1200" w:lineRule="exact"', text: "固定值 60 磅" },
        ],
        "",
        GRID,
      ),
    );
    expect(drawnLineHeight(doc, "固定值 18 磅")).toBe(`${PITCH * 2}px`);
    expect(drawnLineHeight(doc, "固定值 12 磅")).toBe(`${PITCH}px`);
    expect(drawnLineHeight(doc, "固定值 60 磅")).toBe(`${PITCH * 4}px`);
  });

  it("keeps the height a paragraph was told not to snap", () => {
    // 1,994 paragraphs state `w:val="0"` at their own slot and 1,101 of those state a height with it: the off is
    // a statement, and the row the file asks for is the row it gets — exactly the numbers #72 drew.
    const doc = open(
      makeSpacingDocx(
        [
          { spacing: 'w:line="360" w:lineRule="exact"', snapToGrid: 'w:val="0"', text: "不吸网格的固定值" },
          { snapToGrid: 'w:val="false"', text: "不吸网格没说自己行高" },
        ],
        "",
        GRID,
      ),
    );
    expect(byText(doc, "不吸网格的固定值").snapToGrid).toBe(false);
    expect(drawnLineHeight(doc, "不吸网格的固定值")).toBe("24px");
    expect(drawnLineHeight(doc, "不吸网格没说自己行高")).toBe("");
  });

  it("takes a bare w:snapToGrid as the switch stated on, not as silence", () => {
    // 1,400 of the corpus's 3,394 paragraph-slot statements are this spelling — an element with no `w:val`.
    const doc = open(makeSpacingDocx([{ snapToGrid: "", text: "光秃秃的一句吸网格" }], "", GRID));
    expect(byText(doc, "光秃秃").snapToGrid).toBe(true);
    expect(drawnLineHeight(doc, "光秃秃")).toBe(`${PITCH}px`);
  });

  it("reads the grid switch only where the paragraph says it, not where a run does", () => {
    // 72 of the corpus's 3,466 `w:snapToGrid` elements sit in a `w:rPr`: 19 in the paragraph mark's own
    // properties, 53 in a run's. §17.3.1.87 puts the switch on the paragraph, so a file that writes it on the
    // mark or on the glyphs has not taken that paragraph off the grid — reading it there would put 72
    // paragraphs on the opt-out path nobody asked for. The bytes still have to come back.
    const doc = open(
      makeSpacingDocx(
        [
          { spacing: 'w:line="360" w:lineRule="exact"', markRPr: '<w:rPr><w:snapToGrid w:val="0"/></w:rPr>', text: "段落标记里写的不吸" },
          { spacing: 'w:line="360" w:lineRule="exact"', runRPr: '<w:snapToGrid w:val="0"/>', text: "run里写的不吸" },
        ],
        "",
        GRID,
      ),
    );
    for (const needle of ["段落标记里", "run里"]) {
      expect(byText(doc, needle).snapToGrid).toBeUndefined();
      expect(drawnLineHeight(doc, needle)).toBe(`${PITCH * 2}px`);
    }
    const out = save(doc);
    expect(paraXml(out, "段落标记里")).toContain('<w:rPr><w:snapToGrid w:val="0"/></w:rPr>');
    expect(paraXml(out, "run里")).toContain('<w:rPr><w:snapToGrid w:val="0"/></w:rPr>');
  });

  it("gives a 最小值 the grid's multiple as its floor and a 多倍 the taller of the two", () => {
    const doc = open(
      makeSpacingDocx(
        [
          { spacing: 'w:line="363" w:lineRule="atLeast"', text: "最小值 18.15 磅" },
          { spacing: 'w:line="360" w:lineRule="auto"', text: "1.5 倍" },
        ],
        "",
        GRID,
      ),
    );
    // 24.2px on a 20.8px grid is two lines, and the floor keeps `1em` under it so the glyphs cannot be cut.
    expect(drawnLineHeight(doc, "最小值")).toBe(`max(${PITCH * 2}px, 1em)`);
    // A multiplier is measured against each line's own text, which only the browser knows, so the grid can only
    // be a floor here: one line, or that many em, whichever is taller.
    expect(drawnLineHeight(doc, "1.5 倍")).toBe(`max(${PITCH}px, 1.5em)`);
  });

  it("lends the off down the style chain, which is where most of it is written", () => {
    // 267 of the 275 statements `word/styles.xml` makes at a lendable slot are the off, and the paragraph under
    // one says nothing.
    const doc = open(
      makeSpacingDocx([{ pStyle: "Lent", spacing: 'w:line="360" w:lineRule="exact"', text: "行距靠样式、网格也靠样式" }], '<w:snapToGrid w:val="0"/>', GRID),
    );
    expect(byText(doc, "行距靠样式").snapToGrid).toBe(false);
    expect(drawnLineHeight(doc, "行距靠样式")).toBe("24px");
    // The paragraph's own statement beats the style's, both ways.
    const both = open(
      makeSpacingDocx([{ pStyle: "Lent", snapToGrid: "", text: "样式说不吸，段落自己说吸" }], '<w:snapToGrid w:val="0"/>', GRID),
    );
    expect(byText(both, "样式说不吸").snapToGrid).toBe(true);
    expect(drawnLineHeight(both, "样式说不吸")).toBe(`${PITCH}px`);
  });

  it("draws the heights it always drew when the section asks for no grid", () => {
    // The negative control for everything drawn before this round: every fixture in this file — and every page
    // the DOM probes pack — spells a `w:sectPr` with no `w:docGrid`, so no existing line height may move.
    const doc = open(
      makeSpacingDocx([
        { spacing: 'w:line="360" w:lineRule="exact"', text: "固定值 18 磅" },
        { spacing: 'w:line="360" w:lineRule="auto"', text: "1.5 倍" },
        { spacing: 'w:line="578" w:lineRule="exact"', text: "固定值 28.9 磅" },
        { text: "没说行高" },
      ]),
    );
    expect(drawnLineHeight(doc, "固定值 18 磅")).toBe("24px");
    expect(drawnLineHeight(doc, "1.5 倍")).toBe("1.50");
    expect(drawnLineHeight(doc, "固定值 28.9 磅")).toBe("38.53333333333333px");
    expect(drawnLineHeight(doc, "没说行高")).toBe("");
  });

  it("draws no grid for a section that asks for a grid without a pitch, or a grid that is not a line grid", () => {
    // §17.3.2.9 leaves `w:linePitch` optional and `w:type` defaulting to `default`. Neither is a line box the
    // page could measure, and the corpus states neither: 234/234 state a pitch and all say `lines`.
    for (const sect of ['<w:docGrid w:type="lines"/>', '<w:docGrid w:type="lines" w:linePitch="0"/>', '<w:docGrid w:type="linesAndChars" w:linePitch="312"/>']) {
      const doc = open(makeSpacingDocx([{ text: "网格说不出间距" }], "", sect));
      expect(drawnLineHeight(doc, "网格说不出间距")).toBe("");
    }
  });

  it("puts the section's grid back as it came and never claims a borrowed opt-out", () => {
    const doc = open(
      makeSpacingDocx(
        [
          { spacing: 'w:line="360" w:lineRule="exact"', snapToGrid: 'w:val="0"', text: "自己说了不吸" },
          { pStyle: "Lent", spacing: 'w:line="360" w:lineRule="exact"', text: "从样式借的不吸" },
        ],
        '<w:snapToGrid w:val="0"/>',
        GRID,
      ),
    );
    const out = save(doc);
    expect(out).toContain(GRID);
    expect(paraXml(out, "自己说了不吸")).toContain('<w:snapToGrid w:val="0"/>');
    // The field is the paragraph's resolved set, not its own statement: a save that wrote the borrowed off onto
    // the paragraph would make the file claim a property the user never touched.
    expect(paraXml(out, "从样式借的")).not.toContain("snapToGrid");
  });
});
