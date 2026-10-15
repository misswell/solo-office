import { describe, expect, it } from "vitest";
import type { Paragraph } from "../src/core/docx.js";
import { ownParaStyle, paraStyleOf, syncParaSlots, WordDocument } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeChainDocx } from "./helpers/make-paragraph.js";

/**
 * §17.7.2's ladder: the paragraph's own `w:pPr`, then the style its `w:pStyle` names, then that style's
 * `w:basedOn` ancestors in order, then the default paragraph style, property by property — each one from the
 * first level that states it, so a style that leaves a field blank passes the ancestor's value through rather
 * than blanking it.
 *
 * The scale of what that was costing is `.scratch/basedon-chain.ts` over the 86 WPS-written documents:
 *
 * - 1,182 of the 1,337 paragraph styles carry `w:basedOn`, running 1 to 4 styles deep (155 / 1,011 / 169 / 2).
 *   No chain circles and none points at a style the part does not declare, and no `w:pStyle` in the 46,258
 *   paragraphs dangles.
 * - 879 paragraphs in 70 of the files ask for a property only an ancestor states: 501 line heights (every one
 *   of them `w:exact`, so a `px` height the browser was not drawing at all), 327 alignments, 220 indents
 *   (122 left, 65 right, 33 first-line), 115 段前段后 gaps.
 * - 38,475 paragraphs name no style, and 3,147 of those take a property from the default paragraph style
 *   (1,956 alignments, 613 `w:exact` heights, 478 of each gap, 466 first-line indents, 284 right indents).
 *   A further 32 paragraphs name a style whose whole chain states nothing while the default style states a
 *   first-line indent, which is why the default sits beneath the chain and not only beneath a style-less
 *   paragraph.
 * - `w:docDefaults/w:pPrDefault` states no paragraph property in any of the 86, so that last rung is not read.
 *
 * What the write side owes the same files is the other half: 159 paragraphs carry no `w:pPr` at all, 91 of them
 * in files whose default style states a property (`.scratch/para-no-pr.ts`). Reading the chain into their fields
 * must not hand those paragraphs a `w:pPr` they never wrote.
 */

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));

const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const byText = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
/** The paragraph's own `w:pPr`, as written — empty string when the save gave it none. */
const pPrXml = (xml: string, needle: string): string => /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>|<w:pPr\/>/.exec(paraXml(xml, needle))?.[0] ?? "";
const drawnStyle = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<(?:div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>(?:(?!<\/(?:div|h[1-6])>)[\s\S])*<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return / style="([^"]*)"/.exec(m[0])?.[1] ?? "";
  }
  return "";
};
const drawn = (doc: WordDocument, needle: string, prop: string): string =>
  new RegExp(`(?:^|;)${prop}:([^;]*)`).exec(drawnStyle(doc, needle))?.[1] ?? "";

/** The three-level chain the corpus's 501 inherited heights come in: leaf silent, middle indents, root 固定值. */
const CHAIN = [
  { id: "Root", pPr: '<w:spacing w:line="578" w:lineRule="exact" w:before="240"/><w:jc w:val="right"/>' },
  { id: "Mid", basedOn: "Root", pPr: '<w:ind w:left="420" w:firstLine="480"/><w:spacing w:after="60"/>' },
  { id: "Leaf", basedOn: "Mid" },
];

describe("docx 样式链解析：w:basedOn 整链、默认段落样式兜底、写出侧不无中生有", () => {
  it("reads a property the paragraph's own style leaves to its ancestors", () => {
    // The 879 paragraphs: nothing in the named style, the height two levels up and the unit with it.
    const doc = open(makeChainDocx(CHAIN, [{ pStyle: "Leaf", text: "三层链" }]));
    const p = byText(doc, "三层链");
    expect(p.lineSpacing).toBe(578);
    expect(p.lineSpacingRule).toBe("exact");
    expect(p.spacingBefore).toBe(240);
    expect(p.spacingAfter).toBe(60);
    expect(p.align).toBe("right");
    expect(p.indentLeft).toBe(420);
    expect(p.indentFirstLine).toBe(480);
    // 固定值 28.9 磅 at the third hop went from the browser's default line height to a pixel height.
    expect(drawn(doc, "三层链", "line-height")).toBe("38.53333333333333px");
    expect(drawn(doc, "三层链", "margin-left")).toBe("28px");
    expect(drawn(doc, "三层链", "text-indent")).toBe("32px");
  });

  it("takes each property from the first style in the chain that states it", () => {
    // §17.7.2's per-property rule: the leaf's own 段后 is the paragraph's 段后 even though the root states one,
    // and the root's 段前 still comes through where the leaf is silent.
    const doc = open(
      makeChainDocx(
        [
          { id: "Root", pPr: '<w:spacing w:after="999" w:before="240"/><w:jc w:val="right"/>' },
          { id: "Leaf", basedOn: "Root", pPr: '<w:jc w:val="center"/><w:spacing w:after="120"/>' },
        ],
        [{ pStyle: "Leaf", text: "谁先写谁" }],
      ),
    );
    const p = byText(doc, "谁先写谁");
    expect(p.align).toBe("center");
    expect(p.spacingAfter).toBe(120);
    expect(p.spacingBefore).toBe(240);
  });

  it("falls back on the default paragraph style for a paragraph that names no style", () => {
    // 3,147 of the corpus's 38,475 style-less paragraphs, 1,956 of them an alignment and 613 a 固定值 height.
    const doc = open(
      makeChainDocx(
        [
          { id: "Normal", isDefault: true, pPr: '<w:jc w:val="center"/><w:spacing w:line="578" w:lineRule="exact" w:before="240" w:after="120"/><w:ind w:firstLine="480"/>' },
          { id: "Other" },
        ],
        [{ text: "没写样式" }],
      ),
    );
    const p = byText(doc, "没写样式");
    expect(p.styleId).toBeUndefined();
    expect(p.align).toBe("center");
    expect(p.lineSpacing).toBe(578);
    expect(p.lineSpacingRule).toBe("exact");
    expect(p.indentFirstLine).toBe(480);
    expect(drawn(doc, "没写样式", "text-align")).toBe("center");
    expect(drawn(doc, "没写样式", "line-height")).toBe("38.53333333333333px");
  });

  it("puts the default paragraph style beneath a named chain that never reaches it", () => {
    // The 32 paragraphs in 16 files whose named style has no `w:basedOn` and states no first-line indent while
    // the default style does. Everything the chain states still wins.
    const doc = open(
      makeChainDocx(
        [
          { id: "Normal", isDefault: true, pPr: '<w:ind w:firstLine="480"/><w:jc w:val="center"/>' },
          { id: "Alone", pPr: '<w:spacing w:after="60"/><w:jc w:val="left"/>' },
        ],
        [{ pStyle: "Alone", text: "链不通到默认样式" }],
      ),
    );
    const p = byText(doc, "链不通到默认样式");
    expect(p.indentFirstLine).toBe(480);
    expect(p.spacingAfter).toBe(60);
    expect(p.align).toBe("left");
  });

  it("walks a chain whose own ancestor is the default style once, not twice", () => {
    const doc = open(
      makeChainDocx(
        [
          { id: "Normal", isDefault: true, pPr: '<w:jc w:val="center"/><w:spacing w:line="360" w:lineRule="auto"/>' },
          { id: "Child", basedOn: "Normal" },
        ],
        [{ pStyle: "Child", text: "链里已经有默认样式" }],
      ),
    );
    expect(byText(doc, "链里已经有默认样式").align).toBe("center");
    expect(drawn(doc, "链里已经有默认样式", "line-height")).toBe("1.50");
  });

  it("keeps what the paragraph states itself against the chain, and a stated zero against a height three hops up", () => {
    const doc = open(
      makeChainDocx(CHAIN, [
        { pStyle: "Leaf", pPr: '<w:jc w:val="left"/><w:ind w:left="0"/>', text: "自己写了居中" },
        { pStyle: "Leaf", pPr: '<w:spacing w:line="0"/>', text: "自己写了个零" },
      ]),
    );
    const own = byText(doc, "自己写了居中");
    expect(own.align).toBe("left");
    expect(own.indentLeft).toBe(0);
    // The root's height still comes through: `w:line` is a different attribute from `w:jc` and `w:left`.
    expect(own.lineSpacing).toBe(578);
    // A stated zero is a statement, not an absence, and no unit turns a zero into a line.
    expect(byText(doc, "自己写了个零").lineSpacing).toBe(0);
    expect(drawn(doc, "自己写了个零", "line-height")).toBe("");
  });

  it("stops at a chain that circles instead of hanging the read", () => {
    // Neither shape occurs in the 86 files (`.scratch/basedon-chain.ts` counts 0 cycles and 0 dangling
    // `w:basedOn`), so this is the specification's own guard rather than a WPS-verified one.
    const doc = open(
      makeChainDocx(
        [
          { id: "A", basedOn: "B", pPr: '<w:spacing w:line="360" w:lineRule="auto"/>' },
          { id: "B", basedOn: "A", pPr: '<w:jc w:val="right"/><w:spacing w:after="60"/>' },
        ],
        [{ pStyle: "A", text: "互指" }],
      ),
    );
    const p = byText(doc, "互指");
    expect(p.lineSpacing).toBe(360);
    expect(p.align).toBe("right");
    expect(p.spacingAfter).toBe(60);
  });

  it("reads a `w:pStyle` the part does not declare as a paragraph that named nothing", () => {
    // Again unmeasured: all 46,258 paragraphs name a declared style or nothing at all.
    const doc = open(
      makeChainDocx([{ id: "Normal", isDefault: true, pPr: '<w:jc w:val="center"/>' }], [{ pStyle: "Ghost", text: "指向不存在的样式" }]),
    );
    const p = byText(doc, "指向不存在的样式");
    expect(p.styleName).toBeUndefined();
    expect(p.align).toBe("center");
  });

  it("lends `w:numPr` as the element states it, and no page break", () => {
    // Numbering is part of the read now: 58 paragraph styles state `w:pPr/w:numPr` and 48 paragraphs take their
    // list that way (`.scratch/style-numpr-census.ts`, `.scratch/num-lent-shapes.ts`). It travels as one property,
    // because where `w:ind` can be filled attribute by attribute this element either says the list or it does not
    // — and `w:numId` absent reads as 0, which is §17.3.4.25's cancellation, the shape of 11 of those 48.
    // `w:pageBreakBefore` is still not part of it: a style's break before is nothing a paragraph here asks for.
    const doc = open(
      makeChainDocx(
        [
          { id: "Normal", isDefault: true, pPr: '<w:pageBreakBefore/><w:jc w:val="center"/>' },
          { id: "Listed", basedOn: "Normal", pPr: '<w:numPr><w:ilvl w:val="2"/><w:numId w:val="7"/></w:numPr>' },
          { id: "HalfNum", basedOn: "Normal", pPr: '<w:numPr><w:ilvl w:val="1"/></w:numPr>' },
        ],
        [{ pStyle: "Listed", text: "编号从样式来" }, { pStyle: "HalfNum", text: "样式只写 w:ilvl" }],
      ),
    );
    const p = byText(doc, "编号从样式来");
    expect(p.numPr).toEqual({ numId: 7, ilvl: 2 });
    expect(p.align).toBe("center");
    expect(p.pageBreakBefore).toBe(false);
    expect(byText(doc, "样式只写 w:ilvl").numPr).toEqual({ numId: 0, ilvl: 1 });
  });

  it("does not let a default character style stand in for the default paragraph style", () => {
    const doc = open(
      makeChainDocx(
        [
          { id: "CharDefault", type: "character", isDefault: true, pPr: '<w:jc w:val="center"/>' },
          { id: "Normal", pPr: '<w:jc w:val="right"/>' },
        ],
        [{ text: "只有字符样式是默认" }],
      ),
    );
    expect(byText(doc, "只有字符样式是默认").align).toBeUndefined();
  });

  it("saves a paragraph that wrote no `w:pPr` without writing one for it", () => {
    // The 159, and the 91 of them in files whose default style states a property: the chain fills the fields and
    // the page draws them, and the file's statement of nothing stays a statement of nothing.
    const bytes = makeChainDocx(
      [{ id: "Normal", isDefault: true, pPr: '<w:jc w:val="center"/><w:spacing w:line="578" w:lineRule="exact" w:after="120"/><w:ind w:firstLine="480"/>' }],
      [{ text: "什么都没写" }],
    );
    const doc = open(bytes);
    expect(byText(doc, "什么都没写").lineSpacing).toBe(578);
    expect(drawn(doc, "什么都没写", "text-indent")).toBe("32px");
    expect(pPrXml(save(doc), "什么都没写")).toBe("");
  });

  it("writes only the property the reader actually moved onto such a paragraph", () => {
    const doc = open(
      makeChainDocx(
        [{ id: "Normal", isDefault: true, pPr: '<w:jc w:val="center"/><w:spacing w:line="578" w:lineRule="exact" w:after="120"/><w:ind w:firstLine="480"/>' }],
        [{ text: "被改成居左" }, { text: "什么都没写" }],
      ),
    );
    const p = byText(doc, "被改成居左");
    p.align = "left";
    syncParaSlots(p);
    expect(pPrXml(save(doc), "被改成居左")).toBe("<w:pPr><w:jc w:val=\"left\"/></w:pPr>");
    // The paragraph next to it, untouched, still writes nothing.
    expect(pPrXml(save(doc), "什么都没写")).toBe("");
  });

  it("leaves an inherited alignment out of a set the paragraph wrote for itself", () => {
    // 46,258 of the corpus's paragraphs carry their own `w:pPr`; a chain-lent property must not slip into the
    // carried set beside the shading and the outline level that only travel as the file's own bytes. The
    // numbering is the sharpest case: it draws on the page, and the paragraph never asked for it.
    const doc = open(
      makeChainDocx(
        [
          {
            id: "Normal",
            isDefault: true,
            pPr: '<w:jc w:val="center"/><w:ind w:firstLine="480"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>',
          },
        ],
        [{ pPr: '<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:outlineLvl w:val="1"/>', text: "自己写了底纹" }],
      ),
    );
    const p = byText(doc, "自己写了底纹");
    expect(p.align).toBe("center");
    expect(p.numPr).toEqual({ numId: 3, ilvl: 0 });
    expect(drawn(doc, "自己写了底纹", "padding-left")).toBe("8px");
    expect(pPrXml(save(doc), "自己写了底纹")).toBe(
      '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:outlineLvl w:val="1"/></w:pPr>',
    );
  });

  it("spells every field of a paragraph the editor built, which has no read snapshot", () => {
    // The shape a paragraph split with Enter arrives as (`fresh()` copies the fields and no snapshot with them),
    // so the second half of a split paragraph says out loud what the first half said.
    const doc = open(
      makeChainDocx(
        [{ id: "Normal", isDefault: true, pPr: '<w:jc w:val="center"/><w:spacing w:line="578" w:lineRule="exact"/><w:ind w:firstLine="480"/>' }],
        [{ text: "原文" }],
      ),
    );
    const src = byText(doc, "原文");
    const split: Paragraph = {
      kind: "paragraph",
      styleId: src.styleId,
      align: src.align,
      indentFirstLine: src.indentFirstLine,
      lineSpacing: src.lineSpacing,
      lineSpacingRule: src.lineSpacingRule,
      pageBreakBefore: false,
      inlines: [],
    };
    doc.blocks.push(split);
    const out = save(doc);
    expect(pPrXml(out, "原文")).toBe("");
    const last = [...out.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)].map((m) => m[0]).at(-1) ?? "";
    expect(/<w:pPr>[\s\S]*?<\/w:pPr>/.exec(last)?.[0]).toBe(
      '<w:pPr><w:spacing w:line="578" w:lineRule="exact"/><w:ind w:firstLine="480"/><w:jc w:val="center"/></w:pPr>',
    );
  });

  it("records only what the paragraph wrote when a reviewer is asked to decide", () => {
    // The steps `setParagraph` takes on a paragraph that carried no `w:pPr`. The properties on it came from the
    // chain, so the record must say the previous set was empty rather than hand the reviewer a set the file
    // never wrote — §17.13.5.19 records the direct formatting, and refusing puts the paragraph back to borrowing.
    const doc = open(
      makeChainDocx(
        [{ id: "Normal", isDefault: true, pPr: '<w:jc w:val="center"/><w:spacing w:line="578" w:lineRule="exact" w:after="120"/>' }],
        [{ text: "什么都没写" }],
      ),
    );
    const p = byText(doc, "什么都没写");
    const was = paraStyleOf(p);
    p.align = "left";
    syncParaSlots(p);
    p.revFormat = { kind: "format", id: 3, author: "顾峰", date: "2026-03-01T08:00:00Z", prev: ownParaStyle(p.asRead, was) };
    expect(paraXml(save(doc), "什么都没写")).toBe(
      '<w:p><w:pPr><w:jc w:val="left"/><w:pPrChange w:id="3" w:author="顾峰" w:date="2026-03-01T08:00:00Z"><w:pPr/></w:pPrChange></w:pPr>' +
        "<w:r><w:t>什么都没写</w:t></w:r></w:p>",
    );
    expect(doc.resolveRevisions("reject", 3)).toBe(1);
    expect(byText(doc, "什么都没写").align).toBe("center");
    expect(byText(doc, "什么都没写").lineSpacing).toBe(578);
    expect(pPrXml(save(doc), "什么都没写")).toBe("");
  });
});

/** The paragraph's opening tag as the page draws it, so a test can read the `data-*` stamps off it. */
const openTag = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<(?:div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>(?:(?!<\/(?:div|h[1-6])>)[\s\S])*<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return m[0].slice(0, m[0].indexOf(">"));
  }
  return "";
};

/**
 * The three page-keeping rules through the same ladder — the case #126 left open. `w:widowControl` makes it the
 * biggest of them, because its default is on: a style that writes the `w:val="0"` does not add a rule to the
 * paragraphs that borrow it, it removes one every one of them already had.
 *
 * `.scratch/wps-keep-lent.mjs` over the 115 WPS-written documents with a table that are on disk now: 113 of them
 * state `w:widowControl` in `word/styles.xml` (220 times, 160 of those saying off) and 10,205 paragraphs take
 * that off from the chain while their own `w:pPr` says nothing; `w:keepNext` and `w:keepLines` appear 58 times
 * apiece in the styles and lend 83 paragraphs each. `w:docDefaults/w:pPrDefault` states none of the three in any
 * of the 115, so the rung under the default paragraph style still has nothing to lend here.
 */
describe("docx 保持规则从样式链借：三条开关走同一阶梯，孤行控制反向的缺省也照借", () => {
  it("借到样式说的那句关：段落自己没说，data-widow 照样盖上", () => {
    const doc = open(
      makeChainDocx(
        [
          { id: "Off", pPr: '<w:widowControl w:val="0"/>' },
          { id: "Keep", pPr: "<w:keepNext/><w:keepLines/>" },
        ],
        [{ pStyle: "Off", text: "样式说关" }, { pStyle: "Keep", text: "样式说保" }],
      ),
    );
    expect(byText(doc, "样式说关").widowControl).toBe(false);
    expect(openTag(doc, "样式说关")).toContain('data-widow="0"');
    expect(byText(doc, "样式说保").keepNext).toBe(true);
    expect(byText(doc, "样式说保").keepLines).toBe(true);
    expect(openTag(doc, "样式说保")).toContain('data-keep-next="1"');
    expect(openTag(doc, "样式说保")).toContain('data-keep-lines="1"');
    expect(openTag(doc, "样式说保")).not.toContain("data-widow");
  });

  it("沿 w:basedOn 往上找：祖父说的关算到孙子段落头上", () => {
    const doc = open(
      makeChainDocx(
        [
          { id: "Root", pPr: '<w:widowControl w:val="none"/>' },
          { id: "Mid", basedOn: "Root", pPr: '<w:jc w:val="center"/>' },
          { id: "Leaf", basedOn: "Mid" },
        ],
        [{ pStyle: "Leaf", text: "隔两级借" }],
      ),
    );
    expect(byText(doc, "隔两级借").widowControl).toBe(false);
    expect(byText(doc, "隔两级借").align).toBe("center");
  });

  it("子样式裸标签把祖先那句关说回开：取第一个说的，不合并", () => {
    const doc = open(
      makeChainDocx(
        [
          { id: "Off", pPr: '<w:widowControl w:val="0"/>' },
          { id: "Take", basedOn: "Off", pPr: "<w:widowControl/>" },
        ],
        [{ pStyle: "Take", text: "取回开" }, { pStyle: "Off", text: "还是关" }],
      ),
    );
    expect(byText(doc, "取回开").widowControl).toBe(true);
    expect(openTag(doc, "取回开")).not.toContain("data-widow");
    expect(byText(doc, "还是关").widowControl).toBe(false);
  });

  it("段落自己写的压过链，包括把链借来的开取消掉", () => {
    const doc = open(
      makeChainDocx(
        [{ id: "On", pPr: "<w:widowControl/><w:keepNext/>" }],
        [{ pStyle: "On", pPr: '<w:widowControl w:val="0"/><w:keepNext w:val="0"/>', text: "自己写的" }],
      ),
    );
    expect(byText(doc, "自己写的").widowControl).toBe(false);
    expect(byText(doc, "自己写的").keepNext).toBe(false);
    expect(openTag(doc, "自己写的")).toContain('data-widow="0"');
    expect(openTag(doc, "自己写的")).not.toContain("data-keep-next");
  });

  it("默认段落样式说关：连 w:pPr 都没有的段落也跟着关", () => {
    const doc = open(
      makeChainDocx([{ id: "Normal", isDefault: true, pPr: '<w:widowControl w:val="0"/>' }], [{ text: "什么都没写" }]),
    );
    expect(byText(doc, "什么都没写").widowControl).toBe(false);
    expect(openTag(doc, "什么都没写")).toContain('data-widow="0"');
  });

  it("借来的关不抄成段落自己的直接格式，样式那一侧一字不动", () => {
    const bytes = makeChainDocx([{ id: "Off", pPr: '<w:widowControl w:val="0"/>' }], [{ pStyle: "Off", text: "借来的关" }]);
    const out = save(open(bytes));
    expect(paraXml(out, "借来的关")).toBe(
      '<w:p><w:pPr><w:pStyle w:val="Off"/></w:pPr><w:r><w:t>借来的关</w:t></w:r></w:p>',
    );
    const stylesOf = (b: Uint8Array): string => text(OpcPackage.from(b).bytes("word/styles.xml")!);
    expect(stylesOf(writeDocx(open(bytes)))).toBe(stylesOf(bytes));
  });

  it("段落自己写的那句关往返不丢、也不被复制成第二枚", () => {
    const doc = open(
      makeChainDocx([{ id: "Off", pPr: '<w:widowControl w:val="0"/>' }], [
        { pStyle: "Off", pPr: '<w:widowControl w:val="0"/>', text: "自己说关" },
        { pStyle: "Off", text: "借来说关" },
      ]),
    );
    const out = save(doc);
    expect((pPrXml(out, "自己说关").match(/<w:widowControl/g) ?? []).length).toBe(1);
    expect(pPrXml(out, "借来说关")).not.toContain("widowControl");
    expect(byText(open(writeDocx(doc)), "自己说关").widowControl).toBe(false);
  });

  it("回车拆段：上一段自己写的关跟着走，只从链借来的留给新样式说", () => {
    const own = open(
      makeChainDocx(
        [{ id: "Head", next: "Body" }, { id: "Body" }],
        [{ pStyle: "Head", pPr: '<w:widowControl w:val="0"/>', text: "自己写的关" }],
      ),
    );
    expect(own.nextSplit(byText(own, "自己写的关"))!.widowControl).toBe(false);
    const borrowed = open(
      makeChainDocx(
        [{ id: "Head2", next: "Body2", pPr: '<w:widowControl w:val="0"/>' }, { id: "Body2" }],
        [{ pStyle: "Head2", text: "借来的关" }],
      ),
    );
    const split = borrowed.nextSplit(byText(borrowed, "借来的关"))!;
    expect(split.styleId).toBe("Body2");
    expect(split.widowControl).toBeUndefined();
  });
});
