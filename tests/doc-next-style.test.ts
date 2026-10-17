import { describe, expect, it } from "vitest";
import { WordDocument, type Paragraph } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeOutlineDocx, pStyle } from "./helpers/make-outline.js";
import { unstamped } from "./helpers/unstamp.js";

/**
 * §17.7.2.24 `w:next` — the style a paragraph takes when Enter splits the one before it, which is what makes
 * 「标题 1 后回车落回正文」 the behaviour it is.
 *
 * Measured over the 86 WPS-written documents before anything here changed (`.scratch/next-resolve.ts`):
 *
 * - 668 of their 2,378 styles state a next style, and **all 668** targets are declared paragraph styles; none
 *   names the style that states it. A further 358 state nothing while a `w:basedOn` ancestor states one, so the
 *   question is the same chain the formatting ladder already walks.
 * - it is not a decoration: 17,299 of their 30,322 paragraphs sit under a style whose next differs from their
 *   own, and for all of them the two chains state different paragraph properties — the split would draw wrong
 *   without the switch, and 928 would keep an outline level their new style does not have.
 * - the default paragraph style is the exception, and it is the reason `nextStyleOf` takes the *named* style and
 *   not the resolved one: 63 of the 86 declare a next on the default style, 16,783 paragraphs name no style, and
 *   924 of those would land on a heading style. No file looks like it happened: of 230 adjacent paragraph style
 *   changes 0 match the preceding style's `w:next`, and of 934 paragraphs that name a style 0 name a default
 *   style's next. Whether WPS itself ignores a default style's next is 未取证 — the rule here is the one the
 *   corpus's own writes are consistent with, and the one that cannot turn body text into a heading.
 */

const doc = (): WordDocument => WordDocument.openBytes(makeOutlineDocx([{ pPr: pStyle("T1"), text: "标题" }, { text: "正文" }]));
const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);

describe("w:next：样式说「回车之后换成谁」", () => {
  it("样式自己声明的 next 就是它", () => {
    expect(doc().nextStyleOf("T1")).toBe("Body1");
  });

  it("自己没有 next 时沿 w:basedOn 往上取第一个声明的", () => {
    const d = doc();
    expect(d.nextStyleOf("Borrow")).toBe("Body1"); // 一跳：Borrow → T1
    expect(d.nextStyleOf("Deep")).toBe("Body1"); // 两跳：Deep → T2 → T1
    expect(d.styleDef("Deep")?.next).toBeUndefined();
  });

  it("段落没点名样式时不换——默认样式的 next 不落到正文上", () => {
    const d = doc();
    // 语料里 63 份的默认样式确实声明了 next（这份 fixture 的 `Normal` 也声明），而 16,783 个不点名样式的段落
    // 一旦跟着换，其中 924 个会变成标题。这里段落点名了 `Normal`，才按它自己的声明换。
    expect(d.nextStyleOf(undefined)).toBeUndefined();
    expect(d.nextStyleOf("Normal")).toBe("Body1");
  });

  it("next 指到没声明的样式、或指回自己，等于没说", () => {
    const d = doc();
    expect(d.nextStyleOf("Dangling")).toBeUndefined();
    expect(d.nextStyleOf("SelfNext")).toBeUndefined();
    expect(d.nextStyleOf("NoSuchStyle")).toBeUndefined();
  });

  it("点名了一个不存在于本文件的样式，也不换", () => {
    expect(doc().nextStyleOf("T1")).toBe("Body1");
    expect(WordDocument.openBytes(makeOutlineDocx([{ text: "只有正文" }])).nextStyleOf(undefined)).toBeUndefined();
  });

  it("读 next 不把文档弄脏：没编辑过的件保存还是原字节", () => {
    const bytes = makeOutlineDocx([{ pPr: pStyle("T1"), text: "标题" }]);
    const d = WordDocument.openBytes(bytes);
    expect(d.nextStyleOf("T1")).toBe("Body1");
    expect(d.paragraphStyleOf("Body1")?.align).toBe("both");
    expect(d.edited).toBe(false);
    expect(unstamped(writeDocx(d))).toEqual(unstamped(bytes));
  });
});

describe("拆出来的后一段：只说样式，不抄上一段借来的格式", () => {
  /**
   * The split the page's Enter produces: `WordDocument.nextSplit` is the rule itself, and the editor's `fresh`
   * calls it for the element the browser just split (`.scratch/next-dom.ts` measures that the new element comes
   * back without a `data-bid`, so that is the branch a real keystroke reaches). Here the text is put back by hand
   * because the page would supply it.
   */
  const splitLike = (d: WordDocument, previous: Paragraph): Paragraph => {
    const split = d.nextSplit(previous)!;
    split.inlines = [{ text: "后一段" }];
    d.blocks.push(split);
    return split;
  };

  it("后一段写出的是 pStyle 一个 child，正文的对齐与行距仍从样式来", () => {
    const d = doc();
    expect((d.blocks[0] as Paragraph).styleId).toBe("T1");
    splitLike(d, d.blocks[0] as Paragraph);
    const xml = part(writeDocx(d));
    const mine = [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes("后一段")) ?? "";
    expect(/<w:pPr(?:\s[^>]*)?><w:pStyle w:val="Body1"\/><\/w:pPr>/.test(mine)).toBe(true);
    expect(mine).not.toContain("<w:jc");
    expect(mine).not.toContain("<w:spacing");
  });

  it("保存再读回来：后一段是正文样式，大纲级别不再跟着上一段", () => {
    const d = doc();
    const before = d.blocks[0] as Paragraph;
    expect(before.outlineLevel).toBe(0);
    splitLike(d, before);
    const again = WordDocument.openBytes(writeDocx(d));
    const last = again.blocks[again.blocks.length - 1] as Paragraph;
    expect(last.styleId).toBe("Body1");
    expect(last.outlineLevel).toBeUndefined();
    // 大纲窗格按 1 起数（`w:outlineLvl w:val="0"` 是一级标题），所以清单里只剩上一段那一条。
    expect(again.outline().map((o) => o.level)).toEqual([1]);
  });

  it("上一段自己写过的直接格式跟着走，样式链借来的不跟", () => {
    // 语料的形状：标题自己直接写了 `w:ind w:left="720"`（不是从样式借的）。拆段不改直接格式，所以后一段还带着
    // 那个缩进；借来的对齐与行距不跟，那是上一个样式的话。
    const d = WordDocument.openBytes(makeOutlineDocx([{ pPr: `${pStyle("T1")}<w:ind w:left="720"/>`, text: "标题" }]));
    expect((d.blocks[0] as Paragraph).align).toBeUndefined();
    splitLike(d, d.blocks[0] as Paragraph);
    const last = d.blocks[d.blocks.length - 1] as Paragraph;
    expect(last.indentLeft).toBe(720);
    expect(last.align).toBe("both");
    const mine = [...part(writeDocx(d)).matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)]
      .map((m) => m[0])
      .find((s) => s.includes("后一段")) ?? "";
    expect(mine).toContain('<w:pStyle w:val="Body1"/>');
    expect(mine).toContain('<w:ind w:left="720"/>');
    expect(mine).not.toContain("<w:jc");
  });

  it("上一段的「段前分页」不跟到新段——跟着走就是每按一次回车多一页", () => {
    const d = WordDocument.openBytes(makeOutlineDocx([{ pPr: `${pStyle("T1")}<w:pageBreakBefore/>`, text: "标题" }]));
    expect((d.blocks[0] as Paragraph).pageBreakBefore).toBe(true);
    splitLike(d, d.blocks[0] as Paragraph);
    expect((d.blocks[d.blocks.length - 1] as Paragraph).pageBreakBefore).toBeUndefined();
    const mine = [...part(writeDocx(d)).matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)]
      .map((m) => m[0])
      .find((s) => s.includes("后一段")) ?? "";
    expect(mine).not.toContain("<w:pageBreakBefore");
  });
});
