import { describe, expect, it } from "vitest";
import { syncParaSlots, WordDocument, type Paragraph } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import {
  DEFAULT_STYLE,
  EMPHASIS_CHAR_STYLE,
  HEADING_STYLE,
  QUOTE_STYLE,
  makeStyleGalleryDocx,
} from "./helpers/make-style-gallery.js";
import { unstamped } from "./helpers/unstamp.js";

/**
 * 开始页那一格「样式」：把选中的段落换成**件里自己声明的**某一枚段落样式。
 *
 * 换样式不是把新样式那一套属性抄到段落身上——那是「这一段自己写了十几句」，再换回去也撤不掉。落笔只有一句
 * `w:pStyle`，其余的让给新链；段落自己写下的直接格式（槽位看得见的、与旧链不同的）留着，run 自己的 `w:rPr`
 * 与字符样式借来的那几句也留着。纸上读到的字则必须跟着换：腿画在每枚 run 的 `styleOf` 上，不重新落到字段上，
 * 「换成标题 1」就只挪了段距、字还是正文那么大。
 *
 * 授权 46 份共 112 枚段落样式（每份 6～27 枚，`.scratch/para-style-gallery-audience.py`），103 枚带 `w:qFormat`，
 * 默认那枚 6/6 都叫 `Normal`；夹具照这个形状造了一枚默认、一枚标题链、一枚引用、一枚字符样式。
 */

const bytes = (): Uint8Array => makeStyleGalleryDocx();
const part = (b: Uint8Array): string => text(OpcPackage.from(b).bytes("word/document.xml")!);
const stylesOf = (b: Uint8Array): string => text(OpcPackage.from(b).bytes("word/styles.xml")!);
const open = (b: Uint8Array): WordDocument => WordDocument.openBytes(b);
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const byText = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
const pPrXml = (xml: string, needle: string): string => /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>|<w:pPr\/>/.exec(paraXml(xml, needle))?.[0] ?? "";
const runXml = (xml: string, needle: string): string => /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/.exec(paraXml(xml, needle))?.[0] ?? "";
const paraHtml = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<(?:div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>[\s\S]*?<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return m[0];
  }
  return "";
};
const openTag = (doc: WordDocument, needle: string): string => {
  const html = paraHtml(doc, needle);
  return html.slice(0, html.indexOf(">") + 1);
};
// 字体名里那个引号在 HTML 里是 `&quot;`，实体自带的 `;` 会把 CSS 声明切开，所以先还原再读。
const decodeCss = (css: string): string => css.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&");
const drawn = (doc: WordDocument, needle: string, prop: string): string =>
  new RegExp(`(?:^|;)${prop}:([^;]*)`).exec(decodeCss(/ style="([^"]*)"/.exec(openTag(doc, needle))?.[1] ?? ""))?.[1] ?? "";
const runDrawn = (doc: WordDocument, needle: string, prop: string): string => {
  const span = /<span style="([^"]*)"/.exec(paraHtml(doc, needle))?.[1] ?? "";
  return new RegExp(`(?:^|;)${prop}:([^;]*)`).exec(decodeCss(span))?.[1] ?? "";
};

/** 换样式并让槽位跟字段说同一句——编辑器 `setParagraphStyle` 落笔的那一步。 */
const switchTo = (doc: WordDocument, needle: string, styleId: string | undefined): void => {
  const p = byText(doc, needle);
  expect(doc.applyParaStyle(p, styleId)).toBe(true);
  syncParaSlots(p);
};

describe("开始页「样式」画廊：换的是 w:pStyle，不是把新样式抄到段落身上", () => {
  it("只列件里声明的段落样式，按声明顺序，默认那枚标出来；字符样式不在里面", () => {
    expect(open(bytes()).paragraphStyles()).toEqual([
      { id: DEFAULT_STYLE, name: "Normal", isDefault: true },
      { id: HEADING_STYLE, name: "标题 1", isDefault: false },
      { id: QUOTE_STYLE, name: "引用", isDefault: false },
    ]);
    expect(open(bytes()).paragraphStyles().some((s) => s.id === EMPHASIS_CHAR_STYLE)).toBe(false);
  });

  it("甲组 换成标题 1：盘上只多一句 w:pStyle，纸上换成 16 磅 Arial 深蓝粗体", () => {
    const d = open(bytes());
    switchTo(d, "甲组", HEADING_STYLE);
    expect(paraXml(part(writeDocx(d)), "甲组")).toBe(
      '<w:p><w:pPr><w:pStyle w:val="2"/></w:pPr><w:r><w:t>甲组 什么都没说的一段</w:t></w:r></w:p>',
    );
    const p = byText(d, "甲组");
    expect(p.styleId).toBe(HEADING_STYLE);
    expect(p.outlineLevel).toBe(0);
    expect(p.spacingBefore).toBe(340);
    expect(p.keepNext).toBe(true);
    // 字得重新落到 run 上：纸上画的是 run 自己的 `styleOf`，段落只带一个字号。
    expect(p.inlines[0].sizePt).toBe(16);
    expect(p.inlines[0].bold).toBe(true);
    expect(p.inlines[0].color).toBe("#1F3864");
    expect(p.inlines[0].font).toBe("Arial");
    expect(p.inlines[0].fontEastAsia).toBe("黑体");
    expect(openTag(d, "甲组").startsWith("<h1 ")).toBe(true);
    expect(openTag(d, "甲组")).toContain('data-outline="1"');
    expect(openTag(d, "甲组")).toContain('data-keep-next="1"');
    expect(drawn(d, "甲组", "font-size")).toBe("16pt");
    expect(drawn(d, "甲组", "margin-top")).toBe("22.666666666666668px");
    expect(runDrawn(d, "甲组", "font-size")).toBe("16pt");
    expect(runDrawn(d, "甲组", "font-weight")).toBe("600");
    expect(runDrawn(d, "甲组", "color")).toBe("#1F3864");
    expect(runDrawn(d, "甲组", "font-family")).toContain("Arial");

    const again = open(writeDocx(d));
    const back = byText(again, "甲组");
    expect(back.styleId).toBe(HEADING_STYLE);
    expect(back.outlineLevel).toBe(0);
    expect(back.inlines[0].sizePt).toBe(16);
    expect(back.inlines[0].color).toBe("#1F3864");
  });

  it("乙组 自己写的缩进与粗体留着，标题借来的段距不写成它自己说的", () => {
    const d = open(bytes());
    switchTo(d, "乙组", HEADING_STYLE);
    const xml = part(writeDocx(d));
    expect(paraXml(xml, "乙组")).toBe(
      '<w:p><w:pPr><w:pStyle w:val="2"/><w:ind w:left="720"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>乙组 自己写了缩进的一段</w:t></w:r></w:p>',
    );
    expect(pPrXml(xml, "乙组")).not.toContain("<w:spacing");
    const p = byText(d, "乙组");
    expect(p.indentLeft).toBe(720);
    expect(p.spacingBefore).toBe(340);
    expect(drawn(d, "乙组", "margin-left")).toBe("48px");
    expect(drawn(d, "乙组", "margin-top")).toBe("22.666666666666668px");

    const again = open(writeDocx(d));
    expect(byText(again, "乙组").indentLeft).toBe(720);
    expect(byText(again, "乙组").spacingBefore).toBe(340);
  });

  it("丙组 已是标题，换回正文：整枚 w:pPr 撤掉，而不是点名 Normal", () => {
    const d = open(bytes());
    switchTo(d, "丙组", undefined);
    expect(paraXml(part(writeDocx(d)), "丙组")).toBe('<w:p><w:r><w:t>丙组 已是标题的一段</w:t></w:r></w:p>');
    const p = byText(d, "丙组");
    expect(p.styleId).toBeUndefined();
    expect(p.outlineLevel).toBeUndefined();
    expect(p.spacingBefore).toBeUndefined();
    expect(openTag(d, "丙组").startsWith("<div ")).toBe(true);
    expect(openTag(d, "丙组")).not.toContain("data-outline");
    expect(drawn(d, "丙组", "margin-top")).toBe("");
    expect(byText(open(writeDocx(d)), "丙组").styleId).toBeUndefined();
  });

  it("丁组 换成引用：字符样式的脸与墨留着，斜体从段落链借", () => {
    const d = open(bytes());
    switchTo(d, "丁组", QUOTE_STYLE);
    const xml = part(writeDocx(d));
    expect(pPrXml(xml, "丁组")).toBe('<w:pPr><w:pStyle w:val="3"/></w:pPr>');
    expect(runXml(xml, "丁组")).toBe(
      '<w:r><w:rPr><w:rStyle w:val="7"/></w:rPr><w:t>丁组 挂着字符样式的一段</w:t></w:r>',
    );
    const run = byText(d, "丁组").inlines[0];
    expect(run.italic).toBe(true);
    expect(run.font).toBe("Georgia");
    expect(run.color).toBe("#C00000");
    expect(drawn(d, "丁组", "margin-left")).toBe("37.8px");
    expect(runDrawn(d, "丁组", "font-style")).toBe("italic");
    expect(runDrawn(d, "丁组", "font-family")).toContain("Georgia");
    expect(runDrawn(d, "丁组", "color")).toBe("#C00000");

    const back = open(writeDocx(d)).blocks.find(
      (b): b is Paragraph => b.kind === "paragraph" && b.inlines.some((r) => (r.text ?? "").includes("丁组")),
    )!;
    expect(back.inlines[0].italic).toBe(true);
    expect(back.inlines[0].font).toBe("Georgia");
  });

  it("换样式不动 word/styles.xml：件里那几枚样式一字未改", () => {
    const source = bytes();
    const d = open(source);
    switchTo(d, "甲组", HEADING_STYLE);
    switchTo(d, "丙组", undefined);
    switchTo(d, "丁组", QUOTE_STYLE);
    expect(stylesOf(writeDocx(d))).toBe(stylesOf(source));
  });

  it("点名一个件里没有的样式、或本来就是那个样式：什么都不改，也不弄脏", () => {
    const source = bytes();
    const d = open(source);
    const p = byText(d, "甲组");
    expect(d.applyParaStyle(p, "NoSuchStyle")).toBe(false);
    expect(d.applyParaStyle(p, undefined)).toBe(false);
    expect(d.applyParaStyle(byText(d, "丙组"), HEADING_STYLE)).toBe(false);
    expect(d.edited).toBe(false);
    expect(unstamped(writeDocx(d))).toEqual(unstamped(source));
  });
});
