import { describe, expect, it } from "vitest";
import { syncParaSlots, WordDocument, type Paragraph } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { NUMBERING, makeNumberingDocx, numPr, numberingStyles } from "./helpers/make-numbering.js";

/**
 * What a numbered paragraph draws. Three obligations, each measured over the 86 WPS-written documents before
 * anything here changed (`.scratch/num-lent-shapes.ts`, `.scratch/num-from-style.ts`, `.scratch/numfmt-census.ts`):
 *
 * - the list a paragraph joins can come from its style. 48 paragraphs borrow `w:numPr` down the §17.7.2 chain,
 *   all 48 from the style their `w:pStyle` names rather than from an ancestor, and none of them showed a list on
 *   the page before. 37 of the 48 draw a symbol now; the other 11 take a `w:numPr` that writes only `w:ilvl`
 *   (2 of the corpus's 58 style-side `w:numPr` elements, `.scratch/style-numpr-census.ts`), and since the
 *   property is lent whole that reads as numId 0, which §17.3.4.25 says is no list at all.
 * - `w:numId w:val="0"` is the cancellation, and the corpus's favourite: 517 of its 981 own-`w:numPr` paragraphs
 *   are one, 86 of those signed `w:ilvl="-1"`. Every one of them used to draw a bullet, listing a paragraph the
 *   file had taken out of the list.
 * - the symbol is the file's `w:lvlText` template (§17.3.1.65), and all 617 `w:lvl` elements state one. The "."
 *   this renderer used to append is the template's to state: 、 for 91 of the 92 `chineseCounting` levels, ")"
 *   for 72 of the 75 `lowerLetter` ones, and only 84 of the 205 `decimal` ones.
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

const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const byText = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
const pPrOf = (xml: string, needle: string): string =>
  /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>|<w:pPr\/>/.exec(paraXml(xml, needle))?.[0] ?? "";

const drawnPara = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<(?:div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>(?:(?!<\/(?:div|h[1-6])>)[\s\S])*<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return m[0];
  }
  return "";
};
/** The symbol in front of the paragraph that says `needle`, or "" when it stands in no list. */
const drawn = (doc: WordDocument, needle: string): string =>
  /<span class="bullet"(?:\s[^>]*)?>([\s\S]*?)<\/span>/.exec(drawnPara(doc, needle))?.[1] ?? "";
/** The whole mark element, so a test can see what the level put on it. */
const markOf = (doc: WordDocument, needle: string): string =>
  /<span class="bullet"[^>]*>[\s\S]*?<\/span>/.exec(drawnPara(doc, needle))?.[0] ?? "";
const propOf = (doc: WordDocument, needle: string, prop: string): string =>
  new RegExp(`(?:^|;)${prop}:([^;]*)`).exec(/ style="([^"]*)"/.exec(drawnPara(doc, needle))?.[1] ?? "")?.[1] ?? "";

/** The numbering part as text: the one part this reader never regenerates, so a save owes it byte for byte. */
const numberingOf = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/numbering.xml")!);

/** `count` paragraphs joining the same list, each spelling its own `w:numPr`. */
const list = (count: number, at: string, prefix: string): { pPr?: string; text: string }[] =>
  Array.from({ length: count }, (_, i) => ({ pPr: at, text: `${prefix}${i + 1}` }));

describe("docx 编号：符号按 w:lvlText 画、w:numId 0 不画、列表可从样式借来", () => {
  it("draws the file's own template for 中文小写 instead of a digit and a dot", () => {
    const doc = open(makeNumberingDocx([
      { pPr: numPr(1, 0), text: "第一项" },
      { pPr: numPr(1, 0), text: "第二项" },
      { pPr: numPr(1, 0), text: "第三项" },
    ]));
    expect(doc.levelOf(1, 0)).toMatchObject({ numFmt: "chineseCounting", lvlText: "%1、", start: 1 });
    expect(drawn(doc, "第一项")).toBe("一、");
    expect(drawn(doc, "第二项")).toBe("二、");
    expect(drawn(doc, "第三项")).toBe("三、");
    expect(propOf(doc, "第一项", "padding-left")).toBe("8px");
  });

  it("counts 中文小写 the way the enumeration spells it, from wherever w:start says to begin", () => {
    // The level starts at 10, so twelve paragraphs are 十 through 二十一: 十 on its own, 二十 with no second 二,
    // 二十一 with it. A table of the nine digits would hand back decimal digits from 十 upward.
    const doc = open(makeNumberingDocx(list(12, numPr(2, 0), "第")));
    expect(Array.from({ length: 12 }, (_, i) => drawn(doc, `第${i + 1}`))).toEqual([
      "十、", "十一、", "十二、", "十三、", "十四、", "十五、", "十六、", "十七、", "十八、", "十九、", "二十、", "二十一、",
    ]);
  });

  it("takes the punctuation, the format and the starting number from the level", () => {
    const doc = open(makeNumberingDocx([
      { pPr: numPr(3, 0), text: "上层" },
      { pPr: numPr(3, 1), text: "首项" },
      { pPr: numPr(3, 1), text: "次项" },
      { pPr: numPr(5, 1), text: "圈一" },
      { pPr: numPr(5, 1), text: "圈二" },
    ]));
    // `w:lvlText w:val="%2)"` over `w:start w:val="2"`: the letter is the level's, the ")" is the template's, and
    // the first item drawn at that level is b rather than a.
    expect(drawn(doc, "首项")).toBe("b)");
    expect(drawn(doc, "次项")).toBe("c)");
    expect(drawn(doc, "上层")).toBe("1");
    expect(propOf(doc, "首项", "padding-left")).toBe("32px");
    // A level that states no `w:start` begins at 1 (§17.3.4.29 makes it optional), and
    // `decimalEnclosedCircleChinese` is the circled digit rather than the decimal one.
    expect(drawn(doc, "圈一")).toBe("①");
    expect(drawn(doc, "圈二")).toBe("②");
  });

  it("fills a %N that names another level from that level's sequence, or from its w:start", () => {
    const doc = open(makeNumberingDocx([
      { pPr: numPr(4, 0), text: "上层一" },
      { pPr: numPr(4, 0), text: "上层二" },
      { pPr: numPr(4, 1), text: "下层一" },
      { pPr: numPr(4, 1), text: "下层二" },
    ]));
    // `%1.%2` over a level that began at 3: before the dot is what that level has reached, after it the sub-level
    // restarts because each `w:ilvl` keeps its own sequence.
    expect(drawn(doc, "上层一")).toBe("3");
    expect(drawn(doc, "上层二")).toBe("4");
    expect(drawn(doc, "下层一")).toBe("4.1");
    expect(drawn(doc, "下层二")).toBe("4.2");

    // Nothing above it has been drawn, so %1 is that level's `w:start` rather than a 0.
    const alone = open(makeNumberingDocx([{ pPr: numPr(4, 1), text: "只有下层" }]));
    expect(drawn(alone, "只有下层")).toBe("3.1");
  });

  it("draws nothing for `w:numId w:val=\"0\"`, in every shape the corpus cancels with", () => {
    // The three spellings are the measured ones: numId 0 with no `w:ilvl` (4 of the 981), with `w:ilvl="0"` (427),
    // and with `w:ilvl="-1"` (86). All three drew a bullet and an indent: a list member the file refused.
    const doc = open(makeNumberingDocx([
      { pPr: '<w:numPr><w:numId w:val="0"/></w:numPr>', text: "只写 numId 0" },
      { pPr: numPr(0, 0), text: "ilvl 0 与 numId 0" },
      { pPr: numPr(0, -1), text: "ilvl -1 与 numId 0" },
      { pPr: numPr(1, 0), text: "真正的列表项" },
    ]));
    for (const cancelled of ["只写 numId 0", "ilvl 0 与 numId 0", "ilvl -1 与 numId 0"]) {
      expect(byText(doc, cancelled).numPr?.numId).toBe(0);
      expect(drawn(doc, cancelled)).toBe("");
      expect(drawnPara(doc, cancelled)).not.toContain('class="bullet"');
      expect(propOf(doc, cancelled, "padding-left")).toBe("");
    }
    expect(drawn(doc, "真正的列表项")).toBe("一、");
  });

  it("reads a `w:numPr` that leaves `w:ilvl` out as level 0", () => {
    // The element is optional and the corpus omits it — in 4 of the 981 paragraphs, every one of those a
    // cancellation. Absent means the first level here too, since a paragraph cannot be at no level.
    const doc = open(makeNumberingDocx([{ pPr: numPr(1), text: "省略 ilvl" }]));
    expect(byText(doc, "省略 ilvl").numPr).toEqual({ numId: 1, ilvl: 0 });
    expect(drawn(doc, "省略 ilvl")).toBe("一、");
  });

  it("puts a paragraph in its style's list, taking the whole property from the first level that states it", () => {
    const doc = open(makeNumberingDocx([
      { pPr: '<w:pStyle w:val="Numbered"/>', text: "靠样式进列表" },
      { pPr: '<w:pStyle w:val="Unnumbered"/>', text: "样式里被取消" },
    ]));
    // `Numbered` states `<w:ilvl w:val="1"/><w:numId w:val="1"/>`, so the paragraph is drawn at that level.
    expect(byText(doc, "靠样式进列表").numPr).toEqual({ numId: 1, ilvl: 1 });
    expect(drawn(doc, "靠样式进列表")).toBe("（一）");
    expect(propOf(doc, "靠样式进列表", "padding-left")).toBe("32px");
    // `Unnumbered` borrows from `Numbered` and states `w:ilvl` alone. §17.7.2 takes `w:numPr` as one property, so
    // the ancestor's numId does not ride along with it: absent numId is 0, which is the cancellation.
    expect(byText(doc, "样式里被取消").numPr).toEqual({ numId: 0, ilvl: 1 });
    expect(drawn(doc, "样式里被取消")).toBe("");
  });

  it("saves a borrowed list as the paragraph's own set, without writing the list into it", () => {
    // The write-side half of the lending: the page draws the symbol, the file still says nothing about lists.
    const doc = open(makeNumberingDocx([{ pPr: '<w:pStyle w:val="Numbered"/><w:spacing w:after="60"/>', text: "借来的编号" }]));
    expect(doc.edited).toBe(false);
    const xml = saved(doc);
    expect(pPrOf(xml, "借来的编号")).toBe('<w:pPr><w:pStyle w:val="Numbered"/><w:spacing w:after="60"/></w:pPr>');
    expect(xml).not.toContain("<w:numPr>");
    // And after the round trip the style is still what says it, so the symbol is still there.
    expect(drawn(open(writeDocx(doc)), "借来的编号")).toBe("（一）");
  });

  it("writes back the cancellation a paragraph stated, so a save does not put it back in the list", () => {
    // The read gives the paragraph's own `w:numPr` over the style's; the write has to keep it. Drop the element
    // and only the style's numId stands, handing the paragraph back the list its author closed.
    const doc = open(makeNumberingDocx([
      { pPr: '<w:pStyle w:val="Numbered"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>', text: "取消样式给的列表" },
    ]));
    expect(byText(doc, "取消样式给的列表").numPr).toEqual({ numId: 0, ilvl: 0 });
    expect(drawn(doc, "取消样式给的列表")).toBe("");
    const xml = saved(doc);
    expect(pPrOf(xml, "取消样式给的列表")).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>');
    expect(drawn(WordDocument.openBytes(writeDocx(doc)), "取消样式给的列表")).toBe("");
  });

  it("puts a list the user cancels into the paragraph's own set, at `w:numPr`'s place in the sequence", () => {
    // The step 撤销列表 takes on a paragraph whose list came from its style: `setParagraph` writes the
    // cancellation into the field and `syncParaSlots` moves it into the set the file wrote. The set has to gain
    // the element the paragraph never spelled — otherwise the click does nothing — and lose nothing else.
    const doc = open(makeNumberingDocx([{ pPr: '<w:pStyle w:val="Numbered"/><w:spacing w:after="60"/>', text: "借来的编号" }]));
    const p = byText(doc, "借来的编号");
    p.numPr = { numId: 0, ilvl: p.numPr!.ilvl };
    syncParaSlots(p);
    expect(drawn(doc, "借来的编号")).toBe("");
    const xml = part(writeDocx(doc));
    expect(pPrOf(xml, "借来的编号")).toBe(
      '<w:pPr><w:pStyle w:val="Numbered"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="0"/></w:numPr><w:spacing w:after="60"/></w:pPr>',
    );
    expect(drawn(open(writeDocx(doc)), "借来的编号")).toBe("");
  });

  it("counts a letter sequence as a base-26 list that begins on the first letter", () => {
    // §17.18.4's letter forms are a count, not a modulo: a, b, …, z, then aa. What this replaced started one
    // code point early — U+0060 ("`") on the first item, "@" for uppercase — and wrapped back to a on the 27th.
    // No list in the corpus reaches 27 items at a letter level, so the wrap is the spec's, not a measured shape;
    // the first item is the one every one of the 75 `lowerLetter` levels draws.
    const doc = open(makeNumberingDocx(list(27, numPr(6, 0), "项")));
    expect(drawn(doc, "项1")).toBe("a)");
    expect(drawn(doc, "项26")).toBe("z)");
    expect(drawn(doc, "项27")).toBe("aa)");
    const upper = open(makeNumberingDocx([
      { pPr: numPr(6, 1), text: "大写一" },
      { pPr: numPr(6, 1), text: "大写二" },
    ]));
    expect(drawn(upper, "大写一")).toBe("(A)");
    expect(drawn(upper, "大写二")).toBe("(B)");
  });

  it("keeps each numId its own list, and draws a template that states no placeholder as it stands", () => {
    const doc = open(makeNumberingDocx([
      { pPr: numPr(1, 0), text: "中文一" },
      { pPr: numPr(5, 0), text: "圆点一" },
      { pPr: numPr(5, 0), text: "圆点二" },
      { pPr: numPr(1, 0), text: "中文二" },
    ]));
    // The two bullet paragraphs spend no 中文小写 number, and a template with no %N repeats verbatim.
    expect(drawn(doc, "中文一")).toBe("一、");
    expect(drawn(doc, "圆点一")).toBe("•");
    expect(drawn(doc, "圆点二")).toBe("•");
    expect(drawn(doc, "中文二")).toBe("二、");
  });

  it("exposes the numbering part as the model reads it", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(2, 0), text: "台账" }]));
    expect(doc.numberingTable().map((n) => n.numId)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]);
    // numId 11 is the level that states a `w:hint` and no face: a family is not a font, so `font` stays unset
    // and the mark's own chain is what the page wears (§17.18.39 lists three hint values, none of them a name).
    expect(doc.levelOf(11, 0)).toMatchObject({ numFmt: "chineseCounting", lvlText: "%1、", font: undefined });
    expect(doc.numberingTable().find((n) => n.numId === 3)?.levels[1]).toMatchObject({
      numFmt: "lowerLetter",
      lvlText: "%2)",
      start: 2,
      ilvl: 1,
    });
    // A `w:numId` the part cannot resolve is not this corpus's shape — no paragraph in the 86 asks for one — so
    // the fallback records a decision not to lose the mark rather than a measured symbol: the dot before, the
    // dot after.
    const dangling = open(makeNumberingDocx([{ pPr: numPr(99, 0), text: "无定义" }]));
    expect(dangling.levelOf(99, 0)).toBeUndefined();
    expect(drawn(dangling, "无定义")).toBe("•");
  });
});

describe("列表符号的字体：私用区那颗码点按级自己声明的 w:rFonts 画", () => {
  /** The corpus's own spelling of a Wingdings bullet: the codepoint in `w:lvlText`, the font in the level's `w:rPr`. */
  const MARK = String.fromCharCode(0xf0a7);

  it("w:lvl/w:rPr/w:rFonts 进了模型，符号就穿着它", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(7, 0), text: "方块" }]));
    expect(doc.levelOf(7, 0)).toMatchObject({ numFmt: "bullet", lvlText: MARK, font: "Wingdings" });
    // 级没写字号，所以那一段解析出来的 11 磅（docDefaults 的 w:sz w:val="22"）才是这颗符号的磅数。
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-family:Wingdings;font-size:11pt">${MARK}</span>`);
  });

  it("级没写 rFonts：符号穿这一段解析出的脸，不穿页面样式表那张", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(8, 0), text: "方块" }]));
    expect(doc.levelOf(8, 0)?.font).toBeUndefined();
    // 级没说，就按 §17.3.1.4 往下走：这一段 docDefaults 点了 Calibri，那颗符号就穿它——
    // 页面样式表的 Times New Roman 是 App 的默认，不是这个件说的。
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-family:&quot;Calibri&quot;;font-size:11pt">${MARK}</span>`);
  });

  it("只有符号穿这颗字体的脸，正文那串字照旧", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(7, 0), text: "方块" }]));
    const para = drawnPara(doc, "方块");
    expect((para.match(/font-family:Wingdings/g) ?? []).length).toBe(1);
    expect(/<span class="bullet"[^>]*font-family:Wingdings/.test(para)).toBe(true);
  });

  it("读进模型的字体不改写出侧：段落仍然只说自己那个 w:numPr", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(7, 0), text: "方块" }]));
    expect(pPrOf(saved(doc), "方块")).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>');
    expect(pPrOf(saved(doc), "方块")).not.toContain("Wingdings");
  });
});

describe("列表符号的字号：级上写的 w:sz 只落在那颗符号上", () => {
  const MARK = String.fromCharCode(0xf0a7);

  it("w:sz 按半点读进来，换成磅（与正文 run 同一个换算）", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(9, 0), text: "方块" }, { pPr: numPr(10, 0), text: "点数" }]));
    expect(doc.levelOf(9, 0)).toMatchObject({ numFmt: "bullet", font: "Wingdings", sizePt: 8 });
    expect(doc.levelOf(10, 0)).toMatchObject({ numFmt: "decimal", sizePt: 9 });
    // 级没写 w:sz 就是没说，不是说了 0：numId 7 与 8 都只有字体或什么都没有。
    expect(doc.levelOf(7, 0)?.sizePt).toBeUndefined();
    expect(doc.levelOf(8, 0)?.sizePt).toBeUndefined();
  });

  it("那颗符号比正文小：文件说小一号，页面就画小一号", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(10, 0), text: "点数" }]));
    expect(markOf(doc, "点数")).toBe(`<span class="bullet" style="font-family:&quot;Calibri&quot;;font-size:9pt">1.</span>`);
  });

  it("字体与字号同在一颗符号上，顺序照 rPr 里写的两样", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(9, 0), text: "方块" }]));
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-family:Wingdings;font-size:8pt">${MARK}</span>`);
  });

  it("级上的字号不外溢：这一段自己的字照它的样式链，不跟着小", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(9, 0), text: "方块" }]));
    const para = drawnPara(doc, "方块");
    // docDefaults 的 w:sz w:val="22" 是 11 磅，那颗符号的 8 磅只属于它自己。
    expect((para.match(/font-size:8pt/g) ?? []).length).toBe(1);
    expect(/<span class="bullet"[^>]*font-size:8pt/.test(para)).toBe(true);
    expect((para.match(/font-size:9pt/g) ?? []).length).toBe(0);
  });

  it("写出侧一个字号都不添：numbering.xml 逐字节照旧，段落只说自己那个 w:numPr", () => {
    const src = makeNumberingDocx([{ pPr: numPr(9, 0), text: "方块" }]);
    const doc = open(src);
    const out = writeDocx(doc);
    expect(numberingOf(out)).toBe(numberingOf(src));
    // 逼一次重拼，看它会不会把级上的 8 磅当成段落自己写的。
    expect(paragraphs(doc).length).toBe(1);
    const regenerated = saved(doc);
    expect(pPrOf(regenerated, "方块")).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr>');
    expect((regenerated.match(/<w:sz\b/g) ?? []).length).toBe((part(src).match(/<w:sz\b/g) ?? []).length);
    expect(paraXml(regenerated, "方块")).not.toContain("8pt");
  });
});

describe("那颗符号的字号回退：级没写就用段落标记自己解析出的磅数", () => {
  const MARK = String.fromCharCode(0xf0a7);
  /** The corpus's spelling of a mark that states its size: `w:sz` with a `w:szCs` beside it, 326 of 326. */
  const markSz = (half: number): string => `<w:rPr><w:sz w:val="${half}"/><w:szCs w:val="${half}"/></w:rPr>`;

  it("w:pPr/w:rPr/w:sz 进了模型：级没说，那颗符号用段落的 14 磅", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(8, 0) + markSz(28), text: "方块" }]));
    expect(doc.paragraphs()[0].markSizePt).toBe(14);
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-family:&quot;Calibri&quot;;font-size:14pt">${MARK}</span>`);
  });

  it("标记没写字号：往上一跳取样式链，不取 docDefaults", () => {
    const doc = open(makeNumberingDocx([{ pPr: `<w:pStyle w:val="MarkSized"/>${numPr(8, 0)}`, text: "方块" }]));
    expect(doc.paragraphs()[0].markSizePt).toBe(16);
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-family:&quot;Calibri&quot;;font-size:16pt">${MARK}</span>`);
  });

  it("两个都没写：落到 docDefaults 那 11 磅——页面样式表的 12 磅不是文件说的数", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(8, 0), text: "方块" }]));
    expect(doc.paragraphs()[0].markSizePt).toBe(11);
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-family:&quot;Calibri&quot;;font-size:11pt">${MARK}</span>`);
  });

  it("级上写的赢段落标记写的：numId 9 的 8 磅盖掉标记那 11 磅", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(9, 0) + markSz(22), text: "方块" }]));
    expect(doc.levelOf(9, 0)?.sizePt).toBe(8);
    expect(doc.paragraphs()[0].markSizePt).toBe(11);
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-family:Wingdings;font-size:8pt">${MARK}</span>`);
  });

  it("标记的磅数只穿那颗符号：这一段自己的字照它 run 上解析出的大小", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(8, 0) + markSz(28), text: "方块" }]));
    const para = drawnPara(doc, "方块");
    expect((para.match(/font-size:14pt/g) ?? []).length).toBe(1);
    expect(/<span class="bullet"[^>]*font-size:14pt/.test(para)).toBe(true);
    expect(para).toContain("font-size:11pt");
  });

  it("画它不写它：段落的 w:pPr 按原件带回，numbering.xml 逐字节照旧", () => {
    const src = makeNumberingDocx([{ pPr: numPr(8, 0) + markSz(28), text: "方块" }]);
    const doc = open(src);
    const regenerated = saved(doc);
    expect(pPrOf(regenerated, "方块")).toContain('<w:sz w:val="28"/><w:szCs w:val="28"/>');
    expect((regenerated.match(/<w:sz\b/g) ?? []).length).toBe((part(src).match(/<w:sz\b/g) ?? []).length);
    expect(numberingOf(writeDocx(doc))).toBe(numberingOf(src));
    expect(paraXml(regenerated, "方块")).not.toContain("14pt");
  });
});

describe("那颗符号的脸回退：级没点脸就用段落标记自己解析出的那两张", () => {
  /** The corpus's own spelling of a mark that names its faces: both attributes plus the `w:hint` WPS adds. */
  const markFaces = (attrs: string): string => `<w:rPr><w:rFonts ${attrs}/></w:rPr>`;
  /** The East Asian half of the stack the mark wears: the invented family, quoted as the run's stack spells it. */
  const EA_STACK = `font-family:&quot;solo-ea-黑体&quot;`;

  it("w:pPr/w:rPr/w:rFonts 进了模型：一、 穿标记点的那两张，东亚那张在前", () => {
    // numId 11 是语料里那 282 段的形状：级只写 `w:hint="eastAsia"`，一张脸都没点。
    const doc = open(makeNumberingDocx([{ pPr: numPr(11, 0) + markFaces(`w:ascii="Arial" w:eastAsia="黑体" w:hint="eastAsia"`), text: "方块" }]));
    const p = doc.paragraphs()[0];
    expect(doc.levelOf(11, 0)?.font).toBeUndefined();
    expect([p.markFont, p.markFontEastAsia]).toEqual(["Arial", "黑体"]);
    expect(markOf(doc, "方块")).toBe(
      `<span class="bullet" style="font-family:&quot;solo-ea-黑体&quot;,&quot;Arial&quot;;font-size:11pt">一、</span>`,
    );
    expect(doc.eastAsiaFaces()).toContain("黑体");
  });

  it("两张分别取：标记只点 eastAsia，拉丁那一路仍从链上借——不是把借来的擦掉", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(11, 0) + markFaces(`w:eastAsia="黑体"`), text: "方块" }]));
    const p = doc.paragraphs()[0];
    // docDefaults 点了 Calibri，标记没点拉丁那一路，所以两样各归各（§17.7.2 每条属性各自止步）。
    expect([p.markFont, p.markFontEastAsia]).toEqual(["Calibri", "黑体"]);
    expect(markOf(doc, "方块")).toContain(`font-family:&quot;solo-ea-黑体&quot;,&quot;Calibri&quot;;`);
  });

  it("标记一个字都没点：往上一跳取样式链那两张", () => {
    const doc = open(makeNumberingDocx([{ pPr: `<w:pStyle w:val="MarkFaced"/>${numPr(11, 0)}`, text: "方块" }]));
    const p = doc.paragraphs()[0];
    expect([p.markFont, p.markFontEastAsia]).toEqual(["Arial", "宋体"]);
    expect(markOf(doc, "方块")).toContain(`font-family:&quot;solo-ea-宋体&quot;,&quot;Arial&quot;;`);
  });

  it(`空值不算点了脸：w:ascii="" 让位给链上的 Calibri`, () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(11, 0) + markFaces(`w:ascii="" w:eastAsia="黑体"`), text: "方块" }]));
    expect([doc.paragraphs()[0].markFont, doc.paragraphs()[0].markFontEastAsia]).toEqual(["Calibri", "黑体"]);
  });

  it("级上点了脸的赢：numId 9 的 Wingdings 盖过标记那两张", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(9, 0) + markFaces(`w:ascii="Arial" w:eastAsia="黑体"`), text: "方块" }]));
    expect(doc.levelOf(9, 0)?.font).toBe("Wingdings");
    expect(markOf(doc, "方块")).toContain("font-family:Wingdings;");
    expect(markOf(doc, "方块")).not.toContain(EA_STACK);
  });

  it("三层都没点：不写 font-family——页面不能造文件里没有的脸", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(11, 0), text: "方块" }], NUMBERING, numberingStyles("")));
    const p = doc.paragraphs()[0];
    expect([p.markFont, p.markFontEastAsia]).toEqual([undefined, undefined]);
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-size:11pt">一、</span>`);
  });

  it("脸只穿那颗符号：这一段自己的字照它 run 解析出的 Calibri，不跟着换黑体", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(11, 0) + markFaces(`w:ascii="Arial" w:eastAsia="黑体"`), text: "方块" }]));
    const para = drawnPara(doc, "方块");
    expect((para.match(/黑体/g) ?? []).length).toBe(1);
    expect(/<span class="bullet"[^>]*黑体/.test(para)).toBe(true);
    expect(para).toContain(`<span style="font-size:11pt;font-family:&quot;Calibri&quot;;">方块</span>`);
  });

  it("标记点的脸正文没用过也要有规则：@font-face 与那颗符号的栈都指向它", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(11, 0) + markFaces(`w:eastAsia="黑体"`), text: "方块" }]));
    const html = renderWord(doc);
    expect((html.match(/solo-ea-黑体/g) ?? []).length).toBe(2); // 规则里的族名 + 那颗符号的栈
    expect(html).toContain(`@font-face{font-family:"solo-ea-黑体";src:local("黑体");`);
  });

  it("画它不写它：w:pPr 里的 w:rFonts 按原件带回，numbering.xml 逐字节照旧", () => {
    const src = makeNumberingDocx([{ pPr: numPr(11, 0) + markFaces(`w:ascii="Arial" w:eastAsia="黑体"`), text: "方块" }]);
    const doc = open(src);
    const regenerated = saved(doc);
    expect(pPrOf(regenerated, "方块")).toContain(`<w:rFonts w:ascii="Arial" w:eastAsia="黑体"/>`);
    expect(regenerated).not.toContain("solo-ea");
    expect(numberingOf(writeDocx(doc))).toBe(numberingOf(src));
    // 没写 rFonts 的段落，重拼也不给它添一个。
    const bare = open(makeNumberingDocx([{ pPr: numPr(11, 0), text: "干净" }]));
    expect(pPrOf(saved(bare), "干净")).not.toContain("<w:rFonts");
  });
});

describe("那颗符号的颜色回退：级没点色就用段落标记自己解析出的那张", () => {
  /** The corpus's own spelling of a mark that states a colour: `w:val` alone, six hex digits (74/74). */
  const markColor = (val: string): string => `<w:rPr><w:color w:val="${val}"/></w:rPr>`;
  /** A declaration off the *mark's* own style — `propOf` reads the paragraph wrapper, which is not this span. */
  const markProp = (doc: WordDocument, needle: string, prop: string): string =>
    new RegExp(`(?:^|;)${prop}:([^;]*)`).exec(/ style="([^"]*)"/.exec(markOf(doc, needle))?.[1] ?? "")?.[1] ?? "";

  it("w:pPr/w:rPr/w:color 进了模型：一、 穿标记点的那张，三条声明按字体、字号、颜色排", () => {
    const doc = open(
      makeNumberingDocx([
        { pPr: numPr(11, 0) + `<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="黑体"/><w:color w:val="C00000"/></w:rPr>`, text: "方块" },
      ]),
    );
    expect(doc.paragraphs()[0].markColor).toBe("#C00000");
    expect(markOf(doc, "方块")).toBe(
      `<span class="bullet" style="font-family:&quot;solo-ea-黑体&quot;,&quot;Arial&quot;;font-size:11pt;color:#C00000">一、</span>`,
    );
  });

  it("标记没点颜色：往上一跳取样式链那一张，字号仍归 docDefaults", () => {
    // numId 11 的级既没点脸也没点色，链上只有 MarkColored 那一张颜色可取。
    const doc = open(makeNumberingDocx([{ pPr: `<w:pStyle w:val="MarkColored"/>${numPr(11, 0)}`, text: "方块" }]));
    expect(doc.paragraphs()[0].markColor).toBe("#7030A0");
    expect(markOf(doc, "方块")).toContain(`font-size:11pt;color:#7030A0`);
  });

  it("标记点的赢样式链的：同一条属性也各自止步（§17.7.2）", () => {
    const doc = open(
      makeNumberingDocx([{ pPr: `<w:pStyle w:val="MarkColored"/>${numPr(11, 0)}${markColor("C00000")}`, text: "方块" }]),
    );
    expect(doc.paragraphs()[0].markColor).toBe("#C00000");
    expect(markOf(doc, "方块")).toContain("color:#C00000");
    expect(markOf(doc, "方块")).not.toContain("#7030A0");
  });

  it(`w:val="auto" 不算点了颜色：让位给链上那一张`, () => {
    const doc = open(
      makeNumberingDocx([
        { pPr: `<w:pStyle w:val="MarkColored"/>${numPr(11, 0)}${markColor("auto")}`, text: "借链" },
        { pPr: numPr(11, 0) + markColor("auto"), text: "没有链" },
      ]),
    );
    expect(doc.paragraphs()[0].markColor).toBe("#7030A0");
    // 这一段的 docDefaults 也没点色，所以「auto」一路读到底还是没有——不写 color 比造一个文件没说的色强。
    expect(doc.paragraphs()[1].markColor).toBeUndefined();
    expect(markOf(doc, "没有链")).not.toContain("color:");
  });

  it("不是颜色的值不算点了颜色：形状不对就退回链", () => {
    const doc = open(
      makeNumberingDocx([
        { pPr: numPr(11, 0) + markColor("GGGGGG"), text: "坏值" },
        { pPr: `<w:pStyle w:val="MarkColored"/>${numPr(11, 0)}${markColor("GGGGGG")}`, text: "坏值借链" },
      ]),
    );
    expect(doc.paragraphs()[0].markColor).toBeUndefined();
    expect(markOf(doc, "坏值")).not.toContain("color:");
    expect([doc.paragraphs()[1].markColor, markProp(doc, "坏值借链", "color")]).toEqual(["#7030A0", "#7030A0"]);
  });

  it("三层都没点：不写 color——页面不能造文件里没有的颜色", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(11, 0), text: "方块" }]));
    expect(doc.paragraphs()[0].markColor).toBeUndefined();
    expect(markOf(doc, "方块")).toBe(`<span class="bullet" style="font-family:&quot;Calibri&quot;;font-size:11pt">一、</span>`);
  });

  it("颜色只穿那颗符号：这一段自己的字照它 run 解析出的，不跟着变红", () => {
    const doc = open(makeNumberingDocx([{ pPr: numPr(11, 0) + markColor("C00000"), text: "方块" }]));
    const para = drawnPara(doc, "方块");
    // 语料里 74/74 的标记那张色与正文解析出的其实是**同一张**（文件把墨写在 run 上也写在标记上）；这一夹具故意只
    // 让标记点色，为的是量「那张色不外溢」这条判据本身——正文没色时它必须一点色都不沾。
    expect((para.match(/C00000/g) ?? []).length).toBe(1);
    expect(/<span class="bullet"[^>]*C00000/.test(para)).toBe(true);
    expect(para).toContain(`<span style="font-size:11pt;font-family:&quot;Calibri&quot;;">方块</span>`);
  });

  it("画它不写它：w:pPr 里的 w:color 按原件带回，numbering.xml 逐字节照旧", () => {
    const src = makeNumberingDocx([{ pPr: numPr(11, 0) + markColor("C00000"), text: "方块" }]);
    const doc = open(src);
    const regenerated = saved(doc);
    expect(pPrOf(regenerated, "方块")).toContain(`<w:color w:val="C00000"/>`);
    expect(regenerated).not.toContain("#C00000");
    expect(numberingOf(writeDocx(doc))).toBe(numberingOf(src));
    // 没写 w:color 的段落，重拼也不给它添一个。
    const bare = open(makeNumberingDocx([{ pPr: numPr(11, 0), text: "干净" }]));
    expect(pPrOf(saved(bare), "干净")).not.toContain("<w:color");
  });
});
