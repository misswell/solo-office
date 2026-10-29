import { describe, expect, it } from "vitest";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { clearRunStyle, parseSizePt, WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { childSpans, elementContent, OpcPackage, text } from "../src/core/opc.js";
import { makeClearFormatDocx, makeHighlightDocx, makeOnOffDocx, makeRunPropsDocx, makeTextFillDocx, makeUnderlineDocx } from "./helpers/make-run.js";
import { renderInline } from "../src/core/render.js";

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
 * `textFill` 2214, `shd` 1454, `u` 1256, `i` 971, `iCs` 960, `caps` 274, `strike` 38. Those counts are 作废语料
 * readings (#106) and have not been re-run on the authorized 46 件, so they are 待重证; the shapes they named are
 * the ones the fixtures below carry.
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
      bold: true, italic: false, underline: "single", strike: true, sizePt: 12, color: "#808080",
      highlight: "yellow", font: "Times New Roman", fontEastAsia: "宋体",
      // `w:vertAlign w:val="baseline"` is in this run's set, and the run saying it sits on the line is a thing it
      // stated: the field holds the name, which is what stops the writer from answering the file with a silence.
      vertAlign: "baseline",
      // `w:spacing w:val="20"` is the run's own 字符间距 sentence (#174 gave it a field); the `w:w` percentage
      // beside it has none and travels as bytes.
      charSpacingTwips: 20,
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

  it("takes a typed size to the nearest half point before it reaches w:sz", () => {
    // `w:sz` counts half-points (§17.18.60), so 13.7 in the box can only be said as 13.5 on the page.
    expect(parseSizePt("13.7")).toBe(13.5);
    expect(parseSizePt("10.5")).toBe(10.5);
    expect(parseSizePt(" 12 ")).toBe(12);
    expect(parseSizePt("13.7 ")).toBe(13.5);
    // Not a number, not half a point, or past the box's ceiling: no size at all, so the caller writes nothing.
    expect(parseSizePt("abc")).toBeUndefined();
    expect(parseSizePt("")).toBeUndefined();
    expect(parseSizePt("0")).toBeUndefined();
    expect(parseSizePt("0.2")).toBeUndefined();
    expect(parseSizePt("-1")).toBeUndefined();
    expect(parseSizePt("1639")).toBeUndefined();
    expect(parseSizePt("1e3")).toBeUndefined();

    const doc = open(makeRunPropsDocx());
    const typed = parseSizePt("13.7")!;
    const out = rPrOf(retyped(doc, "字号成对", { ...runAt(doc, "字号成对"), sizePt: typed }), "字号成对");
    expect(out).toContain('<w:sz w:val="27"/><w:szCs w:val="27"/>');
  });

  it("clears the eleven modeled run fields and keeps the slots the model has no field for", () => {
    const doc = open(makeClearFormatDocx());
    const cleared = clearRunStyle(runAt(doc, "整组都说的一枚"));
    const rpr = rPrOf(retyped(doc, "整组都说的一枚", cleared), "整组都说的一枚");
    for (const tag of ["b", "i", "u", "strike", "color", "sz", "szCs", "highlight", "vertAlign", "spacing"]) {
      expect(rpr).not.toContain(`<w:${tag}`);
    }
    // 模型没有字段的两样按原件带回：`w:cs` 那一槽的脸留在同一枚 `w:rFonts` 上，`w:lang` 原样。
    expect(rpr).toContain('<w:rFonts w:cs="Arial"/>');
    expect(rpr).toContain('<w:lang w:val="en-US"/>');

    // 两张字面脸都撤掉、元素里只剩它自己的名字时，这枚 `w:rFonts` 整个退场（不落一枚空的）。
    const facesOnly = open(makeClearFormatDocx());
    const run = runAt(facesOnly, "整组都说的一枚");
    run.rPrSlots = (run.rPrSlots ?? []).map((s) => (s.tag === "rFonts" ? { tag: s.tag, xml: '<w:rFonts w:ascii="Arial" w:eastAsia="宋体"/>' } : s));
    expect(rPrOf(retyped(facesOnly, "整组都说的一枚", clearRunStyle(run)), "整组都说的一枚")).not.toContain("w:rFonts");

    // 只有槽的一枚与什么都没说的一枚：模型没有字段可清，盘上的字节就该一字不动。
    const fresh = part(makeClearFormatDocx());
    const same = open(makeClearFormatDocx());
    const xml = retyped(same, "只有槽的一枚", clearRunStyle(runAt(same, "只有槽的一枚")));
    expect(rPrOf(xml, "只有槽的一枚")).toBe(rPrOf(fresh, "只有槽的一枚"));
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
    expect(runAt(doc, "链接里的字").underline).toBe("single");
    // The `w:Hyperlink` style draws it and the file never wrote a `w:u`, so an untouched link comes back as it sat.
    expect(rPrOf(elsewhere(makeRunPropsDocx()), "链接里的字")).toBe('<w:rPr><w:rStyle w:val="Hyperlink"/><w:kern w:val="28"/></w:rPr>');

    // Moved, it is spelled into the carried set at the rank `CT_RPr` gives it: after `w:szCs`, before `w:shd`.
    const on = open(makeRunPropsDocx());
    expect(rPrOf(retyped(on, "字号成对", { ...runAt(on, "字号成对"), underline: "single" }), "字号成对")).toBe(
      '<w:rPr><w:rFonts w:eastAsia="仿宋" w:cs="仿宋"/><w:sz w:val="21"/><w:szCs w:val="24"/><w:u w:val="single"/></w:rPr>',
    );
    // Off, the cancellation is a value the run writes rather than an element it deletes, exactly as `w:b` above
    // becomes `<w:b w:val="0"/>`. That is the shape WPS itself uses: 授权 46 份的 `word/document.xml` 里
    // `<w:u w:val="none"/>` 有 30 枚（同一趟逐枚：`single` 1,565、`thick` 20、裸写 0），段落标记那一侧另有 3 枚
    // （`.scratch/underline-audience.py`）。
    const off = open(makeRunPropsDocx());
    expect(rPrOf(retyped(off, "整段自述", { ...runAt(off, "整段自述"), underline: "none" }), "整段自述")).toContain(
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
 * `w:u` counted on the authorized cohort (46 份 = `.scratch/wps-corpus` + WPS 自带模板，只数 `.docx/.dotx/.dotm/.docm`；
 * `.scratch/underline-audience.py` 逐枚数 `w:rPr/w:u`，另开一趟只数它的属性名）。1,615 枚说过线的 run 住在 22 件里：
 * `single` 1,565 / `none` 30 / `thick` 20，枚枚写了 `w:val`，枚枚只写 `w:val`（没有一枚带 `w:color` 或 `themeColor`）。
 * 段落标记那一侧 316 / 3 / 16，`word/styles.xml` 22 枚全在字符样式自己的 `w:rPr` 里、枚枚 `single`。
 * 旧口径那组（1586 枚 / 488 成字段 / 254 `none` / 768 成字节 / 35 枚样式）随 86 份语料一起作废，这里的数是重量的。
 *
 * 这一族因此按形状名建模，不再收成「有没有线」：`thick` 那 20 枚坐在 2 份件里，收成 `single` 就是把文件说的粗线
 * 画成细线、拼回再写成细线。`LendU` 那一格照 §17.7.2 的借线形状摆（授权口径里 `none` 坐在借得到线的链上有几枚
 * 尚未再数，所以它是形状的夹具而不是量到的一枚）。至于 WPS 那 8 档线形下拉——`CT_TableTools.kuip:115-122` 只报
 * `icon="line_style_{1,7,25,2,4,5,6,18}_display_kd"`，图标资源不在这个包里，档名与线形的对应仍是未取证，本轮不摆
 * 那一圈；`UnderLineColor`（同文件 :124-127）在授权口径里 0 受众（1,615 枚 `w:u` 无一枚说色），同样不摆。
 */
describe("docx underline read as a statement", () => {
  it("reads a stated no-underline as a value the run gave, not a missing element", () => {
    const run = runAt(open(makeUnderlineDocx()), "说明没有");
    expect(run.underline).toBe("none");
    expect(run.asRead?.underline).toBe("none");
  });

  it("stops the ladder per property at the rung that stated one", () => {
    const doc = open(makeUnderlineDocx());
    // §17.7.2 ranks the run's own `w:rPr` above the character style it names, so this run's `none` beats the
    // line the style lends. The other three are the controls: a style that is not contradicted lends a line,
    // and a run that names no style still says which way it goes for itself.
    expect(runAt(doc, "这句话盖掉借来的线").underline).toBe("none");
    expect(runAt(doc, "链上借来的线").underline).toBe("single");
    expect(runAt(doc, "自己点了线").underline).toBe("single");
    expect(runAt(doc, "粗线原样回去").underline).toBe("thick");
  });

  it("hands back a stated none and a stated thick exactly as they sat", () => {
    const before = part(makeUnderlineDocx());
    const after = save(open(makeUnderlineDocx()));
    expect(rPrOf(after, "说明没有")).toBe(rPrOf(before, "说明没有"));
    expect(rPrOf(after, "说明没有")).toBe('<w:rPr><w:u w:val="none"/></w:rPr>');
    // 形状名现在就在模型里：粗线不再靠「这枚 run 说不出来所以随行成字节」保住，而是说出来也拼回去。
    expect(rPrOf(after, "粗线原样回去")).toBe('<w:rPr><w:u w:val="thick"/></w:rPr>');
  });

  it("writes the cancellation when the page turns off a line the style lent", () => {
    const doc = open(makeUnderlineDocx());
    const out = rPrOf(
      retyped(doc, "链上借来的线", { ...runAt(doc, "链上借来的线"), underline: "none" }),
      "链上借来的线",
    );
    expect(out).toBe('<w:rPr><w:rStyle w:val="LendU"/><w:u w:val="none"/></w:rPr>');
  });

  it("writes a line back when the page turns a cancelled run on", () => {
    const doc = open(makeUnderlineDocx());
    const out = rPrOf(retyped(doc, "说明没有", { ...runAt(doc, "说明没有"), underline: "single" }), "说明没有");
    expect(out).toBe('<w:rPr><w:u w:val="single"/></w:rPr>');
  });

  /**
   * The page is where a name can be lost twice over: once by drawing `thick` at a single line's thickness, once by
   * giving the editor nothing to read back. So the two faces are pinned separately — the CSS says how it draws,
   * `data-u` says what the file named, and a name CSS already carries (`single`) is not stamped again.
   */
  it("draws the thickness the name says and stamps only what CSS cannot carry back", () => {
    const doc = open(makeUnderlineDocx());
    expect(renderInline(runAt(doc, "粗线原样回去"))).toBe(
      '<span style="text-decoration:underline;text-decoration-thickness:0.1875em;" data-u="thick">粗线原样回去</span>',
    );
    // `single` draws exactly as it did while the model held a boolean: no extra bytes for the 1,565 runs that say it.
    expect(renderInline(runAt(doc, "自己点了线"))).toBe('<span style="text-decoration:underline;">自己点了线</span>');
    expect(renderInline(runAt(doc, "链上借来的线"))).toBe('<span style="text-decoration:underline;">链上借来的线</span>');
    // A stated `none` draws no line — which is the same picture as a run that never spoke — so the stamp is the only
    // thing that tells the two sentences apart on the way back.
    expect(renderInline(runAt(doc, "说明没有"))).toBe('<span data-u="none">说明没有</span>');
  });
});

/**
 * `w:strike` and `w:vertAlign`, counted three ways and reconciled (`.scratch/onoff-where.ts` cuts the bodies by
 * whose property set each tag sits in, `.scratch/onoff-family.ts` asks the production model what it resolves, and
 * `.scratch/onoff-roundtrip.ts` runs open→save over all 86 files). The bodies state a `w:strike` 56 times: 36 on a
 * run's own set with no `w:val` (on, by the element's default), 18 on a paragraph mark, and **2 stating
 * `<w:strike w:val="0"/>`** in 2 files — both on a run carrying text, and both the reader used to resolve to
 * `strike = true`, a line drawn through text the file said to leave alone. After this round the round trip returns
 * 56 as 56 with nothing moved, and those 2 resolve `false`. `w:vertAlign` appears 3860 times and **every one of
 * them says `baseline`** — 2701 on a run's set, 1159 on a mark — of which this reader models 717 runs, 354 naming a
 * character style, and 0 of those sitting on a chain that lends a superscript: representation kept, not pixels
 * rescued, the same verdict `w:u` reached. `w:dstrike` states `0` twice and is not a field this model has, so it
 * travels as bytes. `LendStrike`, `LendSup` and the two-line run are the shapes the corpus points at but does not
 * carry — built to the patterns above so each sentence has the rung it must beat and the drawing has a case to
 * compose. Nothing here claims a pixel the corpus has not shown.
 */
describe("docx on/off family read as a statement", () => {
  it("reads a stated no-strike as a value the run gave, under either spelling", () => {
    const doc = open(makeOnOffDocx());
    expect(runAt(doc, "说了没划").strike).toBe(false);
    expect(runAt(doc, "说了没划").asRead?.strike).toBe(false);
    // §17.3.2.19's `ST_OnOff` takes the word as readily as the digit, the way `w:i` already did.
    expect(runAt(doc, "盖掉借来的删除线").strike).toBe(false);
  });

  it("stops the ladder per property at the rung that stated one", () => {
    const doc = open(makeOnOffDocx());
    // The run's own `w:rPr` outranks the character style it names (§17.7.2), so the `false` beats the line the
    // style lends, and the style that goes uncontradicted still lends it.
    expect(runAt(doc, "链上借来的删除线").strike).toBe(true);
    expect(runAt(doc, "两条线一起").strike).toBe(true);
    // The same for the alignment: a stated `baseline` beats the superscript the style would lend, and it is a
    // value the model holds rather than an absence — the page has to know the run spoke.
    expect(runAt(doc, "回到基线").vertAlign).toBe("baseline");
  });

  it("hands back a stated cancellation exactly as it sat", () => {
    const before = part(makeOnOffDocx());
    const after = save(open(makeOnOffDocx()));
    expect(rPrOf(after, "说了没划")).toBe('<w:rPr><w:strike w:val="0"/></w:rPr>');
    expect(rPrOf(after, "盖掉借来的删除线")).toBe('<w:rPr><w:rStyle w:val="LendStrike"/><w:strike w:val="false"/></w:rPr>');
    expect(rPrOf(after, "回到基线")).toBe(rPrOf(before, "回到基线"));
  });

  it("writes the cancellation when the page turns off a line the style lent", () => {
    const doc = open(makeOnOffDocx());
    const out = rPrOf(
      retyped(doc, "链上借来的删除线", { ...runAt(doc, "链上借来的删除线"), strike: false }),
      "链上借来的删除线",
    );
    // Spelled, not deleted: dropping the element would answer the file's sentence with a silence no file wrote.
    expect(out).toBe('<w:rPr><w:rStyle w:val="LendStrike"/><w:strike w:val="0"/></w:rPr>');
  });

  it("keeps a superscript run's stated size outside the scale, so the box can still read it", () => {
    const html = renderInline(runAt(open(makeClearFormatDocx()), "整组都说"));
    // 22 磅（`w:sz 44`）写在**外面**那一枚上，缩一圈（`.75em`）套在里面。同一枚上后写的那句 `.75em` 会把
    // 文件说的磅数盖掉：纸上画成 12px，字号框读回空串（`.75em` 不以 pt 结尾，`styleOf` 认不出）。
    expect(html).toContain("font-size:22pt;");
    expect(html).toContain('<span style="vertical-align:super;font-size:.75em;">整组都说的一枚</span>');
    expect(html.indexOf("font-size:22pt")).toBeLessThan(html.indexOf("font-size:.75em"));
  });

  it("draws two lines on one run instead of letting the later declaration replace the earlier", () => {
    const both = runAt(open(makeOnOffDocx()), "两条线一起");
    expect(renderInline(both)).toBe('<span style="text-decoration:line-through underline;">两条线一起</span>');
    // A stated baseline is the run saying it sits on the line: it draws nothing, and it is not a superscript.
    const base = renderInline(runAt(open(makeOnOffDocx()), "回到基线"));
    expect(base).not.toContain("vertical-align");
    expect(base).not.toContain("font-size");
  });

  it("carries both statements through every round trip, not just the first", () => {
    const once = writeDocx(open(makeOnOffDocx()));
    const twice = writeDocx(open(once));
    expect(part(twice)).toBe(part(once));
    const back = open(twice);
    expect(runAt(back, "说了没划").strike).toBe(false);
    expect(runAt(back, "回到基线").vertAlign).toBe("baseline");
  });
});

/** Every `w:highlight` tag, as the file wrote it. */
const hlTags = (xml: string): string[] => [...xml.matchAll(/<w:highlight(?:\s[^>]*)?\/>/g)].map((m) => m[0]);

/**
 * `w:highlight` is the third member of the "the value says which family member" line-up — §17.17.64's
 * `ST_HighlightColor`, seventeen names of which one, `none`, means the same as `0`/`false`/`baseline` does for the
 * on/off and alignment families: the run spoke, and what it spoke was no colour.
 *
 * Counted three ways and reconciled (`.scratch/highlight-counted.ts` walks the tag stack over 298 parts of 86 WPS
 * files, then asks the production model what it resolves): the bodies state `w:highlight` **4605 times** and only
 * two members ever appear — **`none` 4135** and `yellow` 470. Cut by owner: 2417 `none` plus 249 `yellow` on a
 * run's own set, 1718 plus 221 on a paragraph mark. The run layer answers "who can reach it": of those 2666,
 * **2369 are on runs this reader models** (2361 carrying text — 1474 inside table cells, 887 in body paragraphs —
 * and 8 as marks), and the other **48 sit inside runs kept as the bytes they came in with**, which is exactly the
 * 2417. Every one of them states `w:val`; no run leaves it off, and nothing in the corpus uses a name outside the
 * seventeen.
 *
 * Who could lend a colour: **nobody**. `word/styles.xml` states `w:highlight` 0 times, `w:rPrDefault` 0, and the
 * numbering definitions and the header/footer/footnote/endnote parts 0 each, so in these files the 4135 `none`
 * sentences cancel a silence rather than a colour, and the rescue this round makes visible — a typed run no longer
 * resurrecting a lent `yellow` — has 0 audience today. `LendHl` and the `blue` run are shapes the corpus points at
 * without carrying, built to the patterns above. What does have audience is the file's own words: 2417 runs say
 * `w:val="none"`, and reading back off the CSS lost that sentence (`background:none` hands its longhand to
 * Chromium as `initial`, and only the seven names the renderer spells come back as a lookable `rgb(...)` —
 * `.scratch/hl-pending.ts`).
 */
describe("docx highlight read as three states", () => {
  it("tells the cancellation from the silence, under a name the page could not guess", () => {
    const doc = open(makeHighlightDocx());
    // §17.17.64's `none` is a value the run gave, so the model holds it as one rather than as an absence.
    expect(runAt(doc, "说了没高亮").highlight).toBe("none");
    expect(runAt(doc, "说了没高亮").asRead?.highlight).toBe("none");
    expect(runAt(doc, "什么都没点").highlight).toBeUndefined();
    // A name outside the seven the ribbon offers reads as itself; it is not a colour the page can look up.
    expect(runAt(doc, "页面叫不出名字的高亮").highlight).toBe("blue");
  });

  it("stops the ladder per property at the rung that stated one", () => {
    const doc = open(makeHighlightDocx());
    // The run's own `w:rPr` outranks the character style it names (§17.7.2), so `none` beats the colour the style
    // lends — and a style nobody contradicts still lends it.
    expect(runAt(doc, "链上借来的高亮").highlight).toBe("yellow");
    expect(runAt(doc, "盖掉借来的高亮").highlight).toBe("none");
    expect(runAt(doc, "黄底").highlight).toBe("yellow");
    // The borrowed colour is not the run's own sentence: it shows on screen but travels in no `w:rPr` slot.
    expect(runAt(doc, "链上借来的高亮").rPrSlots?.some((s) => s.tag === "highlight")).toBe(false);
    expect(runAt(doc, "盖掉借来的高亮").rPrSlots?.some((s) => s.tag === "highlight")).toBe(true);
  });

  it("hands back a stated cancellation, and a name it cannot draw, exactly as they sat", () => {
    const before = part(makeHighlightDocx());
    const after = save(open(makeHighlightDocx()));
    expect(rPrOf(after, "说了没高亮")).toBe('<w:rPr><w:highlight w:val="none"/></w:rPr>');
    expect(rPrOf(after, "盖掉借来的高亮")).toBe('<w:rPr><w:rStyle w:val="LendHl"/><w:highlight w:val="none"/></w:rPr>');
    expect(rPrOf(after, "页面叫不出名字的高亮")).toBe(rPrOf(before, "页面叫不出名字的高亮"));
    // The run that stated nothing keeps stating nothing: no `w:rPr` is invented to hold a colour that isn't there.
    expect(rPrOf(after, "什么都没点")).toBe("");
    expect(hlTags(after)).toEqual(hlTags(before));
  });

  it("writes the cancellation when the page turns off a colour the style lent", () => {
    const doc = open(makeHighlightDocx());
    const out = rPrOf(
      retyped(doc, "链上借来的高亮", { ...runAt(doc, "链上借来的高亮"), highlight: "none" }),
      "链上借来的高亮",
    );
    // Spelled, not deleted: the file's way of saying no is the enum name, and a silence is an answer no file wrote.
    expect(out).toBe('<w:rPr><w:rStyle w:val="LendHl"/><w:highlight w:val="none"/></w:rPr>');
  });

  it("keeps the name the document chose when the page moves a run from one colour to another", () => {
    const doc = open(makeHighlightDocx());
    const moved = rPrOf(retyped(doc, "黄底", { ...runAt(doc, "黄底"), highlight: "green" }), "黄底");
    expect(moved).toBe('<w:rPr><w:highlight w:val="green"/></w:rPr>');
    // `blue` is a §17.17.64 member the renderer has never met: it must go back out as the name, not as a colour.
    const kept = rPrOf(retyped(doc, "页面叫不出名字的高亮", { ...runAt(doc, "页面叫不出名字的高亮"), text: "页面叫不出名字的高亮字" }), "页面叫不出名字的高亮");
    expect(kept).toBe('<w:rPr><w:highlight w:val="blue"/></w:rPr>');
  });

  it("draws the cancellation as nothing, and stamps the name the run stated on the span it draws", () => {
    const doc = open(makeHighlightDocx());
    expect(renderInline(runAt(doc, "说了没高亮"))).toBe('<span style="background:transparent;" data-hl="none">说了没高亮</span>');
    expect(renderInline(runAt(doc, "黄底"))).toBe('<span style="background:#ffff00;" data-hl="yellow">黄底</span>');
    expect(renderInline(runAt(doc, "页面叫不出名字的高亮"))).toBe('<span style="background:blue;" data-hl="blue">页面叫不出名字的高亮</span>');
    // The run that said nothing gets no stamp: an empty `data-hl` would be a sentence the file never wrote.
    expect(renderInline(runAt(doc, "什么都没点"))).toBe("<span>什么都没点</span>");
    // A borrowed colour draws, but the run's own words are what the stamp carries.
    expect(renderInline(runAt(doc, "链上借来的高亮"))).toBe('<span style="background:#ffff00;" data-hl="yellow">链上借来的高亮</span>');
    expect(renderInline(runAt(doc, "盖掉借来的高亮"))).toBe('<span style="background:transparent;" data-hl="none">盖掉借来的高亮</span>');
  });

  it("carries every state through every round trip, and through an edit elsewhere in the body", () => {
    const once = writeDocx(open(makeHighlightDocx()));
    const twice = writeDocx(open(once));
    expect(part(twice)).toBe(part(once));
    const back = open(twice);
    expect(runAt(back, "说了没高亮").highlight).toBe("none");
    expect(runAt(back, "页面叫不出名字的高亮").highlight).toBe("blue");
    expect(runAt(back, "盖掉借来的高亮").highlight).toBe("none");
    expect(hlTags(elsewhere(makeHighlightDocx()))).toEqual(hlTags(part(makeHighlightDocx())));
  });
});

/** How many `w14:textFill` elements the part states, wherever they sit. */
const fillCount = (xml: string): number => (xml.match(/<w14:textFill\b/g) ?? []).length;

describe("docx 字色与那句扩展色（w14:textFill）", () => {
  it("改字色时那句扩展色跟着让位，其余的孩子原位不动", () => {
    const doc = open(makeTextFillDocx());
    const out = rPrOf(retyped(doc, "镜像的一对", { ...runAt(doc, "镜像的一对"), color: "#0000FF" }), "镜像的一对");
    expect(out).toContain('<w:color w:val="0000FF"/>');
    // 镜像跟着字色走：留下它，这一句就成了「一个 run 报两个颜色」，而谁说了算不由我们定。
    expect(out).not.toContain("textFill");
    expect(out).toContain('<w:kern w:val="28"/>');
    expect(out).toContain('<w:sz w:val="24"/><w:szCs w:val="24"/>');
    expect(out).toContain('w:hint="eastAsia"');
  });

  it("没动字色时不必重写那一句，连它报的主题名也原样留着", () => {
    const doc = open(makeTextFillDocx());
    const out = rPrOf(retyped(doc, "镜像的一对", { ...runAt(doc, "镜像的一对"), bold: true }), "镜像的一对");
    expect(out).toContain('<w14:textFill><w14:solidFill><w14:schemeClr val="text1"/></w14:solidFill></w14:textFill>');
    expect(out).toContain('w:themeColor="text1"');
    expect(out).toContain("<w:b/>");
  });

  it("把字色改回它读进来的那一个，那一句一个字都不必重写", () => {
    const before = rPrOf(part(makeTextFillDocx()), "镜像的一对");
    const doc = open(makeTextFillDocx());
    const run = runAt(doc, "镜像的一对");
    paraAt(doc, "镜像的一对").inlines[0] = { ...run, color: "#0000FF" };
    expect(fillCount(save(doc))).toBe(fillCount(part(makeTextFillDocx())) - 1);
    paraAt(doc, "镜像的一对").inlines[0] = { ...run };
    expect(rPrOf(save(doc), "镜像的一对")).toBe(before);
  });

  it("给一个只说过 w:color 的 run 改色，不替它编一句扩展色", () => {
    const doc = open(makeTextFillDocx());
    const out = rPrOf(retyped(doc, "只说了 w:color", { ...runAt(doc, "只说了 w:color"), color: "#0000FF" }), "只说了 w:color");
    expect(out).toContain('<w:color w:val="0000FF"/>');
    expect(out).not.toContain("w14:");
    expect(out).toContain('<w:kern w:val="28"/>');
  });

  it("套过明暗的那一句也跟着让位，而不是留一个再算不回原来的颜色的形状", () => {
    const doc = open(makeTextFillDocx());
    const out = rPrOf(retyped(doc, "明暗过一道", { ...runAt(doc, "明暗过一道"), color: "#FF0000" }), "明暗过一道");
    expect(out).toContain('<w:color w:val="FF0000"/>');
    expect(out).not.toMatch(/textFill|lumMod|lumOff/);
    expect(out).toContain('<w:sz w:val="21"/>');
  });

  it("只说过扩展色、没说过 w:color 的 run：没点色就留着，点了一色才替它让位", () => {
    const only = "只说了那句扩展色";
    expect(runAt(open(makeTextFillDocx()), only).color).toBeUndefined();
    const kept = elsewhere(makeTextFillDocx());
    expect(fillCount(kept)).toBe(fillCount(part(makeTextFillDocx())));
    expect(rPrOf(kept, only)).toContain("<w14:textFill>");
    const doc = open(makeTextFillDocx());
    const out = rPrOf(retyped(doc, only, { ...runAt(doc, only), color: "#008000" }), only);
    // 这一句同时钉住命名空间：这一组的别名要从「只能在文字命名空间里活」的孩子取。只有扩展色的 run 没有这样的
    // 证人，早先按第一个孩子猜，把整组包成了 `<w14:rPr>` —— 没有读者认得的元素。
    expect(out).toBe("<w:rPr><w:color w:val=\"008000\"/></w:rPr>");
  });

  it("别处加一段不带走任何一句扩展色，且保存两次写出同一句", () => {
    const bytes = makeTextFillDocx();
    expect(fillCount(elsewhere(bytes))).toBe(fillCount(part(bytes)));
    const once = writeDocx(open(bytes));
    expect(fillCount(part(once))).toBe(3);
    expect(part(writeDocx(open(once)))).toBe(part(once));
  });
});
