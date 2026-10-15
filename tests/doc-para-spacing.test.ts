import { describe, expect, it } from "vitest";
import type { Paragraph } from "../src/core/docx.js";
import { LINE_ONE, linePitchTwips, paraSpacingPitch, syncParaSlots, WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { childSpans, elementContent, OpcPackage, text } from "../src/core/opc.js";
import { makeSpacingDocx } from "./helpers/make-paragraph.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const findPara = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
/** 那一段自己写下的 `w:pPr` 整句：形状对不对，看字节而不是看字段。 */
const pPrOf = (xml: string, needle: string): string => {
  const para = [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
  const slot = childSpans(elementContent(para)).find((k) => k.tag === "pPr");
  return slot ? elementContent(slot.xml) : "";
};
const spacingOf = (xml: string, needle: string): string => /<w:spacing[^>]*\/>/.exec(pPrOf(xml, needle))?.[0] ?? "";
/** 段落标记 `w:rPr` 里那一枚同名的 `w:spacing`（字符间距），它跟段距不是一个东西。 */
const markSpacingOf = (xml: string, needle: string): string =>
  /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(pPrOf(xml, needle))?.[0] ?? "";

/**
 * 段前／段后／行距的**写出**侧。读与画早在 #72 与 #130~#132 就归了文件，这一族此前没有任何一条命令能改它。
 *
 * 形状不是编的：`.scratch/spacing-shape.ts` 扫过 201 份 WPS 写出的件的 10,248 枚 `w:spacing`——
 * - 属性先后只有一簇，`before, beforeLines, beforeAutospacing, after, afterLines, afterAutospacing, line, lineRule`；
 * - `w:line` 与 `w:lineRule` 从来没单独出现过（7,810 枚成对，0 枚只写一边）；
 * - 同一个间距用两笔写（缇 + 1/100 行）的有段前 2,007 枚、段后 1,850 枚，商几乎全落 312 缇一档，
 *   1,865 枚前后同档、0 枚不一致；另有 33+91 枚只说 1/100 行而不给缇，那种档量不出来。
 * 这些数与「谁能动、谁不该动」的口径写在 `syncParaSlots` 的头注里。
 *
 * 单位换算按件自己的那一档：改缇时把 1/100 行 那一笔按同一档重算（`w:before="312" w:beforeLines="100"` 改 624
 * 缇 → `w:beforeLines="200"`）；件没说档就不写这一笔，而不是照 App 的行高造一个。
 */
describe("docx 段距与行距的写出", () => {
  const DOC = makeSpacingDocx(
    [
      { spacing: 'w:before="120" w:after="60" w:line="360" w:lineRule="auto"', markRPr: '<w:rPr><w:spacing w:val="20"/></w:rPr>', text: "两笔都说的" },
      { spacing: 'w:before="312" w:beforeLines="100" w:after="312" w:afterLines="100"', text: "按行写档的" },
      { spacing: 'w:before="312"', text: "只说缇的" },
      { spacing: 'w:before="0" w:beforeAutospacing="1"', text: "让软件自定段距的" },
      { spacing: 'w:line="578" w:lineRule="exact"', text: "固定值的" },
      { spacing: 'w:line="363" w:lineRule="atLeast"', text: "最小值的" },
      { pStyle: "Lent", spacing: 'w:before="0"', text: "行距靠样式的" },
      { text: "什么都没说的" },
    ],
    '<w:spacing w:line="480" w:lineRule="auto"/>',
  );

  it("一件没动过地存一遍，那一句一个字节都不变", () => {
    const round = save(open(DOC));
    for (const needle of ["两笔都说的", "按行写档的", "只说缇的", "让软件自定段距的", "固定值的", "最小值的", "行距靠样式的", "什么都没说的"]) {
      expect(spacingOf(round, needle)).toBe(spacingOf(part(DOC), needle));
    }
    // 整段 `w:pPr` 也一样：没写字节的段落不该因为存盘而被重新拼一遍。
    expect(pPrOf(round, "两笔都说的")).toBe(pPrOf(part(DOC), "两笔都说的"));
  });

  it("没动过那三个数时，一个字节都不写", () => {
    const doc = open(DOC);
    const p = findPara(doc, "按行写档的");
    // 把字段按**读到的原值**交回去（对话框里没人改过格子就是这个形状），`moved()` 认得出这不是用户的话。
    p.spacingBefore = 312;
    p.spacingAfter = 312;
    syncParaSlots(p);
    expect(spacingOf(save(doc), "按行写档的")).toBe('<w:spacing w:before="312" w:beforeLines="100" w:after="312" w:afterLines="100"/>');
  });

  it("改段前时按件自己那一档把 1/100 行 也带走，段后原样留在后面", () => {
    const doc = open(DOC);
    const p = findPara(doc, "按行写档的");
    p.spacingBefore = 624;
    syncParaSlots(p);
    expect(spacingOf(save(doc), "按行写档的")).toBe('<w:spacing w:before="624" w:beforeLines="200" w:after="312" w:afterLines="100"/>');
  });

  it("件没给那一档时不写 1/100 行，也不凭空造一个", () => {
    const doc = open(DOC);
    const p = findPara(doc, "只说缇的");
    p.spacingBefore = 624;
    syncParaSlots(p);
    expect(spacingOf(save(doc), "只说缇的")).toBe('<w:spacing w:before="624"/>');
  });

  it("用户说了一个数，就把「让软件自定」那一笔按件里的灭法写掉", () => {
    const doc = open(DOC);
    const p = findPara(doc, "让软件自定段距的");
    expect(p.spacingBefore).toBe(0);
    p.spacingBefore = 300;
    syncParaSlots(p);
    expect(spacingOf(save(doc), "让软件自定段距的")).toBe('<w:spacing w:before="300" w:beforeAutospacing="0"/>');
  });

  it("`w:line` 与 `w:lineRule` 同进同出，且排在缇那几笔之后", () => {
    const doc = open(DOC);
    const p = findPara(doc, "两笔都说的");
    p.lineSpacing = 480;
    p.lineSpacingRule = "auto";
    syncParaSlots(p);
    expect(spacingOf(save(doc), "两笔都说的")).toBe('<w:spacing w:before="120" w:after="60" w:line="480" w:lineRule="auto"/>');
    // 段落标记里那枚 `w:spacing w:val` 说的是字符间距：整枚留在 `w:rPr` 的字节里，段距这条分支碰不到它。
    expect(markSpacingOf(save(doc), "两笔都说的")).toBe('<w:rPr><w:spacing w:val="20"/></w:rPr>');
  });

  it("固定值换成倍数时，单位跟着规则一起换", () => {
    const doc = open(DOC);
    const p = findPara(doc, "固定值的");
    p.lineSpacing = LINE_ONE;
    p.lineSpacingRule = "auto";
    syncParaSlots(p);
    expect(spacingOf(save(doc), "固定值的")).toBe('<w:spacing w:line="240" w:lineRule="auto"/>');

    const other = open(DOC);
    const q = findPara(other, "最小值的");
    q.lineSpacing = 578;
    q.lineSpacingRule = "exact";
    syncParaSlots(q);
    expect(spacingOf(save(other), "最小值的")).toBe('<w:spacing w:line="578" w:lineRule="exact"/>');
  });

  it("只动行距时，段落自己写过的段前留在原位；从样式借来的那个数不被抄进段落", () => {
    const doc = open(DOC);
    const p = findPara(doc, "行距靠样式的");
    expect(p.lineSpacing).toBe(480);
    expect(p.asRead?.lineSpacing).toBe(480);
    p.lineSpacing = 240;
    p.lineSpacingRule = "auto";
    syncParaSlots(p);
    // `w:before="0"` 是这一段自己说的（与「没说」是两件事，#132 钉过），改行距不能把它顶掉。
    expect(spacingOf(save(doc), "行距靠样式的")).toBe('<w:spacing w:before="0" w:line="240" w:lineRule="auto"/>');
  });

  it("什么都没写过的段落，改段前才第一次写下这一句", () => {
    const doc = open(DOC);
    const p = findPara(doc, "什么都没说的");
    expect(p.pPrSlots).toBeUndefined();
    p.spacingAfter = 120;
    p.lineSpacing = 480;
    p.lineSpacingRule = "auto";
    syncParaSlots(p);
    // 没有槽的段落由 `ownParaStyle` 按字段拼（`w:line` 配着默认规则一起写）。
    expect(pPrOf(save(doc), "什么都没说的")).toBe('<w:spacing w:after="120" w:line="480" w:lineRule="auto"/>');
  });

  it("改过的那一句存下去再读回来，第二次存不再改字节", () => {
    const doc = open(DOC);
    const p = findPara(doc, "按行写档的");
    p.spacingBefore = 624;
    syncParaSlots(p);
    const once = writeDocx(doc);
    const again = part(writeDocx(WordDocument.openBytes(once)));
    expect(spacingOf(again, "按行写档的")).toBe('<w:spacing w:before="624" w:beforeLines="200" w:after="312" w:afterLines="100"/>');
  });

  it("那一档量不出来时 `paraSpacingPitch` 不猜", () => {
    const doc = open(DOC);
    expect(paraSpacingPitch(findPara(doc, "按行写档的"))).toBe(312);
    expect(paraSpacingPitch(findPara(doc, "只说缇的"))).toBeUndefined();
    expect(paraSpacingPitch(findPara(doc, "什么都没说的"))).toBeUndefined();
    expect(linePitchTwips("0", "100")).toBeUndefined();
    expect(linePitchTwips("312", "0")).toBeUndefined();
  });
});
