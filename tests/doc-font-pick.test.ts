import { describe, expect, it } from "vitest";
import type { Inline, Paragraph, PropSlot } from "../src/core/docx.js";
import { WordDocument, applyRunPatch, pickedFaceCoversHan } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeChainDocx } from "./helpers/make-paragraph.js";
import { makeThemeDocx } from "./helpers/make-theme.js";

/**
 * §17.3.2.26 把一个 run 的字按书写系统分给三张脸：`w:ascii` 管 ASCII 那一段，`w:hAnsi` 管高 ANSI，`w:eastAsia`
 * 管中日韩。工具栏的字体框说的却是一句话——「这个 run 的脸是 X」——所以这一句该落到哪几槽，得先看件自己怎么说。
 * 数的是本机 WPS 写出的 201 份件的 `word/document.xml`（`.scratch/font-slot-shape.ts`，32,176 个 run，其中 29,643
 * 个带自己的 `w:rFonts`，**没有一个 run 带两枚**）：
 *
 * - 正文带汉字的 run：两槽都写名的 **7,706** 枚里 7,659 枚两张同名、**47** 枚各写一张；只写拉丁槽的 **2,331** 枚；
 *   只写汉字槽的 **34** 枚；两槽都没写的 **5,444** 枚。
 * - 正文纯拉丁的 run：两槽都写的 **1,062** 枚**全部同名**（不同值 0 枚）；只写拉丁槽 13 枚；都没写 5,483 枚。
 * - `w:cs` 与 `w:ascii` 同名 9,843 枚、不同名 2,508 枚——复合文字那一槽因此不是「同一句话的另一半」，模型也没有
 *   它的字段，挑脸时一字不动。
 *
 * 于是这一轮的规矩：件里已经给这个 run 留了汉字槽（写名字或指主题都算留），除非它明说两张不同的脸，挑的那张就接着
 * 管这一槽；件里没留，只有这段文字自己带汉字时才替它补上——不给一段纯拉丁的 run 写一句关于中文的话。真机 WPS 的
 * 字体框改一次究竟写哪几槽，仍未取证。
 */

const slots = (attrs: string, prefix = "w"): PropSlot[] => [{ tag: "rFonts", xml: `<${prefix}:rFonts ${attrs}/>` }];

/**
 * 一次「挑脸」按 ribbon 的做法改模型：走的是生产那一条 `applyRunPatch`（`restyle` 逐 run 调的就是它），这里只把
 * 改出来的那份贴回 run。抄一份摊开的写法就等于给这条规矩第二个主人——第一轮正是这样：本地那份 `pick` 自己问
 * `pickedFaceCoversHan`，把生产里那句让位拿掉之后 17 例照样全绿。
 */
const pick = (run: Inline, face: string): void => {
  Object.assign(run, applyRunPatch(run, { font: face }));
};

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const runsOf = (doc: WordDocument): Inline[] =>
  doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph").flatMap((p) => p.inlines);
const runAt = (doc: WordDocument, needle: string): Inline => runsOf(doc).find((r) => r.text.includes(needle))!;
/** 该段 run 自己那枚 `w:rFonts`：夹具的段落不写 `w:pPr`，所以正文里第一个就是 run 的。 */
const tagOf = (body: string, needle: string): string =>
  /<w:rFonts\b[^>]*>/.exec(body.split("</w:p>").find((p) => p.includes(needle)) ?? "")?.[0] ?? "";
const tags = (xml: string): string[] => [...xml.matchAll(/<[a-zA-Z][-\w.]*:rFonts\b[^>]*/g)].map((m) => m[0]);
const SLOTS: [string, string][] = [
  ["w:ascii", "w:asciiTheme"],
  ["w:hAnsi", "w:hAnsiTheme"],
  ["w:eastAsia", "w:eastAsiaTheme"],
  ["w:cs", "w:cstheme"],
];
const stated = (tag: string, name: string): boolean => new RegExp(`${name}="[^"]+`).test(tag);
const collisions = (xml: string): number => tags(xml).filter((t) => SLOTS.some(([lit, them]) => stated(t, lit) && stated(t, them))).length;

describe("docx 挑一张脸管到哪一槽", () => {
  it("跟着管一件里两张同名的汉字槽", () => {
    // 7,659 枚的形状：件里两张名字一样，那一句「这个 run 的脸」本就同时管着两槽。
    expect(pickedFaceCoversHan(slots('w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri"'), "汉字正文")).toBe(true);
  });

  it("不碰一件里明说了两张不同脸的汉字槽", () => {
    // 那 47 枚各写一张的 run：改拉丁那一张不等于把中文那一张也换了。
    expect(pickedFaceCoversHan(slots('w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"'), "汉字正文")).toBe(false);
  });

  it("汉字正文里只写了拉丁槽的 run 也算留了汉字槽", () => {
    // 2,331 枚：件只说了一半，但这段字确实穿的是这个 run 的脸。只写 `w:hAnsi` 的也一样——那一槽同属拉丁这一段。
    expect(pickedFaceCoversHan(slots('w:ascii="Times New Roman" w:hAnsi="Times New Roman"'), "汉字正文")).toBe(true);
    expect(pickedFaceCoversHan(slots('w:hAnsi="Times New Roman"'), "汉字正文")).toBe(true);
  });

  it("纯拉丁正文里只写了拉丁槽的 run 不领汉字槽", () => {
    // 1,062 枚两槽同名与 13 枚只写拉丁的都在这里：一段没有中文的话不该替它宣布一张中文脸。
    expect(pickedFaceCoversHan(slots('w:ascii="Times New Roman" w:hAnsi="Times New Roman"'), "Latin only")).toBe(false);
    expect(pickedFaceCoversHan(slots('w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri"'), "Latin only")).toBe(true);
  });

  it("汉字槽指着主题时也算留着，无论这段字是不是中文", () => {
    // 这一槽已经在说话，只是说的是「我的脸来自主题那一列」；挑脸要接这句话就得把它换掉，而不是绕过去。
    const themed = slots('w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"');
    expect(pickedFaceCoversHan(themed, "汉字正文")).toBe(true);
    expect(pickedFaceCoversHan(themed, "Latin only")).toBe(true);
  });

  it("两槽都没写的 run 按这段字自己带不带汉字分", () => {
    expect(pickedFaceCoversHan(undefined, "汉字正文")).toBe(true);
    expect(pickedFaceCoversHan([], "Latin only")).toBe(false);
    // 假名与谚文同属 §17.3.2.26 的东亚那一类，与 `render.ts` 的 `unicode-range` 一套码段。
    expect(pickedFaceCoversHan(undefined, "かな")).toBe(true);
    expect(pickedFaceCoversHan(undefined, "한글")).toBe(true);
  });

  it("把 w:hint 当成不是这一槽的说法", () => {
    // 28,074 个 run 带 `w:hint`（26,922 枚写 eastAsia）。按 §17.3.2.26 那是分派提示，不是件给这一槽留的位置，
    // 所以纯拉丁的正文仍不领汉字槽——与 #138 里 `标记只说hint` 同一口径。
    expect(pickedFaceCoversHan(slots('w:ascii="Calibri" w:hAnsi="Calibri" w:hint="eastAsia"'), "Latin only")).toBe(false);
  });

  it("复合文字那一槽不参与这句话", () => {
    // `w:cs` 与 `w:ascii` 不同名的有 2,508 枚：两张脸各管一段，挑一次脸不顺手把第三槽也写了。
    const withCs = slots('w:ascii="Calibri" w:eastAsia="Calibri" w:cs="Arial"');
    expect(pickedFaceCoversHan(withCs, "汉字正文")).toBe(true);
    expect(withCs[0]!.xml).toContain("w:cs");
  });

  it("认得原件用的别的名字空间别名", () => {
    // `w` 只是绑在 words 命名空间上的一个别名（`carriedRunPrefix` 因此从孩子身上取别名而不假定它）。
    expect(pickedFaceCoversHan(slots('a:ascii="Calibri" a:hAnsi="Calibri" a:eastAsia="宋体"', "a"), "汉字正文")).toBe(false);
    expect(pickedFaceCoversHan(slots('a:ascii="Calibri" a:eastAsiaTheme="minorEastAsia"', "a"), "Latin only")).toBe(true);
  });
});

describe("docx 挑一张脸保存出的样子", () => {
  const sameName = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri"/>';
  const twoNames = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/>';
  const latinName = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>';
  /** 没点到的那一段用它自己那枚标签原样回来，所以它写的必须是件里的拼法。 */
  const kept = '<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:eastAsia="Cambria"/>';

  it("两件同名时把两槽一起换成挑的那张，没点到的那段一字不动", () => {
    const bytes = makeChainDocx([], [{ text: "改这段", rPr: sameName }, { text: "没这段", rPr: kept }]);
    const doc = open(bytes);
    pick(runAt(doc, "改这段"), "Times New Roman");
    const body = part(writeDocx(doc));
    const tag = tagOf(body, "改这段");
    expect(tag).toContain('w:ascii="Times New Roman"');
    expect(tag).toContain('w:hAnsi="Times New Roman"');
    expect(tag).toContain('w:eastAsia="Times New Roman"');
    expect(tagOf(body, "没这段")).toBe(kept);
    expect(collisions(body)).toBe(0);
    const again = runAt(open(writeDocx(open(bytes))), "改这段");
    expect(again.font).toBe("Calibri");
    expect(again.fontEastAsia).toBe("Calibri");
  });

  it("件里明说两张脸时只挪拉丁那一槽，另一张读得回", () => {
    const bytes = makeChainDocx([], [{ text: "改这段", rPr: twoNames }]);
    const doc = open(bytes);
    pick(runAt(doc, "改这段"), "Times New Roman");
    const tag = tagOf(part(writeDocx(doc)), "改这段");
    expect(tag).toContain('w:ascii="Times New Roman"');
    expect(tag).toContain('w:hAnsi="Times New Roman"');
    // 汉字那一槽是件自己另说的一句话，原样带走。
    expect(tag).toContain('w:eastAsia="宋体"');
    const again = runAt(open(writeDocx(doc)), "改这段");
    expect(again.font).toBe("Times New Roman");
    expect(again.fontEastAsia).toBe("宋体");
  });

  it("只写了拉丁槽的 run 在汉字正文下补出汉字槽", () => {
    // 2,331 枚的形状：件只说了一半，这一句把另一半也说了。
    const bytes = makeChainDocx([], [{ text: "改这段", rPr: latinName }]);
    const doc = open(bytes);
    pick(runAt(doc, "改这段"), "Times New Roman");
    const body = part(writeDocx(doc));
    expect(tagOf(body, "改这段")).toContain('w:eastAsia="Times New Roman"');
    expect(collisions(body)).toBe(0);
  });

  it("只写了拉丁槽的 run 在纯拉丁正文下不写汉字槽", () => {
    const bytes = makeChainDocx([], [{ text: "改这段", rPr: latinName }]);
    const doc = open(bytes);
    const run = runAt(doc, "改这段");
    run.text = "Latin only";
    pick(run, "Times New Roman");
    const tag = tagOf(part(writeDocx(doc)), "Latin only");
    expect(tag).not.toContain("eastAsia");
    // 地板给的那张中文脸仍留在地板里：这个 run 没为自己宣布过它。
    expect(tag).not.toContain("宋体");
  });

  it("件里没写 `w:rFonts` 的 run 只写挪了的那一槽，不替它认领借来的中文脸", () => {
    // 5,444 / 5,483 枚「两槽都没写」的形状，地板是 67 of 86 那份 Calibri/宋体（`makeChainDocx` 的默认地板）。
    const bytes = makeChainDocx([], [{ text: "改这段" }, { text: "别的一段" }]);
    const doc = open(bytes);
    const run = runAt(doc, "改这段");
    run.text = "Latin only";
    pick(run, "Times New Roman");
    // 模型里那张借来的脸照在（页面要靠它画），只是不该由这个 run 写进件里。
    expect(run.fontEastAsia).toBe("宋体");
    const body = part(writeDocx(doc));
    expect(tagOf(body, "Latin only")).toBe('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>');
    expect(tagOf(body, "别的一段")).toBe("");
  });

  it("汉字正文里没写脸的 run 这一句连着汉字槽一起写", () => {
    const body = saved(makeChainDocx([], [{ text: "改这段" }]), (r) => pick(r, "Times New Roman"));
    expect(tagOf(body, "改这段")).toBe('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman"/>');
  });

  it("汉字槽指着主题的 run 被挑脸时那句主题引用让位，别槽的一句不少", () => {
    const doc = open(makeThemeDocx());
    pick(runAt(doc, "整行交给主题"), "Times New Roman");
    const body = part(writeDocx(doc));
    const tag = tagOf(body, "整行交给主题");
    expect(tag).toContain('w:ascii="Times New Roman"');
    expect(tag).toContain('w:eastAsia="Times New Roman"');
    expect(tag).not.toContain("asciiTheme");
    expect(tag).not.toContain("hAnsiTheme");
    expect(tag).not.toContain("eastAsiaTheme");
    // 复合文字那一槽模型没有字段，那句「来自主题」原样带走。
    expect(tag).toContain('w:cstheme="minorBidi"');
    expect(collisions(tag)).toBe(0);
    // 全篇仍是夹具自己那一枚：`字面先赢` 那次编辑没参与，它的拉丁槽一字没挪，让位规则因此不叫它——挪槽才让位。
    expect(collisions(body)).toBe(1);
    const again = runAt(open(writeDocx(doc)), "整行交给主题");
    expect(again.font).toBe("Times New Roman");
    expect(again.fontEastAsia).toBe("Times New Roman");
  });

  it("给每一段都挑一次脸，正文里也不出现同槽并存", () => {
    const doc = open(makeThemeDocx());
    // 挑的脸必须与每段解析出的那张不同：同名时那一槽没挪，件里原有的那句主题引用也就不该让位。
    for (const run of runsOf(doc)) pick(run, "KaiTi");
    const body = part(writeDocx(doc));
    // 夹具故意写了一枚 `字面先赢`（同槽两张），语料里那种形状是 0 枚：一次挑脸把它换掉，且不再多出任何一枚。
    expect(collisions(part(makeThemeDocx()))).toBe(1);
    expect(collisions(body)).toBe(0);
    expect(tags(body).filter((t) => t.includes('w:cstheme="minorBidi"')).length).toBeGreaterThan(0);
    // 挪过的每一段都还读得回自己那张新脸，没挪的槽仍从主题读。
    expect(runAt(open(writeDocx(doc)), "字面先赢").font).toBe("KaiTi");
    expect(runAt(open(writeDocx(doc)), "地板给的脸").font).toBe("KaiTi");
  });
});

/** 只改「改这段」那一段，其余段落留给上面的断言当对照。 */
function saved(bytes: Uint8Array, mutate: (run: Inline) => void): string {
  const doc = open(bytes);
  mutate(runAt(doc, "改这段"));
  return part(writeDocx(doc));
}
