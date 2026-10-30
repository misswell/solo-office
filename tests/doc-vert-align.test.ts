import { describe, expect, it } from "vitest";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { WordDocument, applyRunPatch } from "../src/core/docx.js";
import { renderInline } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { childSpans, elementContent, OpcPackage, text } from "../src/core/opc.js";
import { makeVertAlignDocx } from "./helpers/make-vert-align.js";

/**
 * 开始页那两枚「上标/下标」（WPS `CT_Home.kuip:136` 把 `Superscript`/`Subscript` 折成一枚 `Font4`，摆在下划线、
 * 删除线之后、突出显示之前）。读、画、存三条腿在 #87/#174 那一带就有了；这一片补的是命令面与回显，所以这里钉的
 * 是**命令落下去**的样子：
 *
 * - 一条命令只改一个字段，写回时同槽只有一句 `w:vertAlign`（不许叠两句，也不许两处各写一句）。
 * - 按回去（`""`）是把槽撤掉，落回原件那半边的字节——**不**替文件写 `w:vertAlign w:val="baseline"`：WPS 真机
 *   按下标/上标那一对开关时写的是哪一句**未取证**，而「撤回」在模型里就是「这句没说过」。
 * - 明说 `baseline` 的那枚 run（§17.3.2.24 的取消极）被换成上下标时，原来那一句让位，不是并排站着。
 * - 纸上那句 `.75em` 的缩一圈是排版层，套在内层；外层那句磅数仍是文件说的那个数（字号框读得回来）。
 */

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const paraAt = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text).join("").includes(needle))!;
const runAt = (doc: WordDocument, needle: string): Inline => paraAt(doc, needle).inlines[0]!;

const PARA_RE = /<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g;
const RUN_RE = /<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*<\/w:r>/g;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(PARA_RE)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
/** One run's own `w:rPr` as written — `""` when the save gave it none. */
const rPrOf = (xml: string, needle: string): string => {
  const run = [...paraXml(xml, needle).matchAll(RUN_RE)].map((m) => m[0])[0];
  return run ? childSpans(elementContent(run)).find((k) => k.tag === "rPr")?.xml ?? "" : "";
};
const count = (xml: string, needle: string): number => (xml.match(new RegExp(needle, "g")) ?? []).length;

/**
 * The ribbon path in the editor is `restyle` → this; a test that wrote the model any other way would be a second owner.
 * The result replaces the run in the array: a withdrawal has no field to copy back (`Object.assign` would keep the old
 * one standing), and `RPR_ORDER` writing is what turns the model's silence into a saved slice.
 */
const patch = (doc: WordDocument, needle: string, change: Parameters<typeof applyRunPatch>[1]): void => {
  const para = paraAt(doc, needle);
  const at = para.inlines.findIndex((r) => r.text.includes(needle));
  para.inlines[at] = applyRunPatch(para.inlines[at]!, change);
};

describe("docx 上标/下标命令：一句 w:vertAlign，撤回是撤槽", () => {
  const bytes = makeVertAlignDocx();

  it("reads the three states the corpus keeps apart, and the silence", () => {
    const doc = open(bytes);
    expect(runAt(doc, "乙组").vertAlign).toBe("superscript");
    expect(runAt(doc, "丙组").vertAlign).toBe("subscript");
    expect(runAt(doc, "丁组").vertAlign, "明说的 baseline 是文件说过的（#87）").toBe("baseline");
    expect(runAt(doc, "甲组").vertAlign, "什么都没写就是什么都没说").toBeUndefined();
    // 一个没动过的件，存回去逐字节相同。
    expect(save(open(bytes))).toBe(part(bytes));
  });

  it("draws the raise, with the page's shrink nested inside the file's own size", () => {
    const run = runAt(open(bytes), "戊组");
    const html = renderInline(run);
    expect(html).toContain("font-size:22pt");
    // 缩一圈套在内层：同一枚 span 上写 `.75em` 会把文件那句 22 磅盖掉，字号框就读不回来了。
    expect(html).toContain("vertical-align:super;font-size:.75em;");
    expect(renderInline(runAt(open(bytes), "丙组"))).toContain("vertical-align:sub");
    expect(renderInline(runAt(open(bytes), "丁组")), "baseline 在纸上什么也不画").not.toContain("vertical-align");
  });

  it("writes one slot when the command lands on a run that said nothing", () => {
    const doc = open(bytes);
    patch(doc, "甲组", { vertAlign: "superscript" });
    const rPr = rPrOf(save(doc), "甲组");
    expect(rPr).toBe('<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>');
  });

  it("takes the slot away on the way back, and does not invent a baseline", () => {
    const doc = open(bytes);
    patch(doc, "乙组", { vertAlign: "" });
    const xml = save(doc);
    expect(rPrOf(xml, "乙组"), "撤回之后那枚 rPr 里什么都不该剩（写端对空 rPr 的既有口径）").toBe("<w:rPr/>");
    // 同一份件里丁组明写着 baseline，所以只能看乙组这一段：撤回不是往件里写一句没写过的 baseline。
    expect(paraXml(xml, "乙组")).not.toContain("baseline");
  });

  it("replaces a stated baseline instead of standing beside it", () => {
    const doc = open(bytes);
    patch(doc, "丁组", { vertAlign: "subscript" });
    const rPr = rPrOf(save(doc), "丁组");
    expect(rPr).toBe('<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>');
    expect(count(rPr, "<w:vertAlign")).toBe(1);
  });

  it("switches the pair rather than stacking them", () => {
    const doc = open(bytes);
    patch(doc, "丙组", { vertAlign: "superscript" });
    const rPr = rPrOf(save(doc), "丙组");
    expect(rPr).toBe('<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>');
    expect(rPr).not.toContain("subscript");
  });
});
