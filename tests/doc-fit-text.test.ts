import { describe, expect, it } from "vitest";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { childSpans, elementContent, OpcPackage, text } from "../src/core/opc.js";
import { renderInline } from "../src/core/render.js";
import { makeFitTextDocx } from "./helpers/make-run.js";

/**
 * `w:rPr/w:fitText`（调整宽度，§17.3.2.14）。授权 46 件里 43 枚 / 2 件——`business plan.dotm` 的 zh_CN 与 zh_HK，
 * 每段 3~4 枚同 `id`、`w:val` 全 `7791`，摊开的是封面表单那条填空线。这一片钉三件事：那枚孩子在**自己**的槽里
 * （模型不给它字段）、纸上带着组的名字与目标宽（怎么摊由真 Chrome 探针 `fittext-dom` 量）、写回时不多写一句。
 */

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const paraAt = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text).join("").includes(needle))!;
const runAt = (doc: WordDocument, needle: string, at = 0): Inline => paraAt(doc, needle).inlines[at]!;

const PARA_RE = /<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g;
const RUN_RE = /<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*<\/w:r>/g;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(PARA_RE)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
const rPrOf = (xml: string, needle: string, at = 0): string => {
  const run = [...paraXml(xml, needle).matchAll(RUN_RE)].map((m) => m[0])[at];
  const slot = run ? childSpans(elementContent(run)).find((k) => k.tag === "rPr") : undefined;
  return slot?.xml ?? "";
};
const count = (xml: string, tag: string): number => (xml.match(new RegExp(`<w:${tag}\\b`, "g")) ?? []).length;

describe("w:rPr/w:fitText 调整宽度", () => {
  const bytes = makeFitTextDocx();

  it("rides in the run's own property set, unmodelled", () => {
    const doc = open(bytes);
    const slots = runAt(doc, "项目名称").rPrSlots ?? [];
    expect(slots.find((s) => s.tag === "fitText")?.xml, "the element travels as bytes").toBe(
      '<w:fitText w:val="7791" w:id="0"/>',
    );
    expect(slots.find((s) => s.tag === "w"), "字符缩放 is the neighbour, not the same slot").toBeDefined();
  });

  it("stamps the group's name and the width the file stated on the run", () => {
    const doc = open(bytes);
    const label = renderInline(runAt(doc, "项目名称"));
    expect(label).toContain('data-fit="0"');
    expect(label).toContain('data-fit-tw="7791"');
    const filled = renderInline(runAt(doc, "项目名称", 1));
    expect(filled, "a run without `w:spacing` keeps the drawer free of letter-spacing").not.toContain("letter-spacing");
  });

  it("keeps the run's own 字距 out of the drawer's letter-spacing", () => {
    // 摊开的字距是量出来的，不是文件说的话；编辑器把内联 `letter-spacing` 读回 run 的 `w:spacing`，所以那一句
    // 只留给文件自己写下的值，fit 那一组把它挪进 `--fit-char-ls`。
    const doc = open(bytes);
    const spaced = renderInline(runAt(doc, "项目名称", 2));
    expect(spaced).toContain("--fit-char-ls:4.1pt");
    expect(spaced).not.toContain("letter-spacing");
    expect(spaced).toContain('data-fit="0"');
  });

  it("draws no group the file did not ask for", () => {
    const doc = open(bytes);
    expect(renderInline(runAt(doc, "零宽的不成组")), "w:val=0 is not a width").not.toContain("data-fit");
  });

  it("gives a run with no `w:id` a group of its own", () => {
    const doc = open(bytes);
    const first = renderInline(runAt(doc, "没有号"));
    const second = renderInline(runAt(doc, "也自成组"));
    expect(first).toContain('data-fit=""');
    expect(second).toContain('data-fit=""');
    expect(first).toContain('data-fit-tw="3000"');
  });

  it("writes the element back once, per run, as the file spelled it", () => {
    const out = save(open(bytes));
    expect(count(out, "fitText"), "3 + 1 + 2 + 1 + 1, each once").toBe(8);
    expect(rPrOf(out, "项目名称")).toContain('<w:fitText w:val="7791" w:id="0"/>');
    expect(rPrOf(out, "零宽的不成组")).toContain('<w:fitText w:val="0" w:id="1"/>');
    expect(rPrOf(out, "没有号")).toContain('<w:fitText w:val="3000"/>');
    expect(rPrOf(out, "WWWWWWWW"), "a width narrower than the text is said too").toContain(
      '<w:fitText w:val="200" w:id="9"/>',
    );
  });

  it("keeps the group through a keystroke inside one of its runs", () => {
    const doc = open(bytes);
    paraAt(doc, "项目名称").inlines[0] = { ...runAt(doc, "项目名称"), text: "项目名称打过" };
    const out = save(doc);
    expect(rPrOf(out, "打过")).toContain('<w:fitText w:val="7791" w:id="0"/>');
    expect(count(out, "fitText")).toBe(8);
  });

  it("moves the run's own 字距 without touching the group or doubling the element", () => {
    const doc = open(bytes);
    runAt(doc, "项目名称", 2).charSpacingTwips = 120;
    const out = save(doc);
    const rpr = rPrOf(out, "项目名称", 2);
    expect(rpr, "the field re-spells the element it moved").toContain('<w:spacing w:val="120"/>');
    expect(rpr, "the group is not the field's to drop").toContain('<w:fitText w:val="7791" w:id="0"/>');
    expect((rpr.match(/<w:spacing\b/g) ?? []).length, "one sentence, not the source bytes plus a copy").toBe(1);
    expect(count(out, "fitText")).toBe(8);
  });
});
