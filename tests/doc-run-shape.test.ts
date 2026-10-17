import { describe, expect, it } from "vitest";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { WordDocument, applyFacePick, applyRunPatch } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeTabRunDocx } from "./helpers/make-run.js";

/**
 * §17.3 的 `EG_RunInnerContent` 是无上限序列：一枚 `w:r` 可以并排装 `w:tab`、`w:br`、好几段 `w:t`。读入侧按孩子
 * 给 inline，写回侧若仍逐枚 inline 发一枚 run，文件里那一枚就变成两枚，各带一份 `w:rPr`——`w:cs`、`w:hint`
 * 这类模型没有字段的槽跟着写两遍。受众按本机 WPS 写出的 201 件正文数（`.scratch/run-split-scope.ts`）：32,207 枚
 * run 里 **2** 枚多孩子，形状都是 `tab,t`，**2** 枚都带自己的 `w:rFonts`，带 `w:rPrChange` 的 **0** 枚。改脸前
 * `w:cs` 27,156 枚、`w:hint` 47,208 枚，旧写法改一遍脸各多 2 枚；并回一枚 run 之后两处增量都是 0。
 *
 * 这条件里另备了两个方向：同名同式的**两枚**兄弟 run 不能并（并的根据是同一枚源 run，不是同一串属性），半句改了
 * 格式的同源两孩子也不能并。链接与修订里那两例这 201 件里没有受众（没有带 tab 的 `w:hyperlink`/`w:ins`），建的
 * 是 wrapper 只套一层的规矩——合并的落笔点若把整枚 run 再包一遍，`w:hyperlink` 就成了嵌套。
 *
 * 并回还有一件前置：#140 那句「件里这段说的是不是中文」得按**整枚 run** 问（`Inline.runText`，读入侧在多个孩子
 * 之间给同一份引用）。按孩子问时，件里没预留 `w:eastAsia` 的那一枚 run 挑一遍脸只给带字的孩子补上那一槽，
 * tab/br 那半句不补，两串性质不同就并不回去——夹具里这样的三枚（br+t、链接里 tab+t、修订里 tab+t）在 #143 之前
 * 会把 6 枚 run 写成 9 枚、`w:cs` 多写 3 份。
 */

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const parasOf = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const runsOf = (doc: WordDocument): Inline[] => parasOf(doc).flatMap((p) => p.inlines);
const count = (xml: string, re: RegExp): number => xml.match(re)?.length ?? 0;
/** 挑脸走生产那条 `applyRunPatch`：`restyle` 逐 run 调的就是它，抄一份摊开的写法就是第二个主人。 */
const pick = (run: Inline, face: string): void => {
  Object.assign(run, applyRunPatch(run, { font: face }));
};
/** 写回件里装了这段文字的那枚 run。 */
const runHolding = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:r>[\s\S]*?<\/w:r>/g)].map((m) => m[0]).find((r) => r.includes(needle)) ?? "";

describe("docx 一枚源 run 写回仍是一枚 run", () => {
  /** 只改 p2 的头一枚：全文因此走重建路径，其余各枚的性质串与原件一样。 */
  const rebuildOne = (): string => {
    const doc = open(makeTabRunDocx());
    const target = parasOf(doc)[1].inlines[0];
    Object.assign(target, applyRunPatch(target, { bold: true }));
    return part(writeDocx(doc));
  };

  it("tab 挨着字：并回一枚 run，两个孩子按原顺序，性质只一份", () => {
    const out = rebuildOne();
    const run = runHolding(out, "挨着 tab 的字");
    expect(run).toContain("<w:tab/>");
    expect(count(run, /<w:rPr\b/g)).toBe(1);
    expect(count(run, /<w:rFonts\b/g)).toBe(1);
    expect(run).toMatch(/<w:tab\/><w:t>挨着 tab 的字<\/w:t>/);
  });

  it("换页符挨着字同样并回；同名同式的两枚兄弟 run 仍是两枚", () => {
    const out = rebuildOne();
    expect(runHolding(out, "分页符后的字")).toMatch(/<w:br w:type="page"\/><w:t>分页符后的字<\/w:t>/);
    expect(out).toContain("<w:r><w:rPr><w:rStyle w:val=\"Hyperlink\"/><w:rFonts w:cs=\"Calibri\"/></w:rPr><w:tab/><w:t>链接里的 tab</w:t></w:r>");
    expect(count(out, /<w:r\b/g)).toBe(6);
    expect(count(out, /<w:rFonts\b/g)).toBe(6);
    const halves = [...out.matchAll(/<w:r>[\s\S]*?<\/w:r>/g)].map((m) => m[0]).filter((r) => r.includes("同名同式的"));
    expect(halves).toHaveLength(2);
  });

  it("链接与修订里的 tab 挨着字：wrapper 只一层，里面一枚 run", () => {
    const out = rebuildOne();
    expect(count(out, /<w:hyperlink\b/g)).toBe(1);
    expect(count(out, /<w:ins\b/g)).toBe(1);
    const link = /<w:hyperlink[^>]*>[\s\S]*?<\/w:hyperlink>/.exec(out)![0];
    expect(count(link, /<w:r\b/g)).toBe(1);
    const ins = /<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/.exec(out)![0];
    expect(count(ins, /<w:r\b/g)).toBe(1);
  });

  it("整件改一遍脸：六枚 run 仍是六枚，`w:cs` 与 `w:hint` 一枚不加", () => {
    // 并回的前提是两半拼出的性质串逐字相同，而 #140 那句「这段说的是不是中文」问的是一枚 run，不是它的某个
    // 孩子：`w:tab`/`w:br` 那半句自己不带字，同一枚 run 里写的是中文，所以它也该拿到 `w:eastAsia`。问法按孩子
    // 时，件里没预留汉字槽的那种 run（`w:rFonts` 只写拉丁槽）改一遍脸会分裂成两枚、`w:cs` 跟着多写一份；
    // 现在四枚多孩子的 run（tab+t、br+t、链接里 tab+t、修订里 tab+t）全部并着回去。
    const before = part(makeTabRunDocx());
    const doc = open(makeTabRunDocx());
    for (const r of runsOf(doc)) pick(r, "黑体");
    const out = part(writeDocx(doc));
    expect(count(out, /w:cs=/g)).toBe(count(before, /w:cs=/g));
    expect(count(out, /w:hint=/g)).toBe(count(before, /w:hint=/g));
    expect(count(out, /<w:r>/g)).toBe(count(before, /<w:r>/g));
    expect(count(out, /<w:rFonts\b/g)).toBe(count(before, /<w:rFonts\b/g));
    expect(runHolding(out, "挨着 tab 的字")).toMatch(/<w:tab\/><w:t>挨着 tab 的字<\/w:t>/);
    expect(runHolding(out, "分页符后的字")).toMatch(/<w:br w:type="page"\/><w:t>分页符后的字<\/w:t>/);
    expect(runHolding(out, "链接里的 tab")).toMatch(/<w:tab\/><w:t>链接里的 tab<\/w:t>/);
    expect(runHolding(out, "修订里的 tab")).toMatch(/<w:tab\/><w:t>修订里的 tab<\/w:t>/);
    expect(count(runHolding(out, "挨着 tab 的字"), /<w:rFonts\b/g)).toBe(1);
    expect(runHolding(out, "挨着 tab 的字")).toContain('w:ascii="黑体"');
  });

  it("汉字这句问的是整枚 run，不是「有字就要写中文」", () => {
    // 同一套规矩的反面：run 通篇不带汉字时，`w:tab` 那半句仍然不该被补上一句关于中文的话。
    const slots = [{ tag: "rFonts", xml: '<w:rFonts w:ascii="Calibri" w:cs="Calibri"/>' }];
    expect(applyFacePick({ text: "", rPrSlots: slots, runText: "abc" }, "黑体")).toEqual({ font: "黑体" });
    expect(applyFacePick({ text: "", rPrSlots: slots, runText: "ab 中" }, "黑体")).toEqual({ font: "黑体", fontEastAsia: "黑体" });
    // 单孩子的 run 没有 `runText`，照旧按自己那句判。
    expect(applyFacePick({ text: "中文", rPrSlots: slots }, "黑体")).toEqual({ font: "黑体", fontEastAsia: "黑体" });
    expect(applyFacePick({ text: "word", rPrSlots: slots }, "黑体")).toEqual({ font: "黑体" });
  });

  it("半句改了格式：同源的两孩子各写各的性质，仍是两枚，但一句都不丢", () => {
    const doc = open(makeTabRunDocx());
    const [, textInline] = parasOf(doc)[0].inlines;
    Object.assign(textInline, applyRunPatch(textInline, { bold: true }));
    const out = part(writeDocx(doc));
    // 前半（只有 tab）保持原件那一串性质，后半带上了粗体：两枚各有各的 `w:rPr`，不该并。
    expect(out).toContain("<w:rFonts w:hint=\"eastAsia\" w:ascii=\"Times New Roman\" w:eastAsia=\"宋体\" w:cs=\"Times New Roman\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"28\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\"/></w:rPr><w:tab/>");
    expect(runHolding(out, "挨着 tab 的字")).toContain("<w:b/>");
    expect(count(out, /<w:tab\/>/g)).toBe(3);
  });

  it("写回再读：tab、换页符与文字都在，顺序不错", () => {
    const doc = open(makeTabRunDocx());
    for (const r of runsOf(doc)) pick(r, "黑体");
    const runs = runsOf(open(writeDocx(doc)));
    expect(runs.filter((r) => r.tab)).toHaveLength(3);
    expect(runs.filter((r) => r.break)).toHaveLength(1);
    const text = runs.map((r) => r.text).join("|");
    expect(text).toContain("挨着 tab 的字");
    expect(text).toContain("修订里的 tab");
    expect(text).toContain("同名同式的后半");
  });
});
