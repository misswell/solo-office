import { describe, expect, it } from "vitest";
import { WordDocument, type Paragraph } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { flatOf, inlineLength } from "../src/core/run-identity.js";
import { CASE_OPTIONS, caseText, changedRanges, type CaseMode } from "../src/core/doc-case.js";
import { makeCaseDocx } from "./helpers/make-case.js";

const WITH_MARK = "hello\u0000world";

describe("docx 更改大小写：字面层，长度不许动", () => {
  it("upper/lower/toggle only touch the cased letters", () => {
    expect(caseText("你好 World 123", "upper")).toBe("你好 WORLD 123");
    expect(caseText("你好 WORLD 123", "lower")).toBe("你好 world 123");
    expect(caseText("Hello World", "toggle")).toBe("hELLO wORLD");
    expect(caseText("Ω ω 7", "toggle")).toBe("ω Ω 7");
  });

  it("keeps the length, so the offsets and the runs stand where they were", () => {
    // `ß` 的大写是两枚字（SS）、`İ` 的小写带一枚组合点：换写面就换长度的那几个字原样留着。
    for (const mode of ["upper", "lower", "title", "sentence", "toggle"] as CaseMode[]) {
      for (const text of ["straße İ. 你好", "a\u0000b", "ǆ"] as string[]) {
        expect(caseText(text, mode).length).toBe(text.length);
      }
    }
  });

  it("sentence case restarts at the punctuation the file speaks, and at a mark", () => {
    expect(caseText("hello world. how are you? i am fine", "sentence")).toBe("Hello world. How are you? I am fine");
    expect(caseText("中文。english 句子", "sentence")).toBe("中文。English 句子");
    expect(caseText(WITH_MARK, "sentence")).toBe("Hello\u0000World");
    // 逗号不是句界：它后面那个词仍按句中处理。
    expect(caseText("one, two. three", "sentence")).toBe("One, two. Three");
    // 汉字把「这一句的开头」用完：它后面那个英文词按句中处理。
    expect(caseText("。你好 hello", "sentence")).toBe("。你好 hello");
    // 数字与空格不算开头，也不算结束开头。
    expect(caseText("end. 123abc", "sentence")).toBe("End. 123Abc");
  });

  it("title case starts a word where a letter or digit stops", () => {
    expect(caseText("the quick-brown FOX", "title")).toBe("The Quick-Brown Fox");
    expect(caseText("hello 123world", "title")).toBe("Hello 123World");
    // 撇号不是字母：Word 同此读数。
    expect(caseText("don't stop", "title")).toBe("Don'T Stop");
    expect(caseText(WITH_MARK, "title")).toBe("Hello\u0000World");
  });

  it("leaves Han alone in every one of the five modes", () => {
    for (const mode of ["sentence", "lower", "upper", "title", "toggle"] as CaseMode[]) {
      expect(caseText("更改大小写：汉字没有", mode)).toBe("更改大小写：汉字没有");
    }
  });

  it("reports the档位 the ribbon shows, with the placeholder first", () => {
    expect(CASE_OPTIONS[0]).toEqual(["", "更改大小写"]);
    const modes = CASE_OPTIONS.slice(1).map(([value]) => value);
    expect(modes).toEqual(["sentence", "lower", "upper", "title", "toggle"]);
    expect(CASE_OPTIONS.slice(1).every(([value, label]) => label !== value && /[\u4e00-\u9fff]/.test(label))).toBe(true);
  });
});

describe("docx 更改大小写：逐区间，汉字那一段不许被划进修订", () => {
  it("finds the maximal contiguous differences", () => {
    expect(changedRanges("你好 WORLD hello", caseText("你好 WORLD hello", "upper"))).toEqual([[9, 14]]);
    expect(changedRanges("the cat", caseText("the cat", "title"))).toEqual([[0, 1], [4, 5]]);
  });

  it("says nothing when nothing changed", () => {
    expect(changedRanges("你好", caseText("你好", "upper"))).toEqual([]);
    expect(changedRanges("", caseText("", "upper"))).toEqual([]);
  });

  it("covers every differing character exactly once, in order", () => {
    for (const mode of ["sentence", "lower", "upper", "title", "toggle"] as CaseMode[]) {
      const before = "Hello World 你好 AGAIN 123 don't";
      const after = caseText(before, mode);
      const ranges = changedRanges(before, after);
      let seen = "";
      let last = 0;
      for (const [from, to] of ranges) {
        expect(from).toBeGreaterThanOrEqual(last);
        expect(before.slice(from, to)).not.toBe(after.slice(from, to));
        for (let i = from; i < to; i += 1) expect(before[i]).not.toBe(after[i]);
        for (const [f2, t2] of ranges) if (f2 !== from) expect(f2 >= to || t2 <= from).toBe(true);
        seen += after.slice(from, to);
        last = to;
      }
      let all = "";
      for (let i = 0; i < before.length; i += 1) all += before[i] !== after[i] ? after[i] : "";
      expect(seen).toBe(all);
    }
  });
});

/**
 * 装机烟测与纸面探针共用的那件（`tests/helpers/make-case.ts`）在这一层要问的是**偏移空间的事实**：
 * `changeCase` 整条算术建在「标记占一枚字符、但没有字面」上——那一枚若是两枚字符或零枚，每一格的落点就全偏了。
 */
describe("docx 更改大小写：装机夹具那四段的偏移形状", () => {
  const open = (): WordDocument => WordDocument.openBytes(makeCaseDocx());
  const paraAt = (doc: WordDocument, needle: string): Paragraph =>
    doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph").find((p) => flatOf(p.inlines).includes(needle))!;

  it("the tab run occupies one offset and spells no literal", () => {
    const yi = paraAt(open(), "乙组");
    expect(yi.inlines.map((r) => r.text)).toEqual(["乙组 one", "", "number"]);
    expect(yi.inlines.map((r) => inlineLength(r))).toEqual([6, 1, 6]);
    // 那枚 NUL 就是标记在偏移空间里的化身：改写不许挪它，也不许把它当成一格字面。
    const flat = flatOf(yi.inlines);
    expect(flat).toBe("乙组 one" + "\u0000" + "number");
    expect(caseText(flat, "upper")).toBe("乙组 ONE" + "\u0000" + "NUMBER");
  });

  it("the sharp-s paragraph keeps its length through all five modes", () => {
    const flat = flatOf(paraAt(open(), "丙组").inlines);
    for (const mode of ["sentence", "lower", "upper", "title", "toggle"] as CaseMode[]) {
      expect(caseText(flat, mode).length).toBe(flat.length);
    }
    expect(caseText(flat, "upper")).toBe("丙组 STRAßE 大写");
  });

  it("writes an untouched fixture back as the same four paragraphs", () => {
    const doc = open();
    const again = WordDocument.openBytes(writeDocx(doc));
    for (const needle of ["甲组", "乙组", "丙组", "丁组"]) {
      expect(flatOf(paraAt(again, needle).inlines)).toBe(flatOf(paraAt(doc, needle).inlines));
    }
    // 丁组那两句格式是这一族「不许碰」的那一半：夹具要带着它们进壳，纸面才谈得上「改字面不丢格式」。
    expect(paraAt(again, "丁组").inlines[0]!.bold).toBe(true);
    expect(paraAt(again, "丁组").inlines[0]!.sizePt).toBe(16);
  });
});