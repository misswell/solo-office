import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import { strToU8, zipSync } from "fflate";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";

/** 一件只有正文的件：`w:body` 由调用方写死，页边一切按最小形状来。 */
const docOf = (body: string): WordDocument =>
  WordDocument.openBytes(
    zipSync(
      {
        "[Content_Types].xml": strToU8(
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
        ),
        "_rels/.rels": strToU8(
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
        ),
        "word/document.xml": strToU8(`<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}</w:body></w:document>`),
      },
      { level: 6 },
    ),
  );

const p = (inner: string): string => `<w:p>${inner}</w:p>`;
const t = (s: string, preserve = false): string =>
  `<w:r><w:t${preserve ? ' xml:space="preserve"' : ""}>${s}</w:t></w:r>`;
/** 一个格子：孩子直接放进 `w:tc`——嵌套表是段落的兄弟，不是段落的孙子（写进 `w:p` 里就成了非法形状）。 */
const cell = (children: string): string => `<w:tc>${children}</w:tc>`;
const tcell = (s: string): string => cell(p(t(s)));
const table = (cells: string): string => `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cells}</w:tr></w:tbl>`;

/**
 * 「字数统计」（WPS 的 aid 792 `idMso="WordCount"`，`dialogHint="1"`）数的是模型里的文字，不是页面上的像素。
 * 四条口径都在这一个文件里钉死：汉字算「中文字符」而中文标点不算、空格只在「计空格」那一行算、制表符与换行
 * 一个字都不算、表格（含格子里的格子）里的段落与正文同一段逻辑。
 *
 * 每一行的期望数都是手算的：`中文 two words` 是 12 个码点（含两个空格）、10 个非空、2 个汉字、2 个非中文单词。
 */
describe("docx 字数统计：页数之外那几行由模型数出来", () => {
  it("把汉字、空白与标点各归各的一行", () => {
    const doc = docOf(
      p(t("中文 two words")) +
        p(t("全角空格　与，标点", true)) +
        p(t("格子")) +
        table(tcell("格子里")) +
        table(cell(p(t("前") + '<w:r><w:tab/><w:br/><w:t>后</w:t></w:r>'))),
    );
    // 段落：正文 3 + 表里那格 1 + 最后一张表那格 1 = 5（空格与制表符不成段）
    expect(doc.wordCounts().paragraphs).toBe(5);
    // 码点：12 + 9 + 2（格子）+ 3（格子里）+ 2（前后）= 28
    expect(doc.wordCounts().characters).toBe(12 + 9 + 2 + 3 + 2);
    // 不计空格：10 + 8（全角空格 U+3000 剔掉）+ 2 + 3 + 2 = 25
    expect(doc.wordCounts().charactersNoSpace).toBe(10 + 8 + 2 + 3 + 2);
    // 汉字：中文 2 + 那句 7（，不算）+ 格子 2 + 格子里 3 + 前后 2 = 16
    expect(doc.wordCounts().cjk).toBe(2 + 7 + 2 + 3 + 2);
    // 非中文单词：只有 two、words
    expect(doc.wordCounts().words).toBe(2);
  });

  it("中文标点不算汉字，混排的词被汉字切开", () => {
    const doc = docOf(p(t("报ABC中DEF", true)));
    // 「报/中」两个汉字，「ABC」「DEF」两个词——汉字当分隔符，所以是 2 词而不是 1 词；八个码点没有空格
    expect(doc.wordCounts()).toMatchObject({ cjk: 2, words: 2, characters: 8, charactersNoSpace: 8 });
  });

  it("空格、制表符与换行都不算字，空段落仍是一个段落", () => {
    const doc = docOf(p('<w:r><w:tab/></w:r><w:r><w:br w:type="page"/></w:r>') + p(t("   ")));
    expect(doc.wordCounts()).toMatchObject({ paragraphs: 2, characters: 3, charactersNoSpace: 0, cjk: 0, words: 0 });
  });

  it("格子里再套一张表，那些字也数得到", () => {
    // 嵌套表是格子里那一段的**兄弟**：`w:tc` 里先是 `w:p`，再是 `w:tbl`
    const doc = docOf(table(cell(p(t("外层")) + table(tcell("嵌套")))));
    expect(doc.wordCounts()).toMatchObject({ paragraphs: 2, characters: 4, charactersNoSpace: 4, cjk: 4, words: 0 });
  });

  it("新建的空白件是一段空文字", () => {
    expect(WordDocument.blank().wordCounts()).toMatchObject({ paragraphs: 1, characters: 0, charactersNoSpace: 0, cjk: 0, words: 0 });
  });
});
