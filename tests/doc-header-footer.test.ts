import { describe, expect, it } from "vitest";
import { WordDocument, twipsToPx } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import {
  BAND_DEFAULT_TEXT,
  BAND_EVEN_TEXT,
  BAND_FIRST_TEXT,
  FIRST_SECTION_HEADER,
  FIRST_SECTION_FOOTER,
  LAST_SECTION_HEADER,
  LAST_SECTION_FOOTER,
  PAGE_FOOTER,
  PAGE_ONLY_FOOTER,
  PAGE_ONLY_FOOTER_STYLED,
  PAGE_NO_RESULT_FOOTER,
  footerPart,
  LOGO_RUN,
  PAGE_BOX_DISAGREEING_RUN,
  PAGE_BOX_RUN,
  PAGE_BOX_TWO_PARA_RUN,
  makeBoxWithWordsFooterDocx,
  makeFooterDocx,
  makeFirstPageHeaderDocx,
  makeLogoFooterDocx,
  makeHeaderFooterDocx,
  makeFullPageDistanceDocx,
  makePageDistanceDocx,
  makeStyledHeaderDocx,
  makeTwoSectionDocx,
  pageFieldRuns,
  pageFieldWithoutResult,
  REFERENCE_EVEN_ONLY,
  REFERENCE_FIRST_PAGE_ONLY,
  REFERENCE_ID_FIRST,
  REFERENCE_TYPE_FIRST,
  STYLED_FIRST_PARA,
  STYLED_HEADER,
  XML_DECL,
} from "./helpers/make-header-footer.js";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const headerOf = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/header1.xml")!);
const footerOf = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/footer1.xml")!);

/** A footer of one plain paragraph — the shape a page number sits in. */
const plainFooter = (content: string): string => `${XML_DECL}<w:ftr xmlns:w="${NS_W}"><w:p><w:r><w:t>${content}</w:t></w:r></w:p></w:ftr>`;

/** A part of the caller's markup, in a file whose body points at it. */
const withHeader = (body: string): Uint8Array => makeHeaderFooterDocx(`${XML_DECL}<w:hdr xmlns:w="${NS_W}">${body}</w:hdr>`);

/**
 * The header band's own `style`, which is the only place its box can come from: `renderWord` emits the band
 * as a sibling of `.page`, so the variables set on the page never reach it and the stylesheet answers with
 * its portrait 794px fallback — a landscape opening section was drawing a portrait header.
 */
const bandStyle = (bytes: Uint8Array): string => /class="hf hf-header[^"]*"[^>]*?style="([^"]*)"/.exec(renderWord(WordDocument.openBytes(bytes)))?.[1] ?? "";

/** Open, set the header's texts, save. */
const edited = (bytes: Uint8Array, paragraphs: string[]): Uint8Array => {
  const doc = WordDocument.openBytes(bytes);
  expect(doc.setHeaderTexts("header", paragraphs)).toBe(true);
  return writeDocx(doc);
};

/** The same part again as a file, so a save's output can be read back. */
const reread = (headerPart: string): WordDocument => WordDocument.openBytes(makeHeaderFooterDocx(headerPart));

describe("docx header/footer parts travel as the file wrote them", () => {
  it("reads each paragraph's text, with a literal tab kept as a tab", () => {
    const doc = WordDocument.openBytes(makeStyledHeaderDocx());
    expect(doc.header?.editable).toBe(true);
    // One tab, from the run. The `w:tab` inside `w:pPr/w:tabs` says where a tab lands; it is not text.
    expect(doc.header?.paragraphs).toEqual(["机密\t第一段", "第二段"]);
  });

  it("saves a part nobody typed in byte for byte", () => {
    expect(headerOf(writeDocx(WordDocument.openBytes(makeStyledHeaderDocx())))).toBe(STYLED_HEADER);
  });

  it("keeps an untouched paragraph, and the declarations its ids need, when another one is edited", () => {
    const out = headerOf(edited(makeStyledHeaderDocx(), ["机密\t第一段", "第二段改"]));
    expect(out).toContain(STYLED_FIRST_PARA);
    expect(out).toContain('mc:Ignorable="w14"');
    expect(out).toContain("xmlns:w14=");
    expect(out).toContain("第二段改");
    expect(reread(out).header?.paragraphs).toEqual(["机密\t第一段", "第二段改"]);
  });

  it("keeps the edited paragraph's own properties, its run format, and its tab", () => {
    const out = headerOf(edited(makeStyledHeaderDocx(), ["机密\t新字", "第二段"]));
    expect(out).toContain('<w:p w14:paraId="11111111" w14:textId="22222222">');
    expect(out).toContain('<w:pStyle w:val="Header"/>');
    expect(out).toContain('<w:jc w:val="right"/>');
    expect(out).toContain('<w:tab w:val="right" w:leader="dot" w:pos="9026"/>');
    expect(out).toContain('<w:rPr><w:b/><w:sz w:val="18"/></w:rPr>');
    expect(out).toContain('<w:bookmarkStart w:id="7" w:name="HeaderMark"/>');
    expect(out).toContain("<w:t>机密</w:t><w:tab/><w:t>新字</w:t>");
    expect(out).not.toContain("第一段");
    expect(reread(out).header?.paragraphs).toEqual(["机密\t新字", "第二段"]);
  });

  it("writes xml:space only where the text needs it", () => {
    const out = headerOf(edited(makeStyledHeaderDocx(), ["机密\t第一段", " 两头 "]));
    // The paragraph that travelled by itself has none; the one that was rewritten has exactly its own.
    expect(out.split('xml:space="preserve"').length - 1).toBe(1);
    expect(out).toContain('<w:t xml:space="preserve"> 两头 </w:t>');
    expect(reread(out).header?.paragraphs).toEqual(["机密\t第一段", " 两头 "]);
  });

  it("appends a paragraph the part did not have, and drops one the model no longer holds", () => {
    const grown = headerOf(edited(makeStyledHeaderDocx(), ["机密\t第一段", "第二段", "第三段"]));
    expect(grown).toContain(STYLED_FIRST_PARA);
    expect(grown).toContain('<w:p w14:paraId="33333333">');
    expect(grown).toContain("<w:p><w:r><w:t>第三段</w:t></w:r></w:p>");
    expect(reread(grown).header?.paragraphs).toEqual(["机密\t第一段", "第二段", "第三段"]);

    const shrunk = headerOf(edited(makeStyledHeaderDocx(), ["机密\t第一段"]));
    expect(shrunk).toContain(STYLED_FIRST_PARA);
    expect(shrunk).not.toContain("第二段");
    expect(shrunk).not.toContain('w14:paraId="33333333"');
  });

  it("refuses to flatten a part whose structure a flat text cannot say", () => {
    const cases: [string, string, string[]][] = [
      ["软换行", `<w:hdr xmlns:w="${NS_W}"><w:p><w:r><w:t>上</w:t><w:br/><w:t>下</w:t></w:r></w:p></w:hdr>`, ["上下"]],
      ["插入修订", `<w:hdr xmlns:w="${NS_W}"><w:p><w:ins w:id="1" w:author="甲"><w:r><w:t>加的字</w:t></w:r></w:ins></w:p></w:hdr>`, ["加的字"]],
      // Refusing the write is not refusing the words: a container this model will not edit still shows them
      // on the page, and its bytes still travel whole.
      ["容器", `<w:hdr xmlns:w="${NS_W}"><w:p><w:r><w:t>字</w:t></w:r></w:p><w:sdt><w:sdtContent><w:p><w:r><w:t>控件</w:t></w:r></w:p></w:sdtContent></w:sdt></w:hdr>`, ["字", "控件"]],
      // The old gate read only `w:`, so a part that binds the namespace under another name slipped past it
      // and had its container rebuilt away.
      ["别的前缀下的容器", `<w0:hdr xmlns:w0="${NS_W}"><w0:p><w0:r><w0:t>字</w0:t></w0:r></w0:p><w0:sdt><w0:sdtContent><w0:p><w0:r><w0:t>控件</w0:t></w0:r></w0:p></w0:sdtContent></w0:sdt></w0:hdr>`, ["字", "控件"]],
    ];
    for (const [name, inner, texts] of cases) {
      const part = `${XML_DECL}${inner}`;
      const doc = WordDocument.openBytes(makeHeaderFooterDocx(part));
      expect(doc.header?.editable, name).toBe(false);
      expect(doc.header?.paragraphs, name).toEqual(texts);
      expect(headerOf(writeDocx(doc)), name).toBe(part);
    }
  });

  it("keeps a part with nothing but paragraphs editable, whatever prefix it binds", () => {
    const part = `${XML_DECL}<w0:hdr xmlns:w0="${NS_W}"><w0:p><w0:r><w0:t>前缀页眉</w0:t></w0:r></w0:p></w0:hdr>`;
    const doc = WordDocument.openBytes(makeHeaderFooterDocx(part));
    expect(doc.header?.editable).toBe(true);
    expect(doc.header?.paragraphs).toEqual(["前缀页眉"]);
    expect(headerOf(edited(makeHeaderFooterDocx(part), ["改过的页眉"]))).toBe(
      `${XML_DECL}<w0:hdr xmlns:w0="${NS_W}"><w0:p><w0:r><w0:t>改过的页眉</w0:t></w0:r></w0:p></w0:hdr>`,
    );
  });

  it("spells a header this session created without inventing attributes", () => {
    const doc = WordDocument.blank();
    doc.ensureHeaderFooter("header");
    doc.setHeaderTexts("header", ["第 1 页"]);
    const out = headerOf(writeDocx(doc));
    expect(out).toContain("<w:t>第 1 页</w:t>");
    expect(out).not.toContain("xml:space");
    expect(WordDocument.openBytes(writeDocx(doc)).header?.paragraphs).toEqual(["第 1 页"]);
  });

  /**
   * Which part a section uses as its header is said by two attributes, and the order they appear in is the
   * file's choice. Matching them by position is what made every header in the local corpus invisible: the
   * search this replaced required `w:type` before `r:id`, where WPS writes `r:id` before `w:type` on all 59
   * header and 57 footer references. A real file therefore reported no header at all — nothing to display,
   * nothing to edit, and no part for a save to splice back. The first-page reference is the same question with
   * the switch on, so it travels here too.
   */
  it("finds the part whatever order the reference puts its attributes in", () => {
    for (const [ref, titlePg] of [[REFERENCE_ID_FIRST, false], [REFERENCE_TYPE_FIRST, false], [REFERENCE_FIRST_PAGE_ONLY, true]] as [string, boolean][]) {
      const doc = WordDocument.openBytes(makeHeaderFooterDocx(STYLED_HEADER, ref, titlePg));
      expect(doc.header?.id, ref).toBe("rId5");
      expect(doc.header?.paragraphs, ref).toEqual(["机密\t第一段", "第二段"]);
    }
  });

  /**
   * 「首页不同」 is the section's own word for which of its bands the opening page wears (§17.10.3): a
   * `w:type="first"` reference is the page's band when `w:titlePg` stands, and not otherwise. The choice cannot
   * be made by whichever reference the section wrote first, and the authorized 46 份 show why the question is real
   * rather than hypothetical (`.scratch/even-odd-audience.py`: 58 节, 38 枚带引用 — `footer:default` 26、
   * `header:default` 6、`footer:even` 4、`footer:first` 2, no `header:even` at all). `w:titlePg` is stated on 2 节,
   * and the 4 节 that hold an even band are not among them: an even band is a part held, never a part shown,
   * because `w:evenAndOddHeaders` is stated **0 枚** across 46/46 `word/settings.xml` — the switch is the file's
   * and this app has no page that asks for it (#164).
   */
  it("wears the first-page band only when the section says the first page differs", () => {
    const off = WordDocument.openBytes(makeFirstPageHeaderDocx());
    expect(off.header?.paragraphs, "no titlePg: the default band").toEqual([BAND_DEFAULT_TEXT]);
    expect(off.header?.part, "no titlePg: the default part").toBe("word/header1.xml");

    const on = WordDocument.openBytes(makeFirstPageHeaderDocx({ titlePg: true }));
    expect(on.header?.paragraphs, "titlePg: the first band").toEqual([BAND_FIRST_TEXT]);
    expect(on.header?.part, "titlePg: the first part").toBe("word/header2.xml");

    // Onto the page, not just into the model: the paper the user looks at wears one of the two, and the other
    // says nothing on it.
    expect(renderWord(off)).toContain(BAND_DEFAULT_TEXT);
    expect(renderWord(off)).not.toContain(BAND_FIRST_TEXT);
    expect(renderWord(on)).toContain(BAND_FIRST_TEXT);
    expect(renderWord(on)).not.toContain(BAND_DEFAULT_TEXT);

    // The even band is held but displayed by no section: 授权 46 份 cite 的 38 枚引用里 `w:type="even"` 只有
    // **4 枚，全在页脚那一侧**，而 **46/46 的 `word/settings.xml` 都没说 `evenAndOddHeaders`**（0 枚）——开关是
    // 文件的，不是渲染器的（#164）。所以一份只 cite even 的件有部件，但这一张纸不是它归属的那张。
    const even = WordDocument.openBytes(makeHeaderFooterDocx(STYLED_HEADER, REFERENCE_EVEN_ONLY));
    expect(even.header).toBeUndefined();
  });

  it("switches to the first-page band without inventing a reference or touching the part it did not show", () => {
    // Showing a different band is a read decision, so the save owes the file its own set of three references
    // and the untouched bytes of the two parts the page did not wear. Editing the one it did wear must still
    // land in that part — the model holds `header2.xml`, and a save that spliced `header1.xml` would write the
    // text onto pages this file says are not first.
    const bytes = makeFirstPageHeaderDocx({ titlePg: true });
    const doc = WordDocument.openBytes(bytes);
    expect(doc.setHeaderTexts("header", ["首页改版"])).toBe(true);
    const out = OpcPackage.from(writeDocx(doc));
    const xml = text(out.bytes("word/document.xml")!);
    expect((xml.match(/<w:headerReference/g) ?? []).length, "three references, none invented").toBe(3);
    expect(xml).toContain('<w:headerReference r:id="rId6" w:type="first"/>');
    expect(xml).toContain("<w:titlePg/>");
    expect(text(out.bytes("word/header2.xml")!)).toContain("<w:t>首页改版</w:t>");
    expect(text(out.bytes("word/header1.xml")!)).toContain(`<w:t>${BAND_DEFAULT_TEXT}</w:t>`);
    expect(text(out.bytes("word/header3.xml")!)).toContain(`<w:t>${BAND_EVEN_TEXT}</w:t>`);
    expect(WordDocument.openBytes(writeDocx(doc)).header?.paragraphs).toEqual(["首页改版"]);
  });

  it("keeps the part the section states for a switched-off first page, and shows the default it holds", () => {
    // A section that asks for a different first page but cites no first-page band: the file still holds a
    // default band, and blanking the page here would take away text the file wrote. What WPS puts on that page
    // is 未取证 — this is the choice that keeps the file's own content on it.
    const doc = WordDocument.openBytes(makeFirstPageHeaderDocx({ titlePg: true, refs: ["default", "even"] }));
    expect(doc.header?.paragraphs).toEqual([BAND_DEFAULT_TEXT]);
    expect(doc.header?.part).toBe("word/header1.xml");
  });

  it("reports no header when no section binds one, even with the part sitting in the package", () => {
    const doc = WordDocument.openBytes(makeHeaderFooterDocx(STYLED_HEADER, ""));
    expect(doc.header).toBeUndefined();
  });

  it("round-trips an ordinary header edit", () => {
    const bytes = withHeader(`<w:p><w:r><w:t>公司机密</w:t></w:r></w:p>`);
    expect(WordDocument.openBytes(bytes).header?.paragraphs).toEqual(["公司机密"]);
    expect(WordDocument.openBytes(edited(bytes, ["已修订"])).header?.paragraphs).toEqual(["已修订"]);
  });

  it("edits a footer as the part it is, not as a header", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(plainFooter("第 1 页")));
    expect(doc.header).toBeUndefined();
    expect(doc.footer?.editable).toBe(true);
    expect(doc.footer?.paragraphs).toEqual(["第 1 页"]);
    expect(doc.setHeaderTexts("footer", ["第 2 页"])).toBe(true);

    const out = writeDocx(doc);
    expect(footerOf(out)).toContain("<w:t>第 2 页</w:t>");
    expect(WordDocument.openBytes(out).footer?.paragraphs).toEqual(["第 2 页"]);
    expect(WordDocument.openBytes(out).header).toBeUndefined();
  });

  it("saves an untouched footer byte for byte", () => {
    const part = plainFooter("机密 · 请勿外传");
    expect(footerOf(writeDocx(WordDocument.openBytes(makeFooterDocx(part))))).toBe(part);
  });
});

describe("docx 页眉页脚：显示", () => {
  it("draws one line per paragraph on the paper, with the tab kept", () => {
    const html = renderWord(WordDocument.openBytes(makeStyledHeaderDocx()));
    // 这一段自己写了 `w:jc w:val="right"`，纸面就得靠右（#211：此前带子里清一色左对齐，段说的话没人问）。
    // 第二段借的是 `Footer` 样式，那份 `word/styles.xml` 里没写 `w:jc`，所以一个属性都不多。
    expect(html).toContain('<div class="hf-p" style="text-align:right">机密\t第一段</div>');
    expect(html).toContain('<div class="hf-p">第二段</div>');
    // 带子排在纸盒**之内**、正文块之后（#157：此前它排在纸盒之外，画在纸上看不到，正文也从没为它让位）。
    // 排在最后一条与页面边框同一先例：编辑器按块在 `.page` 孩子里的下标寻址，任何排在前面的装饰会把下标顶偏。
    expect(html.indexOf('class="page"')).toBeLessThan(html.indexOf('solo-hf-master'));
    expect(html.indexOf('<div class="para"')).toBeLessThan(html.indexOf('solo-hf-master'));
  });

  it("puts the footer on the last sheet and names the band it offers to edit", () => {
    const html = renderWord(WordDocument.openBytes(makeFooterDocx(plainFooter("第 1 页"))));
    expect(html).toContain('data-hf="footer"');
    expect(html).toContain('class="hf hf-footer hf-bottom solo-hf-master"');
    expect(html).toContain('title="点击编辑页脚"');
    expect(html).toContain('<div class="hf-p">第 1 页</div>');
    expect(html.indexOf('class="page"')).toBeLessThan(html.indexOf("hf hf-footer"));
    expect(renderWord(WordDocument.openBytes(makeStyledHeaderDocx()))).toContain('class="hf hf-header solo-hf-master"');
  });

  it("keeps the band on the paper uneditable until it is opened", () => {
    // 带子住在可编辑的纸盒里，`contenteditable="false"` 就是它 resting 状态唯一的来源（挂在纸盒外时那句是继承来的）。
    // 缺了它，`editHF` 开头那句 `isContentEditable` 直接 return：带子点不开，打字落进正文，读回又把它当装饰跳过。
    const html = renderWord(WordDocument.openBytes(makeStyledHeaderDocx()));
    const master = /<div class="hf hf-header[^"]*"[^>]*>/.exec(html)?.[0] ?? "";
    expect(master).toContain('contenteditable="false"');
    expect(master).toContain('data-hf-part="word/header1.xml"');
  });

  it("offers no band for a section that binds nothing", () => {
    const html = renderWord(WordDocument.openBytes(makeHeaderFooterDocx(STYLED_HEADER, "")));
    expect(html).not.toContain('class="hf ');
  });

  it("says a rich band is not editable, in the words the reader sees", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(`${XML_DECL}<w:ftr xmlns:w="${NS_W}"><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p></w:ftr>`));
    expect(doc.footer?.editable).toBe(false);
    expect(renderWord(doc)).toContain("此页脚含表格或图片等富内容，暂不支持编辑");
  });
});

/**
 * A header belongs to a section, and 18 of the 86 files this machine's WPS wrote have more than one: the
 * break hanging off a paragraph ends the opening section, and the body's own `sectPr` is the closing one.
 * Reading only the body's read the last section, which in those files reports a header 4 times and a footer
 * 5 times that the opening page does not have, and points at a different part once each.
 */
describe("docx 页眉页脚：多节文档按节取", () => {
  /** The document.xml text between two markers, so a test can name the section that holds something. */
  const between = (xml: string, from: string, to: string): string => {
    const a = xml.indexOf(from);
    const b = xml.indexOf(to);
    expect(a, xml).toBeGreaterThanOrEqual(0);
    return xml.slice(a, b < 0 ? xml.length : b);
  };
  const docXml = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);

  it("shows the header of the section the document opens with, not the last one", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx());
    expect(doc.header?.part).toBe("word/header1.xml");
    expect(doc.header?.paragraphs).toEqual(["开篇页眉"]);
  });

  it("edits the part the opening section cites and leaves the closing section's alone", () => {
    const bytes = makeTwoSectionDocx();
    const out = edited(bytes, ["改过的开篇"]);
    expect(text(OpcPackage.from(out).bytes("word/header1.xml")!)).toContain("改过的开篇");
    expect(text(OpcPackage.from(out).bytes("word/header2.xml")!)).toBe(LAST_SECTION_HEADER);
    expect(docXml(out)).toBe(docXml(bytes));
  });

  it("offers no band when the opening section says nothing about a header", () => {
    const bytes = makeTwoSectionDocx({ firstRef: "" });
    expect(WordDocument.openBytes(bytes).header).toBeUndefined();
    expect(renderWord(WordDocument.openBytes(bytes))).not.toContain('class="hf ');
    // The closing section's header is the file's own and stays that way: nothing to show, nothing to touch.
    expect(text(OpcPackage.from(writeDocx(WordDocument.openBytes(bytes))).bytes("word/header2.xml")!)).toBe(LAST_SECTION_HEADER);
    expect(docXml(writeDocx(WordDocument.openBytes(bytes)))).toBe(docXml(bytes));
  });

  it("lists every section the file declared, in the order they start", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx());
    expect(doc.sections.map((s) => [s.pageWidthTwips, s.pageHeightTwips])).toEqual([
      [11906, 16838],
      [16838, 11906],
    ]);
  });

  /**
   * #157：纸上那一条带子的**字**得跟着节走。这一节自己 cite 的用这一节的部件，什么都没说的一节接着它前面
   * 最近的那一节——不是接着整份件的最后一份，也不是把开篇那一节的字贴到每一张纸上。
   */
  it("keeps what each section cites itself, and reads the closing section's part too", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx());
    expect(doc.sectionRefs.map((p) => [p.header?.part, p.footer?.part])).toEqual([
      ["word/header1.xml", undefined],
      ["word/header2.xml", undefined],
    ]);
    expect(doc.sectionBand("header", 0)?.paragraphs).toEqual(["开篇页眉"]);
    expect(doc.sectionBand("header", 1)?.paragraphs).toEqual(["结尾页眉"]);
    // The band the editor draws and a save writes comes from the section the page starts in; a later section's
    // part is a separate one, and the save writes only what somebody actually typed into.
    expect(doc.header?.part).toBe("word/header1.xml");
    expect(doc.sectionBand("header", 1)).not.toBe(doc.header);
  });

  it("carries a section that says nothing off the one before it, never off the one after", () => {
    const quiet = WordDocument.openBytes(makeTwoSectionDocx({ lastRef: "" }));
    expect(quiet.sectionRefs[1]).toEqual({});
    expect(quiet.sectionBand("header", 1)?.part).toBe("word/header1.xml");
    // Forward only: the opening section shows nothing, and the closing one's part does not reach back to it.
    const firstQuiet = WordDocument.openBytes(makeTwoSectionDocx({ firstRef: "" }));
    expect(firstQuiet.sectionBand("header", 0)).toBeUndefined();
    expect(firstQuiet.sectionBand("header", 1)?.part).toBe("word/header2.xml");
  });

  it("hands one part to every section that cites it, so editing it is one edit", () => {
    const shared = makeTwoSectionDocx({ lastRef: '<w:headerReference r:id="rId5" w:type="default"/>' });
    const doc = WordDocument.openBytes(shared);
    expect(doc.sectionBand("header", 1)).toBe(doc.header);
  });

  it("draws the later section's band as a母本 the page can copy, not as a second editable band", () => {
    const html = renderWord(WordDocument.openBytes(makeTwoSectionDocx()));
    expect(html).toContain('class="solo-hf-src" data-hf-part="word/header2.xml" data-hf-kind="header"');
    expect(html).toContain("结尾页眉");
    // 母本不占 `.hf` 这个名：`wireHF` 只认 `.hf`，认到两份就是两份都能改、改了各归各。
    expect(/class="[^"]*solo-hf-src[^"]*hf-header/.test(html)).toBe(false);
    expect((html.match(/class="hf hf-header[^"]*"/g) ?? []).length).toBe(1);
    expect(html).toContain("display:none");
    // A one-section file draws none of it.
    expect(renderWord(WordDocument.openBytes(makeStyledHeaderDocx()))).not.toContain("solo-hf-src");
  });

  it("puts the closing section's footer on the last sheet and keeps the opening one as a母本", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx({ footers: "both" }));
    const html = renderWord(doc);
    expect(doc.sectionBand("footer", 1)?.part).toBe("word/footer2.xml");
    const partOf = (which: "header" | "footer"): string =>
      new RegExp(`class="hf hf-${which}[^"]*"[^>]*?data-hf-part="([^"]*)"`).exec(html)?.[1] ?? "";
    expect(partOf("footer"), "纸盒外那一份页脚守的是最后一张，不是开篇那一张").toBe("word/footer2.xml");
    expect(partOf("header")).toBe("word/header1.xml");
    // 第一节那一份不再可见，就退成母本给纸上的复件抄——可见与母本合起来仍是一份部件只画一遍。
    expect(html).toContain('class="solo-hf-src" data-hf-part="word/footer1.xml" data-hf-kind="footer"');
    expect((html.match(/class="hf hf-footer/g) ?? []).length).toBe(1);
  });

  it("writes back the part the last sheet shows, not the one the document opens with", () => {
    const bytes = makeTwoSectionDocx({ footers: "both" });
    const doc = WordDocument.openBytes(bytes);
    const part = doc.sectionBand("footer", 1)!.part;
    expect(doc.setBandSegments("footer", part, [[{ text: "改过的结尾" }]])).toBe(true);
    const out = writeDocx(doc);
    const pkg = OpcPackage.from(out);
    expect(text(pkg.bytes("word/footer2.xml")!)).toContain("改过的结尾");
    // 没改过的三份各归各：开篇那一张的页脚、两张的页眉，还有正文，都按原字节交回去。
    expect(text(pkg.bytes("word/footer1.xml")!)).toBe(FIRST_SECTION_FOOTER);
    expect(text(pkg.bytes("word/header1.xml")!)).toBe(FIRST_SECTION_HEADER);
    expect(text(pkg.bytes("word/document.xml")!)).toBe(text(OpcPackage.from(bytes).bytes("word/document.xml")!));
    expect(WordDocument.openBytes(out).sectionBand("footer", 1)?.paragraphs).toEqual(["改过的结尾"]);
  });

  it("lists every part a section shows to the save, and marks clean ones clean", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx({ footers: "both" }));
    expect(doc.headerFooterParts().map((p) => p.part).sort()).toEqual(["word/footer1.xml", "word/footer2.xml", "word/header1.xml", "word/header2.xml"]);
    expect(doc.headerFooterParts().every((p) => !p.dirty), "nobody typed: no part is a write").toBe(true);
    expect(doc.setBandSegments("header", "word/header2.xml", [[{ text: "结尾改版" }]])).toBe(true);
    expect(doc.headerFooterParts().filter((p) => p.dirty).map((p) => p.part)).toEqual(["word/header2.xml"]);
  });

  it("shows one footer on both sheets when the closing section cites nothing", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx({ footers: "first" }));
    expect(doc.sectionBand("footer", 1)?.part).toBe("word/footer1.xml");
    const html = renderWord(doc);
    expect(/class="hf hf-footer[^"]*"[^>]*?data-hf-part="word\/footer1.xml"/.test(html)).toBe(true);
    // 同一份部件不另画母本：两份都叫 `.hf` 就两份都能改、改了各归各。
    expect(html).not.toContain('data-hf-part="word/footer1.xml" data-hf-kind="footer"');
    expect(doc.headerFooterParts().map((p) => p.part).sort()).toEqual(["word/footer1.xml", "word/header1.xml", "word/header2.xml"]);
  });

  /**
   * #163 第三档：授权口径里 4 个多节模板身份有 **2 个的第一节什么都不 cite、后面的节各 cite 一枚页脚**
   * （labor contract 是 `[沉默, footer:default]`，project plan 是 `[沉默, footer:default, footer:default]`；
   * `.scratch/band-per-section.ts` 读的就是这 8 个实例）。这一支把那个形状钉在单元测试里：第一张纸没有页脚，
   * 末张显示**第二节自己**那一份，向前继承不回头（§17.10.3）。
   */
  it("shows no footer on the first sheet when only a later section cites one", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx({ footers: "last" }));
    expect(doc.footer, "开篇那一节没 cite，文档级页眉页脚就不该替它认领一份").toBeUndefined();
    expect(doc.sectionRefs.map((p) => [p.header?.part ?? "-", p.footer?.part ?? "-"])).toEqual([
      ["word/header1.xml", "-"],
      ["word/header2.xml", "word/footer2.xml"],
    ]);
    expect(doc.sectionBand("footer", 0), "第二节那一份不回头够到第一张纸").toBeUndefined();
    expect(doc.sectionBand("footer", 1)?.paragraphs).toEqual(["结尾页脚"]);
    const html = renderWord(doc);
    const partOf = (which: "header" | "footer"): string =>
      new RegExp(`class="hf hf-${which}[^"]*"[^>]*?data-hf-part="([^"]*)"`).exec(html)?.[1] ?? "";
    expect(partOf("footer")).toBe("word/footer2.xml");
    expect(partOf("header")).toBe("word/header1.xml");
    // 第一节没有页脚，也没有第二份页脚可降级：页脚这一侧纸上只有一份、母本零份。
    expect((html.match(/class="hf hf-footer/g) ?? []).length).toBe(1);
    expect(html).not.toContain('data-hf-part="word/footer2.xml" data-hf-kind="footer"');
    expect(doc.headerFooterParts().map((p) => p.part).sort()).toEqual(["word/footer2.xml", "word/header1.xml", "word/header2.xml"]);
  });

  /** 同一形状的写回侧：#163 在 8 个授权实例上量到「逐节 cite 表一致 / 纸面显示一致」，这里钉住这一支。 */
  it("saves a file whose opening section cites no footer with that section still silent", () => {
    const bytes = makeTwoSectionDocx({ footers: "last" });
    const out = writeDocx(WordDocument.openBytes(bytes));
    expect(text(OpcPackage.from(out).bytes("word/document.xml")!)).toBe(text(OpcPackage.from(bytes).bytes("word/document.xml")!));
    expect(text(OpcPackage.from(out).bytes("word/footer2.xml")!)).toBe(LAST_SECTION_FOOTER);
    const again = WordDocument.openBytes(out);
    expect(again.footer).toBeUndefined();
    expect(again.sectionRefs.map((p) => [p.header?.part ?? "-", p.footer?.part ?? "-"])).toEqual([
      ["word/header1.xml", "-"],
      ["word/header2.xml", "word/footer2.xml"],
    ]);
    expect([again.sectionBand("footer", 0)?.part ?? "-", again.sectionBand("footer", 1)?.part ?? "-"]).toEqual(["-", "word/footer2.xml"]);
  });

  /**
   * 新建那一枚引用**怎么写**也归 #163 的口径管：授权 46 份的 38 枚引用 38/38 是 `r:id` 在前、`w:type` 在后、
   * 自闭合；cite 多枚的 6 节按 `w:type` 分档排（`first` → `default` → `even`），同档里页眉在页脚前
   * （`.scratch/sect-ref-order.ts`）。这一支把两端钉住——落点归开篇那一节（`bandSectIsBody` 为假时走
   * `citeBandReference`），形状按语料拼。
   */
  it("writes a band reference this session owes in the order the files use", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx({ footers: "last" }));
    const made = doc.ensureHeaderFooter("footer");
    expect(made?.part).toBe("word/footer1.xml");
    const xml = docXml(writeDocx(doc));
    // 排在第一节那枚 `header:default` **之后**，不是插到 `sectPr` 最前面。
    expect(between(xml, 'w:rsidR="00A1"', 'w:rsidR="00B2"')).toContain('<w:headerReference r:id="rId5" w:type="default"/><w:footerReference r:id="rId9" w:type="default"/>');
    expect(between(xml, 'w:rsidR="00A1"', 'w:rsidR="00B2"')).not.toContain("w:type=\"default\" r:id=");
    // 第二节自己那一份一个字没动。
    expect(xml.slice(xml.indexOf('w:rsidR="00B2"'))).toContain('<w:footerReference r:id="rId8" w:type="default"/>');
    const again = WordDocument.openBytes(writeDocx(doc));
    expect([again.sectionBand("footer", 0)?.part ?? "-", again.sectionBand("footer", 1)?.part ?? "-"]).toEqual(["word/footer1.xml", "word/footer2.xml"]);
  });

  it("puts a created default band after the first-page reference the section already wrote", () => {
    // 授权那 2 节把 `w:type="first"` 写在 `default` 之前，一节连页脚的 first 都排在页眉的 default 前。
    const doc = WordDocument.openBytes(makeFirstPageHeaderDocx({ titlePg: true, refs: ["first"] }));
    expect(doc.header?.part).toBe("word/header2.xml");
    expect(doc.footer).toBeUndefined();
    doc.ensureHeaderFooter("footer");
    const xml = docXml(writeDocx(doc));
    expect(xml).toMatch(/<w:headerReference r:id="rId6" w:type="first"\/><w:footerReference r:id="rId\d+" w:type="default"\/>/);
    const again = WordDocument.openBytes(writeDocx(doc));
    expect(again.header?.part).toBe("word/header2.xml");
    expect(again.footer?.part).toBe("word/footer1.xml");
  });

  it("saves a two-section file byte-identical when nobody typed, with both header parts intact", () => {
    const bytes = makeTwoSectionDocx();
    const out = writeDocx(WordDocument.openBytes(bytes));
    const pkg = OpcPackage.from(out);
    expect(text(pkg.bytes("word/header1.xml")!)).toBe(text(OpcPackage.from(bytes).bytes("word/header1.xml")!));
    expect(text(pkg.bytes("word/header2.xml")!)).toBe(LAST_SECTION_HEADER);
    expect(text(pkg.bytes("word/document.xml")!)).toBe(text(OpcPackage.from(bytes).bytes("word/document.xml")!));
  });

  it("gives the band the width of the section that owns it, both ways round", () => {
    // The page box and the band have to come from the same section, or a header belonging to a landscape
    // page gets drawn on a portrait one.
    expect(bandStyle(makeTwoSectionDocx())).toContain(`--page-w:${twipsToPx(11906)}px`);
    expect(bandStyle(makeTwoSectionDocx({ firstLandscape: true }))).toContain(`--page-w:${twipsToPx(16838)}px`);
    // And the same column the body's text starts in. The stylesheet's own padding is a fixed 96px = 1440
    // twips, which is the one margin value no file in the corpus uses.
    expect(bandStyle(makeTwoSectionDocx({ firstMarginTwips: 1800 }))).toContain(`--pad-l:${twipsToPx(1800)}px`);
    expect(bandStyle(makeTwoSectionDocx())).toContain("padding-left:var(--pad-l)");
  });

  it("creates a header the opening section cites, and never a second sectPr", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx({ firstRef: "" }));
    const made = doc.ensureHeaderFooter("header");
    expect(made?.part).toBe("word/header3.xml");
    const out = writeDocx(doc);
    const xml = docXml(out);
    expect((xml.match(/<w:sectPr\b/g) ?? []).length, "a save must not spell a section the file never had").toBe(2);
    expect(between(xml, 'w:rsidR="00A1"', 'w:rsidR="00B2"')).toContain(`r:id="${made!.id}"`);
    expect(xml.slice(xml.indexOf('w:rsidR="00B2"'))).not.toContain(made!.id);
    const again = WordDocument.openBytes(out);
    expect(again.header?.part).toBe("word/header3.xml");
    expect(again.header?.editable).toBe(true);
  });

  it("keeps a created band's reference in the same place on a second save", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx({ firstRef: "" }));
    doc.ensureHeaderFooter("header");
    const once = writeDocx(doc);
    const twice = writeDocx(WordDocument.openBytes(once));
    expect(docXml(twice)).toBe(docXml(once));
    expect(text(OpcPackage.from(twice).bytes("word/header3.xml")!)).toBe(text(OpcPackage.from(once).bytes("word/header3.xml")!));
  });
});


describe("docx page fields in a header or footer", () => {
  /** One paragraph of a fixture part, cut out whole, so a save can be compared against the bytes it kept. */
  const paragraphAt = (part: string, n: number): string => {
    const from = [...part.matchAll(/<w:p[\s>]/g)][n]!.index;
    return part.slice(from, part.indexOf("</w:p>", from) + "</w:p>".length);
  };

  it("reads a run-level PAGE field as one carried piece, not a reason to refuse the part", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(PAGE_FOOTER));
    expect(doc.footer?.editable).toBe(true);
    // The number the field cached is not text the reader typed, so it never joins the paragraph's text.
    expect(doc.footer?.paragraphs).toEqual(["第  页", "普通一行"]);
    expect(doc.footer?.segments?.[0]?.map((s) => (s.carry ? `[${s.carry.tag}:${s.carry.show}]` : s.text))).toEqual([
      "第 ",
      "[fldChar:页码]",
      " 页",
    ]);
    expect(doc.footer?.segments?.[0]?.[1]?.carry?.xml).toBe(pageFieldRuns());
    // The second paragraph is read the same way, as text with nothing carried in it.
    expect(doc.footer?.segments?.[1]).toEqual([{ text: "普通一行" }]);
  });

  it("saves a part that carries a field and was not typed in byte for byte", () => {
    for (const part of [PAGE_FOOTER, PAGE_ONLY_FOOTER, PAGE_ONLY_FOOTER_STYLED]) {
      expect(footerOf(writeDocx(WordDocument.openBytes(makeFooterDocx(part))))).toBe(part);
    }
  });

  /**
   * 授权语料那 **2 枚没有 `separate` 的域**（36 枚有；`.scratch/field-stage-shapes.ts`）。文件没写下「这一页是第
   * 几页」的缓存结果，读的人就一个数都不该编：域仍是整枚带着的记号，正文里既不出现「1」也不出现空位，写回按原字节。
   */
  it("carries a field that has no separate and no cached result without inventing a number", () => {
    const bytes = makeFooterDocx(PAGE_NO_RESULT_FOOTER);
    const doc = WordDocument.openBytes(bytes);
    expect(doc.footer?.editable).toBe(true);
    expect(doc.footer?.paragraphs).toEqual(["第  页"]);
    expect(doc.footer?.segments?.[0]?.map((s) => (s.carry ? `[${s.carry.tag}:${s.carry.show}]` : s.text))).toEqual(["第 ", "[fldChar:页码]", " 页"]);
    expect(doc.footer?.segments?.[0]?.[1]?.carry?.xml).toBe(pageFieldWithoutResult());
    const html = renderWord(doc);
    expect(html).toContain("第");
    expect(html).toContain("页");
    expect(html.match(/data-tag="fldChar"/g) ?? []).toHaveLength(1);
    expect(footerOf(writeDocx(doc))).toBe(PAGE_NO_RESULT_FOOTER);
    expect(WordDocument.openBytes(writeDocx(doc)).footer?.paragraphs).toEqual(["第  页"]);
  });

  it("leaves the field's paragraph as its bytes when another one is edited", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(PAGE_FOOTER));
    const fieldPara = doc.footer?.segments?.[0];
    expect(fieldPara && doc.setHeaderSegments("footer", [fieldPara, [{ text: "普通一行改" }]])).toBe(true);
    const out = footerOf(writeDocx(doc));
    expect(out).toContain(paragraphAt(PAGE_FOOTER, 0));
    expect(out).toContain("普通一行改");
  });

  it("gives the field back inside a paragraph the user retyped, where it stood", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(PAGE_FOOTER));
    const field = doc.footer?.segments?.[0]?.[1];
    expect(field && doc.setHeaderSegments("footer", [[{ text: "共 " }, field, { text: " 页完" }], [{ text: "普通一行" }]])).toBe(true);
    const out = footerOf(writeDocx(doc));
    expect(out).toContain(field!.carry!.xml);
    expect(out.indexOf("共 ")).toBeLessThan(out.indexOf('<w:fldChar w:fldCharType="begin"'));
    expect(out.indexOf('<w:fldChar w:fldCharType="end"')).toBeLessThan(out.indexOf(" 页完"));
    // The paragraph's own properties and the text run's typeface survive the rebuild.
    expect(out).toContain('<w:pStyle w:val="Footer"/>');
    expect(doc.footer?.paragraphs).toEqual(["共  页完", "普通一行"]);
  });

  it("sets retyped text in the paragraph's own run, not in the field's marker", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(PAGE_ONLY_FOOTER_STYLED));
    const field = doc.footer?.segments?.[0]?.[0];
    expect(field && doc.setHeaderSegments("footer", [[field, { text: "改过的字" }]])).toBe(true);
    expect(footerOf(writeDocx(doc))).toContain('<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>改过的字</w:t></w:r>');
  });

  it("refuses a flat text write on a part that carries a field", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(PAGE_FOOTER));
    // Flat text cannot say which side of the field it goes, and guessing would drop the field.
    expect(doc.setHeaderTexts("footer", ["第  页", "普通一行改"])).toBe(false);
    expect(doc.footer?.paragraphs).toEqual(["第  页", "普通一行"]);
  });

  it("refuses an instruction it cannot spell back, and keeps the part as its own bytes", () => {
    const part = footerPart(`<w:p><w:r><w:t>时间</w:t></w:r>${pageFieldRuns(" TIME ", "08:00")}</w:p>`);
    const doc = WordDocument.openBytes(makeFooterDocx(part));
    expect(doc.footer?.editable).toBe(false);
    expect(doc.footer?.segments).toBeUndefined();
    expect(footerOf(writeDocx(doc))).toBe(part);
  });

  it("refuses a field that never closes, and one nested in another", () => {
    const unclosed = footerPart(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r></w:p>`,
    );
    expect(WordDocument.openBytes(makeFooterDocx(unclosed)).footer?.editable).toBe(false);
    const nested = footerPart(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    );
    expect(WordDocument.openBytes(makeFooterDocx(nested)).footer?.editable).toBe(false);
  });

  /** The shape the corpus writes: properties naming the control, then the content it holds. */
  const control = (inner: string, pr = '<w:sdtPr><w:id w:val="77"/></w:sdtPr>'): string =>
    `<w:sdt>${pr}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

  it("reads a page number a content control wraps, and edits it inside the control", () => {
    const part = footerPart(
      `${control(`<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>${pageFieldRuns()}</w:p>`)}<w:p><w:r><w:t>联系文字</w:t></w:r></w:p>`,
    );
    const d = WordDocument.openBytes(makeFooterDocx(part));
    expect(d.footer?.editable).toBe(true);
    expect(d.footer?.paragraphs).toEqual(["", "联系文字"]);
    expect(d.footer?.segments?.[0]?.map((s) => (s.carry ? `[${s.carry.tag}:${s.carry.show}]` : s.text))).toEqual(["[fldChar:页码]"]);
    const field = d.footer?.segments?.[0]?.[0];
    expect(field && d.setHeaderSegments("footer", [[{ text: "第 " }, field, { text: " 页" }], [{ text: "联系文字" }]])).toBe(true);
    // The typing goes inside the control; the control's own bytes — its properties, the paragraph's style —
    // stay as the file wrote them.
    const out = footerOf(writeDocx(d));
    expect(out).toContain('<w:sdtPr><w:id w:val="77"/></w:sdtPr>');
    expect(out).toContain('<w:pStyle w:val="Footer"/>');
    expect(out).toContain('<w:t xml:space="preserve">第 </w:t>');
    expect(out).toContain("fldChar");
    const again = WordDocument.openBytes(writeDocx(d));
    expect(again.footer?.paragraphs).toEqual(["第  页", "联系文字"]);
    expect(again.footer?.editable).toBe(true);
  });

  it("reads two controls deep, with typed characters between the two fields", () => {
    const inner = control(`<w:p><w:r><w:t>第</w:t></w:r>${pageFieldRuns()}<w:r><w:t>页</w:t></w:r>${pageFieldRuns(" NUMPAGES ", "3")}</w:p>`);
    const d = WordDocument.openBytes(makeFooterDocx(footerPart(control(inner))));
    expect(d.footer?.editable).toBe(true);
    expect(d.footer?.paragraphs).toEqual(["第页"]);
    expect((d.footer?.segments ?? [])[0]?.filter((s) => s.carry).map((s) => s.carry?.show)).toEqual(["页码", "总页数"]);
    const segs = (d.footer?.segments ?? []).map((s) => s.slice());
    segs[0]!.push({ text: "！" });
    expect(d.setHeaderSegments("footer", segs)).toBe(true);
    const out = footerOf(writeDocx(d));
    expect((out.match(/<w:t[^>]*>！<\/w:t>/g) ?? []).length).toBe(1);
    expect((out.match(/<w:sdtPr>/g) ?? []).length).toBe(2);
    expect((out.match(/<w:sdtContent>/g) ?? []).length).toBe(2);
    expect(WordDocument.openBytes(writeDocx(d)).footer?.paragraphs).toEqual(["第页！"]);
  });

  it("hands an untouched control back as its own bytes", () => {
    const part = footerPart(control(`<w:p>${pageFieldRuns()}</w:p>`));
    expect(footerOf(writeDocx(WordDocument.openBytes(makeFooterDocx(part))))).toBe(part);
  });

  it("drops a paragraph the page no longer holds without dropping the control holding it", () => {
    const part = footerPart(
      control(`<w:p><w:r><w:t>控件里第一句</w:t></w:r></w:p><w:p><w:r><w:t>控件里第二句</w:t></w:r></w:p>`),
    );
    const d = WordDocument.openBytes(makeFooterDocx(part));
    expect(d.footer?.paragraphs).toEqual(["控件里第一句", "控件里第二句"]);
    expect(d.setHeaderSegments("footer", [[{ text: "控件里第一句" }]])).toBe(true);
    const out = footerOf(writeDocx(d));
    expect(out).toContain('<w:sdtPr><w:id w:val="77"/></w:sdtPr>');
    expect(out).toContain("<w:sdtContent>");
    expect(out).not.toContain("控件里第二句");
    expect(WordDocument.openBytes(writeDocx(d)).footer?.editable).toBe(true);
  });

  it("draws the control's field as a chip, not as the number it cached", () => {
    const html = renderWord(WordDocument.openBytes(makeFooterDocx(footerPart(control(`<w:p>${pageFieldRuns(undefined, "2")}</w:p>`)))));
    expect(html).toContain(">页码<");
    expect(html).not.toContain(">2<");
  });

  it("refuses a control that says the edit is not allowed", () => {
    for (const pr of ['<w:sdtPr><w:id w:val="77"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>', "<w:sdtPr><w:placeholder/></w:sdtPr>"]) {
      const part = footerPart(control(`<w:p>${pageFieldRuns()}</w:p>`, pr));
      const d = WordDocument.openBytes(makeFooterDocx(part));
      expect(d.footer?.editable, pr).toBe(false);
      expect(footerOf(writeDocx(d)), pr).toBe(part);
    }
  });

  it("refuses a control holding anything but paragraphs", () => {
    const withTable = footerPart(control(`<w:p><w:r><w:t>字</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`));
    expect(WordDocument.openBytes(makeFooterDocx(withTable)).footer?.editable).toBe(false);
    const noProps = `<w:sdt><w:sdtContent><w:p><w:r><w:t>控件</w:t></w:r></w:p></w:sdtContent></w:sdt>`;
    expect(WordDocument.openBytes(makeFooterDocx(footerPart(noProps))).footer?.editable).toBe(false);
  });

  it("refuses a control whose content is not pure markup, and one nested past the third", () => {
    const spaced = `<w:sdt><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent>\n<w:p>${pageFieldRuns()}</w:p>\n</w:sdtContent></w:sdt>`;
    expect(WordDocument.openBytes(makeFooterDocx(footerPart(spaced))).footer?.editable).toBe(false);
    let deep = `<w:p><w:r><w:t>字</w:t></w:r></w:p>`;
    for (let i = 0; i < 4; i++) deep = control(deep);
    expect(WordDocument.openBytes(makeFooterDocx(footerPart(deep))).footer?.editable).toBe(false);
    let three = `<w:p><w:r><w:t>字</w:t></w:r></w:p>`;
    for (let i = 0; i < 3; i++) three = control(three);
    expect(WordDocument.openBytes(makeFooterDocx(footerPart(three))).footer?.editable).toBe(true);
  });

  it("still shows the words a refused control holds", () => {
    // Refusing the edit is a statement about writing, not about the page: a footer whose page number sits in
    // a locked control, or one nested past the level this model claims, still says 「第 … 页」.
    const locked = footerPart(
      control(`<w:p><w:r><w:t xml:space="preserve">第 </w:t></w:r>${pageFieldRuns()}<w:r><w:t xml:space="preserve"> 页</w:t></w:r></w:p>`, '<w:sdtPr><w:id w:val="77"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>'),
    );
    const d = WordDocument.openBytes(makeFooterDocx(locked));
    expect(d.footer?.editable).toBe(false);
    expect(d.footer?.paragraphs.some((p) => p.includes("第") && p.includes("页"))).toBe(true);
    expect(footerOf(writeDocx(d))).toBe(locked);

    let deep = `<w:p><w:r><w:t>深处有字</w:t></w:r></w:p>`;
    for (let i = 0; i < 4; i++) deep = control(deep);
    const dd = WordDocument.openBytes(makeFooterDocx(footerPart(deep)));
    expect(dd.footer?.editable).toBe(false);
    expect(renderWord(dd)).toContain("深处有字");
  });

  it("labels a page count by what it counts", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(footerPart(`<w:p>${pageFieldRuns(" NUMPAGES ", "12")}</w:p>`)));
    expect(doc.footer?.segments?.[0]?.[0]?.carry?.show).toBe("总页数");
    expect(doc.footer?.paragraphs).toEqual([""]);
  });

  it("takes the field out when the page says it is gone", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(PAGE_FOOTER));
    const kept = (doc.footer?.segments?.[0] ?? []).filter((s) => !s.carry);
    expect(doc.setHeaderSegments("footer", [kept, [{ text: "普通一行" }]])).toBe(true);
    const out = footerOf(writeDocx(doc));
    expect(out).not.toContain("fldChar");
    expect(out).toContain("第 ");
  });

  it("inserts a page number spelled the way these files write one", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(footerPart(`<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>机密</w:t></w:r></w:p>`)));
    expect(doc.insertPageNumber("footer")).toBe(true);
    expect(footerOf(writeDocx(doc))).toContain(
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    // The text the user had is still the text the model holds, and the cached digit did not join it.
    expect(doc.footer?.paragraphs).toEqual(["机密"]);
  });

  it("reads back a page number it inserted, and saves the same bytes twice", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(footerPart(`<w:p><w:r><w:t>机密</w:t></w:r></w:p>`)));
    doc.insertPageNumber("footer");
    const once = writeDocx(doc);
    const again = WordDocument.openBytes(once);
    expect(again.footer?.editable, "an inserted field must not be a reason to refuse the footer later").toBe(true);
    expect(again.footer?.paragraphs).toEqual(["机密"]);
    expect(again.footer?.segments?.[0]?.at(-1)?.carry?.show).toBe("页码");
    expect(footerOf(writeDocx(again))).toBe(footerOf(once));
  });

  it("puts the field where the text said, not at the end", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(footerPart(`<w:p><w:r><w:t>机密文件</w:t></w:r></w:p>`)));
    expect(doc.insertPageNumber("footer", 0, 2)).toBe(true);
    expect(doc.footer?.segments?.[0]?.map((s) => (s.carry ? "·" : s.text))).toEqual(["机密", "·", "文件"]);
    expect(doc.footer?.paragraphs).toEqual(["机密文件"]);
  });

  it("refuses to insert where the part cannot be edited", () => {
    const doc = WordDocument.openBytes(makeFooterDocx(footerPart(`<w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>表</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:p>`)));
    expect(doc.footer?.editable).toBe(false);
    expect(doc.insertPageNumber("footer")).toBe(false);
  });

  it("creates a footer to hold the page number when the document had none", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx());
    expect(doc.footer).toBeUndefined();
    expect(doc.ensureHeaderFooter("footer")).toBeDefined();
    expect(doc.insertPageNumber("footer")).toBe(true);
    const out = writeDocx(doc);
    expect(footerOf(out)).toContain('<w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText>');
    const again = WordDocument.openBytes(out);
    expect(again.footer?.part).toBe("word/footer1.xml");
    expect(again.footer?.editable).toBe(true);
    expect(again.footer?.paragraphs).toEqual([""]);
    expect(again.footer?.segments?.[0]?.length).toBe(1);
  });

  it("draws the field as one box the caret cannot enter, and the band's command beside it", () => {
    // The footer band is the last thing the page draws, so it is the tail of the markup.
    const band = (bytes: Uint8Array): string => {
      const html = renderWord(WordDocument.openBytes(bytes));
      const at = html.lastIndexOf('<div class="hf hf-footer');
      return at < 0 ? "" : html.slice(at);
    };
    expect(band(makeFooterDocx(PAGE_FOOTER))).toContain('contenteditable="false"');
    expect(band(makeFooterDocx(PAGE_FOOTER))).toContain('data-carry="');
    expect(band(makeFooterDocx(PAGE_FOOTER))).toContain(">页码</span>");
    // The number the field cached is not on the page, so nothing can read it back as typed text.
    expect(band(makeFooterDocx(PAGE_FOOTER))).not.toContain(">1<");
    expect(band(makeFooterDocx(PAGE_FOOTER))).toContain('<button class="hf-page" data-hf-page="footer"');
    // A band that holds a plain line still offers the command; one nobody can edit offers nothing.
    expect(band(makeFooterDocx(footerPart("<w:p><w:r><w:t>x</w:t></w:r></w:p>")))).toContain('data-hf-page="footer"');
    expect(band(makeFooterDocx(footerPart("<w:p><w:br/></w:p>")))).not.toContain("hf-page");
  });

  describe("a picture in a header or footer band", () => {
    const logo = (): string => WordDocument.openBytes(makeLogoFooterDocx(LOGO_RUN)).footer?.segments?.[0]?.[0]?.carry?.xml ?? "";

    it("carries a textless logo as one piece and leaves the words beside it editable", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(LOGO_RUN));
      expect(d.footer?.editable).toBe(true);
      // The drawing shows the page no text, so the paragraph's text view stays empty rather than inventing one.
      expect(d.footer?.paragraphs).toEqual(["", "联系文字"]);
      expect(logo()).toContain("<w:drawing>");
      expect(renderWord(d)).toContain(">图片</span>");
      expect(renderWord(d)).toContain("data-hf-page");
    });

    it("hands the logo paragraph back byte for byte when only the other paragraph is typed in", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(LOGO_RUN));
      const segs = (d.footer?.segments ?? []).map((p) => p.map((s) => ({ ...s })));
      segs[1] = [{ text: "改成别的" }];
      expect(d.setHeaderSegments("footer", segs)).toBe(true);
      const out = footerOf(writeDocx(d));
      expect(out).toContain(`<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>${logo()}</w:p>`);
      expect(out).toContain("<w:t>改成别的</w:t>");
    });

    it("writes a change that moves no typed character at all", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(LOGO_RUN));
      // Deleting the chip takes a piece out of the model and nothing out of the text view: hashing only the
      // texts left `edited` false, and the save handed back the bytes it read with the logo still in it.
      expect(d.setHeaderSegments("footer", [[], [{ text: "联系文字" }]])).toBe(true);
      expect(footerOf(writeDocx(d))).not.toContain("<w:drawing>");
      expect(footerOf(writeDocx(d))).toContain('<w:pStyle w:val="Footer"/>');
    });

    it("keeps its shape through a save and a re-read", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(LOGO_RUN));
      const again = WordDocument.openBytes(writeDocx(d));
      expect(again.footer?.editable).toBe(true);
      expect(again.footer?.segments?.[0]?.[0]?.carry?.xml).toBe(logo());
    });

    it("inserts a page number into the words beside a logo, leaving the logo alone", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(LOGO_RUN));
      expect(d.insertPageNumber("footer")).toBe(true);
      const out = footerOf(writeDocx(d));
      expect(out).toContain(logo());
      expect(out).toContain('fldCharType="begin"');
      const again = WordDocument.openBytes(writeDocx(d));
      expect((again.footer?.segments ?? []).flat().filter((s) => s.carry).length).toBe(2);
    });
  });

  describe("a floating text box holding the page number", () => {
    /** Every `w:txbxContent` a saved footer holds, so a claim about "both copies" can be checked as counts. */
    const boxes = (out: string): number => (out.match(/<w:txbxContent>/g) ?? []).length;

    it("reads through the frame: the box's words are the band's words", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(PAGE_BOX_RUN));
      expect(d.footer?.editable).toBe(true);
      // `第 ` and ` 页` are text the page shows; the field between them is not, so it never joins the view.
      expect(d.footer?.paragraphs).toEqual(["第  页", "联系文字"]);
      expect(d.footer?.segments?.[0]?.map((s) => s.carry?.show ?? s.text)).toEqual(["第 ", "页码", " 页"]);
      expect(d.footer?.boxedParagraphs).toEqual([true, false]);
      const page = renderWord(d);
      expect(page).toContain("第 ");
      expect(page).toContain(">页码</span>");
      expect(page).not.toContain(">图片</span>");
      expect(page).toContain('data-hf-page="footer"');
    });

    it("writes an edit into both copies of the box and leaves the frame as the file drew it", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(PAGE_BOX_RUN));
      const segs = (d.footer?.segments ?? []).map((p) => p.map((s) => ({ ...s })));
      segs[0] = [{ text: "第 " }, d.footer!.segments![0]![1]!, { text: " 张" }];
      expect(d.setHeaderSegments("footer", segs)).toBe(true);
      const out = footerOf(writeDocx(d));
      expect(boxes(out)).toBe(2);
      expect((out.match(/<w:t[^>]*> 张<\/w:t>/g) ?? []).length).toBe(2);
      expect(out).toContain('<wp:extent cx="1000000" cy="500000"/>');
      expect(out).toContain("<v:textbox");
      expect(out).toContain('<w:pStyle w:val="Footer"/>');
      // The box's own paragraph keeps its alignment and its typeface — the band's paragraph says nothing of them.
      expect((out.match(/<w:jc w:val="right"\/>/g) ?? []).length).toBe(2);
      expect((out.match(/<w:sz w:val="18"\/>/g) ?? []).length).toBeGreaterThanOrEqual(4);
    });

    it("writes a change that moves no typed character inside the box", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(PAGE_BOX_RUN));
      // Dropping the field takes a piece out of the model and nothing out of the text view. A save that
      // hashed only the words handed the box back as it read it, with the page number still in it.
      expect(d.setHeaderSegments("footer", [[{ text: "第  页" }], [{ text: "联系文字" }]])).toBe(true);
      const out = footerOf(writeDocx(d));
      expect(out).not.toContain("fldChar");
      expect(boxes(out)).toBe(2);
      expect((out.match(/<w:t[^>]*>第  页<\/w:t>/g) ?? []).length).toBe(2);
    });

    it("hands an untouched box back byte for byte", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(PAGE_BOX_RUN));
      d.setHeaderSegments("footer", d.footer?.segments?.map((p) => p.map((s) => ({ ...s }))) ?? []);
      const out = footerOf(writeDocx(d));
      expect(out).toContain(PAGE_BOX_RUN);
    });

    it("keeps the box, its words and its field through a save and a re-read", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(PAGE_BOX_RUN));
      const again = WordDocument.openBytes(writeDocx(d));
      expect(again.footer?.editable).toBe(true);
      expect(again.footer?.paragraphs).toEqual(["第  页", "联系文字"]);
      expect(again.footer?.segments?.[0]?.[1]?.carry?.tag).toBe("fldChar");
      // Re-read reads what the save wrote, so the second save is the same bytes as the first.
      expect(footerOf(writeDocx(again))).toBe(footerOf(writeDocx(d)));
    });

    it("refuses the band when the two copies of the box do not say the same thing", () => {
      const d = WordDocument.openBytes(makeLogoFooterDocx(PAGE_BOX_DISAGREEING_RUN));
      expect(d.footer?.editable).toBe(false);
      expect(d.footer?.paragraphs[0]).toContain("第 ");
      const page = renderWord(d);
      expect(page).toContain("联系文字");
      expect(page).not.toContain("hf-page");
    });

    it("refuses a box whose words are split over two paragraphs", () => {
      expect(WordDocument.openBytes(makeLogoFooterDocx(PAGE_BOX_TWO_PARA_RUN)).footer?.editable).toBe(false);
    });

    it("refuses a box that is not the only run of its paragraph", () => {
      // The box's pieces join the paragraph's, so a paragraph that also holds runs of its own has no honest
      // mapping left: which half of the words belong inside the frame? That band stays held as its own bytes.
      const d = WordDocument.openBytes(makeBoxWithWordsFooterDocx(PAGE_BOX_RUN));
      expect(d.footer?.editable).toBe(false);
      expect(d.footer?.paragraphs[0]).toContain("页脚左边");
      expect(renderWord(d)).toContain("页脚左边");
      expect(WordDocument.openBytes(makeLogoFooterDocx(PAGE_BOX_RUN)).footer?.editable).toBe(true);
    });
  });
});

/**
 * `w:pgMar` states seven distances, not four (ECMA-376 §17.3.2.13), and all 58 of them in the authorized
 * 46 份 state all seven (`.scratch/corpus-recheck.mjs`). `w:header` and `w:footer` are measured from the
 * **edge of the paper**, so they say where a band sits rather than how wide the text column is; `w:gutter` is
 * binding width added to the margin it belongs to. Reading only four left every file's bands at the
 * stylesheet's fixed gap and its binding margin nowhere.
 *
 * 真机取到了带子的落点与让位（夹具 `/tmp/tall-header.docx`：页眉 10 行、四边 1440 缇、`w:header` 720 缇，
 * WPS 正文上沿 ≈3445 缇，#159）。仍未取证的是**带子比整页还高**那一支——本版不裁，让它溢出纸。
 */
describe("docx 页眉页脚：w:pgMar 的三个距离", () => {
  /** One band's own box, which is where its placement has to appear: the band lives inside `.page`. */
  const band = (html: string, which: "header" | "footer"): string =>
    new RegExp(`class="hf hf-${which}[^"]*"[^>]*?style="([^"]*)"`).exec(html)?.[1] ?? "";
  /** The page box's own `style`, which carries the four margins and the binding width. */
  const page = (html: string): string => /<div class="page" style="([^"]*)"/.exec(html)?.[1] ?? "";
  const bodyXml = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);

  it("reads the three distances a section states", () => {
    const doc = WordDocument.openBytes(makeFullPageDistanceDocx());
    expect(doc.sections[0]!.margins).toEqual({
      top: 1440,
      right: 1440,
      bottom: 1440,
      left: 1440,
      header: 720,
      footer: 720,
      gutter: 288,
    });
  });

  it("reads them when the section is a paragraph's break", () => {
    // 18 of the 86 files carry more than one section, and the opening one hangs off a paragraph — the shape
    // whose geometry a save splices into rather than one it rebuilds.
    expect(WordDocument.openBytes(makeFullPageDistanceDocx("paragraph")).sections[0]!.margins.header).toBe(720);
  });

  it("invents nothing for a section that states none", () => {
    const m = WordDocument.openBytes(makePageDistanceDocx('w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"')).sections[0]!.margins;
    expect(m.header).toBeUndefined();
    expect(m.footer).toBeUndefined();
    expect(m.gutter).toBeUndefined();
  });

  it("places the band at the distance its own section states", () => {
    // §17.3.2.7/9 那两个数从**页边**量起，所以带子按自己的摆位落在纸上，而不是靠一句内边距把字撑开。
    // 夹具 `/tmp/tall-header.docx` 在 WPS 真机上量到正文为页眉让位（正文上沿 ≈3445 缇），让位本身由
    // `doc-layout.ts` 的 `applyBandYield` 在铺排前量，不在这一串字面里（`bandYieldPx` 钉的是那道算式）。
    const html = renderWord(
      WordDocument.openBytes(makePageDistanceDocx('w:top="2160" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="1440"')),
    );
    expect(band(html, "header")).toContain(`;top:${twipsToPx(720)}px`);
    expect(band(html, "footer")).toContain(`;bottom:${twipsToPx(1440)}px`);
    // 两份带子都排在正文之后、纸盒之内：正文的下标不能被装饰顶偏。
    expect(html.indexOf('<div class="para"')).toBeLessThan(html.indexOf('solo-hf-master'));
  });

  it("puts a distance past the margin on the paper rather than clamping it away", () => {
    // 距离比边距还大时那一截落在正文里，由 `applyBandYield` 让出来——摆位没有「负的内边距」可写，
    // 也就不会再出现旧那一句 `max(0, …)` 把数夹掉、纸上看不到带子的两支。
    const html = renderWord(WordDocument.openBytes(makePageDistanceDocx('w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="2160" w:footer="2160"')));
    expect(band(html, "header")).toContain(`;top:${twipsToPx(2160)}px`);
    expect(band(html, "footer")).toContain(`;bottom:${twipsToPx(2160)}px`);
  });

  it("puts the band on the paper edge where the file said no distance", () => {
    const html = renderWord(WordDocument.openBytes(makePageDistanceDocx('w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"')));
    expect(band(html, "header")).toContain(";top:0px");
    expect(band(html, "footer")).toContain(";bottom:0px");
    // 发明 Word 的 docDefaults（页眉 720 / 页脚 720）会把一个没人写过的数画到纸上。
    expect(band(html, "header")).not.toContain(`top:${twipsToPx(720)}px`);
  });

  it("adds the binding width to the left margin the page and the bands draw with", () => {
    const html = renderWord(WordDocument.openBytes(makePageDistanceDocx('w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="288"')));
    const withGutter = `${twipsToPx(1440 + 288)}px`;
    expect(page(html)).toContain(`--pad-l:${withGutter}`);
    // The band's text starts at the body's column, so the same width has to reach it.
    expect(band(html, "header")).toContain("padding-left:var(--pad-l)");
    expect(band(html, "footer")).toContain("padding-left:var(--pad-l)");
    // Without a stated gutter nothing is added — the column stays the margin the file wrote.
    expect(page(renderWord(WordDocument.openBytes(makePageDistanceDocx('w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'))))).toContain(
      `--pad-l:${twipsToPx(1440)}px`,
    );
  });

  it("hands the three distances back on a save, whichever sectPr stated them", () => {
    // Every real file's section travels as the bytes it was read from, so a save that dropped one of these
    // would be a rebuild — which is what the two shapes below pin from both sides.
    const stated = 'w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="288"';
    for (const at of ["body", "paragraph"] as const) {
      const bytes = writeDocx(WordDocument.openBytes(makePageDistanceDocx(stated, at)));
      expect(bodyXml(bytes)).toContain(`<w:pgMar ${stated}/>`);
    }
    // A section that stated none is not owed a set of numbers, either.
    const bare = writeDocx(WordDocument.openBytes(makePageDistanceDocx('w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"')));
    expect(bodyXml(bare)).not.toContain("w:header=");
    expect(bodyXml(bare)).not.toContain("w:gutter=");
  });

  it("still states them after the header's words are edited and the file saved", () => {
    const out = edited(makeFullPageDistanceDocx(), ["改过的页眉"]);
    expect(bodyXml(out)).toContain('w:header="720" w:footer="720" w:gutter="288"');
    expect(headerOf(out)).toContain("改过的页眉");
    // Re-read reads the saved file, so the distances are the same on the second open.
    expect(WordDocument.openBytes(out).sections[0]!.margins.gutter).toBe(288);
  });
});
