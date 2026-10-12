import { describe, expect, it } from "vitest";
import { WordDocument, twipsToPx } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import {
  LAST_SECTION_HEADER,
  PAGE_FOOTER,
  PAGE_ONLY_FOOTER,
  PAGE_ONLY_FOOTER_STYLED,
  footerPart,
  LOGO_RUN,
  PAGE_BOX_DISAGREEING_RUN,
  PAGE_BOX_RUN,
  PAGE_BOX_TWO_PARA_RUN,
  makeBoxWithWordsFooterDocx,
  makeFooterDocx,
  makeLogoFooterDocx,
  makeHeaderFooterDocx,
  makeStyledHeaderDocx,
  makeTwoSectionDocx,
  pageFieldRuns,
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
   * nothing to edit, and no part for a save to splice back.
   */
  it("finds the part whatever order the reference puts its attributes in", () => {
    for (const ref of [REFERENCE_ID_FIRST, REFERENCE_TYPE_FIRST, REFERENCE_FIRST_PAGE_ONLY]) {
      const doc = WordDocument.openBytes(makeHeaderFooterDocx(STYLED_HEADER, ref));
      expect(doc.header?.id, ref).toBe("rId5");
      expect(doc.header?.paragraphs, ref).toEqual(["机密\t第一段", "第二段"]);
    }
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
  it("draws one line per paragraph, outside the page, with the tab kept", () => {
    const html = renderWord(WordDocument.openBytes(makeStyledHeaderDocx()));
    expect(html).toContain('<div class="hf-p">机密\t第一段</div>');
    expect(html).toContain('<div class="hf-p">第二段</div>');
    expect(html.indexOf('class="hf hf-header"')).toBeLessThan(html.indexOf('class="page"'));
  });

  it("puts the footer under the page and names the band it offers to edit", () => {
    const html = renderWord(WordDocument.openBytes(makeFooterDocx(plainFooter("第 1 页"))));
    expect(html).toContain('data-hf="footer"');
    expect(html).toContain('class="hf hf-footer hf-bottom"');
    expect(html).toContain('title="点击编辑页脚"');
    expect(html).toContain('<div class="hf-p">第 1 页</div>');
    expect(html.indexOf('class="page"')).toBeLessThan(html.indexOf("hf hf-footer"));
    expect(renderWord(WordDocument.openBytes(makeStyledHeaderDocx()))).toContain('class="hf hf-header"');
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
