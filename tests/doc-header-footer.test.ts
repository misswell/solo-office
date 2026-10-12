import { describe, expect, it } from "vitest";
import { WordDocument, twipsToPx } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import {
  LAST_SECTION_HEADER,
  makeFooterDocx,
  makeHeaderFooterDocx,
  makeStyledHeaderDocx,
  makeTwoSectionDocx,
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
      ["容器", `<w:hdr xmlns:w="${NS_W}"><w:p><w:r><w:t>字</w:t></w:r></w:p><w:sdt><w:sdtContent><w:p><w:r><w:t>控件</w:t></w:r></w:p></w:sdtContent></w:sdt></w:hdr>`, ["字"]],
      // The old gate read only `w:`, so a part that binds the namespace under another name slipped past it
      // and had its container rebuilt away.
      ["别的前缀下的容器", `<w0:hdr xmlns:w0="${NS_W}"><w0:p><w0:r><w0:t>字</w0:t></w0:r></w0:p><w0:sdt><w0:sdtContent><w0:p><w0:r><w0:t>控件</w0:t></w0:r></w0:p></w0:sdtContent></w0:sdt></w0:hdr>`, ["字"]],
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

