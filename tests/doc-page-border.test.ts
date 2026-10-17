import { describe, expect, it } from "vitest";
import { WordDocument, twipsToPx } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { WPS_CLEARED_BORDER_SECTION, WPS_PAGE_BORDER_SECTION, WPS_SECTION, makePageSetupDocx } from "./helpers/make-page-setup.js";

const bodyXml = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const tagOf = (xml: string, name: string): string => new RegExp(`<w:${name}[^>]*/?>`).exec(xml)?.[0] ?? "";
/**
 * The whole element, children and all. `tagOf` stops at the first `>` it finds, which is the entire story for a
 * self-closing `w:pgSz` and a lie about `w:pgBorders`: its border set is its children, so a comparison that used
 * the opening tag alone would pass on a file that lost all four sides.
 */
const spanOf = (xml: string, name: string): string => new RegExp(`<w:${name}\\b[^>]*/>|<w:${name}\\b[^>]*>[\\s\\S]*?<\\/w:${name}>`).exec(xml)?.[0] ?? "";
const pageStyle = (html: string): string => /<div class="page" style="([^"]*)"/.exec(html)?.[1] ?? "";
/** One border side as the page drew it, found by the side the file named rather than by its position. */
const borderDiv = (html: string, side: string): string =>
  new RegExp(`<div class="doc-page-border" data-side="${side}"[^>]*style="([^"]*)"`).exec(html)?.[1] ?? "";

describe("docx 页面边框（w:pgBorders）：读、画、按原件带回", () => {
  it("reads the four sides with the stroke, width, gap and colour the file gave each", () => {
    const sec = WordDocument.openBytes(makePageSetupDocx(WPS_PAGE_BORDER_SECTION)).sections[0]!;
    expect(sec.pageBorders).toEqual({
      offsetFrom: "page",
      sides: {
        // `w:sz` is the file's eighths of a point and `w:space` whole points (§17.3.2.10); both keep their own
        // scale because they are multiplied by different factors on the way to a pixel.
        top: { val: "single", size: 8, space: 6, color: "auto" },
        left: { val: "dashed", size: 12, space: 0, color: "FF0000" },
        bottom: { val: "double", size: 4, space: 12 },
        right: { val: "none", size: 4, space: 0 },
      },
    });
  });

  it("leaves a section that states no border without one, rather than an empty border set", () => {
    // Absence has to survive the read: a section holding `w:pgBorders` with no usable side and a section holding
    // no `w:pgBorders` at all are different statements about the file, and only the first is a save's to keep.
    expect(WordDocument.openBytes(makePageSetupDocx(WPS_SECTION)).sections[0]!.pageBorders).toBeUndefined();
    const empty = WordDocument.openBytes(makePageSetupDocx('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgBorders/><w:cols w:space="425"/></w:sectPr>'));
    expect(empty.sections[0]!.pageBorders).toEqual({ offsetFrom: "page", sides: {} });
  });

  it("reads the gap against the body when the file says so, and against the paper when it says nothing", () => {
    const fromText = WordDocument.openBytes(makePageSetupDocx('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:pgBorders w:offsetFrom="text"><w:top w:val="single" w:sz="6" w:space="3"/></w:pgBorders></w:sectPr>'));
    expect(fromText.sections[0]!.pageBorders!.offsetFrom).toBe("text");
    // None of the corpus's 24 spelled `w:offsetFrom`, so the default §17.3.2.10 gives it is the common case.
    expect(WordDocument.openBytes(makePageSetupDocx(WPS_PAGE_BORDER_SECTION)).sections[0]!.pageBorders!.offsetFrom).toBe("page");
  });

  it("drops a side that names no stroke, and ignores a child that is not one of the four", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx('<w:sectPr><w:pgBorders><w:top w:sz="8"/><w:left w:val="single" w:sz="8"/><w:hairline w:val="single"/></w:pgBorders></w:sectPr>'));
    expect(doc.sections[0]!.pageBorders!.sides).toEqual({ left: { val: "single", size: 8 } });
  });

  it("draws each side its own line at its own gap, and no line for the sides that say none", () => {
    const html = renderWord(WordDocument.openBytes(makePageSetupDocx(WPS_PAGE_BORDER_SECTION)));
    // 6 points is 8 CSS px, and 8 eighths of a point is one pixel wide — the two units the file counts in.
    expect(borderDiv(html, "top")).toBe("position:absolute;pointer-events:none;left:0;right:0;top:8px;border-top:1px solid currentColor;");
    expect(borderDiv(html, "left")).toContain("border-left:2px dashed #FF0000;");
    expect(borderDiv(html, "left")).toContain("left:0px;");
    expect(borderDiv(html, "bottom")).toContain("bottom:16px;");
    expect(borderDiv(html, "bottom")).toContain("border-bottom:1px double currentColor;");
    expect(borderDiv(html, "right")).toBe("");
  });

  it("adds the margin to the gap when the border is measured from the body", () => {
    const html = renderWord(WordDocument.openBytes(makePageSetupDocx('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:pgBorders w:offsetFrom="text"><w:top w:val="single" w:sz="6" w:space="3"/></w:pgBorders></w:sectPr>')));
    // 3 points off the text body, which starts 1440 twips (96 px) in: the two measurements add.
    expect(borderDiv(html, "top")).toContain(`top:${Math.round(4 + twipsToPx(1440))}px;`);
  });

  it("hangs the decoration off the end of the body, so a block keeps the position the editor addresses it by", () => {
    const html = renderWord(WordDocument.openBytes(makePageSetupDocx(WPS_PAGE_BORDER_SECTION)));
    expect(html.indexOf("doc-page-border")).toBeGreaterThan(html.indexOf("正文"));
    // 纸上的装饰都要有一个定位祖先：页面边框如此，#157 之后贴在纸内的两份带子原件也如此，所以这一格不再看条件。
    expect(pageStyle(html)).toContain("position:relative");
    expect(pageStyle(renderWord(WordDocument.openBytes(makePageSetupDocx(WPS_SECTION))))).toContain("position:relative");
    expect(renderWord(WordDocument.openBytes(makePageSetupDocx(WPS_SECTION)))).not.toContain("doc-page-border");
  });

  it("draws nothing for the only page-border shape WPS' own files show: four sides that each say none", () => {
    // 授权 46 份里 `w:pgBorders` 是 **0 枚**（`.scratch/oneoff-gap.mjs`）：这份夹具画的不是眼下有受众的形状，
    // 钉的是「文件写了边框就按原件读回来、写了 none 就不画线」那两条规则本身。
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_CLEARED_BORDER_SECTION));
    expect(doc.sections[0]!.pageBorders).toEqual({
      offsetFrom: "page",
      sides: {
        top: { val: "none", size: 0, space: 0 },
        left: { val: "none", size: 0, space: 0 },
        bottom: { val: "none", size: 0, space: 0 },
        right: { val: "none", size: 0, space: 0 },
      },
    });
    const html = renderWord(doc);
    expect(html).not.toContain("doc-page-border");
    // 纸盒是不是定位祖先已经是常量（上一条），清空的边框不再往这里加东西。
    // A cleared border is still the file's statement. Dropping it would be this model editing the page.
    const xml = bodyXml(writeDocx(doc));
    expect(spanOf(xml, "pgBorders")).toBe('<w:pgBorders><w:top w:val="none" w:sz="0" w:space="0"/><w:left w:val="none" w:sz="0" w:space="0"/><w:bottom w:val="none" w:sz="0" w:space="0"/><w:right w:val="none" w:sz="0" w:space="0"/></w:pgBorders>');
  });


  it("keeps the border's bytes through a page-setup change, and the border readable afterwards", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_PAGE_BORDER_SECTION));
    doc.setPageSetup({ topTwips: 1000, rightTwips: 1000, bottomTwips: 1000, leftTwips: 1000, landscape: false });
    const xml = bodyXml(writeDocx(doc));
    // `applyPageSetup` edits `w:pgSz` and `w:pgMar` where they stand; everything beside them, including a
    // border set the dialog has no field for, is the file's own bytes.
    expect(xml).toContain("<w:pgBorders><w:top w:val=\"single\" w:sz=\"8\" w:space=\"6\" w:color=\"auto\"/><w:left w:val=\"dashed\" w:sz=\"12\" w:space=\"0\" w:color=\"FF0000\"/><w:bottom w:val=\"double\" w:sz=\"4\" w:space=\"12\"/><w:right w:val=\"none\" w:sz=\"4\" w:space=\"0\"/></w:pgBorders>");
    expect(tagOf(xml, "pgMar")).toContain('w:top="1000"');
    expect(doc.sections[0]!.pageBorders!.sides.top!.space).toBe(6);
  });

  it("puts the page box and the margins ahead of a border set that was already there", () => {
    // §17.3 orders `w:pgSz`, `w:pgMar`, `w:paperSrc`, `w:pgBorders`; a `w:pgMar` written after the borders is a
    // file a strict reader refuses, so the elements a page-setup change adds land in sequence, not appended.
    const doc = WordDocument.openBytes(makePageSetupDocx('<w:sectPr><w:pgBorders><w:top w:val="single" w:sz="8"/></w:pgBorders><w:cols w:space="720"/></w:sectPr>'));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: false });
    const xml = bodyXml(writeDocx(doc));
    expect(xml.indexOf("<w:pgSz")).toBeLessThan(xml.indexOf("<w:pgMar"));
    expect(xml.indexOf("<w:pgMar")).toBeLessThan(xml.indexOf("<w:pgBorders"));
    expect(xml.indexOf("<w:pgBorders")).toBeLessThan(xml.indexOf("<w:cols"));
  });

  it("brings the border back with the page when a page-setup step is undone", () => {
    const bytes = makePageSetupDocx(WPS_PAGE_BORDER_SECTION);
    const untouched = spanOf(bodyXml(writeDocx(WordDocument.openBytes(bytes))), "pgBorders");
    const doc = WordDocument.openBytes(bytes);
    const before = doc.pageGeometryState();
    doc.setPageSetup({ topTwips: 720, rightTwips: 720, bottomTwips: 720, leftTwips: 720, landscape: true });
    expect(doc.sections[0]!.margins.top).toBe(720);
    doc.restorePageGeometry(before);
    expect(doc.sections[0]!.margins.top).toBe(1440);
    expect(doc.sections[0]!.pageBorders!.sides.left!.val).toBe("dashed");
    expect(spanOf(bodyXml(writeDocx(doc)), "pgBorders")).toBe(untouched);
  });

  it("leaves a document created here with no page border, since nobody asked for one", () => {
    expect(WordDocument.blank().sections[0]!.pageBorders).toBeUndefined();
    expect(renderWord(WordDocument.blank())).not.toContain("doc-page-border");
  });
});
