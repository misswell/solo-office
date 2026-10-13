import { describe, expect, it } from "vitest";
import { WordDocument, cmToTwips, twipsToCm, twipsToPx } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { BARE_SECTION, WPS_LANDSCAPE_SECTION, WPS_NEAR_A4_SECTION, WPS_SECTION, WPS_SECTION_NO_DISTANCES, WPS_SECTION_UNREAD_CHILDREN, makePageSetupDocx } from "./helpers/make-page-setup.js";
import { makeFullPageDistanceDocx, makeTwoSectionDocx } from "./helpers/make-header-footer.js";

const bodyXml = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
/** One element's own opening tag, so a test can name the bytes a change wrote rather than a re-read value. */
const tagOf = (xml: string, name: string): string => new RegExp(`<w:${name}[^>]*/?>`).exec(xml)?.[0] ?? "";
/** A whole element, opening tag through closing tag: what a section re-splice has to bring back, not just start. */
const spanOf = (xml: string, name: string): string => new RegExp(`<w:${name}\\b[^>]*/>|<w:${name}\\b[^>]*>[\\s\\S]*?<\\/w:${name}>`).exec(xml)?.[0] ?? "";
/** The page box's own `style`, which is where the geometry a user typed has to show up. */
const page = (html: string): string => /<div class="page" style="([^"]*)"/.exec(html)?.[1] ?? "";

describe("docx 页面设置：纸张大小/页边距/距离/装订线/方向", () => {
  it("converts centimetres the way the dialog counts them", () => {
    // The dialog speaks centimetres and the part counts twips (§17.3.2.13): 1440 twips = 2.54 cm exactly.
    expect(cmToTwips(2.54)).toBe(1440);
    expect(twipsToCm(1440)).toBeCloseTo(2.54, 6);
    // Every distance the local corpus states survives a round trip through the dialog's unit.
    for (const tw of [1440, 1800, 851, 992, 283, 425, 720, 2160]) expect(cmToTwips(twipsToCm(tw))).toBe(tw);
    // The same round trip twice over: a page set once must not drift when it is opened again.
    const once = cmToTwips(twipsToCm(11906));
    expect(cmToTwips(twipsToCm(once))).toBe(once);
    expect(twipsToPx(once)).toBe(793.7333333333333);
  });

  it("writes the four margins where the file spelled them and leaves the rest of the section alone", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION));
    doc.setPageSetup({ topTwips: 1000, rightTwips: 1200, bottomTwips: 1000, leftTwips: 1300, landscape: false });
    const xml = bodyXml(writeDocx(doc));
    expect(tagOf(xml, "pgMar")).toBe('<w:pgMar w:top="1000" w:right="1200" w:bottom="1000" w:left="1300" w:header="851" w:footer="992" w:gutter="283"/>');
    // The column and grid settings this model has no field for travel byte for byte, in their place.
    expect(xml).toContain('<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>');
    expect(doc.sections[0]!.margins.gutter).toBe(283);
  });

  it("invents no distance for a section that states none", () => {
    // The three paper-edge distances stay absent unless the dialog carries a number for them: writing Word's
    // docDefaults here would put a value in the file that neither the file nor the user named.
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION_NO_DISTANCES));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1800, landscape: false });
    const mar = tagOf(bodyXml(writeDocx(doc)), "pgMar");
    expect(mar).toBe('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800"/>');
    expect(mar).not.toContain("w:header=");
    expect(mar).not.toContain("w:footer=");
    expect(mar).not.toContain("w:gutter=");
    // A distance the user *did* type is the one thing that earns its place, and §17.3's order is where it goes.
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1800, footerTwips: 850, landscape: false });
    expect(tagOf(bodyXml(writeDocx(doc)), "pgMar")).toBe('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:footer="850"/>');
    expect(WordDocument.openBytes(writeDocx(doc)).sections[0]!.margins.header).toBeUndefined();
  });

  it("turns the paper without resizing it, and says so the only way the corpus does", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: true });
    // All 22 landscape sections in the 86 files this machine's WPS wrote put the long edge in `w:w` **and**
    // spell `w:orient="landscape"`; the 93 portrait ones spell no `w:orient` at all.
    expect(tagOf(bodyXml(writeDocx(doc)), "pgSz")).toBe('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');

    const back = WordDocument.openBytes(makePageSetupDocx(WPS_LANDSCAPE_SECTION));
    back.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: false });
    const xml = bodyXml(writeDocx(back));
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(xml).not.toContain("w:orient");
  });

  it("keeps the attribute order and the attributes it never read", () => {
    // An element that spells its own margins out of §17.3's order, and a page box carrying a `w:code` this
    // model has no field for: the change is the values, and rewriting the element would take both down.
    const doc = WordDocument.openBytes(
      makePageSetupDocx('<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:left="1800" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr>'),
    );
    doc.setPageSetup({ topTwips: 2000, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1000, headerTwips: 720, landscape: true });
    const xml = bodyXml(writeDocx(doc));
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape" w:code="9"/>');
    expect(tagOf(xml, "pgMar")).toBe('<w:pgMar w:left="1000" w:right="1440" w:top="2000" w:bottom="1440" w:header="720"/>');
  });

  it("carries every child it never reads through the one save that rebuilds the section", () => {
    // 「页面设置」 is the only edit that re-spells `w:sectPr` bytes: typing does not, because there the section
    // travels as the file's own bytes untouched. So this is where an unread child can die — `applyPageSetup`
    // walks the children and patches only `w:pgSz` and `w:pgMar`, and everything else in §17.3's list has to
    // come back whole and in place. The shape is the corpus's: over the 86 files this machine's WPS wrote, those
    // two elements plus `w:pgBorders` (24 sections), `w:cols`/`w:docGrid` (all 115), `w:pgNumType` (3) and
    // `w:titlePg` (3) are the only children real files state, and every one came back byte-identical.
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION_UNREAD_CHILDREN));
    doc.setPageSetup({ topTwips: 1000, rightTwips: 1200, bottomTwips: 1000, leftTwips: 1300, landscape: false });
    const xml = bodyXml(writeDocx(doc));
    expect(spanOf(xml, "sectPr")).toBe(
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1200" w:bottom="1000" w:left="1300" w:header="851" w:footer="992" w:gutter="0"/><w:pgBorders><w:top w:val="none" w:sz="0" w:space="0"/><w:left w:val="none" w:sz="0" w:space="0"/><w:bottom w:val="none" w:sz="0" w:space="0"/><w:right w:val="none" w:sz="0" w:space="0"/></w:pgBorders><w:pgNumType w:start="1"/><w:cols w:space="425" w:num="1"/><w:titlePg/><w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/></w:sectPr>',
    );

    // Twice, because the second walk reads what the first one wrote: a section that survives one dialog and is
    // rebuilt by the next is the shape a real editing session reaches, and a child kept by luck rather than by
    // the walk would go here.
    const again = WordDocument.openBytes(writeDocx(doc));
    again.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: true });
    const xml2 = bodyXml(writeDocx(again));
    expect(spanOf(xml2, "sectPr")).toBe(
      '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/><w:pgBorders><w:top w:val="none" w:sz="0" w:space="0"/><w:left w:val="none" w:sz="0" w:space="0"/><w:bottom w:val="none" w:sz="0" w:space="0"/><w:right w:val="none" w:sz="0" w:space="0"/></w:pgBorders><w:pgNumType w:start="1"/><w:cols w:space="425" w:num="1"/><w:titlePg/><w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/></w:sectPr>',
    );
    // And the twice-written bytes still say what the file said: a carried child is kept whole, not kept as
    // padding, and this model reads it back exactly as far as it knows how.
    expect(WordDocument.openBytes(writeDocx(again)).sections[0]!.pageBorders!.sides.top!.val).toBe("none");
  });

  it("adds a page box and margins to a section that had neither, at their place in the sequence", () => {
    // A minimal producer's section carries only its columns. §17.3 puts `w:pgSz` and `w:pgMar` before `w:cols`,
    // and a `w:pgMar` written after `w:docGrid` is a file Word refuses to read.
    const doc = WordDocument.openBytes(makePageSetupDocx(BARE_SECTION));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1800, bottomTwips: 1440, leftTwips: 1800, headerTwips: 851, landscape: false });
    const xml = bodyXml(writeDocx(doc));
    expect(xml).toContain('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851"/><w:cols w:space="720"/></w:sectPr>');
    expect(WordDocument.openBytes(writeDocx(doc)).sections[0]!.margins.header).toBe(851);
  });

  it("changes the section that owns the opening page when that is a paragraph's break", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION, "paragraph"));
    doc.setPageSetup({ topTwips: 900, rightTwips: 900, bottomTwips: 900, leftTwips: 900, landscape: false });
    const xml = bodyXml(writeDocx(doc));
    expect(tagOf(xml, "pgMar")).toContain('w:top="900"');
    // The geometry travelled inside the paragraph's own `w:pPr`: a body-level section appearing beside it
    // would be a second section the file never had.
    expect(xml.indexOf('<w:pgMar w:top="900"')).toBeLessThan(xml.indexOf("</w:pPr>"));
    expect((xml.match(/<w:sectPr/g) ?? []).length).toBe(1);
    expect(WordDocument.openBytes(writeDocx(doc)).sections[0]!.margins.top).toBe(900);
  });

  it("counts as an edit, so a save regenerates instead of handing back the bytes it read", () => {
    // The trap: `writeDocx` returns the source package untouched while `edited` says nothing changed, and the
    // page's geometry is not part of `blocks`. A change that only moved the section has to move the hash.
    const bytes = makeFullPageDistanceDocx();
    const opened = WordDocument.openBytes(bytes);
    expect(opened.edited).toBe(false);
    opened.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, headerTwips: 900, landscape: false });
    expect(opened.edited).toBe(true);
    expect(tagOf(bodyXml(writeDocx(opened)), "pgMar")).toContain('w:header="900"');
  });

  it("states a section only when the user asked for geometry the file never had", () => {
    // Without a page-setup change a save still owes nothing, even when it regenerates the part for other
    // reasons: the fallback spelling this writer would use is the app's own page, not the file's, and writing
    // it is how a document gains a `w:sectPr` it never had.
    const typed = WordDocument.openBytes(makePageSetupDocx("", "none"));
    typed.blocks = [{ kind: "paragraph", inlines: [{ text: "改过的正文" }] }];
    expect(typed.edited).toBe(true);
    expect(tagOf(bodyXml(writeDocx(typed)), "pgMar")).toBe("");

    const doc = WordDocument.openBytes(makePageSetupDocx("", "none"));
    doc.setPageSetup({ topTwips: 1000, rightTwips: 1000, bottomTwips: 1000, leftTwips: 1000, footerTwips: 992, landscape: true });
    const xml = bodyXml(writeDocx(doc));
    // The numbers the user typed are the whole reason a section exists now, and the writer spells them in the
    // order §17.3 gives — the four margins first, then the distances.
    expect(tagOf(xml, "pgMar")).toBe('<w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:footer="992"/>');
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
    expect(WordDocument.openBytes(writeDocx(doc)).sections[0]!.margins.footer).toBe(992);
    // Saving the same document again writes the same bytes: the second file's section now travels as read.
    const saved = writeDocx(doc);
    expect(bodyXml(writeDocx(WordDocument.openBytes(saved)))).toBe(bodyXml(saved));
  });

  it("changes only the first section, which is the page the app draws", () => {
    const doc = WordDocument.openBytes(makeTwoSectionDocx({ firstLandscape: false }));
    expect(doc.sections.length).toBe(2);
    doc.setPageSetup({ topTwips: 720, rightTwips: 720, bottomTwips: 720, leftTwips: 720, landscape: false });
    expect(doc.sections[0]!.margins.top).toBe(720);
    // The closing section's own page box is untouched — 16838×11906, spelled by the file and not ours to move.
    expect(doc.sections[1]!.pageWidthTwips).toBe(16838);
    const xml = bodyXml(writeDocx(doc));
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(xml).toContain('<w:sectPr w:rsidR="00B2"><w:headerReference r:id="rId6" w:type="default"/><w:pgSz w:w="16838" w:h="11906"/></w:sectPr>');
  });

  it("carries the page in an undo step, not only the text", () => {
    // The model half of what `DocEditor` restores: the body's `w:sectPr` bytes and the geometry read off them
    // live outside `blocks`, so a step back has to put both back. (The editor wiring is proven in the DOM.)
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION));
    const before = doc.pageGeometryState();
    doc.setPageSetup({ topTwips: 720, rightTwips: 720, bottomTwips: 720, leftTwips: 720, landscape: true });
    expect(doc.sections[0]!.margins.top).toBe(720);
    doc.restorePageGeometry(before);
    expect(doc.sections[0]!.margins.top).toBe(1440);
    expect(doc.pageGeometryState().bodySectPr).toBe(before.bodySectPr);
    expect(doc.edited).toBe(false);
  });

  it("draws the page the dialog set", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION));
    expect(page(renderWord(doc))).toContain(`--page-w:${twipsToPx(11906)}px`);
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 2880, gutterTwips: 283, landscape: true });
    const html = page(renderWord(doc));
    expect(html).toContain(`--page-w:${twipsToPx(16838)}px`);
    // The binding width still rides on the left margin the new value set.
    expect(html).toContain(`--pad-l:${twipsToPx(2880 + 283)}px`);
  });

  it("states the page a document created here is drawn on", () => {
    // A new document has no file to be faithful to, and 「页面设置」 on it has to start from numbers rather than
    // an empty box. Before this, the writer's fallback spelled `<w:pgSz/>` with no width: a page of unknown size.
    const fresh = WordDocument.blank();
    expect(tagOf(bodyXml(writeDocx(fresh)), "pgSz")).toBe('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(tagOf(bodyXml(writeDocx(fresh)), "pgMar")).toBe('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>');
    fresh.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, gutterTwips: 283, landscape: true });
    const xml = bodyXml(writeDocx(fresh));
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
    expect(tagOf(xml, "pgMar")).toBe('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="283"/>');
    // Reading the saved file back gives the same page the dialog was showing.
    const again = WordDocument.openBytes(writeDocx(fresh));
    expect(again.sections[0]!.margins.gutter).toBe(283);
    expect(page(renderWord(again))).toContain(`--page-w:${twipsToPx(16838)}px`);
  });

  it("counts the sheets the list names in the twips a producer spells", () => {
    // `w:pgSz` stores two distances and has no names (§17.3.2.12), so 「A4」 is the dialog's word for a pair.
    // Each pair below is the sheet as its own standard states it — ISO 216 in whole tenths of a centimetre,
    // 信纸 as 8.5×11 inches — and each converts to the count the files carry: A4's 11906×16838 is what 111 of
    // the corpus's 115 sections write.
    expect([cmToTwips(21), cmToTwips(29.7)]).toEqual([11906, 16838]);
    expect([cmToTwips(29.7), cmToTwips(42)]).toEqual([16838, 23811]);
    expect([cmToTwips(14.8), cmToTwips(21)]).toEqual([8391, 11906]);
    expect([cmToTwips(25), cmToTwips(35.3)]).toEqual([14173, 20013]);
    expect([cmToTwips(17.6), cmToTwips(25)]).toEqual([9978, 14173]);
    expect([cmToTwips(21.59), cmToTwips(27.94)]).toEqual([12240, 15840]);
    expect([cmToTwips(21.59), cmToTwips(35.56)]).toEqual([12240, 20160]);
  });

  it("cuts the page to the sheet the dialog named and moves nothing else", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1800, landscape: false, paperTwips: { width: 12240, height: 15840 } });
    const xml = bodyXml(writeDocx(doc));
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="12240" w:h="15840"/>');
    // The distances the file stated stay its own, and the settings behind them travel byte for byte: the sheet
    // is the only thing that was asked for.
    expect(tagOf(xml, "pgMar")).toBe('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="283"/>');
    expect(xml).toContain('<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>');
    expect(WordDocument.openBytes(writeDocx(doc)).sections[0]!.pageWidthTwips).toBe(12240);
    expect(page(renderWord(doc))).toContain(`--page-w:${twipsToPx(12240)}px`);
  });

  it("reads the two numbers as one sheet, whichever box each was typed in", () => {
    // The direction the page lies in is asked about on its own, so the pair only says which sheet: the long
    // edge goes where the direction puts it. A dialog that filled 宽 with 42 while 横向 is checked and one that
    // filled it with 29.7 while 纵向 is checked both say A3 lying as the radios ask.
    const bytes = makePageSetupDocx(WPS_SECTION);
    const margins = { topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440 };
    const sideways = WordDocument.openBytes(bytes);
    sideways.setPageSetup({ ...margins, landscape: true, paperTwips: { width: 23811, height: 16838 } });
    expect(tagOf(bodyXml(writeDocx(sideways)), "pgSz")).toBe('<w:pgSz w:w="23811" w:h="16838" w:orient="landscape"/>');

    const upright = WordDocument.openBytes(bytes);
    upright.setPageSetup({ ...margins, landscape: false, paperTwips: { width: 23811, height: 16838 } });
    expect(tagOf(bodyXml(writeDocx(upright)), "pgSz")).toBe('<w:pgSz w:w="16838" w:h="23811"/>');
  });

  it("keeps the page lying as it was when only the sheet moves", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_LANDSCAPE_SECTION));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: true, paperTwips: { width: 15840, height: 12240 } });
    // Long edge first and `w:orient` still spelled — the only way the corpus's 22 landscape sections say it.
    expect(tagOf(bodyXml(writeDocx(doc)), "pgSz")).toBe('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>');
  });

  it("leaves a sheet that is on no list at the numbers the file wrote", () => {
    // Four of the 115 sections spell a near-A4 11900×16840. The dialog shows it as 自定义尺寸 with 20.99 and
    // 29.704 in the boxes, and those centimetres have to convert back to the pair the file had — a page pulled
    // to the nearest named sheet would be six twips narrower than the document it came from.
    expect(cmToTwips(Number(twipsToCm(11900).toFixed(3)))).toBe(11900);
    expect(cmToTwips(Number(twipsToCm(16840).toFixed(3)))).toBe(16840);
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_NEAR_A4_SECTION));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1800, bottomTwips: 1440, leftTwips: 1800, landscape: false, paperTwips: { width: 11900, height: 16840 } });
    const xml = bodyXml(writeDocx(doc));
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="11900" w:h="16840"/>');
    expect(tagOf(xml, "pgMar")).toContain('w:right="1800"');
    expect(WordDocument.openBytes(writeDocx(doc)).sections[0]!.pageHeightTwips).toBe(16840);
  });

  it("patches the page box where it stands, keeping its order and what this model has no field for", () => {
    const doc = WordDocument.openBytes(
      makePageSetupDocx('<w:sectPr><w:pgSz w:h="16838" w:w="11906" w:code="9"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'),
    );
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: false, paperTwips: { width: 12240, height: 15840 } });
    // `w:h` first because the file wrote it first, and `w:code` because nothing here owns it. Rewriting the
    // element to §17.3's order would take both down, and a portrait page gains no `w:orient` on its own.
    expect(tagOf(bodyXml(writeDocx(doc)), "pgSz")).toBe('<w:pgSz w:h="15840" w:w="12240" w:code="9"/>');
  });

  it("writes a sheet the user named into a section that stated none, ahead of the margins", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(BARE_SECTION));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1800, bottomTwips: 1440, leftTwips: 1800, landscape: false, paperTwips: { width: 9978, height: 14173 } });
    expect(bodyXml(writeDocx(doc))).toContain('<w:sectPr><w:pgSz w:w="9978" w:h="14173"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/><w:cols w:space="720"/></w:sectPr>');
  });

  it("resizes the section that hangs off a paragraph, and only that one", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION, "paragraph"));
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: false, paperTwips: { width: 14173, height: 20013 } });
    const xml = bodyXml(writeDocx(doc));
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="14173" w:h="20013"/>');
    expect(xml.indexOf('<w:pgSz w:w="14173"')).toBeLessThan(xml.indexOf("</w:pPr>"));
    expect((xml.match(/<w:sectPr/g) ?? []).length).toBe(1);
    expect(WordDocument.openBytes(writeDocx(doc)).sections[0]!.pageHeightTwips).toBe(20013);
  });

  it("sizes the page a document created here is drawn on", () => {
    const fresh = WordDocument.blank();
    fresh.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: false, paperTwips: { width: 14173, height: 20013 } });
    const xml = bodyXml(writeDocx(fresh));
    expect(tagOf(xml, "pgSz")).toBe('<w:pgSz w:w="14173" w:h="20013"/>');
    const again = WordDocument.openBytes(writeDocx(fresh));
    expect(again.sections[0]!.pageWidthTwips).toBe(14173);
    expect(page(renderWord(again))).toContain(`--page-w:${twipsToPx(14173)}px`);
  });

  it("carries the sheet in the undo step, with the margins", () => {
    const doc = WordDocument.openBytes(makePageSetupDocx(WPS_SECTION));
    const before = doc.pageGeometryState();
    doc.setPageSetup({ topTwips: 1440, rightTwips: 1440, bottomTwips: 1440, leftTwips: 1440, landscape: false, paperTwips: { width: 12240, height: 15840 } });
    expect(doc.sections[0]!.pageWidthTwips).toBe(12240);
    doc.restorePageGeometry(before);
    expect(doc.sections[0]!.pageWidthTwips).toBe(11906);
    expect(doc.pageGeometryState().bodySectPr).toBe(before.bodySectPr);
    expect(doc.edited).toBe(false);
  });
});
