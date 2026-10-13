import { describe, expect, it } from "vitest";
import { WordDocument, cmToTwips, twipsToCm, twipsToPx } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { BARE_SECTION, WPS_LANDSCAPE_SECTION, WPS_SECTION, WPS_SECTION_NO_DISTANCES, makePageSetupDocx } from "./helpers/make-page-setup.js";
import { makeFullPageDistanceDocx, makeTwoSectionDocx } from "./helpers/make-header-footer.js";

const bodyXml = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
/** One element's own opening tag, so a test can name the bytes a change wrote rather than a re-read value. */
const tagOf = (xml: string, name: string): string => new RegExp(`<w:${name}[^>]*/?>`).exec(xml)?.[0] ?? "";
/** The page box's own `style`, which is where the geometry a user typed has to show up. */
const page = (html: string): string => /<div class="page" style="([^"]*)"/.exec(html)?.[1] ?? "";

describe("docx 页面设置：页边距/距离/装订线/方向", () => {
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
});
