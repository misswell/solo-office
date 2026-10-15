import { describe, expect, it } from "vitest";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeThemeDocx, themeXml } from "./helpers/make-theme.js";

const pkg = (bytes: Uint8Array): OpcPackage => OpcPackage.from(bytes);
const part = (bytes: Uint8Array, name: string): string => text(pkg(bytes).bytes(name)!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const paraAt = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text).join("").includes(needle))!;
const runAt = (doc: WordDocument, needle: string): Inline => paraAt(doc, needle).inlines[0]!;
/** Save without touching anything, and read the body back out of the bytes that came out. */
const saved = (bytes: Uint8Array): string => {
  const out = writeDocx(open(bytes));
  return part(out, "word/document.xml");
}

/**
 * A `w:rFonts` can name a face without naming one: the four `w:*Theme` attributes hold an `ST_ThemeFont` token,
 * which is a slot of the package's theme part rather than a font. 5,702 of the 54,200 `w:rFonts` in the bodies of
 * the 201 documents this machine's WPS wrote are stated that way, and not one of them also states a literal face
 * (`.scratch/theme-font-audience.ts`), so a reader that ignores them draws those runs with a face nobody wrote.
 * These tests pin which references the file lets this reader answer, and — just as deliberately — which it does
 * not.
 */
describe("docx theme fonts", () => {
  it("reads the document floor's face out of the theme when docDefaults names no literal one", () => {
    // 58 of the 201 put their whole run floor there, and in those files 8,962 body paragraphs state no face at
    // all (`.scratch/theme-font-latin-audience.ts`) — the paragraph below is one of them.
    expect(runAt(open(makeThemeDocx()), "地板给的脸").font).toBe("Calibri");
  });

  it("resolves a face a paragraph borrows from a style that points at the theme", () => {
    // styles.xml states 264 theme-only `w:rFonts`, 250 of them on a Latin slot (`.scratch/theme-font-slots.ts`).
    const run = runAt(open(makeThemeDocx()), "样式给的脸");
    expect(run.font).toBe("Cambria");
    // The same element's `w:eastAsiaTheme="majorEastAsia"` names the slot whose `a:ea` is empty in 200 of 201
    // files, and this style states no language to pick a script from, so nothing is invented for it.
    expect(run.fontEastAsia).toBeUndefined();
  });

  it("keeps a literal face ahead of a theme reference on the same element", () => {
    // §17.3.2.11 ranks neither attribute over the other and 0 of the 201 files write both, so this is pinned as a
    // rule: the face the file named is the face the page draws.
    expect(runAt(open(makeThemeDocx()), "字面先赢").font).toBe("Times New Roman");
  });

  it("answers an East Asian reference from the theme where the run states the language the theme lists", () => {
    // 442 of the 5,537 theme-referencing `w:rPr` blocks in these bodies state a `w:lang@w:eastAsia` whose script
    // the theme's `a:font` list does carry (`.scratch/theme-font-resolvable.ts`): the file said which one.
    expect(runAt(open(makeThemeDocx()), "汉字有语言").fontEastAsia).toBe("仿宋");
  });

  it("leaves the East Asian slot unread where the theme is empty and the run states no language", () => {
    const doc = open(makeThemeDocx());
    expect(runAt(doc, "汉字无语言").fontEastAsia).toBeUndefined();
    // `w:hint` is not a stated language: 4,677 of those blocks lean on it, and answering it here would be this
    // reader picking a script the document never named. Which face a real WPS draws for them is 未取证.
    expect(runAt(doc, "标记只说hint").fontEastAsia).toBeUndefined();
  });

  it("resolves the East Asian slot from the theme where the theme itself names a face", () => {
    // 1 of the 201 files writes `a:ea` with a face; there the reference is fully stated by the package.
    expect(runAt(open(makeThemeDocx({ theme: themeXml(true) })), "汉字无语言").fontEastAsia).toBe("等线");
  });

  it("reads a paragraph mark's faces through the theme", () => {
    const mark = paraAt(open(makeThemeDocx()), "标记的脸");
    expect(mark.markFont).toBe("Cambria");
    expect(mark.markFontEastAsia).toBe("华文中宋");
  });

  it("reads a numbering level's mark face through the theme", () => {
    // The corpus's 421 `w:rFonts` in numbering.xml state none of this, so the rule is what is pinned here.
    expect(open(makeThemeDocx()).numberingTable().find((n) => n.numId === 1)?.levels[0]?.font).toBe("Calibri");
  });

  it("reads nothing when the package has no theme part to point at", () => {
    const doc = open(makeThemeDocx({ theme: null }));
    expect(runAt(doc, "地板给的脸").font).toBeUndefined();
    expect(runAt(doc, "汉字有语言").fontEastAsia).toBeUndefined();
  });

  it("saves an untouched document with its theme references and its theme part, and no invented face", () => {
    const bytes = makeThemeDocx();
    const before = part(bytes, "word/document.xml");
    const body = saved(bytes);
    const out = writeDocx(open(bytes));
    // The reference *is* what the file said, so it is what goes back — the same 5,966 of the 201 corpus files
    // survived a save before this reader could resolve them (`.scratch/theme-font-audience.ts`). A save that
    // traded a `w:*Theme` for a face name would rewrite the file's sentence about where the face comes from, so
    // every `w:rFonts` tag is compared as written rather than counted.
    const fonts = (xml: string): string[] => [...xml.matchAll(/<[a-zA-Z][-\w.]*:rFonts\b[^>]*/g)].map((m) => m[0]);
    expect(fonts(body)).toEqual(fonts(before));
    expect(body.match(/w:ascii="Calibri"/g) ?? []).toHaveLength(0);
    expect(part(out, "word/styles.xml")).toContain('w:asciiTheme="minorHAnsi"');
    expect(pkg(out).bytes("word/theme/theme1.xml")).toEqual(pkg(bytes).bytes("word/theme/theme1.xml"));
    // And the faces still resolve after the round trip, so a save has not quietly taken them away.
    const again = open(out);
    expect(runAt(again, "地板给的脸").font).toBe("Calibri");
    expect(runAt(again, "汉字有语言").fontEastAsia).toBe("仿宋");
  });
});
