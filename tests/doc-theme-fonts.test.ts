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
 * the 201 documents this machine's WPS wrote are stated that way, and the rule that shape follows is per slot —
 * across every part of those files, **0 elements state a literal face and a theme reference for the same slot**,
 * while 5,260 state one slot literally and point another at the theme (`.scratch/theme-slot-pair.ts`). These tests
 * pin which references the file lets this reader answer, and — just as deliberately — which it does not, then which
 * half of the element a face edit is allowed to rewrite.
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
    // §17.3.2.11 ranks neither attribute over the other, and per slot 0 of the 55,585 tags in these files write
    // both (`.scratch/theme-slot-pair.ts`), so this is pinned as a rule for a shape the corpus never reaches.
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

/**
 * The other half of the same attribute set: what a face edit may leave behind. Re-spelling a moved slot into an
 * element that still carried that slot's `w:*Theme` was the one shape the corpus never reaches — simulating that
 * edit over the 201 files produced 2,808 same-slot collisions from 2,804 themed runs
 * (`.scratch/theme-both-on-edit.ts`), against a corpus count of 0 (`.scratch/theme-slot-pair.ts`). A reference is
 * that slot's own sentence about where its face comes from, so moving the slot retires the sentence and the slots
 * nobody moved keep theirs.
 */
describe("docx theme fonts under a face edit", () => {
  /**
   * The `w:rFonts` of the run in `needle`'s paragraph. That paragraph carries no `w:pPr`, so its first `w:rFonts`
   * is the run's own and not the paragraph mark's.
   */
  const fontsOf = (body: string, needle = "整行交给主题"): string =>
    /<w:rFonts\b[^>]*>/.exec(body.split("</w:p>").find((p) => p.includes(needle)) ?? "")?.[0] ?? "";

  const edited = (mutate: (r: Inline) => void): { body: string; out: Uint8Array } => {
    const doc = open(makeThemeDocx());
    mutate(runAt(doc, "整行交给主题"));
    const out = writeDocx(doc);
    return { body: part(out, "word/document.xml"), out };
  };

  it("retires the Latin slot's theme reference when the model moves that face", () => {
    const { body, out } = edited((r) => (r.font = "SimSun"));
    const tag = fontsOf(body);
    expect(tag).toContain('w:ascii="SimSun"');
    expect(tag).not.toContain("asciiTheme");
    expect(tag).not.toContain("hAnsiTheme");
    expect(tag).toContain('w:eastAsiaTheme="minorEastAsia"');
    expect(tag).toContain('w:cstheme="minorBidi"');
    // What comes back out of the saved bytes is what the page then draws: the named face for the moved slot, and
    // the theme's Han face still resolved for the slot the edit never touched.
    const again = runAt(open(out), "整行交给主题");
    expect(again.font).toBe("SimSun");
    expect(again.fontEastAsia).toBe("仿宋");
  });

  it("retires only the East Asian reference when the model moves that face", () => {
    const { body, out } = edited((r) => (r.fontEastAsia = "宋体"));
    const tag = fontsOf(body);
    expect(tag).toContain('w:eastAsia="宋体"');
    expect(tag).not.toContain("eastAsiaTheme");
    expect(tag).toContain('w:asciiTheme="minorHAnsi"');
    expect(tag).toContain('w:hAnsiTheme="minorHAnsi"');
    const again = runAt(open(out), "整行交给主题");
    expect(again.font).toBe("Calibri");
    expect(again.fontEastAsia).toBe("宋体");
  });

  it("leaves the moved slot stating nothing of its own when the change is refused", () => {
    // A refusal already took the literal name away; leaving the reference standing would hand the slot back to the
    // theme by a route the file never wrote, so the slot resolves from the docDefaults floor instead — which is the
    // same fallback any refused field gets. The other three slots are untouched by a Latin refusal.
    const { body, out } = edited((r) => {
      r.font = "SimSun";
      r.font = undefined;
    });
    const tag = fontsOf(body);
    expect(tag).not.toContain("w:ascii=");
    expect(tag).not.toContain("asciiTheme");
    expect(tag).toContain('w:eastAsiaTheme="minorEastAsia"');
    expect(tag).toContain('w:cstheme="minorBidi"');
    expect(runAt(open(out), "整行交给主题").font).toBe("Calibri");
  });

  it("writes no same-slot collision anywhere in the document it saved", () => {
    // The guard the corpus count is: same slot, both attributes, in one element. A face edit on every themed run
    // must retire the fixture's one intentional `字面先赢` pair and add none.
    // `[^"]+` and not `[^"]`: the shorter form demands a one-character value and reports every real face as absent.
    const SLOTS: [string, string][] = [
      ["w:ascii", "w:asciiTheme"],
      ["w:hAnsi", "w:hAnsiTheme"],
      ["w:eastAsia", "w:eastAsiaTheme"],
      ["w:cs", "w:cstheme"],
    ];
    const stated = (tag: string, name: string): boolean => new RegExp(`${name}="[^"]+`).test(tag);
    const collisions = (xml: string): number =>
      (xml.match(/<w:rFonts\b[^>]*>/g) ?? []).filter((t) => SLOTS.some(([lit, them]) => stated(t, lit) && stated(t, them))).length;
    const bytes = makeThemeDocx();
    expect(collisions(part(bytes, "word/document.xml"))).toBe(1);
    const doc = open(bytes);
    let moved = 0;
    for (const p of paragraphs(doc))
      for (const r of p.inlines)
        if (r.rPrSlots?.some((s) => s.tag === "rFonts" && s.xml.includes("Theme"))) {
          r.font = "SimSun";
          r.fontEastAsia = "宋体";
          moved++;
        }
    // Five runs in the fixture hand a slot to the theme; the count is here so a fixture edit cannot empty this test.
    expect(moved).toBe(5);
    // Including `字面先赢`, whose same-slot pair is retired by the same rule — the writer leaves no element that
    // states a slot twice, so nothing downstream has to pick a precedence §17.3.2 does not give.
    expect(collisions(part(writeDocx(doc), "word/document.xml"))).toBe(0);
  });
});
