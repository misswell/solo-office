import { OpcPackage, parseOrderedXml, rawAttrs, rawChildren, rawFind, tagLocalOf } from "./opc.js";

/**
 * The theme's two font slots, and the one question this answers: a `w:rFonts` that names no face of its own but
 * points at `w:asciiTheme="minorHAnsi"` did state a face — it just stated it two parts away.
 *
 * `w:asciiTheme`/`w:hAnsiTheme`/`w:eastAsiaTheme`/`w:cstheme` carry an `ST_ThemeFont` token (`majorAscii`,
 * `minorEastAsia`, …): a *slot* of the package's theme part, whose `a:fontScheme` holds `a:majorFont` and
 * `a:minorFont`, each with `a:latin`/`a:ea`/`a:cs` faces and a list of `a:font` fallbacks keyed by writing system.
 * The token says which slot and which face; the token's own name says nothing about which of the three faces to
 * read — `majorAscii`/`majorHAnsi` are the Latin face, `*EastAsia` the East Asian one, `*Bidi` the complex-script
 * one, and that is the only mapping the token itself supports.
 *
 * Over the 201 documents this machine's WPS wrote (`.scratch/theme-font-audience.ts`, `.scratch/theme-font-slots.ts`,
 * `.scratch/theme-font-latin-audience.ts`): every one carries a theme part with `a:majorFont`/`a:minorFont`, and
 * every one writes a face into `a:latin` of both slots — while `a:ea` is empty in 200 of 201 and `a:cs` in all 201.
 * `w:rFonts` uses that split asymmetrically: 250 of the 264 theme references in `word/styles.xml` point at a Latin
 * slot (and 58 of them are the document floor, `w:docDefaults`), so those resolve here, from bytes the file wrote;
 * the 5,406 in `word/document.xml` point at an East Asian slot, whose face the theme leaves empty. 8,962 body
 * paragraphs sit under those 58 floors and state no literal face at all — today they draw with the page's own
 * default face, which is a number nobody wrote.
 *
 * An empty `a:ea` is not read as "fall back to whichever Han script this document looks like". The theme does
 * carry `a:font` entries per writing system, and picking one needs the text's script; where the *run* states its
 * East Asian language (`w:lang@w:eastAsia`), that is the file speaking and `face()` takes it — 442 of the 5,537
 * theme-referencing `w:rPr` blocks in document.xml, and 149 of styles.xml's 254. Elsewhere the answer would be the
 * machine's locale rather than the document's, so this returns `undefined` and the caller keeps the fallback it
 * already had. Which face a real WPS draws for the rest is 未取证.
 */
export class ThemeFonts {
  /** Nothing to resolve against: every `face()` call answers `undefined`, so the reader keeps its old behaviour. */
  static readonly none = new ThemeFonts(new Map());

  /** slot (`major` | `minor`) → the three faces plus the per-script fallbacks. */
  private constructor(private readonly slots: Map<string, { latin?: string; ea?: string; cs?: string; byScript: Map<string, string> }>) {}

  /**
   * The theme part the document points at, taken off its own relationship when one exists (a theme can be an
   * external target, and a package can hold more than one). A document read from a package with no theme — a
   * freshly built one — gets `none`.
   */
  static fromPackage(pkg: OpcPackage): ThemeFonts {
    let path: string | undefined;
    for (const rel of pkg.rels("word/document.xml").values()) if (rel.type.endsWith("/theme")) {
      path = pkg.resolve("word/document.xml", rel);
      break;
    }
    if (!path || !pkg.parts.has(path)) path = [...pkg.parts.keys()].find((p) => /^word\/theme\/[^/]+\.xml$/i.test(p));
    const xml = path ? pkg.text(path) : undefined;
    if (!xml) return ThemeFonts.none;
    const scheme = rawFind(parseOrderedXml(xml), "fontScheme");
    if (!scheme) return ThemeFonts.none;
    const slots = new Map<string, { latin?: string; ea?: string; cs?: string; byScript: Map<string, string> }>();
    for (const entry of rawChildren(scheme)) {
      const local = tagLocalOf(entry);
      const slot = local === "majorFont" ? "major" : local === "minorFont" ? "minor" : undefined;
      if (!slot) continue;
      const face: { latin?: string; ea?: string; cs?: string; byScript: Map<string, string> } = { byScript: new Map() };
      for (const child of rawChildren(entry)) {
        const name = tagLocalOf(child);
        const typeface = rawAttrs(child)["@_typeface"];
        if (name === "latin" || name === "ea" || name === "cs") {
          // An empty `typeface=""` is the shape these themes use for "no face here", not a face called "".
          if (typeface) face[name] = typeface;
        } else if (name === "font") {
          const script = rawAttrs(child)["@_script"];
          if (script && typeface) face.byScript.set(script, typeface);
        }
      }
      slots.set(slot, face);
    }
    return slots.size === 0 ? ThemeFonts.none : new ThemeFonts(slots);
  }

  /**
   * The face an `ST_ThemeFont` token names, or `undefined` when the theme states none. `script` is only consulted
   * for a slot whose own face is empty: `a:font` is a per-writing-system list, and asking it without a script would
   * be this module choosing one.
   */
  face(token: string | undefined, script?: string): string | undefined {
    const m = /^(major|minor)(Ascii|HAnsi|EastAsia|Bidi)$/.exec(token ?? "");
    if (!m) return undefined;
    const slot = this.slots.get(m[1].toLowerCase());
    if (!slot) return undefined;
    const kind = m[2] === "EastAsia" ? "ea" : m[2] === "Bidi" ? "cs" : "latin";
    return slot[kind] ?? (script === undefined ? undefined : slot.byScript.get(script));
  }
}

/**
 * `w:rFonts@w:hint` and `w:lang@w:eastAsia` are how a run says which writing system its text is in, and the theme's
 * `a:font` list is keyed by the same thing — but with its own tags (`Hans`, `Hant`, `Jpan`, `Hang`). Only these four
 * are translated, and only from an explicit East Asian language: anything else answers `undefined`, because the
 * document did not say.
 */
export function themeFontScript(eastAsiaLang: string | undefined): string | undefined {
  if (!eastAsiaLang) return undefined;
  const v = eastAsiaLang.replace(/_/g, "-");
  if (/^zh-(?:CN|SG)$/i.test(v)) return "Hans";
  if (/^zh-(?:TW|HK|MO)$/i.test(v)) return "Hant";
  if (/^ja/i.test(v)) return "Jpan";
  if (/^ko/i.test(v)) return "Hang";
  return undefined;
}
