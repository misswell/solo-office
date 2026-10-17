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
 * Over the authorized corpus — 46 WPS-bundled templates, `.scratch/authorized-families.mjs` (theme part) and
 * `.scratch/theme-slot-pair-authorized.mjs` (`w:rFonts` slots): **46 of 46 carry a theme part holding both
 * `a:majorFont` and `a:minorFont`**, and across the 92 font slots (46 × major+minor) all 92 write a face into
 * `a:latin`, while `a:ea` is empty in 84 (8 write one) and `a:cs` is empty in all 92. So the asymmetry this class
 * resolves is measured here, not inherited: a themed Latin reference has a face to resolve to, a themed East Asian
 * reference usually does not.
 * The per-part split the old sentence quoted (250 of 264 `w:rFonts` theme references in `word/styles.xml` pointing
 * at Latin, 5,406 in `word/document.xml` pointing at East Asian, 8,962 body paragraphs under 58 floors) was counted
 * over the quarantined 201-file cohort and is **待重证** on the authorized list; what is re-measured is that
 * `w:rFonts` uses the split asymmetrically — 1,957 of the 14,642 `w:rFonts` in the authorized files carry a theme
 * slot, 11,905 carry a literal face, 650 carry both on different slots, and **0 state a literal face and a theme
 * reference in the same slot**.
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
