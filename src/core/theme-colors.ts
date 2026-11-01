import { OpcPackage, parseOrderedXml, rawAttrs, rawChildren, rawFind, tagLocalOf } from "./opc.js";
import { themePartPath } from "./theme-fonts.js";

/** A theme style entry reduced to what this reader can state: the colour slot or the colour, and a line's width. */
interface StylePaint {
  /** `true` when the entry is a solid of `phClr` — the placeholder the referencing shape colours. */
  placeholder: boolean;
  /** The colour the entry names itself, when it names one. */
  color?: string;
  /** `a:ln@w`, EMU. Only a line entry carries one. */
  width?: number;
}

/**
 * The colour half of the theme part: `a:clrScheme`, `a:fillStyleLst` and `a:lnStyleLst`.
 *
 * A DrawingML shape may state its paint outright (`a:solidFill` in its own `a:spPr`), or borrow it from the theme
 * through `wps:style` — `a:fillRef idx="1"` carrying an `a:schemeClr` says "the first entry of `a:fillStyleLst`,
 * painted with this theme colour". Both spellings are in the authorized corpus, so a reader that honours only the
 * first draws the borrowed half of a diagram as nothing.
 *
 * The style lists are addressed by position (`idx` is 1-based). Only the plain forms resolve here: a solid fill with
 * no modifiers, whose colour is either `phClr` or a colour the entry itself names. A gradFill, a pattern, or a `phClr`
 * carrying `a:shade`/`a:tint`/`a:satMod` answers `undefined`, because expanding those needs the colour-transform
 * arithmetic the file never states in words this model owns — and an invented approximation would report a colour the
 * document did not write.
 *
 * Audience, over the authorized corpus (`.scratch/wgp-census.mjs`: the 21 WPS-bundled Word templates — 40 files
 * deduped by basename — plus the 6 Word files of the 114-file list): **11 枚** `wps:style` in **4 件**. Only
 * `business plan.dotm` spells them inline — 7 枚 on its 38-shape `wpg:wgp`, the one drawing this reader draws,
 * where 31 枚 state their paint in their own `a:spPr` and 0 枚 state neither. The other 4 枚 sit on `wp:anchor`
 * shapes in four footers, and this release does not place anchored shapes at all, so they are borrowed paint on a
 * figure the page leaves undrawn. The 114-file cohort's Word parts hold 0 shapes of either kind.
 */
export class ThemeColors {
  /** Nothing to resolve against: every question answers `undefined`, so a caller draws only stated paint. */
  static readonly none = new ThemeColors(new Map(), [], []);

  private constructor(
    private readonly scheme: Map<string, string>,
    private readonly fills: (StylePaint | undefined)[],
    private readonly lines: (StylePaint | undefined)[],
  ) {}

  static fromPackage(pkg: OpcPackage): ThemeColors {
    const path = themePartPath(pkg);
    const xml = path ? pkg.text(path) : undefined;
    if (!xml) return ThemeColors.none;
    const root = parseOrderedXml(xml);
    const scheme = new Map<string, string>();
    for (const slot of rawChildren(rawFind(root, "clrScheme"))) {
      const name = tagLocalOf(slot);
      const hex = slotColor(slot);
      if (name && hex) scheme.set(name, hex);
    }
    const fills = rawChildren(rawFind(root, "fillStyleLst")).map((entry) => stylePaint(entry, scheme));
    const lines = rawChildren(rawFind(root, "lnStyleLst")).map((entry) => {
      const paint = rawChildren(entry).find((kid) => paintTags.has(tagLocalOf(kid) ?? ""));
      const resolved = paint ? stylePaint(paint, scheme) : undefined;
      const w = Number(rawAttrs(entry)["@_w"]);
      return resolved ? { ...resolved, width: Number.isFinite(w) ? w : undefined } : undefined;
    });
    return scheme.size || fills.some(Boolean) || lines.some(Boolean) ? new ThemeColors(scheme, fills, lines) : ThemeColors.none;
  }

  /** The colour an `a:schemeClr@val` names (`dk1`, `accent2`, …), or `undefined` when the theme states none. */
  color(val: string): string | undefined {
    return this.scheme.get(val);
  }

  /** `a:fillStyleLst[idx-1]`, or `undefined` when the theme states no such entry or states one this reader cannot expand. */
  fillStyle(idx: number): StylePaint | undefined {
    return this.fills[idx - 1];
  }

  /** `a:lnStyleLst[idx-1]`, with the width that entry states. */
  lineStyle(idx: number): StylePaint | undefined {
    return this.lines[idx - 1];
  }
}

/** The paint-bearing children of an `a:spPr` or a style entry. `noFill` is one: it says *none*, which is not silence. */
const paintTags = new Set(["solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "noFill"]);

/**
 * The colour a `a:dk1`/`a:lt1`/`a:accentN` slot states. `a:sysClr` names an OS colour and carries `lastClr` — the
 * value the file saw where it was written, which is the only colour this page can draw without asking the machine.
 */
function slotColor(slot: any): string | undefined {
  const kid = rawChildren(slot)[0];
  const name = tagLocalOf(kid);
  if (name !== "srgbClr" && name !== "sysClr") return undefined;
  return hexColor(rawAttrs(kid)[name === "srgbClr" ? "@_val" : "@_lastClr"]);
}

/** A six-digit colour written as the file writes it, or `undefined` for anything that is not one. */
export function hexColor(value: unknown): string | undefined {
  return typeof value === "string" && /^[0-9a-fA-F]{6}$/.test(value) ? `#${value}` : undefined;
}

/**
 * A style entry reduced to a plain solid, or `undefined`. `scheme` is the `a:clrScheme` already read, which is
 * what names a colour an entry points at by slot.
 */
function stylePaint(paint: any, scheme: Map<string, string>): StylePaint | undefined {
  if (tagLocalOf(paint) !== "solidFill") return undefined;
  const kid = rawChildren(paint)[0];
  const name = tagLocalOf(kid);
  // A modifier under the colour (`a:shade`, `a:tint`, `a:satMod`, …) is arithmetic this reader does not run.
  if (rawChildren(kid).length) return undefined;
  const attrs = rawAttrs(kid);
  if (name === "schemeClr") {
    const val = String(attrs["@_val"] ?? "");
    if (!val) return undefined;
    // `phClr` is the placeholder the referencing shape's own `a:schemeClr` colours.
    return val === "phClr" ? { placeholder: true } : { placeholder: false, color: scheme.get(val) };
  }
  if (name === "srgbClr") return { placeholder: false, color: hexColor(attrs["@_val"]) };
  if (name === "sysClr") return { placeholder: false, color: hexColor(attrs["@_lastClr"]) };
  return undefined;
}
