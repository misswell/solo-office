import { OpcPackage, asArray, children as kids, parseOrderedXml, rawAttrs, rawChildren, rawFind, rawFindAll, relsPathOf, tagLocalOf, tagOf, type Rel } from "./opc.js";
import { dataUrl, IMAGE_EXT_BY_MIME, IMAGE_MIME_BY_EXT } from "./media.js";
import { BLANK_LAYOUTS } from "./pptx-blank.js";
import type { Inline } from "./docx.js";

/** `a:bodyPr`: where a shape's text sits inside its box. */
export interface TextBody {
  /** `a:bodyPr@lIns` and friends, in EMU: the text's distance from each box edge. */
  lInsEmu?: number;
  tInsEmu?: number;
  rInsEmu?: number;
  bInsEmu?: number;
  /** `a:bodyPr@anchor`: "t", "ctr" or "b". */
  anchor?: string;
  /** `a:bodyPr@wrap="none"`: lines do not wrap at the box edge. */
  noWrap?: boolean;
}

/** A placeholder the slide inherits from its layout, geometry in slide pixels. */
export interface LayoutPlaceholder {
  type?: string;
  idx?: number;
  x: number;
  y: number;
  width: number;
  height: number;
  /** Level-1 character size the layout's `a:lstStyle` states, in points. */
  sizePt?: number;
  /** Level-1 paragraph alignment, in the editor's words. */
  align?: string;
  /** The layout's `a:bodyPr`, with the master's own where the layout leaves one out. */
  textBody?: TextBody;
  /** Level-1 `marL`/`indent`, in EMU; the hanging indent a bulleted line starts from. */
  marLEmu?: number;
  indentEmu?: number;
  /** Level-1 bullet character, or `<a:buNone/>` suppressing the master's bullet. */
  bullet?: string;
  noBullet?: boolean;
  /** Level-1 `a:lnSpc/a:spcPct`: line spacing as a percentage of single. */
  lineSpacingPct?: number;
  /** Level-1 `a:spcBef/a:spcPts`, in points. */
  spaceBeforePt?: number;
  /** Level-1 `a:spcBef/a:spcPct`, as a percentage of the line height. */
  spaceBeforePct?: number;
  /** Level-1 `a:spcAft/a:spcPts`, in points. */
  spaceAfterPt?: number;
  /** Level-1 `a:spcAft/a:spcPct`, as a percentage of the line height. */
  spaceAfterPct?: number;
}

/** One level of a text style (`p:bodyStyle` or a shape's `a:lstStyle`). */
export interface BodyLevel {
  sizePt?: number;
  align?: string;
  marLEmu?: number;
  indentEmu?: number;
  bullet?: string;
  noBullet?: boolean;
  /** `a:lnSpc/a:spcPct`: line spacing as a percentage of single. */
  lineSpacingPct?: number;
  /** `a:spcBef/a:spcPts`, in points. */
  spaceBeforePt?: number;
  /** `a:spcBef/a:spcPct`, as a percentage of the line height. */
  spaceBeforePct?: number;
  /** `a:spcAft/a:spcPts`, in points. */
  spaceAfterPt?: number;
  /** `a:spcAft/a:spcPct`, as a percentage of the line height. */
  spaceAfterPct?: number;
}

/**
 * The WPS blank template's `p:bodyStyle` levels: step-in 228600 EMU per level with a
 * hanging -228600 indent, sizes 18/16/16/14/14/18/18/18/18 and ●/●/●/Wingdings-0x9F/•…
 * A blank deck built here writes these, so a blank `Presentation` starts from them.
 */
export const WPS_BODY_LEVELS: BodyLevel[] = [18, 16, 16, 14, 14, 18, 18, 18, 18].map((sizePt, i) => ({
  sizePt,
  marLEmu: 228600 + 457200 * i,
  indentEmu: -228600,
  bullet: ["●", "●", "●", "\u009f", "•", "•", "•", "•", "•"][i],
}));

export interface Shape {
  kind: "shape" | "picture" | "table" | "group";
  name: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  geometry?: string;
  fill?: string;
  lineColor?: string;
  lineWidth?: number;
  paragraphs?: SlideParagraph[];
  src?: string;
  rows?: SlideCell[][];
  children?: Shape[];
  placeholder?: string;
  /** `p:ph@idx` of the placeholder this shape fills, when it fills one. */
  placeholderIdx?: number;
  /** A text box the editor inserted: no preset geometry, no fill or line of its own. */
  textBox?: boolean;
  /** Package part of an image inserted in the editor; a save writes it out. */
  mediaPath?: string;
  /** Relationship id the writer assigns to an inserted picture inside its slide part. */
  relId?: string;
  /** The part's own ordered nodes for this element, patched in place when the shape is edited. */
  raw?: any;
  /** `a:txBody/a:lstStyle` levels: what this shape's text inherits before the layout and master speak. */
  textLevels?: BodyLevel[];
  /** The shape's own `a:bodyPr`: insets, vertical anchor and wrapping of its text. */
  textBody?: TextBody;
  /** `a:bodyPr/a:normAutofit@fontScale`: the size to draw at, as a percent of the stated one. */
  fontScalePct?: number;
  /** `a:bodyPr/a:normAutofit@lnSpcReduction`: how much of the line height the fit gives back. */
  lineSpacingReductionPct?: number;
  /** Text was edited; a save rebuilds this shape's text body from the model. */
  textDirty?: boolean;
}

export interface SlideParagraph {
  align?: string;
  level: number;
  bullet?: string;
  runs: Inline[];
  /** `a:pPr/a:lnSpc/a:spcPct`: line spacing as a percentage of single. */
  lineSpacingPct?: number;
  /** `a:pPr/a:spcBef/a:spcPts`, in points. */
  spaceBeforePt?: number;
  /** `a:pPr/a:spcBef/a:spcPct`, as a percentage of the line height. */
  spaceBeforePct?: number;
  /** `a:pPr/a:spcAft/a:spcPts`, in points. */
  spaceAfterPt?: number;
  /** `a:pPr/a:spcAft/a:spcPct`, as a percentage of the line height. */
  spaceAfterPct?: number;
}

export interface SlideCell {
  text: string;
  runs: Inline[];
  fill?: string;
}

export interface Slide {
  index: number;
  name: string;
  background?: string;
  shapes: Shape[];
  notes?: string;
  /** Part path inside the package, e.g. `ppt/slides/slide1.xml`. */
  partPath: string;
  notesPath?: string;
  /** `p:sldId` id and relationship id in `ppt/presentation.xml`. */
  sldId: string;
  relId: string;
  /** Ordered nodes of `p:sld`, patched in place by a save. */
  tree: any[];
  /** The layout part this slide inherits from; a save writes the relationship to it. */
  layoutPath?: string;
  /** Placeholders inherited from the layout; the editor draws a prompt where no shape fills one. */
  layoutPlaceholders?: LayoutPlaceholder[];
  /** Set by an edit; a clean slide is written back byte-identical. */
  dirty: boolean;
  notesDirty: boolean;
}

const EMU_PER_PX = 9525;

/** OOXML `algn` values in, to the editor's own left/center/right/both words. */
const ALGN_IN: Record<string, string | undefined> = { l: "left", ctr: "center", r: "right", just: "both", dist: "both" };

const find = (node: any, name: string): any => kids(node).find((c) => c.name === name)?.node;
const findAll = (node: any, name: string): any[] => kids(node).filter((c) => c.name === name).map((c) => c.node);
const attr = (node: any, name: string): string | undefined => node?.[`@_${name}`];

/**
 * One `a:lvlNpPr` (or `a:pPr`) as the editor's level paint. `a:lnSpc` holds the line
 * spacing, `a:spcBef`/`a:spcAft` the paragraph spacing; both may say points or a
 * percentage of the line height, kept apart so the ladder can pick one rung's answer.
 */
function readLevel(lvl: any): BodyLevel {
  const a = rawAttrs(lvl);
  const rpr = rawFind(rawChildren(lvl), "defRPr");
  const sz = Number(rawAttrs(rpr ?? {})["@_sz"]);
  const buChar = rawFind(rawChildren(lvl), "buChar");
  const ln = rawFind(rawChildren(lvl), "lnSpc");
  const lnPct = ln ? rawFind(rawChildren(ln), "spcPct") : undefined;
  const bef = rawFind(rawChildren(lvl), "spcBef");
  const befPts = bef ? rawFind(rawChildren(bef), "spcPts") : undefined;
  const befPct = bef ? rawFind(rawChildren(bef), "spcPct") : undefined;
  const aft = rawFind(rawChildren(lvl), "spcAft");
  const aftPts = aft ? rawFind(rawChildren(aft), "spcPts") : undefined;
  const aftPct = aft ? rawFind(rawChildren(aft), "spcPct") : undefined;
  const pct = lnPct ? Number(rawAttrs(lnPct)["@_val"]) : NaN;
  const before = befPts ? Number(rawAttrs(befPts)["@_val"]) : NaN;
  const beforePct = befPct ? Number(rawAttrs(befPct)["@_val"]) : NaN;
  const after = aftPts ? Number(rawAttrs(aftPts)["@_val"]) : NaN;
  const afterPct = aftPct ? Number(rawAttrs(aftPct)["@_val"]) : NaN;
  return {
    sizePt: Number.isFinite(sz) && sz > 0 ? sz / 100 : undefined,
    align: a["@_algn"] ? ALGN_IN[a["@_algn"]] : undefined,
    marLEmu: a["@_marL"] !== undefined ? Number(a["@_marL"]) : undefined,
    indentEmu: a["@_indent"] !== undefined ? Number(a["@_indent"]) : undefined,
    bullet: buChar ? rawAttrs(buChar)["@_char"] : undefined,
    noBullet: !!rawFind(rawChildren(lvl), "buNone") || undefined,
    // Zero is a declared "no spacing", which draws the same as silence; the type's
    // hundredths-of-a-point and thousandths-of-a-percent are scaled here once.
    lineSpacingPct: Number.isFinite(pct) && pct > 0 ? pct / 1000 : undefined,
    spaceBeforePt: Number.isFinite(before) && before > 0 ? before / 100 : undefined,
    spaceBeforePct: Number.isFinite(beforePct) && beforePct > 0 ? beforePct / 1000 : undefined,
    spaceAfterPt: Number.isFinite(after) && after > 0 ? after / 100 : undefined,
    spaceAfterPct: Number.isFinite(afterPct) && afterPct > 0 ? afterPct / 1000 : undefined,
  };
}

/** The nine levels of a `a:lstStyle`, absent levels left empty. */
function readLevels(style: any): BodyLevel[] {
  const out: BodyLevel[] = [];
  for (let n = 1; n <= 9; n++) {
    const lvl = style ? rawFind(rawChildren(style), `lvl${n}pPr`) : undefined;
    out.push(lvl ? readLevel(lvl) : {});
  }
  return out;
}

const hasLevel = (l: BodyLevel): boolean =>
  l.sizePt !== undefined ||
  l.align !== undefined ||
  l.marLEmu !== undefined ||
  l.indentEmu !== undefined ||
  l.bullet !== undefined ||
  l.noBullet !== undefined ||
  l.lineSpacingPct !== undefined ||
  l.spaceBeforePt !== undefined ||
  l.spaceBeforePct !== undefined ||
  l.spaceAfterPt !== undefined ||
  l.spaceAfterPct !== undefined;


export class Presentation {
  slides: Slide[] = [];
  slideWidthEmu = 12192000;
  slideHeightEmu = 6858000;
  /** The package the deck was read from; a writer copies every part it does not own. */
  src?: OpcPackage;
  /** Part path of `ppt/presentation.xml`. */
  presPath = "ppt/presentation.xml";
  /** Ordered nodes of `p:presentation`, patched in place when the slide list changes. */
  presTree: any[] = [];
  /** Parts a save must drop (slides the user deleted, with their notes). */
  removedParts = new Set<string>();
  /** True once the slide list changed, which makes a save rewrite `ppt/presentation.xml`. */
  presDirty = false;
  /** Layout a new slide references, as a package path. */
  newSlideLayoutPath?: string;
  /** `layoutChoices()` parses every layout part's name once; the deck's layouts do not change. */
  private layoutCache?: { path: string; name: string }[];
  /** First relationship id a blank deck may hand to a slide; its presentation rels hold the rest. */
  private blankRelBase = 0;
  /**
   * `p:gridSpacing@cx` from `ppt/viewProps.xml`, in EMU. WPS writes 72008 on 69 of the
   * 77 authorized decks (6 more write 76200, 2 write 72000), so the grid pitch the
   * editor draws comes from the deck; a blank deck built here uses WPS's common value.
   */
  gridSpacingEmu = 72008;
  /**
   * Level-1 sizes of the master's `p:txStyles`, in points. A placeholder the layout leaves
   * silent inherits these; the defaults are what the WPS blank template states (36 / 18).
   */
  titleStylePt = 36;
  bodyStylePt = 18;
  /**
   * Level-1 size of the master's `p:otherStyle`, in points: the default for text in shapes
   * that are not placeholders. WPS writes 18pt in 33 of the 35 template decks and 10.35pt
   * in the other two, and leaves the size off the runs themselves.
   */
  otherStylePt = 18;
  /** Master `p:bodyStyle` levels 1–9; a body paragraph's level picks its size, indent and bullet. */
  bodyLevels: BodyLevel[] = WPS_BODY_LEVELS.map((l) => ({ ...l }));
  /** Images inserted in the editor (package part → data URL); a save writes these parts. */
  addedMedia = new Map<string, string>();
  private theme = new Map<string, string>();
  private media = new Map<string, string>();
  private pkg?: OpcPackage;
  /**
   * `p:clrMap` of the deck's master. Backgrounds name colours as `bg1`/`tx2` and the
   * map says which clrScheme slot that is; the spec defaults hold when the master is silent.
   */
  private clrMap = new Map<string, string>([
    ["bg1", "lt1"],
    ["tx1", "dk1"],
    ["bg2", "lt2"],
    ["tx2", "dk2"],
  ]);
  /** CSS background per layout/master part; every slide on that part inherits the same value. */
  private bgCache = new Map<string, string | undefined>();
  /** Master placeholders by part, the geometry a layout leaves to its master. */
  private masterPlaceholderCache = new Map<string, LayoutPlaceholder[]>();

  get slideWidthPx(): number {
    return this.slideWidthEmu / EMU_PER_PX;
  }

  get slideHeightPx(): number {
    return this.slideHeightEmu / EMU_PER_PX;
  }

  /** Grid pitch in slide pixels, for the editor's 查看网格线 overlay. */
  get gridSpacingPx(): number {
    return this.gridSpacingEmu / EMU_PER_PX;
  }

  static openBytes(data: Uint8Array): Presentation {
    const p = new Presentation();
    p.read(OpcPackage.from(data));
    return p;
  }

  /** A new deck with one blank slide and no source package; a save writes the canonical parts. */
  static blank(): Presentation {
    const p = new Presentation();
    p.presTree = parseOrderedXml(PRESENTATION_TEMPLATE);
    // The blank writer emits master, theme, presProps, viewProps and tableStyles
    // under rId1..rId5, so slides start after them.
    p.newSlideLayoutPath = "ppt/slideLayouts/slideLayout1.xml";
    p.blankRelBase = 5;
    p.addSlide();
    return p;
  }

  private read(pkg: OpcPackage): void {
    this.pkg = pkg;
    this.src = pkg;
    this.readTheme(pkg);
    this.readMedia(pkg);
    this.readMasterStyles(pkg);
    const root = find(pkg.xml("ppt/presentation.xml"), "presentation");
    this.presTree = parseOrderedXml(pkg.text("ppt/presentation.xml") ?? "");
    const sldSz = find(root, "sldSz");
    if (sldSz) {
      this.slideWidthEmu = Number(attr(sldSz, "cx") ?? this.slideWidthEmu);
      this.slideHeightEmu = Number(attr(sldSz, "cy") ?? this.slideHeightEmu);
    }
    const rels = pkg.rels("ppt/presentation.xml");
    // The grid pitch lives in the view part, not in the deck's slides; a deck without
    // that part keeps the default above.
    const viewProps = [...rels.values()].find((r) => r.type.endsWith("/viewProps"));
    if (viewProps) {
      const xml = pkg.text(pkg.resolve("ppt/presentation.xml", viewProps)) ?? "";
      const pitch = /<p:gridSpacing\b[^>]*\bcx="(\d+)"/.exec(xml);
      if (pitch) this.gridSpacingEmu = Number(pitch[1]);
    }
    // `p:sldId` carries both the numeric id and the relationship id; the
    // prefix-stripped reader merges the two attributes, so read them raw.
    const lst = rawFind(this.presTree, "sldIdLst");
    const entries = lst ? rawFindAll(rawChildren(lst), "sldId") : [];
    const ordered = entries.length
      ? entries.map((e) => rawAttrs(e))
      : [...rels.keys()].map((id) => ({ "@_id": id, "@_r:id": id }) as Record<string, string>);
    let index = 0;
    for (const entry of ordered) {
      const relId = entry["@_r:id"] ?? entry["@_id"];
      const rel = rels.get(String(relId));
      if (!rel || !rel.type.endsWith("/slide")) continue;
      const path = pkg.resolve("ppt/presentation.xml", rel);
      if (!pkg.parts.has(path)) continue;
      this.slides.push(this.readSlide(pkg, path, index++, String(entry["@_id"] ?? relId), String(relId), rel));
    }
  }

  private readSlide(pkg: OpcPackage, path: string, index: number, sldId: string, relId: string, rel: Rel): Slide {
    const root = find(pkg.xml(path), "sld");
    const rels = pkg.rels(path);
    const cSld = find(root, "cSld");
    const slide: Slide = { index, name: `第 ${index + 1} 页`, shapes: [], partPath: path, sldId, relId, tree: parseOrderedXml(pkg.text(path) ?? ""), dirty: false, notesDirty: false };

    const layoutRel = [...rels.values()].find((r) => r.type.endsWith("/slideLayout"));
    if (layoutRel) {
      const layoutPath = pkg.resolve(path, layoutRel);
      if (!this.newSlideLayoutPath) this.newSlideLayoutPath = layoutPath;
      slide.layoutPath = layoutPath;
      slide.layoutPlaceholders = this.layoutPlaceholdersOf(layoutPath);
    }

    slide.background = this.backgroundCss(find(cSld, "bg"));
    // A slide that names no background shows the layout's, and the layout's own master's;
    // WPS writes the colour only on the master, so without this the dark and grey
    // templates open on white.
    if (!slide.background && slide.layoutPath) slide.background = this.inheritedBackground(pkg, slide.layoutPath);

    const spTree = find(cSld, "spTree");
    const rawTree = rawFind(slide.tree, "spTree");
    // Ordered trees keep whitespace between elements as text entries; only elements pair with shapes.
    const rawKids = rawTree ? rawChildren(rawTree).filter((c) => !!tagOf(c)) : [];
    let child = 0;
    for (const entry of kids(spTree)) {
      const shape = this.build(entry.name, entry.node, rels, path, rawKids[child++]);
      if (!shape) continue;
      this.inheritLayoutFrame(shape, entry.node, slide.layoutPlaceholders);
      slide.shapes.push(shape);
    }
    const notesRel = [...rels.values()].find((r) => r.type.endsWith("/notesSlide"));
    if (notesRel) {
      const notesPath = pkg.resolve(path, notesRel);
      if (pkg.parts.has(notesPath)) {
        slide.notesPath = notesPath;
        slide.notes = notesTextOf(pkg.text(notesPath) ?? "") || undefined;
      }
    }
    return slide;
  }

  private build(name: string, node: any, rels: Map<string, Rel>, sourcePart: string, raw?: any): Shape | undefined {
    if (name === "grpSp") {
      const xfrm = find(find(node, "grpSpPr"), "xfrm");
      const base = this.frame(xfrm);
      // A group's members sit directly under `p:grpSp`; only a slide carries `p:spTree`. The
      // parse side already fell back to the group node, so the raw side must index the same
      // list or every member loses its part nodes (and with them its `a:lstStyle` and `bodyPr`).
      const groupTree = raw ? rawChildren(raw).find((c) => tagLocalOf(c) === "spTree") ?? raw : undefined;
      const rawKids = groupTree ? rawChildren(groupTree).filter((c) => !!tagOf(c)) : [];
      const children: Shape[] = [];
      let child = 0;
      for (const entry of kids(find(node, "spTree") ?? node)) {
        const built = this.build(entry.name, entry.node, rels, sourcePart, rawKids[child++]);
        if (built) children.push(built);
      }
      return { kind: "group", name: "", ...base, children, raw };
    }
    if (name !== "sp" && name !== "pic" && name !== "graphicFrame" && name !== "cxnSp") return undefined;

    const nv = find(node, "nvSpPr") ?? find(node, "nvPicPr") ?? find(node, "nvGraphicFramePr") ?? find(node, "nvCxnSpPr");
    const label = String(attr(find(nv, "cNvPr"), "name") ?? "");
    // A deck may paint with the theme (`a:schemeClr accent1`) instead of an sRGB literal;
    // the master's `p:clrMap` says which clrScheme slot a `bg1`/`tx1` name means.
    const resolve = (scheme: any): string | undefined => this.themeHex(scheme);
    const ph = find(find(nv, "nvPr"), "ph");
    const spPr = find(node, "spPr") ?? find(node, "grpSpPr");
    // A graphicFrame carries its own p:xfrm; shapes and pictures nest it under spPr.
    const base = {
      name: label,
      ...this.frame(find(node, "xfrm") ?? find(spPr, "xfrm")),
      placeholder: ph ? String(attr(ph, "type") ?? "") : undefined,
      placeholderIdx: ph && attr(ph, "idx") !== undefined ? Number(attr(ph, "idx")) : undefined,
    };

    if (name === "pic") {
      const blip = find(find(node, "blipFill"), "blip");
      const embed = blip ? attr(blip, "embed") : undefined;
      const rel = embed ? rels.get(String(embed)) : undefined;
      return {
        kind: "picture",
        ...base,
        src: rel && this.pkg ? this.media.get(this.pkg.resolve(sourcePart, rel)) : undefined,
        geometry: geom(spPr),
        raw,
      };
    }
    if (name === "graphicFrame") {
      const tbl = findDeep(node, "tbl");
      const rows = findAll(tbl, "tr").map((tr) =>
        findAll(tr, "tc").map((tc): SlideCell => {
          const runs = paragraphs(find(tc, "txBody"), resolve).flatMap((p) => p.runs);
          return { text: runs.map((r) => r.text).join(""), runs, fill: solidHex(find(find(tc, "tcPr"), "solidFill"), resolve) };
        }),
      );
      return { kind: "table", ...base, rows: rows.length ? rows : [[]], raw };
    }
    const ln = find(spPr, "ln");
    // The text style is read off the raw part nodes: a parsed child map does not keep an
    // array shape when a level appears once, which rawFind needs.
    const rawTxBody = raw ? rawFind(rawChildren(raw), "txBody") : undefined;
    const lstStyle = rawTxBody ? rawFind(rawChildren(rawTxBody), "lstStyle") : undefined;
    const textLevels = lstStyle ? readLevels(lstStyle) : undefined;
    // "Shrink text on overflow" leaves the size it picked in `a:normAutofit`; without a
    // `fontScale` the producer expects the renderer to fit at draw time, which this model
    // does not measure. Only the stated percentage is a fact the file carries.
    const bodyPr = rawTxBody ? rawFind(rawChildren(rawTxBody), "bodyPr") : undefined;
    const norm = bodyPr ? rawChildren(bodyPr).find((c) => tagLocalOf(c) === "normAutofit") : undefined;
    const normAttrs = rawAttrs(norm ?? {});
    const fontScale = Number(normAttrs["@_fontScale"]);
    const reduction = Number(normAttrs["@_lnSpcReduction"]);
    return {
      kind: "shape",
      ...base,
      textBox: attr(find(nv, "cNvSpPr"), "txBox") === "1" || undefined,
      geometry: geom(spPr),
      fill: solidHex(find(spPr, "solidFill"), resolve),
      lineColor: solidHex(find(ln, "solidFill"), resolve),
      lineWidth: attr(ln, "w") ? Number(attr(ln, "w")) / EMU_PER_PX : undefined,
      paragraphs: paragraphs(find(node, "txBody"), resolve),
      textLevels: textLevels?.some(hasLevel) ? textLevels : undefined,
      textBody: readTextBody(bodyPr),
      fontScalePct: Number.isFinite(fontScale) && fontScale > 0 ? fontScale / 1000 : undefined,
      lineSpacingReductionPct: Number.isFinite(reduction) && reduction > 0 ? reduction / 1000 : undefined,
      raw,
    };
  }

  private frame(xfrm: any): { x: number; y: number; width: number; height: number; rotation?: number } {
    const off = find(xfrm, "off");
    const ext = find(xfrm, "ext");
    return {
      x: Number(attr(off, "x") ?? 0) / EMU_PER_PX,
      y: Number(attr(off, "y") ?? 0) / EMU_PER_PX,
      width: Number(attr(ext, "cx") ?? 0) / EMU_PER_PX,
      height: Number(attr(ext, "cy") ?? 0) / EMU_PER_PX,
      rotation: attr(xfrm, "rot") ? Number(attr(xfrm, "rot")) / 60000 : undefined,
    };
  }

  /**
   * A slide placeholder that states no transform takes the layout's geometry: PowerPoint
   * matches `p:ph@idx` first, then the type, and WPS's own templates leave every one to it.
   */
  private inheritLayoutFrame(shape: Shape, node: any, layout: LayoutPlaceholder[] | undefined): void {
    if (!layout?.length || shape.placeholder === undefined) return;
    const spPr = find(node, "spPr") ?? find(node, "grpSpPr");
    if (find(node, "xfrm") ?? find(spPr, "xfrm")) return;
    const nv = find(node, "nvSpPr") ?? find(node, "nvPicPr") ?? find(node, "nvGraphicFramePr") ?? find(node, "nvCxnSpPr");
    const ph = find(find(nv, "nvPr"), "ph");
    const idx = attr(ph, "idx") !== undefined ? Number(attr(ph, "idx")) : undefined;
    const type = String(attr(ph, "type") ?? "obj");
    const hit =
      (idx !== undefined ? layout.find((p) => p.idx === idx) : undefined) ??
      layout.find((p) => (p.type ?? "obj") === type) ??
      (type === "obj" ? layout.find((p) => p.type === undefined) : undefined);
    if (!hit) return;
    shape.x = hit.x;
    shape.y = hit.y;
    shape.width = hit.width;
    shape.height = hit.height;
  }

  private readTheme(pkg: OpcPackage): void {
    // The theme the slide master references owns the deck's colours. A package also
    // ships themes for handouts and notes, and reading those last would repaint the
    // slides with another palette (WPS's black template has a 44546A notes theme).
    const masterRel = [...pkg.rels("ppt/presentation.xml").values()].find((r) => r.type.endsWith("/slideMaster"));
    const masterPath = masterRel ? pkg.resolve("ppt/presentation.xml", masterRel) : undefined;
    const themeRel = masterPath ? [...pkg.rels(masterPath).values()].find((r) => r.type.endsWith("/theme")) : undefined;
    const owner = themeRel && masterPath ? pkg.resolve(masterPath, themeRel) : undefined;
    const paths = owner && pkg.parts.has(owner)
      ? [owner]
      : [...pkg.parts.keys()].filter((p) => /^ppt\/theme\/theme\d+\.xml$/.test(p));
    for (const path of paths) {
      const clrScheme = findDeep(pkg.xml(path), "clrScheme");
      for (const { name, node } of kids(clrScheme)) {
        // WPS names the two base slots with `sysClr` (windowText / window) and keeps
        // the concrete value in `lastClr`; the other slots are plain sRGB.
        const srgb = find(node, "srgbClr");
        const sys = srgb ? undefined : find(node, "sysClr");
        const val = srgb ? attr(srgb, "val") : attr(sys, "lastClr");
        if (val) this.theme.set(name.toLowerCase(), "#" + val);
      }
    }
  }

  private readMedia(pkg: OpcPackage): void {
    for (const [path, bytes] of pkg.parts) {
      if (!path.startsWith("ppt/media/")) continue;
      const type = IMAGE_MIME_BY_EXT[path.slice(path.lastIndexOf(".") + 1).toLowerCase()];
      if (type) this.media.set(path, dataUrl(type, bytes));
    }
  }

  /** Registers an inserted image and returns the package part a save will write it to. */
  addPicture(url: string): string {
    const mime = url.slice("data:".length, url.indexOf(";"));
    const ext = IMAGE_EXT_BY_MIME[mime] ?? "png";
    const used = new Set(this.addedMedia.keys());
    if (this.pkg) for (const path of this.pkg.parts.keys()) used.add(path);
    let n = 1;
    while (used.has(`ppt/media/image${n}.${ext}`)) n++;
    const path = `ppt/media/image${n}.${ext}`;
    this.addedMedia.set(path, url);
    return path;
  }

  themeColor(name: string): string | undefined {
    return this.theme.get(name);
  }

  /** The clrScheme slot a `a:schemeClr@val` names, after the master's `p:clrMap`. */
  private themeHex(scheme: any): string | undefined {
    const name = String(attr(scheme, "val") ?? "");
    return this.theme.get(this.clrMap.get(name) ?? name) ?? this.theme.get(name);
  }

  /** The layout's `p:bg`, else the master's — backgrounds cascade and stop at the master. */
  private inheritedBackground(pkg: OpcPackage, layoutPath: string): string | undefined {
    const own = this.backgroundOfPart(pkg, layoutPath);
    if (own) return own;
    const rel = [...pkg.rels(layoutPath).values()].find((r) => r.type.endsWith("/slideMaster"));
    return rel ? this.backgroundOfPart(pkg, pkg.resolve(layoutPath, rel)) : undefined;
  }

  /** A layout/master part's background as CSS, cached because slides share it. */
  private backgroundOfPart(pkg: OpcPackage, path: string): string | undefined {
    if (this.bgCache.has(path)) return this.bgCache.get(path);
    const cSld = pkg.parts.has(path) ? findDeep(pkg.xml(path), "cSld") : undefined;
    const css = this.backgroundCss(find(cSld, "bg"));
    this.bgCache.set(path, css);
    return css;
  }

  /** `p:bg` in the editor's CSS: a solid fill, or the linear gradient `a:gradFill` states. */
  private backgroundCss(bg: any): string | undefined {
    if (!bg) return undefined;
    // `p:bgRef` names its colour directly; the bgFillStyleLst index only picks a style
    // the deck already restates, and every WPS background here repeats the scheme colour.
    const ref = find(bg, "bgRef");
    if (ref) return this.colorCss(ref);
    const pr = find(bg, "bgPr");
    const solid = find(pr, "solidFill");
    if (solid) return this.colorCss(solid);
    const grad = find(pr, "gradFill");
    return grad ? this.gradientCss(grad) : undefined;
  }

  /** The first colour under `parent`, scheme references resolved and transforms applied. */
  private colorCss(parent: any): string | undefined {
    return solidHex(parent, (scheme) => this.themeHex(scheme));
  }

  private gradientCss(grad: any): string | undefined {
    const stops: string[] = [];
    for (const { node } of kids(find(grad, "gsLst"))) {
      const color = this.colorCss(node);
      if (color) stops.push(`${color} ${Number(attr(node, "pos") ?? 0) / 1000}%`);
    }
    if (stops.length < 2) return undefined;
    // OOXML measures clockwise from left-to-right; CSS from straight up.
    const deg = Number(attr(find(grad, "lin"), "ang") ?? 0) / 60000 + 90;
    return `linear-gradient(${deg}deg, ${stops.join(", ")})`;
  }

  /** Marks a slide as edited, so a save regenerates it instead of copying it through. */
  markDirty(slide: Slide): void {
    slide.dirty = true;
  }

  /** Lowest-numbered layout part in the package, for a deck whose slides name none. */
  /** Placeholder geometry of a layout: the package's own part, or the one a blank deck generates. */
  private layoutPlaceholdersOf(path: string | undefined): LayoutPlaceholder[] | undefined {
    if (!path) return undefined;
    if (this.pkg) {
      const xml = this.pkg.text(path);
      return xml ? parseLayoutPlaceholders(xml, this.masterPlaceholdersOf(path)) : undefined;
    }
    const m = /slideLayout(\d+)\.xml$/.exec(path);
    const def = m ? BLANK_LAYOUTS[Number(m[1]) - 1] : undefined;
    // A placeholder without geometry takes the master's, which this model does not resolve.
    return def?.placeholders
      .filter((p) => p.x !== undefined)
      .map((p) => ({
        type: p.type,
        idx: p.idx,
        x: (p.x ?? 0) / EMU_PER_PX,
        y: (p.y ?? 0) / EMU_PER_PX,
        width: (p.w ?? 0) / EMU_PER_PX,
        height: (p.h ?? 0) / EMU_PER_PX,
        sizePt: p.szPt,
        align: p.algn ? ALGN_IN[p.algn] : undefined,
        textBody: p.anchor ? { anchor: p.anchor } : undefined,
        marLEmu: p.zeroIndent ? 0 : undefined,
        indentEmu: p.zeroIndent ? 0 : undefined,
        noBullet: p.noBullet,
      }));
  }

  /** The master's placeholders: the geometry a layout borrows where it states no transform. */
  private masterPlaceholdersOf(layoutPath: string): LayoutPlaceholder[] | undefined {
    const pkg = this.pkg;
    if (!pkg) return undefined;
    const rel = [...pkg.rels(layoutPath).values()].find((r) => r.type.endsWith("/slideMaster"));
    if (!rel) return undefined;
    const path = pkg.resolve(layoutPath, rel);
    const cached = this.masterPlaceholderCache.get(path);
    if (cached) return cached;
    const xml = pkg.parts.has(path) ? pkg.text(path) : undefined;
    const list = xml ? parseLayoutPlaceholders(xml) : [];
    this.masterPlaceholderCache.set(path, list);
    return list;
  }

  /** Level-1 title/body sizes and the body's nine levels from the master `p:txStyles`. */
  private readMasterStyles(pkg: OpcPackage): void {
    const rel = [...pkg.rels("ppt/presentation.xml").values()].find((r) => r.type.endsWith("/slideMaster"));
    if (!rel) return;
    const xml = pkg.text(pkg.resolve("ppt/presentation.xml", rel));
    const tree = xml ? parseOrderedXml(xml) : undefined;
    const map = tree ? rawFind(tree, "clrMap") : undefined;
    for (const [from, to] of Object.entries(rawAttrs(map ?? {}))) {
      if (!from.startsWith("@_") || !to) continue;
      this.clrMap.set(from.slice(2), to);
    }
    const styles = tree ? rawFind(tree, "txStyles") : undefined;
    if (!styles) return;
    const styleOf = (name: string): any => rawFind(rawChildren(styles), name);
    const title = readLevels(styleOf("titleStyle"))[0]?.sizePt;
    const body = readLevels(styleOf("bodyStyle"));
    const other = readLevels(styleOf("otherStyle"))[0]?.sizePt;
    if (title !== undefined) this.titleStylePt = title;
    if (body[0]?.sizePt !== undefined) this.bodyStylePt = body[0].sizePt;
    if (other !== undefined) this.otherStylePt = other;
    if (body.some(hasLevel)) this.bodyLevels = body;
  }

  /**
   * The layouts a new slide can use, as the deck names them (`p:cSld@name`), in part
   * order. A blank deck offers the eleven layouts its writer generates.
   */
  layoutChoices(): { path: string; name: string }[] {
    if (this.layoutCache) return this.layoutCache;
    if (!this.pkg) {
      return (this.layoutCache = BLANK_LAYOUTS.map((l, i) => ({ path: `ppt/slideLayouts/slideLayout${i + 1}.xml`, name: l.name })));
    }
    const paths: string[] = [];
    for (const path of this.pkg.parts.keys()) if (/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)) paths.push(path);
    paths.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    return (this.layoutCache = paths.map((path, i) => {
      const tree = parseOrderedXml(this.pkg!.text(path) ?? "");
      const name = rawAttrs(rawFind(tree, "cSld") ?? {})["@_name"];
      return { path, name: name ?? `版式 ${i + 1}` };
    }));
  }

  private firstSlideLayout(): string | undefined {    let found: string | undefined;
    if (!this.pkg) return undefined;
    for (const path of this.pkg.parts.keys()) {
      if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path)) continue;
      if (!found || path.localeCompare(found, undefined, { numeric: true }) < 0) found = path;
    }
    return found;
  }

  /** Appends a blank slide that inherits from a layout; the argument overrides the default. */
  addSlide(layoutPath?: string): Slide {
    // Slides added in this session are not in the source package yet, so the
    // next free number and relationship id have to count them too.
    const used = new Set<number>();
    const takeNumber = (path: string): void => {
      const m = /^ppt\/slides\/slide(\d+)\.xml$/.exec(path);
      if (m) used.add(Number(m[1]));
    };
    if (this.pkg) for (const path of this.pkg.parts.keys()) takeNumber(path);
    for (const slide of this.slides) takeNumber(slide.partPath);
    let n = 1;
    while (used.has(n)) n++;
    const rels = this.pkg?.rels(this.presPath);
    let maxRel = this.blankRelBase;
    const takeRel = (id: string): void => {
      const m = /^rId(\d+)$/.exec(id);
      if (m) maxRel = Math.max(maxRel, Number(m[1]));
    };
    if (rels) for (const id of rels.keys()) takeRel(id);
    for (const slide of this.slides) takeRel(slide.relId);
    const maxId = this.slides.reduce((max, s) => Math.max(max, Number(s.sldId) || 0), 0);
    // A new slide must inherit from a layout; prefer the one existing slides use.
    if (!this.newSlideLayoutPath) this.newSlideLayoutPath = this.firstSlideLayout();
    const chosen = layoutPath ?? this.newSlideLayoutPath;
    const slide: Slide = {
      index: this.slides.length,
      name: `第 ${this.slides.length + 1} 页`,
      shapes: [],
      partPath: `ppt/slides/slide${n}.xml`,
      sldId: String(Math.max(256, maxId + 1)),
      relId: `rId${maxRel + 1}`,
      tree: parseOrderedXml(SLIDE_TEMPLATE),
      layoutPath: chosen,
      dirty: true,
      notesDirty: false,
    };
    slide.layoutPlaceholders = this.layoutPlaceholdersOf(chosen);
    if (this.pkg && chosen) slide.background = this.inheritedBackground(this.pkg, chosen);
    this.slides.push(slide);
    this.presDirty = true;
    return slide;
  }

  /** Removes a slide with its notes; a save drops those parts from the package. */
  removeSlide(index: number): void {
    const [slide] = this.slides.splice(index, 1);
    if (!slide) return;
    this.removedParts.add(slide.partPath);
    this.removedParts.add(relsPathOf(slide.partPath));
    if (slide.notesPath) {
      this.removedParts.add(slide.notesPath);
      this.removedParts.add(relsPathOf(slide.notesPath));
    }
    this.presDirty = true;
    this.slides.forEach((s, i) => ((s.index = i), (s.name = `第 ${i + 1} 页`)));
  }

  /** Moves a slide by `delta` positions in the deck order. */
  moveSlide(index: number, delta: number): void {
    const target = index + delta;
    if (target < 0 || target >= this.slides.length) return;
    const [slide] = this.slides.splice(index, 1);
    this.slides.splice(target, 0, slide);
    this.presDirty = true;
    this.slides.forEach((s, i) => ((s.index = i), (s.name = `第 ${i + 1} 页`)));
  }
}

/** RGB → HSL: hue in degrees, saturation and lightness in 0..1. */
function hslOf(hex: string): [number, number, number] {
  const n = parseInt(hex.replace(/^#/, ""), 16);
  const r = ((n >> 16) & 255) / 255;
  const g = ((n >> 8) & 255) / 255;
  const b = (n & 255) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h * 60, s, l];
}

/** HSL back to `#RRGGBB`, the form the model and the stylesheet carry. */
function hexOf(h: number, s: number, l: number): string {
  const hue = (((h % 360) + 360) % 360) / 360;
  const q = l <= 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channel = (t: number): string => {
    const tt = ((t % 1) + 1) % 1;
    const v = tt < 1 / 6 ? p + (q - p) * 6 * tt : tt < 1 / 2 ? q : tt < 2 / 3 ? p + (q - p) * (2 / 3 - tt) * 6 : p;
    return Math.round(Math.max(0, Math.min(1, v)) * 255)
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  };
  return `#${channel(hue + 1 / 3)}${channel(hue)}${channel(hue - 1 / 3)}`;
}

/**
 * Applies the colour transforms a colour node carries. DrawingML adjusts luminance in
 * HSL, not by scaling each channel: `lumMod`/`lumOff` move L, `tint` lifts it toward
 * white (`L·t + 1−t`), `shade` darkens it (`L·t`).
 */
function transformHex(hex: string, node: any): string {
  const amount = (name: string): number | undefined => {
    const v = Number(attr(find(node, name), "val"));
    return Number.isFinite(v) ? v / 100000 : undefined;
  };
  let [h, s, l] = hslOf(hex);
  const lumMod = amount("lumMod");
  if (lumMod !== undefined) l *= lumMod;
  const lumOff = amount("lumOff");
  if (lumOff !== undefined) l += lumOff;
  const tint = amount("tint");
  if (tint !== undefined) l = l * tint + (1 - tint);
  const shade = amount("shade");
  if (shade !== undefined) l *= shade;
  return hexOf(h, s, Math.max(0, Math.min(1, l)));
}

/**
 * Placeholders a layout or master declares. A content placeholder that states no transform
 * borrows its geometry from `master` (idx first, then type): WPS leaves 111 title/body
 * placeholders of its 330 layouts to the master. The footer, date and slide-number
 * placeholders also live only on the master, but nothing here draws them, so they stay out.
 */
function parseLayoutPlaceholders(xml: string, master?: LayoutPlaceholder[]): LayoutPlaceholder[] {
  const spTree = rawFind(parseOrderedXml(xml), "spTree");
  const out: LayoutPlaceholder[] = [];
  for (const entry of spTree ? rawChildren(spTree) : []) {
    if (tagLocalOf(entry) !== "sp") continue;
    const nv = rawFind(rawChildren(entry), "nvSpPr");
    const ph = nv ? rawFind(rawChildren(nv), "ph") : undefined;
    if (!ph) continue;
    const spPr = rawFind(rawChildren(entry), "spPr");
    const xfrm = spPr ? rawFind(rawChildren(spPr), "xfrm") : undefined;
    const off = xfrm ? rawFind(rawChildren(xfrm), "off") : undefined;
    const ext = xfrm ? rawFind(rawChildren(xfrm), "ext") : undefined;
    const a = rawAttrs(ph);
    const type = a["@_type"];
    const idx = a["@_idx"] !== undefined ? Number(a["@_idx"]) : undefined;
    const content = type === undefined || ["title", "ctrTitle", "subTitle", "body", "obj"].includes(type);
    const masterHit = content
      ? ((idx !== undefined ? master?.find((p) => p.idx === idx) : undefined) ??
        master?.find((p) => (p.type ?? "obj") === (type ?? "obj")))
      : undefined;
    const geometry =
      off && ext
        ? {
            x: Number(rawAttrs(off)["@_x"] ?? 0) / EMU_PER_PX,
            y: Number(rawAttrs(off)["@_y"] ?? 0) / EMU_PER_PX,
            width: Number(rawAttrs(ext)["@_cx"] ?? 0) / EMU_PER_PX,
            height: Number(rawAttrs(ext)["@_cy"] ?? 0) / EMU_PER_PX,
          }
        : masterHit;
    if (!geometry) continue;
    const txBody = rawFind(rawChildren(entry), "txBody");
    const bodyPr = txBody ? rawFind(rawChildren(txBody), "bodyPr") : undefined;
    const lst = txBody ? rawFind(rawChildren(txBody), "lstStyle") : undefined;
    const lvl1 = lst ? rawFind(rawChildren(lst), "lvl1pPr") : undefined;
    const level = lvl1 ? readLevel(lvl1) : {};
    out.push({
      type,
      idx,
      x: geometry.x,
      y: geometry.y,
      width: geometry.width,
      height: geometry.height,
      sizePt: level.sizePt,
      align: level.align,
      // `a:bodyPr` inherits attribute by attribute: the layout may anchor and leave the
      // insets to the master, or the other way round.
      textBody: mergeTextBody(readTextBody(bodyPr), masterHit?.textBody),
      marLEmu: level.marLEmu,
      indentEmu: level.indentEmu,
      bullet: level.bullet,
      noBullet: level.noBullet,
      lineSpacingPct: level.lineSpacingPct,
      spaceBeforePt: level.spaceBeforePt,
      spaceBeforePct: level.spaceBeforePct,
      spaceAfterPt: level.spaceAfterPt,
      spaceAfterPct: level.spaceAfterPct,
    });
  }
  return out;
}

/** `a:bodyPr` attributes; an empty element states nothing and stays undefined. */
function readTextBody(bodyPr: any): TextBody | undefined {
  if (!bodyPr) return undefined;
  const a = rawAttrs(bodyPr);
  const emu = (name: string): number | undefined => {
    const v = a[`@_${name}`];
    return v !== undefined && Number.isFinite(Number(v)) ? Number(v) : undefined;
  };
  const body: TextBody = {
    lInsEmu: emu("lIns"),
    tInsEmu: emu("tIns"),
    rInsEmu: emu("rIns"),
    bInsEmu: emu("bIns"),
    anchor: a["@_anchor"],
    noWrap: a["@_wrap"] === "none" ? true : undefined,
  };
  return Object.values(body).some((v) => v !== undefined) ? body : undefined;
}

/** The nearer `a:bodyPr` answers each attribute it states; the master answers the rest. */
function mergeTextBody(own: TextBody | undefined, fallback: TextBody | undefined): TextBody | undefined {
  if (!own) return fallback;
  if (!fallback) return own;
  const merged: TextBody = {
    lInsEmu: own.lInsEmu ?? fallback.lInsEmu,
    tInsEmu: own.tInsEmu ?? fallback.tInsEmu,
    rInsEmu: own.rInsEmu ?? fallback.rInsEmu,
    bInsEmu: own.bInsEmu ?? fallback.bInsEmu,
    anchor: own.anchor ?? fallback.anchor,
    noWrap: own.noWrap ?? fallback.noWrap,
  };
  return merged;
}

/** A paragraph's own spacing from the parsed tree; zero is a declared "none", same as silence. */
function paragraphSpacing(pPr: any): Pick<SlideParagraph, "lineSpacingPct" | "spaceBeforePt" | "spaceBeforePct" | "spaceAfterPt" | "spaceAfterPct"> {
  const pct = Number(attr(find(find(pPr, "lnSpc"), "spcPct"), "val"));
  const before = Number(attr(find(find(pPr, "spcBef"), "spcPts"), "val"));
  const beforePct = Number(attr(find(find(pPr, "spcBef"), "spcPct"), "val"));
  const after = Number(attr(find(find(pPr, "spcAft"), "spcPts"), "val"));
  const afterPct = Number(attr(find(find(pPr, "spcAft"), "spcPct"), "val"));
  return {
    lineSpacingPct: Number.isFinite(pct) && pct > 0 ? pct / 1000 : undefined,
    spaceBeforePt: Number.isFinite(before) && before > 0 ? before / 100 : undefined,
    spaceBeforePct: Number.isFinite(beforePct) && beforePct > 0 ? beforePct / 1000 : undefined,
    spaceAfterPt: Number.isFinite(after) && after > 0 ? after / 100 : undefined,
    spaceAfterPct: Number.isFinite(afterPct) && afterPct > 0 ? afterPct / 1000 : undefined,
  };
}

function paragraphs(txBody: any, resolve?: (node: any) => string | undefined): SlideParagraph[] {
  return findAll(txBody, "p").map((p) => {
    const pPr = find(p, "pPr");
    const runs: Inline[] = [];
    for (const child of kids(p)) {
      if (child.name === "r" || child.name === "fld") runs.push(inlineOf(child.node, resolve));
      else if (child.name === "br") runs.push({ text: "\n" } as Inline);
    }
    return {
      level: Number(attr(pPr, "lvl") ?? 0),
      align: ALGN_IN[attr(pPr, "algn") ?? ""],
      // Real producers nest `a:buChar char="•"`; an attribute is tolerated too.
      bullet: attr(find(pPr, "buChar"), "char") ?? attr(pPr, "buChar"),
      runs,
      ...paragraphSpacing(pPr),
    };
  });
}

function inlineOf(r: any, resolve?: (node: any) => string | undefined): Inline {
  const rPr = find(r, "rPr");
  return {
    text: textOf(find(r, "t")),
    bold: attr(rPr, "b") === "1",
    italic: attr(rPr, "i") === "1",
    // `Inline.underline` 存的是形状名（docx 那侧按 §17.17.109 带），这一侧的 `a:u` 档名（`sng`/`dbl`/`heavy`/…）
    // 在这份语料里还没有数过的受众，所以仍收成「有线 = single」——与这一行改之前同一眼。
    underline: attr(rPr, "u") && attr(rPr, "u") !== "none" ? "single" : undefined,
    sizePt: attr(rPr, "sz") ? Number(attr(rPr, "sz")) / 100 : undefined,
    color: solidHex(find(rPr, "solidFill"), resolve),
    font: attr(find(rPr, "latin"), "typeface"),
  } as Inline;
}

function textOf(t: any): string {
  if (t === undefined || t === null) return "";
  return typeof t === "string" ? t : (t["#text"] ?? "");
}

function geom(spPr: any): string | undefined {
  return attr(find(spPr, "prstGeom"), "prst");
}

/** A fill/line colour as `#RRGGBB`; a theme reference goes through `resolve`. */
function solidHex(fill: any, resolve?: (node: any) => string | undefined): string | undefined {
  if (!fill) return undefined;
  const srgb = find(fill, "srgbClr");
  const sys = srgb ? undefined : find(fill, "sysClr");
  const scheme = srgb || sys ? undefined : find(fill, "schemeClr");
  const node = srgb ?? sys ?? scheme;
  const hex = srgb ? "#" + attr(srgb, "val") : sys ? "#" + (attr(sys, "lastClr") ?? "") : scheme && resolve ? resolve(scheme) : undefined;
  return hex && hex.length === 7 ? transformHex(hex, node) : undefined;
}

function findDeep(node: any, name: string): any {
  if (!node || typeof node !== "object") return undefined;
  for (const child of kids(node)) {
    if (child.name === name) return child.node;
    const hit = findDeep(child.node, name);
    if (hit) return hit;
  }
  return undefined;
}

// ------------------------------------------------------------- raw part trees

function decodeXml(s: string): string {
  return s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

/**
 * The notes body placeholder's paragraphs, one line each. The slide image and the
 * number placeholders carry no speaker text, and an empty paragraph is a real line.
 */
function notesTextOf(raw: string): string {
  const shapes = rawFindAll(parseOrderedXml(raw), "sp");
  const target =
    shapes.find((sp) => {
      const ph = rawFind(rawChildren(sp), "ph");
      return ph ? rawAttrs(ph)["@_type"] === "body" : false;
    }) ?? shapes.find((sp) => rawFind(rawChildren(sp), "txBody"));
  const txBody = target ? rawFind(rawChildren(target), "txBody") : undefined;
  if (!txBody) return "";
  const lines = rawFindAll(rawChildren(txBody), "p").map((p) =>
    rawFindAll(rawChildren(p), "t")
      .map((t) => rawChildren(t).map((entry) => String(entry["#text"] ?? "")).join(""))
      .join(""),
  );
  const text = decodeXml(lines.join("\n"));
  return text.trim() ? text : "";
}

/** A blank slide, valid on its own; its relationship to a layout is added by the writer. */
const SLIDE_TEMPLATE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

/** A blank presentation root for a deck that never came from a file. */
const PRESENTATION_TEMPLATE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst/><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
