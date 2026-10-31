import { OpcPackage, asArray, children as kids, parseOrderedXml, rawAttrs, rawChildren, rawFind, rawFindAll, relsPathOf, tagLocalOf, tagOf, type Rel } from "./opc.js";
import { dataUrl, IMAGE_EXT_BY_MIME, IMAGE_MIME_BY_EXT } from "./media.js";
import { BLANK_LAYOUTS } from "./pptx-blank.js";
import type { Inline } from "./docx.js";

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
  /** `a:bodyPr@anchor`: "t", "ctr" or "b". */
  anchor?: string;
}

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
  /** Text was edited; a save rebuilds this shape's text body from the model. */
  textDirty?: boolean;
}

export interface SlideParagraph {
  align?: string;
  level: number;
  bullet?: string;
  runs: Inline[];
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
  /** Images inserted in the editor (package part → data URL); a save writes these parts. */
  addedMedia = new Map<string, string>();
  private theme = new Map<string, string>();
  private media = new Map<string, string>();
  private pkg?: OpcPackage;

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

    const bgPr = find(find(cSld, "bg"), "bgPr");
    const bgClr = find(find(bgPr, "solidFill"), "srgbClr");
    if (bgClr) slide.background = "#" + attr(bgClr, "val");

    const spTree = find(cSld, "spTree");
    const rawTree = rawFind(slide.tree, "spTree");
    // Ordered trees keep whitespace between elements as text entries; only elements pair with shapes.
    const rawKids = rawTree ? rawChildren(rawTree).filter((c) => !!tagOf(c)) : [];
    let child = 0;
    for (const entry of kids(spTree)) {
      const shape = this.build(entry.name, entry.node, rels, path, rawKids[child++]);
      if (shape) slide.shapes.push(shape);
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
      const groupTree = raw ? rawFind(rawChildren(raw), "spTree") : undefined;
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
          const runs = paragraphs(find(tc, "txBody")).flatMap((p) => p.runs);
          return { text: runs.map((r) => r.text).join(""), runs, fill: solidHex(find(find(tc, "tcPr"), "solidFill")) };
        }),
      );
      return { kind: "table", ...base, rows: rows.length ? rows : [[]], raw };
    }
    const ln = find(spPr, "ln");
    return {
      kind: "shape",
      ...base,
      textBox: attr(find(nv, "cNvSpPr"), "txBox") === "1" || undefined,
      geometry: geom(spPr),
      fill: solidHex(find(spPr, "solidFill")),
      lineColor: solidHex(find(ln, "solidFill")),
      lineWidth: attr(ln, "w") ? Number(attr(ln, "w")) / EMU_PER_PX : undefined,
      paragraphs: paragraphs(find(node, "txBody")),
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

  private readTheme(pkg: OpcPackage): void {
    for (const [path] of pkg.parts) {
      if (!/^ppt\/theme\/theme\d+\.xml$/.test(path)) continue;
      const clrScheme = findDeep(pkg.xml(path), "clrScheme");
      for (const { name, node } of kids(clrScheme)) {
        const srgb = find(node, "srgbClr");
        if (srgb) this.theme.set(name.toLowerCase(), "#" + attr(srgb, "val"));
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
      return xml ? parseLayoutPlaceholders(xml) : undefined;
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
        anchor: p.anchor,
      }));
  }

  /** Level-1 title/body sizes from the one master `p:txStyles`; layout placeholders inherit them. */
  private readMasterStyles(pkg: OpcPackage): void {
    const rel = [...pkg.rels("ppt/presentation.xml").values()].find((r) => r.type.endsWith("/slideMaster"));
    if (!rel) return;
    const xml = pkg.text(pkg.resolve("ppt/presentation.xml", rel));
    const styles = xml ? rawFind(parseOrderedXml(xml), "txStyles") : undefined;
    if (!styles) return;
    const level1Pt = (name: string): number | undefined => {
      const style = rawFind(rawChildren(styles), name);
      const lvl = style ? rawFind(rawChildren(style), "lvl1pPr") : undefined;
      const rpr = lvl ? rawFind(rawChildren(lvl), "defRPr") : undefined;
      const sz = Number(rawAttrs(rpr ?? {})["@_sz"]);
      return Number.isFinite(sz) && sz > 0 ? sz / 100 : undefined;
    };
    const title = level1Pt("titleStyle");
    const body = level1Pt("bodyStyle");
    if (title !== undefined) this.titleStylePt = title;
    if (body !== undefined) this.bodyStylePt = body;
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

/** Placeholders a layout declares with geometry of its own; the ones that inherit from the master are skipped. */
function parseLayoutPlaceholders(xml: string): LayoutPlaceholder[] {
  const spTree = rawFind(parseOrderedXml(xml), "spTree");
  const out: LayoutPlaceholder[] = [];
  for (const entry of spTree ? rawChildren(spTree) : []) {
    if (tagLocalOf(entry) !== "sp") continue;
    const nv = rawFind(rawChildren(entry), "nvSpPr");
    const ph = nv ? rawFind(rawChildren(nv), "ph") : undefined;
    const spPr = rawFind(rawChildren(entry), "spPr");
    const xfrm = spPr ? rawFind(rawChildren(spPr), "xfrm") : undefined;
    const off = xfrm ? rawFind(rawChildren(xfrm), "off") : undefined;
    const ext = xfrm ? rawFind(rawChildren(xfrm), "ext") : undefined;
    if (!ph || !off || !ext) continue;
    const a = rawAttrs(ph);
    const o = rawAttrs(off);
    const e = rawAttrs(ext);
    const txBody = rawFind(rawChildren(entry), "txBody");
    const bodyPr = txBody ? rawFind(rawChildren(txBody), "bodyPr") : undefined;
    const lst = txBody ? rawFind(rawChildren(txBody), "lstStyle") : undefined;
    const lvl1 = lst ? rawFind(rawChildren(lst), "lvl1pPr") : undefined;
    const defRPr = lvl1 ? rawFind(rawChildren(lvl1), "defRPr") : undefined;
    const sz = Number(rawAttrs(defRPr ?? {})["@_sz"]);
    out.push({
      type: a["@_type"],
      idx: a["@_idx"] !== undefined ? Number(a["@_idx"]) : undefined,
      x: Number(o["@_x"] ?? 0) / EMU_PER_PX,
      y: Number(o["@_y"] ?? 0) / EMU_PER_PX,
      width: Number(e["@_cx"] ?? 0) / EMU_PER_PX,
      height: Number(e["@_cy"] ?? 0) / EMU_PER_PX,
      sizePt: Number.isFinite(sz) && sz > 0 ? sz / 100 : undefined,
      align: lvl1 ? ALGN_IN[rawAttrs(lvl1)["@_algn"] ?? ""] : undefined,
      anchor: bodyPr ? rawAttrs(bodyPr)["@_anchor"] : undefined,
    });
  }
  return out;
}

function paragraphs(txBody: any): SlideParagraph[] {
  return findAll(txBody, "p").map((p) => {
    const pPr = find(p, "pPr");
    const runs: Inline[] = [];
    for (const child of kids(p)) {
      if (child.name === "r" || child.name === "fld") runs.push(inlineOf(child.node));
      else if (child.name === "br") runs.push({ text: "\n" } as Inline);
    }
    return {
      level: Number(attr(pPr, "lvl") ?? 0),
      align: ALGN_IN[attr(pPr, "algn") ?? ""],
      // Real producers nest `a:buChar char="•"`; an attribute is tolerated too.
      bullet: attr(find(pPr, "buChar"), "char") ?? attr(pPr, "buChar"),
      runs,
    };
  });
}

function inlineOf(r: any): Inline {
  const rPr = find(r, "rPr");
  return {
    text: textOf(find(r, "t")),
    bold: attr(rPr, "b") === "1",
    italic: attr(rPr, "i") === "1",
    // `Inline.underline` 存的是形状名（docx 那侧按 §17.17.109 带），这一侧的 `a:u` 档名（`sng`/`dbl`/`heavy`/…）
    // 在这份语料里还没有数过的受众，所以仍收成「有线 = single」——与这一行改之前同一眼。
    underline: attr(rPr, "u") && attr(rPr, "u") !== "none" ? "single" : undefined,
    sizePt: attr(rPr, "sz") ? Number(attr(rPr, "sz")) / 100 : undefined,
    color: solidHex(find(rPr, "solidFill")),
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

function solidHex(fill: any): string | undefined {
  const v = attr(find(fill, "srgbClr"), "val");
  return v ? "#" + v : undefined;
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
