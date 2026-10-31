import { strToU8, zipSync } from "fflate";
import { bytesOfDataUrl, IMAGE_MIME_BY_EXT } from "./media.js";
import { buildXml, parseOrderedXml, rawAttrs, rawChildren, rawFind, rawFindAll, relsPathOf, tagLocalOf } from "./opc.js";
import type { Presentation, Shape, Slide, SlideCell, SlideParagraph } from "./pptx.js";
import { blankDeckParts } from "./pptx-blank.js";
import type { Inline } from "./docx.js";
import { hasLine } from "./docx.js";

/**
 * Serializes an edited deck. Slides the user did not touch are copied through
 * byte-identical; an edited slide is patched in place on the raw part tree, so
 * geometry, effects and text this model does not surface stay in the file.
 */

const XML_ESC = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const EMU_PER_PX = 9525;

/** The editor's alignment words, back to the `algn` values OOXML stores. */
const ALGN_OUT: Record<string, string> = { left: "l", center: "ctr", right: "r", both: "just", justify: "just" };
const CT_P = "application/vnd.openxmlformats-officedocument.presentationml";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

export function writePptx(pres: Presentation): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  if (pres.src) for (const [path, bytes] of pres.src.parts) files[path] = bytes;
  for (const path of pres.removedParts) delete files[path];
  for (const [path, url] of pres.addedMedia) files[path] = bytesOfDataUrl(url);
  // A deck built in the editor has no source package: master, layouts, theme and
  // the document properties are generated here, never overwriting what was written.
  if (!pres.src) for (const [path, xml] of Object.entries(blankDeckParts(pres.slides.length))) if (!files[path]) files[path] = strToU8(xml);

  for (const slide of pres.slides) {
    const relsPath = relsPathOf(slide.partPath);
    const rewritten = slide.dirty || !files[slide.partPath];
    // Notes typed into a deck whose slide never had a notes part need that part
    // created first: the slide's relationship has to name a part that exists.
    const addedNotes = slide.notesDirty && !slide.notesPath && !!(slide.notes ?? "").trim();
    if (addedNotes) slide.notesPath = createNotesPart(slide, files);
    // A picture the editor inserted needs a relationship in the slide part, which
    // the id must take before the shape node naming that id is built.
    const extra = rewritten ? assignPictureRels(pres, slide, !files[relsPath] && !!(slide.layoutPath ?? pres.newSlideLayoutPath)) : [];
    if (rewritten) {
      for (const shape of slide.shapes) patchShape(shape);
      syncShapeTree(rawFind(slide.tree, "spTree"), slide.shapes);
      files[slide.partPath] = strToU8(buildXml(slide.tree));
    }
    if (extra.length || addedNotes) files[relsPath] = strToU8(slideRels(pres, slide, extra));
    else if (!files[relsPath] && slide.partPath) files[relsPath] = strToU8(slideRels(pres, slide));
    if (slide.notesDirty && slide.notesPath && files[slide.notesPath]) patchNotes(slide, files);
  }

  if (pres.presDirty) {
    patchPresentation(pres);
    files[pres.presPath] = strToU8(buildXml(pres.presTree));
    files[relsPathOf(pres.presPath)] = strToU8(presentationRels(pres));
  }

  // A deck built in the editor has no source package, so it also needs the root
  // relationship that names its main part.
  if (!files["_rels/.rels"]) {
    files["_rels/.rels"] = strToU8(
      `${DECL}<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="${XML_ESC(pres.presPath)}"/></Relationships>`,
    );
  }

  files["[Content_Types].xml"] = strToU8(contentTypesXml(pres, files));
  return zipSync(files, { level: 6 });
}

/**
 * Rewrites a shape container from the model: kept elements first, then one node
 * per model shape in order, so a deleted shape's node leaves the file and a
 * shape added in the editor gets a fresh element.
 */
function syncShapeTree(spTree: any, shapes: Shape[]): void {
  if (!spTree) return;
  const children = rawChildren(spTree);
  const kept = children.filter((c) => {
    const tag = tagLocalOf(c);
    return tag !== undefined && !SHAPE_TAGS.has(tag);
  });
  let nextId = 2;
  for (const cNvPr of rawFindAll(children, "cNvPr")) {
    const id = Number(rawAttrs(cNvPr)["@_id"] ?? 0);
    if (Number.isFinite(id)) nextId = Math.max(nextId, id + 1);
  }
  const nodes: any[] = [];
  for (const shape of shapes) {
    const node = shape.raw ?? (shape.kind === "picture" ? pictureElement(shape, nextId++) : shapeElement(shape, nextId++));
    if (!node) continue;
    nodes.push(node);
    if (shape.kind === "group") syncShapeTree(rawFind(rawChildren(node), "spTree"), shape.children ?? []);
  }
  children.length = 0;
  children.push(...kept, ...nodes);
}

const SHAPE_TAGS = new Set(["sp", "pic", "graphicFrame", "grpSp", "cxnSp"]);

function shapeElement(shape: Shape, id: number): any | undefined {
  if (shape.kind !== "shape") return undefined;
  const xfrm: any = {
    "a:xfrm": [
      { "a:off": [], ":@": { "@_x": emu(shape.x), "@_y": emu(shape.y) } },
      { "a:ext": [], ":@": { "@_cx": emu(shape.width), "@_cy": emu(shape.height) } },
    ],
    ":@": shape.rotation ? { "@_rot": String(Math.round(shape.rotation * 60000)) } : {},
  };
  const spPr: any[] = [xfrm, { "a:prstGeom": [{ "a:avLst": [] }], ":@": { "@_prst": shape.geometry ?? "rect" } }];
  if (shape.fill) spPr.push({ "a:solidFill": [{ "a:srgbClr": [], ":@": { "@_val": shape.fill.replace(/^#/, "").toUpperCase() } }] });
  if (shape.lineColor) {
    spPr.push({
      "a:ln": [{ "a:solidFill": [{ "a:srgbClr": [], ":@": { "@_val": shape.lineColor.replace(/^#/, "").toUpperCase() } }] }],
      ":@": shape.lineWidth ? { "@_w": emu(shape.lineWidth) } : {},
    });
  }
  const txBody: any[] = [{ "a:bodyPr": [] }, { "a:lstStyle": [] }, ...(shape.paragraphs ?? []).map(paragraphNode)];
  // A shape that fills a layout placeholder keeps its identity, or PowerPoint draws
  // the prompt again over the text the user typed.
  const phAttrs: Record<string, string> = {};
  if (shape.placeholder) phAttrs["@_type"] = shape.placeholder;
  if (shape.placeholderIdx !== undefined) phAttrs["@_idx"] = String(shape.placeholderIdx);
  return {
    "p:sp": [
      {
        "p:nvSpPr": [
          { "p:cNvPr": [], ":@": { "@_id": String(id), "@_name": shape.name || "形状" } },
          { "p:cNvSpPr": [], ":@": shape.textBox ? { "@_txBox": "1" } : {} },
          { "p:nvPr": shape.placeholder ? [{ "p:ph": [], ":@": phAttrs }] : [] },
        ],
      },
      { "p:spPr": spPr },
      { "p:txBody": txBody },
    ],
  };
}

/** An inserted picture: the slide relationship names the media part, the model names the id. */
function pictureElement(shape: Shape, id: number): any | undefined {
  if (!shape.relId) return undefined;
  const xfrmKids: any[] = [
    { "a:off": [], ":@": { "@_x": emu(shape.x), "@_y": emu(shape.y) } },
    { "a:ext": [], ":@": { "@_cx": emu(shape.width), "@_cy": emu(shape.height) } },
  ];
  const xfrm: any = { "a:xfrm": xfrmKids, ":@": shape.rotation ? { "@_rot": String(Math.round(shape.rotation * 60000)) } : {} };
  return {
    "p:pic": [
      {
        "p:nvPicPr": [
          { "p:cNvPr": [], ":@": { "@_id": String(id), "@_name": shape.name || "图片" } },
          { "p:cNvPicPr": [{ "a:picLocks": [], ":@": { "@_noChangeAspect": "1" } }] },
          { "p:nvPr": [] },
        ],
      },
      { "p:blipFill": [{ "a:blip": [], ":@": { "@_r:embed": shape.relId } }, { "a:stretch": [{ "a:fillRect": [] }] }] },
      { "p:spPr": [xfrm, { "a:prstGeom": [{ "a:avLst": [] }], ":@": { "@_prst": "rect" } }] },
    ],
  };
}

// ------------------------------------------------------------------- shapes

function patchShape(shape: Shape): void {
  const raw = shape.raw;
  if (!raw) return;
  const spPr = rawFind(rawChildren(raw), "spPr");
  const xfrm = (spPr ? rawFind(rawChildren(spPr), "xfrm") : undefined) ?? rawFind(rawChildren(raw), "xfrm");
  if (xfrm) {
    const off = rawFind(rawChildren(xfrm), "off");
    if (off) {
      const a = rawAttrs(off);
      a["@_x"] = emu(shape.x);
      a["@_y"] = emu(shape.y);
    }
    const ext = rawFind(rawChildren(xfrm), "ext");
    if (ext) {
      const a = rawAttrs(ext);
      a["@_cx"] = emu(shape.width);
      a["@_cy"] = emu(shape.height);
    }
    const xa = rawAttrs(xfrm);
    if (shape.rotation) xa["@_rot"] = String(Math.round(shape.rotation * 60000));
    else delete xa["@_rot"];
  }
  if (shape.textDirty && shape.kind === "shape") {
    replaceTextBody(raw, shape.paragraphs ?? [], true);
  }
  if (shape.textDirty && shape.kind === "picture") replaceTextBody(raw, shape.paragraphs ?? []);
  if (shape.textDirty && shape.kind === "table") patchTable(raw, shape.rows ?? []);
  if (shape.kind === "group") for (const child of shape.children ?? []) patchShape(child);
}

const emu = (px: number): string => String(Math.round(px * EMU_PER_PX));

/** CT_Shape child order, for a text body created on a shape the source had none for. */
const SHAPE_BODY_ORDER = ["nvSpPr", "nvCxnSpPr", "spPr", "style", "txBody", "extLst"];
const bodyRank = (entry: any): number => {
  const tag = tagLocalOf(entry);
  // Pretty-printed whitespace is not an element and never bounds the insertion.
  if (tag === undefined) return -1;
  const i = SHAPE_BODY_ORDER.indexOf(tag);
  return i < 0 ? SHAPE_BODY_ORDER.length + 1 : i;
};

/** Swaps every `a:p` inside the element's text body for the model's paragraphs. */
function replaceTextBody(scope: any, paragraphs: SlideParagraph[], create = false): void {
  const kids = rawChildren(scope);
  const txBody = rawFind(kids, "txBody");
  if (txBody) {
    replaceParagraphs(rawChildren(txBody), paragraphs);
    return;
  }
  // A connector or bare autoshape carries no text body; text typed into one still
  // belongs in the file, so the body is created where the schema orders it.
  if (!create || !paragraphs.length) return;
  const at = kids.findIndex((kid) => bodyRank(kid) > SHAPE_BODY_ORDER.indexOf("txBody"));
  kids.splice(at < 0 ? kids.length : at, 0, { "p:txBody": [{ "a:bodyPr": [] }, { "a:lstStyle": [] }, ...paragraphs.map(paragraphNode)] });
}

function patchTable(scope: any, rows: SlideCell[][]): void {
  const tbl = rawFind(rawChildren(scope), "tbl");
  if (!tbl) return;
  const trs = rawFindAll(rawChildren(tbl), "tr");
  trs.forEach((tr, r) => {
    const tcs = rawFindAll(rawChildren(tr), "tc");
    tcs.forEach((tc, c) => {
      const cell = rows[r]?.[c];
      if (!cell) return;
      const want = cell.runs.length ? cell.runs.map((run) => run.text).join("") : cell.text;
      if (cellTextOf(tc) === want) return;
      const txBody = rawFind(rawChildren(tc), "txBody");
      if (txBody) replaceParagraphs(rawChildren(txBody), [{ level: 0, runs: cell.runs.length ? cell.runs : [{ text: cell.text }] }]);
    });
  });
}

/** Joins a cell's text the same way the reader does, so an untouched cell is left alone. */
function cellTextOf(tc: any): string {
  const walk = (node: any): string => {
    let out = "";
    for (const child of rawChildren(node)) {
      const tag = tagLocalOf(child);
      if (tag === "t") out += rawChildren(child).map((text: any) => String(text["#text"] ?? "")).join("");
      else if (tag === "br") out += "\n";
      else if (tag !== undefined) out += walk(child);
    }
    return out;
  };
  return walk(tc);
}

function replaceParagraphs(children: any[], paragraphs: SlideParagraph[]): void {
  const kept = children.filter((c) => {
    const tag = tagLocalOf(c);
    return !!tag && tag !== "p";
  });
  const rebuilt = paragraphs.length ? paragraphs.map(paragraphNode) : [];
  children.length = 0;
  children.push(...kept, ...rebuilt);
}

function paragraphNode(p: SlideParagraph): any {
  const kids: any[] = [];
  const attrs: Record<string, string> = {};
  if (p.level) attrs["@_lvl"] = String(p.level);
  if (p.marLEmu !== undefined) attrs["@_marL"] = String(p.marLEmu);
  if (p.indentEmu !== undefined) attrs["@_indent"] = String(p.indentEmu);
  if (p.align) attrs["@_algn"] = ALGN_OUT[p.align] ?? p.align;
  // CT_TextParagraphProperties order: lnSpc, spcBef, spcAft, then the bullet.
  if (p.lineSpacingPct) kids.push({ "a:lnSpc": [{ "a:spcPct": [], ":@": { "@_val": String(Math.round(p.lineSpacingPct * 1000)) } }] });
  if (p.spaceBeforePt) kids.push({ "a:spcBef": [{ "a:spcPts": [], ":@": { "@_val": String(Math.round(p.spaceBeforePt * 100)) } }] });
  else if (p.spaceBeforePct !== undefined) kids.push({ "a:spcBef": [{ "a:spcPct": [], ":@": { "@_val": String(Math.round(p.spaceBeforePct * 1000)) } }] });
  if (p.spaceAfterPt) kids.push({ "a:spcAft": [{ "a:spcPts": [], ":@": { "@_val": String(Math.round(p.spaceAfterPt * 100)) } }] });
  else if (p.spaceAfterPct !== undefined) kids.push({ "a:spcAft": [{ "a:spcPct": [], ":@": { "@_val": String(Math.round(p.spaceAfterPct * 1000)) } }] });
  if (p.noBullet) kids.push({ "a:buNone": [] });
  else if (p.bullet) kids.push({ "a:buChar": [], ":@": { "@_char": p.bullet } });
  // DrawingML spells a soft line break as a sibling `a:br`, never as text.
  const node: any[] = [];
  // An empty paragraph property node says nothing and would churn text nobody edited.
  if (kids.length || Object.keys(attrs).length) node.push({ "a:pPr": kids, ":@": attrs });
  for (const run of p.runs) {
    run.text.split("\n").forEach((text, i) => {
      if (i > 0) node.push({ "a:br": [] });
      if (text) node.push(runNode(run, text));
    });
  }
  return { "a:p": node };
}

function runNode(run: Inline, text: string): any {
  const attrs: Record<string, string> = {};
  if (run.bold) attrs["@_b"] = "1";
  if (run.italic) attrs["@_i"] = "1";
  if (hasLine(run.underline)) attrs["@_u"] = "sng";
  if (run.sizePt) attrs["@_sz"] = String(Math.round(run.sizePt * 100));
  const children: any[] = [];
  if (run.color) children.push({ "a:solidFill": [{ "a:srgbClr": [], ":@": { "@_val": run.color.replace(/^#/, "").toUpperCase() } }] });
  if (run.font) children.push({ "a:latin": [], ":@": { "@_typeface": run.font } });
  const rPr: any = { "a:rPr": children, ":@": attrs };
  const t: any = { "a:t": [{ "#text": text }] };
  return { "a:r": children.length || Object.keys(attrs).length ? [rPr, t] : [t] };
}

// ------------------------------------------------------------------- notes

function patchNotes(slide: Slide, files: Record<string, Uint8Array>): void {
  const path = slide.notesPath;
  if (!path) return;
  const tree = parseOrderedXml(new TextDecoder("utf-8").decode(files[path]));
  const shapes = rawFindAll(tree, "sp");
  const byPh = shapes.find((sp) => {
    const ph = rawFind(rawChildren(sp), "ph");
    return ph ? rawAttrs(ph)["@_type"] === "body" : false;
  });
  const target = byPh ?? shapes.find((sp) => rawFind(rawChildren(sp), "txBody"));
  const txBody = target ? rawFind(rawChildren(target), "txBody") : undefined;
  if (txBody) {
    const lines = (slide.notes ?? "").split("\n");
    replaceParagraphs(rawChildren(txBody), lines.map((text) => ({ level: 0, runs: [{ text }] })));
    files[path] = strToU8(buildXml(tree));
  }
}

/**
 * A notes part for a slide that never had one, shaped like the one WPS writes:
 * the slide image, the body placeholder and the slide-number placeholder, with
 * relationships back to the slide and (when the deck has one) the notes master.
 */
function createNotesPart(slide: Slide, files: Record<string, Uint8Array>): string {
  let n = 1;
  while (files[`ppt/notesSlides/notesSlide${n}.xml`]) n++;
  const path = `ppt/notesSlides/notesSlide${n}.xml`;
  files[path] = strToU8(notesSlideXml((slide.notes ?? "").split("\n")));
  const master = Object.keys(files)
    .filter((p) => /^ppt\/notesMasters\/notesMaster\d+\.xml$/.test(p))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))[0];
  const rels = [`<Relationship Id="rId1" Type="${NS_R}/slide" Target="${XML_ESC(relativeTarget(path, slide.partPath))}"/>`];
  if (master) rels.push(`<Relationship Id="rId2" Type="${NS_R}/notesMaster" Target="${XML_ESC(relativeTarget(path, master))}"/>`);
  files[relsPathOf(path)] = strToU8(`${DECL}<Relationships xmlns="${NS_PKG}">${rels.join("")}</Relationships>`);
  return path;
}

function notesSlideXml(lines: string[]): string {
  const body = lines
    .map((line) => (line ? `<a:p><a:r><a:rPr lang="zh-CN" altLang="en-US"/><a:t>${XML_ESC(line)}</a:t></a:r></a:p>` : "<a:p/>"))
    .join("");
  return `${DECL}<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="幻灯片图像占位符 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:endParaRPr lang="zh-CN" altLang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="备注占位符 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="幻灯片编号占位符 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN" altLang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

// ---------------------------------------------------------- presentation.xml

function patchPresentation(pres: Presentation): void {
  const lst = rawFind(pres.presTree, "sldIdLst");
  if (!lst) return;
  const children = rawChildren(lst);
  const kept = children.filter((c) => tagLocalOf(c) !== undefined && tagLocalOf(c) !== "sldId");
  const items = pres.slides.map((slide) => ({ "p:sldId": [], ":@": { "@_id": slide.sldId, "@_r:id": slide.relId } }));
  children.length = 0;
  children.push(...items, ...kept);
}

function presentationRels(pres: Presentation): string {
  const base = pres.src?.rels(pres.presPath);
  const slideIds = new Set(pres.slides.map((s) => s.relId));
  const others: string[] = [];
  if (base) {
    for (const [id, rel] of base) {
      if (slideIds.has(id) || rel.type.endsWith("/slide")) continue;
      others.push(`<Relationship Id="${XML_ESC(id)}" Type="${XML_ESC(rel.type)}" Target="${XML_ESC(rel.target)}"${rel.targetMode === "External" ? ' TargetMode="External"' : ""}/>`);
    }
  } else {
    // The generated blank package hangs master, theme and the three property parts
    // under rId1..rId5, which slides start after.
    for (const [id, type, target] of [
      ["rId1", "slideMaster", "slideMasters/slideMaster1.xml"],
      ["rId2", "theme", "theme/theme1.xml"],
      ["rId3", "presProps", "presProps.xml"],
      ["rId4", "viewProps", "viewProps.xml"],
      ["rId5", "tableStyles", "tableStyles.xml"],
    ]) {
      others.push(`<Relationship Id="${id}" Type="${NS_R}/${type}" Target="${target}"/>`);
    }
  }
  const slides = pres.slides.map((s) => {
    const target = "slides/" + s.partPath.slice("ppt/slides/".length);
    return `<Relationship Id="${XML_ESC(s.relId)}" Type="${NS_R}/slide" Target="${XML_ESC(target)}"/>`;
  });
  return `${DECL}<Relationships xmlns="${NS_PKG}">${slides.join("")}${others.join("")}</Relationships>`;
}

/** A slide's relationships: everything the source part had, plus one per inserted picture. */
function slideRels(pres: Presentation, slide: Slide, extra: { id: string; path: string }[] = []): string {
  const rels: string[] = [];
  const src = pres.src?.rels(slide.partPath);
  if (src?.size) {
    for (const [id, rel] of src) {
      rels.push(
        `<Relationship Id="${XML_ESC(id)}" Type="${XML_ESC(rel.type)}" Target="${XML_ESC(rel.target)}"${rel.targetMode === "External" ? ' TargetMode="External"' : ""}/>`,
      );
    }
  } else if (slide.layoutPath ?? pres.newSlideLayoutPath) {
    const target = relativeTarget(slide.partPath, slide.layoutPath ?? pres.newSlideLayoutPath!);
    rels.push(`<Relationship Id="rId1" Type="${NS_R}/slideLayout" Target="${XML_ESC(target)}"/>`);
  }
  for (const e of extra) {
    rels.push(`<Relationship Id="${XML_ESC(e.id)}" Type="${NS_R}/image" Target="${XML_ESC(relativeTarget(slide.partPath, e.path))}"/>`);
  }
  if (slide.notesPath && !rels.some((rel) => rel.includes('/notesSlide"'))) {
    const taken = new Set(rels.map((rel) => /Id="([^"]+)"/.exec(rel)?.[1] ?? ""));
    let n = 0;
    while (taken.has(`rId${++n}`));
    rels.push(`<Relationship Id="rId${n}" Type="${NS_R}/notesSlide" Target="${XML_ESC(relativeTarget(slide.partPath, slide.notesPath))}"/>`);
  }
  return `${DECL}<Relationships xmlns="${NS_PKG}">${rels.join("")}</Relationships>`;
}

/**
 * Reserves a slide-local relationship id per inserted picture and records it on the
 * shape, so the `p:pic` the shape tree builds names the relationship the rels part
 * carries. `layoutTakesRid1` covers a slide whose rels part is generated here.
 */
function assignPictureRels(pres: Presentation, slide: Slide, layoutTakesRid1: boolean): { id: string; path: string }[] {
  const taken = new Set<string>(pres.src?.rels(slide.partPath)?.keys() ?? []);
  if (layoutTakesRid1) taken.add("rId1");
  const out: { id: string; path: string }[] = [];
  let n = 0;
  for (const shape of walkShapes(slide.shapes)) {
    if (shape.kind !== "picture" || !shape.mediaPath) continue;
    while (taken.has(`rId${++n}`));
    shape.relId = `rId${n}`;
    taken.add(shape.relId);
    out.push({ id: shape.relId, path: shape.mediaPath });
  }
  return out;
}

function* walkShapes(shapes: Shape[]): Generator<Shape> {
  for (const shape of shapes) {
    yield shape;
    if (shape.children) yield* walkShapes(shape.children);
  }
}

function relativeTarget(fromPart: string, toPart: string): string {
  const fromDir = fromPart.slice(0, fromPart.lastIndexOf("/")).split("/");
  const to = toPart.split("/");
  let common = 0;
  while (common < fromDir.length && common < to.length - 1 && fromDir[common] === to[common]) common++;
  const up = fromDir.length - common;
  return "../".repeat(up) + to.slice(common).join("/");
}

// ---------------------------------------------------------------- content types

function contentTypesXml(pres: Presentation, files: Record<string, Uint8Array>): string {
  const pkg = pres.src;
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  if (pkg) {
    for (const [ext, type] of pkg.defaultTypes) defaults.set(ext, type);
    for (const [part, type] of pkg.overrideTypes) if (files[part.slice(1)]) overrides.set(part, type);
  }
  for (const slide of pres.slides) {
    overrides.set("/" + slide.partPath, `${CT_P}.slide+xml`);
    if (slide.notesPath && files[slide.notesPath]) overrides.set("/" + slide.notesPath, `${CT_P}.notesSlide+xml`);
  }
  // A deck built in the editor carries no source override for its main part.
  if (files[pres.presPath] && !overrides.has("/" + pres.presPath)) overrides.set("/" + pres.presPath, `${CT_P}.presentation.main+xml`);
  // Generated parts (blank master, layouts, theme, property parts, notes) need
  // their content type wherever the source package did not already declare one.
  const TYPE_BY_PART: [RegExp, string][] = [
    [/^ppt\/slideMasters\/slideMaster\d+\.xml$/, `${CT_P}.slideMaster+xml`],
    [/^ppt\/slideLayouts\/slideLayout\d+\.xml$/, `${CT_P}.slideLayout+xml`],
    [/^ppt\/theme\/theme\d+\.xml$/, "application/vnd.openxmlformats-officedocument.theme+xml"],
    [/^ppt\/presProps\.xml$/, `${CT_P}.presProps+xml`],
    [/^ppt\/viewProps\.xml$/, `${CT_P}.viewProps+xml`],
    [/^ppt\/tableStyles\.xml$/, `${CT_P}.tableStyles+xml`],
    [/^docProps\/core\.xml$/, "application/vnd.openxmlformats-package.core-properties+xml"],
    [/^docProps\/app\.xml$/, "application/vnd.openxmlformats-officedocument.extended-properties+xml"],
  ];
  for (const path of Object.keys(files)) {
    if (overrides.has("/" + path)) continue;
    const hit = TYPE_BY_PART.find(([re]) => re.test(path));
    if (hit) overrides.set("/" + path, hit[1]);
  }
  if (!defaults.has("rels")) defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml");
  if (!defaults.has("xml")) defaults.set("xml", "application/xml");
  for (const path of Object.keys(files)) {
    const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
    if (defaults.has(ext) || overrides.has("/" + path)) continue;
    const mime = IMAGE_MIME_BY_EXT[ext];
    if (mime) defaults.set(ext, mime);
  }
  const defs = [...defaults].map(([ext, type]) => `<Default Extension="${XML_ESC(ext)}" ContentType="${XML_ESC(type)}"/>`);
  const ovs = [...overrides].map(([part, type]) => `<Override PartName="${XML_ESC(part)}" ContentType="${XML_ESC(type)}"/>`);
  return `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defs.join("")}${ovs.join("")}</Types>`;
}
