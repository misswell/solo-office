import { strToU8, zipSync } from "fflate";
import type { Block, BorderSet, Inline, ParaStyle, Paragraph, ParagraphFormatRevision, PropSlot, RunStyle, RowInfo, Table, TableBorder, TableCell, WordDocument } from "./docx.js";
import { bytesOfDataUrl } from "./media.js";
import { BORDER_ORDER, ROW_REVISION_TAGS, preserve, revisionTag } from "./docx.js";
import { childSpans, elementContent } from "./opc.js";
import { setChildren } from "./doc-table-revisions.js";
import { commentDirty, commentElement, commentParaIds, commentsRootInfo, type DocComment } from "./comments.js";

/**
 * Serializes the document model back to WordprocessingML. Only the parts the
 * model edits are regenerated — the rest of the package is copied through, so
 * a save cannot quietly drop a header, a theme or a font table.
 */

const XML_ESC = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const EMU_PER_PX = 9525;

const CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  svg: "image/svg+xml",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
  bin: "application/vnd.openxmlformats-officedocument.oleObject",
};

/** Content types for the parts a save may have to create in a package that never had them. */
function fallbackType(path: string): string | undefined {
  const name = path.slice(path.lastIndexOf("/") + 1);
  const overrides: [RegExp, string][] = [
    [/^word\/document\.xml$/, `${CT_W}.document.main+xml`],
    [/^word\/styles\.xml$/, `${CT_W}.styles+xml`],
    [/^word\/numbering\.xml$/, `${CT_W}.numbering+xml`],
    [/^word\/settings\.xml$/, `${CT_W}.settings+xml`],
    [/^word\/fontTable\.xml$/, `${CT_W}.fontTable+xml`],
    [/^word\/webSettings\.xml$/, `${CT_W}.webSettings+xml`],
    [/^word\/theme\/theme\d+\.xml$/, "application/vnd.openxmlformats-officedocument.theme+xml"],
    [/^word\/header\d+\.xml$/, `${CT_W}.header+xml`],
    [/^word\/footer\d+\.xml$/, `${CT_W}.footer+xml`],
    [/^word\/footnotes\.xml$/, `${CT_W}.footnotes+xml`],
    [/^word\/endnotes\.xml$/, `${CT_W}.endnotes+xml`],
    [/^word\/comments\.xml$/, `${CT_W}.comments+xml`],
    [/^word\/commentsExtended\.xml$/, `${CT_W}.commentsExtended+xml`],
    [/^word\/commentsIds\.xml$/, `${CT_W}.commentsIds+xml`],
    [/^word\/people\.xml$/, `${CT_W}.people+xml`],
    [/^docProps\/core\.xml$/, "application/vnd.openxmlformats-package.core-properties+xml"],
    [/^docProps\/app\.xml$/, "application/vnd.openxmlformats-officedocument.extended-properties+xml"],
    [/^docProps\/(custom|customXml)\.xml$/, "application/vnd.openxmlformats-officedocument.custom-properties+xml"],
  ];
  for (const [re, type] of overrides) if (re.test(path)) return type;
  const ext = name.includes(".") ? name.slice(name.lastIndexOf(".") + 1).toLowerCase() : "";
  if (ext in CONTENT_TYPES) return CONTENT_TYPES[ext];
  return undefined;
}

/**
 * The paragraph mark's own `w:rPr`, carried as the file wrote it, with only the break's record replaced.
 * Accepting or refusing a paragraph-mark revision decides whether that one child is there; the font, size,
 * and colour sitting beside it belong to the mark and must survive the decision either way. `CT_RPr` puts
 * the record after the character properties, so it lands there and not at the front of the set.
 */
function withMarkRevision(p: Paragraph, slot: PropSlot): string {
  const prefix = /^<([A-Za-z_][-\w.]*):rPr/.exec(slot.xml)?.[1] ?? "w";
  const kids = childSpans(elementContent(slot.xml))
    .map((k) => ({ tag: k.tag, xml: k.xml }))
    .filter((k) => !ROW_REVISION_TAGS.has(k.tag));
  const mark = p.markRev;
  if (mark) {
    const tag = revisionTag(mark);
    const at = kids.findIndex((k) => k.tag === "rPrChange");
    kids.splice(at < 0 ? kids.length : at, 0, { tag, xml: `<${prefix}:${tag} ${revAttrs(mark)}/>` });
  }
  return setChildren(slot.xml, kids);
}

/**
 * `w:ind` from the fields the model holds, stating only the attributes one of them carries. The attribute
 * order is the one WPS writes, measured over the 29,889 `w:ind` elements in the 86 documents this repository's
 * shapes come from — `left`, then `right`, then the first line's two spellings, then `w:hanging`
 * (`left,leftChars,right,rightChars,firstLine,firstLineChars` 123 times, `left,hanging` 118).
 */
function indTag(s: ParaStyle): string {
  const attrs: string[] = [];
  const put = (name: string, value: number | undefined): void => {
    if (value !== undefined) attrs.push(` ${name}="${Math.round(value)}"`);
  };
  put("w:left", s.indentLeft);
  put("w:right", s.indentRight);
  put("w:firstLine", s.indentFirstLine);
  put("w:firstLineChars", s.indentFirstLineChars);
  put("w:hanging", s.indentHanging);
  return attrs.length ? `<w:ind${attrs.join("")}/>` : "";
}

/** `CT_PPrBase` children have a required order; Word repairs deviations but validators do not. */
function paragraphProps(p: Paragraph): string {
  // The set the file wrote, child by child: what the page draws comes from the fields, and the fields the
  // editor moves are pushed back into this set by `syncParaSlots`. Spelling the set from the fields instead
  // is how a save quietly removes `w:keepNext`, an outline level, or the paragraph mark's own font.
  if (p.pPrSlots?.length) {
    const kids: PropSlot[] = p.pPrSlots
      .map((s) => ({ tag: s.tag, xml: s.tag === "rPr" ? withMarkRevision(p, s) : s.xml }))
      .filter((s) => s.xml);
    const mark = p.markRev;
    if (mark && !kids.some((k) => k.tag === "rPr"))
      kids.push({ tag: "rPr", xml: `<w:rPr><w:${revisionTag(mark)} ${revAttrs(mark)}/></w:rPr>` });
    const fmt = p.revFormat;
    if (fmt) kids.push({ tag: "pPrChange", xml: paraChangeXml(fmt) });
    return kids.length ? `<w:pPr>${kids.map((k) => k.xml).join("")}</w:pPr>` : "";
  }
  const parts: string[] = [];
  if (p.styleId) parts.push(`<w:pStyle w:val="${XML_ESC(p.styleId)}"/>`);
  if (p.pageBreakBefore) parts.push("<w:pageBreakBefore/>");
  if (p.numPr) parts.push(`<w:numPr><w:ilvl w:val="${p.numPr.ilvl}"/><w:numId w:val="${p.numPr.numId}"/></w:numPr>`);
  const spacing: string[] = [];
  if (p.spacingBefore !== undefined) spacing.push(` w:before="${Math.round(p.spacingBefore)}"`);
  if (p.spacingAfter !== undefined) spacing.push(` w:after="${Math.round(p.spacingAfter)}"`);
  if (p.lineSpacing) spacing.push(` w:line="${Math.round(p.lineSpacing)}" w:lineRule="auto"`);
  if (spacing.length) parts.push(`<w:spacing${spacing.join("")}/>`);
  const ind = indTag(p);
  if (ind) parts.push(ind);
  if (p.align) parts.push(`<w:jc w:val="${XML_ESC(p.align)}"/>`);
  // The paragraph mark's own revision, then the properties record: `CT_PPr` puts `w:rPr` after the
  // base properties and `w:pPrChange` last, and a validator that reads them out of order rejects the part.
  const mark = p.markRev;
  if (mark) parts.push(`<w:rPr><w:${revisionTag(mark)} ${revAttrs(mark)}/></w:rPr>`);
  const fmt = p.revFormat;
  if (fmt) parts.push(paraChangeXml(fmt));
  return parts.length ? `<w:pPr>${parts.join("")}</w:pPr>` : "";
}

/**
 * A `w:pPrChange`: the previous set as the record wrote it when it carries one, because a reviewer decides
 * what the file said and not what this model could name. A record this editor signs on a paragraph it built
 * itself has no source text, so that one is still spelled from the eight properties the model holds.
 */
function paraChangeXml(fmt: ParagraphFormatRevision): string {
  const inner = fmt.prevSlots?.length ? fmt.prevSlots.map((s) => s.xml).join("") : paraStyleTags(fmt.prev);
  return `<w:pPrChange ${revAttrs(fmt)}><w:pPr${inner ? `>${inner}</w:pPr` : "/"}></w:pPrChange>`;
}

/** The properties a `w:pPrChange` carries: the same tags, in the same order, from the recorded set. */
function paraStyleTags(s: ParaStyle): string {
  const parts: string[] = [];
  if (s.styleId) parts.push(`<w:pStyle w:val="${XML_ESC(s.styleId)}"/>`);
  if (s.pageBreakBefore) parts.push("<w:pageBreakBefore/>");
  if (s.numPr) parts.push(`<w:numPr><w:ilvl w:val="${s.numPr.ilvl}"/><w:numId w:val="${s.numPr.numId}"/></w:numPr>`);
  const spacing: string[] = [];
  if (s.spacingBefore !== undefined) spacing.push(` w:before="${Math.round(s.spacingBefore)}"`);
  if (s.spacingAfter !== undefined) spacing.push(` w:after="${Math.round(s.spacingAfter)}"`);
  if (s.lineSpacing) spacing.push(` w:line="${Math.round(s.lineSpacing)}" w:lineRule="auto"`);
  if (spacing.length) parts.push(`<w:spacing${spacing.join("")}/>`);
  const ind = indTag(s);
  if (ind) parts.push(ind);
  if (s.align) parts.push(`<w:jc w:val="${XML_ESC(s.align)}"/>`);
  return parts.join("");
}

/** `CT_RPr` children likewise: rFonts, b, i, strike, color, sz, szCs, highlight, u, vertAlign. */
function runPropsInner(s: RunStyle): string {
  const parts: string[] = [];
  if (s.font) parts.push(`<w:rFonts w:ascii="${XML_ESC(s.font)}" w:hAnsi="${XML_ESC(s.font)}"/>`);
  if (s.bold !== undefined) parts.push(s.bold ? "<w:b/>" : '<w:b w:val="0"/>');
  if (s.italic !== undefined) parts.push(s.italic ? "<w:i/>" : '<w:i w:val="0"/>');
  if (s.strike) parts.push("<w:strike/>");
  if (s.color) parts.push(`<w:color w:val="${XML_ESC(s.color.replace(/^#/, "").toUpperCase())}"/>`);
  if (s.sizePt) parts.push(`<w:sz w:val="${Math.round(s.sizePt * 2)}"/><w:szCs w:val="${Math.round(s.sizePt * 2)}"/>`);
  if (s.highlight) parts.push(`<w:highlight w:val="${XML_ESC(s.highlight)}"/>`);
  if (s.underline) parts.push('<w:u w:val="single"/>');
  if (s.vertAlign) parts.push(`<w:vertAlign w:val="${s.vertAlign}"/>`);
  return parts.join("");
}

const runProps = (s: RunStyle): string => {
  const inner = runPropsInner(s);
  return inner ? `<w:rPr>${inner}</w:rPr>` : "";
};

/** The order `CT_RPr` gives its children; a property the model newly spells has to land in it. */
const RPR_ORDER = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
  "condense", "extend", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
  "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
  "rPrChange",
];
const rprRank = (tag: string): number => {
  const at = RPR_ORDER.indexOf(tag);
  return at < 0 ? RPR_ORDER.length : at;
};

/**
 * The `w:rPr` children the model has a field for, and how that field writes them. `w:sz` and `w:szCs` are one
 * field's two children: the size button writes both, so a size change moves the complex-script size with the
 * Latin one rather than leaving the run at two sizes. Everything not in this list has no field and keeps the
 * text the file wrote for it — `w:rStyle`, `w:kern`, `w:lang`, `w:bCs`, `w:spacing`, `w:shd`, `w:eastAsianLayout`.
 */
const RUN_PROPS: { key: keyof RunStyle; tags: string[]; spell: (r: RunStyle, prefix: string, source: string) => string }[] = [
  {
    key: "font",
    tags: ["rFonts"],
    spell: (r, prefix, source) =>
      !r.font ? "" : source
        ? putLatinFonts(source, prefix, r.font)
        : `<${prefix}:rFonts w:ascii="${XML_ESC(r.font)}" w:hAnsi="${XML_ESC(r.font)}"/>`,
  },
  {
    key: "bold",
    tags: ["b"],
    spell: (r, prefix) => (r.bold === undefined ? "" : r.bold ? `<${prefix}:b/>` : `<${prefix}:b w:val="0"/>`),
  },
  {
    key: "italic",
    tags: ["i"],
    spell: (r, prefix) => (r.italic === undefined ? "" : r.italic ? `<${prefix}:i/>` : `<${prefix}:i w:val="0"/>`),
  },
  { key: "strike", tags: ["strike"], spell: (r, prefix) => (r.strike ? `<${prefix}:strike/>` : "") },
  {
    key: "color",
    tags: ["color"],
    spell: (r, prefix) => (r.color ? `<${prefix}:color w:val="${XML_ESC(r.color.replace(/^#/, "").toUpperCase())}"/>` : ""),
  },
  {
    key: "sizePt",
    tags: ["sz", "szCs"],
    spell: (r, prefix) =>
      r.sizePt ? `<${prefix}:sz w:val="${Math.round(r.sizePt * 2)}"/><${prefix}:szCs w:val="${Math.round(r.sizePt * 2)}"/>` : "",
  },
  {
    key: "highlight",
    tags: ["highlight"],
    spell: (r, prefix) => (r.highlight ? `<${prefix}:highlight w:val="${XML_ESC(r.highlight)}"/>` : ""),
  },
  { key: "underline", tags: ["u"], spell: (r, prefix) => (r.underline ? `<${prefix}:u w:val="single"/>` : "") },
  { key: "vertAlign", tags: ["vertAlign"], spell: (r, prefix) => (r.vertAlign ? `<${prefix}:vertAlign w:val="${r.vertAlign}"/>` : "") },
];

/**
 * `w:rFonts` with the Latin faces put onto the element as it was written. Picking a font owns `w:ascii` and
 * `w:hAnsi`; `w:eastAsia` and `w:cs` name the faces on the same element that no field holds, and a font change
 * is not their owner either.
 */
function putLatinFonts(xml: string, prefix: string, font: string): string {
  const gt = xml.indexOf(">");
  if (gt < 0) return xml;
  const head = xml.slice(0, gt);
  const selfClosed = head.endsWith("/");
  const open = (selfClosed ? head.slice(0, -1) : head).replace(/\s[\w.-]+:(?:ascii|hAnsi)="[^"]*"/g, "");
  const pair = ` ${prefix}:ascii="${XML_ESC(font)}" ${prefix}:hAnsi="${XML_ESC(font)}"`;
  return `${open}${pair}${selfClosed ? "/>" : ">"}` + (selfClosed ? "" : xml.slice(gt + 1));
}

/**
 * The run's `w:rPr` the way the file wrote it, with only the children whose field actually moved re-spelled.
 * The same split `paragraphProps` makes for a `w:pPr` and `applyCellProps` for a `w:tcPr`: the set belongs to
 * the document, and a property the model cannot explain is not the model's to remove. Rebuilding it from nine
 * fields is what took the kerning, the complex-script face and the language out of every run in a document.
 */
function carriedRunProps(r: Inline): string {
  const slots = r.rPrSlots ?? [];
  // A prefix is taken from the set rather than assumed: `w` is only an alias for the words namespace where the
  // part binds it, and a child written with another prefix than its neighbours reads as another element.
  const prefix = /^<([A-Za-z_][-\w.]*):/.exec(slots[0]?.xml ?? "")?.[1] ?? "w";
  const same = (a: unknown, b: unknown): boolean =>
    a === b || (typeof a === "object" && typeof b === "object" && JSON.stringify(a) === JSON.stringify(b));
  const drops = new Set<PropSlot>();
  const inserts: PropSlot[] = [];
  for (const prop of RUN_PROPS) {
    // A field that has not moved off its read snapshot keeps the source text, so an attribute the model has no
    // field for — `w:kern`, `w:themeColor`, `w:eastAsia` — survives a save that never touched it.
    if (same(r.asRead?.[prop.key], r[prop.key])) continue;
    const own = slots.filter((s) => prop.tags.includes(s.tag));
    const xml = prop.spell(r, prefix, own[0]?.xml ?? "");
    for (const s of own) drops.add(s);
    if (xml) inserts.push({ tag: prop.tags[0]!, xml });
  }
  const kids = slots.filter((s) => !drops.has(s));
  for (const ins of inserts) {
    const at = kids.findIndex((k) => rprRank(k.tag) > rprRank(ins.tag));
    kids.splice(at < 0 ? kids.length : at, 0, ins);
  }
  const fmt = r.revFormat;
  if (fmt) {
    // The record keeps the set as the file wrote it: `prev` names only what a field can hold, and a save that
    // rebuilt the record from those fields would leave a reviewer refusing a change no document ever made.
    const prevInner = fmt.prevSlots?.length
      ? `<${prefix}:rPr>${fmt.prevSlots.map((s) => s.xml).join("")}</${prefix}:rPr>`
      : runProps(fmt.prev) || `<${prefix}:rPr/>`;
    kids.push({ tag: "rPrChange", xml: `<${prefix}:rPrChange ${revAttrs(fmt)}>${prevInner}</${prefix}:rPrChange>` });
  }
  // A run whose `w:rPr` the file left empty comes back empty rather than with no properties at all.
  return kids.length ? `<${prefix}:rPr>${kids.map((k) => k.xml).join("")}</${prefix}:rPr>` : `<${prefix}:rPr/>`;
}

/**
 * The run's properties with its `w:rPrChange` appended, which is where OOXML keeps a formatting
 * revision: inside the same `w:rPr`, after the properties the run now has. The record has to hold a
 * `w:rPr` of its own even when nothing was set before.
 */
function revisedRunProps(inline: Inline): string {
  if (inline.rPrSlots) return carriedRunProps(inline);
  const fmt = inline.revFormat;
  const inner = runPropsInner(inline) + (fmt ? `<w:rPrChange ${revAttrs(fmt)}>${runProps(fmt.prev) || "<w:rPr/>"}</w:rPrChange>` : "");
  return inner ? `<w:rPr>${inner}</w:rPr>` : "";
}

function textRun(inline: Inline): string {
  if (inline.image) return imageRun(inline, inline.image.width, inline.image.height, "");
  const props = revisedRunProps(inline);
  const body = inline.tab ? "<w:tab/>" : inline.break ? breakXml(inline) : inline.rev?.kind === "del" ? delTextTag(inline.text) : textTag(inline.text);
  return `<w:r>${props}${body}</w:r>`;
}

/** A break goes back as the file wrote it — its type, its `w:clear`, all of it. Ours own is the plain soft return. */
const breakXml = (inline: Inline): string => inline.brXml ?? "<w:br/>";

const textTag = (text: string): string => (text === "" ? "" : `<w:t${preserve(text)}>${XML_ESC(text)}</w:t>`);

/** Deleted text is spelled `w:delText`, not `w:t` — Word reads the two differently. */
const delTextTag = (text: string): string => (text === "" ? "" : `<w:delText${preserve(text)}>${XML_ESC(text)}</w:delText>`);

function imageRun(inline: Inline, width: number, height: number, rid: string): string {
  const img = inline.image!;
  // The drawing the file wrote travels whole while nothing about the picture has moved. Five fields describe
  // where a picture sits and how big it is; the frame it floats in, its wrap, its shadow, its crop and the
  // blip's own extension list are not among them, and re-spelling the drawing from the fields downgraded
  // every floating picture to an inline one and dropped the rest. Only the relationship id is the writer's.
  const carried = inline.drawing;
  if (rid && carried && carried.width === width && carried.height === height && carried.name === img.name) {
    return `<w:r>${revisedRunProps(inline)}${remapEmbed(carried.xml, carried.embed, rid)}</w:r>`;
  }
  const cx = Math.round(width * EMU_PER_PX);
  const cy = Math.round(height * EMU_PER_PX);
  const name = XML_ESC(img.name || "图片");
  const drawing =
    // Our own drawing declares the prefixes it uses where they are used: a document whose root the file
    // wrote without `xmlns:a` keeps that root, and a picture inserted into it still resolves.
    `<w:drawing><wp:inline xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}" xmlns:r="${NS_R}" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="${name}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="${NS_PIC}"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  return `<w:r>${revisedRunProps(inline)}${drawing}</w:r>`;
}

/**
 * A carried drawing's citation, pointed at the relationship id this save issued. Everything else in its bytes
 * goes back untouched — the id is the only thing in a drawing that belongs to the package rather than to the
 * picture, and a save that renumbered its relationships would otherwise leave the shape citing nothing.
 */
function remapEmbed(xml: string, from: string | undefined, rid: string): string {
  if (!from || from === rid) return xml;
  let out = xml;
  for (const attr of ["r:embed", "r:link"]) {
    const needle = `${attr}="${from}"`;
    if (out.includes(needle)) out = out.split(needle).join(`${attr}="${rid}"`);
  }
  return out;
}

/**
 * Anchors keep their place between runs: the range markers are paragraph children and
 * the reference mark is a run of its own, which is how Word writes both.
 */
function anchorXml(anchor: NonNullable<Inline["anchor"]>, from: Inline): string {
  if (anchor.kind === "start") return `<w:commentRangeStart w:id="${anchor.id}"/>`;
  if (anchor.kind === "end") return `<w:commentRangeEnd w:id="${anchor.id}"/>`;
  // A mark the file wrote keeps the properties the file wrote with it — Word gives this run the
  // `CommentReference` style, but a document may say more about it than that, and re-spelling the run from one
  // field is how a save loses markup it never read. Only a mark this editor inserts has no source to answer to,
  // and that one takes the shape Word writes.
  const slots = from.rPrSlots;
  if (!slots) return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${anchor.id}"/></w:r>`;
  const prefix = /^<([A-Za-z_][-\w.]*):/.exec(slots[0]?.xml ?? "")?.[1] ?? "w";
  const props = slots.length ? `<${prefix}:rPr>${slots.map((s) => s.xml).join("")}</${prefix}:rPr>` : `<${prefix}:rPr/>`;
  return `<w:r>${props}<w:commentReference w:id="${anchor.id}"/></w:r>`;
}

/** Consecutive inlines sharing a link become one `w:hyperlink`. */
function inlinesXml(inlines: Inline[], rels: RelBuilder): string {
  const out: string[] = [];
  let i = 0;
  while (i < inlines.length) {
    // Carried bytes go back as they came, at the place the reader found them: a bookmark or a field is not
    // this writer's to reinterpret, and rebuilding one from attributes would lose whatever it also carried.
    if (inlines[i].carry) {
      out.push(inlines[i].carry!.xml);
      i += 1;
      continue;
    }
    if (inlines[i].anchor) {
      out.push(anchorXml(inlines[i].anchor!, inlines[i]));
      i += 1;
      continue;
    }
    const rev = inlines[i].rev;
    if (rev) {
      // One wrapper per revision, as Word writes it: the runs it covers stay runs. An anchor is a
      // paragraph marker, not revised text, so it stays outside the wrapper instead of becoming "".
      const group: Inline[] = [];
      while (i < inlines.length && !inlines[i].anchor && sameRevision(inlines[i].rev, rev)) group.push(inlines[i++]);
      out.push(`<w:${revisionTag(rev)} ${revAttrs(rev)}>${group.map((r) => runWrapped(r, rels)).join("")}</w:${revisionTag(rev)}>`);
      continue;
    }
    const link = inlines[i].hyperlink;
    if (!link) {
      out.push(runWrapped(inlines[i], rels));
      i += 1;
      continue;
    }
    // Runs join one `w:hyperlink` only while the element itself would be spelled the same: two links to the
    // same target carrying different attributes are two elements, or one run's attributes would be written
    // over its neighbour's.
    const extras = inlines[i].linkAttrs;
    const group: Inline[] = [];
    while (i < inlines.length && inlines[i].hyperlink === link && inlines[i].linkAttrs === extras) group.push(inlines[i++]);
    const inner = group.map((r) => runXml(r, rels)).join("");
    out.push(`<w:hyperlink ${linkAttrs(link, rels, extras)}>${inner}</w:hyperlink>`);
  }
  return out.join("");
}

/**
 * A link to a place inside this document is spelled `w:anchor` and has no relationship at all; a link to
 * the outside is a relationship id. Writing an in-document link as `r:id` points the reader at a file
 * address instead of the bookmark, which is how a table of contents entry stops being one after a save.
 * The target is spelled first and the attributes the file carried after, the way Word writes them
 * (`r:id="rId4" w:history="1"`), because the target is this model's own field while the rest travel as text.
 */
const linkAttrs = (link: string, rels: RelBuilder, extras?: string): string =>
  `${link.startsWith("#") ? `w:anchor="${XML_ESC(link.slice(1))}"` : `r:id="${rels.hyperlink(link)}"`}${extras ? ` ${extras}` : ""}`;

const runXml = (r: Inline, rels: RelBuilder): string =>
  r.image ? imageRun(r, r.image.width, r.image.height, rels.image(r.image.path)) : textRun(r);

/** A link inside a revision keeps its own `w:hyperlink`, which is where Word puts it. */
const runWrapped = (r: Inline, rels: RelBuilder): string =>
  // Carried bytes are the element as it came. Inside a revision they belong in the wrapper rather than beside
  // it, because a run put outside is read as text nobody revised and the record next to it loses its cover.
  r.carry ? r.carry.xml : r.hyperlink ? `<w:hyperlink ${linkAttrs(r.hyperlink, rels, r.linkAttrs)}>${runXml(r, rels)}</w:hyperlink>` : runXml(r, rels);

const sameRevision = (a: Inline["rev"], b: Inline["rev"]): boolean =>
  a !== undefined && b !== undefined && a.kind === b.kind && a.move === b.move && a.id === b.id && a.author === b.author;

const revAttrs = (rev: { id: number; author: string; date?: string }): string =>
  `w:id="${rev.id}" w:author="${XML_ESC(rev.author)}"${rev.date ? ` w:date="${XML_ESC(rev.date)}"` : ""}`;

/** One edge of a border set, spelled the way `w:tblBorders` and `w:tcBorders` spell it. */
const borderAttrs = (b: TableBorder): string =>
  `w:val="${XML_ESC(b.val)}"${b.size !== undefined ? ` w:sz="${b.size}"` : ""}${b.space !== undefined ? ` w:space="${b.space}"` : ""}${b.color ? ` w:color="${XML_ESC(b.color)}"` : ""}`;

const bordersXml = (tag: string, set?: BorderSet): string => {
  if (!set) return "";
  const sides = BORDER_ORDER.filter((side) => set[side]);
  return sides.length ? `<w:${tag}>${sides.map((side) => `<w:${side} ${borderAttrs(set[side]!)}/>`).join("")}</w:${tag}>` : "";
};

/**
 * A cell's `w:tcPr`: what the file wrote, unless the model built this cell itself and has only the
 * typed fields to go on. Children follow the order `CT_TcPrBase` gives them.
 */
function cellPropsXml(cell: TableCell): string {
  if (cell.tcPrXml) return cell.tcPrXml;
  const props: string[] = [];
  if (cell.widthTwips) props.push(`<w:tcW w:w="${Math.round(cell.widthTwips)}" w:type="dxa"/>`);
  if (cell.spanCols && cell.spanCols > 1) props.push(`<w:gridSpan w:val="${cell.spanCols}"/>`);
  if (cell.spanRows !== undefined) props.push(cell.spanRows === 1 ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>");
  const borders = bordersXml("tcBorders", cell.borders);
  if (borders) props.push(borders);
  if (cell.shading) props.push(`<w:shd w:val="clear" w:color="auto" w:fill="${XML_ESC(cell.shading.replace(/^#/, "").toUpperCase())}"/>`);
  if (cell.noWrap) props.push("<w:noWrap/>");
  if (cell.textDirection) props.push(`<w:textDirection w:val="${XML_ESC(cell.textDirection)}"/>`);
  if (cell.valign) props.push(`<w:vAlign w:val="${XML_ESC(cell.valign)}"/>`);
  return props.length ? `<w:tcPr>${props.join("")}</w:tcPr>` : "";
}

function cellXml(cell: TableCell, rels: RelBuilder): string {
  const text = cell.paragraphs.length ? cellBlocksXml(cell.paragraphs, rels) : "<w:p/>";
  return `<w:tc${cell.attrsXml ?? ""}>${cellPropsXml(cell)}${text}</w:tc>`;
}

/**
 * A row's `w:trPr`. The children the file wrote come back one by one in the order they sat; only the
 * row's own record is rebuilt, because whether it survives at all is what accepting or rejecting decides.
 * A row the editor made has no such children, so its height and flags are spelled from the model.
 */
function rowPropsXml(row?: RowInfo): string {
  if (!row) return "";
  const kids = row.slots?.map((s) => s.xml) ?? [];
  if (!kids.length) {
    if (row.keepTogether) kids.push("<w:cantSplit/>");
    if (row.heightTwips !== undefined)
      kids.push(`<w:trHeight w:val="${Math.round(row.heightTwips)}"${row.heightRule ? ` w:hRule="${XML_ESC(row.heightRule)}"` : ""}/>`);
    if (row.header) kids.push("<w:tblHeader/>");
    if (row.hidden) kids.push("<w:hidden/>");
  }
  if (row.rev) {
    const at = kids.findIndex((k) => /<w:trPrChange[\s/>]/.test(k));
    kids.splice(at < 0 ? kids.length : at, 0, `<w:${revisionTag(row.rev)} ${revAttrs(row.rev)}/>`);
  }
  return kids.length ? `<w:trPr>${kids.join("")}</w:trPr>` : "";
}

/** `w:tblGrid` as the file wrote it, or one sized from what the model knows of the new grid. */
function gridXml(t: Table): string {
  const columns = Math.max(t.columns, ...t.rows.map((cells) => cells.reduce((n, c) => n + (c.spanCols ?? 1), 0)));
  if (t.gridXml && (t.columnWidths?.length ?? 0) === columns) return t.gridXml;
  const widths = t.columnWidths?.length === columns ? t.columnWidths : undefined;
  const grid = t.rows[0] ?? [];
  const cols = Array.from({ length: columns }, (_, i) => {
    const width = widths?.[i] ?? grid[i]?.widthTwips;
    return `<w:gridCol${width ? ` w:w="${Math.round(width)}"` : ""}/>`;
  }).join("");
  return `<w:tblGrid>${cols}</w:tblGrid>`;
}

/** `w:tblPr` for a table the editor built; a table off the page returns what the page wrote. */
function tablePropsXml(t: Table): string {
  if (t.tblPrXml) return t.tblPrXml;
  const props = [
    t.styleId ? `<w:tblStyle w:val="${XML_ESC(t.styleId)}"/>` : "",
    // `w:w="0" w:type="auto"` is what "no width declared" already means, so a table that never said one
    // gains nothing from being given the tag and a save that writes it is inventing a property.
    t.widthTwips !== undefined ? `<w:tblW w:w="${Math.round(t.widthTwips)}" w:type="${XML_ESC(t.widthType ?? "dxa")}"/>` : "",
    t.align ? `<w:jc w:val="${XML_ESC(t.align)}"/>` : "",
    t.indentTwips !== undefined ? `<w:tblInd w:w="${Math.round(t.indentTwips)}" w:type="dxa"/>` : "",
    bordersXml("tblBorders", t.borders),
    t.layout ? `<w:tblLayout w:type="${XML_ESC(t.layout)}"/>` : "",
  ].join("");
  return props ? `<w:tblPr>${props}</w:tblPr>` : "";
}

function tableXml(t: Table, rels: RelBuilder): string {
  // A row that somebody added or removed signs itself on the row, before any of the row's cells.
  const between = t.betweenRows ?? [];
  const after = (i: number): string =>
    between
      .filter((s) => s.afterRow === i)
      .map((s) => s.xml)
      .join("");
  const rows = t.rows
    .map((cells, i) => {
      const row = cells[0]?.row;
      // The row's unmodelled children go back between the siblings they were read next to. WPS puts the
      // `w:tblPrEx` band first, ahead of `w:trPr`, so a save that returns it after the cells rewrites
      // every row of every table in the wrong order.
      const ex = row?.extras ?? [];
      const head = ex.filter((s) => s.beforeProps).map((s) => s.xml).join("");
      const mid = ex.filter((s) => !s.beforeProps && s.afterCells === 0).map((s) => s.xml).join("");
      const afterCell = (n: number): string => ex.filter((s) => !s.beforeProps && s.afterCells === n).map((s) => s.xml).join("");
      const orphaned = ex.filter((s) => !s.beforeProps && s.afterCells > cells.length).map((s) => s.xml).join("");
      return `<w:tr${row?.attrsXml ?? ""}>${head}${rowPropsXml(row)}${mid}${cells.map((c, n) => cellXml(c, rels) + afterCell(n + 1)).join("")}${orphaned}</w:tr>${after(i)}`;
    })
    .join("");
  // Markup that sat before the first row keeps that place; markup past the last row the table has now (its
  // rows were deleted) still goes back at the end rather than vanishing with the row it followed.
  return `<w:tbl>${tablePropsXml(t)}${gridXml(t)}${after(-1)}${rows}${between.filter((s) => s.afterRow >= t.rows.length).map((s) => s.xml).join("")}</w:tbl>`;
}

function blockXml(block: Block, rels: RelBuilder): string {
  if (block.kind === "table") return tableXml(block, rels);
  // A container the model does not explain goes back exactly as it came, in the place the reader found it.
  if (block.kind === "carried") return block.xml;
  return `<w:p${block.attrsXml ?? ""}>${paragraphProps(block)}${inlinesXml(block.inlines, rels)}</w:p>`;
}

/**
 * A cell's block children with the rule the Word family enforces already satisfied: no `w:tbl` closes the
 * cell, and no two sit side by side. A cell that ends on a table is read as damage, and two tables with
 * nothing between them are read as one table, so the paragraph that separates or closes them is spelled here
 * instead of relying on every caller to remember it. The body keeps whatever order the file gave it — this
 * normalisation is for cells, where readers are known to require it — and a cell written this way reads back
 * with the paragraph already in the model, so saving the same document twice writes the same bytes.
 */
function cellBlocksXml(blocks: Block[], rels: RelBuilder): string {
  const out: string[] = [];
  blocks.forEach((block, i) => {
    out.push(blockXml(block, rels));
    if (block.kind !== "table") return;
    const next = blocks[i + 1];
    if (!next || next.kind === "table") out.push("<w:p/>");
  });
  return out.join("");
}

function bodyXml(doc: WordDocument, rels: RelBuilder): string {
  return doc.blocks.map((b) => blockXml(b, rels)).join("") + sectPrXml(doc);
}

function sectPrXml(doc: WordDocument): string {
  const preserved = doc.sectionPropertiesXml();
  if (preserved) return preserved;
  // Nothing to be faithful to and nothing the model read: a file whose body never declared a section stays
  // that way. Writing `pgSz`/`pgMar` here would pin page geometry the reader was free to choose — and the
  // spelling this function falls back to (`<w:pgSz/>` with no width or height) declares a page of unknown
  // size, which is worse than saying nothing. A document created here has no file behind it and keeps the
  // page the app opens with.
  //
  // `sectReferenceTags()` is the other half of the question, and it says "" for more than "nothing owed":
  // a band whose section is a paragraph's break is cited there by the model itself, so that section's
  // bytes already travel and spelling a second one beside it is how a file gained a `sectPr` it never had.
  //
  // A third case decides the other way: 「页面设置」 on a file that never spelled a section. The numbers the user
  // typed are then the whole reason a section exists, and a save that dropped them would close the dialog on
  // a page that never moved.
  if (doc.source() && !doc.sectReferenceTags() && !doc.pageSetupStated()) return "";
  const s = doc.sections[0];
  // `w:orient` is spelled only for landscape. None of the 93 portrait sections in the 86 files this machine's
  // WPS wrote says so, while every one of its 22 landscape sections does and puts the long edge in `w:w`.
  const size = s
    ? ` w:w="${s.pageWidthTwips}" w:h="${s.pageHeightTwips}"${s.pageWidthTwips > s.pageHeightTwips ? ' w:orient="landscape"' : ""}`
    : "";
  // The three distances §17.3.2.13 counts from the paper edge follow the four body margins in the order WPS
  // spells all seven, and a section that holds no number for one says nothing about it here.
  const dist = (key: "header" | "footer" | "gutter"): string => (s?.margins[key] === undefined ? "" : ` w:${key}="${s.margins[key]}"`);
  const mar = s
    ? ` w:top="${s.margins.top}" w:right="${s.margins.right}" w:bottom="${s.margins.bottom}" w:left="${s.margins.left}"${dist("header")}${dist("footer")}${dist("gutter")}`
    : ' w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"';
  // Header/footer references lead the section's children in schema order.
  return `<w:sectPr>${doc.sectReferenceTags()}<w:pgSz${size}/><w:pgMar${mar}/></w:sectPr>`;
}

/** The prefixes this writer's own markup needs; whatever else the file declared is added to them. */
const WRITER_PREFIXES = new Set(["w", "r", "wp", "a", "pic"]);

function documentXml(doc: WordDocument, rels: RelBuilder): string {
  const body = bodyXml(doc, rels);
  const carried = doc.documentRootAttributes();
  if (carried) {
    // The file's own root goes back untouched — every declaration it chose, its order, and any attribute
    // this model never read. Only a prefix the tags below actually use and the root left unbound is added,
    // and `bound` reads the carried text rather than the model's prefix map so a single-quoted declaration
    // in the original is still seen as bound.
    const bound = new Set([...carried.matchAll(/\bxmlns:([\w.-]+)\s*=\s*["'][^"']*["']/g)].map((m) => m[1]!));
    const need = (prefix: string, uri: string): string => (bound.has(prefix) ? "" : ` xmlns:${prefix}="${uri}"`);
    return `${DECL}<w:document${carried}${need("w", NS_W)}${need("r", NS_R)}><w:body>${body}</w:body></w:document>`;
  }
  // Carried property sets can name prefixes no tag this writer builds uses (`w14:`, `mc:`). Declaring
  // only our own would save a table whose markup no longer resolves, and dropping the root's
  // `mc:Ignorable` would turn an extension a reader used to skip into one it has to refuse.
  const declared = new Map(Object.entries(doc.documentNamespaces()).filter(([prefix]) => !WRITER_PREFIXES.has(prefix)));
  const ignorable = doc.documentIgnorable();
  if (ignorable && !declared.has("mc")) declared.set("mc", NS_MC);
  const namespaces = [...declared].map(([prefix, uri]) => ` xmlns:${prefix}="${XML_ESC(uri)}"`).join("");
  const skip = ignorable ? ` mc:Ignorable="${XML_ESC(ignorable)}"` : "";
  // `w:` and `r:` are in every body this writer builds. The drawing prefixes go back only where the file
  // itself declared them on the root: 14 of the 86 table-bearing files declare `a` on the drawing rather than
  // on `w:document`, and a save that "helpfully" lifts it to the root has invented a declaration the file
  // never wrote. Markup this writer generates carries its own declarations (see `imageRun`), so it needs none.
  const fromSource = new Set(Object.keys(doc.documentNamespaces()));
  const declares = (prefix: string, uri: string): string => (fromSource.has(prefix) ? ` xmlns:${prefix}="${XML_ESC(uri)}"` : "");
  return (
    `${DECL}<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"${declares("wp", NS_WP)}${declares("a", NS_A)}${declares("pic", NS_PIC)}${namespaces}${skip}>` +
    `<w:body>${body}</w:body></w:document>`
  );
}

function relsXml(rels: RelBuilder): string {
  const tags = rels.entries.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${XML_ESC(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`).join("");
  return `${DECL}<Relationships xmlns="${NS_PKG}">${tags}</Relationships>`;
}

/**
 * A comment body rebuilt from the model is plain formatted runs; a comment nobody edited is
 * written back from the source text, so its rich runs, its `w:paraId` links and any markup
 * this model does not read stay exactly as they were.
 */
function commentsXml(list: DocComment[], rootAttrs: string): string {
  const root = commentsRootInfo(rootAttrs, list.some((c) => !!c.paraIds?.length));
  const body = list
    .map((c) => {
      const ids = c.paraIds ?? commentParaIds(c.raw);
      // The source bytes are reusable only while their paragraph-id keys still match the
      // model: an id the extension parts key on has to exist in comments.xml too, or Word
      // is pointed at a paragraph that is not there. Without a w14 namespace the ids are
      // inert, so keeping the original bytes — rich runs and all — still wins.
      const keysMatch = root.w14 === undefined || JSON.stringify(ids) === JSON.stringify(commentParaIds(c.raw));
      return c.raw !== "" && !commentDirty(c) && keysMatch ? c.raw : commentElement(c, root, commentRuns, ids);
    })
    .join("");
  return `${DECL}<${root.prefix}:comments${root.attrs}>${body}</${root.prefix}:comments>`;
}

/** Text runs only — a body holding a table, picture or field is never editable, so it never comes here. */
const commentRuns = (runs: Inline[]): string => runs.filter((r) => !r.anchor).map(textRun).join("");

interface RelEntry {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

/**
 * Relationship ids the file already wrote are kept; a save allocates its own above the highest of them,
 * and only for a part or a target the source package never related.
 */
class RelBuilder {
  readonly entries: RelEntry[] = [];
  private readonly byKey = new Map<string, string>();
  /** Carried relationships keyed on the *part* they resolve to, so a different spelling of the same
   *  target (`media/x.png` against `/word/media/x.png`) is recognised as the relationship it is. */
  private readonly byPath = new Map<string, string>();
  private next: number;

  constructor(carried: { id: string; type: string; target: string; targetMode: string; path?: string }[]) {
    for (const rel of carried) {
      this.entries.push({ id: rel.id, type: rel.type, target: rel.target, external: rel.targetMode === "External" });
      this.byKey.set(`${rel.type}|${rel.target}`, rel.id);
      if (rel.path && rel.targetMode !== "External" && !this.byPath.has(`${rel.type}|${rel.path}`)) {
        this.byPath.set(`${rel.type}|${rel.path}`, rel.id);
      }
    }
    const highest = carried.reduce((max, r) => Math.max(max, Number(/(\d+)$/.exec(r.id)?.[1] ?? 0)), 0);
    this.next = highest + 1;
  }

  private add(type: string, target: string, external?: boolean): string {
    const key = `${type}|${target}`;
    const seen = this.byKey.get(key);
    if (seen) return seen;
    const id = `rId${this.next++}`;
    this.entries.push({ id, type, target, external });
    this.byKey.set(key, id);
    return id;
  }

  hyperlink(url: string): string {
    return this.add(`${NS_R}/hyperlink`, url, true);
  }

  /** The comments part the body's anchors point at. */
  comments(target: string): string {
    return this.add(`${NS_R}/comments`, target);
  }

  /** A part a save creates that the source package never had, related under its own type. */
  relation(type: string, target: string): string {
    return this.add(type, target);
  }

  /**
   * A picture's relationship. The file's own id comes back whenever this save points at a part the source
   * already related — the part, not its spelling, because `/word/media/x.png` and `media/x.png` are one part
   * and a save that issued a second id for it orphaned the first while markup that travels by the byte kept
   * citing it. A picture inserted here is a part the package never had, and gets a fresh id.
   */
  image(path: string): string {
    const carried = this.byPath.get(`${NS_R}/image|${path}`);
    if (carried) return carried;
    return this.add(`${NS_R}/image`, path.replace(/^word\//, ""));
  }
}

/**
 * The content types as the package wrote them, plus only the declarations this save actually needs.
 * `[Content_Types].xml` is a *set of promises*, not a per-part listing: one `Default Extension="rels"`
 * covers every relationship part, and one `Default Extension="xml" ContentType="application/xml"` covers
 * every plain-XML part the document has no special type for. Spelling the table again from part names
 * invents an `Override` for each of those (`/_rels/.rels`, `/word/settings.xml`) and drops a `Default`
 * this writer has no template entry for — `bin`, `wmf`, a thumbnail format — which un-types bytes the
 * package still carries. So the source table travels: an `Override` leaves only when its part leaves the
 * package, and only a part this package never had gets a new promise, the narrowest one that fits.
 */
function contentTypesXml(carried: string, partsAtRead: Set<string>, paths: { path: string; type?: string }[]): string {
  const decl = /^<\?xml[\s\S]*?\?>/.exec(carried)?.[0];
  const rest = carried.slice((decl ?? "").length);
  // Whitespace between the declaration and the root is the source's own, and it stays. The root match has to
  // start at the root rather than at that whitespace, or this falls through to spelling the table again.
  const gap = /^\s*/.exec(rest)?.[0] ?? "";
  const inner = rest.slice(gap.length);
  const root = /^<[^>]*>/.exec(inner)?.[0];
  const close = inner.lastIndexOf("</");
  if (!root || close < 0) return spellContentTypes(paths);
  const attr = (xml: string, name: string): string => new RegExp(`\\s${name}="([^"]*)"`).exec(xml)?.[1] ?? "";
  const extOf = (path: string): string => (path.includes(".") ? path.slice(path.lastIndexOf(".") + 1).toLowerCase() : "");
  const spans = childSpans(inner.slice(root.length, close));
  const kept = spans.filter((s) =>
    // An extension-wide promise covers a *kind* of part, not one part: it stays whether or not this save
    // still has a part of that extension, because the source wrote it and nothing about it moved.
    s.tag === "Default" ? true : s.tag === "Override" ? paths.some((p) => `/${p.path}` === attr(s.xml, "PartName")) : true,
  );
  const defaults = new Map(kept.filter((s) => s.tag === "Default").map((s) => [attr(s.xml, "Extension").toLowerCase(), attr(s.xml, "ContentType")]));
  const overridden = new Set(kept.filter((s) => s.tag === "Override").map((s) => attr(s.xml, "PartName")));
  const added: string[] = [];
  for (const { path, type } of paths) {
    if (partsAtRead.has(path)) continue; // the source table already made this promise; re-spelling it is not fidelity
    const part = `/${path}`;
    if (overridden.has(part)) continue;
    const ext = extOf(path);
    const resolved = type ?? fallbackType(path);
    if (!resolved) continue;
    const covered = defaults.get(ext);
    if (covered === resolved) continue;
    if (ext in CONTENT_TYPES && CONTENT_TYPES[ext] === resolved && !path.endsWith(".xml")) {
      defaults.set(ext, resolved);
      added.push(`<Default Extension="${ext}" ContentType="${resolved}"/>`);
      continue;
    }
    overridden.add(part);
    added.push(`<Override PartName="${part}" ContentType="${resolved}"/>`);
  }
  return `${decl ?? DECL}${gap}${root}${kept.map((s) => s.xml).join("")}${added.join("")}</Types>`;
}

/** A package this writer built from scratch: no source table to travel, so the names spell the set. */
function spellContentTypes(paths: { path: string; type?: string }[]): string {
  const defaults = ["rels", "xml"].map((ext) => `<Default Extension="${ext}" ContentType="${ext === "rels" ? "application/vnd.openxmlformats-package.relationships+xml" : "application/xml"}"/>`);
  const mediaExt = new Set<string>();
  const overrides: string[] = [];
  for (const { path, type } of paths) {
    const ext = path.includes(".") ? path.slice(path.lastIndexOf(".") + 1).toLowerCase() : "";
    const resolved = type ?? fallbackType(path);
    // An extension-wide default beats an override for media, which can repeat per slide.
    if (resolved && ext in CONTENT_TYPES && !path.endsWith(".xml")) {
      mediaExt.add(ext);
      continue;
    }
    if (resolved) overrides.push(`<Override PartName="/${path}" ContentType="${resolved}"/>`);
  }
  const mediaDefaults = [...mediaExt].map((ext) => `<Default Extension="${ext}" ContentType="${CONTENT_TYPES[ext]}"/>`);
  return `${DECL}<Types xmlns="${NS_CT}">${defaults.join("")}${mediaDefaults.join("")}${overrides.join("")}</Types>`;
}

/**
 * The package root's own relationships, for a file that came with nothing usable.
 *
 * A promise has to be kept by a part: `docProps/core.xml` travels here with the source package when there is
 * one, and a document built from scratch has no metadata part to cite. Emitting the citation anyway produced a
 * relationship that resolves to nothing — which is what a reader sees on 文件 → 新建, the one path with no
 * source bytes to carry. The fix is the promise, not a fabricated part.
 */
function packageRelsXml(have: Set<string>): string {
  const core = have.has("docProps/core.xml") ? `<Relationship Id="rId2" Type="${NS_PKG}/metadata/core-properties" Target="docProps/core.xml"/>` : "";
  return (
    `${DECL}<Relationships xmlns="${NS_PKG}">` +
    `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/>` +
    `${core}</Relationships>`
  );
}

/** The package relationships can be kept as read when they still point at the part this writer emits. */
function packageRelsAreUsable(bytes: Uint8Array | undefined): boolean {
  return bytes !== undefined && /relationships\/officeDocument"[^>]*Target="word\/document\.xml"/.test(new TextDecoder().decode(bytes));
}

export function writeDocx(doc: WordDocument): Uint8Array {
  const src = doc.source();
  // A document nobody edited saves back byte-for-byte: regeneration covers the
  // paragraphs this model reads and would drop fields, bookmarks and content
  // controls it does not.
  if (src && !doc.edited) {
    const untouched: Record<string, Uint8Array> = {};
    for (const [path, bytes] of src.parts) untouched[path] = bytes;
    return zipSync(untouched, { level: 6 });
  }

  const rels = new RelBuilder(doc.sourceRels());
  // Anchors in the regenerated body only mean anything if the part they point at is related.
  if (doc.comments.length) rels.comments(doc.commentsRelTarget());
  const document = documentXml(doc, rels);
  const files: Record<string, Uint8Array> = {
    "word/document.xml": strToU8(document),
  };

  const passthrough = doc.untouchedParts();
  for (const part of passthrough) files[part.path] = part.bytes;
  if (!packageRelsAreUsable(files["_rels/.rels"])) files["_rels/.rels"] = strToU8(packageRelsXml(new Set(Object.keys(files))));

  // Header/footer texts the user changed regenerate here; untouched rich parts
  // (letterhead logos, tables) keep their exact bytes via the passthrough above.
  for (const hf of doc.headerFooterParts()) {
    if (hf.dirty) files[hf.part] = strToU8(hf.xml);
  }

  // The same rule for comments: an unchanged part passes through above, and only a list
  // that actually changed is rebuilt — paragraph-id keys of deleted comments go with it.
  if (doc.commentsDirty()) {
    doc.prepareCommentsForSave();
    const part = doc.commentsPart();
    files[part.path] = strToU8(commentsXml(doc.comments, part.rootAttrs));
    for (const ext of doc.commentExtensionParts()) {
      files[ext.path] = strToU8(ext.xml);
      // A part this package never carried needs both halves of the promise: a relationship
      // from the body and a content-type override. Without either, Word ignores it and the
      // resolve state the user just set silently disappears on the next open.
      rels.relation(ext.relType, ext.target);
    }
  }

  files["word/_rels/document.xml.rels"] = strToU8(relsXml(rels));

  // A picture the model holds is written from the model's own bytes and wins over the copy that travelled
  // with the package — the model is the thing that was edited. Media the model holds no inline for came
  // through the passthrough above and keeps its bytes, which is what the body's `r:embed` still cites.
  const fromModel = new Set<string>();
  for (const inline of everyInline(doc)) {
    const img = inline.image;
    if (!img || fromModel.has(img.path)) continue;
    fromModel.add(img.path);
    files[img.path] = bytesOfDataUrl(img.src);
  }

  const paths = Object.keys(files).map((path) => ({ path, type: passthrough.find((p) => p.path === path)?.type }));
  const partsAtRead = new Set(src ? [...src.parts.keys()] : []);
  files["[Content_Types].xml"] = strToU8(contentTypesXml(src?.text("[Content_Types].xml") ?? "", partsAtRead, paths));
  return zipSync(files, { level: 6 });
}

function everyInline(doc: WordDocument): Inline[] {
  const out: Inline[] = [];
  const walkBlocks = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.kind === "paragraph") out.push(...block.inlines);
      else if (block.kind === "table") for (const row of block.rows) for (const cell of row) walkBlocks(cell.paragraphs);
    }
  };
  walkBlocks(doc.blocks);
  return out;
}
