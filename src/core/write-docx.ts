import { strToU8, zipSync } from "fflate";
import type { Block, BorderSet, CellStyle, Inline, ListAddition, MarginTwips, ParaStyle, Paragraph, ParagraphFormatRevision, PropSlot, RunStyle, RowInfo, RowStyle, Table, TableBorder, TableCell, TableStyle, WordDocument } from "./docx.js";
import { bytesOfDataUrl } from "./media.js";
import { BORDER_ORDER, PBDR_ORDER, ROW_REVISION_TAGS, ownParaStyle, ownRunStyle, paraBorderChanges, paraStyleOf, preserve, revisionTag, runStyleOf, samePropValue, shadingFillToken } from "./docx.js";
import { childSpans, elementContent } from "./opc.js";
import { changeRecordXml, placeRecordXml, setChildren } from "./doc-table-revisions.js";
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
/** Where list definitions live: this part is the only thing a list command on a definition-less file owes. */
const NUMBERING_PART = "word/numbering.xml";
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
 * order is the one WPS writes: of the 2,400 `w:ind` elements the authorized 46 份 spell across `word/*`, 2,306
 * keep this relative order (`left`, then `right`, then the first line's two spellings, then `w:hanging`), and the
 * only multi-attribute shape that does not is `left,leftChars,hanging,firstLineChars` — 94 of them in 4 份, where
 * WPS puts `hanging` before the first line (`.scratch/authorized-slot-census.py`; the 29,889 over the 86 documents
 * is the quarantined cohort's arithmetic). Attribute order asks nothing of a reader. The two paths that spell a
 * `w:ind` agree on that order: this one, for a paragraph the file gave no `w:pPr` and for a `w:pPrChange` this
 * editor signs, and `IND_ORDER` in `syncParaSlots`, which reorders only the six modelled attributes and leaves
 * the rest of a slot's bytes alone. So editing the indent of one of those 4 份 does move `w:hanging` behind
 * `w:firstLineChars`, and the 94 elements WPS wrote the other way are the whole of what that costs.
 */
function indTag(s: ParaStyle): string {
  const attrs: string[] = [];
  const put = (name: string, value: number | undefined): void => {
    if (value !== undefined) attrs.push(` ${name}="${Math.round(value)}"`);
  };
  put("w:left", s.indentLeft);
  put("w:leftChars", s.indentLeftChars);
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
  // Nothing was carried, so the fields spell the set — only the ones the paragraph said for itself, which
  // `ownParaStyle` decides against the read snapshot. A paragraph whose file wrote no `w:pPr` states nothing of
  // its own, yet every property it now shows comes down the style chain (`paraLent`): the style goes on lending
  // it, so spelling it here would write a `w:pPr` the file never had. 159 of the corpus's paragraphs are in that
  // shape, 91 of them in files whose default paragraph style states a property — 待重证, because
  // `.scratch/para-no-pr.ts` counts over a `/tmp` list of paths that was never pinned to a cohort, and the only
  // list it can have been fed is the quarantined one. A
  // paragraph this editor built has no snapshot, so all of its fields are its own and get spelled, which is
  // what keeps a paragraph split in two looking like the one it came from.
  const own = ownParaStyle(p.asRead, paraStyleOf(p));
  const has = (key: keyof ParaStyle): boolean => own[key] !== undefined;
  const parts: string[] = [];
  if (has("styleId")) parts.push(`<w:pStyle w:val="${XML_ESC(p.styleId ?? "")}"/>`);
  // 「换行和分页」那四枚（§17.3.1.14 `w:keepNext`、§17.3.1.13 `w:keepLines`、§17.3.1.65 `w:pageBreakBefore`、
  // §17.3.2.90 `w:widowControl`）在没有 `w:pPr` 的段上只有字段可说——编辑器拆出来的段、以及件根本没写 `w:pPr` 的那几段。
  // 写不写看**状态变没变**而不是「字段真不真」：`pageBreakBefore` 的字段缺省就是 `false`，照 `ownParaStyle` 那套
  // 「非假才写」的说法，取消勾选（真→假）就一句话都没有，纸上照旧断页；反过来照「与读快照不同就写」，新建出来的段
  // （`asRead` 只有样式链那一份）每一段都会多出一句 `<w:pageBreakBefore w:val="0"/>`——替文件编一句它没说过的话。
  // 拼法照件：开 = 裸写，关 = `w:val="0"`（授权 46 份里这四枚说关的一律这么写，`.scratch/toggle-lent-chain.ts`）。
  // `CT_PPr` 把这四枚排在 `w:pStyle` 之后、`w:numPr` 之前。
  for (const tag of ["keepNext", "keepLines", "pageBreakBefore", "widowControl"] as const) {
    const was = p.asRead?.[tag] === true;
    const now = p[tag] === true;
    if (was === now) continue;
    parts.push(now ? `<w:${tag}/>` : `<w:${tag} w:val="0"/>`);
  }
  if (has("numPr")) parts.push(`<w:numPr><w:ilvl w:val="${p.numPr!.ilvl}"/><w:numId w:val="${p.numPr!.numId}"/></w:numPr>`);
  // 框与底纹两条路共用同一个口径 `paraBorderChanges`：**只有这一击动过的边**说出去。带 `w:pPr` 的段落由
  // `syncParaSlots` 逐边认领（那条路保住件原来的字节与模型没有字段的那些属性），只剩字段的这一段才由这里说——
  // 编辑器拆出来的段、以及件根本没写 `w:pPr` 的那 159 段就是这种。拼法照件：一边一枚 `<w:side>`、按 `CT_PBdr`
  // 的边序（26 枚 `w:pBdr` 全是这个形状），底纹恒为 `w:val="clear" w:color="auto" w:fill="十六进制"`
  // （30/30，`.scratch/pbdr-audience.ts`）；那一格填的不是颜色就按 token 原样写回（`auto`：段落层 0 例、
  // `w:tcPr` 层 1,600 例，见 `shadingFillToken`）。件里的框 24 枚在 `word/styles.xml`、正文段落自己写的
  // 0 枚（同一条探测），所以把样式链借来的整个盒子抄到段落身上，每一枚都是替文件新编的一句样式的话——
  // 逐边比而不是整枚比，才知道这一段到底改了哪一边。
  const box = paraBorderChanges(p);
  if (box) {
    const sides = PBDR_ORDER.filter((side) => box[side]);
    parts.push(`<w:pBdr>${sides.map((side) => `<w:${side} ${borderAttrs(box[side]!)}/>`).join("")}</w:pBdr>`);
  }
  if (p.shading !== undefined && !samePropValue(p.asRead?.shading, p.shading)) {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${XML_ESC(shadingFillToken(p.shading))}"/>`);
  }
  // 「对齐到网格」不在 `PARA_STYLE_KEYS` 里——带 `w:pPr` 的段落由 `syncParaSlots` 按件的原字节带着走，这一段
  // 却可能一个 slot 也没有（编辑器自己拆出来的段、或那 159 段没写 `w:pPr` 的），那时候勾选框只剩字段可说。
  // 拼法照件：说开裸写、说关 `w:val="0"`。只说**这一击动过的**，与 `ownParaStyle` 对着读快照同一条口径——
  // 从样式链借来的那个开关是样式的话，抄到段落身上就是替文件编一句。`CT_PPr` 把它排在 `w:spacing` 前。
  // 「允许西文在单词中间换行」（`w:wordWrap`）与上面那一支同一条来路：不在 `PARA_STYLE_KEYS` 里，带 `w:pPr` 的
  // 段落由 `syncParaSlots` 按件的原字节带走，这里只服务没有 slot 的段。拼法照件——写了的 200 枚枚枚裸写（授权
  // 口径 `.scratch/cjk-typo-audience.py`），说关那一句按 §17.4.77 的 on/off 写法补 `w:val="0"`。`CT_PPr` 把它排在
  // `w:snapToGrid` 前（`w:kinsoku` 之后）。
  if (p.wordWrap !== undefined && !samePropValue(p.asRead?.wordWrap, p.wordWrap)) {
    parts.push(p.wordWrap ? "<w:wordWrap/>" : '<w:wordWrap w:val="0"/>');
  }
  if (p.snapToGrid !== undefined && !samePropValue(p.asRead?.snapToGrid, p.snapToGrid)) {
    parts.push(p.snapToGrid ? "<w:snapToGrid/>" : '<w:snapToGrid w:val="0"/>');
  }
  const spacing: string[] = [];
  if (has("spacingBefore")) spacing.push(` w:before="${Math.round(p.spacingBefore!)}"`);
  if (has("spacingAfter")) spacing.push(` w:after="${Math.round(p.spacingAfter!)}"`);
  // A height the model holds without a rule of its own is signed `w:auto`: that is §17.3.1.27's default, and
  // signing it is what the files do anyway — all 3,948 stated heights in the authorized 46 份's body/band/footnote
  // parts name their rule beside it (2,890 `auto`, 1,004 `exact`, 54 `atLeast`; the style layer adds 212 more and
  // signs them all too). The old cohort's 13,396 `w:auto` is quarantined arithmetic.
  // `.scratch/authorized-spacing-census.py`.
  if (has("lineSpacing")) spacing.push(` w:line="${Math.round(p.lineSpacing!)}" w:lineRule="${p.lineSpacingRule ?? "auto"}"`);
  if (spacing.length) parts.push(`<w:spacing${spacing.join("")}/>`);
  const ind = indTag({
    ...(has("indentLeft") ? { indentLeft: p.indentLeft } : {}),
    ...(has("indentLeftChars") ? { indentLeftChars: p.indentLeftChars } : {}),
    ...(has("indentRight") ? { indentRight: p.indentRight } : {}),
    ...(has("indentFirstLine") ? { indentFirstLine: p.indentFirstLine } : {}),
    ...(has("indentFirstLineChars") ? { indentFirstLineChars: p.indentFirstLineChars } : {}),
    ...(has("indentHanging") ? { indentHanging: p.indentHanging } : {}),
  });
  if (ind) parts.push(ind);
  if (has("align")) parts.push(`<w:jc w:val="${XML_ESC(p.align ?? "")}"/>`);
  // `CT_PPr` ranks `w:outlineLvl` below `w:jc` and above `w:rPr`. A paragraph with no slots — one this editor
  // built by splitting another — has only the fields to say what it is, so a heading that does not get its level
  // spelled here comes back off the page as body text.
  if (has("outlineLevel")) parts.push(`<w:outlineLvl w:val="${p.outlineLevel!}"/>`);
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
  if (s.lineSpacing) spacing.push(` w:line="${Math.round(s.lineSpacing)}" w:lineRule="${s.lineSpacingRule ?? "auto"}"`);
  if (spacing.length) parts.push(`<w:spacing${spacing.join("")}/>`);
  const ind = indTag(s);
  if (ind) parts.push(ind);
  if (s.align) parts.push(`<w:jc w:val="${XML_ESC(s.align)}"/>`);
  if (s.outlineLevel !== undefined) parts.push(`<w:outlineLvl w:val="${s.outlineLevel}"/>`);
  return parts.join("");
}

/** `w:rFonts` from the faces the model holds. A face it does not hold is not spelled. */
function rFontsTag(s: RunStyle, prefix = "w"): string {
  const latin = s.font ? ` ${prefix}:ascii="${XML_ESC(s.font)}" ${prefix}:hAnsi="${XML_ESC(s.font)}"` : "";
  const east = s.fontEastAsia ? ` ${prefix}:eastAsia="${XML_ESC(s.fontEastAsia)}"` : "";
  return latin || east ? `<${prefix}:rFonts${latin}${east}/>` : "";
}

/**
 * The faces to spell into a `w:rFonts` for a run whose file wrote none. `w:rFonts` is the one element in this set
 * that carries two fields, so the guard every other property has — a value the run borrowed from its style chain
 * or from `w:docDefaults` is not a value the run states — needs saying again here: a run that only had its Latin
 * slot moved must not have the Chinese face it was *reading through the floor* written into it, or one command
 * would pin a sentence the file never wrote and take that run out from under the floor it belongs to
 * (`.scratch/font-slot-shape.ts`: 5,444 Han-text runs and 5,483 Latin-text ones state no `w:rFonts` of their own).
 */
function movedFaces(r: Inline): RunStyle {
  return {
    font: samePropValue(r.asRead?.font, r.font) ? undefined : r.font,
    fontEastAsia: samePropValue(r.asRead?.fontEastAsia, r.fontEastAsia) ? undefined : r.fontEastAsia,
  };
}

/**
 * `CT_RPr` children spelled from fields alone: the two places that have no bytes to carry — a run the document
 * wrote no `w:rPr` for, and a `w:rPrChange` whose record arrived without the file's children.
 *
 * Derived from `RUN_PROPS`, not written out again beside it. A hand-list of the same set is a second owner of the
 * same sentence, and the two drift: `charSpacingTwips` was registered in the table and left out here, so a run
 * whose record had to be rebuilt from fields lost its 「字距 82」and gained a `w:szCs` the file never wrote
 * (`.scratch/char-spacing-dom.ts` 的「改过字距的字」打字后一段两句字距变成一句).
 */
function runPropsInner(s: RunStyle): string {
  const r = s as unknown as Inline;
  return RUN_PROPS.map((prop) => ({ tag: prop.tags[0]!, xml: prop.spell(r, "w", "") }))
    .filter((p) => p.xml !== "")
    .sort((a, b) => rprRank(a.tag) - rprRank(b.tag))
    .map((p) => p.xml)
    .join("");
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
 * Latin one rather than leaving the run at two sizes. `w:rFonts` is the reverse — one element, two fields — so it
 * lists both faces under `keys` and re-spells whichever moved. Everything not in this list has no field and keeps
 * the text the file wrote for it — `w:rStyle`, `w:kern`, `w:lang`, `w:bCs`, `w:shd`, `w:eastAsianLayout`.
 * `w:spacing` (字符间距) left that set in round #174: 授权 46 件里 32 枚、枚枚非缺省（18 枚字符样式写 `-6`、14 枚 run
 * 写 `82`），所以它有字段、有 CSS、也有下面 `RUN_PROPS` 里的一条写法。
 */
const RUN_PROPS: {
  key: keyof RunStyle;
  /** The fields that may move this element. One entry's several faces, or one field's several tags. */
  keys?: (keyof RunStyle)[];
  tags: string[];
  spell: (r: Inline, prefix: string, source: string) => string;
}[] = [
  {
    key: "font",
    keys: ["font", "fontEastAsia"],
    tags: ["rFonts"],
    spell: (r, prefix, source) =>
      !r.font && !r.fontEastAsia
        ? ""
        : source
          ? putRunFonts(source, prefix, r)
          : rFontsTag(movedFaces(r), prefix),
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
  {
    key: "strike",
    tags: ["strike"],
    // Spelled, not deleted, on the off side: 2 runs in 2 of these files state `<w:strike w:val="0"/>` and carry
    // text the reader resolves (`.scratch/onoff-where.ts`, `.scratch/onoff-audience.ts`), and dropping the element
    // would replace the file's sentence with a silence no file wrote — the same rule `<w:b w:val="0"/>` and
    // `<w:u w:val="none"/>` follow.
    spell: (r, prefix) =>
      r.strike === undefined ? "" : r.strike ? `<${prefix}:strike/>` : `<${prefix}:strike w:val="0"/>`,
  },
  {
    key: "color",
    // `w14:textFill` is the same sentence in the 2010 extension's words, so it moves with the colour. On the
    // authorized 46 份 the extension is spelled 58 times, in 6 份, every one of them in `word/styles.xml` (those 6
    // are three preset templates in their `.docx` and `.dotx` forms) and none in the body or a band — so the cohort
    // proves the shape, not a body-run audience. All 58 sit directly under a `w:rPr` that states a `w:color`, and
    // that `w:color` states a hex value and a `w:themeColor` in all 58; the extension names the same theme slot
    // (`tx1` for `text1`) in all 58. Two spellings cover them: `w14:solidFill/w14:schemeClr tx1` with `lumMod`
    // 85000/`lumOff` 15000 beside `w:color w:val="262626" w:themeColor="text1" w:themeTint="D9"` (56), and
    // `accent1` beside `4874CB` (2) (`.scratch/authorized-slot-census.py`). Whether resolving the extension through
    // `word/theme/theme1.xml` lands on the hex `w:color` already gave is the part this cohort has not been re-run
    // on: the 3296-that-resolve / 360-that-only-differ-in-naming-generation split and the 16-file save observation
    // are the quarantined cohort's arithmetic, and 2,110-as-slot / 1,990-on-a-run-with-text with them.
    // It is a mirror, not an extra, and that is why the write path still drops it with the colour: left standing,
    // a colour change writes a run saying two colours at once, and which one a reader paints with is not ours to
    // choose. Nothing is spelled for a run that never stated one, so a document without the extension does not
    // gain a `w14:` child it never declared.
    tags: ["color", "textFill"],
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
  {
    key: "underline",
    tags: ["u"],
    // The cancellation is written, not deleted: 授权 46 份的 `word/document.xml` 里 `<w:u w:val="none"/>` 有 30 枚
    // （同一趟逐枚：`single` 1,565、`thick` 20、裸写 0），所以「取消」是文件自己说过的话，不是一个可以顺手省略的
    // 默认值。作废语料上那句「952 枚、254 枚坐在链上什么都不说的 run 上」属旧读数（待重证），换成上面这组。
    // A save that drops the element invents a silence the file never wrote — exactly the mistake
    // `<w:b w:val="0"/>` above refuses to make.
    // 线形跟着说话：那 20 枚 `thick` 坐在 2 份件里，模型存的就是形状名，所以拼回去的还是这个名字，而不是把它
    // 收成 `single` 再交出去。一名的降级会把「粗线」在下一位读者眼里变成细线。
    spell: (r, prefix) => (r.underline === undefined ? "" : `<${prefix}:u w:val="${XML_ESC(r.underline)}"/>`),
  },
  { key: "vertAlign", tags: ["vertAlign"], spell: (r, prefix) => (r.vertAlign ? `<${prefix}:vertAlign w:val="${r.vertAlign}"/>` : "") },
  {
    key: "charSpacingTwips",
    tags: ["spacing"],
    // `0` is a value the run wrote (`ST_SignedTwipsMeasure`), so the field is tested for absence, not truthiness:
    // a run that said 「不加不减」 keeps saying it, and a run that never stated spacing gains no element.
    spell: (r, prefix) =>
      r.charSpacingTwips === undefined ? "" : `<${prefix}:spacing w:val="${Math.round(r.charSpacingTwips)}"/>`,
  },
];

/**
 * `w:rFonts` with the faces the model moved put onto the element as it was written. Only the attributes whose
 * field actually changed are replaced: `w:cs`, `w:cstheme` and `w:hint` have no field at all and always travel as
 * the file spelled them, and a Latin face that has not moved keeps its own bytes even when the east-asian one
 * does. A moved field that now holds nothing takes its attribute with it — that is what refusing a font change
 * means.
 *
 * A moved slot takes its theme reference with it. 授权 46 份全部件共 **14,646 枚 `w:rFonts`**（document 14,007 /
 * styles 397 / 页脚 146 / numbering 84 / 页眉 8 / glossary 4；主题槽那一条腿先前数出的 14,642 与这个差 4，差在
 * glossary 部件没进它那趟扫描——分母按部件对过，两个数各说各的口径），其中 1,957 枚带主题槽、11,905 枚带字面脸、
 * 650 枚在**不同槽**上两种都写，而 **0 枚在同一槽里既写字脸又写主题引用**——所以这条规则按槽成立，不按元素成立
 * （`.scratch/theme-slot-pair-authorized.mjs`）。作废语料上那组 55,585 / 54,200 / 796 / 156 / 433 与 5,260 属旧读数。
 * Writing `w:ascii` next to a standing `w:asciiTheme` would put the run
 * in a state §17.3.2 gives no precedence for and the files never reach; leaving the reference standing is what the
 * other half of the pair is for, since a slot the model did not move keeps saying where its face comes from.
 */
function putRunFonts(xml: string, prefix: string, r: Inline): string {
  const gt = xml.indexOf(">");
  if (gt < 0) return xml;
  const head = xml.slice(0, gt);
  const selfClosed = head.endsWith("/");
  const latinMoved = !samePropValue(r.asRead?.font, r.font);
  const eastMoved = !samePropValue(r.asRead?.fontEastAsia, r.fontEastAsia);
  const strip = (s: string, names: string): string => s.replace(new RegExp(`\\s[\\w.-]+:(?:${names})="[^"]*"`, "g"), "");
  let open = (selfClosed ? head.slice(0, -1) : head).trimEnd();
  if (latinMoved) open = strip(open, "ascii|hAnsi|asciiTheme|hAnsiTheme");
  if (eastMoved) open = strip(open, "eastAsia|eastAsiaTheme");
  const face = (name: string, val: string | undefined): string => (val ? ` ${prefix}:${name}="${XML_ESC(val)}"` : "");
  const latin = latinMoved ? face("ascii", r.font) + face("hAnsi", r.font) : "";
  const east = eastMoved ? face("eastAsia", r.fontEastAsia) : "";
  return `${open}${latin}${east}${selfClosed ? "/>" : ">"}` + (selfClosed ? "" : xml.slice(gt + 1));
}

/**
 * The prefix a carried `w:rPr` writes under. A prefix is taken from the set rather than assumed: `w` is only an
 * alias for the words namespace where the part binds it, and a child written with another alias than its
 * neighbours reads as another element. So the set's alias comes from a child that can live in no other
 * namespace — one of the names `CT_RPr` gives — and a set holding nothing but extension children falls back to
 * the `w` the rest of this writer already spells. Taking whichever child happened to come first instead wrapped a
 * run whose set held only a `w14:textFill` in `<w14:rPr>`, an element no reader knows.
 */
function carriedRunPrefix(slots: PropSlot[]): string {
  const xml = slots.find((s) => RPR_ORDER.includes(s.tag))?.xml ?? "";
  return /^<([A-Za-z_][-\w.]*):/.exec(xml)?.[1] ?? "w";
}

/**
 * The run's `w:rPr` the way the file wrote it, with only the children whose field actually moved re-spelled.
 * The same split `paragraphProps` makes for a `w:pPr` and `applyCellProps` for a `w:tcPr`: the set belongs to
 * the document, and a property the model cannot explain is not the model's to remove. Rebuilding it from nine
 * fields is what took the kerning, the complex-script face and the language out of every run in a document.
 */
function carriedRunProps(r: Inline): string {
  const slots = r.rPrSlots ?? [];
  const prefix = carriedRunPrefix(slots);
  const drops = new Set<PropSlot>();
  const inserts: PropSlot[] = [];
  for (const prop of RUN_PROPS) {
    // A field that has not moved off its read snapshot keeps the source text, so an attribute the model has no
    // field for — `w:kern`, `w:themeColor`, `w:cs` — survives a save that never touched it.
    if ((prop.keys ?? [prop.key]).every((k) => samePropValue(r.asRead?.[k], r[k]))) continue;
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
  // Only what the run itself states goes into its `w:rPr`. A property that came down the style ladder is the
  // style's, and `paragraphProps` keeps the same rule for a `w:pPr`: spelling a borrowed bold into every run
  // of a heading would put the heading's formatting on runs whose file never carried it, and take it back the
  // moment the paragraph left the style.
  const inner = runPropsInner(ownRunStyle(inline.asRead, runStyleOf(inline))) + (fmt ? `<w:rPrChange ${revAttrs(fmt)}>${runProps(fmt.prev) || "<w:rPr/>"}</w:rPrChange>` : "");
  return inner ? `<w:rPr>${inner}</w:rPr>` : "";
}

function textRun(inline: Inline): string {
  if (inline.image) return imageRun(inline, inline.image.width, inline.image.height, "");
  return `<w:r>${revisedRunProps(inline)}${runChild(inline)}</w:r>`;
}

/** What one inline puts inside its `w:r`: a tab, the break as the file wrote it, revised text, or text. */
const runChild = (inline: Inline): string =>
  inline.tab ? "<w:tab/>" : inline.break ? breakXml(inline) : inline.rev?.kind === "del" ? delTextTag(inline.text) : textTag(inline.text);

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
  const prefix = carriedRunPrefix(slots);
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
      out.push(`<w:${revisionTag(rev)} ${revAttrs(rev)}>${runsXml(group, (r) => runWrapped(r, rels), (run, r) => underLink(r, rels, run))}</w:${revisionTag(rev)}>`);
      continue;
    }
    const link = inlines[i].hyperlink;
    if (!link) {
      // Plain neighbours go through one emitter so a source run the reader split can be joined back here.
      const group: Inline[] = [];
      while (i < inlines.length && !inlines[i].carry && !inlines[i].anchor && !inlines[i].rev && !inlines[i].hyperlink)
        group.push(inlines[i++]);
      out.push(runsXml(group, (r) => runWrapped(r, rels), (run) => run));
      continue;
    }
    // Runs join one `w:hyperlink` only while the element itself would be spelled the same: two links to the
    // same target carrying different attributes are two elements, or one run's attributes would be written
    // over its neighbour's.
    const extras = inlines[i].linkAttrs;
    const group: Inline[] = [];
    while (i < inlines.length && inlines[i].hyperlink === link && inlines[i].linkAttrs === extras) group.push(inlines[i++]);
    const inner = runsXml(group, (r) => runXml(r, rels), (run) => run);
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

/**
 * The reader gives every child of a `w:r` its own inline — a tab standing next to text is two of them — so a
 * save that emitted one `w:r` per inline would hand back two runs where the file had one, each carrying its
 * own copy of that run's `w:rPr` and with it every slot the model has no field for (`w:cs`, `w:hint`).
 * Neighbours that came from the same source run and whose properties spell identically go back as the one run
 * the file wrote; a run whose other half really did get its own formatting is still two.
 *
 * `single` spells a run that cannot join: inside a `w:hyperlink` or a revision wrapper the caller has already
 * opened the element, so the per-run emitter it hands in is the bare run — passing the wrapper instead would
 * nest a second `w:hyperlink` around text the file put under one. `joined` places a run that DID join, and is
 * the caller's because only the caller knows whether the link wrapper is already open around it.
 */
function runsXml(runs: Inline[], single: (r: Inline) => string, joined: (run: string, first: Inline) => string): string {
  const out: string[] = [];
  for (let i = 0; i < runs.length; ) {
    const first = runs[i];
    const props = joinableRun(first) ? revisedRunProps(first) : "";
    let j = i + 1;
    if (props)
      while (
        j < runs.length &&
        joinableRun(runs[j]) &&
        runs[j].rPrSlots === first.rPrSlots &&
        runs[j].hyperlink === first.hyperlink &&
        runs[j].linkAttrs === first.linkAttrs &&
        revisedRunProps(runs[j]) === props
      )
        j++;
    if (j > i + 1) {
      out.push(joined(`<w:r>${props}${runs.slice(i, j).map(runChild).join("")}</w:r>`, first));
      i = j;
      continue;
    }
    out.push(single(first));
    i += 1;
  }
  return out.join("");
}

/** The link a run of these inlines belongs under, spelled once for whoever closes them into a run. */
const underLink = (r: Inline, rels: RelBuilder, run: string): string =>
  r.hyperlink ? `<w:hyperlink ${linkAttrs(r.hyperlink, rels, r.linkAttrs)}>${run}</w:hyperlink>` : run;

/** Only a carried run with no anchor, picture or carried bytes rejoins — each of those needs its own element. */
const joinableRun = (r: Inline): boolean => Boolean(r.rPrSlots) && !r.image && !r.anchor && !r.carry;

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
  const props = cellPropsInner(cell);
  // 没有 `w:tcPr` 的格子（编辑器自己造的）没有原件文字可挂记录，这一击签下的记录就拼在这里。
  // 「在不在」那一族（`w:cellIns`/`w:cellDel`）只署名、不带改前的属性集，序上在 `w:tcPrChange` 之前。
  for (const tag of ["cellIns", "cellDel"] as const) {
    const place = cell.revs?.find((r) => r.tag === tag);
    if (place) props.push(placeRecordXml(tag, place));
  }
  const rec = cell.revs?.find((r) => r.tag === "tcPrChange");
  if (rec) props.push(changeRecordXml("tcPrChange", rec, slotsXml(rec.prevSlots) || cellPropsInner(rec.prev).join("")));
  return props.length ? `<w:tcPr>${props.join("")}</w:tcPr>` : "";
}

/** 一格说过的属性，按 `CT_TcPrBase` 的孩子序——现下的格与记录里那一套「改之前」共用这一份拼法。 */
function cellPropsInner(cell: CellStyle): string[] {
  const props: string[] = [];
  if (cell.widthTwips) props.push(`<w:tcW w:w="${Math.round(cell.widthTwips)}" w:type="dxa"/>`);
  if (cell.spanCols && cell.spanCols > 1) props.push(`<w:gridSpan w:val="${cell.spanCols}"/>`);
  if (cell.spanRows !== undefined) props.push(cell.spanRows === 1 ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>");
  const borders = bordersXml("tcBorders", cell.borders);
  if (borders) props.push(borders);
  if (cell.shading) props.push(`<w:shd w:val="clear" w:color="auto" w:fill="${XML_ESC(shadingFillToken(cell.shading))}"/>`);
  if (cell.noWrap) props.push("<w:noWrap/>");
  if (cell.textDirection) props.push(`<w:textDirection w:val="${XML_ESC(cell.textDirection)}"/>`);
  if (cell.valign) props.push(`<w:vAlign w:val="${XML_ESC(cell.valign)}"/>`);
  return props;
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
  if (!kids.length) kids.push(...rowPropsInner(row));
  // 没有原件孩子的行没有地方挂签下的记录，只能在这里拼——与表、格那两条同一句「改了、存了、记录没了」。
  // 件里已经带了这一枚孩子的行不双写：`signRowRevision` 把记录换进了 `slots`，这一条只管编辑器自己造的行。
  const rec = row.revFormat;
  if (rec && !kids.some((k) => /<w:trPrChange[\s/>]/.test(k)))
    kids.push(changeRecordXml("trPrChange", rec, slotsXml(rec.prevSlots) || rowPropsInner(rec.prev).join("")));
  if (row.rev) {
    const at = kids.findIndex((k) => /<w:trPrChange[\s/>]/.test(k));
    kids.splice(at < 0 ? kids.length : at, 0, `<w:${revisionTag(row.rev)} ${revAttrs(row.rev)}/>`);
  }
  return kids.length ? `<w:trPr>${kids.join("")}</w:trPr>` : "";
}

/**
 * 一行说过的属性，按 `CT_TrPr` 的孩子序（`.scratch/trpr-order-audience.py` 量到的那一段：`w:cantSplit` 在前、
 * `w:trHeight` 居中、`w:jc` 收尾）——现下的行与记录里那一套「改之前」共用这一份拼法。
 * `w:jc` 也在这一份里：读侧认它（`rowStyleOf`），画侧也画它（`render.ts:rowStyle`），按字段拼时漏了它，
 * 编辑器自己造的行就把这一行说过的摆位丢了。
 */
function rowPropsInner(row: RowStyle): string[] {
  const props: string[] = [];
  if (row.keepTogether) props.push("<w:cantSplit/>");
  if (row.heightTwips !== undefined)
    props.push(`<w:trHeight w:val="${Math.round(row.heightTwips)}"${row.heightRule ? ` w:hRule="${XML_ESC(row.heightRule)}"` : ""}/>`);
  if (row.header) props.push("<w:tblHeader/>");
  if (row.jc) props.push(`<w:jc w:val="${XML_ESC(row.jc)}"/>`);
  if (row.hidden) props.push("<w:hidden/>");
  return props;
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
  const props = tablePropsInner(t);
  // 与格子同一条：编辑器自己造的表没有原件文字，签名留下的记录只能在这里拼出来——不拼就是「改了、存了、
  // 记录没了」，正是那条静默交回原字节的坑。
  const rec = t.revFormat;
  if (rec) props.push(changeRecordXml("tblPrChange", rec, slotsXml(rec.prevSlots) || tablePropsInner(rec.prev).join("")));
  const inner = props.join("");
  return inner ? `<w:tblPr>${inner}</w:tblPr>` : "";
}

/** 表说过的属性，按 `CT_TblPr` 的孩子序——现下的表与记录里那一套「改之前」共用这一份拼法。 */
function tablePropsInner(t: TableStyle): string[] {
  return [
    t.styleId ? `<w:tblStyle w:val="${XML_ESC(t.styleId)}"/>` : "",
    // `w:w="0" w:type="auto"` is what "no width declared" already means, so a table that never said one
    // gains nothing from being given the tag and a save that writes it is inventing a property.
    t.widthTwips !== undefined ? `<w:tblW w:w="${Math.round(t.widthTwips)}" w:type="${XML_ESC(t.widthType ?? "dxa")}"/>` : "",
    t.align ? `<w:jc w:val="${XML_ESC(t.align)}"/>` : "",
    t.indentTwips !== undefined ? `<w:tblInd w:w="${Math.round(t.indentTwips)}" w:type="dxa"/>` : "",
    bordersXml("tblBorders", t.borders),
    // 与 `cellPropsXml` 同一枚形状：`w:val="clear" w:color="auto"` 带头、颜色写在 `w:fill`。
    // 表这一层的底纹在授权语料里 0 枚（`.scratch/shd-audience.ts`），摆的位置是 `CT_TblPr` 给的序。
    t.shading ? `<w:shd w:val="clear" w:color="auto" w:fill="${XML_ESC(shadingFillToken(t.shading))}"/>` : "",
    t.layout ? `<w:tblLayout w:type="${XML_ESC(t.layout)}"/>` : "",
    cellMarginXml(t.cellMargin),
  ].filter(Boolean);
}

/**
 * 编辑器自己造的表那一份默认单元格边距。与 `w:tblCellMar` 在件里的拼法同一句：每条边都是 `w:w` 加
 * `w:type="dxa"`（授权 46 份里 1,502 枚边孩子无一例外）。
 */
function cellMarginXml(m: MarginTwips | undefined): string {
  if (!m) return "";
  const sides = (["top", "left", "bottom", "right"] as const)
    .filter((s) => m[s] !== undefined)
    .map((s) => `<w:${s} w:w="${Math.round(m[s]!)}" w:type="dxa"/>`);
  return sides.length ? `<w:tblCellMar>${sides.join("")}</w:tblCellMar>` : "";
}

/** 一枚记录带着的属性原文：件里怎么写的就怎么回去，模型没有字段的属性也在这一份里。 */
const slotsXml = (slots: PropSlot[]): string => slots.map((s) => s.xml).join("");

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
  // `w:orient` is spelled only for landscape, because that is what the spec says the token is for (§17.3.2.12) —
  // not a corpus measurement: over the authorized 46 WPS templates all 58 sections are portrait (56 write
  // 11906×16838, 2 write 11907×16840) and **0 of the 58 state `w:orient`**, so the landscape half of this rule
  // has no authorized witness and its old numbers (93 portrait / 22 landscape sections) are 待重证.
  // `.scratch/authorized-families2.mjs` prints the section and `w:pgSz` counts.
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
  // itself declared them on the root, and on the authorized 46 份 that is exactly one of them: every `w:document`
  // root declares `w, r, wp, wps, wpg, wpi, wpc, w14, w15, w10, wp14, m, o, v, wne` (`w15`/`wpsCustomData` in 43 of
  // them) and not one declares `a` or `pic` (0 of 46 parts, `.scratch/authorized-slot-census.py`). The 16
  // `w:drawing` in that cohort — 12 件, which is 6 template names each shipped twice (zh_CN and zh_HK), and
  // counting them by file name is what used to report 6 — sit 4 in `document.xml` and 12 in `footer1/2/3`, and each
  // bind `a` inside themselves, on `a:graphic`, all 16 of them. A save that "helpfully" lifted that declaration to
  // the root has invented one the file never wrote. Markup this writer generates carries its own declarations (see
  // `imageRun`), so it needs none.
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

/**
 * `w:nsid` / `w:tmpl` 那两枚 id：WPS 自己写的件里是 8 位大写十六进制，看着随机。这里由 `abstractId` 算出来，
 * 所以同一份件反复保存是同一串字节（重写一遍就会让"只改了列表"的一次保存动满篇）。
 */
const hexId = (seed: string, id: number): string => {
  let h = 0x811c9dc5;
  for (const ch of `${seed}:${id}`) {
    h ^= ch.charCodeAt(0);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  h = Math.imul(h ^ (h >>> 15), 0x27d4eb2d) >>> 0;
  return (h >>> 0).toString(16).toUpperCase().padStart(8, "0");
};

/**
 * A list definition the model had to invent, because the file named no such shape (docx.ts
 * `ensureListNumbering`). The part is never rebuilt from the model — the definitions the document came with,
 * its `w:numPicBullet`, its `mc:` alternates and everything else this writer does not read travel as the bytes
 * that came in, and these are *joined onto* them. §17.4's one ordering rule holds either way: every
 * `w:abstractNum` precedes every `w:num`.
 *
 * The spelling follows what WPS writes for its own single-level definitions: every `w:abstractNum` in the
 * bundle's 11 distinct numbering parts opens `w:nsid` / `w:multiLevelType` / `w:tmpl` (23 of 23), every `w:lvl`
 * carries `w:ilvl` then `w:tentative="0"` (55 of 55), and `w:start` is `1` on 52 of those 55. What it does *not*
 * follow is `w:suff` and `w:pPr/w:ind`: among the single-level definitions 13 of 19 state `w:suff` but split
 * 7 `space` / 6 `nothing`, and only 6 of 19 state any `w:pPr` — a layout claim with no majority to copy, so the
 * level keeps the core the corpus really does spell. A part whose root this cannot find comes back untouched:
 * inventing a root around bytes of unknown shape would be the bigger lie, and the paragraph then draws the
 * definition-less list it was before the click.
 */
function numberingXml(carried: string | undefined, adds: ListAddition[]): string {
  // Whatever prefix the part's own root declared is the prefix these elements use; `w:` is what every file in
  // the corpus writes, and what a part created here declares.
  const prefix = /<([\w-]+):numbering\b/.exec(carried ?? "")?.[1] ?? "w";
  const q = (name: string): string => `${prefix}:${name}`;
  const block =
    adds
      .map(
        (a) =>
          `<${q("abstractNum")} ${q("abstractNumId")}="${a.abstractId}">` +
          `<${q("nsid")} ${q("val")}="${hexId("nsid", a.abstractId)}"/>` +
          `<${q("multiLevelType")} ${q("val")}="singleLevel"/>` +
          `<${q("tmpl")} ${q("val")}="${hexId("tmpl", a.abstractId)}"/>` +
          `<${q("lvl")} ${q("ilvl")}="0" ${q("tentative")}="0"><${q("start")} ${q("val")}="1"/><${q("numFmt")} ${q("val")}="${a.numFmt}"/>` +
          `<${q("lvlText")} ${q("val")}="${XML_ESC(a.lvlText)}"/><${q("lvlJc")} ${q("val")}="left"/></${q("lvl")}>` +
          `</${q("abstractNum")}>`,
      )
      .join("") +
    adds.map((a) => `<${q("num")} ${q("numId")}="${a.numId}"><${q("abstractNumId")} ${q("val")}="${a.abstractId}"/></${q("num")}>`).join("");

  if (!carried) return `${DECL}<${q("numbering")} xmlns:${prefix}="${NS_W}">${block}</${q("numbering")}>`;
  const open = new RegExp(`<${prefix}:numbering\\b[^>]*>`).exec(carried);
  if (!open) return carried;
  if (open[0].endsWith("/>"))
    return `${carried.slice(0, open.index)}${open[0].slice(0, -2)}>${block}</${q("numbering")}>${carried.slice(open.index + open[0].length)}`;
  // The only element named exactly `num` in this part is the top-level definition, so the first one is where the
  // abstract definitions have to precede; a part with no `num` yet takes them at the end, before its close.
  const firstNum = carried.search(new RegExp(`<${prefix}:num(?=[\\s>/])`));
  const where = firstNum >= 0 ? firstNum : carried.lastIndexOf(`</${q("numbering")}>`);
  return where < 0 ? carried : `${carried.slice(0, where)}${block}${carried.slice(where)}`;
}

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

  // A list the document never defined owes bytes too: the part is created when the package had none — 授权 46 份里
  // **24 份根本没有 `word/numbering.xml`**（带的那 22 份去重成 13 个模板），这就是 文件 → 新建 的形状——
  // and joined onto the
  // file's own definitions when it did. The body relationship is the other half of the promise; a part no
  // relationship reaches is a part Word ignores, and the mark the user just asked for would not draw.
  if (doc.listDirty()) {
    files[NUMBERING_PART] = strToU8(numberingXml(src?.text(NUMBERING_PART), doc.listAdditions()));
    if (!rels.entries.some((e) => e.type === `${NS_R}/numbering`)) rels.relation(`${NS_R}/numbering`, "numbering.xml");
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
