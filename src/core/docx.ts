import { OpcPackage, asArray, childSpans, children, decodeEntities, elementContent } from "./opc.js";
import { ThemeFonts, themeFontScript } from "./theme-fonts.js";
import { isMark } from "./run-identity.js";
import { dataUrl } from "./media.js";
import { TABLE_RECORD_TAGS, changeRecordXml, coversWholeSet, decideSetXml, decideSlots, isPropertyRecord, placeRecordXml, recordedSlots, setChildren } from "./doc-table-revisions.js";
import { cellCovering, columnStarts, deleteColumn, gridColumns, removeCell, repairMerges, restateColumnFromCells } from "./doc-table-ops.js";
import {
  allocIds,
  anchorInline,
  bodyShape,
  commentExEntries,
  commentIdsRootAttrs,
  commentIdEntries,
  commentParaIds,
  commentSpans,
  commentState,
  commentsExRootAttrs,
  commentsExXml,
  commentIdsXml,
  commentsRootAttrs,
  NEW_COMMENTS_ROOT,
} from "./comments.js";
import type { CommentAnchor, DocComment } from "./comments.js";

/** Relationship types of the two comment extension parts. */
const REL_COMMENTS_EX = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const REL_COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";

export interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  /**
   * `w:u@w:val`（§17.3.2.25 的 `ST_Underline`）按文件自己写的名字带着：**缺席 = 「没提过」，`"none"` = 「说了不要」，
   * 其余名字 = 「说了要，线是这一档」**——和 `highlight` 同一套三态，只是「要」的那一档从前塌成 `true`，形状就此丢了；
   * 现在留名，画的时候才谈得上画粗线、点线。授权 46 份 1,615 枚 `w:u` 只出现三个名字（`single` 1,565 / `none` 30 /
   * `thick` 20，0 枚裸写；`.scratch/underline-audience.py` 逐枚），`thick` 那 20 枚坐在 2 份件里、另有 16 枚在段落标记上。
   * 认得的档按 `DocumentFormat.OpenXml.Wordprocessing.UnderlineValues` 序列化出来的那串名字；比那串更长的名字这一批件
   * 没喂过，照原样带、照原样写回，画不出来的按单线画（不替文件改口，也不假画）。
   */
  underline?: string;
  strike?: boolean;
  sizePt?: number;
  color?: string;
  /**
   * §17.17.64's `ST_HighlightColor` name, held as the three states the element has: absent is "this run said
   * nothing", `"none"` is "this run said no highlight", and any other member is the colour it chose. `none` is a
   * value rather than an absence because the file wrote it as one. 授权 46 份重数之后的结论是：**`w:highlight` 在任何
   * xml 部件里 0 枚**（同一趟扫描在同批件上找到 1,615 枚 `w:u`，所以这个 0 是「真没有」而不是「没找到」）。也就是说
   * 这一族在授权口径上**没有受众**，三态建模靠的是 §17.17.64 的枚举与「往返不得改写」这条要求，不是靠语料喂出来的。
   * 作废语料上那句「`<w:highlight w:val="none"/>` 4,135 枚」（`.scratch/highlight-counted.ts`）不再是可引的读数。
   * §17.7.2 stops a property at
   * the first rung that stated one, so keeping it as "nothing said" would let a colour lower down win.
   */
  highlight?: string;
  /**
   * The Latin face, `w:rFonts@w:ascii`. Word dispatches a run's text by script: ASCII letters take `w:ascii`,
   * East Asian ones `w:eastAsia` (the field below), high-ANSI `w:hAnsi`. §17.3.2.26.
   */
  font?: string;
  /**
   * `w:rFonts@w:eastAsia` — the face this document chose for its Chinese. It is a second face on the same
   * element, not a fallback for the first: on the authorized 46 份's body parts 4,475 `w:rFonts` elements write
   * both `w:ascii` and `w:eastAsia`, and 161 of them name two different literal faces
   * (`.scratch/corpus-recheck2.mjs`). A run whose
   * Chinese draws in the Latin face is a run the file never described that way.
   */
  fontEastAsia?: string;
  vertAlign?: "baseline" | "superscript" | "subscript";
  /**
   * `w:rPr/w:spacing` (§17.3.2's character spacing): twentieths of a point added after **each** character, and it
   * may be negative (WPS's own 字体→字符间距 加宽/紧缩 writes exactly this). It is a horizontal measure, so it moves
   * where a line breaks — a run this app paints at its unstated advance paints a layout the file did not describe.
   * 授权 46 件量过（`.scratch/layout-elements-audience.py`、`.scratch/layout-second-gate.py`，件按全路径计）：**32 枚
   * 住在 8 件里，32 枚全说非缺省值**——18 枚在 `w:style/w:rPr`（枚枚 `w:type="character"`、枚枚 `w:val="-6"`，即紧缩
   * 0.3 磅），14 枚在 `w:p/w:r/w:rPr`（枚枚 `w:val="82"`，即加宽 4.1 磅）；段落标记那一侧（`w:pPr/w:rPr`）0 枚。
   * 元素受众与效果受众两关都过，所以这一族进模型。`undefined` 是「这句没说」，`0` 是「说了不加不减」——它也是
   * 一句写过的话，不能与沉默混成一格。
   */
  charSpacingTwips?: number;
}

export interface Inline extends RunStyle {
  text: string;
  hyperlink?: string;
  /**
   * The attributes on this link's own `w:hyperlink` opening tag that the model has no field for —
   * `w:history`, `w:docProp`, `w:tm`, `w:invalid` — as the part spelled them, in the file's order. Where the
   * link *points* is not here: `r:id` and `w:anchor` are the `hyperlink` field's job, and carrying them would
   * let a save point at a target the model no longer holds. See `Paragraph.attrsXml`.
   */
  linkAttrs?: string;
  image?: { src: string; width: number; height: number; name: string; path: string };
  /**
   * The `w:drawing` this picture came from, exactly as the file wrote it, alongside the fields the page draws
   * from. A shape carries far more than a picture's bytes — the frame it floats in, its wrap, its effects, the
   * crop and the fill's own extension list — and five fields cannot hold any of that, so a save returns these
   * bytes while the size and name it reads them for have not moved. `width`/`height`/`name` are the snapshot
   * that decides it; `embed` is the relationship id the source used, which the writer remaps to its own.
   */
  drawing?: { xml: string; width: number; height: number; name: string; embed?: string };
  break?: boolean;
  /**
   * The `w:br` this break came from, as the file wrote it. A break carries more than its kind — `w:clear`
   * says whether it also clears the floating shapes beside it — and the model edits none of that, so the
   * token travels whole instead of being re-spelled from a field list that only knows `w:type`. Absent on a
   * break this editor inserts, which is the soft return Word writes with no attribute at all.
   */
  brXml?: string;
  tab?: boolean;
  /**
   * A comment anchor sitting between runs: `w:commentRangeStart`, `w:commentRangeEnd`
   * or the run holding `w:commentReference`. Zero-width in the text, but it has to keep
   * its place or the comment loses the range it describes.
   */
  anchor?: CommentAnchor;
  /**
   * A run-level element this model has no fields for — a bookmark pair, a `w:fldSimple` field, an inline
   * content control. The bytes it came in with are kept and written back unchanged, because a document that
   * loses its bookmarks loses the places other parts of it point to.
   */
  carry?: CarriedInline;
  /** The tracked change (`w:ins` / `w:del`) this run sits inside, if it is one. */
  rev?: Revision;
  /** A `w:rPrChange` on this run: formatting someone revised and nobody has decided yet. */
  revFormat?: FormatRevision;
  /**
   * The run's own `w:rPr`, child by child, in the words the file wrote — every child but its `w:rPrChange`,
   * which `revFormat` spells. A run is rebuilt from its fields otherwise, and the field list cannot hold the
   * kerning, complex-script face, language, shading, ligatures or text fill a Chinese document spells.
   */
  rPrSlots?: PropSlot[];
  /**
   * 同一枚 `w:r` 说过的全部文字，只有它装了不止一个内容孩子（`w:tab`+`w:t`、`w:br`+`w:t`）时才写。§17.3 让一枚
   * run 并排装几个孩子，读出来是几枚 inline，共用的却是同一个 `rPrSlots` 引用；「这段文字说没说的是中文」因此
   * 是整枚 run 的话，不是哪个孩子自己的话——按孩子问，挑脸只给带字的那半补 `w:eastAsia`，tab/br 那半不补，
   * 两串性质不同就并不回一枚 run（#142），多出来的那份 `w:cs`/`w:hint` 是这么被复制的。
   */
  runText?: string;
  /**
   * The style fields as they were read. A property that has not moved off this snapshot keeps the source text
   * it came in with; one that has is re-spelled, which is what lets the toolbar work on a run it did not write.
   */
  asRead?: RunStyle;
}

/**
 * An inline structure carried rather than understood. `key` identifies it across a re-render of the page —
 * the same role `data-bid` plays for blocks — and `show` is what a reader would see inside it, for display
 * only and never written back.
 */
export interface CarriedInline {
  key: number;
  /** The element's local name — `bookmarkStart`, `fldSimple`, `sdt`. It says what was kept, not what it means. */
  tag: string;
  xml: string;
  /**
   * A bookmark's `w:name`, read off the parsed element while the read still has it. The one thing in these
   * bytes anything addresses by name — 「定位」的书签目标 lists them. Read here rather than re-parsed from `xml`
   * downstream: that would be a second reader of the same attribute, and the parsed value comes entity-decoded.
   * Display-only; the write side puts `xml` back as it came.
   */
  name?: string;
  show: string;
  /**
   * Which page-number field these bytes hold, for the one thing `show` can't say: a page field displayed on
   * a sheet has to show the number **that** sheet computes, not the words 「页码」. Decided here, where the
   * `w:instrText` is parsed, so nothing downstream re-reads the instruction to learn it. Display only; the
   * write side still puts `xml` back as it came.
   */
  field?: "PAGE" | "NUMPAGES";
}

/**
 * `w:ins` / `w:del`: text someone revised but nobody has accepted or rejected yet. The model has
 * to carry it, because a run inside a revision is invisible to a reader that only knows `w:r` —
 * and dropping the wrapper on save silently accepts the insertion and deletes the deleted text.
 */
export interface Revision {
  kind: "ins" | "del";
  /** `w:id`. Word numbers revisions per document; it only has to be distinct here. */
  id: number;
  /** `w:author`, taken as written — a revision is signed, not editable. */
  author: string;
  date?: string;
  /**
   * Which half of a tracked move this record is, as the file wrote it: `w:moveFrom` signs the characters
   * leaving a place and behaves exactly like a deletion, `w:moveTo` signs them arriving and behaves like
   * an insertion, and the two halves keep their own `w:id` — nothing in OOXML names the other half, so
   * this model does not guess a pairing either. Absent on an ordinary insertion or deletion, and on
   * anything this editor signs itself, because it never offers a move gesture. Without this the tag
   * name is lost on save and the next reader sees an insertion and a deletion where somebody moved text.
   */
  move?: "from" | "to";
}

/**
 * `w:rPrChange`: the formatting a run had before someone revised it. It is not a wrapper the way
 * `w:ins` and `w:del` are — the run keeps its text and shows its *new* formatting — so it rides on
 * the run rather than on `rev`, where a reader looking for an insertion or deletion would misread it.
 */
export interface FormatRevision {
  kind: "format";
  id: number;
  author: string;
  date?: string;
  /** The properties before the change. OOXML records the whole previous set, so a property missing
   *  here is one that was *not* set, not one to leave alone. */
  prev: RunStyle;
  /**
   * That set as the file wrote it, child by child, when the run carried one. `prev` can only hold the ten
   * properties the model has a field for; a save that rewrote the record from those fields took the record's
   * own `w:kern`, `w:lang` and `w:cs` down with them, and a reviewer deciding the change later would be
   * refusing a formatting change no document ever made.
   */
  prevSlots?: PropSlot[];
}

/**
 * The paragraph properties this model keeps. A `w:pPrChange` records the whole set from before the
 * change, so the same shape serves both the paragraph as it stands and its recorded predecessor.
 * Unlike the paragraph's own fields these are what its `w:pPr` *says* — no style inheritance, because
 * the record has to say what was written, not what it resolved to.
 */
export interface ParaStyle {
  styleId?: string;
  align?: string;
  indentLeft?: number;
  /**
   * The same left indent in hundredths of a character (`w:leftChars`) — how WPS spells a Chinese indent. 授权口径
   * 46 份的 `word/document.xml` 写 2,278 枚 `w:ind`，其中 `w:left` 412 枚、`w:leftChars` 432 枚；那 432 枚里 382 枚
   * 与缇值同写，剩下 50 枚没有缇同伴——48 枚写的是 `0`、2 枚写负数，**正数的 0 枚**，所以「字符数单独说一个正缩进」
   * 这一支在授权语料上没有实例。同写的对子算下来每字符 210 缇（`w:left="840" w:leftChars="400"`，10.5pt）与 398 缇
   * （`1207/303`，≈20pt）这一档，就是段落自己的磅数——两枚说的是同一个缩进，各说一遍。而 `w:defaultTabStop`
   * 46 份全写了（43 份 420、3 份 720），所以 增加缩进量 走一步 = 2 字符 = 420 缇（`.scratch/摆位-authorized.py`、
   * `.scratch/chars-zero-audience.ts`，均经产品代码与逐枚切片两趟口径）。
   *
   * The page draws the twips value, but the button has to move both: leave the character count behind and the file
   * contradicts itself, and Word lays the paragraph out from the count while our page shows the twips.
   */
  indentLeftChars?: number;
  /**
   * The rest of what `w:ind` can state about the first line and the right edge. `indentRight` and
   * `indentHanging` are twips like `indentLeft` (`w:right`, `w:hanging`); `indentFirstLineChars` is hundredths
   * of a character, which is how WPS writes 首行缩进 2 字符 (`w:firstLineChars="200"`).
   *
   * The first line has two spellings, and a file that states both is not contradicting itself. 授权口径 46 份的
   * `word/document.xml` 那 2,278 枚 `w:ind` 里 `w:firstLine` 写了 2,010 枚、`w:firstLineChars` 1,210 枚，逐枚看
   * （`.scratch/chars-zero-audience.ts`）：两枚同写的 1,116 枚中 **608** 枚两枚都是正数，商落在每字符 240 缇
   * （`w:firstLine="480" w:firstLineChars="200"` 348 枚）与 300/301/269 这些档上——都是段落自己的磅数，同一件事
   * 说两遍；**94** 枚写正缇而字符那枚写 `0`；414 枚两枚都写 `0`，那是「说不缩进」的写法。此外 894 枚只写缇，
   * **94** 枚只写字符而写的无一是正数（全是 `0`，且这 94 枚每枚都同时写着自己的 `w:hanging`）。
   *
   * 于是有两支取舍。段自己两枚都写时，上身的仍是缇那一枚——让字符那枚赢（§17.3.1.14 就是这么排的）会把那 94 枚
   * WPS 自己在对话框里给的绝对值抹掉；这一支的取舍按文件留下的数读，WPS 究竟画哪一枚 未取证（本机没有录屏授权）。
   * 段只用字符或只用 `w:hanging` 答过首行时，样式链上那枚缇不再叠上来，见 `indResolve`。
   */
  indentRight?: number;
  indentFirstLine?: number;
  indentFirstLineChars?: number;
  indentHanging?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: number;
  /**
   * How to read `lineSpacing`: as a multiple of one line, a fixed height, or a minimum. `w:line` is a different
   * unit under each of these, so leaving the rule out draws 固定值 22 磅 as 1.83 倍行距 — see `lineRuleOf` for
   * how often that mattered in the files WPS itself wrote.
   */
  lineSpacingRule?: LineSpacingRule;
  /**
   * `w:pageBreakBefore` (§17.3.1.65) as this `w:pPr` states it: `undefined` for nothing said, `true` for a bare
   * or `w:val="1"` element, and `false` for the `w:val="0"` one — which is the only spelling the authorized 46 份
   * has: 200 of its body paragraphs write `<w:pageBreakBefore w:val="0"/>` and none writes it on
   * (`.scratch/toggle-lent-chain.ts`, `.scratch/pbb-spellings.ts`). So presence is not what this field means.
   * `ownParaStyle` drops `false` the way it drops an indent the paragraph only borrowed, so a statement of
   * 「不断页」 travels as the paragraph's own slot bytes and never as a bare element the writer invents.
   */
  pageBreakBefore?: boolean;
  /**
   * The three page-keeping rules (§17.3.1.14 `w:keepNext`, §17.3.1.13 `w:keepLines`, §17.3.2.90
   * `w:widowControl`) as this `w:pPr` states them: `undefined` for nothing said, `true`/`false` for the written switch. A style's
   * statement is what `paraLent` lends, so these live in the same set as the indent a style lends; the paragraph
   * field docs say which of the three defaults to on when nobody says.
   */
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  /**
   * `w:snapToGrid` (§17.3.1.87): whether this paragraph's lines are measured in the section's line grid. The
   * switch defaults to **on**, so `undefined` (nothing in the paragraph's own `w:pPr` or its style chain) and
   * `true` (a bare `<w:snapToGrid/>`) both mean it snaps, and only a written `w:val="0"`/`"false"`/`"none"` takes
   * it off — the same reading `w:widowControl` above gets, for the same reason: the authorized 46 份 write the off.
   * 授权口径的逐槽读数（ElementTree 父链与正则切片两条口径同为 528 枚，`.scratch/authorized-families3.mjs` 同族）：
   * **正文与页眉页脚部件 410 枚**（`word/document.xml` 406 + 页脚 4），其中 **230 枚说关（`w:val="0"`）、180 枚裸写**；
   * 段自己的 `w:pPr` 里那 406 枚的落点是 顶层 `w:body` 段 222、表格内段 132、文本框内段 32、`w:sdtContent` 内段 20。
   * `word/styles.xml` 借出的那 **118 枚 = 110 枚写在 `w:style/w:pPr`（其中 108 枚说关）+ 8 枚写在 `w:style/w:rPr`**——
   * 前者就是本字段要走样式梯子的受众（一段自己不写、被上面的样式告知不贴格）；后者落在**字符样式**的 run 槽上，
   * 本模型只在段落层判定这条开关，那 8 枚画出来会怎样属 **未取证**（要 WPS 真机观察，本机没有录屏授权）。
   * 作废语料上那句「3,394 / 27,511 块、1,994 关、样式 275 枚中 267 关」已被上面这组数换掉。
   *
   * The slot is what counts: 授权 46 份里这 528 枚的父链只有 `w:pPr` 与 `w:style/w:rPr` 两种，**正文与页眉页脚部件里
   * `w:rPr` 落点 0 枚**（410 枚全在 `w:pPr`）。§17.3.1.87 does not put the switch in a body-run slot and
   * this reader does not look — `pPr.snapToGrid` cannot see a child of a child. Slicing `<w:pPr>…</w:pPr>` with a
   * non-greedy regex claims them anyway: the old pass counted 3,396 in the slices against 3,466 in the files, two
   * of them paragraph-mark statements read as paragraph properties and 70 never attributed
   * (`.scratch/wps-snap-where.mjs`) — that split is the quarantined 86 份's arithmetic and 待重证; on the authorized
   * set the same mis-attribution risk is bounded by counting parents with a real parser, which is how the numbers
   * above were taken. The walk itself stays: it is
   * what keeps every element reaching a slot instead of being silently unattributed.
   *
   * Like the three keep flags, this is not in `PARA_STYLE_KEYS`: the element travels in `pPrSlots`, and a
   * `w:pPrChange` record's own `w:snapToGrid` comes back as the record's bytes rather than as a field re-spelled.
   * It is written, though — `syncParaSlots` re-spells the element in the two spellings the authorized 46 份 use (bare
   * for on, `w:val="0"` for off) once the field moves, and `write-docx.ts` spells it for the paragraphs that have no
   * slots at all. Both compare against this snapshot, so a switch the paragraph only borrowed stays the style's.
   */
  snapToGrid?: boolean;
  /**
   * `w:wordWrap` (§17.4.77): WPS 段落对话框「中文版式」里那一格「允许西文在单词中间换行」。三态读法与
   * `w:snapToGrid` 同一条，**缺省相反**：这一枚的缺省是「不许」，所以只有写了才算说了话，裸写那句是「允许」。
   *
   * 受众是量的（`.scratch/cjk-typo-audience.py`，授权 46 件正文的 7,357 枚 `w:pPr`）：**200 枚说了
   * `w:wordWrap`，枚枚裸写不带 `w:val`**（9 份模板名），说关的 **0 枚**，样式层与 `w:docDefaults` 也 **0 枚**——
   * 链上没人借得出这一句，所以解析后的 `true` 只可能是段自己写的。同一段还常连着说 `w:kinsoku`／
   * `w:overflowPunct`／`w:autoSpaceDE`／`w:autoSpaceDN`（各 200 枚，同样裸写）与 `w:topLinePunct w:val="0"`；
   * 那五枚浏览器没有对应的腿（`hanging-punctuation` 在 Chromium 130 里不实现，中西文间距要 `text-autospace`
   * 那一族，而 WPS 侧到底怎么摆更未取证），故本轮不建模，只带 `w:wordWrap` 这一枚有 CSS 腿的。
   *
   * 与 `w:snapToGrid` 同样不在 `PARA_STYLE_KEYS` 里：整枚走 `w:pPr` 的槽位字节，`w:pPrChange` 记录带着它自己的
   * 那句回来；字段动了才由 `syncParaSlots` 按件的两种拼法改写（开=裸写，关=`w:val="0"`）。
   */
  wordWrap?: boolean;
  /**
   * `w:pBdr` (§17.3.2.36): the box drawn around this paragraph, one entry per side the element named. A side
   * whose `w:val` is `none`/`hidden`/`nil` is a side the file **said no line for**, which is a different sentence
   * from not naming the side at all — the same lesson `w:u w:val="none"` taught in §85, and the shape the authorized
   * 46 份 mostly write: of the 26 `w:pBdr` elements in the 46 authorized docx-family files
   * (`.scratch/pbdr-audience.ts`, 7,430 body paragraphs), 16 name all four sides as `none` and 2 name a `none`
   * bottom, while only 8 state a real line (`w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"`). All
   * 26 sit outside the body — 24 in `word/styles.xml` (the 16 and the 8) and the other 2 in a header part — and 0
   * in the 7,430 body paragraphs, so a reader that drew a line for `none` would rule across the 16 chains that
   * refused one, and one that dropped `none` while lending would let an ancestor's line through a cancellation.
   * `w:between` and `w:bar` are read here and not drawn — 0 of the 26 state either, and `w:shadow`
   * is stated by none of them.
   */
  borders?: ParaBorderSet;
  /**
   * `w:pPr/w:shd@w:fill` (§17.3.2.33): the ground behind this paragraph's lines, as the file signed it — either
   * a colour (`#RRGGBB`) or, when the fill says no colour, the token it used instead of one, in practice `auto`.
   * Only the fill is a field; a `w:val` other than `clear` would put a pattern in front of it, and 30/30 of the
   * paragraph-shading elements in the authorized corpus say `w:val="clear" w:color="auto" w:fill="FFFFFF"`
   * (`word/document.xml`'s `w:pPr` only — the same walk finds no `w:shd` under any other parent or in any other
   * part, so `styles.xml` and the header parts have 0 of them). 件数有两说、枚数只有一说：那 30 枚落在 4 个文件路径上
   * （`office rental contract.dotm` 1 枚 + `share transfer agreement.dotm` 14 枚，各在 zh_CN 与 zh_HK 两套模板里有一份），
   * 所以 `.scratch/pbdr-audience.ts` 按路径报 4 份、按文件名去重是 2 份。
   * `w:fill="auto"` has 0 paragraph instances there, but it is not a made-up shape: the same corpus writes it on
   * 1,600 of its 1,646 `w:tcPr/w:shd` (the other 46 say `FFFFFF`), which is the file's own idiom for 「这一格要页面自己的地」. So a stated
   * `auto` is kept as a token rather than collapsed into `undefined`, the same way `w:u w:val="none"` (§85) and
   * `w:color="auto"` (§ `TableBorder.color`) are kept: **said no ground** and **said nothing** are two sentences,
   * and folding the first into the second would let an ancestor style's fill show through a refusal. `undefined`
   * is the second one — no `w:shd`, or one that states no fill.
   */
  shading?: string;
  /**
   * `w:pPr/w:tabs` (§17.4.19) as this one rung states it, `w:val="clear"` included. Unlike every other key in
   * this interface, the chain does not stop at the rung that first states it: `paraLent` folds each rung's list
   * into the one below it per position (`mergeTabStops`), because a style that adds one stop does not cancel
   * the stops the style under it wrote.
   */
  tabStops?: TabStop[];
  numPr?: { numId: number; ilvl: number };
  /**
   * `w:outlineLvl` (§17.3.2.26) as the file signed it: 0 is the first heading level through 8 the ninth, and
   * **9 is body text**, which is why the raw number is kept rather than folded into "a level or none" — the
   * authorized 46 份 do write that 9: of the 7,430 paragraph blocks the census walks, 202 state a level
   * themselves and the 408 resolved levels include 26 at level 9 (`.scratch/outline-census.ts`), and a 9 on the
   * paragraph is a level its style would otherwise lend being taken back.
   */
  outlineLevel?: number;
}

/** `w:spacing@w:lineRule` (§17.3.1.27) — and the only three literals the authorized 46 份 ever state (auto 2,890 / exact 1,004 / atLeast 54, `.scratch/corpus-recheck2.mjs`). */
export type LineSpacingRule = "auto" | "exact" | "atLeast";

/** `w:pPrChange`: the paragraph properties from before someone revised this paragraph. */
export interface ParagraphFormatRevision {
  kind: "format";
  id: number;
  author: string;
  date?: string;
  prev: ParaStyle;
  /**
   * That set as the record wrote it, child by child. `ParaStyle` is the properties the model can name,
   * and a paragraph's `w:pPr` holds far more than those: a refusal rebuilt from the fields put the named ones
   * back and took the record's own `w:shd`, `w:tabs` and `w:snapToGrid` down with them — the same set the
   * reviewer was being asked to decide.
   */
  prevSlots?: PropSlot[];
}

/**
 * What a `w:tblPr` says about the table itself — the properties the page draws, with nothing taken from the
 * table's style. A `w:tblPrChange` records the whole set as it was written before the change, so the same
 * shape serves the table as it stands and its recorded predecessor, the way `ParaStyle` serves a paragraph.
 */
export interface TableStyle {
  styleId?: string;
  widthTwips?: number;
  widthType?: string;
  indentTwips?: number;
  align?: string;
  layout?: string;
  borders?: BorderSet;
  /** `w:tblCellMar`: the padding every cell of this table gets unless it or its row says otherwise. */
  cellMargin?: MarginTwips;
  /**
   * `w:tblPr/w:shd@w:fill`: the ground under the whole table, same two spellings as `ParaStyle.shading`
   * (a colour, or the token a file put where a colour would go). 授权口径实测 0 枚（`.scratch/shd-audience.ts`：
   * 46 份件里 `w:tblPr` 那一层一枚也没有，只有 `w:tcPr` 1,646 与 `w:pPr` 30），所以**这一层的形状是 未取证**：
   * 摆的是 schema（§17.4.60）给的位置与另一家实现（ONLYOFFICE 的 `CT_TblPr` 读端确有 `case "shd"`）认的这一枚孩子，
   * 不是 WPS 自己怎么写。读通是为了「件里要出现时不静默丢」，写回才有对话框那一条腿。
   */
  shading?: string;
}

/** What a `w:trPr` says about its row, minus the row's own record and the children it only carries. */
export interface RowStyle {
  heightTwips?: number;
  heightRule?: string;
  header?: boolean;
  keepTogether?: boolean;
  hidden?: boolean;
  /** `w:jc`: where the row's text sits — under a paragraph's own alignment, over the cell's. */
  jc?: string;
}

/** What a `w:tcPr` says about its cell: width, span, shading, borders, alignment, wrapping. */
export interface CellStyle {
  widthTwips?: number;
  widthType?: string;
  shading?: string;
  spanCols?: number;
  spanRows?: number;
  valign?: string;
  noWrap?: boolean;
  textDirection?: string;
  borders?: BorderSet;
  /** `w:tcMar`: this cell's own padding, which wins over the table's and over its row's exception. */
  margin?: MarginTwips;
}

/**
 * What a record on a table part holds: who signed it, when, and the properties as the file wrote them
 * inside it. The text is kept next to the fields because a refusal has to put the set back *complete* —
 * including the properties this model has no field for, which the fields alone cannot spell.
 */
interface PropertyRecordBase {
  kind: "format";
  /** `w:id`, shared with every other change in the document. */
  id: number;
  author: string;
  date?: string;
  /** The properties the record names, each as the source text of its element. */
  prevSlots: PropSlot[];
}

/** `w:tblPrChange` on a table: the table's own properties from before someone revised them. */
export interface TableFormatRevision extends PropertyRecordBase {
  tag: "tblPrChange";
  prev: TableStyle;
}

/** `w:trPrChange` on a row: the row's own properties from before someone revised them. */
export interface RowFormatRevision extends PropertyRecordBase {
  tag: "trPrChange";
  prev: RowStyle;
}

/**
 * One undecided revision record on a table cell, as its `w:tcPr` wrote it. `cellIns` and `cellDel` sign the
 * cell itself arriving or going — which is how OOXML records a revised column: every cell of it carries the
 * one `w:id` — while `cellMerge` names the single property it changed (`w:vMerge`) and `tcPrChange` records
 * the cell's whole set. The two that move a cell rather than redress it record no properties, so `prev` is
 * the empty set for them.
 */
export interface CellRevision {
  tag: "cellIns" | "cellDel" | "cellMerge" | "tcPrChange";
  /** `w:id`, in the one number space every other change in this document shares. */
  id: number;
  author: string;
  date?: string;
  /** The properties the record names, each as the source text of its element — what a refusal puts back. */
  prevSlots: PropSlot[];
  /** The same properties as the fields the page draws. */
  prev: CellStyle;
}

/**
 * One change still waiting for a decision, as a 审阅窗格 row: what kind it is, who signed it, when,
 * and what it covers. `revisionEntries()` numbers these from 1 in document order, which is how WPS'
 * `Revisions.Item(i)` counts them — the pane lists one row per pending change of any kind, so it
 * needs this one collection rather than the four separate lists.
 *
 * The fields are the ones WPS' own JS API reads off a revision (`wps.js`: `getAllRevisions` pushes
 * `{index, type: rev.Type, author: rev.Author, date: rev.Date, text: range.Text, begin: range.Start,
 * end: range.End}` for `i` from 1 to `Revisions.Count`), with the position given inside the paragraph
 * rather than as a global character position, because that is what this model has.
 */
export interface RevisionEntry {
  /** 1-based, in the order the changes sit in the document. */
  index: number;
  /** The `w:id`, which is what accepting or refusing one takes. */
  id: number;
  /** Our own words for the kind; WPS' numeric `WdRevisionType` values are not evidenced here. A moved
   *  passage reports the half it is — `move-from` where the text was, `move-to` where it landed — and a
   *  row whose `text` is `¶` is a break, so a moved paragraph comes as a pair of those. `row-insert` and
   *  `row-delete` are a whole table row (`w:trPr/w:ins`, `w:trPr/w:del`); its `text` is what the row holds
   *  and its position is the row's first paragraph. The table kinds are a revision on a table part rather
   *  than on characters — `table-format` (`w:tblPrChange`), `row-format` (`w:trPrChange`), `cell-format`
   *  (`w:tcPrChange`), `cell-merge` (`w:cellMerge`), `cell-insert` and `cell-delete` (`w:cellIns`,
   *  `w:cellDel`) — each placed at the first paragraph of the table, row or cell it changed, with no range
   *  of its own because there are no characters to range over. */
  type:
    | "insert"
    | "delete"
    | "run-format"
    | "paragraph-format"
    | "insert-mark"
    | "delete-mark"
    | "move-from"
    | "move-to"
    | "row-insert"
    | "row-delete"
    | "table-format"
    | "row-format"
    | "cell-format"
    | "cell-merge"
    | "cell-insert"
    | "cell-delete";
  author: string;
  date?: string;
  /** The characters the change covers — struck or added text, the revised run, the paragraph whose
   *  properties changed, or `¶` for a break that was added or removed. */
  text: string;
  /** Which paragraph the change sits in, counting the paragraphs in the order `revisionEntries()`
   *  walks them, table cells included. `paragraphs()` is that list. */
  para: number;
  /** Where its range starts and ends inside that paragraph, in characters. A break has no characters
   *  of its own before the end of the paragraph it terminates, so both are the paragraph's length.
   *  A change covering more than one paragraph reports the paragraph and range where it begins. */
  begin: number;
  end: number;
}

export interface Paragraph {
  kind: "paragraph";
  /**
   * A `w:ins` or `w:del` on this paragraph's own mark (`w:pPr/w:rPr`): the break that ends this
   * paragraph was added or removed and nobody has decided yet. It belongs to the paragraph it
   * terminates, so refusing either kind merges this paragraph with the one after it.
   */
  markRev?: Revision;
  /** A `w:pPrChange` on this paragraph: properties someone revised and nobody has decided yet. */
  revFormat?: ParagraphFormatRevision;
  styleId?: string;
  styleName?: string;
  align?: string;
  numPr?: { numId: number; ilvl: number };
  indentLeft?: number;
  indentLeftChars?: number;
  /** The rest of what `w:ind` states about this paragraph, resolved through its style. See `ParaStyle`. */
  indentRight?: number;
  indentFirstLine?: number;
  indentFirstLineChars?: number;
  indentHanging?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: number;
  /** `w:spacing@w:lineRule`: which unit `lineSpacing` is written in. See `ParaStyle`. */
  lineSpacingRule?: LineSpacingRule;
  /**
   * `w:pageBreakBefore` (§17.3.1.65): a page break is forced before this paragraph, and the renderer's
   * `page-break-before:always` is the layout's only page break this field owns. Read as the switch, not as the
   * element's presence: every one of the authorized 46 份 that writes the element writes `w:val="0"` — 200 body
   * paragraphs, and 180 of them reachable as model paragraphs, the other 20 travelling verbatim inside a `w:sdt`
   * container — and before that reading the renderer forced a break on all 180 (`.scratch/pbb-before-after.ts`,
   * counted per file, so its printed column is files and not breaks). No style lends it: `word/styles.xml`
   * in the 46 writes `w:pageBreakBefore` 0 times and 0 paragraphs ask the chain (`.scratch/toggle-lent-chain.ts`),
   * so the paragraph's own statement is the whole story and `false` and silence draw the same page.
   * The installed app is measured on this: `office rental contract.dotm` states the switch 21 times and draws 0
   * breaks now against 21 on the previous build (`.scratch/pbb-dom-count.js`, both readings via `.scratch/cdp-eval.mjs`).
   */
  pageBreakBefore?: boolean;
  /**
   * `w:keepNext` (§17.3.1.14): this paragraph's last line stays on the page where the next block starts.
   * `undefined` means neither this paragraph's `w:pPr` nor its style chain said anything. The packer honours it
   * through the `data-keep-next` the renderer stamps.
   */
  keepNext?: boolean;
  /** `w:keepLines` (§17.3.1.13): this paragraph's lines do not split across pages. Same reading as `keepNext`. */
  keepLines?: boolean;
  /**
   * `w:widowControl` (§17.3.2.90): a page break inside this paragraph leaves at least two lines on each side of
   * the seam. The reading is inverted against the two keep flags: `ST_OnOff` here defaults to **on**, so
   * `undefined` (nobody in the paragraph's own `w:pPr` or its style chain said) means the rule applies and only
   * `false` — a written `w:val="0"`/`"false"`/`"none"` — turns it off. That is why the chain has to be read, and
   * the chain is where nearly all of this corpus's answer lives. Over the 7,393 `w:p` of the authorized 46 份 —
   * body, table cells and text boxes alike, since the file writes the element in all three — asked one rung of
   * §17.7.2's ladder at a time (`.scratch/toggle-lent-chain.ts`): 246 paragraphs state it themselves (160 off,
   * 86 on), 58 borrow it from the named style one hop up (30 off, 28 on), 1,370 borrow **off** from two hops up,
   * 5,601 borrow **off** from the default paragraph style, and 118 have nobody above them to ask. So 7,001
   * paragraphs are told not to widow-protect without writing a word about it: a reader that only looked at the
   * paragraph's own `w:pPr` widow-protected every one of them, which is what #126 did. `w:docDefaults` states none
   * of it — all 46 份 carry `w:docDefaults/w:pPrDefault/w:pPr` and all 46 leave it empty (`.scratch/dd-control.ts`),
   * so the default paragraph style is the last rung with anything to lend. The model side agrees: 2,937 of the
   * 3,120 paragraphs reachable through `doc.blocks` resolve to `false`, and none resolves to `undefined` while its
   * own `w:pPr` stated the switch (`.scratch/widow-model.ts`; that walk stops at `carried` containers, so it counts
   * a smaller denominator than the file's own). See `LENT_PARA_KEYS` for the same split across the family.
   */
  widowControl?: boolean;
  /**
   * `w:snapToGrid`: whether this paragraph's lines are measured in the section's line grid. Same three-state
   * reading as `widowControl` (defaults to on, and only a written off takes it off), and read through the style
   * chain for the same reason — the authorized 46 份 lend it: 404 paragraphs state it (224 off, 180 on), 52 take
   * it off from a style (40 one hop up, 12 two), 154 from the default paragraph style, 16 from a style that says
   * on, and 6,783 get nothing from anyone (`.scratch/toggle-lent-chain.ts`). See `ParaStyle` for the per-slot
   * counts and `render.ts`'s `lineHeight` for what the page does with it.
   */
  snapToGrid?: boolean;
  /**
   * `w:wordWrap`：这一段允不允许西文在单词中间断开。三态与 `ParaStyle` 上那一条同口径（缺省「不许」，裸写那句
   * 才是「允许」），解析后同样把样式链借来的算进来——授权口径里链上 0 枚，所以这里的 `true` 只可能是段自己写的。
   * 画侧那条腿在 `render.ts`（`word-break:break-all`）。
   */
  wordWrap?: boolean;
  /**
   * `w:pBdr`, this paragraph's own sides merged over the ones its style chain names — the merge is per side,
   * because §17.7.2 resolves a property at a time and a paragraph that cancels only its bottom line is not
   * cancelling the box its style drew. See `ParaStyle` for what the authorized 46 份 write there and why
   * `w:val="none"` is a statement rather than a silence. Nothing writes it
   * unless the field moved: read paragraphs carry the element's own bytes in `pPrSlots`.
   */
  borders?: ParaBorderSet;
  /**
   * `w:pPr/w:shd@w:fill`, this paragraph's own or the one its style chain states; see `ParaStyle` for why a
   * stated `auto` is kept as a token instead of read as a silence, and for the 30 枚 / 4 个路径 / 2 份名 that
   * audience actually is. It exists because the page has to draw a ground the file named, and every one of those
   * 30 says `w:val="clear"` with a fill.
   * The pattern forms have no field and are not
   * drawn — 0 of the authorized 46 份 asks for one, since all 1,676 `w:shd` those files write say `w:val="clear"`
   * (`.scratch/shd-audience.ts`) — and their bytes still travel in `pPrSlots`.
   */
  shading?: string;
  /**
   * `w:pPr/w:tabs` resolved for this paragraph: its own stops over the ones its style chain wrote, per position
   * (`mergeTabStops`), with every `w:val="clear"` already spent. So this list only ever names a stop that paints.
   * Not in `PARA_STYLE_KEYS`: no command moves it yet, and the element's bytes travel in `pPrSlots` like the box
   * and the ground beside it.
   */
  tabStops?: TabStop[];
  /**
   * `w:outlineLvl`, this paragraph's own or the one its style chain states; see `ParaStyle` for what the number
   * means and `headingLevel` for the heading the page draws from it.
   */
  outlineLevel?: number;
  /**
   * The size this paragraph's own mark resolves to, in points: `w:pPr/w:rPr/w:sz` (§17.3.1.4), and where the mark
   * states none, what the paragraph's style chain and `w:rPrDefault` lend it — the same ladder a body run of this
   * paragraph reads. It exists for one thing the page draws: a numbered paragraph's symbol is a character run of
   * its own (§17.3.1.12), and where the *level* states no size this is the number the file left standing. Without
   * it the mark falls through to the page stylesheet's 12pt, a size no document asked for. The audience that
   * justified it — 331 of the corpus's 396 such paragraphs resolving elsewhere than 12pt, with
   * `NumberLevel.sizePt` covering the other 70 (`.scratch/mark-size.ts`) — was measured on the quarantined 86 份
   * cohort, and the paragraph-mark-size audience has not been re-counted on the authorized 46: 待重证.
   * Nothing writes it: the mark's own properties travel as `pPrSlots` text, so a save spells the file's bytes.
   */
  markSizePt?: number;
  /**
   * The two faces this paragraph's own mark resolves to — `w:pPr/w:rPr/w:rFonts@w:ascii` and `@w:eastAsia`
   * (§17.3.1.4), each falling back on its own to what the style chain and `w:rPrDefault` lend, the way
   * `readRunStyle` reads one attribute at a time (§17.7.2 stops per property). They exist for the same thing
   * `markSizePt` does: the numbering symbol is a character run of its own (§17.3.1.12), and where the *level*
   * names no font this is what the file left standing. The audience that justified it (424 listed paragraphs on a
   * level that names none: 172 getting the face for their own script from this mark, 234 more from the chain, 18
   * from nothing at all, `.scratch/mark-font.ts`) was measured on the quarantined 86 份 cohort and has not been
   * re-counted on the authorized 46: 待重证. The two faces are a stack, not a choice — the browser dispatches
   * by script, so `三、` takes the East Asian one and `1.` the Latin one, exactly as the paragraph's body text does.
   * Nothing writes them: the mark's properties travel as `pPrSlots` text, so a save spells the file's bytes.
   */
  markFont?: string;
  /** The `@w:eastAsia` half of the pair above — the face this mark's Chinese characters take. */
  markFontEastAsia?: string;
  /**
   * The colour this paragraph's own mark resolves to: the `w:rPr/w:color` that sits beside the `w:rFonts` above
   * (§17.3.1.4), falling back on the style chain and `w:rPrDefault` the way `readRunStyle` does — where
   * `w:val="auto"` counts as no statement, and a shape that is not a colour counts as one neither. It exists for
   * the same reason the two faces do: the numbering symbol is a character run of its own (§17.3.1.12).
   *
   * The level has no say here, because no file in hand tests one. The measurements that said so — 8 of 617 levels
   * writing `w:lvl/w:rPr/w:color` with **8/8 of those having no listed paragraph sitting on them**
   * (`.scratch/mark-color.ts`), and 82 of the 473 listed paragraphs resolving a colour at this rung (74 from the
   * mark itself, 8 from the chain), all 82 differing from the ink the App's own stylesheet gives a page, 74/74 of
   * the mark-named ones agreeing with the paragraph's own text (`.scratch/make-markcolor-fixture.ts`), 391
   * resolving nothing) — all came from the quarantined 86 份 cohort, and neither the level side nor the paragraph
   * side has been re-counted on the authorized 46: 待重证, so this is the reading the files left standing, not a
   * measured precedence. What this field fixes is not a difference from the body but the symbol having been left
   * out of the ink the file chose for its paragraph. Nothing writes it: the mark's properties travel as `pPrSlots`
   * text, so a save spells the file's bytes.
   */
  markColor?: string;
  /**
   * The attributes written on this paragraph's own opening tag, as the part spelled them — `w14:paraId`,
   * `w14:textId`, `w:rsidR`, `w:rsidP`. The model has no field for any of them and a save writes this text
   * back onto the tag, the way `TableCell.tcPrXml` writes a property set it cannot explain. They are kept as
   * text rather than as names because the prefix is bound outside the element: `w14` means the words
   * document namespace only where an `xmlns:w14` says so, and a parsed name cannot give that binding back.
   */
  attrsXml?: string;
  /**
   * This paragraph's own `w:pPr`, child by child, as the part wrote it. The fields above are what the page
   * draws; a save takes the *set* from here, so a property the model has no field for — `w:keepNext`,
   * `w:widowControl`, `w:tabs`, `w:pBdr`, the paragraph mark's own fonts, bold, size and
   * colour — comes back the way it went in instead of being quietly stripped by a document that regenerates.
   * A `w:pPrChange` is not a slot: the record is the `revFormat` field, and a save spells it last as before.
   * A paragraph the editor made has no slots and is spelled from the fields.
   */
  pPrSlots?: PropSlot[];
  /**
   * The paragraph properties as the model worked them out at the moment the paragraph was read, style
   * inheritance included. A save compares against it, so a value the paragraph only borrowed from its style
   * is not written onto the paragraph as though somebody had set it, while a value the page really did change
   * still lands.
   */
  asRead?: ParaStyle;
  inlines: Inline[];
}

export interface TableCell {
  paragraphs: Block[];
  widthTwips?: number;
  /** `w:tcW/@w:type`: `dxa`, `pct`, `auto`, `nil` — the unit the width is counted in. */
  widthType?: string;
  shading?: string;
  spanCols?: number;
  /** 1 starts a vertical merge, 0 continues the one above, undefined is an ordinary cell. */
  spanRows?: number;
  /**
   * What the cell's row says about itself (`w:trPr`), the same object shared by every cell of the row:
   * a row has no object of its own in this model, and a record or a height belongs to the row, not to a
   * cell. Sharing one object is also what carries it through a rebuild of the table.
   */
  row?: RowInfo;
  /** `w:vAlign`: where the cell's text sits when the row is taller than it. */
  valign?: string;
  /** `w:noWrap`: the cell refuses to break its text across lines. */
  noWrap?: boolean;
  /** `w:textDirection`: vertical or rotated text, e.g. `tbRlV`. */
  textDirection?: string;
  /** `w:tcBorders`, one entry per side the cell draws for itself. */
  borders?: BorderSet;
  /** `w:tcMar`: this cell's own padding, side by side; a side it does not name comes from row or table. */
  margin?: MarginTwips;
  /**
   * The cell's own `w:tcPr`, exactly as the file wrote it. What the page draws comes from the fields
   * above; what a save writes comes from here, so a property set the model cannot change — `w:cnfStyle`,
   * the row-conditional band, a pending `w:tcPrChange` — returns the way it went in. `w:tcMar` is read
   * into `margin` but written from here: reading a padding is not authority to invent one.
   */
  tcPrXml?: string;
  /** The attributes on this cell's own `w:tc` opening tag, exactly as written — see `Paragraph.attrsXml`. */
  attrsXml?: string;
  /**
   * Every revision record this cell's `w:tcPr` carries and nobody has decided. A list because the schema lets
   * one cell hold several at once — a column inserted while 修订 was on, whose cell was then rewidened, is two
   * changes with two `w:id`s — and dropping the second would leave a reviewer's change undecided in silence.
   * Each stays inside `tcPrXml` as the file wrote it: this is what the pane lists and a decision rewrites,
   * not a second copy of the cell's properties.
   */
  revs?: CellRevision[];
}

/**
 * One edge of a border set (`w:tblBorders`, `w:tcBorders`). `val` is the stroke name Word uses —
 * `single`, `dashed`, `dotted`, `double`, `thick`, `none`, `hidden`, `nil`, `space` — `size` is in
 * eighths of a point and `space` is the gap in points, both as the file gives them.
 */
export interface TableBorder {
  val: string;
  size?: number;
  space?: number;
  color?: string;
}

/** The sides a border set can name. `start`/`end` are the bidi-aware spellings of left/right. */
export type BorderSet = Partial<
  Record<"top" | "left" | "bottom" | "right" | "start" | "end" | "insideH" | "insideV" | "tl2br" | "tr2bl", TableBorder>
>;

/**
 * The sides `CT_PBdr` (§17.3.2.36) lets a paragraph state, in the order it gives them. It is not `CT_Borders`:
 * a paragraph names `w:between` (the rule between this paragraph and the next one that shares it) and `w:bar`
 * where a table names `w:insideH`/`w:insideV`. An edge carries the same four attributes in both, so a side is
 * still a `TableBorder`.
 */
export const PBDR_ORDER = ["top", "left", "bottom", "between", "right", "bar"] as const;
export type ParaBorderSet = Partial<Record<(typeof PBDR_ORDER)[number], TableBorder>>;

/** The four sides a cell-margin set names, in the order `CT_TblCellMar` gives them. */
export const MARGIN_SIDES = ["top", "left", "bottom", "right"] as const;

/**
 * Cell padding, twips per side: `w:tblCellMar` (§17.3.2.82) says it for a whole table and `w:tcMar`
 * (§17.3.2.69) for a single cell. A side is present only when the file said it — `w:w="0"` means flush,
 * which is a different fact from saying nothing, and a page that cannot tell them apart puts its own
 * spacing back on a table that asked for none. The audience behind that, on the authorized 46 份: 10 份 hold a
 * table and **all 10** state margins — `w:tblCellMar` on all 58 tables, a row-level `w:tblPrEx/w:tblCellMar` on 693
 * of the 698 rows, `w:tcMar` on 888 cells in 2 of them — those 10 being 5 template names each shipped twice (zh_CN
 * and zh_HK), and counting by file name is what used to report 5 and 1. Every one of the cohort's 6,556 side
 * elements that states a `w:w` signs its unit `w:type="dxa"`; none leaves it out
 * (`.scratch/authorized-table-margin.py`). The old reading of all 115 table-bearing files saying at least one side
 * is the quarantined cohort's arithmetic.
 */
export type MarginTwips = Partial<Record<(typeof MARGIN_SIDES)[number], number>>;

/**
 * What a row says about itself, in `w:trPr`. One object shared by the row's cells, because the row has
 * no object of its own here: `rev` is 整行的插入或删除, the rest is height, header repeat, splitting and
 * hiding. `slots` carries the children the model does not own — `w:gridBefore`, `w:trPrChange` — verbatim.
 */
export interface RowInfo {
  /**
   * A `w:ins` or `w:del` on the row itself: somebody added or removed the whole row and nobody has
   * decided yet.
   */
  rev?: Revision;
  /**
   * A `w:trPrChange` on this row: its height, header repeat, splitting or hiding someone revised and nobody
   * has decided yet. It stays inside `slots` as the file wrote it, the way a cell's record stays in its
   * `tcPrXml`, so carrying the properties the model has no field for is one job, not two.
   */
  revFormat?: RowFormatRevision;
  /** `w:trHeight/@w:val`, twips. */
  heightTwips?: number;
  /** `w:trHeight/@w:hRule`: `atLeast`, `exact` or absent for the row's own content height. */
  heightRule?: string;
  /** `w:tblHeader`: this row repeats at the top of every page the table runs across. */
  header?: boolean;
  /** `w:cantSplit`: the row must not break across a page. */
  keepTogether?: boolean;
  /** `w:hidden`: the row is there but draws nowhere. */
  hidden?: boolean;
  /**
   * `w:jc`: where this row's text sits. The authorized 46 份 hold 190 枚 on 4 templates (all `center`), and
   * 1,294 cells have the row's word as the only one — no `w:tcPr/w:jc` anywhere in the cohort (0 枚), no
   * paragraph `w:pPr/w:jc` in those cells. The precedence is the schema's: paragraph > cell > row, so a
   * paragraph that names its own alignment keeps it while the rest of the cell takes the row's.
   */
  jc?: string;
  /**
   * The row's `w:tblPrEx` band, read down to the one property the page draws: cell padding as this row is
   * to lay itself out, overriding the table's. The band is a child of `w:tr` rather than of `w:trPr` — the
   * spelling the files show (all 1,660 rows that carry margins in it, WPS writing it ahead of the row's
   * properties) came from the quarantined cohort, and `w:tblPrEx` has not been re-counted on the authorized
   * 46 份: 待重证 — so it is read from the parsed row and carried back as the source text it came in
   * with (`extras`), never re-spelled from this field.
   */
  cellMargin?: MarginTwips;
  slots?: PropSlot[];
  /**
   * The row's own children that are neither `w:trPr` nor `w:tc` — the `w:tblPrEx` band, which says this row
   * draws with table borders and shading of its own rather than the table's. It is the row's FIRST child there,
   * and where it sits is carried too: a save that returns it
   * after the cells rewrites the rows in the wrong order. The counts that measured how many rows that was (6257
   * putting it before `w:trPr`, 123 before the first cell when the row has no properties at all, 6380 in all,
   * across the 86 real WPS table-bearing files) came from the quarantined cohort and are 待重证 on the authorized
   * 46 — including whether every row really puts it first, which the old "always" asserted. `afterCells` counts
   * the cells it followed and `beforeProps` says whether it preceded `w:trPr`.
   */
  extras?: RowExtra[];
  /** The attributes on this row's own `w:tr` opening tag, exactly as written — see `Paragraph.attrsXml`. */
  attrsXml?: string;
}

/**
 * One child of `w:trPr`, kept as its own source text under the tag it came with. `w:trPr` is the one
 * property set where the model has to rewrite a single child — the row's own signature, decided by a
 * reviewer — without disturbing where its neighbours sit, so the set is carried child by child.
 */
export interface PropSlot {
  tag: string;
  xml: string;
}

/** A row child the model has no field for, kept with the siblings it was read between. See `RowInfo.extras`. */
export interface RowExtra extends PropSlot {
  /** How many `w:tc` preceded it when it was read. */
  afterCells: number;
  /** True when it came before this row's `w:trPr`. */
  beforeProps: boolean;
}

export interface Table {
  kind: "table";
  columns: number;
  rows: TableCell[][];
  styleId?: string;
  /** `w:tblW`: the table's own width, counted in `widthType`. */
  widthTwips?: number;
  widthType?: string;
  /** `w:tblInd`: how far the table sits from its column band's left edge, twips. */
  indentTwips?: number;
  /** `w:jc`: `left`, `center`, `right` — where the table sits inside the text column. */
  align?: string;
  /** `w:tblLayout/@w:type`: `fixed` honours `columnWidths`, anything else fits columns to content. */
  layout?: string;
  /** `w:tblBorders`, the table's own lines including the inside ones. */
  borders?: BorderSet;
  /** `w:tblCellMar`: the padding this table gives every cell that does not say its own. */
  cellMargin?: MarginTwips;
  /** `w:tblPr/w:shd@w:fill`: the ground under the whole table; see `TableStyle.shading` for the audience. */
  shading?: string;
  /** `w:tblGrid/@w:w` of each column, in order — what a fixed layout sizes its columns from. */
  columnWidths?: number[];
  /** The source text of `w:tblGrid` whole, so every `w:gridCol` and its attributes come back. */
  gridXml?: string;
  /**
   * The table's own `w:tblPr`, exactly as the file wrote it. The fields above are what the page draws;
   * a save writes this, so `w:tblLook`, `w:tblCellMar` and a pending `w:tblPrChange` are not re-spelled
   * from a partial set but returned — see the note on `TableCell.tcPrXml`. The margins are the one
   * property the page now reads as fields and the save still takes from here: reading a number is not
   * the same authority as being allowed to change one, and no command has that yet.
   */
  tblPrXml?: string;
  /**
   * A `w:tblPrChange` on this table: its width, alignment, borders or layout someone revised and nobody has
   * decided yet. The record stays inside `tblPrXml` as the file wrote it, so the properties the model has no
   * field for are carried once rather than re-spelled from a partial set.
   */
  revFormat?: TableFormatRevision;
  /**
   * Children of `w:tbl` the model has no fields for, each with the row it sat after (`-1` = before the first
   * row, which is where the schema puts them once the grid is closed). The OOXML table allows range markup and
   * containers between its rows, and WPS does use it: a `w:bookmarkEnd` written between two rows is how a
   * bookmark that started in the paragraph above the table gets closed. Ten of the 86 table-bearing files in
   * the quarantined cohort carried one, so the slot is real; on the authorized 46 份 the same walk finds 0
   * `w:bookmarkStart` and 0 `w:bookmarkEnd` between table rows (the 383 paired bookmarks those files do write
   * sit 0-strong inside any `w:pPr` too), so the audience there is 未取证. Without this slot the start survives a
   * save while the end goes, leaving a bookmark no reader can pair.
   */
  betweenRows?: { afterRow: number; tag: string; xml: string }[];
}

/**
 * A child of the body, or of a table cell, that the model has no fields for: a `w:sdt` content control, a
 * `w:customXml` island, anything a producer wrote between the paragraphs. The page shows the text inside it
 * and edits none of it, and a save writes back the source text it came in with, in the place it held.
 * Without this kind, a document with a table of contents in it loses the control and everything anchored
 * in it on the first keystroke — and the shapes are the ECMA-376 ones, since no file here has ever contained
 * one; whether the authorized 46 份 holds such a container has not been re-counted, so that absence is 待重证.
 */
export interface Carried {
  kind: "carried";
  /** The element's local name — `sdt`, `customXml`. It says what was kept, not what it means. */
  tag: string;
  /** The element's own source text, open tag to close tag, unchanged. */
  xml: string;
  /** What a reader would see inside it, for the page only. Never written back. */
  text: string;
}

export type Block = Paragraph | Table | Carried;

/**
 * One direct child of a paragraph, as the reader found it: either the node whose meaning the model knows, or
 * the source text of the one it does not. Which of the two it is decides whether the child is understood or
 * carried, and the pair is only built from source spans when those spans and the parsed tree agree.
 */
/**
 * A paragraph child, as the parsed tree explains it and — where the parsed tree and the source spans
 * accounted for the same paragraph — as the part actually wrote it. A step with no `node` is machinery this
 * model carries rather than reads, and its `span` is all of it.
 */
type ParaChild = { tag: string; node?: any; span?: { tag: string; xml: string }; name?: string };

/** The tags a row's own record can be spelled as, inside its `w:trPr`. */
export const ROW_REVISION_TAGS: ReadonlySet<string> = new Set(["ins", "del", "moveFrom", "moveTo"]);

/** `CT_Borders`, the sides named by both `w:tblBorders` and `w:tcBorders`. */
export const BORDER_ORDER = ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl"] as const;

/** The four sides a `w:pgBorders` element can state a line for, in the order §17.3.2.10 lists its children. */
export const PAGE_BORDER_SIDES = ["top", "left", "bottom", "right"] as const;
export type PageBorderSide = (typeof PAGE_BORDER_SIDES)[number];

/**
 * One side of a page border: the line the file asks for and the gap it is held off the edge with. `size` is the
 * file's `w:sz`, eighths of a point, spelled as `TableBorder` spells it so the one stroke table draws a table's
 * rule and the page's; `space` counts whole points (§17.3.2.10), which is why the two keep separate names —
 * they are multiplied by different factors on the way to a pixel.
 *
 * `val` is kept as the file spells it rather than turned into a shape, because the page's border styles are a
 * long enumeration, the drawing side maps them, and a token this model did not invent has to be able to come
 * back unchanged.
 */
export interface PageBorder {
  val: string;
  size?: number;
  space?: number;
  color?: string;
  /** `w:frame`, which asks for the line as a frame rather than a border around the text. */
  frame?: boolean;
}

/**
 * `w:docGrid` (§17.3.2.9), the line grid the section lays its text on.
 *
 * Measured over the authorized corpus (46 WPS-written files, `.scratch/corpus-recheck2.mjs` and
 * `.scratch/charspace.mjs`): **58 `w:sectPr` carry 58 `w:docGrid`, and every one of those states a
 * `w:linePitch`** — `w:type="lines"` 49 times, `w:type="linesAndChars"` 6, and 3 state no `w:type` at all, which
 * §17.3.2.9 defaults to `default`. The first section of 46 files takes its pitch from 312 缇 (41 files), 360 (3)
 * or 285 (2). `w:charSpace` says `0` 51 times, says nothing 3 times, and says `-2478` 4 times — the last four are
 * `plans and reports/business plan.dotm`, one pair per locale, and they are this corpus asking for a *character*
 * grid. So this model holds the two numbers that can move a line box (`type` and `linePitch`) and keeps no field
 * for the horizontal half, while the bytes of all three travel back unchanged through `sectionPropertiesXml`.
 *
 * A `linesAndChars` section still puts its text on a line grid — the character grid is the extra, not the whole
 * of what §17.3.2.9 spells — so `lineGridPitchTwips` counts it. `default` and `snapToColumns` (0 instances here)
 * and a stated `lines` that cannot state a pitch are no line grid; and a section that wrote no `w:docGrid` at all
 * asks for none, which is what every fixture this repository wrote before this round states: the page then draws
 * the line heights the file itself states, or the ones nobody states, exactly as it did before this field existed.
 */
export interface DocGrid {
  type: string;
  linePitch?: number;
}

export interface Section {
  pageWidthTwips: number;
  pageHeightTwips: number;
  /**
   * `w:pgMar`. The four body margins are what the page box draws its text column with. `header` and `footer`
   * are a different measurement: ECMA-376 §17.3.2.13 counts them from the **edge of the paper**, not from the
   * body, so a band's distance is only meaningful against the page margin it lives inside. `gutter` is extra
   * width on the binding side, added to the margin it belongs to. All three stay optional: WPS states all
   * seven attributes on all 58 `w:pgMar` in the authorized 46 份（`.scratch/pgmar.mjs`）, but a file that says nothing is not owed a
   * number this model invented.
   */
  margins: { top: number; right: number; bottom: number; left: number; header?: number; footer?: number; gutter?: number };
  /**
   * `w:pgBorders`, the line drawn around the page. Absent means the section states no page border at all, and that
   * is exactly what the authorized corpus says: **0 `w:pgBorders` elements among its 58 `w:sectPr`, 46 份**
   * (`.scratch/pgmar.mjs`). The model still reads and writes the element — §17.3.2.10 gives it its shape and
   * 「页面设置」 can be asked for a frame — but it has never had a real file to check that shape against.
   * `offsetFrom` says what the gap is measured against; a file that omits it lands on the default §17.3.2.10 gives
   * it — the page edge. (How WPS itself reads an omitted `w:offsetFrom` is 未取证; this is the spec's word, not an
   * observation.) A side that states no `w:val` is left out rather than drawn as a line of unknown shape.
   *
   * The reading this field used to cite — 24 of 115 sections carrying the element, all four sides each, and
   * **all 96 of those sides `w:val="none"`**, i.e. real WPS files that carry a border switched off rather than
   * drawn — came from the quarantined cohort and is **not** evidence on the authorized set. So a document with a
   * visible frame around the page is, here too, a shape that has never been observed: the reading and the drawing
   * of one are spec-derived, and their fidelity to WPS is 未取证.
   */
  pageBorders?: { offsetFrom: "page" | "text"; sides: Partial<Record<PageBorderSide, PageBorder>> };
  /** `w:docGrid`: the line grid every snapping paragraph in this section is measured in. See `DocGrid`. */
  docGrid?: DocGrid;
}

/**
 * The page a document is drawn on when its file spelled no section: A4 with a one-inch margin all round. This
 * has one owner because three places ask the same question — the view's fallback, a new document's first page,
 * and 「页面设置」 pre-filling a file that stated nothing — and a page the app shows and a save writes have to be
 * the same box.
 */
export const DEFAULT_SECTION: Section = { pageWidthTwips: 11906, pageHeightTwips: 16838, margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 } };

/**
 * 「这一节到底有没有一行网格、一格多高（缇）」唯一的一处判定。§17.3.2.9 给 `w:docGrid@w:type` 四种拼法，授权语料
 * 用其中三种：**58 枚 `w:sectPr` 里的 58 枚写了 `w:docGrid`，49 枚说 `lines`、6 枚说 `linesAndChars`、3 枚没说
 * `w:type`**（`DocGrid` 头记的是同一批数）。`linesAndChars` 算「有行网格」：那句字符格是**多**要的一格，不是行格的
 * 替代——把它当没有网格，就是让 `business plan.dotm` 与 `purchase contract.dotm` 这两份模板的行落在自己的间距上，
 * 而 WPS 的铺排照 312 缇（或 285）那一档走。`default`/`snapToColumns`（本批 0 枚）与整句没写一样，都算「没有行网格」。
 * 说了有网格却说不出的间距（`w:linePitch` 缺失或非正）也不算——那要模型替文件编一个数（本批 58 枚人人都说了）。
 *
 * 渲染（px）与「段落」对话框里那句小字（缇）都问这一个问题，所以判定只有这一份，两处各自只换算单位。
 */
export function lineGridPitchTwips(sec: Section): number | undefined {
  const g = sec.docGrid;
  if (!g || (g.type !== "lines" && g.type !== "linesAndChars")) return undefined;
  const pitch = g.linePitch;
  return pitch !== undefined && pitch > 0 ? pitch : undefined;
}

/** A copy, because a section a caller mutates must not move the page every other document falls back to. */
const copySection = (s: Section): Section => ({
  pageWidthTwips: s.pageWidthTwips,
  pageHeightTwips: s.pageHeightTwips,
  margins: { ...s.margins },
  // The sides copy by value, the lines inside them do not: a `PageBorder` is read once and never edited in
  // place, so giving each copy its own map is enough to keep a caller's add-or-drop of a side contained.
  pageBorders: s.pageBorders ? { offsetFrom: s.pageBorders.offsetFrom, sides: { ...s.pageBorders.sides } } : undefined,
  docGrid: s.docGrid ? { ...s.docGrid } : undefined,
});

/**
 * What 「页面设置」 changes on the section the page is drawn from: the seven `w:pgMar` distances and the paper's
 * direction. Distances left out are left alone — a file that never stated `w:header` keeps saying nothing about
 * it, because the number this model would write is a guess at the producer's default, not the user's intent.
 *
 * `landscape` is the direction the page is wanted in, and it is what decides which edge is which: the two
 * dimensions already on the section are re-assigned rather than replaced, so A4 stays A4 when it turns. 授权口径
 * 46 份把这一支的受众数成了 **0**：58 枚 `w:sectPr` 全带 `w:pgSz`，58 枚都把长边写在 `w:h`（11906×16838 的 56 枚、
 * 11907×16840 的 2 枚），**58 枚全不写 `w:orient`**——所以件里没有任何一枚横放的样本可参照（`.scratch/pagesz-theme-numfonts-authorized.py`）。
 * 「横放时两枚怎么写」因此 未取证，这一支只能按 §17.3.2.13 的「orient 说方向、w/h 说尺寸」来摆：turning back to
 * portrait removes the attribute, because writing `w:orient="portrait"` is a spelling no file here uses.
 */
export interface PageSetup {
  topTwips: number;
  rightTwips: number;
  bottomTwips: number;
  leftTwips: number;
  headerTwips?: number;
  footerTwips?: number;
  gutterTwips?: number;
  /**
   * The sheet of paper in twips, as the dialog's 宽/高 boxes state it, and always as a pair — one edge from
   * the new paper and the other from the old one would be a sheet no standard defines. Left out, the section
   * keeps the sheet it had and only its direction changes. `w:pgSz` has no presets in this model: 授权口径 46 份的
   * 58 枚 `w:sectPr` 各写一枚 `w:pgSz`，字面只有两种——`w:w="11906" w:h="16838"` 56 枚、`w:w="11907" w:h="16840"`
   * 2 枚（同一张 A4 差 1 缇的两种写法），属性顺序 58 枚全是 `w:w` 在前 `w:h` 在后。所以这里只收两个数、不给预设名：
   * 「A4」那一格属于对话框，不属于模型（`.scratch/pagesz-theme-numfonts-authorized.py`）。
   */
  paperTwips?: { width: number; height: number };
  landscape: boolean;
}

/** The page geometry a save owes, with the bytes it was read from — everything `setPageSetup` can move. */
export interface PageGeometry {
  bodySectPr: string;
  sections: Section[];
  /** True once the user set geometry on a document whose sections said nothing, so a save owes a `w:sectPr`. */
  stated: boolean;
}

/**
 * One paragraph of a header/footer part, taken apart in the order the file wrote it: the text this model
 * edits, with the fields it carries standing between the pieces. A field is not text — its bytes go back
 * unchanged and the number it last cached never joins what the user typed — so it needs an entry of its
 * own, and the paragraph needs an order to put that entry in.
 */
export interface HfSegment {
  text?: string;
  /** A page-number field kept as bytes. Same `key` job the carried inline plays in a body paragraph. */
  carry?: CarriedInline;
}

/**
 * Where one band paragraph says it sits on the paper. Read from the paragraph's own `w:pPr` — the same two
 * children a body paragraph is asked — and used only for drawing: the part travels back as its own bytes, so
 * nothing here is ever written.
 */
export interface BandPlace {
  /**
   * `w:jc`: this paragraph's own statement, or the one its `w:pStyle` chain lends beneath it. `undefined` means
   * nobody in the file said, and the page draws what its stylesheet says, as it did before.
   * 授权 46 份的眉脚部件共 76 段里，**12 段自己说 `center`**（6 个全路径件 / 3 份模板名），另有 2 段从样式链借到
   * `center`、48 段借到 `left`、2 段借到 `both`、12 段谁都没说（`.scratch/band-para-align-audience.py`，链在同包
   * `word/styles.xml` 里按 `w:pStyle → w:basedOn` 走）。此前带子里一枚排版属性都没贴，所以那 14 段居中全画在左沿。
   */
  align?: string;
  /**
   * `w:pPr/w:ind`: this paragraph's own statement over the one its `w:pStyle` chain lends, resolved by the same
   * referee the body paragraphs use (`indResolve`), so 段自己用任一枚答过这一件事时链上另一枚不许叠上来。
   *
   * 授权 46 份的眉脚部件里这一族有受众，两半各是量的（`.scratch/band-ind-tabs-audience.py` 与
   * `.scratch/band-place-slots-audience.py`，同 46 件口径、逐段配对）：**段自己说 4 段**，枚枚只说
   * `w:right="360"`（2 个全路径件 / 1 份模板名）；**链上借来 8 段**，说的是 `w:firstLine` 422 / 883 配
   * `w:firstLineChars="200"`（首行缩进 2 字符），链上 0 段借 left/right。此前带子里这一段与正文那一格
   * 用同一句话答的两枚都被丢掉：右缩进画到纸沿、首行缩进一格不缩。
   */
  ind?: IndentFields;
  /** `w:pPr/w:framePr` — the frame that paragraph is drawn as. See `FramePlace`. */
  frame?: FramePlace;
}

/**
 * The part of `w:framePr` (§17.4.2.5) this page can draw. A file puts a frame's box where its own sentence
 * names, not where the column flows, so a footer that says `w:hAnchor="margin" w:xAlign="right"` is the page
 * number at the **right** margin rather than at the left.
 *
 * Only the shapes the authorized 46 份 actually write are honoured. Every one of its 6 枚 `w:framePr`（全在页脚，
 * 4 个全路径件 / 2 份模板名）says exactly `w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:y="1"` plus
 * `w:xAlign` right ×4 / center ×2, and **0 枚 says `w:hRule`, `w:x`, `w:w`, `w:h`, `w:*Space*` or
 * `w:hAnchor="page"`**（`.scratch/layout-families-audience.py` 与 `.scratch/framepr-tblp-detail.py` 两趟同口径）。
 * So `hAnchor` gates the alignment: a box anchored to the **page** would swing outside the margin column this
 * band is drawn in, and nothing measured here says how far, so it is left undrawn rather than guessed.
 * `w:wrap` decides how body text flows around the box; a band has no body text to flow, so it is not read.
 */
export interface FramePlace {
  /** `w:hAnchor`: whose box the offsets count from. `margin` and `text` share this band's content column. */
  hAnchor?: string;
  /** `w:vAnchor`: named so a later reader can see which anchor the `y` below counts from. */
  vAnchor?: string;
  /** `w:xAlign`: the edge of the anchor box this frame is pushed to. */
  xAlign?: string;
  /** `w:yAlign`: same sentence, vertically. 0 枚 writes it; read so a file that does is not silently dropped. */
  yAlign?: string;
  /** `w:x`: offset from that edge, in twips. With `w:hRule` unmodelled this is read as the left-edge offset the attribute's own default (§17.4.2.5 `hRule="left"`) describes. */
  xTwips?: number;
  /** `w:y`: offset from the anchor, in twips. All 6 枚 say `1` — 0.07px, so this moves ink only where a file says more. */
  yTwips?: number;
}

/**
 * A header or footer part the model can show and (when simple) edit. Rich parts — tables, containers,
 * tracked changes — stay byte-identical on save and are not editable. Two shapes that look rich are read
 * through instead: a `PAGE` or `NUMPAGES` field sitting at run level travels as one atom with the text
 * around it still editable, and a floating text box whose inner paragraph holds the page number is read as
 * *that paragraph's own pieces*, so the words the page shows stay words. A picture with no text in it stays
 * one atom.
 */
export interface HeaderFooter {
  /** Package path ("word/header1.xml"). */
  part: string;
  /** Relationship id used by the body's sectPr reference. */
  id: string;
  /** True when the part is plain paragraphs only, so text edits round-trip. */
  editable: boolean;
  /** Plain text per paragraph, in order — the text segments of `segments`, joined. */
  paragraphs: string[];
  /** Texts as read, so a save knows whether the part needs regenerating. */
  baseline: string[];
  /**
   * The paragraphs with their fields, set on every part this model can edit. Paragraph i of this is
   * paragraph i of the bytes a save splices into, so the two count paragraphs the same way by
   * construction. Absent on a part held readonly, where nothing is ever written back per paragraph.
   */
  segments?: HfSegment[][];
  /** `segments` as read, so a save can tell a paragraph that lost a field from one that only gained a letter. */
  baselineSegments?: HfSegment[][];
  /**
   * Per paragraph: whether its pieces came out of a floating text box rather than out of the band itself.
   * A save writes those pieces back into every copy of the box the file kept, and leaves the frame alone.
   */
  boxedParagraphs?: boolean[];
  /**
   * Per paragraph, what it says about its own place on the band (`BandPlace`). Paragraph i lines up with
   * paragraph i of `source`, the same pairing `segments` keeps, so the array keeps a hole (`undefined`) where a
   * paragraph says nothing rather than shifting every later paragraph up a slot. A band with nothing stated
   * anywhere leaves `place` off entirely rather than carrying 76 empty objects.
   * Draw-only: a save splices text into the part's own bytes and never reads this array.
   */
  place?: (BandPlace | undefined)[];
  /**
   * The part as the file wrote it, `""` for one this session created. A save splices its edits into
   * these bytes rather than rebuilding the part, so the root's declarations and every paragraph the
   * user did not type in travel untouched.
   */
  source: string;
}

/**
 * What a save of this part would put down — each paragraph's text and kept bytes. The "did the model
 * change" hash has to read this rather than `paragraphs`, because a field's cached number is deliberately
 * not text the page shows: inserting a page number, or deleting one, moves nothing in the text view, and a
 * save that saw no change handed back the bytes it read with the field still in them.
 */
const hfSavedState = (hf: HeaderFooter | undefined): unknown =>
  hf?.segments?.map((segs) => segs.map((s) => s.carry?.xml ?? s.text ?? "")) ?? hf?.paragraphs ?? null;

export interface NumberLevel {
  numFmt: string;
  lvlText: string;
  start: number;
  ilvl?: number;
  indent?: number;
  /**
   * The font this level names for its own mark (`w:lvl/w:rPr/w:rFonts`). It is part of the glyph, not decoration.
   * 授权口径数到的是：**110 枚 `w:lvl` 里 84 枚自己写 `w:rFonts`**（`.scratch/authorized-families3.mjs` 同一条腿），
   * 所以「级自带的那张脸要跟着那颗符号走」这条规则有受众。作废语料上更细的那一层——170 枚 `bullet` 级把符号写成
   * 私有码位（F06E 44、F0A7 42…）、Wingdings 158 / Symbol 6、`w:sym` 0 次（`.scratch/sym-count.ts`）——在授权 46 份里
   * 被重新数过，结论换了：**`w:numFmt="bullet"` 级 0 枚**（值域只有 decimal 88、chineseCounting 14、lowerLetter 4、
   * lowerRoman 2、decimalEnclosedCircleChinese 2），**`w:sym` 在任何 xml 部件里都 0 枚**（同一趟扫描的对照 `w:u` 找到
   * 1,615 枚，所以这个 0 是「真没有」而不是「没找到」）。也就是：PUA 码位那套读法在授权语料上没有实例，`w:sym` 那条
   * 腿依旧只能按 §17.3.1.12 的规范形状读，不作「实测过」。
   */
  font?: string;
  /**
   * The size this level names for its own mark, in points (`w:lvl/w:rPr/w:sz` ÷ 2 — the file counts half-points,
   * the same unit `readRunStyle` converts for body runs). §17.3.1.12 gives the symbol a character run of its own,
   * so this is the mark's size and not the paragraph's. 授权 46 份的受众是 **110 枚级里 2 枚自己写 `w:sz`**——比作废
   * 语料上那个「70 段坐在写了字号的级上、其中 42 枚 `bullet` 比正文小（42/42）」的读数小一个数量级，那套分法（含
   * 5 个不同值、28 枚 `decimal` 与正文解析出的磅数相等）没在授权口径上重做，属 **待重证**。规则本身不靠那个分母：
   * 级写了就按级的画。同一条 `w:rPr` 的其余孩子仍按授权口径留着不建模——`w:b`/`w:i`/`w:u`/`w:color`/`w:caps`/`w:strike`
   * 在 110 枚级上的分布没有重新数过（作废语料上「8 枚加粗级与段落自身一致」那一句是 待重证）。
   */
  sizePt?: number;
}

/**
 * The definition `ensureListNumbering` writes when a file names neither shape. Both marks are ordinary characters,
 * not private-use codepoints: the bullets the quarantined cohort wrote are PUA glyphs whose font the level states,
 * and an invented definition carries no such font, so a codepoint would draw tofu. On the authorized 46 份 that
 * shape has no instance at all — 110 枚级里 `w:numFmt="bullet"` 0 枚——所以这条选择不靠那批作废件说话，只按「补出来的
 * 定义没有自带脸」这条理由立着。`%1.` 是授权口径上的多数写法：88 枚 `decimal` 级里 66 枚在 `%N` 之后跟一个 `.`
 * （另有 `、` 10、`）` 6、`．` 2、`)` 2、什么都不跟 2，`.scratch/authorized-families3.mjs`）。
 */
const LIST_SHAPE: Record<"bullet" | "number", { numFmt: string; lvlText: string }> = {
  bullet: { numFmt: "bullet", lvlText: "•" },
  number: { numFmt: "decimal", lvlText: "%1." },
};

/** One list definition a save owes `word/numbering.xml`, with the ids it was handed. */
export interface ListAddition {
  kind: "bullet" | "number";
  numId: number;
  abstractId: number;
  lvlText: string;
  numFmt: string;
}

const TWIPS_PER_PX = 15;
export const twipsToPx = (tw: number): number => tw / TWIPS_PER_PX;

/**
 * 「页面设置」上的单位是厘米，而 `w:pgSz`/`w:pgMar` 以缇计（ECMA-376 §17.3.2.13），一条 = 1440/2.54 缇。
 * 换算只在输入框与模型之间做一次：模型和字节始终存缇，所以一次改动的往返不会因为反复取整而漂。
 */
const TWIPS_PER_CM = 1440 / 2.54;
export const twipsToCm = (tw: number): number => tw / TWIPS_PER_CM;
export const cmToTwips = (cm: number): number => Math.round(cm * TWIPS_PER_CM);

const REL_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const XML_ESC = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * `xml:space="preserve"` goes on a text node only where the text would otherwise lose something at its
 * edges, which is the rule the files follow. The counts that showed WPS following it — the attribute on 960 of
 * the 5,382 `w:t` in the local body corpus, and on exactly the 17 of its 70 header/footer `w:t` that open or close
 * with whitespace, never on the other 53 — came from a pass whose denominator the authorized 46 份 re-measure
 * never covered: 待重证. Spelling it on every one is not a wrong document, it is a
 * different one, and "same bytes when nothing moved" is what a save is measured against here.
 */
export const preserve = (text: string): string => (/^\s|\s$/.test(text) ? ' xml:space="preserve"' : "");

/**
 * What the page shows for one header/footer paragraph: every `w:t` of it in source order, with each
 * literal `w:tab` put as the "\t" the body model gives a tab too. `w:pPr` is stepped over — the `<w:tab>`
 * children of its `w:tabs` define where a tab lands, they are not text the reader sees.
 */
function hfParagraphText(span: { xml: string }): string {
  let text = "";
  for (const kid of childSpans(elementContent(span.xml))) {
    if (kid.tag === "pPr") continue;
    for (const m of kid.xml.matchAll(/<(?:[\w.-]+:)?t\b([^>]*)>([^<]*)<\/(?:[\w.-]+:)?t>|<(?:[\w.-]+:)?tab[ />]/g)) {
      text += m[1] === undefined ? "\t" : decodeEntities(m[2] ?? "");
    }
  }
  return text;
}

/**
 * The page fields this model is willing to keep: `PAGE` and `NUMPAGES`, with or without the formatting
 * switch. The shape is the one the authorized 46 份 actually wrote: its 38 band fields spell `PAGE` 36 times and
 * `NUMPAGES` twice, 28 of the 38 with the ` \* MERGEFORMAT ` switch and 10 without, 34 with the instruction
 * padded by spaces and 4 bare (`.scratch/band-shape-authorized.py`). Nothing else appears at run level in a
 * band, so this regular expression covers all 38 — but the guard is not free: a field this reader does not
 * recognize holds the whole part as its own bytes, and the 46 份 simply never test that branch.
 */
const PAGE_FIELD = /^\s*(PAGE|NUMPAGES)(\s+\\\*\s+MERGEFORMAT)?\s*$/;

/**
 * Content beside a header/footer's paragraphs that a flat piece-by-piece edit cannot account for: a table,
 * a picture, a container, a tracked change, an anchored comment. The prefixes are matched loosely on purpose
 * — a part that binds the wordprocessingml namespace under another name is still rich, and reading only
 * `w:` is how a container slipped through and got dropped.
 */
const HF_BLOCKER = /<(?:[\w.-]+:)?(tbl|drawing|pict|object|fldSimple|fldChar|hyperlink|sdt|ins|del|br|sym)[ />]|<(?:[\w.-]+:)?comment(?:Range|Reference)/;

/** What a paragraph may hold besides runs, and still be a paragraph this model can write back. */
const HF_PARA_STRUCTURAL = new Set(["pPr", "bookmarkStart", "bookmarkEnd"]);

/** What a run of one of those paragraphs may hold: its properties, text, a tab, and a field's two nodes. */
const HF_RUN_KIDS = new Set(["rPr", "t", "tab", "fldChar", "instrText"]);

/**
 * A run that holds one of these instead of text is not a paragraph this model has to refuse. Either it is a
 * picture, and it travels whole as one atom; or it is a floating text box, and its inner paragraph's pieces
 * join the band's. The pass that measured it — every one of the 16 runs of this shape being a `w:rPr` plus an
 * `mc:AlternateContent` wrapping an anchored box, every one sitting alone in its paragraph — was made on the local
 * corpus, and this shape has not been re-counted on the authorized 46 份: 待重证.
 */
const HF_RUN_CARRIED = new Set(["AlternateContent", "drawing", "pict", "object"]);

/**
 * What no run may hide behind a chip: a revision, a hyperlink, a container, a comment anchor, or a break
 * the page shows as text. Each of those has its own meaning the flat read would swallow, and a band that
 * holds one stays held as its own bytes rather than half-explained.
 */
const HF_RUN_DENY = /<(?:[\w.-]+:)?(ins|del|hyperlink|sdt|fldSimple|br|sym)[ />]|<(?:[\w.-]+:)?comment(?:Range|Reference)/;

/**
 * Every element named `tag` at any depth under `xml`, in source order. A floating box hides its words twice:
 * the 16 boxes the quarantined cohort measured each carry a `w:txbxContent` inside the Choice's drawing *and*
 * another inside the Fallback's VML shape, both holding the same one paragraph — the pass never counted that
 * shape on the authorized 46 份, so it is 待重证 there. Reading a box means finding both copies; writing one means
 * updating both, or the page tells two different stories depending on which branch a reader takes.
 */
function spansDeep(xml: string, tag: string, found: { xml: string }[] = []): { xml: string }[] {
  for (const kid of childSpans(elementContent(xml))) {
    if (kid.tag === tag) found.push(kid);
    else spansDeep(kid.xml, tag, found);
  }
  return found;
}

/** One literal substring swapped for another. `String.replace` would read `$` in the replacement as markup. */
const replaceOnce = (xml: string, from: string, to: string): string => {
  const at = xml.indexOf(from);
  return at < 0 ? xml : xml.slice(0, at) + to + xml.slice(at + from.length);
};

/** What a box's pieces say, in the only form two copies of the same box can be compared in: text and atoms. */
const hfBoxShape = (segs: HfSegment[]): string => JSON.stringify(segs.map((s) => (s.carry ? `<${s.carry.tag}>` : s.text ?? "")));

/** Whether an element's children account for every byte between its tags, so splicing them back says the same thing. */
const losslessKids = (xml: string): boolean => {
  const content = elementContent(xml);
  return childSpans(content).map((k) => k.xml).join("") === content;
};

/**
 * What a content control's properties may say for the read-through to be honest. Every one of the 3 measured
 * declares `w:id`, and one a `w:docPartObj` (a gallery building block); `w:alias`/`w:tag` only name it. A
 * `w:lock` or `w:placeholder` is a different promise — WPS would refuse the edit too — and `w:dataBinding`
 * means the text is not typing but a field some other store fills, so a control holding either stays held.
 */
const HF_SDT_PR = new Set(["id", "alias", "tag", "docPartObj"]);

/**
 * A header/footer part's children, walked in the order the page shows them, with `put` asked for every
 * paragraph it holds — whether that paragraph sits at the root or inside a content control.
 *
 * WPS wraps a footer's page number in `w:sdt`: the authorized 46 份 put 4 of them in a band part, arranged as two
 * pairs nested two deep, each pair holding two fields — one counting the page, one counting pages — and typed
 * characters beside them (`.scratch/band-detail-authorized.py`, `.scratch/field-stage-shapes.ts`). The control is
 * the file's business, so it travels as its own bytes and only the paragraph inside it is re-spelled. Reading
 * through `sectionBand` over that cohort, all 32 bands it answers for come back editable and none is held, so the
 * nesting is walked, not dodged. Two pairs is as deep as that cohort goes, so the third level named below is
 * past anything measured and nothing exercises it.
 *
 * `null` means this band is not a band this model may claim: a table or a revision at the root, a control
 * that declares no properties or properties beyond a name, one nested past the third level, or bytes between
 * a container's children that the walk would drop. The reader calls this first and only keeps the part
 * editable when it gives back a walk, so the save never meets a `null` the read did not already refuse.
 */
function bandParagraphs(
  kids: { tag: string; xml: string }[],
  put: (span: { xml: string }) => string,
  depth = 0,
): string | null {
  let out = "";
  for (const k of kids) {
    if (k.tag === "p") {
      out += put(k);
      continue;
    }
    if (k.tag === "sdt") {
      if (depth >= 3 || !losslessKids(k.xml)) return null;
      const parts = childSpans(elementContent(k.xml));
      const pr = parts.find((s) => s.tag === "sdtPr");
      const content = parts.find((s) => s.tag === "sdtContent");
      if (!pr || !content || !parts.every((s) => s === pr || s === content)) return null;
      if (!childSpans(elementContent(pr.xml)).every((s) => HF_SDT_PR.has(s.tag))) return null;
      if (!losslessKids(content.xml)) return null;
      const body = bandParagraphs(childSpans(elementContent(content.xml)), put, depth + 1);
      if (body === null) return null;
      out += replaceOnce(k.xml, content.xml, replaceOnce(content.xml, elementContent(content.xml), body));
      continue;
    }
    if (HF_BLOCKER.test(k.xml)) return null;
    out += k.xml;
  }
  return out;
}

/**
 * The paragraphs a band the walk refuses still has to show: its own, plus the ones a content control holds,
 * however deep. Refusing the edit says nothing about the words — a footer whose page number sits in a locked
 * control is not an empty footer, and drawing it empty would tell the user their page lost its text. The
 * bytes are a separate question: the part travels as the file wrote it, so this list only ever feeds the
 * text view. Capped at eight levels because nesting is the file's choice and the page is not a stack to lose.
 */
function bandTextSpans(kids: { tag: string; xml: string }[], depth = 0): { xml: string }[] {
  const found: { xml: string }[] = [];
  for (const k of kids) {
    if (k.tag === "p") found.push(k);
    else if (k.tag === "sdt" && depth < 8) {
      const content = childSpans(elementContent(k.xml)).find((s) => s.tag === "sdtContent");
      if (content) found.push(...bandTextSpans(childSpans(elementContent(content.xml)), depth + 1));
    }
  }
  return found;
}

/** Which marker of a field a run holds — `begin`, `separate`, `end` — by attribute, whatever prefix spells it. */
const fldStageOf = (run: string): string | undefined => {
  const m = /<(?:[\w.-]+:)?fldChar\b([^>]*)>/.exec(run);
  return m ? /(?:^|\s)(?:[\w.-]+:)?fldCharType="(\w+)"/.exec(m[1] ?? "")?.[1] : undefined;
};

/** The text a run shows: its `w:t` nodes in order, each literal `w:tab` as the "\t" the body model uses. */
const hfRunText = (run: string): string =>
  [...run.matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([^<]*)<\/(?:[\w.-]+:)?t>|<(?:[\w.-]+:)?tab[ />]/g)]
    .map((m) => (m[1] === undefined ? "\t" : decodeEntities(m[1] ?? "")))
    .join("");

/**
 * One paragraph of a header/footer part, taken apart in the order the file wrote it: text where the file put
 * text, one carried atom per page field, and `ok: false` for anything else — a hyperlink, a picture, a
 * container, a nested or unpaired field, an instruction this model cannot spell back.
 *
 * Reading a field this way instead of matching it on the part as a whole is what lets `PAGE` stop being a
 * reason to refuse the entire footer. The number the field caches between its `separate` and its `end` is
 * not text the user typed, so it never joins the paragraph's text view; it travels inside the atom.
 */
function hfParagraphSegments(
  span: { xml: string },
  newKey: () => number,
  depth = 0,
): { segs: HfSegment[]; ok: boolean; boxed: boolean } {
  const kids = childSpans(elementContent(span.xml));
  const segs: HfSegment[] = [];
  let text = "";
  let boxed = false;
  let i = 0;
  const rich = (): { segs: HfSegment[]; ok: boolean; boxed: boolean } => ({ segs: [], ok: false, boxed: false });
  while (i < kids.length) {
    const kid = kids[i]!;
    if (HF_PARA_STRUCTURAL.has(kid.tag)) {
      i++;
      continue;
    }
    if (kid.tag !== "r") return rich();
    const runKids = childSpans(elementContent(kid.xml));
    if (!runKids.every((k) => HF_RUN_KIDS.has(k.tag))) {
      if (!runKids.every((k) => k.tag === "rPr" || HF_RUN_CARRIED.has(k.tag)) || HF_RUN_DENY.test(kid.xml)) return rich();
      // A run that is one picture and nothing else travels whole, the way a page field does; its `w:rPr`
      // rides along with it and is not text. But first: is this picture a box with words in it?
      const boxes = spansDeep(kid.xml, "txbxContent");
      if (boxes.length) {
        // The band's page number sits in an anchored text box in 12 of the authorized 46 份: 24 `w:txbxContent`
        // blocks, exactly two in each of those 12 parts, and each of those parts holds exactly one
        // `mc:AlternateContent` with one `mc:Choice` and one `mc:Fallback` — so the two blocks are the two copies of
        // one box, counted rather than assumed (`.scratch/band-detail-authorized.py`). All 24 carry the field's
        // instruction, and `sectionBand` answers 12 editable bands of this shape. The box is not a label over the
        // user's words — it *is* the paragraph, so its pieces join this paragraph's pieces, and the frame
        // stays where the file put it. Bounded on purpose: one box-shaped run, the box the only run of its
        // paragraph, one paragraph inside, nothing but the box's own copy holding text, and every copy of
        // the box saying the same thing. Anything outside that shape keeps the band held as its own bytes.
        if (
          depth > 0 ||
          kids.filter((k) => k.tag === "r").length !== 1 ||
          runKids.filter((k) => HF_RUN_CARRIED.has(k.tag)).length !== 1
        )
          return rich();
        let outside = kid.xml;
        for (const b of boxes) outside = replaceOnce(outside, b.xml, "");
        if (hfRunText(outside) !== "") return rich();
        const inner: HfSegment[][] = [];
        for (const b of boxes) {
          const paras = childSpans(elementContent(b.xml)).filter((s) => s.tag === "p");
          if (paras.length !== 1) return rich();
          const parsed = hfParagraphSegments(paras[0]!, newKey, depth + 1);
          if (!parsed.ok || parsed.boxed) return rich();
          inner.push(parsed.segs);
        }
        const shape = hfBoxShape(inner[0]!);
        if (inner.some((s) => hfBoxShape(s) !== shape)) return rich();
        if (text !== "") segs.push({ text });
        text = "";
        segs.push(...inner[0]!);
        boxed = true;
        i++;
        continue;
      }
      // No box, and words inside the picture: a chip would replace text the page shows with a label and hand
      // back a band that reads "editable" with nothing to edit. Such a part stays held as its own bytes.
      if (hfRunText(kid.xml) !== "") return rich();
      if (text !== "") segs.push({ text });
      text = "";
      segs.push({ carry: { key: newKey(), tag: runKids.find((k) => HF_RUN_CARRIED.has(k.tag))!.tag, xml: kid.xml, show: "图片" } });
      i++;
      continue;
    }
    if (fldStageOf(kid.xml) !== "begin") {
      text += hfRunText(kid.xml);
      i++;
      continue;
    }
    let end = -1;
    let instr = "";
    for (let j = i + 1; j < kids.length; j++) {
      const k = kids[j]!;
      if (k.tag !== "r" || !childSpans(elementContent(k.xml)).every((c) => HF_RUN_KIDS.has(c.tag))) break;
      const stage = fldStageOf(k.xml);
      if (stage === "begin") break;
      if (stage === "end") {
        end = j;
        break;
      }
      const m = /<(?:[\w.-]+:)?instrText\b[^>]*>([^<]*)<\/(?:[\w.-]+:)?instrText>/.exec(k.xml);
      if (m) instr += decodeEntities(m[1] ?? "");
    }
    if (end < 0 || !PAGE_FIELD.test(instr)) return rich();
    if (text !== "") segs.push({ text });
    text = "";
    segs.push({
      carry: {
        key: newKey(),
        tag: "fldChar",
        xml: kids.slice(i, end + 1).map((k) => k.xml).join(""),
        show: instr.includes("NUMPAGES") ? "总页数" : "页码",
        field: instr.includes("NUMPAGES") ? "NUMPAGES" : "PAGE",
      },
    });
    i = end + 1;
  }
  if (text !== "") segs.push({ text });
  return { segs, ok: true, boxed };
}

/** The paragraph's text view: what the pieces of it that are text say, with a field contributing nothing. */
const hfSegmentsText = (segs: HfSegment[]): string => segs.map((s) => s.text ?? "").join("");

/** Text that sits side by side is one run's worth: what the page shows as one stretch of typing. */
const coalesceHfSegments = (segs: HfSegment[]): HfSegment[] => {
  const out: HfSegment[] = [];
  for (const s of segs) {
    if (s.carry) {
      out.push({ carry: s.carry });
      continue;
    }
    const t = s.text ?? "";
    if (t === "") continue;
    const prev = out.at(-1);
    if (prev?.text !== undefined) prev.text += t;
    else out.push({ text: t });
  }
  return out;
};

/** A copy of a paragraph's pieces, so what a save compares against cannot move under it. */
const copyHfSegments = (segs: HfSegment[]): HfSegment[] => segs.map((s) => (s.carry ? { carry: { ...s.carry } } : { text: s.text }));

/** Whether a paragraph still says what the file wrote: same pieces, same order, same bytes inside each. */
const sameHfSegments = (a: (HfSegment[] | undefined)[], b: (HfSegment[] | undefined)[]): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

/** Paragraph i's pieces, or the flat texts for a part that was never taken apart. */
const hfSegmentsAt = (hf: HeaderFooter, i: number): HfSegment[] | undefined =>
  hf.segments?.[i] ?? (hf.paragraphs[i] === undefined ? undefined : [{ text: hf.paragraphs[i]! }]);

/**
 * A header/footer part after an edit, spliced into the bytes it came in with rather than rebuilt from
 * the texts: the root tag keeps every declaration the file wrote — the authorized 46 份 put 646 `xmlns`
 * declarations on their 38 band roots, exactly 17 each, and every one of the 38 also carries `mc:Ignorable`
 * (`.scratch/band-shape-authorized.py`) — a paragraph whose text did not change travels
 * byte for byte, and a paragraph that did keeps its own `w:pPr` and its first run's `w:rPr`. Rebuilding
 * the part from `paragraphs` instead — which is what this did — takes the style, the alignment, the font
 * size and the namespace declarations off every page the user typed in. The audience for that is the same
 * cohort's bands: 74 of the 76 paragraphs in those 38 parts write a `w:pPr`, 62 of them naming a style and 12
 * stating an alignment — slots a rebuilt part would have to re-invent.
 */
function headerFooterXml(hf: HeaderFooter): string {
  const root = hf.source ? childSpans(hf.source).find((s) => s.tag === "hdr" || s.tag === "ftr") : undefined;
  if (!root) return spelledHeaderFooter(hf);
  const open = root.xml.slice(0, root.xml.indexOf(">") + 1);
  const close = /<\/[^>]*>$/.exec(root.xml)?.[0] ?? "";
  const prologue = /^\s*<\?xml[^>]*\?>\s*/.exec(hf.source)?.[0] ?? "";
  const seen = childSpans(elementContent(root.xml));
  let at = -1;
  const put = (span: { xml: string }): string => {
    const segs = hfSegmentsAt(hf, ++at);
    if (segs === undefined) return "";
    return sameHfSegments([segs], [hf.baselineSegments?.[at]])
      ? span.xml
      : hf.boxedParagraphs?.[at]
        ? editedBoxParagraph(span, segs, hf.baselineSegments?.[at])
        : editedParagraph(span, segs, hf.baselineSegments?.[at]);
  };
  // The same walk the read took, so paragraph i of the model is paragraph i of the bytes whether the file put
  // it at the root or inside a content control. A band that got here was walked by the read; holding its bytes
  // beats re-spelling it by a rule the read did not check.
  const body = bandParagraphs(seen, put);
  if (body === null) return hf.source;
  const paras = at + 1;
  const extra = hf.paragraphs
    .slice(paras)
    .map((_, i) => coalesceHfSegments(hfSegmentsAt(hf, paras + i) ?? []))
    .map((segs) => spelledParagraph(segs, "w"))
    .join("");
  return `${prologue}${open}${body}${extra}${close}`;
}

/** A paragraph the model holds pieces for, rewritten in place: properties kept, its text re-spelled, its fields carried. */
function editedParagraph(span: { xml: string }, segs: HfSegment[], asRead?: HfSegment[]): string {
  const prefix = /^<([\w.-]+):/.exec(span.xml)?.[1] ?? "w";
  const attrs = (/^<[\w.-]*:?p\b([^>]*)>/.exec(span.xml)?.[1] ?? "").replace(/\/\s*$/, "");
  const kids = childSpans(elementContent(span.xml));
  // Properties and empty markers come back as written; run content does not, because the text the model
  // holds is every `w:t` of the paragraph merged into one string, and keeping the runs too would say it
  // twice. A paragraph with several differently formatted runs is therefore set in the first one's.
  const kept = kids.filter((k) => k.tag === "pPr" || /^bookmark(Start|End)$/.test(k.tag)).map((k) => k.xml).join("");
  // The paragraph's own typeface comes from a run the model did not carry away as an atom. A page field's
  // runs — its instruction and the cached digit — are the *field's* formatting, and setting a sentence in
  // the digit's size would do that to every paragraph that opens with a number. Only when the whole
  // paragraph is one field is there no run of its own to take, and then the digit's size is the only one
  // this paragraph was ever given, so typed text beside it wears that rather than the document default.
  const runs = kids.filter((k) => k.tag === "r");
  const atoms = (asRead ?? []).filter((s) => s.carry).map((s) => s.carry!.xml);
  const first =
    runs.find((k) => !atoms.some((a) => a.includes(k.xml))) ?? runs.find((k) => /<(?:[\w.-]+:)?t\b/.test(k.xml)) ?? runs[0];
  const rPrefix = /^<([\w.-]+):/.exec(first?.xml ?? "")?.[1] ?? prefix;
  const props = first ? childSpans(elementContent(first.xml)).find((k) => k.tag === "rPr")?.xml ?? "" : "";
  const parts = coalesceHfSegments(segs)
    .map((s) =>
      s.carry
        ? s.carry.xml
        : s.text && s.text !== ""
          ? `<${rPrefix}:r>${props}${textXml(s.text, rPrefix)}</${rPrefix}:r>`
          : "",
    )
    .join("");
  return `<${prefix}:p${attrs}>${kept}${parts}</${prefix}:p>`;
}

/**
 * A paragraph whose pieces came out of a floating text box, written back **into the box**: the frame, its
 * geometry and its properties travel as the file wrote them, and only the paragraph inside each copy of
 * `w:txbxContent` is re-spelled. Both copies — the Choice's drawing and the Fallback's VML shape — are
 * rebuilt from the same pieces, because a page that read one text in the branch it took and another in the
 * branch it did not is two different documents.
 */
function editedBoxParagraph(span: { xml: string }, segs: HfSegment[], asRead?: HfSegment[]): string {
  const kids = childSpans(elementContent(span.xml));
  const run = kids.find((k) => k.tag === "r" && spansDeep(k.xml, "txbxContent").length > 0);
  if (!run) return editedParagraph(span, segs, asRead);
  let runXml = run.xml;
  for (const box of spansDeep(run.xml, "txbxContent")) {
    const para = childSpans(elementContent(box.xml)).find((s) => s.tag === "p");
    if (!para) continue;
    runXml = replaceOnce(runXml, box.xml, replaceOnce(box.xml, para.xml, editedParagraph(para, segs, asRead)));
  }
  const prefix = /^<([\w.-]+):/.exec(span.xml)?.[1] ?? "w";
  const attrs = (/^<[\w.-]*:?p\b([^>]*)>/.exec(span.xml)?.[1] ?? "").replace(/\/\s*$/, "");
  const kept = kids.filter((k) => k.tag === "pPr" || /^bookmark(Start|End)$/.test(k.tag)).map((k) => k.xml).join("");
  return `<${prefix}:p${attrs}>${kept}${runXml}</${prefix}:p>`;
}

/**
 * The text as markup: a tab the model holds as "\t" goes back as the `w:tab` the file used for it, so a
 * footer that splits its text across tab stops keeps them through a text edit instead of losing them.
 */
const textXml = (text: string, prefix: string): string =>
  text
    .split("\t")
    .map((piece) => (piece === "" ? "" : `<${prefix}:t${preserve(piece)}>${XML_ESC(piece)}</${prefix}:t>`))
    .join(`<${prefix}:tab/>`);

/** A paragraph the part did not have: one run per piece of text, fields carried as the bytes they came in. */
const spelledParagraph = (segs: HfSegment[], prefix: string): string => {
  const body = coalesceHfSegments(segs)
    .map((s) => (s.carry ? s.carry.xml : s.text && s.text !== "" ? `<${prefix}:r>${textXml(s.text, prefix)}</${prefix}:r>` : ""))
    .join("");
  return body === "" ? `<${prefix}:p/>` : `<${prefix}:p>${body}</${prefix}:p>`;
};

/** A part this session created, with no bytes to splice into. */
function spelledHeaderFooter(hf: HeaderFooter): string {
  const root = hf.part.includes("footer") ? "ftr" : "hdr";
  const segs = hf.segments ?? hf.paragraphs.map((t) => [{ text: t }]);
  const body = segs.length ? segs.map((s) => spelledParagraph(s, "w")).join("") : "<w:p/>";
  return `${DECL}<w:${root} xmlns:w="${NS_W}" xmlns:r="${REL_R}">${body}</w:${root}>`;
}

/**
 * A page-number field the way the files this machine was measured against write one: the instruction inside
 * `w:instrText` with its edges kept by `xml:space="preserve"`, `begin` and `end` marking the field off, a
 * `separate` before the number it last showed. The parts that spelled it this way — 16 of the 21 run-level fields
 * in the quarantined cohort's header and footer parts, all 21 caching a number — have not been re-counted on the
 * authorized 46 份: 待重证, so the caching below is the shape §17.6.4 allows rather than a measured idiom — which is why this one
 * caches `1` too: a reader that lays out pages recomputes it, and one that does not still has a digit to
 * put where the page number belongs. Our own page shows neither digit, but the field's label, because this
 * product has no pagination to compute a page from yet.
 */
const pageFieldXml = (instruction: string, prefix: string): string =>
  `<${prefix}:r><${prefix}:fldChar ${prefix}:fldCharType="begin"/></${prefix}:r>` +
  `<${prefix}:r><${prefix}:instrText xml:space="preserve"> ${instruction}  \\* MERGEFORMAT </${prefix}:instrText></${prefix}:r>` +
  `<${prefix}:r><${prefix}:fldChar ${prefix}:fldCharType="separate"/></${prefix}:r>` +
  `<${prefix}:r><${prefix}:t>1</${prefix}:t></${prefix}:r>` +
  `<${prefix}:r><${prefix}:fldChar ${prefix}:fldCharType="end"/></${prefix}:r>`;

/** Where a field goes among a paragraph's pieces: at `offset` characters of text, or at the end. */
const putHfCarry = (segs: HfSegment[], offset: number | undefined, carry: CarriedInline): HfSegment[] => {
  if (offset === undefined) return coalesceHfSegments([...segs, { carry }]);
  const out: HfSegment[] = [];
  let seen = 0;
  let placed = false;
  for (const s of segs) {
    const t = s.text;
    if (!placed && t !== undefined && offset <= seen + t.length) {
      placed = true;
      const cut = offset - seen;
      if (t.slice(0, cut) !== "") out.push({ text: t.slice(0, cut) });
      out.push({ carry });
      if (t.slice(cut) !== "") out.push({ text: t.slice(cut) });
    } else {
      out.push(s);
    }
    if (t !== undefined) seen += t.length;
  }
  if (!placed) out.push({ carry });
  return coalesceHfSegments(out);
};

/** A comment body is one `w:p` per line of the editor's text. */
const commentParagraphs = (text: string): Inline[][] => text.split("\n").map((line) => [{ text: line }]);

/** `w:date` is an `xsd:dateTime`; Word writes second precision, so match it. */
/** Word's `w:date` shape: UTC to the second. */
export const stampNow = (): string => new Date().toISOString().replace(/\.\d{3}Z$/, "Z");

/**
 * One `w:style` as this model keeps it: the properties it states itself (`pPr`, `rPr`), the style whose
 * unstated properties it takes over (`basedOn`), the style a paragraph takes when this one is split by Enter
 * (`next`), and the name the UI shows.
 */
export interface StyleDef {
  name: string;
  /** `w:type` (§17.7.2.3): absent means paragraph, the only kind the 样式 gallery offers. */
  type?: string;
  basedOn?: string;
  next?: string;
  rPr?: RunStyle;
  pPr?: Partial<ParaStyle>;
}

/**
 * What one section's own `w:sectPr` cites for the two bands — nothing inherited from anybody. Reading it
 * per section is what lets a page in a later section show *that* section's header; which band a section's
 * page shows is `WordDocument.sectionBand`, and a save writes every part some section shows (`headerFooterParts`).
 */
export interface SectionBands {
  header?: HeaderFooter;
  footer?: HeaderFooter;
}

export class WordDocument {
  blocks: Block[] = [];
  sections: Section[] = [];
  header?: HeaderFooter;
  footer?: HeaderFooter;
  /** 每一节自己 cite 的那两份带子，与 `sections` 同序同长；只有读端填，画端只读。 */
  sectionRefs: SectionBands[] = [];
  /** Bands read so far, by package path: one part cited by several sections stays one object. */
  private bandByPart = new Map<string, HeaderFooter>();
  comments: DocComment[] = [];
  private styles = new Map<string, StyleDef>();
  /** The `w:default="1"` paragraph style: what a paragraph that names no style falls back on (§17.7.2). */
  private defaultParaStyleId?: string;
  /**
   * `w:docDefaults/w:rPrDefault`, kept to the properties the ladder lends: the rung every paragraph chain ends
   * on, §17.7.2. Only one rung, and it goes in below the chain rather than inside it — see `paragraph`.
   */
  private docDefaultRun: RunStyle = {};
  /**
   * The package's theme fonts, read once at the top of `read` because §17.3.2's four `w:rFonts` theme attributes are
   * how a `w:rPr` can name a face without naming one. The audience that first made that worth reading — 58 of the
   * 201 WPS-written documents putting their whole document floor there (`.scratch/theme-font-audience.ts`) — is
   * 待重证 on the authorized 46 份, whose theme-slot census (`.scratch/theme-slot-pair-authorized.mjs`) counts the
   * themed `w:rFonts` themselves but never this styles-floor split. See `ThemeFonts` for what resolves and what
   * deliberately does not.
   */
  private theme: ThemeFonts = ThemeFonts.none;
  /** Resolved style chains, keyed by the style id the paragraph names. See `paraLent`. */
  private lentMemo = new Map<string | undefined, Partial<ParaStyle> | undefined>();
  /** The same memo for `runLent`, keyed by which kind of chain was walked as well as by the style. */
  private runLentMemo = new Map<string, Partial<RunStyle> | undefined>();
  /** Resolved `w:next` per style id, the way `nextStyleOf` walks it. See §17.7.2.24. */
  private nextMemo = new Map<string, string | undefined>();
  private numbering = new Map<number, NumberLevel[]>();
  /**
   * The list definitions this document *did not come with*, keyed by the `w:numId` they were given. An id is only
   * ever handed out here, so a paragraph that cites one of these is asking for a definition the file has to grow.
   */
  private inventedNumbers = new Map<number, { kind: "bullet" | "number"; abstractId: number }>();
  private media = new Map<string, string>();
  private pkg?: OpcPackage;
  private hyperlinks = new Map<string, string>();
  private bodyRels: { id: string; type: string; target: string; targetMode: string; path: string }[] = [];
  private bodySectPr = "";
  /**
   * Set once the user asked for page geometry on a document whose file spelled no section at all. It decides
   * whether a save owes a `w:sectPr`: writing one that the file never had is what this model refuses on a plain
   * save, and what a page-setup change explicitly asks for.
   */
  private pageSetupSeen = false;
  /** Namespace declarations read off the document root, so carried markup keeps its prefixes bound. */
  private rootNamespaces: Record<string, string> = {};
  private rootIgnorable = "";
  /** The document root's own attribute text, read as bytes: a save writes this rather than spelling a root. */
  private rootOpenAttrs = "";
  private sectPrefix = "w";
  private sectHas = { header: false, footer: false };
  /**
   * Whether the section that starts the document is the body's own `sectPr` or one hanging off a paragraph.
   * A header belongs to the section whose page it sits on, and the page the editor draws is the first one,
   * so a citation the model still owes goes into that `sectPr` — not into the body's, which is the document's
   * **last** section and would hand the closing pages a header the opening page does not have.
   */
  private bandSectIsBody = true;
  /** Ids handed to carried inline structures, so the page can name the one it is holding back. */
  private carryKey = 1;
  private readFingerprint = "";
  private commentsPartPath = "word/comments.xml";
  private commentsTarget = "comments.xml";
  private commentsRoot = NEW_COMMENTS_ROOT;
  /**
   * The comments as they were when read, in the shape `commentsDirty` compares against. A package with no
   * comments part read *no* comments, and that is a state, not a blank: left as `""` every document without
   * comments looks dirty, and the first save invents a `word/comments.xml` part, a content-type override and
   * a relationship for a list nobody has.
   * relationship for a list nobody has.
   */
  private commentsAtRead = "[]";
  private commentsXmlAtRead = "";
  private commentsExXmlAtRead = "";
  private commentIdsXmlAtRead = "";
  private commentsExRoot = "";
  private commentIdsRoot = "";
  private commentsExPath = "word/commentsExtended.xml";
  private commentsIdsPath = "word/commentsIds.xml";
  private commentsExTarget = "commentsExtended.xml";
  private commentsIdsTarget = "commentsIds.xml";
  /** Which comment each `w14:paraId` in the comments part belonged to when it was read. */
  private paraOwner = new Map<string, number>();
  title?: string;

  static openBytes(data: Uint8Array): WordDocument {
    const doc = new WordDocument();
    doc.read(OpcPackage.from(data));
    return doc;
  }

  /** A new document with one empty paragraph, ready to type into. */
  static blank(): WordDocument {
    const doc = new WordDocument();
    doc.blocks.push({ kind: "paragraph", inlines: [] });
    // Its page is stated rather than left for a fallback, so a save writes the box the app drew — `<w:pgSz/>`
    // with no width on it is a page of unknown size, and 「页面设置」 on a new document has to start from numbers.
    doc.sections = [copySection(DEFAULT_SECTION)];
    return doc;
  }

  /** The package this document was read from; a save preserves the parts the model does not own. */
  source(): OpcPackage | undefined {
    return this.pkg;
  }

  /** True once the model differs from what was read, so a save must regenerate its parts. */
  get edited(): boolean {
    return this.pkg !== undefined && this.fingerprint() !== this.readFingerprint;
  }
  /**
   * Image bytes are counted by length rather than hashed: a data URL would dwarf the document. A picture
   * replaced in place still has to read as an edit, though, because a regenerated save writes its media part
   * from the model — leaving the length out would keep the old bytes on the disc behind a new picture.
   */
  private fingerprint(): string {
    return (
      JSON.stringify(this.blocks, (key, value) => (key === "src" && typeof value === "string" ? value.length : value)) +
      JSON.stringify(this.bands().map(hfSavedState)) +
      JSON.stringify(this.comments.map(commentState)) +
      // Page geometry lives outside `blocks`: the body's `w:sectPr` travels as bytes the model does not own,
      // and the section list is what the page is drawn from. Left out of the hash, a margin change would read
      // as an untouched document — and `writeDocx` returns the source package byte for byte while `edited` says
      // nothing changed, so the user's dialog would close and the file on disk keep the old page.
      JSON.stringify([this.bodySectPr, this.sections, this.pageSetupSeen])
    );
  }

  private read(pkg: OpcPackage): void {
    this.pkg = pkg;
    this.theme = ThemeFonts.fromPackage(pkg);
    this.readStyles(pkg);
    this.readNumbering(pkg);
    this.readMedia(pkg);

    const docx = pkg.xml("word/document.xml");
    const rels = pkg.rels("word/document.xml");
    for (const [id, r] of rels) {
      this.bodyRels.push({ id, type: r.type, target: r.target, targetMode: r.targetMode, path: pkg.resolve("word/document.xml", r) });
      if (r.type.endsWith("/hyperlink")) this.hyperlinks.set(id, r.target);
    }
    const body = docx?.document?.body;
    if (!body) throw new Error("word/document.xml 里找不到 w:body");
    const source = pkg.text("word/document.xml") ?? "";
    // The body-level `sectPr` is where a document binds headers, footers and column layout, and this model
    // reads none of that — so the direct child keeps its bytes and goes back verbatim. "Direct child" is the
    // whole rule: a section break in the middle of a document hangs its `sectPr` off a paragraph's `w:pPr`,
    // and that one travels with the paragraph. Picking the last `sectPr` anywhere in the body text used to
    // claim such a paragraph's section as the body's own, and the save then wrote it a second time at the
    // end of the body — one section break became two.
    this.bodySectPr = bodyChildSpans(source).find((s) => s.tag === "sectPr")?.xml ?? "";
    const root = rootMarkup(source);
    this.rootNamespaces = root.prefixes;
    this.rootIgnorable = root.ignorable;
    this.rootOpenAttrs = root.attrs;
    // Every section the document says it has, in the order the sections start, read off the bytes a save
    // splices into rather than off the parsed tree — two representations counting sections differently is
    // how a header ends up cited by the section that does not own it.
    const chain = this.sectionChain(source);
    this.sections = chain.map((s) => sectionGeometry(s.xml));
    this.bandSectIsBody = chain.length === 0 || chain[0]!.fromBody;
    this.readHeaderFooters(pkg, rels, chain);
    this.readComments(pkg, rels);
    // The body read off its own source text rather than only off the parsed tree, so a child the model
    // has no fields for — a content control, a customXml island — keeps its bytes and its place instead
    // of dropping out of the document on the first save.
    this.blocks.push(...this.blockChildren(body, bodyChildSpans(source), rels));
    const core = pkg.text("docProps/core.xml");
    this.title = core ? /<dc:title>([^<]*)<\/dc:title>/.exec(core)?.[1] : undefined;
    this.readFingerprint = this.fingerprint();
  }

  private readStyles(pkg: OpcPackage): void {
    this.defaultParaStyleId = undefined;
    this.docDefaultRun = {};
    this.lentMemo.clear();
    this.runLentMemo.clear();
    this.nextMemo.clear();
    if (!pkg.parts.has("word/styles.xml")) return;
    const st = pkg.xml("word/styles.xml");
    const root = unwrap(st, "styles");
    const defaults = root?.docDefaults?.rPrDefault?.rPr;
    // §17.7.2 puts `w:rPrDefault` beneath the *whole* ladder, so it is kept as the one floor the run read ends on
    // and not folded into every style: merged in, a style that states no size would read as stating the default's,
    // and a chain whose ancestor holds the real one would stop at the first hop.
    const stated = (defaults ? runStyle(defaults, this.theme) : {}) as Record<string, unknown>;
    const floor: Record<string, unknown> = {};
    for (const k of LENT_RUN_KEYS) if (stated[k] !== undefined) floor[k] = stated[k];
    this.docDefaultRun = floor as RunStyle;
    for (const s of asArray(root?.style)) {
      const id = s["@_styleId"];
      this.styles.set(id, {
        name: s.name?.["@_val"] ?? id,
        ...(s["@_type"] === undefined ? {} : { type: String(s["@_type"]) }),
        basedOn: s.basedOn?.["@_val"],
        next: s.next?.["@_val"],
        rPr: s.rPr ? runStyle(s.rPr, this.theme) : undefined,
        pPr: s.pPr ? paraStyle(s.pPr) : undefined,
      });
      if (this.defaultParaStyleId === undefined && s["@_type"] === "paragraph" && s["@_default"] === "1") this.defaultParaStyleId = id;
    }
  }

  /**
   * What a paragraph borrows from its styles, resolved the way §17.7.2 describes the chain: the style the
   * paragraph names first, then each `w:basedOn` ancestor in order and, when it names no style — or names one
   * the part does not declare — the document's default paragraph style. Each property comes from the first
   * style in that chain that states it, so no chain has to merge a whole `w:pPr`, and an attribute nobody
   * states stays absent rather than being filled in from the specification's own default.
   *
   * The depth is not a hypothetical. On the authorized 46 份 every file ships a `word/styles.xml` with 406
   * paragraph styles in it, 298 of them carrying `w:basedOn`, and all 298 naming a style the part declares — chains
   * one to three styles deep, no cycles, no dangling reference (`.scratch/basedon-chain.ts`). Counted over the
   * 7,357 paragraphs of `word/document.xml` that write a `w:pPr` (the body holds 7,430 paragraphs; the other 73
   * state no properties at all, 32 of them as a bare `<w:p/>`): 1,518 name a style and 5,839 name none, and none
   * names one the part does not declare. 808 paragraphs ask for a property only an ancestor states — 790 alignments, 44 first-line
   * indents, and 18 line heights, every one of them a multiple: no 固定值 reaches a paragraph through the chain in
   * this cohort, so the exact-height path #131 opened has no audience this far up the ladder. A further 3,829 name
   * no style while the document's default paragraph style states the property (3,731 alignments, 772 line heights —
   * again all multiples — 320 first-line indents, 78 段后), and 70 name a style whose whole chain states nothing
   * while the default style states a property; all 70 are alignments, none a line height. Below the default style
   * the ladder has one more rung, §17.7.2's `w:docDefaults/w:pPrDefault`, which states no paragraph property in any
   * of the 46 份 and so is not read here; `w:basedOn` is walked with a seen-set because a chain that circled would
   * otherwise hang the read, and no chain in this cohort circles.
   */
  private paraLent(styleId: string | undefined): Partial<ParaStyle> | undefined {
    if (this.lentMemo.has(styleId)) return this.lentMemo.get(styleId);
    const out: Record<string, unknown> = {};
    const start = styleId !== undefined && this.styles.has(styleId) ? styleId : this.defaultParaStyleId;
    const seen = new Set<string>();
    for (let cur = start; cur !== undefined && !seen.has(cur); ) {
      seen.add(cur);
      const stated = this.styles.get(cur)?.pPr as Record<string, unknown> | undefined;
      // §17.7.2 puts the default paragraph style beneath the *whole* chain rather than only beneath a paragraph
      // that names no style, so a chain that never reaches it still ends there. The case that showed it — 32
      // paragraphs in 16 of the 86 naming a style whose ancestors state no first-line indent while the default
      // style does — is the quarantined cohort's count, 待重证 on the authorized 46 份.
      const based = this.styles.get(cur)?.basedOn;
      cur = based ?? (this.defaultParaStyleId === undefined || seen.has(this.defaultParaStyleId) ? undefined : this.defaultParaStyleId);
      if (!stated) continue;
      // `w:tabs` 是唯一一枚**不按「哪一级先说就整批算它的」**走的：链上每一级的清单逐位置合进下面那一级，
      // 刚走到的这一级比已经攒下的那些都低（梯子从段落点名的样式往上爬），所以它是被盖的那一方。
      if (stated.tabStops !== undefined) out.tabStops = mergeTabStops(stated.tabStops as TabStop[], out.tabStops as TabStop[] | undefined);
      for (const key of LENT_PARA_KEYS) {
        if (out[key] !== undefined) continue;
        if (stated[key] !== undefined) out[key] = stated[key];
      }
    }
    const lent = Object.keys(out).length ? (out as Partial<ParaStyle>) : undefined;
    this.lentMemo.set(styleId, lent);
    return lent;
  }

  /**
   * What a run borrows from one style chain. The same ladder `paraLent` walks, for character properties instead of
   * paragraph ones: the style named, then each `w:basedOn` ancestor, and — for a paragraph's chain, where §17.7.2
   * puts it beneath the whole chain — the default paragraph style.
   *
   * One chain only, and it stops there. The floor every character has, §17.7.2's `w:docDefaults/w:rPrDefault`, goes
   * in under the whole ladder at `paragraph()` rather than being folded into each chain here: folded in, a
   * character style that states no size would carry the default's size up over the paragraph chain's own, and the
   * run would draw at the bottom rung while a style above the bottom stated the real one.
   *
   * The depth is measured, not assumed — but on the cohort that is no longer evidence, and the styles audience has
   * not been re-counted on the authorized 46 份, so every number in this paragraph is 待重证
   * (`.scratch/run-lent-depth.ts`): 1,961 of 2,378 styles state a
   * `w:rPr` (1,186 paragraph styles, 738 character ones, 37 table ones), and 1,892 styles name a `w:basedOn`.
   * Against that chain, 551 property statements belong to an ancestor rather than to the style itself — 442 sizes,
   * 72 bolds, 37 colours — of which 536 sit one hop up, 13 two and 2 three. Every one of the 551 is missed by a
   * reader that stops at the named style, which is the whole point of walking the chain.
   *
   * `fromPara` is the difference between the two kinds of style: a paragraph's chain ends at the default
   * paragraph style, a character style's ends at its own `w:basedOn` root, because the paragraph default is not
   * a character style's ancestor.
   */
  private runLent(styleId: string | undefined, fromPara = true): Partial<RunStyle> | undefined {
    const key = `${fromPara ? "p" : "c"}:${styleId ?? ""}`;
    if (this.runLentMemo.has(key)) return this.runLentMemo.get(key);
    const out: Record<string, unknown> = {};
    let cur = fromPara && (styleId === undefined || !this.styles.has(styleId)) ? this.defaultParaStyleId : styleId;
    const seen = new Set<string>();
    while (cur !== undefined && !seen.has(cur)) {
      seen.add(cur);
      const stated = this.styles.get(cur)?.rPr as Record<string, unknown> | undefined;
      if (stated) for (const k of LENT_RUN_KEYS) if (out[k] === undefined && stated[k] !== undefined) out[k] = stated[k];
      const based = this.styles.get(cur)?.basedOn;
      cur = fromPara
        ? based ?? (this.defaultParaStyleId === undefined || seen.has(this.defaultParaStyleId) ? undefined : this.defaultParaStyleId)
        : based;
    }
    const lent = Object.keys(out).length ? (out as Partial<RunStyle>) : undefined;
    this.runLentMemo.set(key, lent);
    return lent;
  }

  private readNumbering(pkg: OpcPackage): void {
    if (!pkg.parts.has("word/numbering.xml")) return;
    const root = unwrap(pkg.xml("word/numbering.xml"), "numbering");
    const abstracts = new Map<string, NumberLevel[]>();
    for (const a of asArray(root?.abstractNum)) {
      const key = String(a["@_abstractNumId"]);
      abstracts.set(
        key,
        asArray(a.lvl).map((l: any) => {
          const rFonts = asArray(l.rPr?.rFonts)[0];
          const szHalf = num(l.rPr?.sz?.["@_val"]);
          return {
            numFmt: l.numFmt?.["@_val"] ?? "decimal",
            lvlText: l.lvlText?.["@_val"] ?? "%1.",
            start: Number(l.start?.["@_val"] ?? 1),
            ilvl: num(l["@_ilvl"]),
            indent: num(l.pPr?.ind?.["@_left"]),
            font:
              rFonts?.["@_ascii"] ??
              rFonts?.["@_hAnsi"] ??
              // A level may hand its bullet font to the theme instead of naming a face. That is the same glyph the two
              // rounds above chase (a PUA character from a particular face), so leaving the reference unread draws a
              // box where the file said Wingdings-like art. None of the 201 files did it yet — their numbering.xml
              // holding 421 `w:rFonts` and no theme attribute among them (`.scratch/theme-font-audience.ts`) is the
              // quarantined cohort's count, and `word/numbering.xml` has not been re-counted on the authorized
              // 46 份: 待重证 — so this is the same `w:rFonts` read the same way everywhere else, not a shape
              // measured to be needed.
              this.theme.face(rFonts?.["@_asciiTheme"] ?? rFonts?.["@_hAnsiTheme"], themeFontScript(asArray(l.rPr?.lang)[0]?.["@_eastAsia"])),
            sizePt: szHalf === undefined ? undefined : szHalf / 2,
          };
        }),
      );
    }
    for (const n of asArray(root?.num)) {
      const levels = abstracts.get(String(n.abstractNumId?.["@_val"]));
      if (levels) this.numbering.set(Number(n["@_numId"]), orderLevels(levels));
    }
  }

  private readMedia(pkg: OpcPackage): void {
    for (const [path, bytes] of pkg.parts) {
      if (!path.startsWith("word/media/")) continue;
      const type = path.endsWith(".png") ? "image/png" : path.endsWith(".gif") ? "image/gif" : path.endsWith(".svg") ? "image/svg+xml" : "image/jpeg";
      this.media.set(path, dataUrl(type, bytes));
    }
  }

  private paragraph(p: any, rels: Map<string, Rel2>, span?: string): Paragraph {
    const pPr = p.pPr ?? {};
    const styleId = pPr.pStyle?.["@_val"];
    const lent = this.paraLent(styleId);
    // The runs below read their own properties off this chain too, so a heading whose size lives in the style
    // draws at the size the file chose rather than at whatever the page's stylesheet happens to use. The floor
    // goes on here, once, so it sits under both the paragraph's chain and any character chain a run names.
    const lentRuns = { ...this.docDefaultRun, ...this.runLent(styleId) };
    const ind = indOf(pPr.ind);
    const sp = spacingOf(pPr.spacing);
    // §17.3.1.4: `w:pPr/w:rPr` is the paragraph's own mark — the paragraph break, and in WPS' lists the
    // numbering symbol too (§17.3.1.12 gives that symbol a character run, whose properties the level states and
    // this mark inherits when the level says nothing). Half-points here, the same unit `readRunStyle` converts.
    // The audience that sizes this rung is 待重证: the counts below came from the 86 WPS-written documents
    // (`.scratch/mark-size.ts`) and the paragraph-mark-size audience has not been re-counted on the authorized
    // 46 份 — 455 of 466 listed paragraphs state a mark run, 326 of them a size; of the 396 whose *level* states
    // no size, **every one** resolves to a number — 298 from this mark, 98 from the paragraph's style chain, 0
    // from nothing at all; 331 of the 396 differ from the page stylesheet's 12pt, so a mark that ignores this
    // chain would draw at a size no file asked for.
    const markSz = num(pPr.rPr?.sz?.["@_val"]);
    // Faces the same way, one attribute at a time: a mark that names only `w:eastAsia` must not hand back an
    // undefined Latin face that erases the chain's, and an empty `w:ascii=""` states nothing (§17.7.2).
    const markRFonts = asArray(pPr.rPr?.rFonts)[0];
    // The same one-attribute-at-a-time rule, with the theme answering only where the mark states no literal face.
    const markScript = themeFontScript(asArray(pPr.rPr?.lang)[0]?.["@_eastAsia"]);
    const markAscii = markRFonts?.["@_ascii"] || this.theme.face(markRFonts?.["@_asciiTheme"], markScript) || undefined;
    const markEast = markRFonts?.["@_eastAsia"] || this.theme.face(markRFonts?.["@_eastAsiaTheme"], markScript) || undefined;
    // Colour the same way, with the two filters `readRunStyle` applies to a run's own: `auto` is no statement,
    // and a value that is not a colour is no statement either — the mark then falls through to the chain.
    const markColorRaw = asArray(pPr.rPr?.color)[0]?.["@_val"];
    const markColor = markColorRaw && markColorRaw !== "auto" ? hex(markColorRaw) : undefined;
    const para: Paragraph = {
      kind: "paragraph",
      styleId,
      styleName: styleId ? this.styles.get(styleId)?.name : undefined,
      align: pPr.jc?.["@_val"] ?? lent?.align,
      ...indResolve(ind, lent),
      spacingBefore: sp.spacingBefore ?? lent?.spacingBefore,
      spacingAfter: sp.spacingAfter ?? lent?.spacingAfter,
      lineSpacing: sp.lineSpacing ?? lent?.lineSpacing,
      lineSpacingRule: sp.lineSpacingRule ?? lent?.lineSpacingRule,
      pageBreakBefore: onOff(pPr.pageBreakBefore) ?? false,
      keepNext: onOff(pPr.keepNext) ?? lent?.keepNext,
      keepLines: onOff(pPr.keepLines) ?? lent?.keepLines,
      widowControl: onOff(pPr.widowControl) ?? lent?.widowControl,
      snapToGrid: onOff(pPr.snapToGrid) ?? lent?.snapToGrid,
      wordWrap: onOff(pPr.wordWrap) ?? lent?.wordWrap,
      // The box, per side: the paragraph's own statement first, the chain's sides it does not name behind. A
      // paragraph that names no side at all takes the chain's object by reference (see `mergeParaBorders`), so
      // the write side still sees it as the style's.
      borders: mergeParaBorders(paraBordersOf(asArray(pPr.pBdr)[0]), lent?.borders),
      shading: shadingOf(asArray(pPr.shd)[0]) ?? lent?.shading,
      // 制表位逐位置合：段落自己那一枚顶掉链上同位置的，链上其余的留着（见 `mergeTabStops`）。
      tabStops: mergeTabStops(lent?.tabStops, tabStopsOf(asArray(pPr.tabs)[0])),
      numPr: numPrOf(pPr.numPr) ?? lent?.numPr,
      outlineLevel: outlineOf(pPr.outlineLvl) ?? lent?.outlineLevel,
      markSizePt: markSz === undefined ? lentRuns.sizePt : markSz / 2,
      markFont: markAscii ?? lentRuns.font,
      markFontEastAsia: markEast ?? lentRuns.fontEastAsia,
      markColor: markColor ?? lentRuns.color,
      inlines: [],
    };
    // The paragraph's own tag can carry an identity this model has no field for — `w14:paraId`, a revision
    // save id. It is read off the source text and written back onto the tag, so a save does not make every
    // paragraph in the document look like a paragraph that was never there.
    const own = ownAttrs(span);
    if (own) para.attrsXml = own;
    // The whole `w:pPr` also travels as text. Reading it off the parsed tree would answer "which properties
    // does the model have a field for", which is not the same question as "what did the file write".
    const pprXml = elementOf(childSpans(elementContent(span ?? "")), "pPr");
    const carried = paraSlotsOf(childSpans(elementContent(pprXml)));
    if (carried.length) para.pPrSlots = carried;
    para.asRead = paraReadSnapshot(para);
    // The paragraph's own mark can be revised: `w:pPr/w:rPr` holds the break that ends this paragraph,
    // signed the same way a run is. Its character properties are read for the numbering symbol's own three —
    // size, the two faces, colour (`markSizePt`, `markFont`, `markFontEastAsia`, `markColor`) — and the break's
    // signature is what the rest of this loop is for: losing it would leave a merge nobody can refuse.
    for (const c of children(pPr.rPr ?? {})) {
      if (isRevision(c.name)) {
        para.markRev = revisionOf(c.name, c.node);
        break;
      }
    }
    const paraChange = pPr.pPrChange;
    if (paraChange) {
      const recorded = paraSlotsOf(spansUnder(pprXml, "pPrChange", "pPr"));
      para.revFormat = {
        kind: "format",
        id: Number(paraChange["@_id"]) || 0,
        author: String(paraChange["@_author"] ?? ""),
        date: paraChange["@_date"],
        prev: paraStyle(paraChange.pPr ?? {}),
        ...(recorded.length ? { prevSlots: recorded } : {}),
      };
    }
    for (const step of this.paragraphChildren(p, span)) {
      // A child this model carries instead of reads — a bookmark, a field, an inline control — travels as
      // the bytes it came in with.
      if (step.node === undefined) {
        if (step.span) {
          const carry: CarriedInline = { key: this.carryKey++, tag: step.tag, xml: step.span.xml, show: carriedText(step.span.xml) };
          if (step.name !== undefined) carry.name = step.name;
          para.inlines.push({ text: "", carry });
        }
        continue;
      }
      const name = step.tag;
      if (name === "r") para.inlines.push(...this.run(step.node, rels, step.span?.xml, lentRuns));
      else if (name === "commentRangeStart" || name === "commentRangeEnd") {
        const id = Number(step.node?.["@_id"]);
        if (Number.isFinite(id)) para.inlines.push(anchorInline({ id, kind: name === "commentRangeStart" ? "start" : "end" }));
      } else if (isRevision(name)) {
        this.revisionRuns(step.node, rels, para.inlines, revisionOf(name, step.node), step.span?.xml, lentRuns);
      } else if (name === "hyperlink") {
        const target = this.linkTarget(step.node, rels);
        const extras = linkOwnAttrs(step.span?.xml);
        for (const inline of this.linkRuns(step.node, rels, step.span?.xml, lentRuns)) {
          // The attributes belong to the element, and the element only survives as an element while the model
          // can name its target: a link it cannot point at is already plain runs, and half-carrying its
          // attributes onto a run that has no link would write them nowhere.
          if (target) {
            inline.hyperlink = target;
            if (extras) inline.linkAttrs = extras;
          }
          linkLook(inline);
          para.inlines.push(inline);
        }
      }
    }
    return para;
  }

  /**
   * A paragraph's direct children in the order the part wrote them. The parsed tree says what each child
   * means, the source spans say what it was written as, and the spans decide the order: a bookmark that sat
   * between two runs has to come back between those two runs. Where the two accounts of the paragraph
   * disagree, only the parsed tree is used and an unknown child is skipped as it always was — a paragraph
   * whose children cannot be paired is not one to start guessing positions for.
   */
  private paragraphChildren(p: any, span: string | undefined): ParaChild[] {
    const parsed = children(p).filter((child) => child.name !== "pPr");
    const own = span === undefined ? [] : childSpans(elementContent(span)).filter((s) => s.tag !== "pPr");
    if (span === undefined || !sameTags(parsed, own)) {
      return parsed.map((child) => ({ tag: child.name, node: child.node }));
    }
    const queues = new Map<string, any[]>();
    for (const child of parsed) {
      const queue = queues.get(child.name);
      if (queue) queue.push(child.node);
      else queues.set(child.name, [child.node]);
    }
    const readable = (tag: string): boolean => tag === "r" || tag === "hyperlink" || tag === "commentRangeStart" || tag === "commentRangeEnd" || isRevision(tag);
    return own.map((s): ParaChild => {
      const node = readable(s.tag) ? queues.get(s.tag)?.shift() : undefined;
      if (node === undefined) {
        // Machinery this model carries rather than reads. The bytes are all it needs — except a bookmark
        // start, whose `w:name` is read off the parsed element here while the read still has it: nothing
        // downstream can pair the raw `xml` with its element again, and the 「定位」 bookmark target
        // addresses these by name.
        const parsedNode = queues.get(s.tag)?.shift();
        const name = parsedNode?.["@_name"];
        return { tag: s.tag, span: s, ...(name === undefined ? {} : { name: String(name) }) };
      }
      // A run that is only field machinery keeps the name of the machinery it holds, so the model and the
      // page can say which part of the field this is rather than just "a run".
      const kind = s.tag === "r" ? machineryOf(node) : "";
      if (kind) return { tag: kind, span: s };
      return { tag: s.tag, node, span: s };
    });
  }

  /**
   * An element's children lined up the same way: what each one means, next to what the part wrote it as.
   * Where the two accounts disagree the spans are left go rather than guessed at — handing a run another
   * run's `w:rPr` would draw the page wrong, and that is worse than rebuilding it from its fields.
   */
  private childSteps(node: any, span: string | undefined): { tag: string; node: any; span?: { tag: string; xml: string } }[] {
    const parsed = children(node);
    const own = span === undefined ? [] : childSpans(elementContent(span));
    if (span === undefined || !sameTags(parsed, own)) return parsed.map((c) => ({ tag: c.name, node: c.node }));
    const queues = new Map<string, any[]>();
    for (const child of parsed) {
      const queue = queues.get(child.name);
      if (queue) queue.push(child.node);
      else queues.set(child.name, [child.node]);
    }
    return own.map((s) => ({ tag: s.tag, node: queues.get(s.tag)?.shift(), span: s }));
  }

  /** The runs inside a `w:hyperlink`, which may themselves sit inside a `w:ins` or `w:del`. */
  private linkRuns(node: any, rels: Map<string, Rel2>, span?: string, fromPara?: Partial<RunStyle>): Inline[] {
    const out: Inline[] = [];
    for (const child of this.childSteps(node, span)) {
      if (child.node === undefined) continue;
      if (child.tag === "r") out.push(...this.run(child.node, rels, child.span?.xml, fromPara));
      else if (isRevision(child.tag)) this.revisionRuns(child.node, rels, out, revisionOf(child.tag, child.node), child.span?.xml, fromPara);
    }
    return out;
  }

  /** A revision's runs are ordinary runs with a signature on them; deleted text is spelled `w:delText`. */
  private revisionRuns(node: any, rels: Map<string, Rel2>, into: Inline[], rev: Revision, span?: string, fromPara?: Partial<RunStyle>): void {
    for (const child of this.childSteps(node, span)) {
      if (child.node === undefined) continue;
      if (child.tag === "r") {
        for (const inline of this.run(child.node, rels, child.span?.xml, fromPara)) into.push({ ...inline, rev });
        continue;
      }
      // A revision wraps a link as readily as a run, and the link keeps its own target.
      if (isRevision(child.tag)) {
        // Word nests a revision inside one when revised text is revised again; the inner signature is
        // the change that moved these characters, so it wins over the one we are already reading.
        this.revisionRuns(child.node, rels, into, revisionOf(child.tag, child.node), child.span?.xml, fromPara);
        continue;
      }
      if (child.tag !== "hyperlink") continue;
      const target = this.linkTarget(child.node, rels);
      const extras = linkOwnAttrs(child.span?.xml);
      for (const inline of this.linkRuns(child.node, rels, child.span?.xml, fromPara)) {
        if (target) {
          inline.hyperlink = target;
          if (extras) inline.linkAttrs = extras;
        }
        linkLook(inline);
        into.push({ ...inline, rev });
      }
    }
  }

  /**
   * Where a `w:hyperlink` points. A link to the outside arrives through its relationship; a link to a place
   * in this document has no relationship and is spelled `w:anchor`, which this model carries as a `#`-prefixed
   * target so the writer can tell the two apart and put the bookmark name back.
   */
  private linkTarget(node: any, rels: Map<string, Rel2>): string | undefined {
    const id = node["@_id"];
    const throughRel = id === undefined ? undefined : this.hyperlinks.get(String(id)) ?? rels.get(String(id))?.target;
    const anchor = node["@_anchor"];
    return throughRel ?? (anchor === undefined || anchor === "" ? undefined : `#${String(anchor)}`);
  }

  private run(r: any, rels: Map<string, Rel2>, span?: string, fromPara?: Partial<RunStyle>): Inline[] {
    // Three rungs hold a run's properties, and the file can put a field on any of them: the run's own `w:rPr`,
    // the character style it names, and the paragraph it sits in. §17.7.2 ranks them in that order, so only a
    // field the run leaves open is taken from below. `w:rStyle` says where the properties come from and is not
    // itself one of them, so it is read off the parsed node rather than through `runStyle`.
    const own = runStyle(r.rPr ?? {}, this.theme);
    const named = r.rPr?.rStyle?.["@_val"];
    const fromChar = typeof named === "string" ? this.runLent(named, false) : undefined;
    const style: RunStyle = { ...fromPara, ...fromChar, ...own };
    // The run's own `w:rPr` travels as text, the same way the paragraph's does: read off the parsed tree the
    // answer is "which properties does the model have a field for", and that list is not what a Chinese
    // document writes. An empty `w:rPr` counts as written: the set the file chose is the one that goes back.
    const props = elementOf(childSpans(elementContent(span ?? "")), "rPr");
    const carried = props ? runSlotsOf(childSpans(elementContent(props))) : undefined;
    const change = r.rPr?.rPrChange;
    const record = runSlotsOf(spansUnder(props ?? "", "rPrChange", "rPr"));
    // A formatting revision signs the run it changed and carries the properties from before.
    const revFormat: FormatRevision | undefined = change
      ? {
          kind: "format",
          id: Number(change["@_id"]) || 0,
          author: String(change["@_author"] ?? ""),
          date: change["@_date"],
          prev: runStyle(change.rPr ?? {}, this.theme),
          ...(record.length ? { prevSlots: record } : {}),
        }
      : undefined;
    // The record is on the run, so everything the run holds carries it. The read snapshot is the *resolved*
    // set, always: a field that came up from the paragraph or the character style is not the run's own, and a
    // save that spelled it into `w:rPr` would write a property the file never put on the run.
    const base = {
      ...style,
      ...(carried ? { rPrSlots: carried } : {}),
      asRead: { ...style },
      ...(revFormat ? { revFormat } : {}),
    };
    const out: Inline[] = [];
    let unmodelled = false;
    // The run's children come paired with the part's own text, not just off the parsed tree: a drawing's
    // frame and effects live below the fields the model reads, and the only faithful copy of them is the
    // bytes the file wrote.
    for (const child of this.childSteps(r, span)) {
      switch (child.tag) {
        case "t":
        case "delText":
          out.push({ ...base, text: typeof child.node === "string" ? child.node : (child.node["#text"] ?? "") });
          break;
        case "br": {
          // Every `w:br` is a break — a page break no less than a soft return. Which one it was is the file's
          // own business: the token travels as it came rather than being inferred from `w:type` alone.
          out.push({ ...base, text: "", break: true, ...(child.span ? { brXml: child.span.xml } : {}) });
          break;
        }
        case "tab":
          out.push({ ...base, text: "", tab: true });
          break;
        case "commentReference": {
          const id = Number(child.node?.["@_id"]);
          // The reference mark keeps the `w:rPr` the file wrote for it. Word styles this run
          // `CommentReference`, and a save that re-spelled the run from that one field would drop whatever else
          // the file put in there; a mark this editor inserts has no file behind it and gets Word's own shape.
          if (Number.isFinite(id))
            out.push({ text: "", anchor: { id, kind: "ref" }, ...(base.rPrSlots ? { rPrSlots: base.rPrSlots } : {}) });
          break;
        }
        case "sym":
          out.push({ ...base, text: String.fromCharCode(parseInt(child.node?.["@_char"] ?? "20", 16)) });
          break;
        case "drawing":
        case "pict": {
          const img = imageOf(child.node, rels, this.media, this.pkg);
          if (img) {
            const embed = child.span ? embedOf(child.span.xml) : undefined;
            out.push({
              ...base,
              text: "",
              image: img,
              ...(child.span
                ? { drawing: { xml: child.span.xml, width: img.width, height: img.height, name: img.name, ...(embed ? { embed } : {}) } }
                : {}),
            });
            break;
          }
          // A picture this reader cannot resolve — an external `r:link`, a frame it does not know — still has
          // bytes, and the run that holds them is not empty. Carrying the drawing keeps the shape on the page
          // and the citation in the body intact; reading it as nothing would delete both.
          if (child.span)
            out.push({ ...base, text: "", carry: { key: this.carryKey++, tag: child.tag, xml: child.span.xml, show: carriedText(child.span.xml) } });
          break;
        }
        default:
          if (child.tag !== "rPr") unmodelled = true;
      }
    }
    // A run built around something the model has no case for — a shape wrapped in `mc:AlternateContent`, an
    // OLE object — reads as nothing at all, and a paragraph that reads as nothing is one a save deletes.
    // The bytes go back where they came from instead of the drawing going down with the reader's blind spot.
    if (!out.length && span && unmodelled)
      return [{ ...base, text: "", carry: { key: this.carryKey++, tag: "r", xml: span, show: carriedText(span) } }];
    // 一枚 run 装了并排的孩子时，这串文字属于整枚 run，孩子各拿一份引用：孩子自己没有 `w:rPr`，共用的是上面
    // 那一套槽，所以「件里这段话说的是中文」只能这样问回去。一个孩子时不必写——问题本来就只有一句文字。
    if (out.length > 1) {
      const runText = out.map((inline) => inline.text).join("");
      for (const inline of out) inline.runText = runText;
    }
    return out;
  }

  private table(tbl: any, rels: Map<string, Rel2>, source?: string): Table {
    // Properties the model owns come off the parsed tree; the ones it does not are kept from the source
    // text in the slot they came from. Without that second half, editing one character of body text
    // would rewrite a table as a bare grid and report nothing lost.
    const own = source ? childSpans(elementContent(source)) : [];
    const spanOf = (tag: string): string | undefined => own.find((s) => s.tag === tag)?.xml;
    const columnWidths = asArray(tbl?.tblGrid?.gridCol)
      .map((g: any) => num(g?.["@_w"]))
      .filter((w: number | undefined): w is number => w !== undefined);
    const trSpans = own.filter((s) => s.tag === "tr");
    const rows = asArray(tbl.tr).map((tr: any, r: number): TableCell[] => {
      const rowSpan = trSpans[r];
      const rowOwn = rowSpan ? childSpans(elementContent(rowSpan.xml)) : [];
      // A row is whatever else it holds (`w:tblPrEx`, which WPS puts first), its `w:trPr`, and its cells.
      // The model has no fields for that first group, so it travels as the source text it came in with,
      // anchored to the siblings it was read between.
      let cellsSoFar = 0;
      let sawProps = false;
      const extras: RowExtra[] = [];
      for (const s of rowOwn) {
        if (s.tag === "tc") {
          cellsSoFar += 1;
          continue;
        }
        if (s.tag === "trPr") {
          sawProps = true;
          continue;
        }
        extras.push({ tag: s.tag, xml: s.xml, afterCells: cellsSoFar, beforeProps: !sawProps });
      }
      const slots = slotsOf(childSpans(elementContent(elementOf(rowOwn, "trPr"))));
      const info = rowInfoOf(tr.trPr, slots, ownAttrs(rowSpan?.xml), extras, marginOf(tr.tblPrEx?.tblCellMar));
      // A row's `w:trPrChange` is one of the slots the model carries rather than a field it owns, so it is
      // read twice: the parsed node for what the page draws, the slot's text for what a decision rewrites.
      if (info) info.revFormat = rowFormatRevision(tr.trPr, slots);
      const tcSpans = rowOwn.filter((s) => s.tag === "tc");
      return asArray(tr.tc).map((tc: any, c: number): TableCell => {
        const cellSpan = tcSpans[c];
        const cellOwn = cellSpan ? childSpans(elementContent(cellSpan.xml)) : [];
        const tcPr = elementOf(cellOwn, "tcPr");
        const cell: TableCell = {
          paragraphs: this.blockChildren(tc, cellOwn, rels),
          ...cellStyleOf(tc.tcPr),
          row: info,
          tcPrXml: tcPr || undefined,
          attrsXml: ownAttrs(cellSpan?.xml) || undefined,
        };
        const revs = cellRevisions(tc.tcPr, tcPr || undefined);
        if (revs.length) cell.revs = revs;
        return cell;
      });
    });
    const table: Table = { kind: "table", rows, columns: Math.max(1, ...rows.map((r) => r.length)), ...tableStyleOf(tbl.tblPr) };
    // The schema lets range markup and containers sit between a table's rows, and WPS does write them there:
    // a `w:bookmarkEnd` closing a bookmark that started above the table. These have no fields to read, so each
    // travels with the row it followed.
    const betweenRows: { afterRow: number; tag: string; xml: string }[] = [];
    let rowsSoFar = -1;
    for (const s of own) {
      if (s.tag === "tr") rowsSoFar += 1;
      else if (s.tag !== "tblPr" && s.tag !== "tblGrid") betweenRows.push({ afterRow: rowsSoFar, tag: s.tag, xml: s.xml });
    }
    if (betweenRows.length) table.betweenRows = betweenRows;
    const grid = spanOf("tblGrid");
    if (grid) table.gridXml = grid;
    if (columnWidths.length) table.columnWidths = columnWidths;
    const tblPr = spanOf("tblPr");
    if (tblPr) table.tblPrXml = tblPr;
    const changed = tableFormatRevision(tbl.tblPr, tblPr);
    if (changed) table.revFormat = changed;
    return table;
  }

  /**
   * An element's block children — the body's or a table cell's — in the order the part gives them. Paragraphs
   * and tables become model objects; anything else the model has no fields for is carried as the source text
   * it came in with. The parsed tree says what a block means and the spans say what it was written as, so the
   * two are read together: the spans decide the order, which is the only record of where a container sat
   * between two paragraphs.
   */
  private blockChildren(node: any, spans: { tag: string; xml: string }[], rels: Map<string, Rel2>): Block[] {
    // A `w:sectPr` is the body's own section list rather than a block, and a `w:tcPr` is the cell's property
    // set, which the reader takes apart where it reads the cell — neither is content to carry.
    const kept = (tag: string): boolean => tag !== "sectPr" && tag !== "tcPr";
    const parsed = children(node).filter((child) => kept(child.name));
    const own = spans.filter((span) => kept(span.tag));
    const out: Block[] = [];
    if (sameTags(parsed, own)) {
      const queues = new Map<string, any[]>();
      for (const child of parsed) {
        const queue = queues.get(child.name);
        if (queue) queue.push(child.node);
        else queues.set(child.name, [child.node]);
      }
      for (const span of own) {
        const child = queues.get(span.tag)?.shift();
        if (span.tag === "p") out.push(this.paragraph(child, rels, span.xml));
        else if (span.tag === "tbl") out.push(this.table(child, rels, span.xml));
        else out.push(carriedOf(span));
      }
      return out;
    }
    // The two accounts of what is in there disagree — a part the span scan could not follow, most likely
    // because a close tag did not match its open one. Then the parsed tree is the only list of blocks and a
    // container has no position to be given back; the tables still read their properties from the spans a
    // positional pairing reaches.
    const tables = own.filter((s) => s.tag === "tbl").map((s) => s.xml);
    let carried = 0;
    for (const child of parsed) {
      if (child.name === "p") out.push(this.paragraph(child.node, rels));
      else if (child.name === "tbl") out.push(this.table(child.node, rels, tables[carried++]));
    }
    return out;
  }

  levelOf(numId: number, ilvl: number): NumberLevel | undefined {
    return this.numbering.get(numId)?.[ilvl];
  }

  /** Numbering definitions by id, for a writer that has to regenerate word/numbering.xml. */
  numberingTable(): { numId: number; levels: NumberLevel[] }[] {
    return [...this.numbering].map(([numId, levels]) => ({ numId, levels }));
  }

  /**
   * The numbering a 「项目符号」/「1. 编号」 click needs: the file's own definition when it names *that* shape,
   * one written here when it does not.
   *
   * Reusing the document's definitions was the whole strategy before, and the numbers that said what that bought
   * (`.scratch/list-availability-audience.ts`: 项目符号 having nothing to point at in **80 of 86** files and 编号 in
   * **48 of 86**, 47 carrying no `word/numbering.xml` at all, and in the 38 files where 编号 did answer the id
   * picked being the first one whose level 0 was not a bullet with **29 of those 38** saying
   * `w:numFmt val="chineseCounting"` — the button labelled 1. drew 一、) came from the quarantined 86 份 cohort;
   * the numbering audience has not been re-counted on the authorized 46 份, so every one of those counts is
   * 待重证. The 文件 → 新建 shape behind 「47 carry none」 is the shape a blank document has whatever the count. So
   * the pick is now exact (`decimal` for 编号), and a file that names neither shape gets a definition
   * instead of a grey button, the way WPS writes one when you click on a blank page.
   *
   * What gets invented is the core a level really does state: `w:start`, `w:numFmt`, `w:lvlText`,
   * `w:lvlJc` — all five shapes that probe counted in column D carry those four, in §17.4's order (same 待重证).
   * What is
   * deliberately *not* invented is `w:ind` and `w:suff`: those are a layout claim, and what WPS puts in a new
   * document's own definition is **未取证**.
   */
  ensureListNumbering(kind: "bullet" | "number"): number {
    const want = LIST_SHAPE[kind].numFmt;
    for (const [numId, levels] of this.numbering) {
      if (this.inventedNumbers.has(numId)) continue;
      if (levels[0]?.numFmt === want) return numId;
    }
    for (const [numId, held] of this.inventedNumbers) if (held.kind === kind) return numId;
    const ids = this.nextListIds();
    this.inventedNumbers.set(ids.numId, { kind, abstractId: ids.abstractId });
    this.numbering.set(ids.numId, [{ ...LIST_SHAPE[kind], start: 1, ilvl: 0 }]);
    return ids.numId;
  }

  /**
   * The document's own 默认制表位 (`w:defaultTabStop` in `word/settings.xml`), in twips — the distance
   * 增加缩进量／减少缩进量 moves a paragraph that spells its indent in twips. On the authorized 46 份 every one
   * states it: 43 spell `w:val="420"` and 3 spell `720`, and none leaves the element out
   * (`.scratch/authorized-tab-stop.py`; the earlier 36-of-39 over distinct bundle bodies is 待重证 — that walk
   * deduped by content hash, a different audience from the 件 the rest of this file counts). `word/settings.xml`
   * is not a part this model owns, so this reads it on demand and a save still carries the bytes the file wrote.
   */
  defaultTabStop(): number | undefined {
    const raw = this.pkg ? new TextDecoder().decode(this.pkg.bytes("word/settings.xml") ?? new Uint8Array()) : "";
    const v = /<(?:[\w-]+:)?defaultTabStop\s+(?:[\w-]+:)?val="(\d+)"/.exec(raw)?.[1];
    return v === undefined ? undefined : Number(v);
  }

  /**
   * 「字数统计」报的数，由模型自己数出来。范围只有**正文**——表格里的那些段落算（它们是 `w:tbl` 里的 `w:p`，
   * 与正文同一段逻辑），页眉页脚、脚注尾注、文本框都不算：那几条在 WPS 的对话框里是勾选项，它的默认勾选
   * **未取证**，所以只数没有争议的那一块。
   *
   * 四条口径都是**推导**（WPS 真机把这几个数报成什么，本轮没有实据）：
   * - 段落数数的是段落标记，空段落也算一个；嵌套表顺着 `w:tc` 的块递归下去，一格里的表不漏。
   * - 字符数按**码点**数（`for…of`），否则一个增补平面的字算成两个；「不计空格」把 Unicode `White_Space`
   *   剔掉——里面含全角空格 U+3000 与不换行空格 U+00A0，这两个在中文件里最常见。
   * - 「中文字符和朝鲜语单词」只数汉字、假名、谚文这些**字位**，中文标点不计；「非中文单词」是空白与这些
   *   字位切出来的、至少含一个字母或数字的片段。`字数 = 中文字符 + 非中文单词`，与 Word 一族那三行的连法一致。
   * - 制表符、换行、域指令一个都不计：它们由 `w:tab`/`w:br`/`w:instrText` 说，不在 `inlines[].text` 上，
   *   而这里只认文字节点。`w:del` 里的删字仍带着 `text`，**算进字数**——这是本轮的取舍，与页面上画删除线
   *   却还留着它们同一套逻辑（未接受的修订还没消失）。
   */
  wordCounts(): WordCounts {
    return countTexts(this.bodyTexts());
  }

  /**
   * 正文里那一段段的文字，一段一个元素（空段也占一个）：表格连同格子里的段落都算，页眉页脚／脚注尾注／文本框不算。
   * 这是数的那张口的**取材**那一半——`wordCounts()` 数它，选区数的是它的切片。
   */
  bodyTexts(): string[] {
    return paragraphTextsOf(this.blocks);
  }

  /**
   * Where 增加缩进量／减少缩进量 puts this paragraph's left indent: the pair to write, or `undefined` when the
   * paragraph already sits on that edge and the click would change nothing.
   *
   * The unit follows how the file spells that paragraph's indent, because the two spellings are the same number
   * twice (see `ParaStyle.indentLeftChars`): a paragraph stating `w:leftChars` moves by 2 字符 and its twips
   * value follows the per-character width the file itself used (210 for a 10.5pt paragraph, 399 for a 19.9pt
   * one — read off its own pair), while a paragraph that only states twips moves by one `w:defaultTabStop` and
   * gets no character count invented for it. Stepping only one of the two is what would make the file disagree
   * with itself, and the page would show the twips while Word showed the characters.
   *
   * It stops at 0 rather than going negative: WPS writes `w:left="0"` out loud 187 times in the bodies of the
   * cohort that was quarantined — indent values are 待重证 on the authorized 46 份 — and that is its way of
   * saying 回到边内. Whether the button keeps going past the left margin onto negative
   * values (that cohort also held `w:left="-1"`, `-41`, `-178`) is 未取证 — no file here was made by clicking it.
   */
  nextIndent(p: Paragraph, dir: 1 | -1): { indentLeft: number; indentLeftChars?: number } | undefined {
    const stepChars = 200;
    const tab = this.defaultTabStop() ?? 420;
    const chars = p.indentLeftChars;
    const left = p.indentLeft ?? 0;
    let nextLeft: number;
    let nextChars: number | undefined;
    if (chars !== undefined) {
      nextChars = Math.max(0, chars + dir * stepChars);
      nextLeft = chars > 0 && left > 0 ? Math.round((nextChars * left) / chars) : Math.max(0, left + dir * tab);
    } else {
      nextLeft = Math.max(0, left + dir * tab);
    }
    if (nextLeft === left && nextChars === chars) return undefined;
    return { indentLeft: nextLeft, ...(nextChars === undefined ? {} : { indentLeftChars: nextChars }) };
  }

  /** Whether `word/numbering.xml` owes new bytes: only definitions handed out above make it dirty. */
  listDirty(): boolean {
    return this.inventedNumbers.size > 0;
  }

  /** The definitions to splice into (or to start) `word/numbering.xml`, oldest first. */
  listAdditions(): ListAddition[] {
    return [...this.inventedNumbers].map(([numId, held]) => ({
      kind: held.kind,
      numId,
      abstractId: held.abstractId,
      lvlText: LIST_SHAPE[held.kind].lvlText,
      numFmt: LIST_SHAPE[held.kind].numFmt,
    }));
  }

  /**
   * Ids nothing in this package has used yet. Both accounts are read: the model's own map for `w:num`, and the
   * part's bytes for `w:abstractNumId`, because the model keeps levels, not the ids the file numbered them with.
   */
  private nextListIds(): { numId: number; abstractId: number } {
    const raw = this.pkg ? new TextDecoder().decode(this.pkg.bytes("word/numbering.xml") ?? new Uint8Array()) : "";
    // Either spelling of an id counts: `<w:num w:numId="1">` puts it on the element, `<w:numId w:val="1"/>` puts it
    // on a child, and the prefix is whatever the part declared. A name that merely *starts* with the same letters
    // (`w:numIdMacAtCleanup`) must not read as one, so `=` or the part's own `val` attribute has to follow.
    const ids = (name: string): number[] =>
      [...raw.matchAll(new RegExp(`(?:[\\w-]+:)?${name}(?:="(-?\\d+)"|\\s+(?:[\\w-]+:)?val="(-?\\d+)")`, "g"))].map(
        (m) => Number(m[1] ?? m[2]),
      );
    const numId = Math.max(0, ...this.numbering.keys(), ...ids("numId"));
    const abstractId = Math.max(0, ...[...this.inventedNumbers.values()].map((h) => h.abstractId), ...ids("abstractNumId"));
    return { numId: numId + 1, abstractId: abstractId + 1 };
  }

  /** Referenced style definitions, for a writer that has to regenerate word/styles.xml. */
  styleDef(styleId: string): StyleDef | undefined {
    return this.styles.get(styleId);
  }

  /**
   * What a paragraph that names this style borrows from it and its ancestors: the same ladder §17.7.2 builds for
   * a paragraph read off the part, so a paragraph this editor builds draws and saves like one the file wrote.
   */
  paragraphStyleOf(styleId: string | undefined): Partial<ParaStyle> | undefined {
    return this.paraLent(styleId);
  }

  /** What a run in a paragraph that names `styleId` borrows, floor included — the same resolution `paragraph` reads. */
  paraRunStyle(styleId: string | undefined): RunStyle {
    return { ...this.docDefaultRun, ...this.runLent(styleId) };
  }

  /**
   * The paragraph styles this part declares, in the order it declares them: the gallery 开始页 offers.
   * §17.7.2.3 gives `w:type` the default `paragraph`, so a style that states none is one of these; character,
   * table and numbering styles are not candidates. The default paragraph style is marked because picking it is
   * 「不点名样式」 — the paragraph's `w:pStyle` goes away rather than naming the default.
   *
   * 授权 46 份共 **112 枚段落样式**（每份 6～27 枚，`.scratch/para-style-gallery-audience.py` 读 `word/styles.xml`
   * 与 `word/document.xml` 的 `w:pStyle`），其中 103 枚带 `w:qFormat`——WPS 画廊露出来的正是这一批；这一格先摆
   * **文件自己声明的每一枚段落样式**，不按 `w:qFormat` 筛，也不替它编名字（有的件在正文里只用到其中几枚）。
   */
  paragraphStyles(): { id: string; name: string; isDefault: boolean }[] {
    const out: { id: string; name: string; isDefault: boolean }[] = [];
    for (const [id, def] of this.styles) {
      if (def.type !== undefined && def.type !== "paragraph") continue;
      out.push({ id, name: def.name, isDefault: id === this.defaultParaStyleId });
    }
    return out;
  }

  /**
   * 把这一段换成 `styleId` 点名的段落样式；`styleId === undefined` 是「不点名」——正文/默认样式，`w:pStyle` 撤掉。
   *
   * 换样式不是把新样式的属性抄到段落身上：那会把「这一段只点名了一个样式」变成「这一段自己写了十几句」，
   * 再换回去也撤不掉。段落自己写下的那一份直接格式留着（槽位看得见的那几项原样认作它的，槽位看不见的就按
   * 「与旧链不同」认，和 `nextSplit` 同一条判据）；剩下的位置留给新链——字段改成新链的值，`asRead` 也换成新链，
   * 写出侧对着它一句话都不多写，保存回来就是 `w:pStyle` 换了一句。
   *
   * run 与段落标记也一样：腿是画在每枚 run 自己的 `styleOf` 上的（`render.ts` 的段落只带字号），所以新链的
   * 字号、字体、颜色、粗细得重新落到 run 字段上，否则「换成标题 1」在纸上只挪了段距、字还是正文那么大。
   * run 自己 `w:rPr` 说过的（槽位里有那一枚孩子）和字符样式借来的(`w:rStyle`)不动——直接格式优先于样式。
   */
  applyParaStyle(p: Paragraph, styleId: string | undefined): boolean {
    if (styleId !== undefined && !this.styles.has(styleId)) return false;
    if (p.styleId === styleId) return false;
    const before = this.paragraphStyleOf(p.styleId);
    const nextRuns = this.paraRunStyle(styleId);
    const lent: ParaStyle = { ...this.paragraphStyleOf(styleId) };
    const fields: Record<string, unknown> = { ...lent };
    const own = ownParaStyle(before, paraStyleOf(p));
    for (const key of PARA_STYLE_KEYS) {
      if (key === "styleId") continue;
      if (own[key] !== undefined || slotStates(p.pPrSlots, PARA_SLOT_STATES[key])) fields[key] = p[key];
    }
    // 那八项不在 `PARA_STYLE_KEYS` 里（整枚靠槽位带回），与 `nextSplit` 同一条判据：槽位说过、或与旧链不同，
    // 都是这一段自己写的；框按引用比，因为链借来的那个物件是原样交回来的。
    for (const key of MODEL_ONLY_PARA_KEYS) {
      const stated = p[key];
      if (stated === undefined) continue;
      if (slotStates(p.pPrSlots, PARA_SLOT_STATES[key]) || stated !== before?.[key]) fields[key] = stated;
    }
    const markSlot = p.pPrSlots?.find((s) => s.tag === "rPr")?.xml ?? "";
    const markSays = (tag: string): boolean => new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${tag}[\\s/>]`).test(markSlot);
    const markFonts = /<(?:[A-Za-z_][\w.-]*:)?rFonts\b[^>]*>/.exec(markSlot)?.[0] ?? "";
    const markFace = (names: string): boolean => new RegExp(`(?:[A-Za-z_][\\w.-]*:)?(?:${names})\\s*=`).test(markFonts);
    for (const key of [...PARA_STYLE_KEYS, ...MODEL_ONLY_PARA_KEYS]) delete p[key];
    Object.assign(p, fields);
    p.styleId = styleId;
    p.styleName = styleId === undefined ? undefined : this.styles.get(styleId)?.name;
    if (!markSays("sz")) p.markSizePt = nextRuns.sizePt;
    if (!markSays("color")) p.markColor = nextRuns.color;
    if (!markFace("ascii|hAnsi|asciiTheme|hAnsiTheme")) p.markFont = nextRuns.font;
    if (!markFace("eastAsia|eastAsiaTheme")) p.markFontEastAsia = nextRuns.fontEastAsia;
    for (const run of p.inlines) {
      const charId = runStyleRef(run.rPrSlots);
      const charLent = charId === undefined ? undefined : this.runLent(charId, false);
      for (const key of RUN_STYLE_KEYS) {
        const direct = run.rPrSlots?.length
          ? slotStates(run.rPrSlots, RUN_SLOT_STATES[key])
          : !samePropValue(run.asRead?.[key], run[key]);
        if (direct) continue;
        const value = charLent?.[key] ?? nextRuns[key];
        (run as unknown as Record<string, unknown>)[key] = value;
        run.asRead = { ...run.asRead, [key]: value };
      }
    }
    p.asRead = { ...lent };
    return true;
  }

  /**
   * §17.7.2.24 `w:next`: the style a paragraph takes when Enter splits the one that names `styleId`. The split
   * keeps the paragraph's own direct properties and drops the ones the previous style was lending it, so the
   * new paragraph reads as a paragraph whose `w:pPr` says only `w:pStyle`.
   *
   * Only a style the paragraph *names* answers the question — a paragraph that names none falls back on the
   * default paragraph style for formatting (§17.7.2) but not for a next style. 授权 46 份的这条腿数清楚了：
   * `word/styles.xml`（连 glossary 那份）共 **623 枚样式、406 枚段落样式，其中 188 枚写 `w:next`，188 枚全是段落样式写的**；
   * 而 **46 份的默认段落样式（46 枚）没有一枚写 `w:next`**——这一句是关键，因为它就是「回车该落回哪里」的那条兜底：
   * 正文段若不点名样式（授权口径下 `word/document.xml` 的 7,430 段里有 **5,912 段不点名**）就走默认样式，而默认样式
   * 既然不指 next，第二段不会莫名其妙变成标题。作废语料上那组「668 of 2,378 写 next / 358 靠 basedOn 继承 / 63 of 86
   * 的默认段落样式写了 next / 16,783 of 30,322 段不点名 / 924 落进标题」属旧读数，其中默认样式那一句已被上面这组换掉，
   * basedOn 继承那 358 枚与「相邻两段样式改动 230 次里 0 次等于前一枚的 next」在授权口径上 **待重证**。
   * A heading's next, which is the case the feature is for, always
   * travels as an explicit `w:pStyle`.
   *
   * A `w:next` that names a style the part does not declare, or the style itself, says nothing a split could
   * apply, so it reads as no next at all.
   */
  nextStyleOf(styleId: string | undefined): string | undefined {
    if (styleId === undefined) return undefined;
    if (this.nextMemo.has(styleId)) return this.nextMemo.get(styleId);
    const seen = new Set<string>();
    let next: string | undefined;
    for (let cur: string | undefined = styleId; cur !== undefined && !seen.has(cur); cur = this.styles.get(cur)?.basedOn) {
      seen.add(cur);
      const stated = this.styles.get(cur)?.next;
      if (stated !== undefined && stated !== styleId && this.styles.has(stated)) {
        next = stated;
        break;
      }
    }
    this.nextMemo.set(styleId, next);
    return next;
  }

  /**
   * The paragraph Enter leaves behind when `previous`'s style names a §17.7.2.24 next style, or nothing when it
   * names none — the caller then copies the shape it had, which is the rule for a style that says no more.
   *
   * 后一段按**新样式的链**画，带上上一段自己写的直接格式（拆段不改直接格式），丢掉上一段样式链借给它的那些。
   * 自己写的那一套这样认：这一段现在显示的、它**旧的**样式链并没有说的那些项。读入快照 `asRead` 当不了参照——
   * 它存的是解析完样式链之后的那一套，文件自己写的 `w:ind w:left="720"` 与链借来的在里头是同一个值，对着它比
   * 只剩「用户后来改过的」，回车就把上一段自己写的缩进、行距整个忘掉（`.scratch/next-dom.ts` 在真页面上量到的
   * 正是这条）。剩下的口子只有一处：自己写的值恰好与旧链相同的那一项分不出来，回车后落回新样式的值。
   *
   * `asRead` 交的是**新链**那一份，于是写出侧对着它只把 `w:pStyle` 说出去——保存回来就是「这段只点名了一个样式」
   * 的形状，而不是把上一个样式的段落属性抄成这一段自己的。
   *
   * 空段上回车结束列表是同一条规则：上一段没有内容（只有空白文字，占位的标记一个都没有），新段就不该再占一个编号。
   */
  nextSplit(previous: Paragraph): Paragraph | undefined {
    const nextId = this.nextStyleOf(previous.styleId);
    if (nextId === undefined) return undefined;
    const lent: ParaStyle = { ...this.paragraphStyleOf(nextId) };
    const before = this.paragraphStyleOf(previous.styleId);
    const own = ownParaStyle(before, paraStyleOf(previous));
    delete own.styleId;
    const fields: ParaStyle = { ...lent, ...own };
    // 这八项不在 `PARA_STYLE_KEYS` 里（写出靠 `w:pPr` 的槽位带回），`ownParaStyle` 认不到它们，所以「上一段自己
    // 写的那一份」得按同一条判据再认一次：与旧链说的不一样，就是这一段自己写的，跟着走——回车不该把上一段自己写的
    // 保持规则或它自己那句「不吸网格」忘掉（授权口径里 226 枚 `w:snapToGrid w:val="0"` 是段落自己写的，
    // `.scratch/corpus-recheck2.mjs`）。`.scratch/next-dom.ts` 在真页面上量到的正是这一族丢掉的形状。
    // 「不一样」在这里按**引用**认，不是 `samePropValue`：框是个物件，而 `mergeParaBorders` 恰好在这一段一个边
    // 都没说的时候把链上那个物件原样交回来。于是引用相同＝「这个盒子是链的，新段该向自己的链要」，引用不同＝
    // 「这一段自己说过」。按内容比会把「自己说过、说的又与链相同」那一击判成借来的（本编辑器给一段加了链上也
    // 有的那条底线，回车后那一半就没了）；按引用比不会，且对四个布尔字段与 `shading` 那个字符串与原判据等价。
    // 件里 26 枚 `w:pBdr` 有 24 枚在 `word/styles.xml`、2 枚在页眉部件的段落，正文段落自己写的 0 枚
    // （`.scratch/pbdr-audience.ts`），所以这条判据量到的是**这一击**的形状，不是件里已有的形状。
    for (const key of MODEL_ONLY_PARA_KEYS) {
      const stated = previous[key];
      if (stated === undefined || stated === before?.[key]) continue;
      // 一次写这些字段中的一个，键是联合的，TS 只认三者交集那一格（`undefined`），所以按名写进去。判据还是上面
      // 那一条引用比较，与 `fields[key] = stated` 同义。
      Object.assign(fields, { [key]: stated });
    }
    // 「拆段不改直接格式」在这一条上有例外：`w:pageBreakBefore` 说的不是这一段的排版，是「这一段之前必须断一页」，
    // 跟着走就是每按一次回车多一页。新段落按它自己的样式画。
    delete fields.pageBreakBefore;
    // 空段上回车结束列表，是同一条规则：这一段没有内容，它就不该再占一个编号。
    if (previous.inlines.every((i) => i.text.trim() === "" && !isMark(i))) delete fields.numPr;
    return {
      kind: "paragraph",
      ...fields,
      styleId: nextId,
      styleName: this.styleDef(nextId)?.name,
      asRead: lent,
      inlines: [],
    };
  }

  /** The body-level section properties exactly as read — headers, footers and columns live here. */
  sectionPropertiesXml(): string {
    const xml = this.bodySectPr;
    if (!xml) return "";
    const tags = this.sectReferenceTags();
    return tags ? this.insertBandRefs(xml, tags) : xml;
  }

  /**
   * Whether a save owes page geometry the read file never had: true only after 「页面设置」 was used on a
   * document with no section bytes to edit.
   */
  pageSetupStated(): boolean {
    return this.pageSetupSeen;
  }

  /**
   * Applies 「页面设置」 to the section the page is drawn from. Which bytes that is depends on how the document
   * starts: one that opens with a section break hangs its `w:sectPr` off a paragraph, and one whose file never
   * spelled a section has no bytes to edit at all — only then does the geometry live in the model alone, and a
   * save writes the section the user asked for. A document with several sections has its **first** one changed,
   * because that is the page this app draws and the one the dialog showed numbers from.
   */
  setPageSetup(set: PageSetup): void {
    if (this.bandSectIsBody && this.bodySectPr) {
      this.bodySectPr = applyPageSetup(this.bodySectPr, set);
      this.sections = [sectionGeometry(this.bodySectPr), ...this.sections.slice(1)];
      return;
    }
    // The same paragraph `citeBandReference` looks for: the first block-level section break, which is the one
    // the opening page belongs to.
    for (const b of this.blocks) {
      if (b.kind !== "paragraph") continue;
      const at = b.pPrSlots?.findIndex((s) => s.tag === "sectPr") ?? -1;
      if (!b.pPrSlots || at < 0) continue;
      const patched = applyPageSetup(b.pPrSlots[at]!.xml, set);
      b.pPrSlots[at] = { tag: "sectPr", xml: patched };
      this.sections = [sectionGeometry(patched), ...this.sections.slice(1)];
      return;
    }
    const base = this.sections[0] ?? copySection(DEFAULT_SECTION);
    const box = pageBoxOf(set, base.pageWidthTwips, base.pageHeightTwips);
    this.sections = [
      {
        pageWidthTwips: box.w,
        pageHeightTwips: box.h,
        margins: {
          top: set.topTwips,
          right: set.rightTwips,
          bottom: set.bottomTwips,
          left: set.leftTwips,
          header: set.headerTwips ?? base.margins.header,
          footer: set.footerTwips ?? base.margins.footer,
          gutter: set.gutterTwips ?? base.margins.gutter,
        },
        // A section that is being built from nothing still has to keep the borders of whatever it was read
        // off — dropping them here would erase a page border as a side effect of changing a margin.
        pageBorders: base.pageBorders,
      },
      ...this.sections.slice(1),
    ];
    this.pageSetupSeen = true;
  }

  /**
   * Everything 「页面设置」 moves, so one step back puts the page where it was. The paragraph's own section break
   * needs no mention here — it travels inside `blocks`, which an undo step already clones — but the body's
   * `w:sectPr` and the geometry read off it sit outside that clone.
   */
  pageGeometryState(): PageGeometry {
    return { bodySectPr: this.bodySectPr, sections: structuredClone(this.sections), stated: this.pageSetupSeen };
  }

  restorePageGeometry(state: PageGeometry): void {
    this.bodySectPr = state.bodySectPr;
    this.sections = structuredClone(state.sections);
    this.pageSetupSeen = state.stated;
  }

  /**
   * Namespace prefixes the document root declared. A save writes them again: carried table property
   * text may name one of them, and a prefix that is not bound makes the part unreadable.
   */
  documentNamespaces(): Record<string, string> {
    return this.rootNamespaces;
  }

  /**
   * The prefixes the root listed in `mc:Ignorable`, as it wrote them. Markup compatibility is what lets a
   * reader that has never seen `w14:` skip the extension instead of failing on it, and a save that drops
   * the list turns a file that opened into one that will not.
   */
  documentIgnorable(): string {
    return this.rootIgnorable;
  }

  /**
   * The `w:document` start tag's attribute text, exactly as the file wrote it and leading with its first
   * space: ` xmlns:w="…" … mc:Ignorable="…"` — or `""` for a root that declared nothing, which is what a
   * document created here looks like. A save puts these bytes back rather than spelling a root of its own:
   * the list is the file's chosen prefix names *and* their order, and an attribute this model never read
   * would otherwise be dropped by the very act of rewriting the tag.
   */
  documentRootAttributes(): string {
    return this.rootOpenAttrs;
  }

  /**
   * Reference tags the model's header/footer still need inside the body's `sectPr`. Empty for a document
   * with more than one section: the body's `sectPr` is its **last** section, and a citation belongs to the
   * section the band sits on — `citeBandReference` puts it there instead.
   */
  sectReferenceTags(): string {
    if (!this.bandSectIsBody) return "";
    const tag = (kind: "header" | "footer", hf?: HeaderFooter): string => (hf && !this.sectHas[kind] ? this.bandRefTag(kind, hf.id) : "");
    return tag("header", this.header) + tag("footer", this.footer);
  }

  /** One band reference this session owes, spelled the way the files spell it (see {@link insertBandRefs}). */
  private bandRefTag(kind: "header" | "footer", id: string): string {
    const p = this.sectPrefix;
    return `<${p}:${kind}Reference r:id="${id}" ${p}:type="default"/>`;
  }

  /**
   * Splices band references into a section's `w:sectPr`. Element order means nothing to a reader — this model
   * matches a band on its `w:type`, never on its position (§17.10.3) — so the only question is what shape a save
   * hands back. 授权 46 份里 cite 多枚的 **6 节按 `w:type` 分档写（`first` 在前、`default` 其次、`even` 最后），
   * 同一档里页眉在页脚前**：形状是 `footer:default > footer:even`（2 节）、`header:default > footer:even`（2 节）、
   * `footer:first > header:default > footer:default`（2 节）（`.scratch/sect-ref-order.ts`）。这一条只量到 6 节，
   * 而且 6 节里没有任何一枚 `w:type="first"` 的页眉，所以「first 档里页眉页脚谁前谁后」是 **未取证**。
   *
   * The tags a save owes are always `default`, so they go after the `first` ones and after the `default`
   * 页眉 — which is also where a newly invented 页眉 lands relative to a 页脚 the section already cited
   * (`header:default > footer:default`). Both splices used to put the new tag straight after the opening tag,
   * which is the right place only for a section that cited nothing.
   */
  private insertBandRefs(sectXml: string, tags: string): string {
    const p = this.sectPrefix;
    const gt = sectXml.indexOf(">", sectXml.indexOf("<" + p + ":sectPr"));
    if (gt < 0) return sectXml;
    if (sectXml[gt - 1] === "/") return `${sectXml.slice(0, gt)}>${tags}</${p}:sectPr>`;
    const sortsBefore = new RegExp(`^\\s*<${p}:(?:(?:header|footer)Reference\\b[^>]*\\b${p}:type="first"|headerReference\\b[^>]*\\b${p}:type="default")[^>]*/>`);
    let at = gt + 1;
    for (;;) {
      const m = sortsBefore.exec(sectXml.slice(at));
      if (!m) break;
      at += m[0].length;
    }
    return `${sectXml.slice(0, at)}${tags}${sectXml.slice(at)}`;
  }

  /**
   * Cites a header/footer this session created from the `sectPr` that owns the band, when that is a
   * paragraph's section break rather than the body's. The reference goes into the first top-level paragraph
   * carrying one, which is the section the document starts with.
   */
  private citeBandReference(kind: "header" | "footer", id: string): void {
    const tag = this.bandRefTag(kind, id);
    for (const b of this.blocks) {
      if (b.kind !== "paragraph" || !b.pPrSlots?.length) continue;
      const at = b.pPrSlots.findIndex((s) => s.tag === "sectPr");
      if (at < 0) continue;
      const slot = b.pPrSlots[at]!;
      b.pPrSlots[at] = { tag: "sectPr", xml: this.insertBandRefs(slot.xml, tag) };
      this.sectHas = { ...this.sectHas, [kind]: true };
      return;
    }
  }

  /**
   * Rewrites the texts of an editable header/footer; rich parts refuse. A part that carries a page field
   * refuses too, because flat text cannot say where the field sits — taking the text alone would drop it.
   */
  setHeaderTexts(which: "header" | "footer", paragraphs: string[]): boolean {
    const hf = which === "header" ? this.header : this.footer;
    if (!hf?.editable) return false;
    if (hf.segments?.some((s) => s.some((x) => x.carry))) return false;
    return this.setHeaderSegments(which, paragraphs.map((t) => [{ text: t }]));
  }

  /** Rewrites an editable header/footer paragraph by paragraph, pieces and all; rich parts refuse. */
  setHeaderSegments(which: "header" | "footer", segments: HfSegment[][]): boolean {
    const hf = which === "header" ? this.header : this.footer;
    if (!hf) return false;
    return this.setBandSegments(which, hf.part, segments);
  }

  /**
   * The band one package path is. One part cited by several sections is one object (`bandByPart`), so the page
   * that shows it, the click that edits it, and the save that writes it all answer to the same holder — which is
   * what lets a change to a later section's 页眉 reach the bytes instead of only the model. A part
   * `ensureHeaderFooter` made this session is not in that map, so the document-level band answers for it.
   */
  bandOf(which: "header" | "footer", part: string): HeaderFooter | undefined {
    const held = this.bandByPart.get(part);
    if (held) return held;
    const own = which === "header" ? this.header : this.footer;
    return own?.part === part ? own : undefined;
  }

  /** Rewrites the band one part holds by its package path — not only the document-level pair. */
  setBandSegments(which: "header" | "footer", part: string, segments: HfSegment[][]): boolean {
    const hf = this.bandOf(which, part);
    if (!hf?.editable) return false;
    hf.segments = segments.map(coalesceHfSegments);
    hf.paragraphs = hf.segments.map(hfSegmentsText);
    return true;
  }

  /**
   * Puts a `PAGE` field in an editable header or footer, spelled as the files write it and carried as one
   * atom, so the text around it stays editable and a save gives the field back unchanged. `paragraph` counts
   * from the top, or from the end when negative; `offset` counts the characters of text before it goes in;
   * `part` names the band the page shows, and leaving it out means the document-level one.
   */
  insertPageNumber(which: "header" | "footer", paragraph = -1, offset?: number, part?: string): boolean {
    const key = part ?? (which === "header" ? this.header?.part : this.footer?.part);
    if (!key) return false;
    const hf = this.bandOf(which, key);
    if (!hf?.editable) return false;
    const segs = (hf.segments ?? hf.paragraphs.map((t) => [{ text: t }])).map(copyHfSegments);
    const at = paragraph < 0 ? segs.length + paragraph : paragraph;
    const target = segs[at];
    if (!target) return false;
    const prefix = /^<([\w.-]+):(?:hdr|ftr)/.exec(hf.source)?.[1] ?? "w";
    target.splice(0, target.length, ...putHfCarry(target, offset, { key: this.carryKey++, tag: "fldChar", xml: pageFieldXml("PAGE", prefix), show: "页码", field: "PAGE" }));
    return this.setBandSegments(which, key, segs);
  }

  /** Adds an empty header/footer to a document that had none, wiring rels and sectPr. */
  ensureHeaderFooter(which: "header" | "footer"): HeaderFooter | undefined {
    const existing = which === "header" ? this.header : this.footer;
    if (existing) return existing;
    const kind = which;
    const taken = new Set(this.bodyRels.map((r) => r.target));
    let n = 1;
    while (taken.has(`${kind}${n}.xml`)) n++;
    const highest = this.bodyRels.reduce((max, r) => Math.max(max, Number(/(\d+)$/.exec(r.id)?.[1] ?? 0)), 0);
    const id = `rId${highest + 1}`;
    this.bodyRels.push({ id, type: `${REL_R}/${kind}`, target: `${kind}${n}.xml`, targetMode: "Internal", path: `word/${kind}${n}.xml` });
    const hf: HeaderFooter = { part: `word/${kind}${n}.xml`, id, editable: true, paragraphs: [""], baseline: ["__new__"], source: "" };
    if (which === "header") this.header = hf;
    else this.footer = hf;
    // A document that opens with a section break binds its header from that paragraph's `sectPr`; the
    // body's own is the last section and citing from there would start the header partway through.
    if (!this.bandSectIsBody) this.citeBandReference(kind, id);
    return hf;
  }

  /**
   * Every band some page shows, once per part: the document-level pair first, then the parts the sections
   * cite. Two things are decided from this list and they must agree — what counts as an edit (`fingerprint`)
   * and what a save may rewrite (`headerFooterParts`). A later section's 页脚 left out of either one is an
   * edit that never happened: the model holds the new words, `edited` says nothing changed, and the save hands
   * back the file as it came.
   */
  private bands(): HeaderFooter[] {
    const out: HeaderFooter[] = [];
    const seen = new Set<string>();
    for (const hf of [this.header, this.footer, ...this.bandByPart.values()]) {
      if (!hf || seen.has(hf.part)) continue;
      seen.add(hf.part);
      out.push(hf);
    }
    return out;
  }

  /**
   * Header/footer parts with their regenerated XML; `dirty` ones a save must write.
   *
   * Every band a section shows is in here, not only the document-level pair (#157): a page whose 页脚 belongs
   * to a later section edits that section's part, and a save that could not see the part would drop the edit on
   * the floor. A part nobody touched reports `dirty: false`, so `writeDocx` leaves its bytes alone — the
   * unchanged-file-stays-unchanged rule is per part, not per document.
   */
  headerFooterParts(): { part: string; xml: string; dirty: boolean }[] {
    return this.bands().map((hf) => ({ part: hf.part, xml: headerFooterXml(hf), dirty: hf.segments ? !sameHfSegments(hf.segments, hf.baselineSegments ?? []) : JSON.stringify(hf.paragraphs) !== JSON.stringify(hf.baseline) }));
  }

  // -------------------------------------------------------------- revisions

  /** Tracked changes still waiting for a decision, newest authorship first. */
  revisions(): Revision[] {
    const out: Revision[] = [];
    const seen = new Set<string>();
    this.eachParagraph((p) => {
      for (const run of p.inlines) {
        const rev = run.rev;
        if (!rev) continue;
        const key = `${rev.kind}:${rev.id}`;
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(rev);
      }
    });
    return out;
  }

  /**
   * The paragraph-mark revisions (`w:pPr/w:rPr/w:ins`, `w:del`) still waiting, one entry per `w:id`.
   * Word's pane lists these as 段落改动 next to the text ones; `revisionEntries()` is the one ordered
   * collection that holds them together with the rest.
   */
  markRevisions(): Revision[] {
    const out: Revision[] = [];
    const seen = new Set<string>();
    this.eachParagraph((p) => {
      const mark = p.markRev;
      if (!mark) return;
      const key = `${mark.kind}:${mark.id}`;
      if (seen.has(key)) return;
      seen.add(key);
      out.push(mark);
    });
    return out;
  }

  /** The paragraph property changes (`w:pPrChange`) still waiting, one entry per `w:id`. */
  paragraphFormatRevisions(): ParagraphFormatRevision[] {
    const out: ParagraphFormatRevision[] = [];
    const seen = new Set<number>();
    this.eachParagraph((p) => {
      const fmt = p.revFormat;
      if (!fmt || seen.has(fmt.id)) return;
      seen.add(fmt.id);
      out.push(fmt);
    });
    return out;
  }

  /**
   * The highest `w:id` a change in this document carries, or 0 when nothing is revised. The editor
   * hands out ids above it: a revision it signs shares one number space with the revisions read from
   * the file, and two changes with the same `w:id` would be decided together by accepting either.
   */
  maxRevisionId(): number {
    let max = 0;
    this.eachParagraph((p) => {
      max = Math.max(max, p.markRev?.id ?? 0, p.revFormat?.id ?? 0);
      for (const run of p.inlines) max = Math.max(max, run.rev?.id ?? 0, run.revFormat?.id ?? 0);
    });
    // A row's own signature lives on its cells, where the paragraph walk cannot reach it, and it shares
    // the one `w:id` number space with everything else — so do the records on a table, a row or a cell.
    this.eachTable((t) => {
      max = Math.max(max, t.revFormat?.id ?? 0);
      for (const row of t.rows) {
        const info = row[0]?.row;
        max = Math.max(max, info?.rev?.id ?? 0, info?.revFormat?.id ?? 0);
        for (const cell of row) for (const rec of cell.revs ?? []) max = Math.max(max, rec.id);
      }
    });
    return max;
  }

  /**
   * Every paragraph inside these blocks that is a paragraph and nothing inside it, in the order a reader
   * meets them — the place a pane can scroll a table's change to.
   */
  private firstParagraphIn(blocks: Block[]): Paragraph | undefined {
    for (const b of blocks) {
      if (b.kind === "paragraph") return b;
      if (b.kind !== "table") continue;
      for (const row of b.rows) for (const cell of row) {
        const inner = this.firstParagraphIn(cell.paragraphs);
        if (inner) return inner;
      }
    }
    return undefined;
  }

  /** What a reader would see in these blocks, table cells included — the text a pane names a change by. */
  private textOfBlocks(blocks: Block[]): string {
    return blocks
      .map((b) =>
        b.kind === "paragraph"
          ? b.inlines.map((i) => i.text).join("")
          : b.kind === "table"
            ? b.rows.flat().map((c) => this.textOfBlocks(c.paragraphs)).join(" ")
            : b.text,
      )
      .join(" ");
  }

  /**
   * The changes waiting on a table part rather than on characters, keyed on the paragraph a pane can scroll
   * to: a table's own record at the first paragraph of the table, a row's at the row's first, a cell's at the
   * cell's first. `text` is whatever the part holds, because a row that was added or a width that was changed
   * has no characters of its own to show. One record covers every cell of a revised column — the file gives
   * them the same `w:id` — and `revisionEntries()` joins those into one row of the list.
   */
  private tablePending(): Map<Paragraph, { key: string; entry: Omit<RevisionEntry, "index" | "para" | "begin" | "end"> }[]> {
    const out = new Map<Paragraph, { key: string; entry: Omit<RevisionEntry, "index" | "para" | "begin" | "end"> }[]>();
    const put = (at: Paragraph | undefined, key: string, entry: Omit<RevisionEntry, "index" | "para" | "begin" | "end">): void => {
      if (!at) return;
      const held = out.get(at);
      if (held) held.push({ key, entry });
      else out.set(at, [{ key, entry }]);
    };
    const walk = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind !== "table") continue;
        const cells = b.rows.flat();
        if (b.revFormat) {
          const rec = b.revFormat;
          put(this.firstParagraphIn(cells.flatMap((c) => c.paragraphs)), `table:${rec.id}`, {
            id: rec.id,
            type: "table-format",
            author: rec.author,
            date: rec.date,
            text: this.textOfBlocks(cells.flatMap((c) => c.paragraphs)),
          });
        }
        for (const row of b.rows) {
          const info = row[0]?.row;
          const inside = row.flatMap((c) => c.paragraphs);
          const anchor = this.firstParagraphIn(inside);
          const text = this.textOfBlocks(inside);
          // The row leads what sits inside it: a reviewer reads "added a row" first, then the text in it.
          if (info?.rev) {
            const rev = info.rev;
            put(anchor, `row:${rev.kind}:${rev.id}`, {
              id: rev.id,
              type: rev.kind === "del" ? "row-delete" : "row-insert",
              author: rev.author,
              date: rev.date,
              text,
            });
          }
          if (info?.revFormat) {
            const rec = info.revFormat;
            put(anchor, `row-format:${rec.id}`, {
              id: rec.id,
              type: "row-format",
              author: rec.author,
              date: rec.date,
              text,
            });
          }
          for (const cell of row) {
            for (const rec of cell.revs ?? [])
              put(this.firstParagraphIn(cell.paragraphs), `cell:${rec.tag}:${rec.id}`, {
                id: rec.id,
                type: CELL_REVISION_TYPES[rec.tag],
                author: rec.author,
                date: rec.date,
                text: this.textOfBlocks(cell.paragraphs),
              });
          }
        }
        for (const cell of cells) walk(cell.paragraphs);
      }
    };
    walk(this.blocks);
    return out;
  }

  /**
   * Every change waiting for a decision, in the order it sits in the document and numbered from 1.
   * A table part's own record leads, then within a paragraph the runs come first, then the paragraph's
   * own property record, then its mark — the break is the last thing in a paragraph. One revision can
   * run across several paragraphs and keeps a single row, placed where it first appears, because the
   * file gives it one `w:id`.
   */
  revisionEntries(): RevisionEntry[] {
    const out: RevisionEntry[] = [];
    const by = new Map<string, RevisionEntry>();
    const put = (
      key: string,
      text: string,
      para: number,
      begin: number,
      end: number,
      make: () => Omit<RevisionEntry, "index" | "para" | "begin" | "end">,
    ): void => {
      const held = by.get(key);
      if (held) {
        held.text += text;
        // Several revised runs in one paragraph are one range of one `w:id`; where the same revision
        // continues into the next paragraph the row keeps the place it began at.
        if (held.para === para) held.end = end;
        return;
      }
      const entry: RevisionEntry = { index: out.length + 1, para, begin, end, ...make() };
      by.set(key, entry);
      out.push(entry);
    };
    let at = 0;
    const pending = this.tablePending();
    this.eachParagraph((p) => {
      const whole = p.inlines.map((i) => i.text).join("");
      // A table part's record is the change the part itself underwent, so it leads the text inside it.
      for (const held of pending.get(p) ?? []) put(held.key, held.entry.text, at, 0, 0, () => held.entry);
      let run = 0;
      for (const r of p.inlines) {
        const from = run;
        run += r.text.length;
        const rev = r.rev;
        if (rev) put(`${rev.kind}:${rev.id}`, r.text, at, from, run, () => ({ id: rev.id, type: entryType(rev), author: rev.author, date: rev.date, text: r.text }));
        const fmt = r.revFormat;
        if (fmt) put(`format:${fmt.id}`, r.text, at, from, run, () => ({ id: fmt.id, type: "run-format", author: fmt.author, date: fmt.date, text: r.text }));
      }
      const pfmt = p.revFormat;
      if (pfmt) put(`paragraph-format:${pfmt.id}`, whole, at, 0, whole.length, () => ({ id: pfmt.id, type: "paragraph-format", author: pfmt.author, date: pfmt.date, text: whole }));
      const mark = p.markRev;
      if (mark) put(`mark:${mark.kind}:${mark.id}`, "¶", at, whole.length, whole.length, () => ({ id: mark.id, type: mark.move ? entryType(mark) : mark.kind === "ins" ? "insert-mark" : "delete-mark", author: mark.author, date: mark.date, text: "¶" }));
      at += 1;
    });
    return out;
  }

  /** The paragraphs in the order `revisionEntries()` walks them, table cells included: a row's `para`
   *  is its index in this list. */
  paragraphs(): Paragraph[] {
    const out: Paragraph[] = [];
    this.eachParagraph((p) => out.push(p));
    return out;
  }

  /**
   * Which row of the list covers a place in the document — WPS' `getRevisionIndexByBegin`, which takes
   * a position and returns the 修订索引 of the change there, counting from 1. Where two changes overlap
   * it answers the innermost one, the range that ends closest to where it began, because that is the
   * change a reviewer clicking on the marked text means. Undefined when nothing is waiting there.
   */
  revisionIndexAt(at: Paragraph, offset: number): number | undefined {
    const list = this.paragraphs();
    let best: RevisionEntry | undefined;
    for (const e of this.revisionEntries()) {
      if (list[e.para] !== at || offset < e.begin || offset > e.end) continue;
      if (!best || e.end - e.begin < best.end - best.begin) best = e;
    }
    return best?.index;
  }

  /** 接受 / 拒绝 the `index`-th change counting the way `revisionEntries()` lists them. */
  resolveRevisionAt(outcome: "accept" | "reject", index: number): boolean {
    const entry = this.revisionEntries()[index - 1];
    if (!entry) return false;
    return this.resolveRevisions(outcome, entry.id) > 0;
  }

  /**
   * 按作者接受 / 拒绝: every change one reviewer signed, and nothing anybody else signed. Authorship
   * is compared exactly as the file wrote it, because a revision is signed rather than editable.
   */
  resolveRevisionsByAuthor(outcome: "accept" | "reject", author: string): number {
    const ids = [...new Set(this.revisionEntries().filter((e) => e.author === author).map((e) => e.id))];
    return ids.reduce((done, id) => done + (this.resolveRevisions(outcome, id) > 0 ? 1 : 0), 0);
  }

  /**
   * 接受 / 拒绝修订, Word's own two outcomes: accepting lets an insertion stand and a deletion
   * take effect, refusing puts the deleted text back and takes the inserted text out. A formatting
   * revision decides the same way — accepting keeps the new properties, refusing puts the recorded
   * ones back — and so does a row's own record, where the change is a row of the table rather than
   * characters: the row stands when an insertion is accepted or a deletion refused, and goes the
   * other two ways. `id` decides one revision only, which is what the mark under the caret needs.
   * Returns the paragraphs and rows touched.
   */
  resolveRevisions(outcome: "accept" | "reject", id?: number): number {
    let touched = this.resolveTableParts(outcome, id);
    touched += this.resolveRows(outcome, id);
    touched += this.resolveParagraphs(outcome, id);
    this.eachParagraph((p) => {
      const next: Inline[] = [];
      let changed = false;
      for (const run of p.inlines) {
        const rev = run.rev;
        const fmt = run.revFormat;
        const decidesRev = !!rev && (id === undefined || rev.id === id);
        const decidesFmt = !!fmt && (id === undefined || fmt.id === id);
        if (!decidesRev && !decidesFmt) {
          next.push(run);
          continue;
        }
        changed = true;
        const rest: Inline = decidesFmt && outcome === "reject" ? restoreStyle(run, fmt!) : { ...run };
        if (decidesFmt) delete rest.revFormat;
        if (decidesRev) {
          const keepsText = outcome === "accept" ? rev!.kind === "ins" : rev!.kind === "del";
          if (!keepsText) continue;
          delete rest.rev;
        }
        next.push(rest);
      }
      if (changed) {
        p.inlines = next;
        touched += 1;
      }
    });
    return touched;
  }

  /**
   * Deciding a row's own record (`w:trPr/w:ins`, `w:trPr/w:del`). An inserted row stands when it is
   * accepted and goes when it is refused; a deleted row goes when it is accepted and stays when it is
   * refused — and its cells keep whatever marks the file gave their text, because those are separate
   * signatures with separate `w:id`s that the reviewer decides for themselves. A table whose last row
   * goes is no longer a table: an empty `w:tbl` is not something Word writes or reads back, so the slot
   * keeps an empty paragraph instead of nothing — the same shape 删除表格 leaves, and the one a block list
   * needs to have a line the caret can sit on.
   */
  private resolveRows(outcome: "accept" | "reject", id?: number): number {
    let touched = 0;
    const walk = (blocks: Block[]): Block[] =>
      blocks.flatMap((b): Block[] => {
        if (b.kind !== "table") return [b];
        const stays: TableCell[][] = [];
        for (const row of b.rows) {
          const rev = row[0]?.row?.rev;
          if (!rev || (id !== undefined && rev.id !== id)) {
            stays.push(row);
            continue;
          }
          touched += 1;
          if (outcome === "accept" ? rev.kind === "ins" : rev.kind === "del") {
            // The row's cells all hold this one object, so clearing the record once clears it for the row.
            delete row[0]!.row!.rev;
            stays.push(row);
          }
        }
        b.rows = stays;
        b.columns = Math.max(1, ...stays.map((r) => r.length));
        // A table whose last row went is no longer a table: an empty `w:tbl` is not something Word writes.
        if (!stays.length) return [{ kind: "paragraph", inlines: [] }];
        for (const row of stays) for (const cell of row) cell.paragraphs = walk(cell.paragraphs);
        return [b];
      });
    this.blocks = walk(this.blocks);
    return touched;
  }

  /**
   * Deciding a record on a table part rather than on characters. `w:tblPrChange`, `w:trPrChange` and
   * `w:tcPrChange` are property sets: accepting lets the set as it now stands stay and refusing puts the
   * recorded one back — either done by rewriting the text the model carries for that set, so a property the
   * model has no field for comes through the decision unchanged. `w:cellMerge` puts back only the merge it
   * named. `w:cellIns` and `w:cellDel` decide whether the cell exists: it stays when an insertion is accepted
   * or a removal refused, and the other two ways depend on how far the record reaches — one `w:id` on every
   * cell of a column is OOXML signing a revised *column*, so the grid column goes with it, while a record on
   * fewer cells names only those and takes only them, because dropping the column would remove cells nobody
   * revised. A column only goes if the table has another one to keep, and a cell only leaves a row that has
   * another cell to keep: marks that would empty the grid or the row are left pending, because a table whose
   * removal was accepted and whose cells are still there is a worse answer than one that says no.
   * Returns the tables, rows and cells touched.
   */
  private resolveTableParts(outcome: "accept" | "reject", id?: number): number {
    let touched = 0;
    /** The cells whose record says they are not there, with the grid column each sits in. */
    const goers: { t: Table; row: TableCell[]; cell: TableCell; at: number; rec: CellRevision }[] = [];
    /** 拒绝时把某格的宽度摆回去过的那一列：`w:tblGrid` 没有记录族，得由该列的格反述（见 `restateColumnFromCells`）。 */
    const widened = new Map<Table, Set<number>>();
    this.eachTable((t) => {
      const tableRec = t.revFormat;
      if (tableRec && (id === undefined || tableRec.id === id)) {
        if (outcome === "reject") restoreTableStyle(t, tableRec.prev);
        if (t.tblPrXml) t.tblPrXml = decideSetXml(t.tblPrXml, tableRec, outcome) || undefined;
        delete t.revFormat;
        touched += 1;
      }
      for (const row of t.rows) {
        const info = row[0]?.row;
        const rowRec = info?.revFormat;
        if (info && rowRec && (id === undefined || rowRec.id === id)) {
          if (outcome === "reject") restoreRowStyle(info, rowRec.prev);
          const left = decideSlots(info.slots ?? [], rowRec, outcome);
          if (left.length) info.slots = left;
          else delete info.slots;
          delete info.revFormat;
          touched += 1;
        }
        for (const cell of row) {
          const mine = (cell.revs ?? []).filter((r) => id === undefined || r.id === id);
          if (!mine.length) continue;
          const settled: CellRevision[] = [];
          for (const rec of mine) {
            // A record of the cell's place decides whether the cell is there; one of its properties does not.
            const properties = isPropertyRecord(rec.tag);
            const goes = (rec.tag === "cellIns" && outcome === "reject") || (rec.tag === "cellDel" && outcome === "accept");
            if (goes) {
              goers.push({ t, row, cell, at: columnStarts(row)[row.indexOf(cell)] ?? 0, rec });
              continue;
            }
            touched += 1;
            settled.push(rec);
            // A place-record has no properties of its own to put back, so it only ever leaves the set.
            if (cell.tcPrXml) cell.tcPrXml = decideSetXml(cell.tcPrXml, rec, properties ? outcome : "accept") || undefined;
          }
          // 一格挂着多条记录时一次摆完，按文档顺序先说的那句算（见 `restoreCellRecords`）。
          if (outcome === "reject") restoreCellRecords(cell, settled.filter((rec) => isPropertyRecord(rec.tag)));
          if (outcome === "reject" && settled.some((rec) => rec.tag === "tcPrChange" && rec.prev.widthTwips !== undefined)) {
            const at = columnStarts(row)[row.indexOf(cell)];
            if (at !== undefined) {
              const cols = widened.get(t) ?? new Set<number>();
              cols.add(at);
              widened.set(t, cols);
            }
          }
          if (settled.length) {
            cell.revs = cell.revs!.filter((r) => !settled.includes(r));
            if (!cell.revs.length) delete cell.revs;
          }
        }
      }
    });
    /** A cell that is gone — with its column or on its own — must stop offering the record that sent it. */
    const settle = (cell: TableCell, rec: CellRevision): void => {
      if (cell.revs?.includes(rec)) {
        cell.revs = cell.revs.filter((r) => r !== rec);
        if (!cell.revs.length) delete cell.revs;
      }
      if (cell.tcPrXml) cell.tcPrXml = decideSetXml(cell.tcPrXml, rec, "accept") || undefined;
    };
    for (const t of new Set(goers.map((g) => g.t))) {
      const mine = goers.filter((g) => g.t === t);
      const drops = mine.filter((g) => coversColumn(t, g.at, g.rec));
      const columns = [...new Set(drops.map((g) => g.at))];
      let changed = 0;
      // A column only goes while the table has another one to keep: a mark that would empty the grid is left
      // pending, because a table whose removal was accepted and whose cells are still there is a worse answer
      // than one that says no.
      if (columns.length < gridColumns(t)) {
        // Right to left, so the columns still to come keep the places the walk wrote down for them.
        for (const at of columns.sort((a, b) => b - a)) {
          if (gridColumns(t) === 1) break;
          deleteColumn(t, at, { row: 0, col: 0 });
          for (const g of drops.filter((x) => x.at === at)) {
            changed += 1;
            // Most of these cells went with the column; one that merely lost a span of it is still here.
            settle(g.cell, g.rec);
          }
        }
      }
      for (const g of mine.filter((x) => !drops.includes(x))) {
        const place = g.row.indexOf(g.cell);
        // Already taken by a dropped column: settled, and there is nothing left to remove.
        if (place < 0) settle(g.cell, g.rec);
        // A row has to keep at least one cell, so a mark on a one-cell row that is not a whole column waits.
        else if (removeCell(g.row, place)) {
          settle(g.cell, g.rec);
          changed += 1;
        }
      }
      if (changed) {
        repairMerges(t);
        t.columns = gridColumns(t);
        touched += changed;
      }
    }
    // 拒绝把哪一列的格摆回了改前的宽度，就把那一列的网格数也摆回去——件里那一枚 `w:gridCol` 没有记录族，
    // 网格只是那份数的镜像；`restateColumnFromCells` 只在整列说得出同一个数时动，说不出就不猜。
    for (const [t, cols] of widened) for (const at of cols) restateColumnFromCells(t, at);
    return touched;
  }

  /**
   * Deciding the paragraph-level revisions: a `w:pPrChange` puts the properties back, and a mark
   * revision adds or removes the break that ends the paragraph it sits on. A break that goes takes
   * the following paragraph's content with it, because the two were never more than one paragraph
   * before somebody revised the boundary between them.
   */
  private resolveParagraphs(outcome: "accept" | "reject", id?: number): number {
    let touched = 0;
    const gone = new Set<Paragraph>();
    for (const group of this.paragraphGroups()) {
      for (let i = 0; i < group.length; i++) {
        const p = group[i];
        if (gone.has(p)) continue;
        const fmt = p.revFormat;
        if (fmt && (id === undefined || fmt.id === id)) {
          if (outcome === "reject") this.restoreParaStyle(p, fmt);
          delete p.revFormat;
          touched += 1;
        }
        // Two Enters in a row leave two marks, so folding one paragraph away can uncover the next.
        for (;;) {
          const mark = p.markRev;
          if (!mark || (id !== undefined && mark.id !== id)) break;
          // An inserted break stands when it is accepted and goes when it is refused; a deleted one is
          // the other way round. Either way the record is decided, and a mark with nothing to join
          // simply stops being pending.
          const joins = outcome === "accept" ? mark.kind === "del" : mark.kind === "ins";
          delete p.markRev;
          touched += 1;
          if (!joins) break;
          let k = i + 1;
          while (k < group.length && gone.has(group[k])) k += 1;
          const neighbour = k < group.length ? group[k] : undefined;
          if (!neighbour) break;
          p.inlines.push(...neighbour.inlines);
          p.markRev = neighbour.markRev;
          gone.add(neighbour);
          i = k;
        }
      }
    }
    if (gone.size) {
      this.blocks = this.blocks.filter((b) => b.kind !== "paragraph" || !gone.has(b));
      for (const b of this.blocks) {
        if (b.kind !== "table") continue;
        for (const row of b.rows)
          for (const cell of row)
            cell.paragraphs = cell.paragraphs.filter((p) => p.kind !== "paragraph" || !gone.has(p));
      }
    }
    return touched;
  }

  /**
   * The paragraphs that can be joined with a neighbour: each stretch of body paragraphs, and each table
   * cell's. A paragraph separated by a table or by a cell wall has no neighbour to merge into, so those
   * boundaries are where the groups end.
   */
  private paragraphGroups(): Paragraph[][] {
    const out: Paragraph[][] = [];
    const walk = (blocks: Block[]): void => {
      let run: Paragraph[] = [];
      const flush = (): void => {
        if (run.length) out.push(run);
        run = [];
      };
      for (const b of blocks) {
        if (b.kind === "paragraph") {
          run.push(b);
          continue;
        }
        // A table ends the run of paragraphs around it, and each cell starts a run of its own: text on
        // either side of a table was never one paragraph, however the file lays them out. A container the
        // model only carries has no paragraphs to join either.
        flush();
        if (b.kind === "table") for (const row of b.rows) for (const cell of row) walk(cell.paragraphs);
      }
      flush();
    };
    walk(this.blocks);
    return out;
  }

  /**
   * Putting a paragraph back the properties its `w:pPrChange` recorded. The record holds the whole
   * previous set, so a property it does not mention was not set — and alignment, indent, the three page-keeping
   * rules, the grid switch, the box and the shading come from the paragraph's style when the record leaves them
   * out, which is where they belonged before.
   */
  private restoreParaStyle(p: Paragraph, fmt: ParagraphFormatRevision): void {
    const prev = fmt.prev;
    for (const key of [...PARA_STYLE_KEYS, ...MODEL_ONLY_PARA_KEYS]) delete p[key];
    Object.assign(p, prev);
    p.styleName = p.styleId ? this.styles.get(p.styleId)?.name : undefined;
    if (fmt.prevSlots?.length) {
      // The record wrote more than the model has a field for, so the paragraph takes its set as written;
      // syncing the slots from fields would re-spell the named ones and delete the rest.
      p.pPrSlots = fmt.prevSlots;
    } else {
      // The record spells the properties the paragraph owns, so the owned children of its `w:pPr` follow the
      // restored fields before anything the new style lends is folded back in.
      syncParaSlots(p);
    }
    const inherited = this.paraLent(p.styleId);
    p.align = p.align ?? inherited?.align;
    // 缩进六枚按读的时候同一句话合起来：`prev` 带回来的是段落自己那一句（上面 `syncParaSlots` 那句说的就是它），
    // 所以「段落只要用任一枚答过这一件事，链上的另一枚就不许叠上来」在这里同样成立（见 `indResolve`）。
    Object.assign(p, indResolve(p, inherited));
    // The spacing goes back the same way it was read: a record that left `w:spacing` out returns the paragraph
    // to whatever its style lends, and one that stated a zero keeps that zero.
    p.spacingBefore = p.spacingBefore ?? inherited?.spacingBefore;
    p.spacingAfter = p.spacingAfter ?? inherited?.spacingAfter;
    p.lineSpacing = p.lineSpacing ?? inherited?.lineSpacing;
    p.lineSpacingRule = p.lineSpacingRule ?? inherited?.lineSpacingRule;
    // Numbering comes back as one property, the way it was read: a record that leaves `w:numPr` out hands the
    // paragraph back the list its style names, and one that wrote `w:numId w:val="0"` keeps that cancellation.
    p.numPr = p.numPr ?? inherited?.numPr;
    // Same for the outline level: a record that leaves `w:outlineLvl` out returns the paragraph to the level its
    // style lends, and one that wrote `w:val="9"` keeps that — 9 is body text (§17.3.2.26), so the paragraph has
    // taken back a level the chain was lending it.
    p.outlineLevel = p.outlineLevel ?? inherited?.outlineLevel;
    // 那七项模型有字段、但 `ownParaStyle` 认不到的（三条保持规则、「对齐到网格」、「允许西文在单词中间换行」、框与底纹）走同一条：记录说了
    // 就按记录的，没说就把手里的清掉、把样式链借的那份还回来。留着不放的话，被否决的那一句还挂在字段上——画错，
    // 下一击 `syncParaSlots` 还会把它替段落写进件里。
    p.keepNext = p.keepNext ?? inherited?.keepNext;
    p.keepLines = p.keepLines ?? inherited?.keepLines;
    p.widowControl = p.widowControl ?? inherited?.widowControl;
    p.snapToGrid = p.snapToGrid ?? inherited?.snapToGrid;
    p.wordWrap = p.wordWrap ?? inherited?.wordWrap;
    // 框是逐边的物件，链上那份按引用还回来，写侧才仍然认得出「这一段自己没说」；底纹那枚 token（`auto`）与颜色
    // 一样是说过的一句话，不是 `undefined`，所以 `??` 就够。
    p.borders = p.borders ?? inherited?.borders;
    p.shading = p.shading ?? inherited?.shading;
    // 退回完之后，这一段重新有一份「它当时显示的是什么」的参照：写出侧对着它判哪一句是用户后来改的口。
    p.asRead = paraReadSnapshot(p);
  }

  /** Tracked formatting changes on runs (`w:rPrChange`) still waiting for a decision, one entry per `w:id`. */
  formatRevisions(): FormatRevision[] {
    const out: FormatRevision[] = [];
    const seen = new Set<number>();
    this.eachParagraph((p) => {
      for (const run of p.inlines) {
        const fmt = run.revFormat;
        if (!fmt || seen.has(fmt.id)) continue;
        seen.add(fmt.id);
        out.push(fmt);
      }
    });
    return out;
  }

  /** Every paragraph in the body, including the ones inside table cells. */
  private eachParagraph(fn: (p: Paragraph) => void): void {
    const walk = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind === "paragraph") fn(b);
        else if (b.kind === "table") for (const row of b.rows) for (const cell of row) walk(cell.paragraphs);
      }
    };
    walk(this.blocks);
  }

  /** Every table in the document, including the ones nested inside a cell, outermost first. */
  private eachTable(fn: (t: Table) => void): void {
    const walk = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind !== "table") continue;
        fn(b);
        for (const row of b.rows) for (const cell of row) walk(cell.paragraphs);
      }
    };
    walk(this.blocks);
  }

  /**
   * 「显示原文」那一趟摆的那张脸：把挂着属性改动记录的五族（run / 段落 / 表 / 行 / 格）各按记录里的
   * `prev` 复制一份出来给渲染器，**模型本身一个字不动**——记录照旧挂着，写回仍写现值，接受/拒绝的
   * 清单也照旧数得到它。摆的是「拒绝」用的同一批 `restore*`：「按记录把这一整套换掉」只许有一处说法，
   * 视图自己再拼一遍 prev→字段，两份迟早各说各话。
   * 复制走 `structuredClone`，与编辑器存历史快照（`doc-edit.ts` 的 `snapshot()`）同一条路；对嵌套表就是
   * 连着格里的段落一起复制，所以递归里改的是副本。
   */
  originalView<T extends Block>(blocks: T[]): T[] {
    const out = structuredClone(blocks);
    const walk = (list: Block[]): void => {
      for (const b of list) {
        if (b.kind === "paragraph") {
          // 段落这一族与「拒绝」共用 `restoreParaStyle`，所以样式链那半边的合并（`paraLent` 借的对齐、缩进、
          // 段距）也在同一条路上重算一遍，不是把字段扣上去就完。
          if (b.revFormat) this.restoreParaStyle(b, b.revFormat);
          b.inlines = b.inlines.map((i) => (i.revFormat ? restoreStyle(i, i.revFormat) : i));
        } else if (b.kind === "table") {
          if (b.revFormat) restoreTableStyle(b, b.revFormat.prev);
          /** 这一张表里，有格的宽度被记录摆回改前的那些列——网格得跟着反述一次（同 `resolveTableParts`）。 */
          const widened = new Set<number>();
          for (const row of b.rows) {
            const info = row[0]?.row;
            if (info?.revFormat) restoreRowStyle(info, info.revFormat.prev);
            for (const cell of row) {
              const records = (cell.revs ?? []).filter((rec) => isPropertyRecord(rec.tag));
              restoreCellRecords(cell, records);
              if (records.some((rec) => rec.tag === "tcPrChange" && rec.prev.widthTwips !== undefined)) {
                const at = columnStarts(row)[row.indexOf(cell)];
                if (at !== undefined) widened.add(at);
              }
              walk(cell.paragraphs);
            }
          }
          for (const at of widened) restateColumnFromCells(b, at);
        }
      }
    };
    walk(out);
    return out;
  }

  /**
   * The document's headings in reading order: `level` is `headingLevel`, so 1 is the first heading level down
   * to 9, and body text is not in the list at all. Order is the same walk `renderBlocks` takes — a table's
   * cells come in row order, nested tables inside them — which is why a pane can address a row by counting
   * `[data-outline]` elements in the page instead of learning where the paragraph sits in the model.
   *
   * A heading the model carries as source bytes rather than paragraphs has no level to report and is not here;
   * the page shows its text, and the pane lists what the model can point at.
   */
  outline(): { level: number; text: string }[] {
    const out: { level: number; text: string }[] = [];
    this.eachParagraph((p) => {
      const level = headingLevel(p);
      if (level === undefined) return;
      const text = p.inlines.map((i) => i.text).join("").trim();
      out.push({ level, text });
    });
    return out;
  }

  /**
   * The `w:eastAsia` faces this document draws its Chinese with, first-use order, deduplicated — the runs' and the
   * numbered marks' alike, because the page has to declare every family something on it reaches through the ladder.
   * `render.ts` turns these into `@font-face` rules; a face nobody names on the page costs nothing to declare.
   */
  eastAsiaFaces(): string[] {
    const out: string[] = [];
    const seen = new Set<string>();
    const add = (face: string | undefined): void => {
      if (face === undefined || seen.has(face)) return;
      seen.add(face);
      out.push(face);
    };
    this.eachParagraph((p) => {
      // The mark's own resolved face rides along: a numbered paragraph's symbol builds its stack from it, and a
      // family with no rule behind it is a family the browser skips.
      add(p.markFontEastAsia);
      for (const run of p.inlines) add(run.fontEastAsia);
    });
    return out;
  }

  // --------------------------------------------------------------- comments

  /** Where a save writes the comment list: package path plus the root declarations to keep. */
  commentsPart(): { path: string; rootAttrs: string } {
    return { path: this.commentsPartPath, rootAttrs: this.commentsRoot };
  }

  /** The body relationship's target for the comments part, verbatim as read. */
  commentsRelTarget(): string {
    return this.commentsTarget;
  }

  /** True when the list is not what was read, so `word/comments.xml` has to be regenerated. */
  commentsDirty(): boolean {
    return JSON.stringify(this.comments.map(commentState)) !== this.commentsAtRead;
  }

  /** Comment ids are plain integers; a new one takes the next free number. */
  nextCommentId(): number {
    return this.comments.reduce((max, c) => Math.max(max, c.id), -1) + 1;
  }

  /** Registers a comment; the caller owns anchor placement in the body. A reply carries no anchors. */
  addComment(id: number, text: string, author: string, initials: string, parentId?: number): DocComment {
    const created: DocComment = {
      id,
      author,
      initials,
      date: stampNow(),
      paragraphs: commentParagraphs(text),
      editable: true,
      // No source text: a comment created here is always regenerated.
      raw: "",
      baseline: "",
      parentId,
    };
    this.comments.push(created);
    this.supplyIds();
    return created;
  }

  /** Replaces a comment's body. A body this model cannot rebuild refuses, so its markup survives. */
  setCommentBody(id: number, text: string): boolean {
    return this.setCommentRuns(id, commentParagraphs(text));
  }

  /**
   * Replaces a comment's body with runs, marks and all — what the panel reads back from the
   * field the reader edits. Pictures and anchors are dropped rather than written: a body that
   * holds either is already `editable === false` and stays verbatim, and the comments part has
   * no relationships of its own to point image or anchor at.
   */
  setCommentRuns(id: number, paragraphs: Inline[][]): boolean {
    const found = this.comments.find((c) => c.id === id);
    if (!found || !found.editable) return false;
    found.paragraphs = paragraphs.map((runs) => runs.filter((r) => !r.image && !r.anchor));
    this.supplyIds();
    return true;
  }

  /**
   * `w15:commentEx/@done`: the resolved state WPS and Word grey out rather than hide. The thread
   * is the unit — resolving a comment resolves the replies under it too, because a thread with an
   * unresolved reply still counts as unresolved in the pane, and Word writes the flag to every row.
   */
  setCommentResolved(id: number, resolved: boolean): boolean {
    const found = this.comments.find((c) => c.id === id);
    if (!found) return false;
    const seen = new Set<number>([found.id]);
    const thread = [found];
    for (let at = 0; at < thread.length; at++) {
      for (const c of this.comments) {
        if (c.parentId === thread[at].id && !seen.has(c.id)) {
          seen.add(c.id);
          thread.push(c);
        }
      }
    }
    for (const c of thread) c.resolved = resolved;
    this.supplyIds();
    return true;
  }

  /** Drops a comment from the list; the caller clears its anchors from the body. */
  removeComment(id: number): boolean {
    const at = this.comments.findIndex((c) => c.id === id);
    if (at < 0) return false;
    const [gone] = this.comments.splice(at, 1);
    // A reply must not be orphaned into a thread that no longer exists: it re-attaches to
    // whatever its parent answered, so deleting the middle of a thread keeps both ends.
    for (const c of this.comments) if (c.parentId === gone.id) c.parentId = gone.parentId;
    return true;
  }

  /**
   * Gives every comment the paragraph ids the extension parts key on. Only a save calls it:
   * `commentState` deliberately ignores ids, so supplying them cannot make an unedited
   * document look edited. Ids read from the package are kept; a comment that never had one
   * (or gained a paragraph) takes the next free value.
   */
  private supplyIds(): void {
    const taken = new Set<string>();
    for (const c of this.comments) for (const id of c.paraIds ?? []) taken.add(id);
    for (const entry of commentExEntries(this.commentsExXmlAtRead)) taken.add(entry.paraId);
    for (const entry of commentIdEntries(this.commentIdsXmlAtRead)) taken.add(entry.paraId);
    for (const c of this.comments) {
      const have = c.paraIds ?? [];
      c.paraIds = have.length < c.paragraphs.length ? [...have, ...allocIds(c.paragraphs.length - have.length, taken)] : have;
      if (!c.durableId) c.durableId = allocIds(1, taken)[0];
    }
  }

  /** Called before a save writes comment parts, so regenerated bodies keep their paragraph ids. */
  prepareCommentsForSave(): void {
    this.supplyIds();
  }

  /**
   * The extension parts rebuilt from the current list: one `w15:commentEx` per comment
   * carrying its resolve flag and reply link, one `w16cid:commentId` per durable id, and
   * every row keyed on a paragraph this model does not own kept verbatim. Rows of deleted
   * comments go with them — a leftover would point Word at a paragraph that is gone.
   */
  commentExtensionParts(): { path: string; xml: string; relType: string; target: string }[] {
    if (!this.pkg) return [];
    this.supplyIds();
    const rows = this.comments.filter((c) => c.paraIds?.[0]);
    const owned = new Set(rows.map((c) => c.paraIds![0]));
    const parentKey = (c: DocComment): string | undefined =>
      c.parentId === undefined ? undefined : this.comments.find((x) => x.id === c.parentId)?.paraIds?.[0];
    // A part this package carried has to be rewritten even after its last comment is gone:
    // the rebuilt root is what drops the deleted comments' rows, and passing the source bytes
    // through instead would leave Word pointed at paragraph ids that no longer exist. A part
    // it never carried is only worth creating while there is state to put in it.
    const parts: { path: string; xml: string; relType: string; target: string }[] = [];

    const exRows = this.foreignRows(commentExEntries(this.commentsExXmlAtRead), owned);
    if (rows.length || this.commentsExXmlAtRead) parts.push({
      path: this.commentsExPath,
      xml: commentsExXml(
        this.commentsExRoot,
        rows.map((c) => ({ paraId: c.paraIds![0], parentParaId: parentKey(c), resolved: c.resolved === true })),
        exRows.map((e) => e.raw)
      ),
      relType: REL_COMMENTS_EX,
      target: this.commentsExTarget,
    });

    const idRows = this.foreignRows(commentIdEntries(this.commentIdsXmlAtRead), owned);
    if (rows.length || this.commentIdsXmlAtRead) parts.push({
      path: this.commentsIdsPath,
      xml: commentIdsXml(
        this.commentIdsRoot,
        rows
          .map((c) => ({ paraId: c.paraIds![0], durableId: c.durableId ?? "" }))
          .filter((r) => r.durableId !== ""),
        idRows.map((e) => e.raw)
      ),
      relType: REL_COMMENTS_IDS,
      target: this.commentsIdsTarget,
    });
    return parts;
  }

  /**
   * Extension rows this model does not own: keyed on a paragraph no current comment claims.
   * A row that belonged to a comment the user deleted goes with it; a row that cannot be
   * attributed at all is kept verbatim, because losing markup we do not understand is the
   * worse failure.
   */
  private foreignRows<T extends { paraId: string; raw: string }>(entries: T[], owned: Set<string>): T[] {
    const alive = new Set(this.comments.map((c) => c.id));
    return entries.filter((e) => {
      if (owned.has(e.paraId)) return false;
      const owner = this.paraOwner.get(e.paraId);
      return owner === undefined || alive.has(owner);
    });
  }

  private readComments(pkg: OpcPackage, rels: Map<string, Rel2>): void {
    const rel = [...rels.values()].find((r) => r.type.endsWith("/comments"));
    if (!rel) return;
    const part = pkg.resolve("word/document.xml", rel);
    const raw = pkg.text(part);
    if (raw === undefined) return;
    this.commentsPartPath = part;
    this.commentsTarget = rel.target;
    this.commentsRoot = commentsRootAttrs(raw);
    const root = unwrap(pkg.xml(part), "comments");
    const spans = commentSpans(raw);
    for (const node of asArray(root?.comment)) {
      const id = Number(node["@_id"]);
      if (!Number.isFinite(id)) continue;
      const paragraphs: Inline[][] = [];
      for (const child of children(node)) {
        if (child.name !== "p") continue;
        const runs: Inline[] = [];
        for (const inner of children(child.node)) {
          // A reference mark inside a comment body links a reply; it is not a body anchor.
          if (inner.name === "r") runs.push(...this.run(inner.node, rels).filter((i) => !i.anchor));
        }
        paragraphs.push(runs);
      }
      const span = spans.find((s) => s.id === id)?.xml ?? "";
      this.comments.push({
        id,
        author: String(node["@_author"] ?? ""),
        initials: node["@_initials"],
        date: node["@_date"],
        paragraphs,
        editable: !/<[\w.-]+:(tbl|drawing|pict|fldSimple|fldChar|hyperlink|sdt|ins|del|moveFrom|moveTo)[ />]/.test(span),
        raw: span,
        baseline: bodyShape(paragraphs),
      });
    }
    this.commentsXmlAtRead = raw;
    for (const span of spans) for (const id of commentParaIds(span.xml)) if (!this.paraOwner.has(id)) this.paraOwner.set(id, span.id);
    for (const c of this.comments) c.paraIds = commentParaIds(c.raw);
    this.readCommentExtensions(pkg, rels);
    this.commentsAtRead = JSON.stringify(this.comments.map(commentState));
  }

  /** A related part's text and where it lives; an empty text when the package has no such part. */
  private extensionPart(pkg: OpcPackage, rels: Map<string, Rel2>, type: string, name: string): { text: string; path: string; target: string } {
    const rel = [...rels.values()].find((r) => r.type.endsWith(type));
    if (!rel) return { text: pkg.text(`word/${name}`) ?? "", path: `word/${name}`, target: name };
    const path = pkg.resolve("word/document.xml", rel);
    return { text: pkg.text(path) ?? "", path, target: rel.target };
  }

  /**
   * `commentsExtended.xml` and `commentsIds.xml` joined onto the comments by paragraph id:
   * the resolve flag, the reply link and the durable id. Both parts are optional in OOXML —
   * WPS creates them only when it needs them (`WordMainPart::AddCommentsExtendedPart`) — so a
   * document can carry comments with neither, and then there is no state to read.
   */
  private readCommentExtensions(pkg: OpcPackage, rels: Map<string, Rel2>): void {
    const ex = this.extensionPart(pkg, rels, "/commentsExtended", "commentsExtended.xml");
    this.commentsExPath = ex.path;
    this.commentsExTarget = ex.target;
    this.commentsExXmlAtRead = ex.text;
    this.commentsExRoot = commentsExRootAttrs(ex.text);

    const ids = this.extensionPart(pkg, rels, "/commentsIds", "commentsIds.xml");
    this.commentsIdsPath = ids.path;
    this.commentsIdsTarget = ids.target;
    this.commentIdsXmlAtRead = ids.text;
    this.commentIdsRoot = commentIdsRootAttrs(ids.text);

    const byPara = new Map(commentExEntries(ex.text).map((e) => [e.paraId, e]));
    const durable = new Map(commentIdEntries(ids.text).map((e) => [e.paraId, e.durableId]));
    const owners = new Map<string, number>();
    for (const c of this.comments) if (c.paraIds?.[0]) owners.set(c.paraIds[0], c.id);
    for (const c of this.comments) {
      const key = c.paraIds?.[0];
      if (!key) continue;
      const row = byPara.get(key);
      if (row) {
        c.resolved = row.resolved;
        const parent = row.parentParaId ? owners.get(row.parentParaId) : undefined;
        if (parent !== undefined && parent !== c.id) c.parentId = parent;
      }
      c.durableId = durable.get(key);
    }
  }

  /**
   * The document's sections in the order they start. A section break hangs its `sectPr` off the `w:pPr` of
   * the paragraph that ends the section; the body's own direct `sectPr` child is the last section. Measured
   * on the authorized 46 份: 58 `w:sectPr` in all, of which 12 are the paragraph-level kind stated in 8 files,
   * and 8 of the 46 documents hold more than one section (`.scratch/pgmar.mjs`, `.scratch/corpus-recheck.mjs`);
   * every body-level `sectPr` sits at the end of `w:body`. The old per-file split — 115 sections in 86
   * documents, 68 with one, 7 with two, 11 with three — is the quarantined cohort's arithmetic, 待重证 here.
   */
  private sectionChain(source: string): { xml: string; fromBody: boolean }[] {
    const out: { xml: string; fromBody: boolean }[] = [];
    for (const kid of bodyChildSpans(source)) {
      if (kid.tag === "sectPr") {
        out.push({ xml: kid.xml, fromBody: true });
        continue;
      }
      if (kid.tag !== "p") continue;
      const pPr = childSpans(elementContent(kid.xml)).find((s) => s.tag === "pPr");
      if (!pPr) continue;
      const inner = childSpans(elementContent(pPr.xml)).find((s) => s.tag === "sectPr");
      if (inner) out.push({ xml: inner.xml, fromBody: false });
    }
    return out;
  }

  private readHeaderFooters(pkg: OpcPackage, rels: Map<string, Rel2>, chain: { xml: string; fromBody: boolean }[]): void {
    // The band the editor draws sits on the page the document **starts** with, so it belongs to the first
    // section — and only forward: a section that cites nothing continues the section before it, never one
    // after it. Reading the body's `sectPr` instead read the last section, which on these files reports a
    // header 4 times and a footer 5 times that the opening page does not have, and cites a different part
    // outright once each.
    const raw = chain[0]?.xml ?? "";
    if (!raw) return;
    const ns = /xmlns:([A-Za-z0-9_.-]+)="http:\/\/schemas\.openxmlformats\.org\/wordprocessingml\/2006\/main"/.exec(raw);
    this.sectPrefix = ns ? ns[1]! : "w";
    const shown = (kind: "header" | "footer", xml: string) => {
      const list = childSpans(elementContent(xml))
        .filter((s) => s.tag === `${kind}Reference`)
        .map((s) => attrsOf(s.xml));
      // Which band the page shows is the section's own statement, not the order it wrote its references in
      // (§17.10.3): the first-page band belongs to the opening page only when the section states `w:titlePg`, and
      // an `even` band belongs to `w:evenAndOddHeaders`. Measured on the authorized 46 份 (`.scratch/even-odd-audience.py`:
      // 46 件都有 `word/settings.xml`，58 节、38 枚带引用): the switch is stated **0 枚** — the element never appears —
      // while 4 节 do cite an `even` **footer** and none an `even` header, and none of those 4 states `w:titlePg`
      // (the cohort states it on 2 节). So an even band is a page the file has ready but never asks to show, and the
      // switch is the file's decision, not the renderer's. What WPS paints when a document *does* turn it on is 未取证.
      const firstPage = childSpans(elementContent(xml)).some((s) => s.tag === "titlePg");
      // A section that asks for a different first page but cites no first-page band still has the part it holds:
      // blanking the page here would take away text the file wrote. What WPS shows for that shape is 未取证.
      return (
        (firstPage ? list.find((a) => a["type"] === "first" && a["id"]) : undefined) ??
        list.find((a) => a["type"] === "default" && a["id"])
      );
    };
    const picks: Record<"header" | "footer", Record<string, string> | undefined> = { header: shown("header", raw), footer: shown("footer", raw) };
    // Whether the section already cites the band this page shows — the question a save asks before adding a
    // reference of its own. A band the section cites but the page does not show is not a band the model holds.
    this.sectHas = { header: !!picks.header?.["id"], footer: !!picks.footer?.["id"] };
    for (const kind of ["header", "footer"] as const) {
      const id = picks[kind]?.["id"];
      if (!id) continue;
      const hf = this.bandPart(pkg, rels, id, kind);
      if (!hf) continue;
      if (kind === "header") this.header = hf;
      else this.footer = hf;
    }
    // Read in the order the sections start, so the page-number atoms get their keys in the order the document
    // opens with — the same order as before this walked the whole chain. This holds only what each section
    // *itself* cites; which band a section's page shows is `sectionBand` below, and what a save writes is the
    // part a page shows (`headerFooterParts`).
    const opening: SectionBands = {};
    if (this.header) opening.header = this.header;
    if (this.footer) opening.footer = this.footer;
    this.sectionRefs = [opening];
    for (let k = 1; k < chain.length; k++) {
      const own: SectionBands = {};
      for (const kind of ["header", "footer"] as const) {
        const id = shown(kind, chain[k]!.xml)?.["id"];
        const hf = id ? this.bandPart(pkg, rels, id, kind) : undefined;
        if (hf) own[kind] = hf;
      }
      this.sectionRefs.push(own);
    }
  }

  /**
   * 第 `section` 节的纸显示的那一份带子（#157）：本节自己 cite 的用本节的，什么都没说的一节接着**它前面**
   * 最近的那一节——只向前，不向后（§17.10.3 的 `linkToPrevious`；「向后不继承」这一条在 WPS 真机上仍未取证，
   * #58）。整条链都没 cite 时回落到文档级那一份：那正是这一趟用户在页眉里新建的部件，它属于开篇那一节，
   * 往后每一节都接着它。
   */
  sectionBand(kind: "header" | "footer", section: number): HeaderFooter | undefined {
    for (let k = Math.max(0, Math.min(section, this.sectionRefs.length - 1)); k >= 0; k--) {
      const own = this.sectionRefs[k]?.[kind];
      if (own) return own;
    }
    return kind === "header" ? this.header : this.footer;
  }

  /**
   * The part one `w:headerReference` / `w:footerReference` cites, read as a band. Cached by package path,
   * because a chain of sections routinely cites one part from several of them and the page must not hold two
   * copies of it: identity is what lets an edit of the band the user types into show up on every sheet that
   * section's page paints (`doc-edit.ts` 的 `paintBandCopies`).
   */
  private bandPart(pkg: OpcPackage, rels: Map<string, Rel2>, id: string, kind: "header" | "footer"): HeaderFooter | undefined {
    const rel = rels.get(id);
    if (!rel) return undefined;
    const part = pkg.resolve("word/document.xml", rel);
    if (!pkg.parts.has(part)) return undefined;
    const held = this.bandByPart.get(part);
    if (held) return held;
    const partRaw = pkg.text(part) ?? "";
    // Read by name, not by position: WPS puts `r:id` before `w:type` on its header and footer references, where
    // the order this used to look for is Word's, so a real WPS page reported no header at all — nothing to show,
    // nothing to edit, nothing to splice back.
    // The texts come off the same span scan a save splices into, so paragraph i of the model is
    // paragraph i of the bytes. Pairing a parsed tree against spans would let the two count paragraphs
    // differently, and the save would then rewrite the wrong one.
    const rootSpan = childSpans(partRaw).find((s) => s.tag === (kind === "header" ? "hdr" : "ftr"));
    const seen = rootSpan ? childSpans(elementContent(rootSpan.xml)) : [];
    const found: { xml: string }[] = [];
    const walked = bandParagraphs(seen, (s) => (found.push(s), ""));
    // A band the walk refuses keeps the flat read it always had: its paragraphs as text, whole part held as
    // its own bytes. Refusing is not the same as dropping what the page was showing, so the text view
    // reaches into the controls the walk would not edit.
    const paraSpans = walked === null ? bandTextSpans(seen) : found;
    const parsed = paraSpans.map((s) => hfParagraphSegments(s, () => this.carryKey++));
    // Anything the paragraphs do not account for — a table or a container sitting beside them — is still a
    // reason to hold the part as its own bytes; within them the piece-by-piece read is the stricter rule.
    const editable = walked !== null && parsed.every((p) => p.ok);
    const segments = editable ? parsed.map((p) => p.segs) : undefined;
    const paragraphs = segments ? segments.map(hfSegmentsText) : paraSpans.map(hfParagraphText);
    const place = paraSpans.map((s) => this.bandPlace(s));
    const hf: HeaderFooter = {
      part,
      id,
      editable,
      paragraphs,
      baseline: [...paragraphs],
      ...(segments ? { segments, baselineSegments: segments.map(copyHfSegments) } : {}),
      ...(segments && parsed.some((p) => p.boxed) ? { boxedParagraphs: parsed.map((p) => p.boxed) } : {}),
      ...(place.some((p) => p !== undefined) ? { place } : {}),
      source: partRaw,
    };
    this.bandByPart.set(part, hf);
    return hf;
  }

  /**
   * One band paragraph's own sentence about where it sits, with the alignment its style chain lends beneath it.
   *
   * Read off the same span scan a save splices into, so paragraph i of the page is paragraph i of the bytes —
   * pairing a parsed tree against these spans would let the two count paragraphs differently.
   *
   * The paragraph's own `w:jc` wins; under it sits `w:pStyle`'s chain, which is where most of the band's answer
   * lives (of those 76 段, 12 say it themselves and 52 take one from the chain — 48 of those saying `left`, which
   * is the number the chain really lends and not a guess). `w:framePr` is the paragraph's alone: 授权 46 份的
   * `word/styles.xml` writes it 0 times, so there is no chain to walk for it (`.scratch/layout-families-audience.py`).
   */
  private bandPlace(span: { xml: string }): BandPlace | undefined {
    const pr = childSpans(elementContent(span.xml)).find((k) => k.tag === "pPr");
    if (!pr) return undefined;
    const inner = childSpans(elementContent(pr.xml));
    const attrs = (tag: string): Record<string, string> | undefined => {
      const el = inner.find((k) => k.tag === tag);
      return el ? attrsOf(el.xml) : undefined;
    };
    const twips = (raw: string | undefined): number | undefined => {
      const n = raw === undefined ? NaN : Number(raw);
      return Number.isFinite(n) ? n : undefined;
    };
    const out: BandPlace = {};
    const lent = this.paraLent(attrs("pStyle")?.val);
    const align = attrs("jc")?.val ?? lent?.align;
    if (align) out.align = align;
    // 缩进按正文同一条裁判合：段落自己那一句先答，答过的那一件事链上不许再叠（`indResolve`）。带子的段不是一份
    // `Paragraph`，它住在这一次 span 扫描里，所以属性是从 `attrsOf` 的裸名取的（见 `indOfAttrs`）。
    const ownInd = attrs("ind");
    const ind = indResolve(indOfAttrs((a) => ownInd?.[a]), lent);
    if (Object.values(ind).some((v) => v !== undefined)) out.ind = ind;
    const frame = attrs("framePr");
    if (frame) {
      const f: FramePlace = {};
      for (const [key, name] of [["hAnchor", "hAnchor"], ["vAnchor", "vAnchor"], ["xAlign", "xAlign"], ["yAlign", "yAlign"]] as const) {
        if (frame[name]) f[key] = frame[name];
      }
      const x = twips(frame.x);
      const y = twips(frame.y);
      if (x !== undefined) f.xTwips = x;
      if (y !== undefined) f.yTwips = y;
      out.frame = f;
    }
    return out.align === undefined && out.frame === undefined && out.ind === undefined ? undefined : out;
  }

  /**
   * Relationships the body carried, every one of them and with the part each resolves to.
   *
   * Dropping the picture and hyperlink relationships here because "the writer re-derives those" was
   * half true: the writer re-derives the ones its own tags name, but a drawing, a `w:fallback` VML
   * `v:imagedata` or a hyperlink inside a structure that travels by the byte still cites the id the
   * file wrote. Measured on 86 real WPS files of the quarantined cohort: the originals hold zero dangling
   * citations, and a save that renumbered image relationships left one carried `r:id` pointing at nothing — the
   * shape lost; relationship tables have not been re-audited on the authorized 46 份, so that zero is 待重证 there.
   * So the id is the file's to keep, and `path` is what lets a save recognise the relationship it is
   * about to invent as one that already exists.
   */
  sourceRels(): { id: string; type: string; target: string; targetMode: string; path: string }[] {
    return this.bodyRels;
  }

  /**
   * Parts this model does not own (theme, settings, headers, docProps…), byte-identical.
   * Saving regenerates only the parts the model edits and copies the rest through, so a
   * property we cannot show yet is not silently dropped from the user's file.
   */
  /**
   * Every part this writer does not own, with its bytes and its declared type as the package had them.
   * Media belongs in this list. The pictures the model holds are rewritten from their own data below, but
   * the media a *carried* drawing, an OLE object or a letterhead header points at has no model entry, and
   * dropping it leaves the relationship the body still cites pointing at a part that is no longer there.
   * An image the user deleted therefore leaves its bytes behind: an orphan costs a few kilobytes, a dangling
   * citation costs the picture next to every other one when a reader repairs the file.
   */
  untouchedParts(): { path: string; bytes: Uint8Array; type?: string }[] {
    if (!this.pkg) return [];
    const owned = new Set(["[Content_Types].xml", "word/document.xml", "word/_rels/document.xml.rels"]);
    return [...this.pkg.parts]
      .filter(([path]) => !owned.has(path) && !path.endsWith("/"))
      .map(([path, bytes]) => ({ path, bytes, type: this.pkg!.contentType(path) }));
  }
}

interface Rel2 {
  id: string;
  type: string;
  target: string;
  targetMode: string;
}

/**
 * The relationship id a drawing cites its picture through, as the source wrote it. A save swaps this for the id
 * its own relationship builder hands out and returns the rest of the drawing untouched — this one token is all
 * of the drawing that is the writer's to change.
 */
function embedOf(xml: string): string | undefined {
  return /\br:(?:embed|link)="([^"]+)"/.exec(xml)?.[1];
}

function imageOf(node: any, rels: Map<string, Rel2>, media: Map<string, string>, pkg?: OpcPackage): Inline["image"] {
  const stack = [node];
  let embed: string | undefined;
  let extent: any;
  let name = "";
  while (stack.length) {
    const cur = stack.pop();
    if (!cur || typeof cur !== "object") continue;
    for (const [k, v] of Object.entries(cur)) {
      if (k === "blip") embed = (v as any)["@_embed"] ?? (v as any)["@_r:embed"];
      if (k === "ext" || k === "extent") {
        const cx = (v as any)["@_cx"];
        if (cx && !extent) extent = v;
      }
      if (k === "@_name" && typeof v === "string") name = v;
      if (v && typeof v === "object") stack.push(v);
    }
  }
  const rel = embed ? rels.get(String(embed)) : undefined;
  const path = rel && pkg ? pkg.resolve("word/document.xml", rel) : undefined;
  const src = path ? media.get(path) : undefined;
  if (!src || !path) return undefined;
  const EMU_PER_PX = 9525;
  return {
    src,
    width: extent ? Number(extent["@_cx"]) / EMU_PER_PX : 320,
    height: extent ? Number(extent["@_cy"]) / EMU_PER_PX : 240,
    name,
    path,
  };
}

/** Whether a paragraph child is a wrapper that signs the runs it holds. */
const isRevision = (name: string): boolean => name === "ins" || name === "del" || name === "moveFrom" || name === "moveTo";

/** The signature on a revision wrapper. A move is Word's way of signing the same text leaving one
 *  place and arriving in another: `w:moveFrom` is the deletion it behaves like, `w:moveTo` the
 *  insertion. Either read as nothing at all would take the moved text out of the document. */
function revisionOf(tag: string, node: any): Revision {
  const id = Number(node?.["@_id"]);
  return {
    kind: tag === "del" || tag === "moveFrom" ? "del" : "ins",
    id: Number.isFinite(id) ? id : 0,
    author: String(node?.["@_author"] ?? ""),
    date: node?.["@_date"],
    ...(tag === "moveFrom" ? { move: "from" as const } : tag === "moveTo" ? { move: "to" as const } : {}),
  };
}

/** The element a revision's record is written as: a move keeps its own tag, so that reading a file and
 *  saving it again does not turn a moved passage into an unrelated deletion and insertion. */
export const revisionTag = (rev: Revision): "ins" | "del" | "moveFrom" | "moveTo" =>
  rev.move === "from" ? "moveFrom" : rev.move === "to" ? "moveTo" : rev.kind;

/** What a pane calls a pending record: the move's own name where the file signed a move, otherwise
 *  whether text came in or went out. */
const entryType = (rev: Revision): "insert" | "delete" | "move-from" | "move-to" =>
  rev.move === "from" ? "move-from" : rev.move === "to" ? "move-to" : rev.kind === "ins" ? "insert" : "delete";

/** Every run property this model carries, so a rejected formatting change can clear the ones the record dropped. */
/**
 * 「这条线上有没有线」——`underline` 换成形状名之后，问在不在的 callers 一律走这里，不留 `if (r.underline)`
 * 那种把 `"none"` 也当成「有」的写法。`undefined`（没提过）与 `"none"`（说了不要）在纸上都是无，在字节上是两句话：
 * 前者不拼元素，后者拼 `<w:u w:val="none"/>`（#85 定的那条，写端自己守着）。
 */
export const hasLine = (underline?: string): boolean => underline !== undefined && underline !== "none";

/**
 * 字号框里打进来的那个数：`w:sz` 数的是**半点**（`ST_HpsMeasure`，§17.18.60），所以落笔前先取到半点上——
 * 13.7 在件里只能说成 13.5，框里也就不该留着 13.7。只认十进制数（`10.5`、`12`）；不是数、不到半个点、
 * 或超过 1638 磅（Word 界面那一格的上限；WPS 的上限未取证）一律不认，交回 `undefined`，调用方据此不落笔。
 */
export function parseSizePt(text: string): number | undefined {
  const t = text.trim();
  if (!/^\d+(?:\.\d+)?$/.test(t)) return undefined;
  const pt = Math.round(Number(t) * 2) / 2;
  return pt >= 0.5 && pt <= 1638 ? pt : undefined;
}

export const RUN_STYLE_KEYS: (keyof RunStyle)[] = ["bold", "italic", "underline", "strike", "sizePt", "color", "highlight", "font", "fontEastAsia", "vertAlign", "charSpacingTwips"];

/**
 * The run properties a style may lend a run — all of them, both faces included. `font` was held back while the
 * model had only the one face from `w:rFonts@w:ascii`: 350 of the quarantined cohort's style-side `w:rFonts`
 * elements name a different face for `w:eastAsia` (`.scratch/run-font-shape.ts`), and handing a run the Latin one
 * as its only face would draw its Chinese in a face the file never chose for it. That styles-layer split is
 * 待重证 on the authorized 46 份, but the shape itself is not: of their body parts' `w:rFonts` elements 4,475
 * write both `w:ascii` and `w:eastAsia` and 161 of those name two different literal faces
 * (`.scratch/corpus-recheck2.mjs`). With `fontEastAsia` read as the second face, the pair says what the element
 * says — on the old count 2,928 of 61,702 modelled runs end up with two different faces from the ladder, itself
 * 待重证 — and the page dispatches them by script (`render.ts`). The 93 style-side elements that state only
 * `w:eastAsia` belong to that same 待重证 count, and they reach their runs too.
 */
export const LENT_RUN_KEYS: (keyof RunStyle)[] = RUN_STYLE_KEYS;

/** A run's formatting on its own — the properties, without the text, anchors or revision records hanging off it. */
export function runStyleOf(run: Inline): RunStyle {
  const out: Record<string, unknown> = {};
  for (const key of RUN_STYLE_KEYS) if (run[key] !== undefined) out[key] = run[key];
  return out as RunStyle;
}

/**
 * Putting a run's formatting back the way the `w:rPrChange` record had it. The record holds the whole
 * previous property set, so a property it does not mention is one that was *not* set — keeping the
 * current one would leave the change half refused.
 */
function restoreStyle(run: Inline, fmt: FormatRevision): Inline {
  const out: Inline = { ...run };
  for (const key of RUN_STYLE_KEYS) delete out[key];
  Object.assign(out, fmt.prev);
  // And the children no field can spell come back with them, because the record wrote them too: a refusal
  // puts the run's whole property set away, not an approximation of it built from the field list.
  if (fmt.prevSlots) {
    out.rPrSlots = fmt.prevSlots;
    out.asRead = fmt.prev;
  }
  return out;
}

/**
 * Which unit a `w:spacing@w:line` is written in. The scope of the 3,948 figure, so it is not read as an all-parts
 * total: `w:spacing` elements whose parent `w:pPr` is itself the child of a real `w:p`, summed over
 * `word/document.xml` (3,924) + `word/footnotes.xml` (12) + `word/endnotes.xml` (12) of the authorized 46 份 —
 * the per-part table is `.scratch/spacing-part-scope.py`, and `.scratch/corpus-recheck2.mjs` summed the same three.
 * The other parts are separate populations: the 页眉页脚 parts hold 74 such `w:pPr` blocks and write **no** `w:spacing`
 * in them, `word/styles.xml` writes 246 inside `w:style`'s `w:pPr` (a different lender, not a paragraph's own), and
 * 32 of the corpus's 4,226 `w:spacing` elements sit in a `w:rPr` (character spacing, §17.3.2.25).
 *
 * Counted that way (`.scratch/corpus-recheck2.mjs`): all 3,948 state a line height, and every
 * one states its rule beside it — the pair counts 3,948 and either attribute alone counts 0 — so there is no
 * "height with no unit" to resolve against §17.3.1.27's default. The rules themselves are only ever `w:auto`
 * 2,890, `w:exact` 1,004 and `w:atLeast` 54, so nothing outside the three literals appears and an unrecognised one
 * stays unread rather than being guessed at; the element itself still travels as the file's own text.
 *
 * 同一趟扫描把「只说段前段后而不说行距」那一支也数了：正文部件里写段前或段后的 `w:pPr/w:spacing` 共 252 枚，
 * **252 枚同枚都还说了行距**，一枚「只有间距没有行高」的都没有（`.scratch/摆位-authorized.py`）。作废口径 86 份那组
 * 「25,801 枚说行距或说间距、5,602 枚只说间距」的算术是旧读数，不再引用。
 */
function lineRuleOf(spacing: any): LineSpacingRule | undefined {
  const v = spacing?.["@_lineRule"];
  return v === "auto" || v === "exact" || v === "atLeast" ? v : undefined;
}

/**
 * What a `w:spacing` states, as fields, one attribute at a time — §17.3.1.27's rule resolves each of them
 * separately, so a paragraph may write its own height and take the rule from its style, and that is what the
 * file said rather than a mix to be corrected.
 *
 * Read through the style because a paragraph can ask for a line height without writing one. 授权口径 46 份把这一档
 * 的受众数过（`.scratch/basedon-chain.ts`，输出存 `.scratch/basedon-authorized.json`；正文 7,357 段写了 `w:pPr`）：
 * 段自己说行距的 3,924 段之外，**264 段**由 `w:pStyle` 命名那一枚样式直接给（`w:auto` 226、`w:exact` 34、
 * `w:atLeast` 4），同一枚样式还借出段前 198 段、段后 206 段；再往上的 `w:basedOn` 链只有 **18 段**的行距落在第 2 跳
 * （全部 `w:auto`，今天这一档读不到，另有 jc 790 / firstLine 44 同形）；默认段落样式那一级给 **772 段**行距
 * （也全部 `w:auto`，同样读不到，另有 jc 3,731 / firstLine 320 / 段后 78）。作废口径那组 5,110 / 64 / 57 / 501 /
 * 613 的数是旧读数，不再引用。The walk stays on §17.7.2's ladder, and the direct half of the population is
 * measured there: of the 7,381 `w:pPr` blocks those three parts hold under a real `w:p`
 * (document 7,357 + footnotes 12 + endnotes 12), 3,948 carry a `w:spacing` that states a line height of their own —
 * `word/document.xml` alone is 3,924 of 7,357, which is the pair the sentence above quotes.
 */
function spacingOf(spacing: any): Partial<ParaStyle> {
  const out: Record<string, number | LineSpacingRule> = {};
  const before = num(spacing?.["@_before"]);
  if (before !== undefined) out.spacingBefore = before;
  const after = num(spacing?.["@_after"]);
  if (after !== undefined) out.spacingAfter = after;
  const line = num(spacing?.["@_line"]);
  if (line !== undefined) out.lineSpacing = line;
  const rule = lineRuleOf(spacing);
  if (rule !== undefined) out.lineSpacingRule = rule;
  return out as Partial<ParaStyle>;
}

/**
 * What a `w:ind` states, as fields. Only the spellings the files really use on their own are read: 授权口径 46 份
 * 的 `word/document.xml` 那 2,278 枚 `w:ind` 里，`w:firstLineChars` 有 1,210 枚，其中 94 枚不写 `w:firstLine`，
 * 而这 94 枚写的都是 `0`（没有一枚用字符单独说一个正的缩进）；`w:leftChars` 432 枚里 50 枚不写 `w:left`，同样
 * 是 48 枚 `0` 加 2 枚负数；`w:hangingChars` 62 枚、`w:rightChars` 8 枚全部带着自己的缇同伴
 * （`.scratch/chars-zero-audience.ts`）。作废口径 86 份那组「29,207 枚里 232 枚只写正字符数」的数是旧读数，
 * 不再引用。所以字符写法不需要自己的字段受众，需要的只是「段落自己答过这一件事」这个判断本身——它由
 * `indResolve` 用来挡住样式链，见那里的两段受众。
 *
 * Attributes the element does not state stay out of the result rather than coming back as `undefined`: a
 * `w:pPrChange` has to record which properties it wrote, and a paragraph has to tell "the file said zero" from
 * "the file said nothing", since the second is what lets its style's indent through.
 */
function indOf(ind: any): Partial<ParaStyle> {
  return indOfAttrs((attr) => ind?.[`@_${attr}`]);
}

/**
 * 一枚 `w:ind` 的六个槽，按属性名取——取的方式不定：正文那一条路从解析好的树上读 `@_left`，带子那一条路从
 * **保存要拼回去的那同一次 span 扫描**里读 `attrsOf` 的裸名。两处必须共用同一张「属性名 → 模型字段」表，
 * 否则同一句话在正文与带子里会解析成两件不同的事。
 */
function indOfAttrs(get: (attr: string) => string | undefined): Partial<ParaStyle> {
  const out: Record<string, number> = {};
  const take = (field: keyof ParaStyle, attr: string): void => {
    const v = num(get(attr));
    if (v !== undefined) out[field] = v;
  };
  take("indentLeft", "left");
  take("indentLeftChars", "leftChars");
  take("indentRight", "right");
  take("indentFirstLine", "firstLine");
  take("indentFirstLineChars", "firstLineChars");
  take("indentHanging", "hanging");
  return out as Partial<ParaStyle>;
}

/** 段落自己那句 `w:ind` 与链上那枚合出来的六枚缩进——同一件事的两枚写法在这里定谁答过。 */
export type IndentFields = Pick<ParaStyle, "indentLeft" | "indentLeftChars" | "indentRight" | "indentFirstLine" | "indentFirstLineChars" | "indentHanging">;

/**
 * 段落自己的一句 `w:ind` 加上样式链借来的那枚，合出页面要画的六属性。
 *
 * §17.3.1.14 让「字符」那一枚压过「缇」那一枚，§17.7.2 让段落自己压过样式。两句话合起来的意思是：段落只要用
 * 任一枚答过这一件事，链上的另一枚就不许再叠上来——今天写的正是那件会叠错的事。授权口径 46 份的
 * `word/document.xml` 里 `w:ind` 共 2,278 枚（`.scratch/chars-zero-audience.ts` 经产品代码数）：**94** 枚写
 * `w:firstLineChars="0"` 而完全不写 `w:firstLine`，其中 **74** 枚的样式链写了 `w:firstLine="883"`；另有 **62**
 * 枚只写自己的 `w:hanging`（既不写 `w:firstLine` 也不写 `w:firstLineChars`），其中 **34** 枚的链也写了 883。
 * 那 108 段自己的 `w:ind` 长得是 `w:left="425" w:leftChars="0" w:hanging="425" w:firstLineChars="0"` 这一类——
 * 段首缩进与左边界等长，正是「首行不缩进、其余行缩进」的悬挂形状，而 `firstLineChars="0"` 是把「首行不缩进」
 * 用字符那一枚又说明白的一次。按下面这两条之前，页面画的是链上那 883 缇：段首被缩进去一格，把文件自己说的
 * 「不缩进」盖掉了（`hanging` 也一并被 `firstLine ?? -hanging` 那一步跳过）。
 *
 * 反过来那两支没有实例：段自己写了 `w:firstLine`/`w:firstLineChars` 而页面仍带着链上 `hanging` 的 0 段，段只写
 * `w:firstLine="0"` 而页面带着链上 `firstLine` 的也是 0 段（`.scratch/hanging-borrow-audience.ts`），所以这里
 * 不开那一头的口。段自己两枚都写时仍照旧由缇那枚上身（见 `indentFirstLine` 的文档）。
 */
function indResolve(own: Partial<ParaStyle>, lent: Partial<ParaStyle> | undefined): IndentFields {
  const saidLeft = own.indentLeft !== undefined || own.indentLeftChars !== undefined;
  const saidFirstLine = own.indentFirstLine !== undefined || own.indentFirstLineChars !== undefined || own.indentHanging !== undefined;
  return {
    indentLeft: own.indentLeft ?? (saidLeft ? undefined : lent?.indentLeft),
    indentLeftChars: own.indentLeftChars ?? lent?.indentLeftChars,
    indentRight: own.indentRight ?? lent?.indentRight,
    indentFirstLine: own.indentFirstLine ?? (saidFirstLine ? undefined : lent?.indentFirstLine),
    indentFirstLineChars: own.indentFirstLineChars ?? lent?.indentFirstLineChars,
    indentHanging: own.indentHanging ?? lent?.indentHanging,
  };
}

/**
 * The numbering a `w:pPr` states, taken as one property: `w:numPr` is not a set of attributes a paragraph can
 * merge piece by piece the way it merges `w:ind`, so the first style in the chain that writes it settles the
 * paragraph's list, and the styles below it get no say.
 *
 * `w:numId` absent reads as 0, which is what the files here mean by it: the authorized 46 份 write 298 paragraphs
 * carrying their own `w:numPr`, 52 of those are `w:numId w:val="0"` — the value that cancels numbering
 * (§17.3.4.25) — and not one of the 298 leaves `w:numId` out (`.scratch/corpus-recheck2.mjs`, `.scratch/pgmar.mjs`).
 * The style side is 待重证: the old reading (11 of the 48 borrowing paragraphs pointing at a `w:numPr` that writes
 * only `w:ilvl`, `.scratch/num-lent-shapes.ts`, `.scratch/style-numpr-census.ts`) came from the quarantined cohort.
 * Reading an absent `w:numId` as 0 is what §17.3.4.28 says either way, and it leaves such a paragraph unnumbered —
 * which is what the page showed before this property was read at all.
 */
function numPrOf(numPr: any): { numId: number; ilvl: number } | undefined {
  if (!numPr) return undefined;
  return { numId: Number(numPr.numId?.["@_val"] ?? 0), ilvl: Number(numPr.ilvl?.["@_val"] ?? 0) };
}

/**
 * The outline level a `w:outlineLvl` states (§17.3.2.26), as the file signed it: 0…8 are heading levels one to
 * nine and 9 is body text. Of the authorized 46 份's 7,430 paragraph blocks, 202 state a level themselves and 206
 * more take one from their style chain — the 408 levels that resolves to are 298 at 0, 60 at 1, 14 at 2, 8 at 3,
 * 2 at 4 and **26 at 9** — and 172 paragraph styles state one (`.scratch/outline-census.ts`). The old census over
 * the 86 documents (46,417 paragraphs, 264 stating a level as 0×6 / 1×160 / 2×42 / 9×56, 582 more lent, 433 styles
 * stating one) is the quarantined cohort's arithmetic. A value outside 0…9 is read as no statement: the files here
 * never write one.
 */
function outlineOf(outlineLvl: any): number | undefined {
  if (outlineLvl === undefined) return undefined;
  const n = Number(outlineLvl["@_val"]);
  return Number.isInteger(n) && n >= 0 && n <= 9 ? n : undefined;
}

/**
 * The heading a paragraph is drawn as: 1 for `w:outlineLvl w:val="0"` up to 9 for 8, and nothing for 9 — the
 * value that puts the paragraph back to body text, whether it states that itself or borrows it.
 */
export function headingLevel(p: Paragraph): number | undefined {
  const level = p.outlineLevel;
  return level !== undefined && level <= 8 ? level + 1 : undefined;
}

/**
 * What a `w:pPr` says on its own, with nothing taken from the paragraph's style. A `w:pPrChange`
 * records the properties as they were written, while the model resolves a paragraph's alignment and
 * indent through its style, so the two have to be read by different rules.
 */
function paraStyle(pPr: any): ParaStyle {
  const s: ParaStyle = {};
  const styleId = pPr.pStyle?.["@_val"];
  if (styleId !== undefined) s.styleId = String(styleId);
  const align = pPr.jc?.["@_val"];
  if (align !== undefined) s.align = String(align);
  Object.assign(s, indOf(pPr.ind), spacingOf(pPr.spacing));
  const pbb = onOff(pPr.pageBreakBefore);
  if (pbb !== undefined) s.pageBreakBefore = pbb;
  const keepNext = onOff(pPr.keepNext);
  if (keepNext !== undefined) s.keepNext = keepNext;
  const keepLines = onOff(pPr.keepLines);
  if (keepLines !== undefined) s.keepLines = keepLines;
  const widow = onOff(pPr.widowControl);
  if (widow !== undefined) s.widowControl = widow;
  const snap = onOff(pPr.snapToGrid);
  if (snap !== undefined) s.snapToGrid = snap;
  const wrap = onOff(pPr.wordWrap);
  if (wrap !== undefined) s.wordWrap = wrap;
  const numPr = numPrOf(pPr.numPr);
  if (numPr) s.numPr = numPr;
  const outline = outlineOf(pPr.outlineLvl);
  if (outline !== undefined) s.outlineLevel = outline;
  // The box. `w:pBdr` names one child per side it says anything about — including `w:val="none"`, which is a
  // statement and not a silence — and `w:shd@w:fill` is the ground under the lines, kept as the file's own token
  // when it names none. Both are read the same way from a `w:style`'s `w:pPr`, from a `w:pPrChange` record, and
  // from a paragraph's own properties.
  const borders = paraBordersOf(asArray(pPr.pBdr)[0]);
  if (borders) s.borders = borders;
  const shd = shadingOf(asArray(pPr.shd)[0]);
  if (shd) s.shading = shd;
  const stops = tabStopsOf(asArray(pPr.tabs)[0]);
  if (stops) s.tabStops = stops;
  return s;
}

/** Every paragraph property a `w:pPrChange` carries, so a refusal can clear the ones the record dropped. */
export const PARA_STYLE_KEYS: (keyof ParaStyle)[] = [
  "styleId", "align", "indentLeft", "indentLeftChars", "indentRight", "indentFirstLine", "indentFirstLineChars",
  "indentHanging", "spacingBefore", "spacingAfter", "lineSpacing", "lineSpacingRule", "pageBreakBefore", "numPr",
  "outlineLevel",
];

/**
 * The properties a paragraph may take from its styles, in the order `paraLent` fills them. The style's own
 * `w:pPr` slice states more than this (`paraStyle` reads it whole); what is not listed here is not lent.
 *
 * The five page-keeping switches are asked of the ladder one rung at a time over the 7,393 `w:p` that
 * `word/document.xml` of the authorized 46 份 holds — body paragraphs, table cells and text boxes alike, since the
 * file writes these elements in all three — and this is where each of them actually sits
 * (`.scratch/toggle-lent-chain.ts`; `w:val="0"` counted as off, a bare element as on):
 *
 * | | 段自己 | 样式第 1 跳 | 第 2 跳 | 默认段落样式 | docDefaults | 无人说 |
 * |---|---|---|---|---|---|---|
 * | `w:widowControl` | 160 off / 86 on | 30 off / 28 on | 1,370 off | 5,601 off | 0 | 118 |
 * | `w:keepNext` | 200 off | 182 on | 0 | 0 | 0 | 7,011 |
 * | `w:keepLines` | 200 off | 182 on | 0 | 0 | 0 | 7,011 |
 * | `w:snapToGrid` | 224 off / 180 on | 24 off / 16 on | 12 off | 154 off | 0 | 6,783 |
 * | `w:pageBreakBefore` | 200 off | 0 | 0 | 0 | 0 | 7,193 |
 *
 * Every row sums to 7,393, and the same family comes back through `WordDocument.openBytes`: of the 3,120
 * paragraphs that walk reaches, 2,937 resolve to widow-**off**, 182 to `keepNext` on, 232 to `snapToGrid` off, and
 * none is left `undefined` while its own `w:pPr` stated the switch (`.scratch/widow-model.ts`; the walk stops at
 * `carried` containers, so its denominator is smaller than the file's own by design).
 *
 * This table answers *which rung speaks*; it is not the element count, and the two differ in one cell. Counting
 * the elements by offset instead (`PPR_ORDER`-free, `.scratch/ppr-ladder-referee.ts`) gives `word/document.xml`
 * 246 `w:widowControl` (86 bare / 160 off), 200 `w:keepNext`, 200 `w:keepLines`, 406 `w:snapToGrid` (180 bare /
 * 226 off) and 200 `w:pageBreakBefore`, plus 63 / 90 / 90 / 118 / **0** in `word/styles.xml` and 4 more
 * `w:snapToGrid` in the bands — four of the five rows agree with the 段自己 column exactly, and `w:snapToGrid` is
 * 2 short: the walk pairs `<w:p>` with `</w:p>` and does not reach those two paragraphs as themselves, while the
 * two statements are ordinary live ones — not inside a `w:pPrChange` (the 46 write that record 0 times,
 * `.scratch/snap-in-pprchange.ts`) and not a second copy in one `w:pPr` (`.scratch/snap-grid-where.ts` finds 406
 * blocks holding one element each, and none outside a `w:pPr`). Cite the element counts for 枚数 and this table
 * for the ladder.
 *
 * So the four that are in this list are in it because 7,001 paragraphs are told not to widow-protect without
 * writing a word — 30 from their own named style, 1,370 from a style one further up, 5,601 from the default
 * paragraph style — alongside 190 told not to snap and 182 told to keep with the next block. A chain-blind reader
 * gets all three wrong, which is what #126 and #127 fixed; and the answers mostly come from below the paragraph's
 * own style, since 6,971 of the 7,001 sit two rungs down or lower.
 *
 * `w:pageBreakBefore` stays out, now on a measured case rather than on §17.7.2's wording: `word/styles.xml` in the
 * 46 writes the element 0 times (`.scratch/pbb-spellings.ts`), so there is nothing here for the chain to lend.
 * What the body does write is 200 `<w:pageBreakBefore w:val="0"/>` and not one on — a statement the paragraph
 * makes for itself, and one that `onOff` reads as *no break*. Reading the element's presence as "break" instead
 * forced a page break on all 180 model paragraphs of those 200 (`.scratch/pbb-before-after.ts`, whose printed
 * column is files × statements — read it per file, never as a total).
 *
 * `w:docDefaults/w:pPrDefault/w:pPr` is present in all 46 份 and empty in all 46 (`.scratch/dd-control.ts`), so
 * the rung under the default paragraph style has nothing to lend — for this family and for the indent and
 * spacing above it.
 *
 * `w:pBdr` and `w:shd` join on the box the files actually draw. Over the 46 authorized docx-family files
 * (`.scratch/pbdr-audience.ts`), 24 of the 26 `w:pBdr` elements are in `word/styles.xml` — 16 of them naming all
 * four sides `none`, 8 naming a real bottom line — and 30 paragraph `w:shd` fills sit in 4 files' bodies under
 * styles that state none. Both are lent **whole**: `paraLent` stops at the first chain rung that states the key,
 * so a style that names only its bottom is answered by an ancestor that names the other three only when the
 * paragraph itself writes them (`mergeParaBorders` merges per side against the chain's set). No file in hand
 * makes a two-rung border chain, so that half is the shape of this ladder, not a measured case.
 */
const LENT_PARA_KEYS: (keyof ParaStyle)[] = [
  "align", "indentLeft", "indentLeftChars", "indentRight", "indentFirstLine", "indentFirstLineChars",
  "indentHanging", "spacingBefore", "spacingAfter", "lineSpacing", "lineSpacingRule", "numPr", "outlineLevel",
  "keepNext", "keepLines", "widowControl", "snapToGrid", "wordWrap", "borders", "shading",
];

/** Every paragraph property the model holds for it, from its own `w:pPr` and from its styles. */
export function paraStyleOf(p: Paragraph): ParaStyle {
  const out: Record<string, unknown> = {};
  for (const key of PARA_STYLE_KEYS) if (p[key] !== undefined) out[key] = p[key];
  return out as ParaStyle;
}

/**
 * 模型有字段、但不在 `PARA_STYLE_KEYS` 里的那八项。它们不参与 `ownParaStyle` 的「这一段自己说了哪些」那本账
 * （整枚靠 `w:pPr` 的槽位字节带回），可**退回一条格式改动**时它们与别项是同一件事：记录没说的就是当时没说，
 * 字段里留着的那些得清掉、再让样式链重新借给它——不清的话拒绝完这一段还顶着用户刚被否决的那一句，画错了，
 * 下一次存盘还会把被否决的话替它写回件里。
 *
 * `w:pBdr` 与 `w:shd` 在这一条上与前面那几项（三条保持规则与两枚开关）没有区别：`w:pPrChange` 记的是直接格式，一条「把边框去掉」的改动被
 * 否决后，段落字段里留着的那一套框必须由样式链重新借给它，而不是继续画用户已经撤回（并被撤回）的那个盒子。
 */
const MODEL_ONLY_PARA_KEYS = ["keepNext", "keepLines", "widowControl", "snapToGrid", "wordWrap", "borders", "shading", "tabStops"] as const;

/**
 * 拆段时新那一段从上一段带走的那一整套段落属性：字段那本账的**全部**，两件都算——`PARA_STYLE_KEYS` 写得出
 * 句子的，与只靠 `w:pPr` 槽位带回的那八项。
 *
 * 这一个清单由模型自己拼，不再由调用手数：手数的那一份（`doc-edit.ts` 的 `fresh`）漏掉过槽位族的四项，
 * 于是回车把上一段自己写的保持规则、那句「不吸网格」整个忘掉，而 `nextSplit` 只在样式写了 `w:next` 时才走，
 * 授权件里那种换样式一次都没命中过（`nextStyleOf` 的说明），也就是说丢那四项的正是天天在走的那条路。
 * 这一项加进来的是同一族：`w:pBdr` 与 `w:shd`。
 *
 * 值按引用带走（`numPr` 早就是这么办的），所以改这些字段是**换掉整个物件**，不是往共享的那个物件里改。
 *
 * 同一本账还有第二个用处：格式改动记录（`w:pPrChange`/`w:rPrChange`）「改之前那一句」带的也是这**全部**，
 * 不是 `paraStyleOf` 那半本。`paraStyleOf` 只列得出句子的`PARA_STYLE_KEYS`，是写出侧判「哪一句归段落」的账，
 * 拿它当记录那本账会漏掉槽位族那八项——单动「对齐到网格」或「允许西文在单词中间换行」时字段看不出改了口，
 * 记录就不留；留下了但带的是半本的，`restoreParaStyle` 把字段与字节放回段落上时两者不同步，被否决的那一句
 * 还剩一半在纸上画着。写侧那半本不动：补上槽位族会让 `syncParaSlots` 把样式借来的话替段落重写一遍。
 */
export function carriedShape(p: Paragraph): ParaStyle {
  const out: Record<string, unknown> = {};
  for (const key of [...PARA_STYLE_KEYS, ...MODEL_ONLY_PARA_KEYS]) if (p[key] !== undefined) out[key] = p[key];
  return out as ParaStyle;
}

/**
 * The snapshot the write side measures the user against: every property the paragraph **showed** when it came
 * in — its own `w:pPr` first, the style chain behind. A key it neither stated nor borrowed is absent, and
 * `moved()`/`ownParaStyle` read that absence as "this one is not this paragraph's to write".
 *
 * `snapToGrid` is in this set (round #134) because the checkbox needs both halves of it: a paragraph borrowing
 * `w:snapToGrid w:val="0"` from its style shows the box unchecked, and only against this snapshot does the write
 * side recognise that statement as the style's rather than as a click. The three keep flags joined for the same
 * reason in #220 — 「换行和分页」的四颗勾 can turn a borrowed 「开」 off, which takes a `w:val="0"` on the
 * paragraph, and the write side can only tell that apart from silence by what the paragraph drew on the way in.
 */
function paraReadSnapshot(p: Paragraph): ParaStyle {
  return {
    styleId: p.styleId,
    align: p.align,
    indentLeft: p.indentLeft,
    indentLeftChars: p.indentLeftChars,
    indentRight: p.indentRight,
    indentFirstLine: p.indentFirstLine,
    indentFirstLineChars: p.indentFirstLineChars,
    indentHanging: p.indentHanging,
    spacingBefore: p.spacingBefore,
    spacingAfter: p.spacingAfter,
    lineSpacing: p.lineSpacing,
    lineSpacingRule: p.lineSpacingRule,
    // 断点那一格只认「说了要」：件里确实有说不要的（授权 46 份正文 200 枚 `w:val="0"`），而那一枚的话由段落自己的
    // 槽位字节带着走——`keep` 与 `owned` 都对 false 一言不发，所以这里把 false 记成没说过，写出侧才不会替文件补一枚
    // 裸的 `<w:pageBreakBefore/>`，把「不断页」改成「断页」。
    pageBreakBefore: p.pageBreakBefore || undefined,
    numPr: p.numPr,
    outlineLevel: p.outlineLevel,
    // 三条保持规则（#220）：与 `snapToGrid` 同一条来路。写侧要拿它判「段上没有这一枚时，用户拨的这一下来得算不算数」——
    // 从样式链借来「开」的段，用户勾掉它得在段落上写一句 `w:val="0"` 才盖得住，而参照只能是入件时画着的那一份。
    // `nextSplit` 不读这本账（它拿旧链 `before` 比，见那里的说明），所以这里补上不改变拆段的口径。
    keepNext: p.keepNext,
    keepLines: p.keepLines,
    widowControl: p.widowControl,
    snapToGrid: p.snapToGrid,
    wordWrap: p.wordWrap,
    // 框与底纹进这本账与 `snapToGrid` 同一个理由：写出侧只把**这一击动过的**说出去，而「动过」是对着入件时
    // 显示的那一份比的。样式链借来的那个盒子留在账上，段落自己没有 `w:pBdr` 的槽位时才不会替文件编一句。
    borders: p.borders,
    shading: p.shading,
  };
}

/**
 * Whether a property field still holds what the read snapshot has for it. Every writer in this file and
 * `write-docx.ts` spells a property only once it has moved: a value the paragraph or run only borrowed from its
 * style is the style's to state, and writing it again as direct formatting is how a save makes the file claim
 * properties it never wrote.
 */
export const samePropValue = (a: unknown, b: unknown): boolean =>
  a === b || (typeof a === "object" && typeof b === "object" && JSON.stringify(a) === JSON.stringify(b));

/**
 * Out of a set of paragraph properties, the ones the paragraph said for itself.
 *
 * A paragraph's fields are filled from its style chain as well as from its own `w:pPr` (§17.7.2), so a field
 * being set says nothing about who set it, and neither the writer nor a `w:pPrChange` record may claim a
 * borrowed property as the paragraph's own: §17.13.5.19 records the direct formatting, and 159 of the
 * quarantined cohort's paragraphs state no `w:pPr` at all while still showing a whole set of properties, because
 * a style lends them (`.scratch/para-no-pr.ts`) — an audience not re-counted on the authorized 46 份, so the
 * number is 待重证 while the shape it measured is what §17.13.5.19 requires anyway. The read snapshot answers the
 * question — a value still equal to the one the
 * paragraph was read with came from the chain, and a paragraph this editor built has no snapshot, so every
 * field on it is its own.
 */
export function ownParaStyle(asRead: ParaStyle | undefined, s: ParaStyle): ParaStyle {
  const out: Record<string, unknown> = {};
  for (const key of PARA_STYLE_KEYS) {
    const v = s[key];
    if (v === undefined || v === false) continue;
    if (asRead && samePropValue(asRead[key], v)) continue;
    out[key] = v;
  }
  return out as ParaStyle;
}

/**
 * Out of a set of run properties, the ones the run said for itself — the same question `ownParaStyle` answers
 * for a paragraph, and the reason a run may be filled from the chain without the save claiming the chain's work.
 *
 * Unlike the paragraph set, `false` is a value the run wrote: `<w:b w:val="0"/>` turns off a bold the style
 * lends, and dropping it here would turn an explicit cancellation back into silence.
 */
export function ownRunStyle(asRead: RunStyle | undefined, s: RunStyle): RunStyle {
  const out: Record<string, unknown> = {};
  for (const key of RUN_STYLE_KEYS) {
    const v = s[key];
    if (v === undefined) continue;
    if (asRead && samePropValue(asRead[key], v)) continue;
    out[key] = v;
  }
  return out as RunStyle;
}

/** One property key and the `w:pPr`/`w:rPr` children that would state it for the paragraph or run itself. */
interface SlotStates {
  tags: string[];
  /** For an element that holds several properties (`w:ind`, `w:spacing`, `w:rFonts`), the attributes that carry this one. */
  attrs?: string[];
}

/**
 * 段落这一族各自住在哪枚 `w:pPr` 孩子里。`w:ind` 与 `w:spacing` 一枚元素托着几个属性，所以要再看属性名；
 * 其余一枚元素就是一句话。
 */
const PARA_SLOT_STATES: Partial<Record<keyof ParaStyle, SlotStates>> = {
  styleId: { tags: ["pStyle"] },
  align: { tags: ["jc"] },
  indentLeft: { tags: ["ind"], attrs: ["left"] },
  indentLeftChars: { tags: ["ind"], attrs: ["leftChars"] },
  indentRight: { tags: ["ind"], attrs: ["right"] },
  indentFirstLine: { tags: ["ind"], attrs: ["firstLine"] },
  indentFirstLineChars: { tags: ["ind"], attrs: ["firstLineChars"] },
  indentHanging: { tags: ["ind"], attrs: ["hanging"] },
  spacingBefore: { tags: ["spacing"], attrs: ["before", "beforeLines", "beforeAutospacing"] },
  spacingAfter: { tags: ["spacing"], attrs: ["after", "afterLines", "afterAutospacing"] },
  lineSpacing: { tags: ["spacing"], attrs: ["line"] },
  lineSpacingRule: { tags: ["spacing"], attrs: ["lineRule"] },
  pageBreakBefore: { tags: ["pageBreakBefore"] },
  numPr: { tags: ["numPr"] },
  outlineLevel: { tags: ["outlineLvl"] },
  keepNext: { tags: ["keepNext"] },
  keepLines: { tags: ["keepLines"] },
  widowControl: { tags: ["widowControl"] },
  snapToGrid: { tags: ["snapToGrid"] },
  wordWrap: { tags: ["wordWrap"] },
  borders: { tags: ["pBdr"] },
  shading: { tags: ["shd"] },
  tabStops: { tags: ["tabs"] },
};

/**
 * The same question for a run's `w:rPr`. `w:color` 的那句话在 2010 扩展里写作 `w14:textFill`（写出侧同一枚
 * `RUN_PROPS` 的两个孩子），所以两种拼法都算 run 自己说的；`w:rFonts` 一枚托两张脸，按属性分。
 */
const RUN_SLOT_STATES: Partial<Record<keyof RunStyle, SlotStates>> = {
  bold: { tags: ["b"] },
  italic: { tags: ["i"] },
  underline: { tags: ["u"] },
  strike: { tags: ["strike"] },
  sizePt: { tags: ["sz"] },
  color: { tags: ["color", "textFill"] },
  highlight: { tags: ["highlight"] },
  font: { tags: ["rFonts"], attrs: ["ascii", "hAnsi", "asciiTheme", "hAnsiTheme"] },
  fontEastAsia: { tags: ["rFonts"], attrs: ["eastAsia", "eastAsiaTheme"] },
  vertAlign: { tags: ["vertAlign"] },
  charSpacingTwips: { tags: ["spacing"] },
};

/**
 * Whether the paragraph or run itself wrote this property — 认槽位，不认字段值。槽位是这一段的原件写下的字，
 * 「自己写的值恰好与旧样式链相同」仍然是它自己写的；`ownParaStyle`/`ownRunStyle` 按值比，认不出那一格。
 */
function slotStates(slots: PropSlot[] | undefined, spec: SlotStates | undefined): boolean {
  if (!slots?.length || spec === undefined) return false;
  const slot = slots.find((s) => spec.tags.includes(s.tag));
  if (!slot) return false;
  if (spec.attrs === undefined) return true;
  return spec.attrs.some((a) => new RegExp(`(?:[A-Za-z_][\\w.-]*:)?${a}\\s*=`).test(slot.xml));
}

/** The character style a run names (`w:rPr/w:rStyle w:val`), read off the set the run carries. */
function runStyleRef(slots: PropSlot[] | undefined): string | undefined {
  const xml = slots?.find((s) => s.tag === "rStyle")?.xml;
  const val = xml ? /(?:[A-Za-z_][\w.-]*:)?val="([^"]*)"/.exec(xml)?.[1] : undefined;
  return val ? val : undefined;
}

/** An on/off property: present means true unless the file says otherwise (`w:noWrap`, `w:cantSplit`). */
function onOff(node: any): boolean | undefined {
  if (node === undefined) return undefined;
  const v = String(node["@_val"] ?? "1").toLowerCase();
  return v !== "0" && v !== "false" && v !== "none";
}

function str(v: unknown): string | undefined {
  return v === undefined || v === null ? undefined : String(v);
}

/** The one element with this local name in a scanned property set, `""` when the set is not there. */
function elementOf(spans: { tag: string; xml: string }[], tag: string): string {
  return spans.find((s) => s.tag === tag)?.xml ?? "";
}

/**
 * The children of an element reached by name through its wrappers, each with the text the file wrote for it.
 * A `w:rPrChange` sits one level deeper than its run's own properties, and reading it off the parsed tree
 * would answer with the fields the model has rather than with the set the record holds.
 */
function spansUnder(xml: string, ...path: string[]): { tag: string; xml: string }[] {
  let current = xml;
  for (const tag of path) current = elementOf(childSpans(elementContent(current)), tag);
  return childSpans(elementContent(current));
}

/**
 * The attributes on an element's own opening tag, copied as the part wrote them: ` w14:paraId="1A2B3C4D"
 * w:rsidR="00A1234"`. Empty for a tag that carries nothing. This is text and not a name-and-value record
 * because the prefix it spells is bound elsewhere — `w14` only means the words document namespace where an
 * `xmlns:w14` says so — and because a save that puts these back on the same tag has to put them back with
 * the same prefix, in the same order, and with the same escaping it was given.
 */
function ownAttrs(xml: string | undefined): string {
  if (!xml) return "";
  const gt = xml.indexOf(">");
  if (gt < 0) return "";
  const head = xml.slice(1, gt);
  const space = head.search(/\s/);
  if (space < 0) return "";
  const attrs = head.slice(space + 1).trim();
  const text = attrs.endsWith("/") ? attrs.slice(0, -1).trim() : attrs;
  return text ? ` ${text}` : "";
}

/**
 * The attributes on a `w:hyperlink`'s own opening tag that this model does not interpret — `w:history`,
 * `w:docProp`, `w:tm`, `w:invalid` — kept as written, in the file's order and with the file's prefixes.
 * The two spellings of *where the link points* (`r:id`, `w:anchor`) are left out because the model owns
 * those: carrying them would let the saved file point somewhere the model no longer holds. Both quote
 * styles are read because XML allows either, and a saved file has to be able to spell them back as given.
 */
function linkOwnAttrs(span: string | undefined): string | undefined {
  const head = ownAttrs(span).trim();
  if (!head) return undefined;
  const kept: string[] = [];
  for (const match of head.matchAll(/(?:[\w.-]+:)?[\w.-]+\s*=\s*(?:"[^"]*"|'[^']*')/g)) {
    const eq = match[0]!.indexOf("=");
    const name = match[0]!.slice(0, eq).trim();
    const local = name.includes(":") ? name.slice(name.indexOf(":") + 1) : name;
    if (local === "id" || local === "anchor") continue;
    kept.push(match[0]!);
  }
  return kept.length ? kept.join(" ") : undefined;
}

/**
 * The attributes on an element's own opening tag, keyed by local name: `type`, `id`, `val`. Which one
 * comes first and what prefix it wears are the file's choice — WPS writes `r:id` then `w:type` where Word
 * writes the other way round — so an element that names a relationship is looked up by name, not by the
 * position a particular producer happened to put it in.
 */
function attrsOf(xml: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of ownAttrs(xml).matchAll(/(?:^|\s)(?:[\w.-]+:)?([\w.-]+)\s*=\s*"([^"]*)"/g)) out[m[1]!] = m[2]!;
  return out;
}

/**
 * The attributes of one element's own opening tag, with the named values moved. An attribute the file already
 * spelled keeps the position and the prefix it was read with and only changes its value, so a section written
 * with `w:top` before `w:left` still reads that way after a page-setup change; `null` drops one, which is how
 * turning back to portrait removes `w:orient` rather than spelling a value no file measured here uses (the 86
 * wrote `w:orient` only for landscape and `w:orient="portrait"` nowhere; neither is re-counted on the authorized
 * 46 份, so it is 待重证). A name
 * the file never spelled is added at the place ECMA-376 gives it among its neighbours (`order`), with the
 * prefix borrowed from the element's own attributes rather than assumed.
 *
 * `xmlns` declarations are kept byte for byte and are never renamed: their local half (`xmlns:w` → `w`) would
 * otherwise collide with the `w:w` of a `w:pgSz` and rewrite a namespace as a page width.
 */
function patchElementAttrs(xml: string, fallbackPrefix: string, updates: Record<string, string | null>, order: string[]): string {
  const gt = xml.indexOf(">");
  if (gt < 0) return xml;
  const head = xml.slice(1, gt);
  const space = head.search(/\s/);
  const name = space < 0 ? head : head.slice(0, space);
  const spelled = ownAttrs(xml).trim();
  const given: { key: string; name: string; text: string }[] = [];
  for (const match of spelled.matchAll(/(?:[\w.-]+:)?[\w.-]+\s*=\s*"[^"]*"/g)) {
    const eq = match[0]!.indexOf("=");
    const attr = match[0]!.slice(0, eq).trim();
    const local = attr.includes(":") ? attr.slice(attr.indexOf(":") + 1) : attr;
    given.push({ key: attr.startsWith("xmlns") ? attr : local, name: attr, text: match[0]! });
  }
  // A prefix written on one of the element's own value attributes is the prefix its names use. An `xmlns`
  // declaration is not evidence about which alias the words namespace got here, and neither is a colon inside
  // a quoted value, so both are passed over.
  const prefixed = given.find((a) => !a.name.startsWith("xmlns") && a.name.includes(":"));
  const prefix = prefixed ? prefixed.name.slice(0, prefixed.name.indexOf(":")) : fallbackPrefix;
  const rank = (key: string): number => {
    const at = order.indexOf(key);
    return at < 0 ? order.length : at;
  };
  // One pass over what the file spelled: named here, its value moves; named as `null`, it goes; anything
  // else the file wrote stays exactly as it was read.
  const kept = given
    .filter((a) => updates[a.key] !== null)
    .map((a) => ({ key: a.key, text: a.key in updates ? `${a.name}="${updates[a.key]}"` : a.text }));
  for (const [key, value] of Object.entries(updates)) {
    if (value === null || kept.some((a) => a.key === key)) continue;
    const text = `${prefix}:${key}="${value}"`;
    const at = kept.findIndex((a) => rank(a.key) > rank(key));
    if (at < 0) kept.push({ key, text });
    else kept.splice(at, 0, { key, text });
  }
  const attrs = kept.length ? ` ${kept.map((a) => a.text).join(" ")}` : "";
  const closed = xml[gt - 1] === "/";
  return `<${name}${attrs}${closed ? "/>" : `>`}${xml.slice(gt + 1)}`;
}

/**
 * The order ECMA-376 §17.3 (CT_SectPr) allows a section's children in, which is also the order the 115 sections
 * of the quarantined 86 use: references, then `w:pgSz`, then `w:pgMar`, then the column and grid settings. The
 * authorized 46 份 have 58 `w:sectPr` and the re-measure did not walk their child order, so that half is 待重证 —
 * the sequence below is the schema's either way. A geometry element a minimal file never spelled goes in at its
 * place in this sequence rather than at the end, because a `w:pgMar` written after `w:docGrid` is a file Word
 * refuses to open.
 */
const SECT_CHILD_ORDER = [
  "headerReference",
  "footerReference",
  "footnotePr",
  "endnotePr",
  "type",
  "pgSz",
  "pgMar",
  "paperSrc",
  "pgBorders",
  "lnNumType",
  "pgNumType",
  "cols",
  "formProt",
  "vAlign",
  "noEndnote",
  "titlePg",
  "textDirection",
  "bidi",
  "rtlGutter",
  "docGrid",
  "printerSettings",
];

/** The attributes `w:pgSz` spells, in the order the quarantined 86 wrote them (待重证 on the authorized 46 份). */
const PG_SZ_ATTRS = ["w", "h", "orient"];
/** The attributes `w:pgMar` spells, in §17.3.2.13's order — the authorized 46 份 state all seven in every one of
 *  their 58 `w:pgMar`; the order they state them in is 待重证. */
const PG_MAR_ATTRS = ["top", "right", "bottom", "left", "header", "footer", "gutter"];

/** The page box 「页面设置」 asks for: the sheet wanted (or the one the section already had), turned as asked. */
function pageBoxOf(set: PageSetup, widthTwips: number, heightTwips: number): { w: number; h: number } {
  // The two numbers say which sheet, `landscape` only says which way it lies — so the long edge lands in `w:w`
  // for landscape and in `w:h` for portrait. A dialog that sets Letter while the page lies sideways therefore
  // writes 15840×12240, the same spelling the quarantined 86's 22 landscape sections used — `w:pgSz` is
  // 待重证 on the authorized 46 份, so this is the schema's reading of "which edge is which", not a count.
  const a = set.paperTwips?.width ?? widthTwips;
  const b = set.paperTwips?.height ?? heightTwips;
  const long = Math.max(a, b);
  const short = Math.min(a, b);
  return { w: set.landscape ? long : short, h: set.landscape ? short : long };
}

/** The seven `w:pgMar` values this setup writes; a distance the dialog left out is not written at all. */
function marginAttrsOf(set: PageSetup): Record<string, string> {
  const attrs: Record<string, string> = {
    top: String(set.topTwips),
    right: String(set.rightTwips),
    bottom: String(set.bottomTwips),
    left: String(set.leftTwips),
  };
  for (const [key, value] of [["header", set.headerTwips], ["footer", set.footerTwips], ["gutter", set.gutterTwips]] as const) {
    if (value !== undefined) attrs[key] = String(value);
  }
  return attrs;
}

/**
 * A section's own bytes with the page geometry written into them. Only `w:pgSz` and `w:pgMar` change, and each
 * is edited where it stands; a section that never spelled one gets it at its place in §17.3's sequence. Every
 * other child keeps its bytes **and** the text written between the children — the 115 sections this machine's WPS
 * wrote in the quarantined 86 have none (待重证 on the authorized 46 份), but a producer that indents its markup
 * should not lose it here, because
 * the change the user asked for was the page, not the file's layout.
 */
function applyPageSetup(xml: string, set: PageSetup): string {
  const prefix = /^<([A-Za-z_][-\w.]*):/.exec(xml)?.[1] ?? "w";
  const content = elementContent(xml);
  const read = (tag: string): Record<string, string> => attrsOf(childSpans(content).find((k) => k.tag === tag)?.xml ?? "");
  // A section that spelled no page box is not a page of unknown size: this is the box the app drew it with.
  const sz = read("pgSz");
  const box = pageBoxOf(set, Number(sz["w"]) || DEFAULT_SECTION.pageWidthTwips, Number(sz["h"]) || DEFAULT_SECTION.pageHeightTwips);
  const size: Record<string, string | null> = { w: String(box.w), h: String(box.h), orient: set.landscape ? "landscape" : null };
  const mar: Record<string, string | null> = marginAttrsOf(set);
  const written = (tag: string): Record<string, string | null> => (tag === "pgSz" ? size : mar);
  const order = (tag: string): string[] => (tag === "pgSz" ? PG_SZ_ATTRS : PG_MAR_ATTRS);

  // Walk the children in the order they were read, keeping the bytes between them.
  const kids: { tag: string; xml: string }[] = childSpans(content);
  const pieces: { tag: string; xml: string; before: string }[] = [];
  let rest = content;
  for (const k of kids) {
    const at = rest.indexOf(k.xml);
    if (at < 0) continue;
    pieces.push({ tag: k.tag, xml: k.tag === "pgSz" || k.tag === "pgMar" ? patchElementAttrs(k.xml, prefix, written(k.tag), order(k.tag)) : k.xml, before: rest.slice(0, at) });
    rest = rest.slice(at + k.xml.length);
  }
  for (const tag of ["pgSz", "pgMar"] as const) {
    if (pieces.some((p) => p.tag === tag)) continue;
    const attrs = Object.entries(written(tag))
      .filter(([, v]) => v !== null)
      .sort((a, b) => order(tag).indexOf(a[0]) - order(tag).indexOf(b[0]))
      .map(([k, v]) => `${prefix}:${k}="${v}"`)
      .join(" ");
    const at = pieces.findIndex((p) => SECT_CHILD_ORDER.indexOf(p.tag) > SECT_CHILD_ORDER.indexOf(tag));
    const el = { tag, xml: `<${prefix}:${tag} ${attrs}/>`, before: at < 0 ? "" : pieces[at]!.before };
    if (at >= 0) pieces[at] = { ...pieces[at]!, before: "" };
    if (at < 0) pieces.push(el);
    else pieces.splice(at, 0, el);
  }
  const gt = xml.indexOf(">");
  const open = xml[gt - 1] === "/" ? `${xml.slice(0, gt - 1)}>` : xml.slice(0, gt + 1);
  const name = /^<([^/\s>]+)/.exec(xml)?.[1] ?? "";
  return `${open}${pieces.map((p) => p.before + p.xml).join("")}${rest}</${name}>`;
}

/**
 * The children of one `w:trPr` as slots, each with its own source text. `w:trPr` is the only property set
 * a save rewrites child by child: the row's signature is decided one at a time, and the neighbour that
 * stays has to stay where the file put it. The signature itself is not a slot — it leaves the set here and
 * comes back as `RowInfo.rev`, so a row has one record of its own tracked change, not two.
 */
function slotsOf(spans: { tag: string; xml: string }[]): PropSlot[] | undefined {
  const kept = spans.filter((s) => !ROW_REVISION_TAGS.has(s.tag));
  return kept.length ? kept.map((s) => ({ tag: s.tag, xml: s.xml })) : undefined;
}

/**
 * The `w:pPr` children that travel as text: everything except the record of a formatting change, which the
 * `revFormat` field spells so a decision can drop it. `w:pPr/w:rPr` is carried whole — the paragraph mark's
 * fonts, bold, size and colour are the model's blind spot, and losing them on a save is how a document gets
 * shorter lines than it came in with.
 */
function paraSlotsOf(spans: { tag: string; xml: string }[]): PropSlot[] {
  return spans.filter((s) => s.tag !== "pPrChange").map((s) => ({ tag: s.tag, xml: s.xml }));
}

/**
 * The `w:rPr` children that travel as text: everything except the record of a formatting change, which
 * `revFormat` spells so a decision can drop it. `w:rStyle`, `w:lang`, `w:kern`, `w:bCs`, `w:shd` and
 * `w:eastAsianLayout` are all in the set and all outside the model's eleven fields. `w:spacing` (字符间距) is in
 * the set too and *inside* them since round #174: it travels as the file wrote it until the field moves, and then
 * `carriedRunProps` drops the child and re-spells it — a field with a slot is not a field copied as bytes.
 */
function runSlotsOf(spans: { tag: string; xml: string }[]): PropSlot[] {
  return spans.filter((s) => s.tag !== "rPrChange").map((s) => ({ tag: s.tag, xml: s.xml }));
}

/** The order `CT_PPr` gives its children; Word repairs a set written out of order and validators do not. */
const PPR_ORDER = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers",
  "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct",
  "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
  "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
  "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
];
const pprRank = (tag: string): number => {
  const at = PPR_ORDER.indexOf(tag);
  return at < 0 ? PPR_ORDER.length : at;
};

/**
 * Putting a paragraph's `w:pPr` slots in line with the properties the model spells for itself, so a toolbar
 * click on alignment, a list button or the 段落 dialog lands *and* the properties the model cannot explain stay.
 * This is the same split `applyCellProps` makes for a cell's `w:tcPr`: the set stays the file's, only the
 * named child moves.
 *
 * A prefix is copied from the set rather than assumed: `w` is only an alias for the words namespace where the
 * part binds it, and a child written with a different prefix than its neighbours would read as another element.
 *
 * `w:spacing` and `w:ind` are re-spelled attribute by attribute rather than replaced, because both state more
 * than the model has fields for. `w:ind` holds six of them (`left`, `leftChars`, `right`, `firstLine`,
 * `firstLineChars`, `hanging`) and leaves the character spellings it does not hold (`w:rightChars`,
 * `w:hangingChars`) in place with the file's own bytes. `w:spacing` holds four (`w:before`, `w:after`,
 * `w:line`, `w:lineRule`) and the child usually states more. On the authorized set, 3,948 `w:spacing` elements sit in
 * a `w:pPr` block — over `word/document.xml` + `word/footnotes.xml` + `word/endnotes.xml` of the 46 份, the split
 * `.scratch/spacing-part-scope.py` prints (`word/styles.xml`'s 246 are a style's own paragraph properties, and the
 * 页眉页脚 parts write none) (`.scratch/corpus-recheck2.mjs`): 18 of them name `w:beforeLines` and 64 `w:afterLines` —
 * the same gap in 1/100 行 — and 50 name both `w:beforeAutospacing` and `w:afterAutospacing`. Dropping those on a
 * 段前 click is how a save
 * quietly rewrites a paragraph the user only nudged.
 *
 * The two spellings of one gap must not be left to disagree: readers honour `w:beforeLines` over `w:before`,
 * so writing only the absolute would make the click invisible in WPS. Where the element states both, their
 * quotient is the document's own 行 pitch — on the authorized set that is 312 缇 in 68 of the 80 elements that state
 * a pair (段前 16 枚、段后 64 枚, `.scratch/corpus-recheck2.mjs`); whether the two gaps inside one element ever
 * disagree is 待重证 (the old reading said 1,865 elements, never once) — and that pitch re-spells the 行 count in
 * step. The 2 elements that state a 行 count and no absolute (`w:beforeLines` alone) give no pitch, so the one the
 * user just replaced is the one that goes: a paragraph the page has to draw from a stated distance.
 *
 * `w:pBdr` 与 `w:shd` 在这一条上与 `w:spacing` 是同一族，只是**逐边**而非逐属性：一枚 `w:pBdr` 里几条边各说各的
 * （§17.3.2.36 的孩子是 `CT_PBdr` 的六条边），所以只有这一击动过的那几边由字段重写，其余的边按件的字节带着走，
 * 连它身上模型没有字段的那些属性（`w:themeColor`/`w:frame`）一起留着。#144 之前这里**不认领**这两枚孩子，
 * 于是「画对了、存下去没了」——见 `paraBorderChanges`。
 */
export function syncParaSlots(p: Paragraph): void {
  const slots = p.pPrSlots;
  if (!slots) return;
  const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
  const prefix = /^<([A-Za-z_][-\w.]*):/.exec(slots[0]?.xml ?? "")?.[1] ?? "w";
  const wanted: PropSlot[] = [];
  // A property the paragraph spelled for itself follows the field, because a reviewer's decision or a toolbar
  // click may have moved it. One it only borrowed from its style is written only when it really moved.
  const keep = (tag: string, key: keyof ParaStyle, value: unknown, xml: string): void => {
    if (value === undefined || value === false) return;
    if (slots.some((s) => s.tag === tag) || !samePropValue(p.asRead?.[key], value)) wanted.push({ tag, xml });
  };
  keep("pStyle", "styleId", p.styleId, `<${prefix}:pStyle w:val="${esc(p.styleId ?? "")}"/>`);
  keep("jc", "align", p.align, `<${prefix}:jc w:val="${esc(p.align ?? "")}"/>`);
  keep(
    "numPr",
    "numPr",
    p.numPr,
    p.numPr ? `<${prefix}:numPr><w:ilvl w:val="${p.numPr.ilvl}"/><w:numId w:val="${p.numPr.numId}"/></${prefix}:numPr>` : "",
  );
  const moved = (key: keyof ParaStyle): boolean => !samePropValue(p.asRead?.[key], p[key]);
  // 缩进那一句：六个属性跟着字段走，件里说过而模型没有字段的那些（`w:rightChars`、`w:hangingChars`）留在原位。
  // 一个都没动时整句按原件的字节留着——属性之间怎么空格、值怎么写，都是文件自己的事。
  const IND_ATTRS: [keyof ParaStyle, string][] = [
    ["indentLeft", "left"], ["indentLeftChars", "leftChars"], ["indentRight", "right"],
    ["indentFirstLine", "firstLine"], ["indentFirstLineChars", "firstLineChars"], ["indentHanging", "hanging"],
  ];
  const IND_ORDER = ["left", "leftChars", "right", "rightChars", "firstLine", "firstLineChars", "hangingChars", "hanging"];
  const indSlot = slots.find((s) => s.tag === "ind");
  const indMoved = IND_ATTRS.some(([key]) => moved(key));
  if (indMoved) {
    const written = attrsWritten(indSlot?.xml ?? "");
    for (const [key, name] of IND_ATTRS) {
      const value = p[key] as number | undefined;
      putAttr(written, name, value === undefined ? undefined : String(Math.round(value)), IND_ORDER, prefix);
    }
    wanted.push({ tag: "ind", xml: written.length ? `<${prefix}:ind ${written.map((k) => k.xml).join(" ")}/>` : "" });
  }
  // 段前段后与行距那一句，同一条口径：模型握着的四个属性跟着字段走，件里另写的 `w:beforeLines`/`w:afterLines`
  // （1/100 行）与 `w:*Autospacing` 留在原位。只动**这一击动过的那一笔**——没动的连件里古怪的写法一起照抄。
  const SPACING_ATTRS: [keyof ParaStyle, string][] = [
    ["spacingBefore", "before"], ["spacingAfter", "after"], ["lineSpacing", "line"], ["lineSpacingRule", "lineRule"],
  ];
  const SPACING_ORDER = ["before", "beforeLines", "beforeAutospacing", "after", "afterLines", "afterAutospacing", "line", "lineRule"];
  const spacingSlot = slots.find((s) => s.tag === "spacing");
  const spacingMoved = SPACING_ATTRS.some(([key]) => moved(key));
  if (spacingMoved) {
    const written = attrsWritten(spacingSlot?.xml ?? "");
    // 改之前的那份字节才是件的说法：档位从「件自己两笔一起说」的那一对里读，改完的值里读不出来了。
    const said = new Map(written.map((k) => [k.name, k.value]));
    /** 件认的一行多高（缇）：`w:before`／(`w:beforeLines`/100)。只说行数不说绝对值的那种给不出这一档。 */
    const pitchOf = (abs: string, lines: string): number | undefined =>
      linePitchTwips(said.get(abs), said.get(lines));
    const saidOn = (name: string): boolean => {
      const v = said.get(name);
      return v !== undefined && v !== "0" && v !== "false";
    };
    const putGap = (key: keyof ParaStyle, name: string, linesName: string, autoName: string): void => {
      if (!moved(key)) return;
      const value = p[key] as number | undefined;
      putAttr(written, name, value === undefined ? undefined : String(Math.round(value)), SPACING_ORDER, prefix);
      // 行单位那一笔跟得上就跟上：读者认 `w:beforeLines` 多过认 `w:before`，留着一个对不上的行数，这一击在
      // WPS 里等于没发生。件只说了行数、说不出档位时——授权口径 46 份的 `w:pPr` 里 `w:before`/`w:beforeLines`
      // 成对 16 枚、`w:after`/`w:afterLines` 64 枚，只写行数不写缇数的只有 2 枚（`.scratch/corpus-recheck2.mjs`）
      // ——被替掉的那句就摘掉。
      if (said.has(linesName)) {
        const pitch = pitchOf(name, linesName);
        putAttr(written, linesName, value === undefined || pitch === undefined ? undefined : String(Math.round((value * 100) / pitch)), SPACING_ORDER, prefix);
      }
      // 「按行自动加空」与刚填进去的数不能两头都说；关掉的字面是件自己写的那种（作废口径里 13 枚这么写，
      // `w:*Autospacing` 在授权 46 份上没重测：待重证）。
      if (said.has(autoName) && value !== undefined && saidOn(autoName)) putAttr(written, autoName, "0", SPACING_ORDER, prefix);
    };
    putGap("spacingBefore", "before", "beforeLines", "beforeAutospacing");
    putGap("spacingAfter", "after", "afterLines", "afterAutospacing");
    if (moved("lineSpacing") || moved("lineSpacingRule")) {
      const line = p.lineSpacing as number | undefined;
      putAttr(written, "line", line === undefined ? undefined : String(Math.round(line)), SPACING_ORDER, prefix);
      // 件里 `w:line` 与 `w:lineRule` 从来没单独出现过（授权口径 46 份的 document+footnotes+endnotes 三件里，
      // 挂在真 `w:p` 的 `w:pPr` 下 3,948 枚成对、0 枚只写一边，`.scratch/corpus-recheck2.mjs` 与
      // `.scratch/spacing-part-scope.py`），所以行高连着
      // 它的单位一起写；模型只存了数没存规则时按 §17.3.1.27 的缺省签名 `w:auto`。
      putAttr(written, "lineRule", line === undefined ? undefined : String(p.lineSpacingRule ?? "auto"), SPACING_ORDER, prefix);
    }
    wanted.push({ tag: "spacing", xml: written.length ? `<${prefix}:spacing ${written.map((k) => k.xml).join(" ")}/>` : "" });
  }
  // 「如果定义了文档网格，则对齐到网格」那一句：件里只有两种拼法——说开**裸写**、说关写 `w:val="0"`
  // （授权口径 46 份的 `word/document.xml` 里 406 枚，裸写 180、说关写 `"0"` 226，两笔合起来就是全部，
  // 没有一枚写 `"1"` 或 `"false"`；`.scratch/corpus-recheck2.mjs`）。字段是三态的：
  // `undefined` = 这一段没说（照样式链或缺省的「吸」），
  // 所以只有动过且说得出开/关才整枚换掉；没说过的原件一个字都不动。
  // 「允许西文在单词中间换行」（`w:wordWrap`）用与下一支同样的三态笔，只是缺省相反：这一枚**只有写了才算说话**，
  // 而授权口径里写了的 200 枚枚枚裸写（=允许）、说关 0 枚（`.scratch/cjk-typo-audience.py`，正文 7,357 枚 `w:pPr`）。
  // 说关的拼法照 §17.4.77 的 `w:val="0"` 补齐，为的是让对话框那一次取消也落得了地，不是从件里抄来的形状。
  const wrapMoved = moved("wordWrap") && p.wordWrap !== undefined;
  if (wrapMoved) {
    wanted.push({
      tag: "wordWrap",
      xml: p.wordWrap ? `<${prefix}:wordWrap/>` : `<${prefix}:wordWrap ${prefix}:val="0"/>`,
    });
  }
  const snapMoved = moved("snapToGrid") && p.snapToGrid !== undefined;
  if (snapMoved) {
    wanted.push({
      tag: "snapToGrid",
      xml: p.snapToGrid ? `<${prefix}:snapToGrid/>` : `<${prefix}:snapToGrid ${prefix}:val="0"/>`,
    });
  }
  // 「换行和分页」那四枚（孤行控制 `w:widowControl` §17.3.2.90、与下段同页 `w:keepNext` §17.3.1.14、段中不分页
  // `w:keepLines` §17.3.1.13、段前分页 `w:pageBreakBefore` §17.3.1.65）与上面两枚开关同一支三态笔：字段
  // `undefined` = 这一段没说（照样式链与缺省），`true` 裸写，`false` 写 `w:val="0"`。说关的拼法不是从 ECMA 的
  // 句法推的：授权 46 份正文里 180 枚 `w:pageBreakBefore` 全写 `w:val="0"`、0 枚写开（#166 量过），
  // 而三条保持规则在段自己那一层也全是 `w:val="0"`（`w:keepNext` 200 枚、`w:keepLines` 200 枚、
  // `w:widowControl` 160 枚，`.scratch/toggle-lent-chain.ts`）。没动过的段落一个字节不动：`moved()`
  // 只认字段与入件时那一份不同，从样式链借来的话照旧归样式。
  const KEEP_SWITCHES: [keyof ParaStyle, string][] = [
    ["keepNext", "keepNext"], ["keepLines", "keepLines"], ["widowControl", "widowControl"], ["pageBreakBefore", "pageBreakBefore"],
  ];
  /**
   * 槽位此刻说的那个布尔：裸写 = 开，`w:val="0"/"false"/"off"` = 关；没有这一枚元素 = 这一段没说。
   * 认领的判据是「字节与字段说的不是同一件事」，不是「字段动过」：同一趟编辑里先勾上再取消，槽位已经在段上了，
   * 只比 `moved()` 会把它留在原处（勾子撤回了，纸上却还断着页）；反过来，件自己写 `w:val="false"` 而字段也是关，
   * 两者同义，这一枚一个字节都不该重写。
   */
  const slotSwitch = (tag: string): boolean | undefined => {
    const xml = slots.find((s) => s.tag === tag)?.xml;
    if (xml === undefined) return undefined;
    const val = /(?:[A-Za-z_][\w.-]*:)?val="([^"]*)"/.exec(xml)?.[1];
    return val === undefined ? true : !(val === "0" || val === "false" || val === "off");
  };
  const keepMoved = new Set<string>();
  for (const [key, tag] of KEEP_SWITCHES) {
    const value = p[key];
    if (value === undefined) continue;
    const stated = slotSwitch(tag);
    // 段上已经有这一枚：与**字节**比，写的是不是同一件事（这样同一趟里勾上再取消看得见，件自己的
    // `w:val="false"` 也不会被当成「说了开」重写一遍）。段上还没有这一枚：与入件时画着的那一份比——
    // 从样式链借来的「开」要在段落上写一句关才算数，而缺省本来就是关的那几枚（`w:pageBreakBefore`）
    // 不该替段落多写一枚 `w:val="0"`。
    if (stated === undefined ? value !== (p.asRead?.[key] === true) : stated !== value) keepMoved.add(tag);
  }
  for (const [key, tag] of KEEP_SWITCHES) {
    if (!keepMoved.has(tag)) continue;
    wanted.push({ tag, xml: p[key] ? `<${prefix}:${tag}/>` : `<${prefix}:${tag} ${prefix}:val="0"/>` });
  }
  // 框：逐边对着读快照比，只有动过的那几边重写；没动的边——这一段自己写的、还有从样式链借来的——一个字节不动。
  // 借来的那一份尤其要说清：`asRead.borders` 里就有它（入件时页面上画的那一整套），把它抄进段落自己的
  // `w:pBdr` 就是替文件编一句「这一段说过框」，而它说的只有被改过的那一边。
  const box = moved("borders") ? paraBorderChanges(p) : undefined;
  if (box) {
    wanted.push({ tag: "pBdr", xml: paraBorderSlotXml(slots.find((s) => s.tag === "pBdr")?.xml, box, prefix) });
  }
  // 地：字段只装着 `w:fill` 那一句，图案（`w:val`）与边线色（`w:color`）是件的说法，原样留着；件没写过这一枚
  // 孩子时才按授权语料的形状补上（46 份里 30 枚段落底纹，`.scratch/pbdr-audience.ts`；这 30 枚连同其余各层，
  // 1,676 枚 `w:shd` 无一例外写 `w:val="clear"`（`.scratch/shd-audience.ts`），而那句 `w:color="auto"` 只在作废
  // 口径的 30/30 里取过证，授权 46 份上是待重证）。
  const shdFill = moved("shading") ? p.shading : undefined;
  const shdMoved = shdFill !== undefined;
  if (shdMoved) {
    wanted.push({ tag: "shd", xml: shadingSlotXml(slots.find((s) => s.tag === "shd")?.xml, shdFill, prefix) });
  }
  // 断页那一枚也只在这支三态笔说了话时认领：`keep` 对 false 一言不发，若这里照样认领，件自己那句
  // `<w:pageBreakBefore w:val="0"/>`（授权 46 份里 180 枚）会在认领时连同槽位一起消失。
  const owned = new Set(["pStyle", "numPr", "jc", ...keepMoved,
    ...(indMoved ? ["ind"] : []), ...(spacingMoved ? ["spacing"] : []),
    ...(wrapMoved ? ["wordWrap"] : []),
    ...(snapMoved ? ["snapToGrid"] : []), ...(box ? ["pBdr"] : []), ...(shdMoved ? ["shd"] : [])]);
  let next = slots.filter((s) => !owned.has(s.tag));
  for (const w of wanted) {
    const at = next.findIndex((s) => pprRank(s.tag) > pprRank(w.tag));
    next = at < 0 ? [...next, w] : [...next.slice(0, at), w, ...next.slice(at)];
  }
  p.pPrSlots = next.length ? [...next].sort((a, b) => pprRank(a.tag) - pprRank(b.tag)) : undefined;
}

/** `CT_Border` 的属性序（§17.3.2.36 的每一条边）：模型认前四枚，后四枚是件的，改这一边时也留在原位。 */
const BORDER_ATTR_ORDER = ["val", "sz", "space", "color", "themeColor", "themeShade", "themeTint", "frame"];
/** `CT_Shd` 的属性序（§17.13.12）：授权口径 46 份里 30/30 枚段落底纹只写头一枚与末两枚，中间那四枚主题色 0 枚。 */
const SHD_ATTR_ORDER = ["val", "color", "themeColor", "themeFill", "themeColorTint", "themeColorShade", "fill"];

/**
 * 一枚框槽（`w:pBdr` / `w:tblBorders` / `w:tcBorders`）里已有的孩子，按标签索引、连件的字节一起留着。
 * 按 `childSpans` 拆而不按正则：`CT_Border` 只有属性，可 `<w:top w:val="single"></w:top>` 这一样也合规，
 * 而重建是把槽里的孩子**逐枚摆回去**的——正则只认自闭合那种，遇到成对写法就会把这一边整枚抹掉。
 */
function borderSidesOf(xml: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const kid of childSpans(elementContent(xml))) out.set(kid.tag, kid.xml);
  return out;
}

/**
 * 这一段这一击**动过的边**：对着读快照逐边比，内容没变的边不说。`asRead.borders` 是入件时页面上画的那一整套
 * （含样式链借来的），所以「没变」的那几边就算报出去也只是把样式的话抄到段落身上；而 §17.7.2 是**逐边**解决的，
 * 只报动过的那几条在语义上完整——没报的边继续从样式借。没有快照的段（编辑器自己拆出来的）每一边都算动过。
 */
export function paraBorderChanges(p: Paragraph): ParaBorderSet | undefined {
  const now = p.borders;
  if (!now) return undefined;
  const was = p.asRead?.borders;
  const out: ParaBorderSet = {};
  let any = false;
  for (const side of PBDR_ORDER) {
    const edge = now[side];
    if (!edge || (was && samePropValue(was[side], edge))) continue;
    out[side] = edge;
    any = true;
  }
  return any ? out : undefined;
}

/** 一边的字面：件写过的属性留着，模型认得的那四枚跟着字段走。 */
function borderSideXml(side: string, base: string | undefined, edge: TableBorder, prefix: string): string {
  const written = attrsWritten(base ?? "");
  putAttr(written, "val", edge.val, BORDER_ATTR_ORDER, prefix);
  putAttr(written, "sz", edge.size === undefined ? undefined : String(edge.size), BORDER_ATTR_ORDER, prefix);
  putAttr(written, "space", edge.space === undefined ? undefined : String(edge.space), BORDER_ATTR_ORDER, prefix);
  putAttr(written, "color", edge.color, BORDER_ATTR_ORDER, prefix);
  return `<${prefix}:${side}${written.length ? ` ${written.map((k) => k.xml).join(" ")}` : ""}/>`;
}

/**
 * 重建一枚框槽：`changes` 点名的边由字段重写，没点名的边（含这一段/这一张表自己写的其余边）按件的字节带着走。
 * 孩子按 `sides` 的边序摆，`sides` 之外还留在件里的那些排在后面——槽里没有模型不认的孩子时，这两段是同一件事。
 */
function borderSlotXml(
  tag: string,
  own: string | undefined,
  changes: Record<string, TableBorder | undefined>,
  sides: readonly string[],
  prefix: string,
): string {
  const base = borderSidesOf(own ?? "");
  for (const side of sides) {
    const edge = changes[side];
    if (edge) base.set(side, borderSideXml(side, base.get(side), edge, prefix));
  }
  const kids = [...sides.filter((side) => base.has(side)), ...[...base.keys()].filter((side) => !sides.includes(side))].map(
    (side) => base.get(side)!,
  );
  return kids.length ? `<${prefix}:${tag}>${kids.join("")}</${prefix}:${tag}>` : "";
}

/** 重建这一段那一枚 `w:pBdr`：动过的边由字段重写，其余按件的字节走。 */
function paraBorderSlotXml(own: string | undefined, changes: ParaBorderSet, prefix: string): string {
  return borderSlotXml("pBdr", own, changes, PBDR_ORDER, prefix);
}

/**
 * 底纹那一枚孩子（`w:pPr`、`w:tblPr`、`w:tcPr` 三层共用同一条写法）：只有 `w:fill` 跟着字段走，
 * `w:val`/`w:color` 是件的说法，件没写过时按授权语料的形状补——46 份件里 1,676 枚 `w:shd` 全是
 * `w:val="clear" w:color="auto"` 加一个 `w:fill`（`.scratch/shd-audience.ts`），没有第二种组合。
 */
function shadingSlotXml(own: string | undefined, fill: string, prefix: string): string {
  const written = attrsWritten(own ?? "");
  const said = new Map(written.map((k) => [k.name, k.value]));
  putAttr(written, "val", said.get("val") ?? "clear", SHD_ATTR_ORDER, prefix);
  putAttr(written, "color", said.get("color") ?? "auto", SHD_ATTR_ORDER, prefix);
  putAttr(written, "fill", shadingFillToken(fill), SHD_ATTR_ORDER, prefix);
  return `<${prefix}:shd ${written.map((k) => k.xml).join(" ")}/>`;
}

/**
 * `w:shd@w:fill` 的写法：模型里 `#RRGGBB` 是读端给一枚颜色画的形状，回到件里是六位大写、不带 `#`；
 * 别的值（`auto`）是文件在颜色的位置上写的那个 token，转成大写就是替件签一个 schema 里没有的词（`w:fill="AUTO"`）。
 */
export const shadingFillToken = (v: string): string => {
  const bare = v.startsWith("#") ? v.slice(1) : v;
  return /^[0-9A-Fa-f]{6}$/.test(bare) ? bare.toUpperCase() : v;
};

/**
 * `w:lineRule="auto"` 那一族里 `w:line` 的单位：240 分之一行（§17.3.1.27）。`exact`/`atLeast` 的两笔按缇数，
 * 与这一个数无关——同一个字段两种单位，所以线（渲染）、格（对话框）与下拉（ribbon）三处必须共用这一个换算。
 */
export const LINE_ONE = 240;

/**
 * 件的同一句 `w:spacing` 里两笔一起说时，它们的商就是这个文档认的「一行多高」（缇）：`w:before`／(`w:beforeLines`/100)。
 * 授权口径 46 份的 `w:pPr` 里这样成对写下的有段前 16 枚、段后 64 枚，商落在 312 缇的是其中 68 枚
 * （`w:type="lines"` 的 `w:docGrid` 那一档，见 #130；`.scratch/corpus-recheck2.mjs`）。只说一边说不出来。
 */
export const linePitchTwips = (absTwips: string | undefined, lines: string | undefined): number | undefined => {
  const a = Number(absTwips);
  const l = Number(lines);
  return Number.isFinite(a) && a > 0 && Number.isFinite(l) && l > 0 ? (a * 100) / l : undefined;
};

/**
 * 这一段自己那句 `w:spacing` 里件认的一行多高（缇），给对话框把「0.5 行」这样的字面摆出来用。只读这一段的
 * `w:pPr` 直接孩子：段落标记 `w:rPr` 里那个同名的 `w:spacing` 说的是字符间距（授权口径 46 份的正文部件里
 * `w:rPr/w:spacing@w:val` 只有 14 枚，`.scratch/corpus-recheck2.mjs`），它整枚待在 `rPr` 槽的字节里，
 * 不是这里的输入。
 */
export function paraSpacingPitch(p: Paragraph): number | undefined {
  const xml = p.pPrSlots?.find((s) => s.tag === "spacing")?.xml ?? "";
  const get = (name: string): string | undefined =>
    new RegExp(`(?:^|\\s)(?:[A-Za-z_][-\\w.]*:)?${name}\\s*=\\s*"([^"]*)"`).exec(xml)?.[1];
  return linePitchTwips(get("before"), get("beforeLines")) ?? linePitchTwips(get("after"), get("afterLines"));
}

/** 一个属性集里已经写着的属性：名字去掉前缀，字节连前缀一起照抄（`w:rightChars` 那种没字段的就留在原位）。 */
function attrsWritten(xml: string): { name: string; xml: string; value: string }[] {
  return [...xml.matchAll(/([A-Za-z_][-\w.]*:)?([\w-]+)\s*=\s*"([^"]*)"/g)].map((m) => ({
    name: m[2]!,
    xml: `${m[1] ?? ""}${m[2]}="${m[3]}"`,
    value: m[3]!,
  }));
}

/**
 * 把模型握着的那一个属性摆进文件自己的序列：`spell` 为空就是这一笔不要了；原来写着就原地换值，没写过就按
 * `order` 的先后插进第一个比它晚的属性前面——顺序是件定的。作废口径里那一簇（201 份的 10,248 枚 `w:spacing`
 * 全落在 `before, beforeLines, beforeAutospacing, after, afterLines, afterAutospacing, line, lineRule` 这一条上）
 * 的属性先后没有重测，是待重证；授权 46 份能给出的是形状本身：`word/document.xml` + `word/footnotes.xml` +
 * `word/endnotes.xml` 三件里，挂在真 `w:p` 的 `w:pPr` 下成对写 `w:line`+`w:lineRule` 的 3,948
 * 枚中 3,696 枚只写这一对，全部属性组合共 8 种（`.scratch/corpus-recheck2.mjs`；部件拆分见
 * `.scratch/spacing-part-scope.py`：3,924 + 12 + 12）。
 */
function putAttr(
  written: { name: string; xml: string; value: string }[],
  name: string,
  spell: string | undefined,
  order: readonly string[],
  prefix: string,
): void {
  const at = written.findIndex((k) => k.name === name);
  if (spell === undefined) {
    if (at >= 0) written.splice(at, 1);
    return;
  }
  if (at >= 0) written[at] = { name, xml: `${prefix}:${name}="${spell}"`, value: spell };
  else {
    const later = written.findIndex((k) => order.indexOf(k.name) > order.indexOf(name));
    if (later < 0) written.push({ name, xml: `${prefix}:${name}="${spell}"`, value: spell });
    else written.splice(later, 0, { name, xml: `${prefix}:${name}="${spell}"`, value: spell });
  }
}


/**
 * 一件属性集（`w:tblPr` / `w:tcPr` / `w:sectPr` 那一类）里换掉**一枚孩子**：原有的整枚原地替换，没有的按
 * `order` 的序插进第一枚比它晚的孩子前面，给空串就是摘掉这一枚。其余孩子与孩子**之间**的文字按原件带着走，
 * 与 `applyPageSetup` 对 `w:sectPr` 做的事同一条口径——用户要改的是这一枚孩子，不是这个文件的排版。
 */
function replaceElementChild(setXml: string, tag: string, xml: string, order: readonly string[]): string {
  const rank = (t: string): number => {
    const at = order.indexOf(t);
    return at < 0 ? order.length : at;
  };
  const content = elementContent(setXml);
  const pieces: { tag: string; xml: string; before: string }[] = [];
  let rest = content;
  let carry = "";
  let saw = false;
  for (const kid of childSpans(content)) {
    const at = rest.indexOf(kid.xml);
    if (at < 0) continue;
    carry += rest.slice(0, at);
    rest = rest.slice(at + kid.xml.length);
    if (kid.tag === tag) {
      saw = true;
      if (!xml) continue;
      pieces.push({ tag, xml, before: carry });
    } else pieces.push({ tag: kid.tag, xml: kid.xml, before: carry });
    carry = "";
  }
  rest = carry + rest;
  if (xml && !saw) {
    const at = pieces.findIndex((p) => rank(p.tag) > rank(tag));
    const piece = { tag, xml, before: at < 0 ? "" : pieces[at]!.before };
    if (at >= 0) pieces[at] = { ...pieces[at]!, before: "" };
    if (at < 0) pieces.push(piece);
    else pieces.splice(at, 0, piece);
  }
  const body = pieces.map((p) => p.before + p.xml).join("") + rest;
  const name = /^<([^/\s>]+)/.exec(setXml)?.[1] ?? "";
  return body ? `<${name}>${body}</${name}>` : `<${name}/>`;
}

/** 一个孩子集里某一枚孩子的原文，没有就是这一件没写过这一枚。 */
const childOf = (setXml: string, tag: string): string | undefined => childSpans(elementContent(setXml)).find((s) => s.tag === tag)?.xml;

/** 一枚元素自己的名字空间前缀别名：新写的孩子要与同组的邻居用同一个前缀，否则读起来是另一枚元素。 */
const tagPrefix = (xml: string): string => /^<([A-Za-z_][-\w.]*):/.exec(xml)?.[1] ?? "w";

/** `CT_TblPr`（§17.4.60）给孩子的序：新写的一枚按它插，不是摆在末尾。 */
const TBLPR_CHILD_ORDER = [
  "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
  "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption",
  "tblDescription", "tblPrChange",
];
/** `CT_TcPr`（§17.4.73）给孩子的序。 */
const TCPR_CHILD_ORDER = [
  "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
  "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
];
/**
 * `CT_TrPr` 给孩子的序，按件里量到的那一段钉：授权 46 份的 598 枚 `w:trPr` 里说行高的 580 枚**全部排在最前**
 * 或紧跟一枚 `w:cantSplit`（22 枚），而 `w:jc` 190 枚没有一枚走在 `w:trHeight` 之前
 * （`.scratch/trpr-order-audience.py`）。其余位置按 schema 序摆，件里这一族只说到这三枚。
 */
const TRPR_CHILD_ORDER = [
  "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader",
  "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange",
];
/**
 * `CT_TblBorders`（§17.4.76）的孩子序：表比一格多出「内部横/内部竖」两条，斜线两边都有。
 * 与模型那一份 `BORDER_ORDER` 同序同集合——`w:between` 是 `CT_PBdr`（§17.3.2.36）才有的那一枚，
 * 摆进这一族就是在编 schema 里不存在的孩子。
 */
export const TBL_BORDER_SIDES = [...BORDER_ORDER] as const;
/** `CT_TcBorders`（§17.4.77）的孩子序：一格只有自己的四条与两条斜线，说不出「内部」——那是表才有的那一枚孩子。 */
export const TC_BORDER_SIDES = ["top", "start", "left", "bottom", "end", "right", "tl2br", "tr2bl"] as const;

/**
 * 这一击**动过的边**：把补丁前后的两套框逐边比，没变的边不进这一份。表与格没有段落那样的读快照（`asRead`），
 * 参照就是件里带回来的那一份；没点名的边继续按件的字节走，连模型没有字段的那些属性（`w:themeColor`、`w:frame`）
 * 与「写过一条没有 `w:val` 的边」这种形状一起留着——`bordersOf` 只认带 `w:val` 的边，整槽重拼就会替件把那几条抹掉。
 */
export function borderDelta(before: BorderSet | undefined, after: BorderSet | undefined): BorderSet {
  const out: BorderSet = {};
  for (const side of BORDER_ORDER) {
    const edge = after?.[side];
    if (edge && !samePropValue(before?.[side], edge)) out[side] = edge;
  }
  return out;
}

/**
 * 表这一击动过的边写回它自己的 `w:tblPr`，`shading` 给了就把同一击的底纹也写回同一枚 `w:tblPr` 的 `w:shd` 槽。
 * 表若没有 `tblPrXml`（编辑器自己造的表），这里什么都不做：写出器本来就按字段拼 `w:tblBorders` 与 `w:shd`，
 * 这条只管「件自己写过」的那一半。
 * 只点底纹、没动边时，`w:tblBorders` 一个字节也不许多动：那一枚孩子原样留在件里，`replaceElementChild`
 * 只在 `slot` 非空时才碰它——`borderSlotXml` 对「一件没改」的输入会照抄件里已有的那些边，所以这条不会凭空造。
 */
export function syncTableBox(t: Table, changes: BorderSet, shading?: string): void {
  const own = t.tblPrXml;
  if (!own) return;
  let next = own;
  const slot = borderSlotXml("tblBorders", childOf(own, "tblBorders"), changes, TBL_BORDER_SIDES, tagPrefix(own));
  if (slot) next = replaceElementChild(next, "tblBorders", slot, TBLPR_CHILD_ORDER);
  if (shading !== undefined)
    next = replaceElementChild(
      next,
      "shd",
      shadingSlotXml(childOf(next, "shd"), shading, tagPrefix(next)),
      TBLPR_CHILD_ORDER,
    );
  t.tblPrXml = next;
}

/**
 * 「表格属性」那一击碰得到的四句：对齐 `w:jc`、左缩进 `w:tblInd`、固定列宽 `w:tblLayout`、默认单元格边距
 * `w:tblCellMar`。**宽度不在这一份里**：`w:tblW` 已经有一条腿跟着网格走（`setTableWidth`，`doc-table-ops.ts`），
 * 第二个写同一句的地方只会把对方顶掉。
 *
 * 缺的键 = 这一句不改；`""` = 把件里那一句**删掉**，回到「没说」的那一档（左对齐 / 不缩进 / 自动列宽）。
 * 删而不是补一句缺省值：授权 46 份的 58 张表里 `w:tblLayout` 58 枚全是 `fixed`，`w:tblInd` 30 枚写 `0`
 * 而 28 枚写别的数，`autofit` 一枚也没写过（`.scratch/tblpr-children-audience.py`）——替文件写一句它从没写过的
 * `autofit` 就是编。与 `syncTableBox` 同一条纪律：没点名的孩子一个字节不动，动过的按 `CT_TblPr` 的序回原位。
 */
export type TablePropsPatch = {
  align?: "left" | "center" | "right" | "";
  indentTwips?: number | "";
  layout?: "fixed" | "";
  /** 只写点名的边；没点名的边沿用件里那一句（连它的 `w:type` 一起原样留着）。 */
  cellMargin?: MarginTwips;
};

/**
 * 「表格属性」那一击**行**那一段碰得到的句子：行自己那句对齐 `w:trPr/w:jc`（§17.4.75）。
 * 缺的键 = 这一句不改；`""` = 摘掉那一句，回到「没说」那一档（左）。
 *
 * **行高、允许跨页断行、标题行重复不在这一份里**：前两句各有各的腿（`setCellHeight` 落 `w:trHeight`、
 * 表格工具「标题行重复」落 `w:tblHeader`），对话框再摆一格就是两个主人写同一句；`w:cantSplit` 则是
 * 铺排里还没有落笔处——引擎把每一行当成一枚整单位铺（`doc-layout.ts:unitsOfBlock` 那一支），勾与不勾
 * 画出来是同一张纸，摆一格就是写一句纸面不认的字节。
 */
export type RowPropsPatch = { align?: "left" | "center" | "right" | "" };

/**
 * 「表格属性」那一次「确定」的全部落笔：表自己那几句 + 选区盖住的每一行那几句。
 * 合成一份而不是两次点击，是为了撤销栈——对话框一次「确定」在用户手里是一件事，拆成两击要退两次。
 */
export type TablePropsDialogPatch = TablePropsPatch & { rows?: RowPropsPatch };

/** 这四句写回表自己的 `w:tblPr`，模型字段同一步跟上——再画一遍与再存一遍说的得是同一件事。 */
export function syncTableProps(t: Table, patch: TablePropsPatch): void {
  if (patch.align !== undefined) t.align = patch.align || undefined;
  if (patch.indentTwips !== undefined) t.indentTwips = patch.indentTwips === "" ? undefined : patch.indentTwips;
  if (patch.layout !== undefined) t.layout = patch.layout || undefined;
  if (patch.cellMargin !== undefined) t.cellMargin = { ...t.cellMargin, ...patch.cellMargin };
  const own = t.tblPrXml;
  // 编辑器自己造的表没有原件文字：写出器按字段拼（`tablePropsInner`），这里只跟上字段。
  if (!own) return;
  const p = tagPrefix(own);
  let next = own;
  if (patch.align !== undefined)
    next = replaceElementChild(next, "jc", patch.align ? `<${p}:jc ${p}:val="${patch.align}"/>` : "", TBLPR_CHILD_ORDER);
  if (patch.indentTwips !== undefined)
    next = replaceElementChild(
      next,
      "tblInd",
      patch.indentTwips === "" ? "" : `<${p}:tblInd ${p}:w="${Math.round(patch.indentTwips)}" ${p}:type="dxa"/>`,
      TBLPR_CHILD_ORDER,
    );
  if (patch.layout !== undefined)
    next = replaceElementChild(
      next,
      "tblLayout",
      patch.layout ? `<${p}:tblLayout ${p}:type="${patch.layout}"/>` : "",
      TBLPR_CHILD_ORDER,
    );
  if (patch.cellMargin !== undefined) {
    // 四条边一条都没点名（`{}`）时一个字也不碰：`w:tblCellMar` 可能还带着别的说法，重拼就是重写。
    const sides = MARGIN_SIDES.filter((s) => patch.cellMargin![s] !== undefined);
    if (sides.length) next = replaceElementChild(next, "tblCellMar", cellMarSlotXml(childOf(next, "tblCellMar"), patch.cellMargin, p), TBLPR_CHILD_ORDER);
  }
  t.tblPrXml = next;
}

/**
 * `w:tblCellMar` 的一句：件里有的边原样留着（边与边之间的文字也留着），点名的边换数，缺的按
 * `top/left/bottom/right` 的序插进去。新写的一条照件的拼法带 `w:w` 与 `w:type="dxa"`——授权 46 份里
 * 751 枚 `w:top`/`w:bottom` 与 751 枚 `w:left`/`w:right` 全是这两句，不带 `w:type` 的一条也没有
 * （`.scratch/tblpr-children-audience.py` 的同一趟读数）。
 */
function cellMarSlotXml(existing: string | undefined, sides: MarginTwips, p: string): string {
  const sideXml = (name: (typeof MARGIN_SIDES)[number], n: number): string =>
    `<${p}:${name} ${p}:w="${Math.round(n)}" ${p}:type="dxa"/>`;
  if (!existing) {
    const kids = MARGIN_SIDES.filter((s) => sides[s] !== undefined).map((s) => sideXml(s, sides[s]!));
    return `<${p}:tblCellMar>${kids.join("")}</${p}:tblCellMar>`;
  }
  let next = existing;
  for (const side of MARGIN_SIDES) {
    const n = sides[side];
    if (n === undefined) continue;
    next = replaceElementChild(next, side, sideXml(side, n), MARGIN_SIDES);
  }
  return next;
}

/** 一格这一击动过的边与地写回它自己的 `w:tcPr`；与表同一条口径，没有原件的格子由写出器按字段拼。 */
export function syncCellBox(cell: TableCell, changes: BorderSet, shading?: string): void {
  const own = cell.tcPrXml;
  if (!own) return;
  let next = own;
  const slot = borderSlotXml("tcBorders", childOf(own, "tcBorders"), changes, TC_BORDER_SIDES, tagPrefix(own));
  if (slot) next = replaceElementChild(next, "tcBorders", slot, TCPR_CHILD_ORDER);
  if (shading !== undefined)
    next = replaceElementChild(next, "shd", shadingSlotXml(childOf(next, "shd"), shading, tagPrefix(next)), TCPR_CHILD_ORDER);
  cell.tcPrXml = next;
}

/**
 * 一格这一击改的竖向对齐写回它自己的 `w:tcPr`。`w:vAlign` 在原件里是槽（模型没读它之前一直原样带着），
 * 所以这一击是**换槽**而不是追加：同一枚孩子只留一处，位置按 `CT_TcPr` 的孩子序。没有原件的格子由写出器
 * 按 `cell.valign` 拼（`write-docx.ts` 那条早就有）。
 */
export function syncCellVAlign(cell: TableCell, valign: string): void {
  const own = cell.tcPrXml;
  if (!own) return;
  const prefix = tagPrefix(own);
  cell.tcPrXml = replaceElementChild(own, "vAlign", `<${prefix}:vAlign ${prefix}:val="${valign}"/>`, TCPR_CHILD_ORDER);
}

/** 逐枚带着的孩子集里新写的一枚该用哪个名字空间前缀：问集里已有的邻居，没有邻居才按 `w:`。 */
const slotPrefix = (slots: PropSlot[] | undefined): string => (slots?.length ? tagPrefix(slots[0]!.xml) : "w");

/**
 * 逐枚孩子带着的属性集（行的 `w:trPr` 正是这一种，`rowInfoOf` 按枚拆开带）里换掉一枚孩子：同一枚只留一处，
 * 新写的一枚按 `order` 落在第一枚比它晚的邻居之前，其余的位置一个字节不动。与 `replaceElementChild` 同一
 * 分工——那一位管「整枚属性集带在件里」的文字形状，这一位管模型的 `slots` 形状。
 */
function upsertSlot(slots: PropSlot[], tag: string, xml: string, order: readonly string[]): PropSlot[] {
  const rank = (t: string): number => {
    const at = order.indexOf(t);
    return at < 0 ? order.length : at;
  };
  const rest = slots.filter((s) => s.tag !== tag);
  const at = rest.findIndex((s) => rank(s.tag) > rank(tag));
  if (at < 0) return [...rest, { tag, xml }];
  return [...rest.slice(0, at), { tag, xml }, ...rest.slice(at)];
}

/**
 * 行高那一击写回行自己那句 `w:trPr/w:trHeight`：字段是给渲染与铺排读的（`render.ts:rowStyle`、`doc-layout.ts`
 * 按这一枚算行盒），孩子是给「拒绝」与写出读的——`write-docx.ts:rowPropsXml` 只在行**没有**带着的孩子时才按
 * 字段拼，读过件的行是按枚带回来的，光改字段存盘会发现那一句还是件里原来的数。没有孩子的行（编辑器自己造的）
 * 没有字节可换，字段就够了，写出器本来就按它们拼这一枚。
 */
export function syncRowHeight(row: RowInfo, twips: number, rule?: string): void {
  row.heightTwips = twips;
  if (rule) row.heightRule = rule;
  else delete row.heightRule;
  if (!row.slots?.length) return;
  const prefix = slotPrefix(row.slots);
  row.slots = upsertSlot(
    row.slots,
    "trHeight",
    `<${prefix}:trHeight ${prefix}:val="${Math.round(twips)}"${rule ? ` ${prefix}:hRule="${rule}"` : ""}/>`,
    TRPR_CHILD_ORDER,
  );
}

/**
 * 「标题行重复」那一击写回行自己那句 `w:trPr/w:tblHeader`（§17.4.49）。与行高同两分：字段给纸面与铺排读
 * （`render.ts` 盖章、`doc-layout.ts` 在页缝处重画这一行），孩子给写出与「拒绝」读——读过件的行是按枚带回
 * `w:trPr` 孩子的，光改字段存盘会发现那一句还在（或不在）。
 *
 * 关掉时**摘掉**那一枚孩子，不写 `<w:tblHeader w:val="0"/>`：模型里「没说」就是这一句的缺席，`w:val="0"`
 * 读回来也是缺席，两种写法在纸上是同一张纸；写一枚显式的「否」只会多一句没人读的字节。件里原本的
 * `w:val="0"` 因此也一并退成缺席——与 `setCellAlign("left")` 清掉 `w:jc` 同一条口径。
 *
 * 只动这一枚孩子，别的孩子（含别人的待裁决记录）原样留在原位，插入位置按 `CT_TrPr` 的孩子序。
 */
export function syncRowHeader(row: RowInfo, on: boolean): void {
  if (on) row.header = true;
  else delete row.header;
  if (!row.slots?.length) return;
  if (!on) {
    row.slots = row.slots.filter((s) => s.tag !== "tblHeader");
    return;
  }
  const prefix = slotPrefix(row.slots);
  row.slots = upsertSlot(row.slots, "tblHeader", `<${prefix}:tblHeader/>`, TRPR_CHILD_ORDER);
}

/**
 * 行对齐那一击写回行自己那句 `w:trPr/w:jc`（§17.4.75，`ST_JcTable` 只有 `left`/`center`/`right` 三档）。
 * 与行高、表头那两条同两分：字段给纸面读（`render.ts:rowStyle` 把这一句落在 `<tr>` 上，格与段从它继承），
 * 孩子给写出与「拒绝」读——读过件的行是按枚带回 `w:trPr` 孩子的，光改字段存盘会发现那一句还是件里的。
 *
 * `""` = 摘掉那一枚孩子，不写 `<w:jc w:val="left"/>`：这一族缺省就是左，「没说」与「明说 left」在纸上是
 * 同一张纸，补一句文件没说的 `left` 就是替它发明（与 `setCellAlign` 清掉段那句 `w:jc` 同一条口径，
 * 两页对着同一份文档不写出两种字节）。件里那句是认不出的档（如 `start`）时，选 `center` 就是原位换掉它。
 */
export function syncRowAlign(row: RowInfo, align: string): void {
  if (align) row.jc = align;
  else delete row.jc;
  if (!row.slots?.length) return;
  if (!align) {
    row.slots = row.slots.filter((s) => s.tag !== "jc");
    return;
  }
  const prefix = slotPrefix(row.slots);
  row.slots = upsertSlot(row.slots, "jc", `<${prefix}:jc ${prefix}:val="${align}"/>`, TRPR_CHILD_ORDER);
}

/** 一套属性在孩子里说过的那几枚：别人的待裁决不算进这一份快照。 */
const statedChildren = (setXml: string): PropSlot[] =>
  childSpans(elementContent(setXml))
    .filter((k) => !TABLE_RECORD_TAGS.has(k.tag))
    .map((k) => ({ tag: k.tag, xml: k.xml }));

/** 一件部件自己说过的字段——「拒绝」按 `restore*` 把这一份捡回模型。 */
const statedFields = <T extends object>(part: T, keys: readonly string[]): Record<string, unknown> => {
  const out: Record<string, unknown> = {};
  const src = part as Record<string, unknown>;
  for (const key of keys) if (src[key] !== undefined) out[key] = src[key];
  return out;
};

/** 表自己说过的那一整套 `w:tblPr` 字段。 */
export const tableShapeOf = (t: Table): TableStyle => statedFields(t, TABLE_STYLE_KEYS) as TableStyle;
/** 一格自己说过的那一整套 `w:tcPr` 字段。 */
export const cellShapeOf = (c: TableCell): CellStyle => statedFields(c, CELL_STYLE_KEYS) as CellStyle;
/** 一行自己说过的那一整套 `w:trPr` 字段。 */
export const rowShapeOf = (row: RowInfo): RowStyle => statedFields(row, ROW_STYLE_KEYS) as RowStyle;

/**
 * 开着修订改表这一圈（`w:tblBorders`、`w:tblPr/w:shd`）时留下的 `w:tblPrChange`。记的是**改之前那一整套自己
 * 说过**的属性，与段落、run 那两条同口径：从 `w:tblStyle` 借来的不进快照，一张表本来什么也没说时记的是一枚空套，
 * 拒绝就把它放回「借用」那一边。两份都得留：文字那一份挂进带着的 `w:tblPr` 末尾，因为 `decideSetXml` 与写出读的都是
 * 这一枚孩子，而件里没建模的属性（`w:tblLook`、`w:tblCellMar`）连自己的字一起回去；字段那一份是给「拒绝」把属性
 * 捡回模型用的。没有 `w:tblPr` 的表（编辑器自己造的）没有文字可挂，记录只进字段，由写出器按字段拼这一枚孩子。
 * 同一个人记过就不再重签：一条 `w:id` 只裁一次，两击得一起退。
 */
export function signTableRevision(t: Table, sign: { id: number; author: string; date?: string }, was: TableStyle, wasXml?: string): TableFormatRevision {
  const mine = t.revFormat?.author === sign.author ? t.revFormat : undefined;
  if (mine) return mine;
  const rec: TableFormatRevision = { kind: "format", tag: "tblPrChange", ...sign, prev: was, prevSlots: wasXml ? statedChildren(wasXml) : [] };
  t.revFormat = rec;
  if (t.tblPrXml) {
    const recorded = rec.prevSlots.map((s) => s.xml).join("");
    t.tblPrXml = replaceElementChild(t.tblPrXml, "tblPrChange", changeRecordXml("tblPrChange", sign, recorded, tagPrefix(t.tblPrXml)), TBLPR_CHILD_ORDER);
  }
  return rec;
}

/**
 * 开着修订改一格那一圈时留下的 `w:tcPrChange`，与表那一条同两分：文字挂进 `w:tcPr`（按 `CT_TcPr` 的孩子序，
 * 落在别人的 `w:cellIns`/`w:cellDel`/`w:cellMerge` 之后），字段留作「拒绝」的参照。
 * `CT_TcPr`（§17.4.73）里 `w:tcPrChange` 只容一枚，所以新签的这枚把上一枚顶掉——模型这一侧同步顶，否则
 * `revs` 里留一条文字里已经没有的记录，「拒绝」就会把一套没写过的属性摆回件里。
 * 同一个人记过就不再重签：一条 `w:id` 只裁一次，两击得一起退。
 */
export function signCellRevision(cell: TableCell, sign: { id: number; author: string; date?: string }, was: CellStyle, wasXml?: string): CellRevision {
  const mine = (cell.revs ?? []).find((r) => r.tag === "tcPrChange" && r.author === sign.author);
  if (mine) return mine;
  const rec: CellRevision = { tag: "tcPrChange", ...sign, prev: was, prevSlots: wasXml ? statedChildren(wasXml) : [] };
  if (cell.tcPrXml) {
    const recorded = rec.prevSlots.map((s) => s.xml).join("");
    cell.tcPrXml = replaceElementChild(cell.tcPrXml, "tcPrChange", changeRecordXml("tcPrChange", sign, recorded, tagPrefix(cell.tcPrXml)), TCPR_CHILD_ORDER);
  }
  cell.revs = [...(cell.revs ?? []).filter((r) => r.tag !== "tcPrChange"), rec];
  return rec;
}

/**
 * 开着修订改**行**那一圈（行高、重复表头、断页、隐藏）时留下的 `w:trPrChange`，与表、格那两条同两分：字节挂进
 * 行按枚带着的 `w:trPr` 孩子集（按 `CT_TrPr` 的孩子序，`TRPR_CHILD_ORDER`），字段留作「拒绝」的参照。
 * 记的是改之前那一整套**自己说过**的属性：`w:trPr` 的孩子里，别人那条待裁决的签名（`w:ins`/`w:del`）不算行
 * 说过的属性——读侧 `rowFormatRevision` 把它们从 `prevSlots` 里滤掉，这里同一口径，否则「拒绝」会把同一枚签名
 * 摆回两份。`CT_TrPr` 只容一枚 `w:trPrChange`，新签的这枚顶掉上一枚，模型这一侧同步顶。
 *
 * 同一个人记过就不再重签：一条 `w:id` 只裁一次，两击得一起退。没有孩子的行（编辑器自己造的）没有字节可挂，
 * 记录只进字段，由写出器按字段拼这一枚孩子——那一处也认这一枚记录，否则这一句「改之前不是这样」只活在内存里，
 * 存盘即蒸发。
 */
export function signRowRevision(
  row: RowInfo,
  sign: { id: number; author: string; date?: string },
  was: RowStyle,
  wasSlots: PropSlot[],
): RowFormatRevision {
  const mine = row.revFormat?.author === sign.author ? row.revFormat : undefined;
  if (mine) return mine;
  const recorded = wasSlots.filter((s) => !TABLE_RECORD_TAGS.has(s.tag) && !ROW_REVISION_TAGS.has(s.tag));
  const rec: RowFormatRevision = { kind: "format", tag: "trPrChange", ...sign, prev: was, prevSlots: recorded };
  if (row.slots?.length) {
    row.slots = upsertSlot(
      row.slots,
      "trPrChange",
      changeRecordXml("trPrChange", sign, recorded.map((s) => s.xml).join(""), slotPrefix(row.slots)),
      TRPR_CHILD_ORDER,
    );
  }
  row.revFormat = rec;
  return rec;
}

/**
 * 一格「在不在」的记录：`w:cellIns`（§17.4.66）与 `w:cellDel`（§17.4.67）只署名，不带改前的属性集，判定它们
 * 的那条路（`resolveTableParts`）读的就是这一枚孩子本身。与 `signCellRevision` 同一分工：原件写过 `w:tcPr` 的
 * 格子按 `CT_TcPr` 的孩子序把字节插进原件（别的孩子一个字节不动），没写过的格子留 `cell.revs` 给写出器去拼——
 * 那一处也认这两枚记录，否则这一句「不在了」就只活在内存里，存盘即蒸发。
 *
 * 同一格上同一句话只签一枚：已经挂着 `cellDel` 的格再点一次，不该多出一个 id 让审核者摇头两遍。
 */
export function signCellPlaceRevision(cell: TableCell, sign: { id: number; author: string; date?: string }, tag: "cellIns" | "cellDel"): CellRevision {
  const mine = (cell.revs ?? []).find((r) => r.tag === tag);
  if (mine) return mine;
  const rec: CellRevision = { tag, ...sign, prev: {}, prevSlots: [] };
  if (cell.tcPrXml) {
    cell.tcPrXml = replaceElementChild(cell.tcPrXml, tag, placeRecordXml(tag, sign, tagPrefix(cell.tcPrXml)), TCPR_CHILD_ORDER);
  }
  cell.revs = [...(cell.revs ?? []), rec];
  return rec;
}

/** One edge of `w:tblBorders` or `w:tcBorders`. */
function borderOf(node: any): TableBorder | undefined {
  const val = str(node?.["@_val"]);
  if (!val) return undefined;
  return { val, size: num(node?.["@_sz"]), space: num(node?.["@_space"]), color: str(node?.["@_color"]) };
}

/** A border set read from the parsed tree: `w:tcBorders` names one element per side it draws. */
function bordersOf(set: any): BorderSet | undefined {
  if (!set) return undefined;
  const out: BorderSet = {};
  let any = false;
  for (const side of BORDER_ORDER) {
    const b = borderOf(set[side]);
    if (b) {
      out[side] = b;
      any = true;
    }
  }
  return any ? out : undefined;
}

/** `w:pBdr` (§17.3.2.36): the six sides a paragraph may state, one element per side it says anything about. */
function paraBordersOf(set: any): ParaBorderSet | undefined {
  if (!set) return undefined;
  const out: ParaBorderSet = {};
  let any = false;
  for (const side of PBDR_ORDER) {
    const b = borderOf(set[side]);
    if (b) {
      out[side] = b;
      any = true;
    }
  }
  return any ? out : undefined;
}

/**
 * `w:shd@w:fill` (§17.3.2.33) as the element signed it: a colour, or the token the file used where a colour
 * would go — `w:fill="auto"` in 1,600 of the authorized 46 份's 1,676 `w:shd` elements, the other 76 naming a
 * hex. Keeping the token is what makes the ladder's
 * `??` mean 「这一段说不要地」 instead of 「这一段没说」; see `ParaStyle.shading` for the two audiences.
 * A value the model can't name travels verbatim and draws nothing, so the byte round-trip never invents one.
 * 授权口径实测（`.scratch/shd-audience.ts`，46 份件）：1,676 枚全是 `w:val="clear" w:color="auto"` 加一个
 * `w:fill`，其中 `w:tcPr` 1,646 枚（`auto` 1,600 / 十六进制 46）、`w:pPr` 30 枚（十六进制），`w:tblPr` 0 枚。
 */
function shadingOf(shd: any): string | undefined {
  const fill = str(shd?.["@_fill"]);
  return fill === undefined ? undefined : hex(fill) ?? fill;
}

/**
 * The paragraph's own sides over the ones its style chain names. Per side, not per element: §17.7.2 resolves one
 * property at a time, and a `w:pBdr` whose bottom says `none` under a style whose bottom says `single` is the
 * paragraph taking that one line back, not the paragraph refusing a box. The corpus says that sentence 18 times
 * among its 26 `w:pBdr` elements — 16 styles naming all four sides `none`, 2 header-part paragraphs naming a
 * `none` bottom (`.scratch/pbdr-audience.ts`) — but 0 body paragraphs write a refusal under a style that lends a
 * line, so the per-side merge is what this ladder needs for the click that has no measured instance yet, and the
 * measured cases are all one rung up. A paragraph naming no side of
 * its own keeps the chain's object **by reference**, which is how `nextSplit` tells a borrowed box from one the
 * paragraph wrote.
 */
function mergeParaBorders(own?: ParaBorderSet, lent?: ParaBorderSet): ParaBorderSet | undefined {
  if (!own) return lent;
  if (!lent) return own;
  return { ...lent, ...own };
}

/**
 * `w:pPr/w:tabs/w:tab` (§17.4.19) as one stop. `posTwips` is measured from the **left edge of the text area** —
 * the page's left margin plus `w:gutter`, which is what `render.ts` puts in `--pad-l` — not from the paragraph's
 * own indent. `leader` keeps the file's `w:leader` token; the page does not draw it yet.
 *
 * 受众（授权 46 份，`.scratch/tabs-audience.py`）：370 块 `w:tabs` 全在 `w:pPr` 底下——`word/document.xml` 230、
 * `word/styles.xml` 90、带子 12，其中 106 块一次写两枚以上；520 枚 `w:tab` 枚枚写 `w:pos`（0 枚漏），
 * 档只拼了四种——`left` 304、`right` 152、`center` 58、`clear` 6，`decimal`/`bar` 各 0 枚，所以模型不收它们；
 * 带 `w:leader` 的 94 枚。正文里真正走这些位的制表符（`w:r/w:tab`）116 枚、带子 2 枚。
 */
export type TabStop = { posTwips: number; val: "left" | "center" | "right" | "clear"; leader?: string };

/** The `w:tabs` block as written, `w:val="clear"` included. A block that names no stop says nothing. */
function tabStopsOf(tabs: any): TabStop[] | undefined {
  const out: TabStop[] = [];
  for (const t of asArray(tabs?.tab)) {
    const pos = Number(t?.["@_pos"]);
    // 没有 `w:pos` 的制表位落不到任何位置上；`w:val` 没说就按 ST_TabVal 的默认值 `left` 算。授权 46 份 520 枚
    // 枚枚两样都写，这两句只是不给坏件留半枚画歪的位置。
    if (!Number.isFinite(pos)) continue;
    const val = str(t?.["@_val"]);
    out.push({
      posTwips: pos,
      val: val === "center" || val === "right" || val === "clear" ? val : "left",
      leader: str(t?.["@_leader"]),
    });
  }
  return out.length ? out : undefined;
}

/**
 * The paragraph's stops over the chain's, **one position at a time** (§17.7.2 resolves a tab stop as its own
 * property): `upper` replaces the stop `lower` wrote at the same `w:pos` and leaves the rest of that list alone,
 * while `w:val="clear"` spends itself on the position it names and keeps no place of its own. Nearest-wins over
 * the whole element — what `paraLent` does for every other key — would let one stop in the paragraph erase a
 * style's whole set, and 106 of the 370 blocks in the authorized 46 份 write two or more stops at once.
 */
function mergeTabStops(lower?: readonly TabStop[], upper?: readonly TabStop[]): TabStop[] | undefined {
  if (!upper?.length) return lower ? [...lower] : undefined;
  const byPos = new Map<number, TabStop>();
  for (const s of lower ?? []) byPos.set(s.posTwips, s);
  for (const s of upper) {
    if (s.val === "clear") byPos.delete(s.posTwips);
    else byPos.set(s.posTwips, s);
  }
  const out = [...byPos.values()].sort((a, b) => a.posTwips - b.posTwips);
  return out.length ? out : undefined;
}

/**
 * A cell-margin set read from the parsed tree: `w:tblCellMar` and `w:tcMar` each name one element per side
 * they pad. A side counts only when the file gave it a `w:w`, because `w:w="0"` (flush) is a fact the page
 * has to draw, not the absence of one. `w:type` says what the number is counted in; a side spelled anything
 * other than `dxa` is left unsaid — `auto` asks for the next layer's answer, which is exactly what leaving
 * it out does here.
 */
function marginOf(set: any): MarginTwips | undefined {
  if (!set) return undefined;
  const out: MarginTwips = {};
  let any = false;
  for (const side of MARGIN_SIDES) {
    const node = set[side];
    if (!node) continue;
    const type = str(node["@_type"]);
    if (type && type !== "dxa") continue;
    const w = num(node["@_w"]);
    if (w === undefined) continue;
    out[side] = w;
    any = true;
  }
  return any ? out : undefined;
}

/** What a row says about itself: its own revision, its height, whether it repeats or may split. */
function rowInfoOf(
  trPr: any,
  slots?: PropSlot[],
  attrs?: string,
  extras?: RowExtra[],
  cellMargin?: MarginTwips,
): RowInfo | undefined {
  const info: RowInfo = { ...rowStyleOf(trPr) };
  const rev = trPr?.ins ? revisionOf("ins", trPr.ins) : trPr?.del ? revisionOf("del", trPr.del) : undefined;
  if (rev) info.rev = rev;
  if (cellMargin) info.cellMargin = cellMargin;
  if (slots?.length) info.slots = slots;
  if (extras?.length) info.extras = extras;
  if (attrs) info.attrsXml = attrs;
  return Object.keys(info).length ? info : undefined;
}

/**
 * What a `w:trPr` says about its row's shape — height, header repeat, splitting, hiding. Read this way for
 * the row as it stands and for the set its `w:trPrChange` recorded, so a refusal cannot put back properties
 * the reader would spell differently from the ones it took off.
 */
function rowStyleOf(trPr: any): RowStyle {
  const out: RowStyle = {};
  const height = num(trPr?.trHeight?.["@_val"]);
  if (height !== undefined) out.heightTwips = height;
  const rule = str(trPr?.trHeight?.["@_hRule"]);
  if (rule) out.heightRule = rule;
  if (onOff(trPr?.tblHeader)) out.header = true;
  if (onOff(trPr?.cantSplit)) out.keepTogether = true;
  if (trPr?.hidden !== undefined) out.hidden = onOff(trPr.hidden) ?? true;
  const jc = str(trPr?.jc?.["@_val"]);
  if (jc) out.jc = jc;
  return out;
}

/**
 * What a `w:tblPr` says about its table — style, width, place, layout, borders. The recorded set of a
 * `w:tblPrChange` is read by the same function for the same reason.
 */
function tableStyleOf(tblPr: any): TableStyle {
  return {
    styleId: str(tblPr?.tblStyle?.["@_val"]),
    widthTwips: num(tblPr?.tblW?.["@_w"]),
    widthType: str(tblPr?.tblW?.["@_type"]),
    indentTwips: num(tblPr?.tblInd?.["@_w"]),
    align: str(tblPr?.jc?.["@_val"]),
    layout: str(tblPr?.tblLayout?.["@_type"]),
    borders: bordersOf(tblPr?.tblBorders),
    cellMargin: marginOf(tblPr?.tblCellMar),
    shading: shadingOf(tblPr?.shd),
  };
}

/** The `w:id`, `w:author` and `w:date` every record on a table part carries. */
function recordAttrs(rec: any): { id: number; author: string; date?: string } {
  return { id: Number(rec?.["@_id"]) || 0, author: String(rec?.["@_author"] ?? ""), date: str(rec?.["@_date"]) };
}

/** The set a record names, which the schema wraps in an element of the set's own name. */
const recordedNode = (rec: any, name: string): any => rec?.[name] ?? rec ?? {};

/** The source text of the record carrying this tag inside a property set the model only carries as text. */
function recordSpan(setXml: string | undefined, tag: string): string | undefined {
  return setXml ? childSpans(elementContent(setXml)).find((s) => s.tag === tag)?.xml : undefined;
}

/** A table's pending `w:tblPrChange`, read from the set the file wrote and the node the parser made. */
function tableFormatRevision(tblPr: any, tblPrXml: string | undefined): TableFormatRevision | undefined {
  const rec = tblPr?.tblPrChange;
  if (!rec) return undefined;
  const span = recordSpan(tblPrXml, "tblPrChange");
  return {
    kind: "format",
    tag: "tblPrChange",
    ...recordAttrs(rec),
    prev: tableStyleOf(recordedNode(rec, "tblPr")),
    prevSlots: span ? recordedSlots(span, "tblPrChange") : [],
  };
}

/**
 * A row's pending `w:trPrChange`, whose recorded set is the row's own children in the slots they came in.
 * A signature inside the recorded set is left out: `w:trPr`'s `w:ins` and `w:del` are carried as the row's
 * own revision, and putting a second copy of one back is how a refusal would end up with two signatures.
 */
function rowFormatRevision(trPr: any, slots: PropSlot[] | undefined): RowFormatRevision | undefined {
  const rec = trPr?.trPrChange;
  if (!rec) return undefined;
  const span = slots?.find((s) => s.tag === "trPrChange")?.xml;
  return {
    kind: "format",
    tag: "trPrChange",
    ...recordAttrs(rec),
    prev: rowStyleOf(recordedNode(rec, "trPr")),
    prevSlots: span ? recordedSlots(span, "trPrChange").filter((s) => !ROW_REVISION_TAGS.has(s.tag)) : [],
  };
}

/** The records a cell's `w:tcPr` can hold, in the order the schema writes them. */
const CELL_REVISION_TAGS: readonly CellRevision["tag"][] = ["cellIns", "cellDel", "cellMerge", "tcPrChange"];

/** What each of those shows as in the list of changes waiting, which is the reviewer's word for it. */
const CELL_REVISION_TYPES: Record<CellRevision["tag"], RevisionEntry["type"]> = {
  cellIns: "cell-insert",
  cellDel: "cell-delete",
  cellMerge: "cell-merge",
  tcPrChange: "cell-format",
};

/**
 * Every revision record one cell carries. The tags come off the parsed tree because a cell can hold several
 * at once — a column inserted while 修订 was on, whose cell was then rewidened, is two changes with two
 * `w:id`s — and the properties come off the text the file wrote, because that is what a decision rewrites.
 * `w:cellIns` and `w:cellDel` sign the cell's place rather than its shape, so they record no properties.
 */
function cellRevisions(tcPr: any, tcPrXml: string | undefined): CellRevision[] {
  const out: CellRevision[] = [];
  for (const tag of CELL_REVISION_TAGS) {
    const rec = tcPr?.[tag];
    if (!rec) continue;
    const span = recordSpan(tcPrXml, tag);
    out.push({
      tag,
      ...recordAttrs(rec),
      prevSlots: span ? recordedSlots(span, tag) : [],
      prev: tag === "cellIns" || tag === "cellDel" ? {} : cellStyleOf(recordedNode(rec, "tcPr")),
    });
  }
  return out;
}

/** The cell's own formatting, from `w:tcPr`: width, merge, shading, borders, alignment, wrapping. */
function cellStyleOf(tcPr: any): CellStyle {
  const out: CellStyle = {
    widthTwips: num(tcPr?.tcW?.["@_w"]),
    widthType: str(tcPr?.tcW?.["@_type"]),
    shading: shadingOf(tcPr?.shd),
    spanCols: num(tcPr?.gridSpan?.["@_val"]),
    spanRows: tcPr?.vMerge === undefined ? undefined : tcPr.vMerge["@_val"] === "restart" ? 1 : 0,
    valign: str(tcPr?.vAlign?.["@_val"]),
    textDirection: str(tcPr?.textDirection?.["@_val"]),
    borders: bordersOf(tcPr?.tcBorders),
    margin: marginOf(tcPr?.tcMar),
  };
  const wrap = onOff(tcPr?.noWrap);
  if (wrap !== undefined) out.noWrap = wrap;
  return out;
}

/** Every table property a `w:tblPrChange` can carry, so a refusal clears the ones the record left out. */
export const TABLE_STYLE_KEYS: (keyof TableStyle)[] = [
  "styleId",
  "widthTwips",
  "widthType",
  "indentTwips",
  "align",
  "layout",
  "borders",
  "cellMargin",
  "shading",
];

/** Every row property a `w:trPrChange` can carry. */
export const ROW_STYLE_KEYS: (keyof RowStyle)[] = ["heightTwips", "heightRule", "header", "keepTogether", "hidden", "jc"];

/** Every cell property a `w:tcPrChange` or a `w:cellMerge` can carry. */
export const CELL_STYLE_KEYS: (keyof CellStyle)[] = [
  "widthTwips",
  "widthType",
  "shading",
  "spanCols",
  "spanRows",
  "valign",
  "noWrap",
  "textDirection",
  "borders",
  "margin",
];

/**
 * Putting a table back the properties its `w:tblPrChange` recorded. The record holds the whole set, so a
 * property it does not name was not set — which is why these clear before they assign.
 */
function restoreTableStyle(t: Table, prev: TableStyle): void {
  for (const key of TABLE_STYLE_KEYS) delete t[key];
  Object.assign(t, prev);
}

/** Putting a row back the properties its `w:trPrChange` recorded. */
function restoreRowStyle(row: RowInfo, prev: RowStyle): void {
  for (const key of ROW_STYLE_KEYS) delete row[key];
  Object.assign(row, prev);
}

/**
 * 同一格挂着多条属性记录时，`prev` 按**文档顺序先说的那一句**算。每条记录说的是「这次改之前」的样子，
 * 后一条记的是前一条改完之后的状态；按文档顺序逐条盖上去，后一条就会把最早那一句顶掉。
 * 例：`w:cellMerge` 的 prev 说这是 `continue`，随后 `w:tcPrChange` 的 prev 说 `restart`——逐条盖出来仍是
 * `restart`，原文与最终态画成同一套范围（#176）。整套记录（`w:tcPrChange`）只认领还没人答过的键，
 * 这样它清掉的是自己那一刻「本来没有」的属性，不会抹掉更早那条记录答过的属性。
 */
function restoreCellRecords(cell: TableCell, records: CellRevision[]): void {
  const answered = new Set<keyof CellStyle>();
  for (const rec of records) {
    const whole = coversWholeSet(rec.tag);
    // `prev` 常常把每种属性都列成 `undefined`；单属性记录只认值非空的那几枚，别把没说的键也认领掉。
    const keys = whole
      ? CELL_STYLE_KEYS
      : (Object.keys(rec.prev) as (keyof CellStyle)[]).filter((key) => rec.prev[key] !== undefined);
    for (const key of keys) {
      if (answered.has(key)) continue;
      answered.add(key);
      const value = rec.prev[key];
      if (value !== undefined) (cell as unknown as Record<string, unknown>)[key] = value;
      else if (whole) delete (cell as unknown as Record<string, unknown>)[key];
    }
  }
}

/**
 * Whether this record sits on every cell of the table that reaches grid column `at`. One `w:id` shared down a
 * column is how OOXML signs a revised *column*; a record on fewer cells than the column has names those cells
 * and nothing else, and dropping the column for it would take cells nobody revised. A row that does not reach
 * the column (a shorter row, or a merge that spans past it) says nothing either way.
 */
function coversColumn(t: Table, at: number, rec: CellRevision): boolean {
  return t.rows.every((row) => {
    const i = cellCovering(row, at);
    return i < 0 || row[i].revs?.some((r) => r.id === rec.id && r.tag === rec.tag) === true;
  });
}

/**
 * One `w:rPr` read into fields. `theme` resolves the four `w:rFonts` theme attributes — a run that names no face of
 * its own but points at `minorHAnsi` did name one, two parts away — and is left off only by callers whose `w:rPr`
 * comes from somewhere with no theme to ask.
 */
function runStyle(rPr: any, theme?: ThemeFonts): RunStyle {
  const s: RunStyle = {};
  if (rPr.b !== undefined) s.bold = rPr.b["@_val"] !== "0" && rPr.b["@_val"] !== "false";
  if (rPr.i !== undefined) s.italic = rPr.i["@_val"] !== "0" && rPr.i["@_val"] !== "false";
  // `w:u` belongs to this same family and spells its cancellation with an enum name instead of `0`: §17.3.2.25's
  // `ST_Underline` says `none`. Reading that as silence left the property the odd one out — a run that said
  // "no underline" under a style that lends one kept the style's line, because §17.7.2 stops per property at the
  // first rung that *stated* one, and this rung stated it. 授权 46 份把这句话写得很勤：**`w:u` 1,615 枚 = `single`
  // 1,565 + `none` 30 + `thick` 20，0 枚裸写**（ElementTree 逐枚；1,593 枚在 `word/document.xml`、22 枚在
  // `word/styles.xml`）。作废语料上那一句「952 枚 none、其中 254 枚坐在本 reader 解析成字段的 run 上」是
  // 待重证的旧读数，现在换成上面这组。那 30 枚一枚都不落在「链上有人借给它一条线」的位置上：样式层那 22 枚
  // `w:u` 全部是**字符样式**写的 `single`，只有 run 自己点名（`w:rStyle`）才借得到，而这 30 枚里没有一枚带
  // `w:rStyle`（28 枚两个样式指针都没写，2 枚只写 `w:pStyle`），段落样式层 0 枚写 `w:u`。所以这一条同样是
  // representation fix，不是 pixel fix——它买到的是「一句文件写过的话，不该被一次不相干的编辑改掉」。
  // 形状名从这一轮起留在模型里（`thick` 那 20 枚 + 段落标记上 16 枚是这一档的受众）：从前 `true` 把「有线」和
  // 「什么线」压成一个答案，纸上只能画细线，而 `w:u` 没被碰过的那一路靠原字节带回，画与存说的是两句话。
  if (rPr.u !== undefined) s.underline = rPr.u["@_val"] || "single";
  // The last two members of the same on/off family. `w:strike` is §17.3.2.19's `ST_OnOff`, so `0` and `false`
  // cancel it just as they cancel `w:b` — reading the element as "struck, whatever it says" put a line through the
  // 2 runs stating `<w:strike w:val="0"/>` among the 56 `w:strike` of the quarantined 86's body parts (36 of those
  // say on with no `w:val`, 18 sit on a paragraph mark) and gave the reader text to hang it on
  // (`.scratch/onoff-where.ts`, `.scratch/onoff-audience.ts`); `w:strike` has no authorized reading yet — 待重证.
  // `w:vertAlign`'s cancellation is an enum name again (§17.3.2.24's `ST_VerAlignRun`: `baseline`), and those
  // bodies state nothing else — all 3860 of them, 717 on runs this reader resolves, none sitting on a chain that
  // lends a superscript — likewise 待重证 on the authorized 46 份, so that half is representation rather
  // than pixels, the same verdict `w:u` got above.
  if (rPr.strike !== undefined) s.strike = rPr.strike["@_val"] !== "0" && rPr.strike["@_val"] !== "false";
  if (rPr.sz) s.sizePt = Number(rPr.sz["@_val"]) / 2;
  const color = rPr.color?.["@_val"];
  if (color && color !== "auto") s.color = hex(color);
  // `w:spacing` 的 `w:val` 是 `ST_SignedTwipsMeasure`，所以负数是合法值、`0` 是「说了不加不减」。只有数不出来
  // 才算没说——把 `-6` 读成沉默，就是替文件撤了它自己撤掉的那 0.3 磅。
  if (rPr.spacing !== undefined) {
    const stated = Number(rPr.spacing["@_val"]);
    if (Number.isFinite(stated)) s.charSpacingTwips = stated;
  }
  // `none` is kept as the value it is, not folded into the absence above it: §17.17.64 spells its cancellation with
  // that enum name, and the quarantined 86 use it for 4,135 of the 4,605 `w:highlight` elements they state — the
  // other 470 all saying `yellow` and none leaving `w:val` off (`.scratch/highlight-counted.ts`). `w:highlight`
  // has not been re-counted on the authorized 46 份, so that split is 待重证; see `RunStyle.highlight`.
  if (rPr.highlight?.["@_val"]) s.highlight = rPr.highlight["@_val"];
  if (rPr.rFonts) {
    // One element, several faces, and a document may state any subset of them. Each attribute is read on its
    // own: `w:ascii="宋体"` alone must not hand back an `undefined` that erases the chain's Chinese face, since
    // §17.7.2 stops per property at the first level that stated one.
    const ascii = rPr.rFonts["@_ascii"];
    if (ascii) s.font = ascii;
    const east = rPr.rFonts["@_eastAsia"];
    if (east) s.fontEastAsia = east;
    // A face can also be stated by pointing at the theme instead of naming one, and these files do it for
    // 250 of styles.xml's references and 58 documents' whole run floor (`.scratch/theme-font-audience.ts`).
    // Only where no literal face was written: §17.3.2 gives no precedence between the two, and per slot the files
    // never need one on the authorized set: its 46 份 state `w:asciiTheme` and `w:hAnsiTheme` 1,796 times each,
    // `w:eastAsiaTheme` 1,143 and `w:csTheme` 1,155, while 11,324 slots name a literal `w:ascii`/`w:eastAsia` face
    // (`.scratch/corpus-recheck2.mjs`). Whether the two spellings ever share **one** slot is measured on that same
    // authorized set: of its 14,642 `w:rFonts` tags, 1,957 point a slot at the theme, 11,905 name a literal face and
    // 650 do both in different slots — while a slot that writes a literal name *and* a theme reference is **0 枚 / 0 件**
    // (`.scratch/theme-slot-pair-authorized.mjs`). So each slot is asked on its
    // own and a slot that named a face is not asked.
    if (theme) {
      const script = themeFontScript(asArray(rPr.lang)[0]?.["@_eastAsia"]);
      if (!s.font) {
        const f = theme.face(rPr.rFonts["@_asciiTheme"] ?? rPr.rFonts["@_hAnsiTheme"], script);
        if (f) s.font = f;
      }
      if (!s.fontEastAsia) {
        const f = theme.face(rPr.rFonts["@_eastAsiaTheme"], script);
        if (f) s.fontEastAsia = f;
      }
    }
  }
  const va = rPr.vertAlign?.["@_val"];
  if (va === "superscript" || va === "subscript" || va === "baseline") s.vertAlign = va;
  return s;
}

/** §17.3.2.26 按书写系统分派时算东亚的那一类码位，与 `render.ts` 的 `unicode-range` 同一套码段。 */
export const EAST_ASIAN_TEXT = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;

/**
 * 一次「挑脸」管不管得到汉字槽。ribbon 的字体框说的是一个 run 的脸，而 §17.3.2.26 把一个 run 的字按书写系统
 * 分给 `w:ascii`/`w:hAnsi`/`w:eastAsia` 三槽，所以这一句落到哪几槽得先看件自己怎么说。授权口径 46 份把正文、页眉
 * 页脚、脚注尾注那几件部件合起来数（`.scratch/corpus-recheck2.mjs`）：1,800 枚 `w:rFonts` 指主题、11,324 枚至少写
 * 了 `w:ascii` 或 `w:eastAsia` 的一张实名，两槽都写而两张名**不同**的 161 枚——`corpus-recheck.mjs` 独立数出同一枚
 * 161（并给出两槽都写的 4,475 枚、分布在 15 件里），两份脚本的 161 相符，故引用「4,475 枚中 161 枚不同」这一对；逐槽看（`.scratch/theme-slot-pair-authorized.mjs`，
 * 各部件 14,642 枚标签）同一槽既写实名又指主题的 **0** 枚 / **0** 件。作废口径那组按正文文字分类的数（汉字正文
 * 两槽同值 7,706、不同值 47；纯拉丁两槽全同值 1,062；只写拉丁槽的汉字正文 2,331，`.scratch/font-slot-shape.ts`）
 * 没有按这个口径重测，是待重证。于是：件里已经给这个 run 留了汉字槽（写名或指主题都算留），除非它明说两张不同
 * 的脸，挑的那张就接着管这一槽；件里没留，只有这段文字自己带汉字时才替它补上——不给一段纯拉丁的 run 写一句
 * 关于中文的话。
 * `w:cs`/`w:cstheme` 模型里没有字段，照旧按原件搬。真机 WPS 的字体框改一次究竟写哪几槽仍未取证。
 */
export function pickedFaceCoversHan(slots: PropSlot[] | undefined, text: string): boolean {
  const tag = slots?.find((s) => s.tag === "rFonts")?.xml ?? "";
  const attr = (name: string): string => new RegExp(`[\\w.-]+:${name}="([^"]*)"`).exec(tag)?.[1]?.trim() ?? "";
  const latin = attr("ascii") || attr("hAnsi");
  const east = attr("eastAsia");
  if (east || attr("eastAsiaTheme")) return !(east && latin && east !== latin);
  return EAST_ASIAN_TEXT.test(text);
}

/**
 * 一次「挑脸」在一个 run 上摊开到哪几槽——这条规矩只有这一个主人：`patched`（选区那一下）与 `armPending`
 * （悬光标处攒着的那一下）都调它，测试也调它。主人放在 core 而不是 UI，是因为 `src/ui` 那一层要能被这份件
 * 自己说过的话决定，不该再抄一遍摊开的写法。
 */
export function applyFacePick(run: Pick<Inline, "text" | "rPrSlots" | "runText">, face: string): Pick<RunStyle, "font" | "fontEastAsia"> {
  // 问的是整枚 run：件里那一套 `w:rPr` 是 run 共用的，`w:tab` 那半句自己不带字，但同一枚 run 里写的是中文。
  return pickedFaceCoversHan(run.rPrSlots, run.runText ?? run.text) ? { font: face, fontEastAsia: face } : { font: face };
}

/**
 * 一条格式命令落在一个 run 上的样子。挑脸那一句先问 `applyFacePick` 该管哪几槽——不该管的槽一个字都不动，写端才
 * 认得出它没挪；空串是「这句撤回」，字段删掉，落回原件那半边的字节。
 * 主人放在 core 而不是 UI：`src/ui/doc-edit.ts` 的 `restyle` 与 `armPending` 都只是它的调用方，UI 里再抄一份摊开
 * 的写法就是第二个主人（抄的那份永远不会漏，本文件第一轮即如此）。
 */
export type RunPatch = { [K in keyof RunStyle]?: RunStyle[K] | "" };

export function applyRunPatch(inline: Inline, patch: RunPatch): Inline {
  const out: Inline = { ...inline };
  const entries = Object.entries(patch) as [keyof RunStyle, RunStyle[keyof RunStyle]][];
  if (patch.font) entries.push(...(Object.entries(applyFacePick(inline, patch.font)) as [keyof RunStyle, RunStyle[keyof RunStyle]][]));
  for (const [key, value] of entries) {
    if (value === undefined) continue;
    if (typeof value === "string" && value === "") delete out[key];
    else (out as unknown as Record<string, unknown>)[key] = value;
  }
  return out;
}

/**
 * 「清除格式」落在一枚 run 上的样子：把模型自己有字段的那一整组直接格式（`RUN_STYLE_KEYS` 那 11 项）删掉，
 * 这段字回到段落样式与样式阶梯给它的那一套；段落格式（对齐/缩进/行距）在段上，不在这里动。
 *
 * 模型没有字段的槽一枚也不碰（`w:rStyle`、`w:lang`、`w:kern`、`w:bCs`/`w:iCs`、`w:w`、`w:fitText`…）——
 * 那些话不是这个模型说的，替它删就是替件作一句没人说过的决定。授权 46 件正文与带子里 7,300 枚 run，
 * 6,972 枚带模型字段，**只有 34 枚**（`w:rStyle` 26、`w:lang` 8）点了这一枚会一个字也不动
 * （`.scratch/clear-format-audience.py`）。WPS 真机这一枚清到哪一层（字符样式 `w:rStyle` 清不清、
 * 与 `w:b` 并排的 `w:bCs` 清不清）仍未取证。
 */
export function clearRunStyle(inline: Inline): Inline {
  const out: Inline = { ...inline };
  for (const key of RUN_STYLE_KEYS) delete out[key];
  return out;
}

/**
 * A link with no colour of its own is drawn underlined. That is a look this reader puts on, not a `w:u` the
 * file wrote — the `w:Hyperlink` character style is what underlines it — so the read snapshot is told as well,
 * or a save would spell into the run an underline nobody asked for.
 */
function linkLook(inline: Inline): void {
  if (inline.color) return;
  inline.underline = "single";
  if (inline.asRead) inline.asRead = { ...inline.asRead, underline: "single" };
}

/**
 * 节的边界，从模型这一侧数出来：一个段落带上 `w:pPr/w:sectPr`，它就是它所**结束**的那一节（`sectionChain`
 * 干的是同一件事，只不过那边读的是字节）。返回这些段落在 `blocks` 里的下标，按文档顺序。
 *
 * 由此两条：块 `b` 属于第 `1 + #{e < b}` 节——带断点的段落自己属于它结束的那一节；第 N 节（N>1）的第一块
 * 是 `ends[N−2] + 1`。正文最后那个 `w:sectPr` 不进 `blocks`（`blockChildren` 把它挡掉），它永远是最后一节
 * 的收尾，所以节数 = `ends.length + 1`。
 */
export function sectionEndsOf(blocks: readonly Block[]): number[] {
  const ends: number[] = [];
  blocks.forEach((b, i) => {
    if (b.kind === "paragraph" && b.pPrSlots?.some((s) => s.tag === "sectPr")) ends.push(i);
  });
  return ends;
}

/**
 * A section's page box, read off its `sectPr` bytes. Word's own defaults for anything the section does not
 * say: A4 (11906×16838 twips) and a 1-inch margin all round.
 */
function sectionGeometry(xml: string): Section {
  const kids = childSpans(elementContent(xml));
  const attrs = (name: string): Record<string, string> | undefined => {
    const s = kids.find((k) => k.tag === name);
    return s ? attrsOf(s.xml) : undefined;
  };
  const num = (a: Record<string, string> | undefined, k: string, d: number): number => {
    const v = a?.[k];
    return v === undefined || v === "" || Number.isNaN(Number(v)) ? d : Number(v);
  };
  const stated = (a: Record<string, string> | undefined, k: string): number | undefined => {
    const v = a?.[k];
    return v === undefined || v === "" || Number.isNaN(Number(v)) ? undefined : Number(v);
  };
  const mar = attrs("pgMar");
  const grid = attrs("docGrid");
  const borderEl = kids.find((k) => k.tag === "pgBorders");
  const sides: Partial<Record<PageBorderSide, PageBorder>> = {};
  if (borderEl) {
    for (const k of childSpans(elementContent(borderEl.xml))) {
      const side = k.tag as PageBorderSide;
      if (!PAGE_BORDER_SIDES.includes(side)) continue;
      const a = attrsOf(k.xml);
      // A side with no `w:val` states no line — the shape is the one attribute that cannot be guessed at.
      if (!a["val"]) continue;
      const line: PageBorder = { val: a["val"] };
      const sz = stated(a, "sz");
      if (sz !== undefined) line.size = sz;
      const space = stated(a, "space");
      if (space !== undefined) line.space = space;
      if (a["color"]) line.color = a["color"];
      if (a["frame"] !== undefined) line.frame = a["frame"] === "true" || a["frame"] === "1";
      sides[side] = line;
    }
  }
  return {
    pageWidthTwips: num(attrs("pgSz"), "w", DEFAULT_SECTION.pageWidthTwips),
    pageHeightTwips: num(attrs("pgSz"), "h", DEFAULT_SECTION.pageHeightTwips),
    margins: {
      top: num(mar, "top", DEFAULT_SECTION.margins.top),
      right: num(mar, "right", DEFAULT_SECTION.margins.right),
      bottom: num(mar, "bottom", DEFAULT_SECTION.margins.bottom),
      left: num(mar, "left", DEFAULT_SECTION.margins.left),
      header: stated(mar, "header"),
      footer: stated(mar, "footer"),
      gutter: stated(mar, "gutter"),
    },
    pageBorders: borderEl ? { offsetFrom: attrsOf(borderEl.xml)["offsetFrom"] === "text" ? "text" : "page", sides } : undefined,
    // `w:docGrid`: the element and its `w:type` are read together, because a section that writes the element
    // without a type asks for §17.3.2.9's default (`default`), which is no line grid — the same way an omitted
    // `w:offsetFrom` above lands on the default that element gives it.
    docGrid: grid ? { type: grid["type"] || "default", linePitch: stated(grid, "linePitch") } : undefined,
  };
}

function unwrap(parsed: any, rootName: string): any {
  return parsed?.[rootName];
}

/**
 * The body's own direct children in source order, each with the bytes it was written as. Both the table
 * carry and the containers kept as they came read from here, because the position a child held in the part
 * is the only record of where a save has to put it back.
 */
function bodyChildSpans(source: string): { tag: string; xml: string }[] {
  const root = childSpans(source).find((s) => s.tag === "document");
  if (!root) return [];
  const body = childSpans(elementContent(root.xml)).find((s) => s.tag === "body");
  return body ? childSpans(elementContent(body.xml)) : [];
}

/** Whether the parsed tree and the source spans list the same children, tag for tag. */
function sameTags(parsed: { name: string }[], spans: { tag: string }[]): boolean {
  const counts = new Map<string, number>();
  for (const child of parsed) counts.set(child.name, (counts.get(child.name) ?? 0) + 1);
  for (const span of spans) {
    const left = (counts.get(span.tag) ?? 0) - 1;
    if (left < 0) return false;
    counts.set(span.tag, left);
  }
  return [...counts.values()].every((n) => n === 0);
}

/** A body or cell child the model has no fields for, kept as the text it came in with. */
function carriedOf(span: { tag: string; xml: string }): Carried {
  return { kind: "carried", tag: span.tag, xml: span.xml, text: carriedText(span.xml) };
}

/**
 * What field machinery a run is made of — a `w:fldChar` boundary or the `w:instrText` instruction, with
 * nothing beside them but their own properties — or an empty string for a run this model can read. A run of
 * machinery has no word in it, and dropping it turns a live field into the number it happened to show the
 * last time somebody updated it.
 */
function machineryOf(r: any): string {
  let name = "";
  for (const kid of children(r)) {
    if (kid.name === "fldChar" || kid.name === "instrText") name ||= kid.name;
    else if (kid.name !== "rPr") return "";
  }
  return name;
}

/**
 * The text a reader sees inside an element: each paragraph's runs, joined, with a space where the part put
 * a paragraph break. `w:instrText` is deliberately not collected — a field's instruction is not prose the
 * page should show as if it were the document's own words.
 */
function carriedText(xml: string): string {
  const runsOf = (chunk: string): string =>
    [...chunk.matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([^<]*)<\/(?:[\w.-]+:)?t>/g)].map((m) => decodeEntities(m[1] ?? "")).join("");
  return xml
    .split(/<(?:[\w.-]+:)?p[\s/>]/)
    .map(runsOf)
    .filter((s) => s !== "")
    .join(" ");
}

/**
 * What the document root's start tag declared: its attribute text exactly as the file wrote it, every
 * namespace prefix it bound, and the prefixes it told a reader to skip if it does not know them.
 *
 * The attribute text is what a save puts back — carried markup names prefixes by the names the file chose,
 * and the declaration list and the ignorable list are only part of it. Reading the other two as well is what
 * lets a writer that has to spell a fresh root bind at least the prefixes its own tags use.
 */
function rootMarkup(source: string): { prefixes: Record<string, string>; ignorable: string; attrs: string } {
  const open = /<(?:[\w-]+:)?document\b[^>]*>/.exec(source)?.[0] ?? "";
  const prefixes: Record<string, string> = {};
  for (const m of open.matchAll(/xmlns:([\w.-]+)\s*=\s*"([^"]*)"/g)) prefixes[m[1]!] = m[2]!;
  return {
    prefixes,
    ignorable: /\bmc:Ignorable\s*=\s*"([^"]*)"/.exec(open)?.[1] ?? "",
    attrs: open.replace(/^<(?:[\w-]+:)?document\b/, "").replace(/\/?>$/, ""),
  };
}

/** Levels are addressed by their w:ilvl, which need not be contiguous or in document order. */
function orderLevels(levels: NumberLevel[]): NumberLevel[] {
  const out: NumberLevel[] = [];
  levels.forEach((l, i) => {
    out[l.ilvl ?? i] = l;
  });
  return out;
}

function num(v: unknown): number | undefined {
  if (v === undefined || v === null || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

function hex(v: unknown): string | undefined {
  if (typeof v !== "string") return undefined;
  if (/^[0-9A-Fa-f]{6}$/.test(v)) return "#" + v;
  if (/^[0-9A-Fa-f]{8}$/.test(v)) return "#" + v.slice(2);
  return undefined;
}

/** 「字数统计」报的那五个数：整篇与选区共用这一个形状。 */
export type WordCounts = {
  paragraphs: number;
  characters: number;
  charactersNoSpace: number;
  cjk: number;
  words: number;
};

/**
 * 给定一段段文字，数出那五个数。口径（码点、`White_Space`、汉字／假名／谚文字位、`字数 = 中文字符 + 非中文单词`）
 * 逐条解释在 `WordDocument.wordCounts()` 上面——那里说的是整篇，这里是**同一份口径换个取材**：
 * 一个元素就是一段（空段落也占一个元素，所以段落数是 `texts.length`），选区传进来的是它自己那几段被切出来的片段。
 */
export function countTexts(texts: readonly string[]): WordCounts {
  // 汉字（含扩充 A 与兼容区）、假名、谚文（含字母区与兼容 Jamo）。增补平面 B–F 用 \u{…} 写。
  const CJK = /[\u3040-\u30ff\u3130-\u318f\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\u1100-\u11ff\u{20000}-\u{2fa1f}]/gu;
  let characters = 0;
  let charactersNoSpace = 0;
  let cjk = 0;
  let words = 0;
  for (const text of texts) {
    for (const ch of text) {
      characters++;
      if (!/\s/u.test(ch)) charactersNoSpace++;
    }
    cjk += text.match(CJK)?.length ?? 0;
    for (const token of text.replace(CJK, " ").split(/\s+/)) if (/[\p{L}\p{N}]/u.test(token)) words++;
  }
  return { paragraphs: texts.length, characters, charactersNoSpace, cjk, words };
}

/** 这些块里那一段段的文字，表格顺着 `w:tc` 递归下去；整篇传 `this.blocks`，选区中间那几块也可以单独传。 */
export function paragraphTextsOf(blocks: readonly Block[]): string[] {
  const out: string[] = [];
  const walk = (bs: readonly Block[]): void => {
    for (const b of bs) {
      if (b.kind === "table") {
        for (const row of b.rows) for (const cell of row) walk(cell.paragraphs);
        continue;
      }
      if (b.kind !== "paragraph") continue;
      out.push(b.inlines.map((i) => i.text ?? "").join(""));
    }
  };
  walk(blocks);
  return out;
}
