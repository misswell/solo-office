import { OpcPackage, asArray, childSpans, children, decodeEntities, elementContent } from "./opc.js";
import { isMark } from "./run-identity.js";
import { dataUrl } from "./media.js";
import { coversWholeSet, decideSetXml, decideSlots, recordedSlots, setChildren } from "./doc-table-revisions.js";
import { cellCovering, columnStarts, deleteColumn, gridColumns, removeCell, repairMerges } from "./doc-table-ops.js";
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
  underline?: boolean;
  strike?: boolean;
  sizePt?: number;
  color?: string;
  /**
   * §17.17.64's `ST_HighlightColor` name, held as the three states the element has: absent is "this run said
   * nothing", `"none"` is "this run said no highlight", and any other member is the colour it chose. `none` is a
   * value rather than an absence because the file wrote it as one — the 86 WPS documents state
   * `<w:highlight w:val="none"/>` 4,135 times (`.scratch/highlight-counted.ts`) — and §17.7.2 stops a property at
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
   * element, not a fallback for the first: 45,054 of the corpus's 61,702 modelled runs write both attributes
   * with the same name, and 1,250 write two different ones (`.scratch/run-font-eastasia.ts`). A run whose
   * Chinese draws in the Latin face is a run the file never described that way.
   */
  fontEastAsia?: string;
  vertAlign?: "baseline" | "superscript" | "subscript";
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
  show: string;
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
   * The same left indent in hundredths of a character (`w:leftChars`) — how WPS spells a Chinese indent. Over the
   * 39 distinct WPS-written bodies in the bundle (`.scratch/indent-step.ts`, deduped by content hash) a positive
   * `w:leftChars` arrives alongside its twips value every single time: `200` with `420` 3 times, `400` with `840`
   * 23 times, `303` with `1207` 16 times. The 缇-per-character those pairs work out to (210, 402, 399) are the
   * paragraph's own font sizes, so the two spellings say the same indent twice — and 420 缇 is also the
   * `w:defaultTabStop` 36 of those 39 files state, which is what makes one step of 增加缩进量 2 字符 = 420 缇 at 10.5pt.
   *
   * The page draws the twips value (no paragraph states a positive `w:leftChars` without it), but the button has
   * to move both: leave the character count behind and the file contradicts itself, and Word lays the paragraph
   * out from the count while our page shows the twips.
   */
  indentLeftChars?: number;
  /**
   * The rest of what `w:ind` can state about the first line and the right edge. `indentRight` and
   * `indentHanging` are twips like `indentLeft` (`w:right`, `w:hanging`); `indentFirstLineChars` is hundredths
   * of a character, which is how WPS writes 首行缩进 2 字符 (`w:firstLineChars="200"`).
   *
   * The first line has two spellings, and a file that states both is not contradicting itself: over the 86
   * WPS-written documents the 1,147 `w:ind` elements carrying a positive pair work out to a per-character
   * width of 180/210/240/264/300/320/440 twips, which are font sizes (9pt, 五号 10.5pt, 12pt, 16pt, 22pt).
   * So they say the same indent twice — except in 88 elements that state a positive `w:firstLine` next to
   * `w:firstLineChars="0"` (83 of them inside a `word/document.xml`, and 70 of those read `w:firstLine="642"`),
   * which is WPS spelling an absolute 磅 value it was given in the dialog. A rule that let the character form
   * win — which is what §17.3.1.14 says — would remove 88 indents WPS itself drew to fix 4 in the opposite
   * shape, so the resolved value is the twips one when it is stated non-zero and the character count otherwise.
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
  pageBreakBefore?: boolean;
  numPr?: { numId: number; ilvl: number };
  /**
   * `w:outlineLvl` (§17.3.2.26) as the file signed it: 0 is the first heading level through 8 the ninth, and
   * **9 is body text**, which is why the raw number is kept rather than folded into "a level or none" —
   * 56 of the corpus's 264 self-stating paragraphs write that 9 (`.scratch/outline-census.ts`), and a 9 on the
   * paragraph is a level its style would otherwise lend being taken back.
   */
  outlineLevel?: number;
}

/** `w:spacing@w:lineRule` (§17.3.1.27) — and the only three literals the 86 WPS-written documents ever state. */
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
}

/** What a `w:trPr` says about its row, minus the row's own record and the children it only carries. */
export interface RowStyle {
  heightTwips?: number;
  heightRule?: string;
  header?: boolean;
  keepTogether?: boolean;
  hidden?: boolean;
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
  pageBreakBefore?: boolean;
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
   * it the mark falls through to the page stylesheet's 12pt, a size no document asked for — 331 of the corpus's
   * 396 such paragraphs resolve elsewhere than 12pt (`NumberLevel.sizePt` covers the other 70, `.scratch/mark-size.ts`).
   * Nothing writes it: the mark's own properties travel as `pPrSlots` text, so a save spells the file's bytes.
   */
  markSizePt?: number;
  /**
   * The two faces this paragraph's own mark resolves to — `w:pPr/w:rPr/w:rFonts@w:ascii` and `@w:eastAsia`
   * (§17.3.1.4), each falling back on its own to what the style chain and `w:rPrDefault` lend, the way
   * `readRunStyle` reads one attribute at a time (§17.7.2 stops per property). They exist for the same thing
   * `markSizePt` does: the numbering symbol is a character run of its own (§17.3.1.12), and where the *level*
   * names no font this is what the file left standing. Over the 86 WPS-written documents, 424 listed paragraphs sit
   * on a level that names none: 172 get the face for their own script from this mark, 234 more from the chain, 18
   * from nothing at all (`.scratch/mark-font.ts`). The two faces are a stack, not a choice — the browser dispatches
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
   * The level has no say here, because the corpus never tests one: 8 of its 617 levels do write
   * `w:lvl/w:rPr/w:color`, and **8/8 of those have no listed paragraph sitting on them** (`.scratch/mark-color.ts`),
   * so no file in hand says which of the two should win. Of the 473 listed paragraphs in the 86 WPS-written
   * documents, **82** resolve a colour at this rung — 74 from the mark itself, 8 from the chain — and all 82 differ
   * from the ink the App's own stylesheet gives a page, which is the number that changes on screen. In 74/74 of the
   * mark-named ones it is the same colour the paragraph's own text resolves to
   * (`.scratch/make-markcolor-fixture.ts`), so what this field fixes is not a difference from the body but the
   * symbol having been left out of the ink the file chose for its paragraph. 391 resolve nothing and stay as they
   * are drawn. Nothing writes it: the mark's properties travel as `pPrSlots` text, so a save spells the file's bytes.
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
  /**
   * The cell's own `w:tcPr`, exactly as the file wrote it. What the page draws comes from the fields
   * above; what a save writes comes from here, so a property set the model cannot explain — `w:tcMar`,
   * `w:cnfStyle`, a pending `w:tcPrChange` — returns the way it went in.
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
  slots?: PropSlot[];
  /**
   * The row's own children that are neither `w:trPr` nor `w:tc` — the `w:tblPrEx` band, which says this row
   * draws with table borders and shading of its own rather than the table's. Across the 86 real WPS tables
   * it is always the row's FIRST child (6257 rows put it before `w:trPr`, 123 before the first cell when the
   * row has no properties at all), so where it sits is carried too: a save that returns it after the cells
   * rewrites 6380 rows in the wrong order. `afterCells` counts the cells it followed and `beforeProps` says
   * whether it preceded `w:trPr`.
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
  /** `w:tblGrid/@w:w` of each column, in order — what a fixed layout sizes its columns from. */
  columnWidths?: number[];
  /** The source text of `w:tblGrid` whole, so every `w:gridCol` and its attributes come back. */
  gridXml?: string;
  /**
   * The table's own `w:tblPr`, exactly as the file wrote it. The fields above are what the page draws;
   * a save writes this, so `w:tblLook`, `w:tblCellMar` and a pending `w:tblPrChange` are not invented
   * but returned — see the note on `TableCell.tcPrXml`.
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
   * the corpus carry one, and without this slot the start survives a save while the end goes, leaving a
   * bookmark no reader can pair.
   */
  betweenRows?: { afterRow: number; tag: string; xml: string }[];
}

/**
 * A child of the body, or of a table cell, that the model has no fields for: a `w:sdt` content control, a
 * `w:customXml` island, anything a producer wrote between the paragraphs. The page shows the text inside it
 * and edits none of it, and a save writes back the source text it came in with, in the place it held.
 * Without this kind, a document with a table of contents in it loses the control and everything anchored
 * in it on the first keystroke — and the shapes are the ECMA-376 ones, since no file in the WPS corpus
 * here has ever contained one.
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
type ParaChild = { tag: string; node?: any; span?: { tag: string; xml: string } };

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

export interface Section {
  pageWidthTwips: number;
  pageHeightTwips: number;
  /**
   * `w:pgMar`. The four body margins are what the page box draws its text column with. `header` and `footer`
   * are a different measurement: ECMA-376 §17.3.2.13 counts them from the **edge of the paper**, not from the
   * body, so a band's distance is only meaningful against the page margin it lives inside. `gutter` is extra
   * width on the binding side, added to the margin it belongs to. All three stay optional: WPS states all
   * seven attributes on all 115 `w:pgMar` in the 86 real files, but a file that says nothing is not owed a
   * number this model invented.
   */
  margins: { top: number; right: number; bottom: number; left: number; header?: number; footer?: number; gutter?: number };
  /**
   * `w:pgBorders`, the line drawn around the page. Absent means the section states no page border at all — the
   * common case, and one this model must not fill in: 24 of the 115 sections in the 86 files this machine's WPS
   * wrote carry the element, the other 91 carry nothing. `offsetFrom` says what the gap is measured against, and
   * none of those 24 spelled it, so they all land on the default §17.3.2.10 gives it — the page edge. (How WPS
   * itself reads an omitted `w:offsetFrom` is 未取证; this is the spec's word, not an observation.) A side that
   * states no `w:val` is left out rather than drawn as a line of unknown shape.
   *
   * What those 24 actually look like is worth stating, because it is the opposite of what a reader of this field
   * would expect: every one writes all four sides, each with `w:val`, `w:sz` and `w:space` and no `w:color`, and
   * **all 96 of those sides say `none`**. Real WPS files carry a page border that was turned off, not one drawn —
   * so a document with a visible frame around the page is, in this corpus, a shape that has never been observed.
   * The reading and drawing of one is spec-derived; its fidelity to WPS is 未取证.
   */
  pageBorders?: { offsetFrom: "page" | "text"; sides: Partial<Record<PageBorderSide, PageBorder>> };
}

/**
 * The page a document is drawn on when its file spelled no section: A4 with a one-inch margin all round. This
 * has one owner because three places ask the same question — the view's fallback, a new document's first page,
 * and 「页面设置」 pre-filling a file that stated nothing — and a page the app shows and a save writes have to be
 * the same box.
 */
export const DEFAULT_SECTION: Section = { pageWidthTwips: 11906, pageHeightTwips: 16838, margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 } };

/** A copy, because a section a caller mutates must not move the page every other document falls back to. */
const copySection = (s: Section): Section => ({
  pageWidthTwips: s.pageWidthTwips,
  pageHeightTwips: s.pageHeightTwips,
  margins: { ...s.margins },
  // The sides copy by value, the lines inside them do not: a `PageBorder` is read once and never edited in
  // place, so giving each copy its own map is enough to keep a caller's add-or-drop of a side contained.
  pageBorders: s.pageBorders ? { offsetFrom: s.pageBorders.offsetFrom, sides: { ...s.pageBorders.sides } } : undefined,
});

/**
 * What 「页面设置」 changes on the section the page is drawn from: the seven `w:pgMar` distances and the paper's
 * direction. Distances left out are left alone — a file that never stated `w:header` keeps saying nothing about
 * it, because the number this model would write is a guess at the producer's default, not the user's intent.
 *
 * `landscape` is the direction the page is wanted in, and it is what decides which edge is which: the two
 * dimensions already on the section are re-assigned rather than replaced, so A4 stays A4 when it turns. Measured
 * on the 86 files this machine's WPS wrote: all 22 landscape sections spell `w:orient="landscape"` **and** put
 * the long edge in `w:w`, and none of the 93 portrait sections spells `w:orient` at all. So turning back to
 * portrait removes the attribute; writing `w:orient="portrait"` would be a spelling no file here uses.
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
   * keeps the sheet it had and only its direction changes. `w:pgSz` has no presets: every one of the 115
   * sections in the 86 files this machine's WPS wrote spells its own two numbers (89 and 22 of them A4's
   * 11906×16838 and 16838×11906, 4 a near-A4 11900×16840), so a name like 「A4」 belongs to the dialog, not here.
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
   * The font this level names for its own mark (`w:lvl/w:rPr/w:rFonts`). It is part of the glyph, not decoration:
   * over the 86 WPS-written documents the 170 `bullet` levels put the mark in `w:lvlText` as a literal
   * private-use codepoint (F06E 44, F0A7 42, F075 38, F06C 30, F0B7 6, F0D8 4, and 6 plain `o`), and say which
   * font it belongs to in the level's own `w:rPr` (Wingdings 158, Symbol 6, other 6) — drawn in the paragraph's
   * font those codepoints are tofu (`.scratch/sym-count.ts`). `w:sym` appears 0 times in that corpus: the
   * element this model already reads for body runs is not how WPS spells a list mark, so the level's font is
   * where the shape comes from here.
   */
  font?: string;
  /**
   * The size this level names for its own mark, in points (`w:lvl/w:rPr/w:sz` ÷ 2 — the file counts half-points,
   * the same unit `readRunStyle` converts for body runs). §17.3.1.12 gives the symbol a character run of its own,
   * so this is the mark's size and not the paragraph's: across the 86 WPS-written documents **70 paragraphs** sit on
   * a level that states one — 42 of them `bullet` levels whose mark is smaller than that paragraph's own text
   * (42/42), 28 `decimal` ones whose stated size equals the text already resolved to (28/28, so the page looks the
   * same either way). 5 distinct values over the 68 levels that write it; the split is `.scratch/mark-size.ts`,
   * which also cross-checks every one of those 70 against the raw `w:lvl` with no disagreement. The other children
   * of that `w:rPr` were counted in the same pass and left un-modelled on purpose: `w:b`/`w:bCs` appear on 8 used
   * levels and all 8 agree with the paragraph's own boldness, so drawing them changes no page, while `w:i`, `w:u`,
   * `w:color`, `w:caps`, `w:strike` and the rest appear on 4–28 levels that **no** paragraph in the corpus lists
   * with them.
   */
  sizePt?: number;
}

/**
 * The definition `ensureListNumbering` writes when a file names neither shape. Both marks are ordinary characters,
 * not private-use codepoints: the corpus's own bullets are PUA glyphs whose font the level states (Wingdings 158,
 * Symbol 6 of the 170 bullet levels, `.scratch/sym-count.ts`), and an invented definition carries no such font, so
 * a codepoint would draw tofu. `%1.` is the ten-digit shape the same probe saw on 24 of the decimal levels.
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
 * edges, which is the rule the files follow: of the 5,382 `w:t` in the local body corpus it puts the
 * attribute on 960, and of the 70 in its header/footer parts on exactly the 17 that open or close with
 * whitespace — never on the other 53. Spelling it on every one is not a wrong document, it is a
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
 * switch. Measured on the local corpus, where of the 21 fields that sit at run level in a header or footer
 * part 16 spell ` PAGE  \* MERGEFORMAT `, 2 spell `PAGE` bare and 2 count pages; the 1 that is none of these
 * stays a reason to hold the part as its own bytes.
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
 * join the band's. Measured on the local corpus, every one of the 16 runs of this shape is a `w:rPr` plus an
 * `mc:AlternateContent` wrapping an anchored box, and every one sits alone in its paragraph.
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
 * the measured 16 all carry a `w:txbxContent` inside the Choice's drawing *and* another inside the Fallback's
 * VML shape, both holding the same one paragraph. Reading a box means finding both copies; writing one means
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
 * WPS wraps a footer's page number in `w:sdt`: measured on the local corpus, 3 of the bands still held
 * read-only are one or two content controls deep around a paragraph carrying `PAGE`, and in two of them
 * `NUMPAGES` and typed characters as well. The control itself is the file's business, so it travels as its
 * own bytes and only the paragraph inside it is re-spelled.
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
        // The 16 bands this corpus holds are an anchored text box whose single inner paragraph carries the
        // page number and the two characters the page shows around it. The box is not a label over the
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
      carry: { key: newKey(), tag: "fldChar", xml: kids.slice(i, end + 1).map((k) => k.xml).join(""), show: instr.includes("NUMPAGES") ? "总页数" : "页码" },
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
 * the texts: the root tag keeps every declaration the file wrote (the local corpus averages 17 prefixes
 * per part, every one of them carrying `mc:Ignorable`), a paragraph whose text did not change travels
 * byte for byte, and a paragraph that did keeps its own `w:pPr` and its first run's `w:rPr`. Rebuilding
 * the part from `paragraphs` instead — which is what this did — takes the style, the alignment, the font
 * size and the namespace declarations off every page the user typed in: of the 86 parts in that corpus
 * a text edit could reach, 85 carry a `w:pStyle` and 18 a `w:jc`.
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
 * `separate` before the number it last showed. 16 of the 21 run-level fields in that corpus's header and
 * footer parts spell the instruction exactly this way, and all 21 cache a number — which is why this one
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
  basedOn?: string;
  next?: string;
  rPr?: RunStyle;
  pPr?: Partial<ParaStyle>;
}

export class WordDocument {
  blocks: Block[] = [];
  sections: Section[] = [];
  header?: HeaderFooter;
  footer?: HeaderFooter;
  comments: DocComment[] = [];
  private styles = new Map<string, StyleDef>();
  /** The `w:default="1"` paragraph style: what a paragraph that names no style falls back on (§17.7.2). */
  private defaultParaStyleId?: string;
  /**
   * `w:docDefaults/w:rPrDefault`, kept to the properties the ladder lends: the rung every paragraph chain ends
   * on, §17.7.2. Only one rung, and it goes in below the chain rather than inside it — see `paragraph`.
   */
  private docDefaultRun: RunStyle = {};
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
      JSON.stringify([hfSavedState(this.header), hfSavedState(this.footer)]) +
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
    const stated = (defaults ? runStyle(defaults) : {}) as Record<string, unknown>;
    const floor: Record<string, unknown> = {};
    for (const k of LENT_RUN_KEYS) if (stated[k] !== undefined) floor[k] = stated[k];
    this.docDefaultRun = floor as RunStyle;
    for (const s of asArray(root?.style)) {
      const id = s["@_styleId"];
      this.styles.set(id, {
        name: s.name?.["@_val"] ?? id,
        basedOn: s.basedOn?.["@_val"],
        next: s.next?.["@_val"],
        rPr: s.rPr ? runStyle(s.rPr) : undefined,
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
   * The depth is not a hypothetical. Over the 86 WPS-written documents (`.scratch/basedon-chain.ts`), 1,182 of
   * its 1,337 paragraph styles carry `w:basedOn`, the chains run one to four styles deep with no cycles and no
   * dangling reference, and 879 paragraphs ask for a property only an ancestor states: 501 line heights — every
   * one of them 固定值, so they go from the browser's default line height to a `px` height the moment they are
   * read — plus 327 alignments, 220 indents (122 left, 65 right, 33 first-line) and 115 段前段后 gaps. A further
   * 38,475 paragraphs name no style at all, and 3,147 of those take a property from the default paragraph style
   * (1,956 alignments and 613 固定值 line heights among them), including 32 paragraphs whose named chain states
   * nothing while the default style states a first-line indent. Below the default style the ladder has one more
   * rung, §17.7.2's `w:docDefaults/w:pPrDefault`, which states no paragraph property in any of the 86 and so is
   * not read here; `w:basedOn` is walked with a seen-set because a chain that circled would otherwise hang the
   * read.
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
      // that names no style, so a chain that never reaches it still ends there: 32 paragraphs in 16 of the 86
      // name a style whose ancestors state no first-line indent while the default style does.
      const based = this.styles.get(cur)?.basedOn;
      cur = based ?? (this.defaultParaStyleId === undefined || seen.has(this.defaultParaStyleId) ? undefined : this.defaultParaStyleId);
      if (!stated) continue;
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
   * The depth is measured, not assumed (`.scratch/run-lent-depth.ts`): 1,961 of the corpus's 2,378 styles state a
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
            font: rFonts?.["@_ascii"] ?? rFonts?.["@_hAnsi"],
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
    // Across the 86 WPS-written documents 455 of 466 listed paragraphs state a mark run, 326 of them a size
    // (`.scratch/mark-size.ts`); of the 396 whose *level* states no size, **every one** resolves to a number —
    // 298 from this mark, 98 from the paragraph's style chain, 0 from nothing at all. 331 of the 396 differ from
    // the page stylesheet's 12pt, so a mark that ignores this chain draws at a size no file asked for.
    const markSz = num(pPr.rPr?.sz?.["@_val"]);
    // Faces the same way, one attribute at a time: a mark that names only `w:eastAsia` must not hand back an
    // undefined Latin face that erases the chain's, and an empty `w:ascii=""` states nothing (§17.7.2).
    const markRFonts = asArray(pPr.rPr?.rFonts)[0];
    const markAscii = markRFonts?.["@_ascii"] || undefined;
    const markEast = markRFonts?.["@_eastAsia"] || undefined;
    // Colour the same way, with the two filters `readRunStyle` applies to a run's own: `auto` is no statement,
    // and a value that is not a colour is no statement either — the mark then falls through to the chain.
    const markColorRaw = asArray(pPr.rPr?.color)[0]?.["@_val"];
    const markColor = markColorRaw && markColorRaw !== "auto" ? hex(markColorRaw) : undefined;
    const para: Paragraph = {
      kind: "paragraph",
      styleId,
      styleName: styleId ? this.styles.get(styleId)?.name : undefined,
      align: pPr.jc?.["@_val"] ?? lent?.align,
      indentLeft: ind.indentLeft ?? lent?.indentLeft,
      indentLeftChars: ind.indentLeftChars ?? lent?.indentLeftChars,
      indentRight: ind.indentRight ?? lent?.indentRight,
      indentFirstLine: ind.indentFirstLine ?? lent?.indentFirstLine,
      indentFirstLineChars: ind.indentFirstLineChars ?? lent?.indentFirstLineChars,
      indentHanging: ind.indentHanging ?? lent?.indentHanging,
      spacingBefore: sp.spacingBefore ?? lent?.spacingBefore,
      spacingAfter: sp.spacingAfter ?? lent?.spacingAfter,
      lineSpacing: sp.lineSpacing ?? lent?.lineSpacing,
      lineSpacingRule: sp.lineSpacingRule ?? lent?.lineSpacingRule,
      pageBreakBefore: pPr.pageBreakBefore !== undefined,
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
    para.asRead = {
      styleId: para.styleId,
      align: para.align,
      indentLeft: para.indentLeft,
      indentLeftChars: para.indentLeftChars,
      indentRight: para.indentRight,
      indentFirstLine: para.indentFirstLine,
      indentFirstLineChars: para.indentFirstLineChars,
      indentHanging: para.indentHanging,
      spacingBefore: para.spacingBefore,
      spacingAfter: para.spacingAfter,
      lineSpacing: para.lineSpacing,
      lineSpacingRule: para.lineSpacingRule,
      pageBreakBefore: para.pageBreakBefore || undefined,
      numPr: para.numPr,
      outlineLevel: para.outlineLevel,
    };
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
        if (step.span) para.inlines.push({ text: "", carry: { key: this.carryKey++, tag: step.tag, xml: step.span.xml, show: carriedText(step.span.xml) } });
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
      // A run that is only field machinery keeps the name of the machinery it holds, so the model and the
      // page can say which part of the field this is rather than just "a run".
      const kind = node !== undefined && s.tag === "r" ? machineryOf(node) : "";
      if (node !== undefined && kind) return { tag: kind, span: s };
      return node === undefined ? { tag: s.tag, span: s } : { tag: s.tag, node, span: s };
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
    const own = runStyle(r.rPr ?? {});
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
          prev: runStyle(change.rPr ?? {}),
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
      const info = rowInfoOf(tr.trPr, slots, ownAttrs(rowSpan?.xml), extras);
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
   * Reusing the document's definitions was the whole strategy before, and the numbers say what that bought
   * (`.scratch/list-availability-audience.ts`, 86 WPS-written documents): 项目符号 had nothing to point at in
   * **80 of 86** files and 编号 in **48 of 86** — 47 carry no `word/numbering.xml` at all, which is also the shape
   * of 文件 → 新建. In the 38 files where 编号 did answer, the id picked was the first one whose level 0 was not a
   * bullet, and in **29 of those 38** that level says `w:numFmt val="chineseCounting"` — the button labelled 1.
   * drew 一、. So the pick is now exact (`decimal` for 编号), and a file that names neither shape gets a definition
   * instead of a grey button, the way WPS writes one when you click on a blank page.
   *
   * What gets invented is the core the corpus really does state for a level: `w:start`, `w:numFmt`, `w:lvlText`,
   * `w:lvlJc` — all five shapes counted in column D of that probe carry those four, in §17.4's order. What is
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
   * 增加缩进量／减少缩进量 moves a paragraph that spells its indent in twips. Measured over the 39 distinct
   * WPS-written bodies in the bundle: 36 state 420 and 3 state 720, and none leaves it out
   * (`.scratch/indent-step.ts`). `word/settings.xml` is not a part this model owns, so this reads it on demand
   * and a save still carries the bytes the file wrote.
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
   * It stops at 0 rather than going negative: WPS writes `w:left="0"` out loud 187 times in these same bodies,
   * which is its way of saying 回到边内. Whether the button keeps going past the left margin onto negative
   * values (the corpus holds `w:left="-1"`, `-41`, `-178`) is 未取证 — no file here was made by clicking it.
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

  /**
   * §17.7.2.24 `w:next`: the style a paragraph takes when Enter splits the one that names `styleId`. The split
   * keeps the paragraph's own direct properties and drops the ones the previous style was lending it, so the
   * new paragraph reads as a paragraph whose `w:pPr` says only `w:pStyle`.
   *
   * Only a style the paragraph *names* answers the question — a paragraph that names none falls back on the
   * default paragraph style for formatting (§17.7.2) but not for a next style. The 86 WPS-written documents say
   * why (`.scratch/next-resolve.ts`): 668 of their 2,378 styles state `w:next`, a further 358 inherit it from a
   * `w:basedOn` ancestor, and every one of the 668 targets is a declared paragraph style that is not the style
   * stating it. Left alone, that reaches body text: 63 of the 86 have a `w:next` on their default paragraph
   * style, 16,783 of their 30,322 paragraphs name no style, and 924 of those would land on a heading style —
   * every second Enter turning ordinary prose into a heading. Nothing in the corpus behaves that way: of its
   * 230 adjacent paragraph style changes, 0 match the preceding style's `w:next`, and of its 934 paragraphs that
   * name a style, 0 name a default style's next. A heading's next, which is the case the feature is for, always
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
    const own = ownParaStyle(this.paragraphStyleOf(previous.styleId), paraStyleOf(previous));
    delete own.styleId;
    const fields: ParaStyle = { ...lent, ...own };
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
    if (!tags) return xml;
    const open = new RegExp(`(<${this.sectPrefix}:sectPr[^>]*>)`);
    return open.test(xml) ? xml.replace(open, `$1${tags}`) : xml;
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
    const p = this.sectPrefix;
    const tag = (kind: "header" | "footer", hf?: HeaderFooter): string =>
      hf && !this.sectHas[kind] ? `<${p}:${kind}Reference ${p}:type="default" r:id="${hf.id}"/>` : "";
    return tag("header", this.header) + tag("footer", this.footer);
  }

  /**
   * Cites a header/footer this session created from the `sectPr` that owns the band, when that is a
   * paragraph's section break rather than the body's. The reference goes after the opening tag of the
   * first top-level paragraph carrying one, which is the section the document starts with.
   */
  private citeBandReference(kind: "header" | "footer", id: string): void {
    const p = this.sectPrefix;
    const tag = `<${p}:${kind}Reference ${p}:type="default" r:id="${id}"/>`;
    for (const b of this.blocks) {
      if (b.kind !== "paragraph" || !b.pPrSlots?.length) continue;
      const at = b.pPrSlots.findIndex((s) => s.tag === "sectPr");
      if (at < 0) continue;
      const slot = b.pPrSlots[at]!;
      b.pPrSlots[at] = {
        tag: "sectPr",
        xml: /\/>$/.test(slot.xml)
          ? slot.xml.replace(/(\s*)\/>$/, `>${tag}</${p}:sectPr>`)
          : slot.xml.replace(new RegExp(`(<${p}:sectPr[^>]*>)`), `$1${tag}`),
      };
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
    if (!hf?.editable) return false;
    hf.segments = segments.map(coalesceHfSegments);
    hf.paragraphs = hf.segments.map(hfSegmentsText);
    return true;
  }

  /**
   * Puts a `PAGE` field in an editable header or footer, spelled as the files write it and carried as one
   * atom, so the text around it stays editable and a save gives the field back unchanged. `paragraph` counts
   * from the top, or from the end when negative; `offset` counts the characters of text before it goes in.
   */
  insertPageNumber(which: "header" | "footer", paragraph = -1, offset?: number): boolean {
    const hf = which === "header" ? this.header : this.footer;
    if (!hf?.editable) return false;
    const segs = (hf.segments ?? hf.paragraphs.map((t) => [{ text: t }])).map(copyHfSegments);
    const at = paragraph < 0 ? segs.length + paragraph : paragraph;
    const target = segs[at];
    if (!target) return false;
    const prefix = /^<([\w.-]+):(?:hdr|ftr)/.exec(hf.source)?.[1] ?? "w";
    target.splice(0, target.length, ...putHfCarry(target, offset, { key: this.carryKey++, tag: "fldChar", xml: pageFieldXml("PAGE", prefix), show: "页码" }));
    return this.setHeaderSegments(which, segs);
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

  /** Header/footer parts with their regenerated XML; `dirty` ones a save must write. */
  headerFooterParts(): { part: string; xml: string; dirty: boolean }[] {
    const out: { part: string; xml: string; dirty: boolean }[] = [];
    for (const hf of [this.header, this.footer]) {
      if (!hf) continue;
      out.push({ part: hf.part, xml: headerFooterXml(hf), dirty: hf.segments ? !sameHfSegments(hf.segments, hf.baselineSegments ?? []) : JSON.stringify(hf.paragraphs) !== JSON.stringify(hf.baseline) });
    }
    return out;
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
   * goes is no longer a table: an empty `w:tbl` is not something Word writes or reads back.
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
        if (!stays.length) return [];
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
            const properties = rec.tag === "tcPrChange" || rec.tag === "cellMerge";
            const goes = (rec.tag === "cellIns" && outcome === "reject") || (rec.tag === "cellDel" && outcome === "accept");
            if (goes) {
              goers.push({ t, row, cell, at: columnStarts(row)[row.indexOf(cell)] ?? 0, rec });
              continue;
            }
            touched += 1;
            settled.push(rec);
            // A place-record has no properties of its own to put back, so it only ever leaves the set.
            if (cell.tcPrXml) cell.tcPrXml = decideSetXml(cell.tcPrXml, rec, properties ? outcome : "accept") || undefined;
            if (properties && outcome === "reject") restoreCellStyle(cell, rec.prev, coversWholeSet(rec.tag));
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
   * previous set, so a property it does not mention was not set — and alignment and indent come from
   * the paragraph's style when the record leaves them out, which is where they belonged before.
   */
  private restoreParaStyle(p: Paragraph, fmt: ParagraphFormatRevision): void {
    const prev = fmt.prev;
    for (const key of PARA_STYLE_KEYS) delete p[key];
    Object.assign(p, prev);
    p.styleName = p.styleId ? this.styles.get(p.styleId)?.name : undefined;
    if (fmt.prevSlots?.length) {
      // The record wrote more than the model has a field for, so the paragraph takes its set as written;
      // syncing the slots from fields would re-spell the named ones and delete the rest.
      p.pPrSlots = fmt.prevSlots;
      p.asRead = { ...prev };
    } else {
      // The record spells the properties the paragraph owns, so the owned children of its `w:pPr` follow the
      // restored fields before anything the new style lends is folded back in.
      syncParaSlots(p);
    }
    const inherited = this.paraLent(p.styleId);
    p.align = p.align ?? inherited?.align;
    // Every `w:ind` attribute resolves on its own, the way the paragraph's first read did it: a record that
    // left the indent out hands the style's indent back, and one that said `w:firstLine="0"` keeps that zero.
    p.indentLeft = p.indentLeft ?? inherited?.indentLeft;
    p.indentLeftChars = p.indentLeftChars ?? inherited?.indentLeftChars;
    p.indentRight = p.indentRight ?? inherited?.indentRight;
    p.indentFirstLine = p.indentFirstLine ?? inherited?.indentFirstLine;
    p.indentFirstLineChars = p.indentFirstLineChars ?? inherited?.indentFirstLineChars;
    p.indentHanging = p.indentHanging ?? inherited?.indentHanging;
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
   * on the 86 files this machine's WPS wrote: 115 sections in 86 documents (68 one, 7 two, 11 three), and
   * every body-level `sectPr` sits at the end of `w:body`.
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
    // Read by name, not by position: WPS puts `r:id` before `w:type` on every one of the 59 header and 57
    // footer references in the local corpus, where the order this used to look for is Word's, so a real
    // WPS page reported no header at all — nothing to show, nothing to edit, nothing to splice back.
    const bound = (kind: "header" | "footer") =>
      childSpans(elementContent(raw))
        .filter((s) => s.tag === `${kind}Reference`)
        .map((s) => attrsOf(s.xml));
    // Which band the page shows is the section's own statement, not the order it wrote its references in
    // (§17.10.3): the first-page band belongs to the opening page only when the section states `w:titlePg`, and
    // an `even` band belongs to `w:evenAndOddHeaders` — which none of the 86 corpus files' `word/settings.xml`
    // turns on, even though 29 of their 115 sections still cite an even reference. The templates show why the
    // switch is the file's and not the renderer's: 14 sections cite default, even *and* first for both bands
    // while stating no `w:titlePg`, which is a page where 首页不同 is available and off.
    const firstPage = childSpans(elementContent(raw)).some((s) => s.tag === "titlePg");
    const shown = (kind: "header" | "footer") => {
      const list = bound(kind);
      const defaultBand = list.find((a) => a["type"] === "default" && a["id"]);
      // A section that asks for a different first page but cites no first-page band still has the part it holds:
      // blanking the page here would take away text the file wrote. What WPS shows for that shape is 未取证.
      return (firstPage ? list.find((a) => a["type"] === "first" && a["id"]) : undefined) ?? defaultBand;
    };
    const picks: Record<"header" | "footer", Record<string, string> | undefined> = { header: shown("header"), footer: shown("footer") };
    // Whether the section already cites the band this page shows — the question a save asks before adding a
    // reference of its own. A band the section cites but the page does not show is not a band the model holds.
    this.sectHas = { header: !!picks.header?.["id"], footer: !!picks.footer?.["id"] };
    for (const kind of ["header", "footer"] as const) {
      const id = picks[kind]?.["id"];
      if (!id) continue;
      const rel = rels.get(id);
      if (!rel) continue;
      const part = pkg.resolve("word/document.xml", rel);
      if (!pkg.parts.has(part)) continue;
      const partRaw = pkg.text(part) ?? "";
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
      const hf: HeaderFooter = {
        part,
        id,
        editable,
        paragraphs,
        baseline: [...paragraphs],
        ...(segments ? { segments, baselineSegments: segments.map(copyHfSegments) } : {}),
        ...(segments && parsed.some((p) => p.boxed) ? { boxedParagraphs: parsed.map((p) => p.boxed) } : {}),
        source: partRaw,
      };
      if (kind === "header") this.header = hf;
      else this.footer = hf;
    }
  }

  /**
   * Relationships the body carried, every one of them and with the part each resolves to.
   *
   * Dropping the picture and hyperlink relationships here because "the writer re-derives those" was
   * half true: the writer re-derives the ones its own tags name, but a drawing, a `w:fallback` VML
   * `v:imagedata` or a hyperlink inside a structure that travels by the byte still cites the id the
   * file wrote. Measured on 86 real WPS files: the originals hold zero dangling citations, and a save
   * that renumbered image relationships left one carried `r:id` pointing at nothing — the shape lost.
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
export const RUN_STYLE_KEYS: (keyof RunStyle)[] = ["bold", "italic", "underline", "strike", "sizePt", "color", "highlight", "font", "fontEastAsia", "vertAlign"];

/**
 * The run properties a style may lend a run — all of them, both faces included. `font` was held back while the
 * model had only the one face from `w:rFonts@w:ascii`: 350 of the corpus's style-side `w:rFonts` elements name a
 * different face for `w:eastAsia` (`.scratch/run-font-shape.ts`), and handing a run the Latin one as its only
 * face would draw its Chinese in a face the file never chose for it. With `fontEastAsia` read as the second face,
 * the pair says what the element says: 2,928 of 61,702 modelled runs end up with two different faces from the
 * ladder, and the page dispatches them by script (`render.ts`). The 93 style-side elements that state only
 * `w:eastAsia` now reach their runs too.
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
 * Which unit a `w:spacing@w:line` is written in. Counted over the 25,801 `w:spacing` elements in the 86
 * WPS-written documents that state a height or a gap (their `word/document.xml`, `word/styles.xml` and
 * header/footer parts together): `w:auto` 13,396, `w:exact` 5,812, `w:atLeast` 991, and 5,602 elements that
 * state 段前段后 but no line at all. Two facts fall out of that. Every element that states a positive `w:line`
 * does state its rule — there is no "height with no unit" to resolve against §17.3.1.27's default — and 6,802
 * of the 20,198 positive heights (a third) are 固定值 or 最小值, counted in twips rather than in the 240ths of a
 * line the drawing used for all of them; on the rendered page that is 5,859 heights in `px` and 967 in `max()`
 * out of the 24,458 line heights the 86 files draw (`.scratch/spacing-draw-impact.ts`). Nothing outside the
 * three literals appears, so an unrecognised one stays unread rather than being guessed at; the element itself
 * still travels as the file's own text.
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
 * Read through the style because 5,110 of the paragraphs in the 86 WPS-written documents ask for a line height
 * that way: they state no `w:line` themselves while the style named by their `w:pStyle` does (4,847 of them
 * `w:auto`, 247 `w:exact`, 16 `w:atLeast` — `.scratch/spacing-lent.ts`). The same elements lend 段前 to 64
 * paragraphs and 段后 to 57 more. A further 501 heights, 54 段前 and 61 段后 come only from higher up the
 * `w:basedOn` chain, and 613 heights with 478 of each gap from the default paragraph style, which is what
 * `paraLent` walks to; the counts are in `.scratch/basedon-chain.ts`.
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
 * What a `w:ind` states, as fields. Only the spellings the files really use on their own are read: over the
 * 29,207 `w:ind` elements in the `word/document.xml` of the 86 WPS-written documents, a positive
 * `w:firstLineChars` appears with no `w:firstLine` beside it 232 times, while a positive `w:leftChars`,
 * `w:rightChars` or `w:hangingChars` never does — there the character form always arrives alongside its twips
 * value, so only the first line needs a field of its own counted in characters.
 *
 * Attributes the element does not state stay out of the result rather than coming back as `undefined`: a
 * `w:pPrChange` has to record which properties it wrote, and a paragraph has to tell "the file said zero" from
 * "the file said nothing", since the second is what lets its style's indent through.
 */
function indOf(ind: any): Partial<ParaStyle> {
  const out: Record<string, number> = {};
  const take = (field: keyof ParaStyle, attr: string): void => {
    const v = num(ind?.[`@_${attr}`]);
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

/**
 * The numbering a `w:pPr` states, taken as one property: `w:numPr` is not a set of attributes a paragraph can
 * merge piece by piece the way it merges `w:ind`, so the first style in the chain that writes it settles the
 * paragraph's list, and the styles below it get no say.
 *
 * `w:numId` absent reads as 0, which is what the files here mean by it. Of the 981 paragraphs that write their
 * own `w:numPr`, 517 write `w:numId w:val="0"` — the value that cancels numbering (§17.3.4.25) — and not one of
 * them leaves `w:numId` out. On the style side, 11 of the 48 paragraphs that borrow the property borrow a `w:numPr`
 * that writes only `w:ilvl` (2 of the 58 style-side `w:numPr` elements, in 6 of the 86 files); read as 0 they
 * leave those paragraphs unnumbered, which is what the page showed before this was read at all
 * (`.scratch/num-lent-shapes.ts`, `.scratch/style-numpr-census.ts`).
 */
function numPrOf(numPr: any): { numId: number; ilvl: number } | undefined {
  if (!numPr) return undefined;
  return { numId: Number(numPr.numId?.["@_val"] ?? 0), ilvl: Number(numPr.ilvl?.["@_val"] ?? 0) };
}

/**
 * The outline level a `w:outlineLvl` states (§17.3.2.26), as the file signed it: 0…8 are heading levels one to
 * nine and 9 is body text. Of the 86 WPS-written documents' 46,417 paragraphs, 264 state one themselves
 * (0 × 6, 1 × 160, 2 × 42, 9 × 56) and 582 more take it from their style chain, while 433 paragraph styles state
 * one (`.scratch/outline-census.ts`). A value outside 0…9 is read as no statement: the files here never write one.
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
  if (pPr.pageBreakBefore !== undefined) s.pageBreakBefore = true;
  const numPr = numPrOf(pPr.numPr);
  if (numPr) s.numPr = numPr;
  const outline = outlineOf(pPr.outlineLvl);
  if (outline !== undefined) s.outlineLevel = outline;
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
 * `w:pageBreakBefore` is deliberately out: no paragraph in the 86 WPS-written documents asks a style for a
 * page break it does not state itself (`.scratch/basedon-chain.ts` counts the `w:pPr` children the styles do
 * write, and `w:pageBreakBefore` is not among the ones a paragraph borrows).
 */
const LENT_PARA_KEYS: (keyof ParaStyle)[] = [
  "align", "indentLeft", "indentLeftChars", "indentRight", "indentFirstLine", "indentFirstLineChars",
  "indentHanging", "spacingBefore", "spacingAfter", "lineSpacing", "lineSpacingRule", "numPr", "outlineLevel",
];

/** Every paragraph property the model holds for it, from its own `w:pPr` and from its styles. */
export function paraStyleOf(p: Paragraph): ParaStyle {
  const out: Record<string, unknown> = {};
  for (const key of PARA_STYLE_KEYS) if (p[key] !== undefined) out[key] = p[key];
  return out as ParaStyle;
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
 * borrowed property as the paragraph's own: §17.13.5.19 records the direct formatting, and 159 of the corpus's
 * paragraphs state no `w:pPr` at all while still showing a whole set of properties, because a style lends them
 * (`.scratch/para-no-pr.ts`). The read snapshot answers the question — a value still equal to the one the
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
 * turning back to portrait removes `w:orient` rather than spelling a value no file in the corpus uses. A name
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
 * The order ECMA-376 §17.3 (CT_SectPr) allows a section's children in, which is the order the 115 sections in
 * the 86 files this machine's WPS wrote use: references, then `w:pgSz`, then `w:pgMar`, then the column and grid
 * settings. A geometry element a minimal file never spelled goes in at its place in this sequence rather than at
 * the end, because a `w:pgMar` written after `w:docGrid` is a file Word refuses to open.
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

/** The attributes `w:pgSz` spells, in the order the corpus uses. */
const PG_SZ_ATTRS = ["w", "h", "orient"];
/** The attributes `w:pgMar` spells, in the order all 115 in the corpus use (§17.3.2.13). */
const PG_MAR_ATTRS = ["top", "right", "bottom", "left", "header", "footer", "gutter"];

/** The page box 「页面设置」 asks for: the sheet wanted (or the one the section already had), turned as asked. */
function pageBoxOf(set: PageSetup, widthTwips: number, heightTwips: number): { w: number; h: number } {
  // The two numbers say which sheet, `landscape` only says which way it lies — so the long edge lands in `w:w`
  // for landscape and in `w:h` for portrait. A dialog that sets Letter while the page lies sideways therefore
  // writes 15840×12240, the same spelling the corpus's 22 landscape sections use.
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
 * other child keeps its bytes **and** the text written between the children — the 115 sections in the 86 files
 * this machine's WPS wrote have none, but a producer that indents its markup should not lose it here, because
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
 * `revFormat` spells so a decision can drop it. `w:rStyle`, `w:lang`, `w:kern`, `w:bCs`, `w:spacing`, `w:shd`
 * and `w:eastAsianLayout` are all in the set and all outside the model's ten fields.
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
 * Putting a paragraph's `w:pPr` slots in line with the four properties the model spells for itself, so a
 * toolbar click on alignment or a list button lands *and* the properties the model cannot explain stay.
 * This is the same split `applyCellProps` makes for a cell's `w:tcPr`: the set stays the file's, only the
 * named child moves.
 *
 * A prefix is copied from the set rather than assumed: `w` is only an alias for the words namespace where the
 * part binds it, and a child written with a different prefix than its neighbours would read as another element.
 *
 * `w:spacing` stays out: it states more attributes than the model has fields for, and rewriting it from the
 * fields would take `w:lineRule` down with it. `w:ind` is re-spelled, but only over the attributes the model
 * holds (`left`, `leftChars`, `right`, `firstLine`, `firstLineChars`, `hanging`) — the character spellings it
 * does not hold (`w:rightChars`, `w:hangingChars`) stay in place with the file's own bytes, and the whole child
 * is left untouched whenever none of the six moved. That is what lets 增加缩进量 move a Chinese indent without
 * the file saying 420 缇 in one attribute and 4 字符 in the next.
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
  keep("pageBreakBefore", "pageBreakBefore", p.pageBreakBefore, `<${prefix}:pageBreakBefore/>`);
  keep("jc", "align", p.align, `<${prefix}:jc w:val="${esc(p.align ?? "")}"/>`);
  keep(
    "numPr",
    "numPr",
    p.numPr,
    p.numPr ? `<${prefix}:numPr><w:ilvl w:val="${p.numPr.ilvl}"/><w:numId w:val="${p.numPr.numId}"/></${prefix}:numPr>` : "",
  );
  // 缩进那一句：六个属性跟着字段走，件里说过而模型没有字段的那些（`w:rightChars`、`w:hangingChars`）留在原位。
  // 一个都没动时整句按原件的字节留着——属性之间怎么空格、值怎么写，都是文件自己的事。
  const IND_ATTRS: [keyof ParaStyle, string][] = [
    ["indentLeft", "left"], ["indentLeftChars", "leftChars"], ["indentRight", "right"],
    ["indentFirstLine", "firstLine"], ["indentFirstLineChars", "firstLineChars"], ["indentHanging", "hanging"],
  ];
  const IND_ORDER = ["left", "leftChars", "right", "rightChars", "firstLine", "firstLineChars", "hangingChars", "hanging"];
  const indSlot = slots.find((s) => s.tag === "ind");
  const indMoved = IND_ATTRS.some(([key]) => !samePropValue(p.asRead?.[key], p[key]));
  if (indMoved) {
    const written = [...(indSlot?.xml ?? "").matchAll(/([A-Za-z_][-\w.]*:)?([\w-]+)\s*=\s*"([^"]*)"/g)].map((m) => ({
      name: m[2],
      xml: `${m[1] ?? ""}${m[2]}="${m[3]}"`,
    }));
    for (const [key, name] of IND_ATTRS) {
      const value = p[key] as number | undefined;
      const at = written.findIndex((k) => k.name === name);
      if (value === undefined) {
        if (at >= 0) written.splice(at, 1);
        continue;
      }
      const spell = { name, xml: `${prefix}:${name}="${Math.round(value)}"` };
      if (at >= 0) written[at] = spell;
      else {
        const after = written.findIndex((k) => IND_ORDER.indexOf(k.name) > IND_ORDER.indexOf(name));
        if (after < 0) written.push(spell);
        else written.splice(after, 0, spell);
      }
    }
    wanted.push({ tag: "ind", xml: written.length ? `<${prefix}:ind ${written.map((k) => k.xml).join(" ")}/>` : "" });
  }
  const owned = new Set(["pStyle", "pageBreakBefore", "numPr", "jc", ...(indMoved ? ["ind"] : [])]);
  let next = slots.filter((s) => !owned.has(s.tag));
  for (const w of wanted) {
    const at = next.findIndex((s) => pprRank(s.tag) > pprRank(w.tag));
    next = at < 0 ? [...next, w] : [...next.slice(0, at), w, ...next.slice(at)];
  }
  p.pPrSlots = next.length ? [...next].sort((a, b) => pprRank(a.tag) - pprRank(b.tag)) : undefined;
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

/** What a row says about itself: its own revision, its height, whether it repeats or may split. */
function rowInfoOf(trPr: any, slots?: PropSlot[], attrs?: string, extras?: RowExtra[]): RowInfo | undefined {
  const info: RowInfo = { ...rowStyleOf(trPr) };
  const rev = trPr?.ins ? revisionOf("ins", trPr.ins) : trPr?.del ? revisionOf("del", trPr.del) : undefined;
  if (rev) info.rev = rev;
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
    shading: hex(tcPr?.shd?.["@_fill"]),
    spanCols: num(tcPr?.gridSpan?.["@_val"]),
    spanRows: tcPr?.vMerge === undefined ? undefined : tcPr.vMerge["@_val"] === "restart" ? 1 : 0,
    valign: str(tcPr?.vAlign?.["@_val"]),
    textDirection: str(tcPr?.textDirection?.["@_val"]),
    borders: bordersOf(tcPr?.tcBorders),
  };
  const wrap = onOff(tcPr?.noWrap);
  if (wrap !== undefined) out.noWrap = wrap;
  return out;
}

/** Every table property a `w:tblPrChange` can carry, so a refusal clears the ones the record left out. */
export const TABLE_STYLE_KEYS: (keyof TableStyle)[] = ["styleId", "widthTwips", "widthType", "indentTwips", "align", "layout", "borders"];

/** Every row property a `w:trPrChange` can carry. */
export const ROW_STYLE_KEYS: (keyof RowStyle)[] = ["heightTwips", "heightRule", "header", "keepTogether", "hidden"];

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
 * Putting a cell back the properties a record named. A `w:tcPrChange` is a snapshot of the whole set, so the
 * keys it does not name were not set and clear first; a `w:cellMerge` names the one property it touched
 * (`w:vMerge`, which is `spanRows` here), and the rest of the cell keeps what it has.
 */
function restoreCellStyle(cell: TableCell, prev: CellStyle, whole: boolean): void {
  if (whole) {
    for (const key of CELL_STYLE_KEYS) delete cell[key];
    Object.assign(cell, prev);
    return;
  }
  Object.assign(cell, Object.fromEntries(Object.entries(prev).filter(([, v]) => v !== undefined)));
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

function runStyle(rPr: any): RunStyle {
  const s: RunStyle = {};
  if (rPr.b !== undefined) s.bold = rPr.b["@_val"] !== "0" && rPr.b["@_val"] !== "false";
  if (rPr.i !== undefined) s.italic = rPr.i["@_val"] !== "0" && rPr.i["@_val"] !== "false";
  // `w:u` belongs to this same family and spells its cancellation with an enum name instead of `0`: §17.3.2.25's
  // `ST_Underline` says `none`. Reading that as silence left the property the odd one out — a run that said
  // "no underline" under a style that lends one kept the style's line, because §17.7.2 stops per property at the
  // first rung that *stated* one, and this rung stated it. The files write the sentence: 952 `w:u` with `val`
  // `none` in the 86 body parts, 254 of them on runs this reader resolves into fields (`.scratch/underline-counted.ts`,
  // `.scratch/underline-audience.ts`), and not one of those 254 leaves `w:val` off — the shape read here is the
  // shape WPS writes. None of the 254 sits on a chain that lends a line today, so this is a representation fix,
  // not a pixel fix; what it buys is that the statement survives an edit that should not have changed it.
  if (rPr.u !== undefined) s.underline = rPr.u["@_val"] !== "none";
  // The last two members of the same on/off family. `w:strike` is §17.3.2.19's `ST_OnOff`, so `0` and `false`
  // cancel it just as they cancel `w:b` — reading the element as "struck, whatever it says" put a line through the
  // 2 runs in these 86 body parts that state `<w:strike w:val="0"/>` (of the 56 `w:strike` they write, 36 say on
  // with no `w:val`, 18 sit on a paragraph mark) and give the reader text to hang it on
  // (`.scratch/onoff-where.ts`, `.scratch/onoff-audience.ts`). `w:vertAlign`'s cancellation is an enum name again
  // (§17.3.2.24's `ST_VerAlignRun`: `baseline`), and the bodies state nothing else — all 3860 of them, 717 on runs
  // this reader resolves, none sitting on a chain that lends a superscript — so that half is representation rather
  // than pixels, the same verdict `w:u` got above.
  if (rPr.strike !== undefined) s.strike = rPr.strike["@_val"] !== "0" && rPr.strike["@_val"] !== "false";
  if (rPr.sz) s.sizePt = Number(rPr.sz["@_val"]) / 2;
  const color = rPr.color?.["@_val"];
  if (color && color !== "auto") s.color = hex(color);
  // `none` is kept as the value it is, not folded into the absence above it: §17.17.64 spells its cancellation with
  // that enum name, and the files use it for 4,135 of the 4,605 `w:highlight` elements they state — the other 470
  // all say `yellow`, and no run in these 86 documents leaves `w:val` off (`.scratch/highlight-counted.ts`).
  if (rPr.highlight?.["@_val"]) s.highlight = rPr.highlight["@_val"];
  if (rPr.rFonts) {
    // One element, several faces, and a document may state any subset of them. Each attribute is read on its
    // own: `w:ascii="宋体"` alone must not hand back an `undefined` that erases the chain's Chinese face, since
    // §17.7.2 stops per property at the first level that stated one.
    const ascii = rPr.rFonts["@_ascii"];
    if (ascii) s.font = ascii;
    const east = rPr.rFonts["@_eastAsia"];
    if (east) s.fontEastAsia = east;
  }
  const va = rPr.vertAlign?.["@_val"];
  if (va === "superscript" || va === "subscript" || va === "baseline") s.vertAlign = va;
  return s;
}

/**
 * A link with no colour of its own is drawn underlined. That is a look this reader puts on, not a `w:u` the
 * file wrote — the `w:Hyperlink` character style is what underlines it — so the read snapshot is told as well,
 * or a save would spell into the run an underline nobody asked for.
 */
function linkLook(inline: Inline): void {
  if (inline.color) return;
  inline.underline = true;
  if (inline.asRead) inline.asRead = { ...inline.asRead, underline: true };
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
