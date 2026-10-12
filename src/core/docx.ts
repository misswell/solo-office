import { OpcPackage, asArray, childSpans, children, decodeEntities, elementContent } from "./opc.js";
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
  highlight?: string;
  font?: string;
  vertAlign?: "superscript" | "subscript";
}

export interface Inline extends RunStyle {
  text: string;
  hyperlink?: string;
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
   * which `revFormat` spells. A run is rebuilt from its fields otherwise, and nine fields cannot hold the
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
   * That set as the file wrote it, child by child, when the run carried one. `prev` can only hold the nine
   * properties the model has a field for; a save that rewrote the record from those fields took the record's
   * own `w:kern`, `w:lang` and `w:eastAsia` down with them, and a reviewer deciding the change later would be
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
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: number;
  pageBreakBefore?: boolean;
  numPr?: { numId: number; ilvl: number };
}

/** `w:pPrChange`: the paragraph properties from before someone revised this paragraph. */
export interface ParagraphFormatRevision {
  kind: "format";
  id: number;
  author: string;
  date?: string;
  prev: ParaStyle;
  /**
   * That set as the record wrote it, child by child. `ParaStyle` is the eight properties the model can name,
   * and a paragraph's `w:pPr` holds far more than eight: a refusal rebuilt from those fields put the eight
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
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: number;
  pageBreakBefore?: boolean;
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
   * `w:outlineLvl`, `w:widowControl`, `w:tabs`, `w:pBdr`, the paragraph mark's own fonts, bold, size and
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

export interface Section {
  pageWidthTwips: number;
  pageHeightTwips: number;
  margins: { top: number; right: number; bottom: number; left: number };
}

/**
 * A header or footer part the model can show and (when simple) edit. Rich parts —
 * tables, pictures, fields — stay byte-identical on save and are not editable.
 */
export interface HeaderFooter {
  /** Package path ("word/header1.xml"). */
  part: string;
  /** Relationship id used by the body's sectPr reference. */
  id: string;
  /** True when the part is plain paragraphs only, so text edits round-trip. */
  editable: boolean;
  /** Plain text per paragraph, in order. */
  paragraphs: string[];
  /** Texts as read, so a save knows whether the part needs regenerating. */
  baseline: string[];
  /**
   * The part as the file wrote it, `""` for one this session created. A save splices its edits into
   * these bytes rather than rebuilding the part, so the root's declarations and every paragraph the
   * user did not type in travel untouched.
   */
  source: string;
}

export interface NumberLevel {
  numFmt: string;
  lvlText: string;
  start: number;
  ilvl?: number;
  indent?: number;
}

const TWIPS_PER_PX = 15;
export const twipsToPx = (tw: number): number => tw / TWIPS_PER_PX;

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
  const paras = seen.filter((s) => s.tag === "p");
  let at = -1;
  const body = seen
    .map((span) => {
      if (span.tag !== "p") return span.xml;
      const text = hf.paragraphs[++at];
      if (text === undefined) return "";
      return text === hf.baseline[at] ? span.xml : editedParagraph(span, text);
    })
    .join("");
  const extra = hf.paragraphs.slice(paras.length).map((t) => spelledParagraph(t, "w")).join("");
  return `${prologue}${open}${body}${extra}${close}`;
}

/** A paragraph the model holds text for, rewritten in place: properties kept, runs collapsed into one. */
function editedParagraph(span: { xml: string }, text: string): string {
  const prefix = /^<([\w.-]+):/.exec(span.xml)?.[1] ?? "w";
  const attrs = (/^<[\w.-]*:?p\b([^>]*)>/.exec(span.xml)?.[1] ?? "").replace(/\/\s*$/, "");
  const kids = childSpans(elementContent(span.xml));
  // Properties and empty markers come back as written; run content does not, because the text the model
  // holds is every `w:t` of the paragraph merged into one string, and keeping the runs too would say it
  // twice. A paragraph with several differently formatted runs is therefore set in the first one's.
  const kept = kids.filter((k) => k.tag === "pPr" || /^bookmark(Start|End)$/.test(k.tag)).map((k) => k.xml).join("");
  const first = kids.find((k) => k.tag === "r");
  const rPrefix = /^<([\w.-]+):/.exec(first?.xml ?? "")?.[1] ?? prefix;
  const props = first ? childSpans(elementContent(first.xml)).find((k) => k.tag === "rPr")?.xml ?? "" : "";
  const run = text === "" ? "" : `<${rPrefix}:r>${props}${textXml(text, rPrefix)}</${rPrefix}:r>`;
  return `<${prefix}:p${attrs}>${kept}${run}</${prefix}:p>`;
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

/** A paragraph the part did not have: one run, one text, nothing invented beside them. */
const spelledParagraph = (text: string, prefix: string): string =>
  text === "" ? `<${prefix}:p/>` : `<${prefix}:p><${prefix}:r>${textXml(text, prefix)}</${prefix}:r></${prefix}:p>`;

/** A part this session created, with no bytes to splice into. */
function spelledHeaderFooter(hf: HeaderFooter): string {
  const root = hf.part.includes("footer") ? "ftr" : "hdr";
  const body = hf.paragraphs.length ? hf.paragraphs.map((t) => spelledParagraph(t, "w")).join("") : "<w:p/>";
  return `${DECL}<w:${root} xmlns:w="${NS_W}" xmlns:r="${REL_R}">${body}</w:${root}>`;
}

/** A comment body is one `w:p` per line of the editor's text. */
const commentParagraphs = (text: string): Inline[][] => text.split("\n").map((line) => [{ text: line }]);

/** `w:date` is an `xsd:dateTime`; Word writes second precision, so match it. */
/** Word's `w:date` shape: UTC to the second. */
export const stampNow = (): string => new Date().toISOString().replace(/\.\d{3}Z$/, "Z");

export class WordDocument {
  blocks: Block[] = [];
  sections: Section[] = [];
  header?: HeaderFooter;
  footer?: HeaderFooter;
  comments: DocComment[] = [];
  private styles = new Map<string, { name: string; basedOn?: string; rPr?: RunStyle; pPr?: Partial<Paragraph> }>();
  private numbering = new Map<number, NumberLevel[]>();
  private media = new Map<string, string>();
  private pkg?: OpcPackage;
  private hyperlinks = new Map<string, string>();
  private bodyRels: { id: string; type: string; target: string; targetMode: string; path: string }[] = [];
  private bodySectPr = "";
  /** Namespace declarations read off the document root, so carried markup keeps its prefixes bound. */
  private rootNamespaces: Record<string, string> = {};
  private rootIgnorable = "";
  /** The document root's own attribute text, read as bytes: a save writes this rather than spelling a root. */
  private rootOpenAttrs = "";
  private sectPrefix = "w";
  private sectHas = { header: false, footer: false };
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
      JSON.stringify([this.header?.paragraphs ?? null, this.footer?.paragraphs ?? null]) +
      JSON.stringify(this.comments.map(commentState))
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
    this.readHeaderFooters(pkg, rels);
    this.readComments(pkg, rels);
    // The body read off its own source text rather than only off the parsed tree, so a child the model
    // has no fields for — a content control, a customXml island — keeps its bytes and its place instead
    // of dropping out of the document on the first save.
    for (const child of children(body)) {
      if (child.name === "sectPr") this.sections.push(section(child.node));
    }
    this.blocks.push(...this.blockChildren(body, bodyChildSpans(source), rels));
    const core = pkg.text("docProps/core.xml");
    this.title = core ? /<dc:title>([^<]*)<\/dc:title>/.exec(core)?.[1] : undefined;
    this.readFingerprint = this.fingerprint();
  }

  private readStyles(pkg: OpcPackage): void {
    if (!pkg.parts.has("word/styles.xml")) return;
    const st = pkg.xml("word/styles.xml");
    const root = unwrap(st, "styles");
    const defaults = root?.docDefaults?.rPrDefault?.rPr;
    const defaultStyle: RunStyle = defaults ? runStyle(defaults) : {};
    for (const s of asArray(root?.style)) {
      const id = s["@_styleId"];
      const rPr = s.rPr ? { ...defaultStyle, ...runStyle(s.rPr) } : defaultStyle;
      this.styles.set(id, {
        name: s.name?.["@_val"] ?? id,
        basedOn: s.basedOn?.["@_val"],
        rPr,
        pPr: s.pPr ? { align: s.pPr.jc?.["@_val"], indentLeft: num(s.pPr.ind?.["@_left"]) } : undefined,
      });
    }
  }

  private readNumbering(pkg: OpcPackage): void {
    if (!pkg.parts.has("word/numbering.xml")) return;
    const root = unwrap(pkg.xml("word/numbering.xml"), "numbering");
    const abstracts = new Map<string, NumberLevel[]>();
    for (const a of asArray(root?.abstractNum)) {
      const key = String(a["@_abstractNumId"]);
      abstracts.set(
        key,
        asArray(a.lvl).map((l: any) => ({
          numFmt: l.numFmt?.["@_val"] ?? "decimal",
          lvlText: l.lvlText?.["@_val"] ?? "%1.",
          start: Number(l.start?.["@_val"] ?? 1),
          ilvl: num(l["@_ilvl"]),
          indent: num(l.pPr?.ind?.["@_left"]),
        })),
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
    const para: Paragraph = {
      kind: "paragraph",
      styleId,
      styleName: styleId ? this.styles.get(styleId)?.name : undefined,
      align: pPr.jc?.["@_val"] ?? this.styles.get(styleId)?.pPr?.align,
      indentLeft: num(pPr.ind?.["@_left"]) ?? this.styles.get(styleId)?.pPr?.indentLeft,
      spacingBefore: num(pPr.spacing?.["@_before"]),
      spacingAfter: num(pPr.spacing?.["@_after"]),
      lineSpacing: num(pPr.spacing?.["@_line"]),
      pageBreakBefore: pPr.pageBreakBefore !== undefined,
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
    const numPr = pPr.numPr;
    if (numPr) {
      para.numPr = {
        numId: Number(numPr.numId?.["@_val"] ?? 0),
        ilvl: Number(numPr.ilvl?.["@_val"] ?? 0),
      };
    }
    para.asRead = {
      styleId: para.styleId,
      align: para.align,
      indentLeft: para.indentLeft,
      spacingBefore: para.spacingBefore,
      spacingAfter: para.spacingAfter,
      lineSpacing: para.lineSpacing,
      pageBreakBefore: para.pageBreakBefore || undefined,
      numPr: para.numPr,
    };
    // The paragraph's own mark can be revised: `w:pPr/w:rPr` holds the break that ends this paragraph,
    // signed the same way a run is. Nothing here reads the mark's character properties — the break is
    // what matters, and losing the signature would leave a merge nobody can refuse.
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
      if (name === "r") para.inlines.push(...this.run(step.node, rels, step.span?.xml));
      else if (name === "commentRangeStart" || name === "commentRangeEnd") {
        const id = Number(step.node?.["@_id"]);
        if (Number.isFinite(id)) para.inlines.push(anchorInline({ id, kind: name === "commentRangeStart" ? "start" : "end" }));
      } else if (isRevision(name)) {
        this.revisionRuns(step.node, rels, para.inlines, revisionOf(name, step.node), step.span?.xml);
      } else if (name === "hyperlink") {
        const id = step.node["@_id"];
        for (const inline of this.linkRuns(step.node, rels, step.span?.xml)) {
          if (id) inline.hyperlink = this.hyperlinks.get(String(id)) ?? rels.get(String(id))?.target;
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
  private linkRuns(node: any, rels: Map<string, Rel2>, span?: string): Inline[] {
    const out: Inline[] = [];
    for (const child of this.childSteps(node, span)) {
      if (child.node === undefined) continue;
      if (child.tag === "r") out.push(...this.run(child.node, rels, child.span?.xml));
      else if (isRevision(child.tag)) this.revisionRuns(child.node, rels, out, revisionOf(child.tag, child.node), child.span?.xml);
    }
    return out;
  }

  /** A revision's runs are ordinary runs with a signature on them; deleted text is spelled `w:delText`. */
  private revisionRuns(node: any, rels: Map<string, Rel2>, into: Inline[], rev: Revision, span?: string): void {
    for (const child of this.childSteps(node, span)) {
      if (child.node === undefined) continue;
      if (child.tag === "r") {
        for (const inline of this.run(child.node, rels, child.span?.xml)) into.push({ ...inline, rev });
        continue;
      }
      // A revision wraps a link as readily as a run, and the link keeps its own target.
      if (isRevision(child.tag)) {
        // Word nests a revision inside one when revised text is revised again; the inner signature is
        // the change that moved these characters, so it wins over the one we are already reading.
        this.revisionRuns(child.node, rels, into, revisionOf(child.tag, child.node), child.span?.xml);
        continue;
      }
      if (child.tag !== "hyperlink") continue;
      const id = child.node["@_id"];
      for (const inline of this.linkRuns(child.node, rels, child.span?.xml)) {
        if (id) inline.hyperlink = this.hyperlinks.get(String(id)) ?? rels.get(String(id))?.target;
        linkLook(inline);
        into.push({ ...inline, rev });
      }
    }
  }

  private run(r: any, rels: Map<string, Rel2>, span?: string): Inline[] {
    const style = runStyle(r.rPr ?? {});
    // The run's own `w:rPr` travels as text, the same way the paragraph's does: read off the parsed tree the
    // answer is "which properties does the model have a field for", and nine fields is not what a Chinese
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
    // The record is on the run, so everything the run holds carries it.
    const base = {
      ...style,
      ...(carried ? { rPrSlots: carried, asRead: style } : {}),
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

  /** Referenced style definitions, for a writer that has to regenerate word/styles.xml. */
  styleDef(styleId: string): { name: string; basedOn?: string; rPr?: RunStyle; pPr?: Partial<Paragraph> } | undefined {
    return this.styles.get(styleId);
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

  /** Reference tags the model's header/footer still need inside a sectPr. */
  sectReferenceTags(): string {
    const p = this.sectPrefix;
    const tag = (kind: "header" | "footer", hf?: HeaderFooter): string =>
      hf && !this.sectHas[kind] ? `<${p}:${kind}Reference ${p}:type="default" r:id="${hf.id}"/>` : "";
    return tag("header", this.header) + tag("footer", this.footer);
  }

  /** Rewrites the texts of an editable header/footer; rich parts refuse. */
  setHeaderTexts(which: "header" | "footer", paragraphs: string[]): boolean {
    const hf = which === "header" ? this.header : this.footer;
    if (!hf?.editable) return false;
    hf.paragraphs = paragraphs;
    return true;
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
    return hf;
  }

  /** Header/footer parts with their regenerated XML; `dirty` ones a save must write. */
  headerFooterParts(): { part: string; xml: string; dirty: boolean }[] {
    const out: { part: string; xml: string; dirty: boolean }[] = [];
    for (const hf of [this.header, this.footer]) {
      if (!hf) continue;
      out.push({ part: hf.part, xml: headerFooterXml(hf), dirty: JSON.stringify(hf.paragraphs) !== JSON.stringify(hf.baseline) });
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
      // The record wrote more than the eight fields can name, so the paragraph takes its set as written;
      // syncing the slots from fields would re-spell the eight and delete the rest.
      p.pPrSlots = fmt.prevSlots;
      p.asRead = { ...prev };
    } else {
      // The record spells the properties the paragraph owns, so the owned children of its `w:pPr` follow the
      // restored fields before anything the new style lends is folded back in.
      syncParaSlots(p);
    }
    const inherited = p.styleId ? this.styles.get(p.styleId)?.pPr : undefined;
    p.align = p.align ?? inherited?.align;
    p.indentLeft = p.indentLeft ?? inherited?.indentLeft;
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

  private readHeaderFooters(pkg: OpcPackage, rels: Map<string, Rel2>): void {
    const raw = this.bodySectPr;
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
    this.sectHas = { header: bound("header").length > 0, footer: bound("footer").length > 0 };
    for (const kind of ["header", "footer"] as const) {
      // The default reference wins; a section that binds only a first-page or an even one still has the
      // part it holds.
      const list = bound(kind);
      const pick = list.find((a) => a["type"] === "default" && a["id"]) ?? list.find((a) => a["id"]);
      const id = pick?.["id"];
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
      const paragraphs = (rootSpan ? childSpans(elementContent(rootSpan.xml)) : [])
        .filter((s) => s.tag === "p")
        .map(hfParagraphText);
      const hf: HeaderFooter = {
        part,
        id,
        // Anything a flat string cannot say — a field, a picture, a container, a tracked change, a soft
        // return, an anchored comment — is a reason to leave the part as its own bytes. The prefixes are
        // matched loosely on purpose: a part that binds the wordprocessingml namespace under another name
        // is still rich, and reading only `w:` is how a container slipped through and got dropped.
        editable: !/<(?:[\w.-]+:)?(tbl|drawing|pict|object|fldSimple|fldChar|hyperlink|sdt|ins|del|br|sym)[ />]|<(?:[\w.-]+:)?comment(?:Range|Reference)/.test(partRaw),
        paragraphs,
        baseline: [...paragraphs],
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
export const RUN_STYLE_KEYS: (keyof RunStyle)[] = ["bold", "italic", "underline", "strike", "sizePt", "color", "highlight", "font", "vertAlign"];

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
  // puts the run's whole property set away, not an approximation of it built from nine fields.
  if (fmt.prevSlots) {
    out.rPrSlots = fmt.prevSlots;
    out.asRead = fmt.prev;
  }
  return out;
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
  const indent = num(pPr.ind?.["@_left"]);
  if (indent !== undefined) s.indentLeft = indent;
  const before = num(pPr.spacing?.["@_before"]);
  if (before !== undefined) s.spacingBefore = before;
  const after = num(pPr.spacing?.["@_after"]);
  if (after !== undefined) s.spacingAfter = after;
  const line = num(pPr.spacing?.["@_line"]);
  if (line !== undefined) s.lineSpacing = line;
  if (pPr.pageBreakBefore !== undefined) s.pageBreakBefore = true;
  const numPr = pPr.numPr;
  if (numPr) s.numPr = { numId: Number(numPr.numId?.["@_val"] ?? 0), ilvl: Number(numPr.ilvl?.["@_val"] ?? 0) };
  return s;
}

/** Every paragraph property a `w:pPrChange` carries, so a refusal can clear the ones the record dropped. */
export const PARA_STYLE_KEYS: (keyof ParaStyle)[] = ["styleId", "align", "indentLeft", "spacingBefore", "spacingAfter", "lineSpacing", "pageBreakBefore", "numPr"];

/** The paragraph's own properties, without its text or the revision records hanging off it. */
export function paraStyleOf(p: Paragraph): ParaStyle {
  const out: Record<string, unknown> = {};
  for (const key of PARA_STYLE_KEYS) if (p[key] !== undefined) out[key] = p[key];
  return out as ParaStyle;
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
 * and `w:eastAsianLayout` are all in the set and all outside the model's nine fields.
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
 * named child moves. `w:spacing` and `w:ind` are deliberately left alone — the fields spell one attribute of
 * each, and rewriting the element would take `w:firstLine`, `w:hanging` and `w:lineRule` down with it.
 *
 * A prefix is copied from the set rather than assumed: `w` is only an alias for the words namespace where the
 * part binds it, and a child written with a different prefix than its neighbours would read as another element.
 */
export function syncParaSlots(p: Paragraph): void {
  const slots = p.pPrSlots;
  if (!slots) return;
  const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
  const prefix = /^<([A-Za-z_][-\w.]*):/.exec(slots[0]?.xml ?? "")?.[1] ?? "w";
  const same = (a: unknown, b: unknown): boolean =>
    a === b || (typeof a === "object" && typeof b === "object" && JSON.stringify(a) === JSON.stringify(b));
  const wanted: PropSlot[] = [];
  // A property the paragraph spelled for itself follows the field, because a reviewer's decision or a toolbar
  // click may have moved it. One it only borrowed from its style is written only when it really moved.
  const keep = (tag: string, key: keyof ParaStyle, value: unknown, xml: string): void => {
    if (value === undefined || value === false) return;
    if (slots.some((s) => s.tag === tag) || !same(p.asRead?.[key], value)) wanted.push({ tag, xml });
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
  const owned = new Set(["pStyle", "pageBreakBefore", "numPr", "jc"]);
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
  if (rPr.u !== undefined && rPr.u["@_val"] !== "none") s.underline = true;
  if (rPr.strike !== undefined) s.strike = true;
  if (rPr.sz) s.sizePt = Number(rPr.sz["@_val"]) / 2;
  const color = rPr.color?.["@_val"];
  if (color && color !== "auto") s.color = hex(color);
  if (rPr.highlight?.["@_val"]) s.highlight = rPr.highlight["@_val"];
  if (rPr.rFonts) s.font = rPr.rFonts["@_ascii"];
  const va = rPr.vertAlign?.["@_val"];
  if (va === "superscript" || va === "subscript") s.vertAlign = va;
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

function section(sectPr: any): Section {
  const pgSz = sectPr.pgSz ?? {};
  const mar = sectPr.pgMar ?? {};
  return {
    pageWidthTwips: Number(pgSz["@_w"] ?? 11906),
    pageHeightTwips: Number(pgSz["@_h"] ?? 16838),
    margins: {
      top: Number(mar["@_top"] ?? 1440),
      right: Number(mar["@_right"] ?? 1440),
      bottom: Number(mar["@_bottom"] ?? 1440),
      left: Number(mar["@_left"] ?? 1440),
    },
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
