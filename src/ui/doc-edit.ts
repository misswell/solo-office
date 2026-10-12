import type { Block, Carried, FormatRevision, Inline, ParaStyle, Paragraph, ParagraphFormatRevision, PropSlot, Revision, RevisionEntry, RowInfo, RunStyle, Table, TableCell, WordDocument } from "../core/docx.js";
import type { CommentAnchor, DocComment } from "../core/comments.js";
import { bodyShape, commentTexts } from "../core/comments.js";
import { RUN_STYLE_KEYS, paraStyleOf, runStyleOf, stampNow, syncParaSlots, twipsToPx } from "../core/docx.js";
import type { CellWhere, TableEdit } from "../core/doc-table-ops.js";
import { cellRangeFor, cellRect, cellsInRect, cellsInSweep, deleteColumnHere, deleteRowHere, gridColumns, insertCellHere, insertColumnHere, insertRowHere, mergeHere, pasteAppendTableHere, pasteAsNestedTableHere, splitHere, splitTableAt } from "../core/doc-table-ops.js";
import { IMAGE_EXT_BY_MIME } from "../core/media.js";
import { CARRY_CLASS, CARRY_STYLE, RUN_CARRY_CLASS, highlightColor, renderWord } from "../core/render.js";
import { CommentPanel } from "./doc-comments.js";

/**
 * Word-like editing over the whole page: one `contenteditable` owns the body, so
 * Enter, Backspace, arrows and table Tab stay the browser's own behavior and the
 * model is re-derived from the DOM after every mutation. Blocks carry a stamped
 * id, so a paragraph that merely moved keeps its style, numbering and spacing.
 *
 * Styling a non-empty selection edits the model directly instead of the DOM —
 * a range that spans several paragraphs then costs nothing special.
 */

/** What one structure command came back with, including which table it ran on and how the page names it. */
type CommandOutcome = { ok: true; at: CellWhere; prefix: number[]; table: Table } | { ok: false; reason: string };

const HIGHLIGHT_NAMES = ["yellow", "green", "cyan", "magenta", "red", "darkYellow", "lightGray"];

/** Chromium reports colors as `rgb(...)`; routing both sides through one element compares them. */
const probe = document.createElement("span");
function cssValue(property: "color" | "background", value: string): string {
  probe.style.color = "";
  probe.style.background = "";
  if (property === "color") probe.style.color = value;
  else probe.style.background = value;
  return (property === "color" ? probe.style.color : probe.style.background) || value;
}

const HIGHLIGHT_BY_CSS = new Map<string, string>(HIGHLIGHT_NAMES.map((n) => [cssValue("background", highlightColor(n)), n]));

function hexOf(css: string): string | undefined {
  const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(css);
  if (!rgb) return undefined;
  return "#" + rgb.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, "0")).join("").toUpperCase();
}

const boldWeight = (weight: string): boolean | undefined => {
  if (!weight) return undefined;
  if (weight === "bold" || weight === "bolder") return true;
  const n = Number.parseFloat(weight);
  return Number.isFinite(n) ? n >= 600 : undefined;
};

/** The run style a DOM element contributes on top of its ancestors'. */
function styleOf(el: Element): RunStyle {
  const tag = el.tagName.toLowerCase();
  const inline = (el as HTMLElement).style;
  const s: RunStyle = {};
  if (tag === "b" || tag === "strong") s.bold = true;
  if (tag === "i" || tag === "em") s.italic = true;
  if (tag === "u") s.underline = true;
  if (tag === "s" || tag === "strike" || tag === "del") s.strike = true;
  if (tag === "sup") s.vertAlign = "superscript";
  if (tag === "sub") s.vertAlign = "subscript";
  const byWeight = boldWeight(inline.fontWeight);
  if (byWeight !== undefined) s.bold = byWeight;
  if (inline.fontStyle === "italic" || inline.fontStyle === "oblique") s.italic = true;
  const deco = `${inline.textDecoration} ${inline.textDecorationLine}`;
  if (deco.includes("underline")) s.underline = true;
  if (deco.includes("line-through")) s.strike = true;
  if (inline.fontSize.endsWith("pt")) s.sizePt = Number.parseFloat(inline.fontSize);
  if (inline.color) s.color = hexOf(inline.color);
  const bg = inline.backgroundColor || inline.background;
  if (bg) s.highlight = HIGHLIGHT_BY_CSS.get(cssValue("background", bg));
  if (inline.verticalAlign === "super") s.vertAlign = "superscript";
  if (inline.verticalAlign === "sub") s.vertAlign = "subscript";
  return s;
}

/** The tracked change a marked span carries back to us — its signature, not a style. */
function revisionOf(el: Element): Revision | undefined {
  const kind = el.classList.contains("rev") ? el.getAttribute("data-rev") : null;
  if (kind !== "ins" && kind !== "del") return undefined;
  const move = moveOf(el, "data-rev-move");
  return {
    kind,
    id: Number(el.getAttribute("data-rev-id") ?? 0),
    author: el.getAttribute("data-rev-author") ?? "",
    date: el.getAttribute("data-rev-date") ?? undefined,
    ...(move ? { move } : {}),
  };
}

/**
 * The rows of this table only. A table nested in a cell is inside the same element, and its rows are not
 * this table's rows.
 */
function rowsOf(el: Element): Element[] {
  return [...el.children].flatMap((part) =>
    part.tagName === "TR" ? [part] : ["TBODY", "THEAD", "TFOOT"].includes(part.tagName) ? [...part.children].filter((tr) => tr.tagName === "TR") : [],
  );
}

/** The cells of this row only, spelled the way Word spells them: `td`, and `th` for a heading row. */
function cellsOf(tr: Element): Element[] {
  return [...tr.children].filter((td) => td.tagName === "TD" || td.tagName === "TH");
}

/**
 * The tracked change on a table row (`w:trPr/w:ins`, `w:trPr/w:del`). It is read off the row, not the
 * cells, and this model has no row object, so the caller stamps it on every cell of the row.
 */
function rowRevisionOf(el: Element): Revision | undefined {
  const kind = el.getAttribute("data-row-rev");
  if (kind !== "ins" && kind !== "del") return undefined;
  return {
    kind,
    id: Number(el.getAttribute("data-row-rev-id") ?? 0),
    author: el.getAttribute("data-row-rev-author") ?? "",
    date: el.getAttribute("data-row-rev-date") ?? undefined,
  };
}

/**
 * A row that is on the page but not in the index — a copy of a row the browser made, which brings the mark
 * along with the text. It has no height or flags of its own, only the signature it was copied with.
 */
function revisionRowOf(el: Element): RowInfo | undefined {
  const rev = rowRevisionOf(el);
  return rev ? { rev } : undefined;
}

/**
 * Which stamped ids this rebuild has already settled on one model object for. The browser duplicates a row
 * when you copy it inside a table, and the duplicate arrives wearing the original's id.
 */
type ClaimSets = { blocks: Set<number>; rows: Set<number>; cells: Set<number> };
const emptyClaims = (): ClaimSets => ({ blocks: new Set(), rows: new Set(), cells: new Set() });

/**
 * Which half of a tracked move the file signed this record as. The editor never makes a move, but a file
 * that came in with one has to keep the tag through whatever the reviewer does next, or the moved passage
 * saves back as an unrelated deletion and insertion.
 */
const moveOf = (el: Element, attr: string): "from" | "to" | undefined => {
  const move = el.getAttribute(attr);
  return move === "from" || move === "to" ? move : undefined;
};

/** The formatting revision a marked span carries back to us: its signature and the properties from before. */
function formatChangeOf(el: Element): FormatRevision | undefined {
  const raw = el.classList.contains("rev-fmt") ? el.getAttribute("data-rev-format") : null;
  if (raw === null) return undefined;
  const encoded = el.getAttribute("data-rev-prev");
  let prev: RunStyle;
  try {
    prev = JSON.parse(decodeURIComponent(encoded ?? "{}")) as RunStyle;
  } catch {
    // A record whose before-state cannot be read is not a change anybody can decide; the text it marks
    // is ordinary text either way, so dropping the record loses nothing but the phantom mark.
    return undefined;
  }
  return {
    kind: "format",
    id: Number(raw) || 0,
    author: el.getAttribute("data-rev-format-author") ?? "",
    date: el.getAttribute("data-rev-format-date") ?? undefined,
    prev,
  };
}

type Context = RunStyle & { hyperlink?: string; rev?: Revision; revFormat?: FormatRevision };

/**
 * The paragraph-level revisions a rendered paragraph carries back to us. A paragraph object is reused
 * when the page is re-read, so a mark the page no longer shows has to be cleared here as well as read —
 * otherwise a revision somebody decided would come back on the next keystroke.
 */
function readParagraphRevisions(el: Element, p: Paragraph): void {
  const kind = el.getAttribute("data-rev-mark");
  if (kind === "ins" || kind === "del") {
    const move = moveOf(el, "data-rev-mark-move");
    p.markRev = {
      kind,
      id: Number(el.getAttribute("data-rev-mark-id")) || 0,
      author: el.getAttribute("data-rev-mark-author") ?? "",
      date: el.getAttribute("data-rev-mark-date") ?? undefined,
      ...(move ? { move } : {}),
    };
  } else delete p.markRev;
  const raw = el.getAttribute("data-para-rev");
  if (raw === null) {
    delete p.revFormat;
    return;
  }
  try {
    p.revFormat = {
      kind: "format",
      id: Number(raw) || 0,
      author: el.getAttribute("data-para-rev-author") ?? "",
      date: el.getAttribute("data-para-rev-date") ?? undefined,
      prev: JSON.parse(decodeURIComponent(el.getAttribute("data-para-prev") ?? "{}")) as ParaStyle,
    };
  } catch {
    // A record whose properties cannot be read is not one anybody can decide; the paragraph's own text
    // is unaffected, so dropping it loses the phantom mark and nothing else.
    delete p.revFormat;
  }
}

const CARRIED: (keyof Inline)[] = ["bold", "italic", "underline", "strike", "sizePt", "color", "highlight", "font", "vertAlign", "hyperlink", "break", "brXml", "tab"];

/** Two runs join only if they share their revision, or an edit would rewrite someone's tracked change. */
const sameRevision = (a?: Revision, b?: Revision): boolean =>
  a?.kind === b?.kind && a?.id === b?.id && a?.author === b?.author && a?.date === b?.date;

const sameRunProps = (a: RunStyle, b: RunStyle): boolean => RUN_STYLE_KEYS.every((k) => a[k] === b[k]);

/** Two records join only if they are the same change, before-state included — the point of the record. */
const sameFormatChange = (a?: FormatRevision, b?: FormatRevision): boolean =>
  a?.id === b?.id &&
  a?.author === b?.author &&
  a?.date === b?.date &&
  (!a || !b || (sameRunProps(a.prev, b.prev) && sameRunSlots(a.prevSlots, b.prevSlots)));

/**
 * Two runs join only if they carry the same properties text. A run's `w:rPr` travels as the file wrote it, so
 * joining one that holds `w:kern` to one that does not would have to pick a winner — and a run cut in half by
 * `splitAt` keeps the one array, so splitting and rejoining a run costs nothing.
 */
const sameRunSlots = (a?: PropSlot[], b?: PropSlot[]): boolean =>
  a === b || (!!a === !!b && (!!a ? a.length === b!.length && a.every((s, i) => s.xml === b![i]?.xml) : true));

const sameStyle = (a: Inline, b: Inline): boolean =>
  CARRIED.every((k) => a[k] === b[k]) &&
  sameRunSlots(a.rPrSlots, b.rPrSlots) &&
  !!a.image === !!b.image &&
  a.image?.path === b.image?.path &&
  // An anchor is an element of its own: joining it to a neighbour would move or swallow the marker.
  a.anchor?.id === b.anchor?.id &&
  a.anchor?.kind === b.anchor?.kind &&
  // A carried structure is the same kind of element: it has no text of its own, and words joined into it
  // would be written back as the bytes it came in with, which is to say lost.
  a.carry?.key === b.carry?.key &&
  a.anchor?.kind === b.anchor?.kind &&
  sameRevision(a.rev, b.rev) &&
  sameFormatChange(a.revFormat, b.revFormat);

/** NBSP is how the renderer makes an empty paragraph take vertical space; it is not document text. */
const cleanText = (text: string): string => text.replace(/\u00a0/g, " ");

const inlineLength = (inline: Inline): number => (inline.text !== "" ? inline.text.length : 1);

const isMark = (inline: Inline): boolean =>
  !!inline.image || !!inline.break || !!inline.tab || !!inline.anchor || !!inline.carry;

/**
 * The body as one comparable string. A mark (image, break, tab, anchor) stands for one character,
 * which is how `splitAt` counts it too; NUL is the placeholder because document text never holds one.
 */
const flatOf = (inlines: Inline[]): string => inlines.map((i) => (i.text !== "" ? i.text : "\u0000")).join("");

const charLength = (inlines: Inline[]): number => inlines.reduce((n, i) => n + inlineLength(i), 0);

function merge(inlines: Inline[]): Inline[] {
  const out: Inline[] = [];
  for (const inline of inlines) {
    if (inline.text === "" && !isMark(inline)) continue;
    const last = out[out.length - 1];
    if (last && !isMark(inline) && sameStyle(last, inline)) last.text += inline.text;
    else out.push({ ...inline });
  }
  return out;
}

export interface DocPath {
  path: number[];
  offset: number;
}

/** What the four 选择 cells commands answer: whether the page can show such a range, and how many cells it lit up. */
export type TableSelection = { ok: boolean; reason?: string; selected?: number };

/** What the page shows about tracked changes: the marks, the text as if accepted, or as if refused. */
export type RevisionView = "marks" | "final" | "original";

/** The body as one character stream, plus the paragraph each span of that stream came from and the
 *  stretch of paragraphs it belongs to: neighbours inside one stretch can be joined or split, and a
 *  paragraph on the other side of a table wall never is. */
type BodyText = { runs: Inline[]; spans: { paragraph: Paragraph; length: number; group: number }[] };

export type StylePatch = Partial<Pick<RunStyle, "bold" | "italic" | "underline" | "strike" | "sizePt" | "color" | "highlight">>;

/** One row of the comment list, shaped for the panel rather than for the file format. */
export interface CommentEntry {
  id: number;
  /** The number the reference mark shows: position in the list, not the `w:id`. */
  ordinal: number;
  author: string;
  initials?: string;
  date?: string;
  paragraphs: string[];
  /** The same body as runs, so the panel can edit it without flattening its formatting. */
  runs: Inline[][];
  /** False when the body holds markup this model cannot rebuild; the text then reads only. */
  editable: boolean;
  /** False when no anchor for it survives in the body — the panel still lists it, greyed out. */
  anchored: boolean;
  /** `w15:commentEx/@done`: shown greyed out, the way WPS and Word keep a resolved thread. */
  resolved: boolean;
  /** The comment this one answers, if it is a reply. */
  parentId?: number;
  /** Ids answering this comment, in list order. */
  replies: number[];
}

const REVISION_STYLE_ID = "solo-doc-revision-styles";

/**
 * Revision marks are classed, not inline-styled, so reading the DOM back cannot mistake "marked
 * because inserted" for run formatting. The editor injects its own stylesheet because `index.html`
 * is not ours to touch.
 */
function injectRevisionStyles(): void {
  if (document.getElementById(REVISION_STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = REVISION_STYLE_ID;
  style.textContent = `
.page .rev-ins { text-decoration: underline; color: #1a6b1f; }
.page .rev-del { text-decoration: line-through; color: #b00020; opacity: .85; }
/* A formatting revision leaves the run's own properties showing and marks the range with a dotted
   underline of its own — a mark, so class-based and never an inline style the run would read back. */
.page .rev-fmt { text-decoration-line: underline; text-decoration-style: dotted; text-decoration-color: #6b4fa0; }
/* The paragraph mark is the break that ends a paragraph, so it gets Word's pilcrow at the end of the
   line: underlined when it was inserted, struck when it was deleted. It is drawn from a class, and a
   pseudo-element is not a text node, so the paragraph's characters are untouched by showing it. */
.page .rev-mark-ins::after, .page .rev-mark-del::after { content: "¶"; margin-left: 2px; font-weight: 400; }
.page .rev-mark-ins::after { text-decoration: underline; color: #1a6b1f; }
.page .rev-mark-del::after { text-decoration: line-through; color: #b00020; }
.page .para-rev { box-shadow: inset 2px 0 0 #6b4fa0; }
/* A row somebody added or removed keeps the row itself marked: the tr carries the same rev-ins /
   rev-del class as a span does, so the view modes hide it the way they hide marked text. Word draws a
   change bar in the margin for this; a bar on the row's first cell is our own cue, chosen because a
   class cannot be mistaken for cell formatting when the page is read back. */
.page tr.rev-row { color: #1a6b1f; }
.page tr.rev-row.rev-del { color: #b00020; }
.page tr.rev-row > td:first-child { box-shadow: inset 3px 0 0 currentColor; }
/* A row box draws no text, so a decoration set on the tr never reaches the words in its cells — Chromium
   computes none of it. The row's paragraphs carry the mark instead. */
.page tr.rev-row.rev-del .para { text-decoration: line-through; color: #b00020; }
.page tr.rev-row.rev-ins .para { text-decoration: underline; color: #1a6b1f; }
/* The flash a pane's row gets when it is clicked or stepped to: drawn on the marks the change leaves
   in the page, and gone with the next render, which rebuilds them from the model. */
.page .solo-rev-hit { background: rgba(255, 214, 102, .55); }
/* 选中整行/整列/整格/整个表格 lit-up cells. An editable page keeps only one Range, so a column — whose cells are not
   neighbours — cannot be lit by the browser's own selection at all: these cells carry the mark instead, and the
   outline is what still shows on a cell whose shading the page draws with an inline style. */
.page td.solo-cell-sel { background: rgba(58, 111, 244, .16); outline: 2px solid rgba(58, 111, 244, .55); outline-offset: -2px; }
/* The view only hides marks; the runs stay in the page, so a save writes what was read either way. */
.page.rev-final .rev-del { display: none; }
.page.rev-final .rev-ins { text-decoration: none; color: inherit; }
.page.rev-original .rev-ins { display: none; }
.page.rev-original .rev-del { text-decoration: none; color: inherit; }
.page.rev-final .rev-fmt, .page.rev-original .rev-fmt { text-decoration-line: none; }
.page.rev-final .rev-mark-ins::after, .page.rev-final .rev-mark-del::after,
.page.rev-original .rev-mark-ins::after, .page.rev-original .rev-mark-del::after { content: none; }
.page.rev-final .para-rev, .page.rev-original .para-rev { box-shadow: none; }
.page.rev-final tr.rev-row.rev-ins .para, .page.rev-original tr.rev-row.rev-del .para { text-decoration: none; color: inherit; }
.page.rev-final tr.rev-row.rev-ins > td:first-child, .page.rev-original tr.rev-row.rev-del > td:first-child { box-shadow: none; }
`;
  document.head.append(style);
}

export class DocEditor {
  private page: HTMLElement | null = null;
  private owned = new Map<number, Block>();
  /** Rows and cells are stamped the same way blocks are: the file's own properties live on them. */
  private ownedRows = new Map<number, RowInfo>();
  private ownedCells = new Map<number, TableCell>();
  /** Ids this rebuild has already handed out, so a stamped element the browser duplicated is a copy. */
  private claimed: ClaimSets = emptyClaims();
  /**
   * The carried inline structures of the whole model by their key, and which of them this rebuild has placed.
   * An inline is not stamped the way a block is, so the page's `data-carry` is the only thing tying a
   * bookmark's span back to the bytes it stands for — and one key must land in exactly one paragraph.
   */
  private carries = new Map<number, Inline>();
  private placedCarries = new Set<number>();
  private nextId = 1;
  private undoStack: { blocks: Block[]; comments: DocComment[]; caret: DocPath | null }[] = [];
  private redoStack: { blocks: Block[]; comments: DocComment[]; caret: DocPath | null }[] = [];
  private burst = 0;
  private composing = false;
  private panel: CommentPanel | null = null;
  /**
   * Comment ids the body anchored at or after the last render. Only these may be pruned when
   * their markers go away: a comment anchored in a header or footer has no body anchor to find.
   */
  private bodyAnchored = new Set<number>();
  /** The comment last revealed: `stepComment` counts from it, so next/previous walk the list. */
  private activeComment?: number;
  /** The change last revealed, by `w:id`: `stepRevision` counts from it, so 下一条/上一条 walk the list. */
  private activeRevision?: number;
  /** The last selection that was inside the page, so a toolbar click can still reach it. */
  private pageSelection: { start: DocPath; end: DocPath } | null = null;
  /**
   * The cells 选中整行/整列/整格/整个表格 lit up. A column is not one Range's worth of text and an editable page holds only
   * one range, so what the browser reports back is narrower than what the command picked: while the selection
   * stays inside this set, the rectangle the commands act on is the one the command computed, not the one the
   * page can show. Leaving the set — a click or a drag anywhere else — hands the selection back to the page.
   */
  private paintedCells: { prefix: string; cells: Set<string> } | null = null;
  private authorName = "Solo Office";
  private authorInitials = "S";
  private revisionView: RevisionView = "marks";
  /** 修订: edits made while this is on are recorded as tracked changes rather than applied. */
  private trackRevisions = false;
  /** The next `w:id` to hand out, so a revision we create cannot collide with one read from the file. */
  private revSeq = 0;
  /**
   * 复制表格: the table the two paste commands below put back when the caller does not name a source. This is
   * the editor's own copy, not the system clipboard — the system one holds text, and reading a table *out of*
   * pasted HTML is a separate job nobody has taken on yet, so `copyTable()` hands the shell the table to carry
   * between documents and the paste commands take it back as an argument. Kept after a paste, the way a
   * clipboard is: the same table can go in twice.
   */
  private clipTable?: Table;

  constructor(
    private readonly host: HTMLElement,
    private readonly doc: WordDocument,
    private readonly hooks: { changed(): void },
  ) {
    for (const id of anchorIdsIn(this.doc.blocks)) this.bodyAnchored.add(id);
    this.revSeq = this.doc.maxRevisionId();
    injectRevisionStyles();
    this.panel = new CommentPanel(this.host, this);
    document.addEventListener("selectionchange", () => this.trackSelection());
    this.render();
  }

  render(): void {
    this.host.innerHTML = renderWord(this.doc);
    this.page = this.host.querySelector<HTMLElement>(".page");
    this.panel?.mount();
    if (!this.page) return;
    this.page.contentEditable = "true";
    this.page.spellcheck = false;
    if (this.zoomLevel !== 1) this.page.style.zoom = String(this.zoomLevel);
    this.reindex();
    this.imagePaths();
    this.wireHF();
    this.wire();
    this.applyRevisionView();
  }

  private hfEditing: HTMLElement | null = null;

  private wireHF(): void {
    for (const el of this.host.querySelectorAll<HTMLElement>(".hf")) {
      el.addEventListener("click", () => this.editHF(el));
    }
  }

  /** Click a header/footer area to edit it; blur commits the paragraph texts. */
  private editHF(el: HTMLElement): void {
    if (el.isContentEditable) return;
    const which = el.dataset.hf === "footer" ? "footer" : "header";
    const hf = which === "header" ? this.doc.header : this.doc.footer;
    if (!hf?.editable) return;
    el.contentEditable = "true";
    el.spellcheck = false;
    this.hfEditing = el;
    el.focus();
    el.addEventListener(
      "blur",
      () => {
        el.contentEditable = "false";
        if (this.hfEditing === el) this.hfEditing = null;
        const paragraphs = [...el.querySelectorAll(".hf-p")].map((p) => (p.textContent ?? "").replace(/\u00a0/g, " "));
        const before = which === "header" ? this.doc.header : this.doc.footer;
        if (!before || JSON.stringify(before.paragraphs) === JSON.stringify(paragraphs)) return;
        this.doc.setHeaderTexts(which, paragraphs);
        this.hooks.changed();
      },
      { once: true },
    );
  }

  /** Opens the header/footer for editing, creating one when the document had none. */
  openHF(which: "header" | "footer"): void {
    const hf = which === "header" ? this.doc.header : this.doc.footer;
    if (!hf) {
      const made = this.doc.ensureHeaderFooter(which);
      if (!made) return;
      this.render();
    }
    const el = this.host.querySelector<HTMLElement>(`.hf-${which}`);
    if (el) {
      this.editHF(el);
      el.focus();
    }
  }

  private zoomLevel = 1;

  /** Ribbon zoom: the page scales, the chrome stays put. */
  setZoom(z: number): void {
    this.zoomLevel = z;
    this.page?.style.setProperty("zoom", String(z));
  }

  /** Ids are re-stamped per render because the model and the DOM are built in the same pass. */
  private reindex(): void {
    if (!this.page) return;
    this.owned.clear();
    this.ownedRows.clear();
    this.ownedCells.clear();
    this.nextId = 1;
    this.indexBlocks(this.page, this.doc.blocks);
  }

  /**
   * Stamps a block list against the elements that show it, so a block keeps its own model object when the
   * browser moves or splits the element around it. A table is walked all the way down: the row and the
   * cell carry properties the file wrote that no paragraph can restate, so each needs its own id to be
   * found again by the next pass.
   */
  private indexBlocks(container: Element, blocks: Block[]): void {
    [...container.children].forEach((el, i) => {
      const block = blocks[i];
      if (!block) return;
      if (block.kind === "carried") {
        // A container has no text to re-derive and nothing to index below it — it is stamped so the next
        // pass finds the same one, and only while the element is still the container's own.
        if (el.classList.contains(CARRY_CLASS)) this.owned.set(this.stamp(el), block);
        return;
      }
      if (block.kind !== "table" || el.tagName !== "TABLE") {
        if (block.kind === "paragraph") this.owned.set(this.stamp(el), block);
        return;
      }
      this.owned.set(this.stamp(el), block);
      rowsOf(el).forEach((tr, r) => {
        const row = block.rows[r];
        if (!row) return;
        // One record for the row: the cells of a row share the object, so the row stays one row however
        // its cells are moved about.
        if (row[0]?.row) this.ownedRows.set(this.stamp(tr, "rid"), row[0].row);
        cellsOf(tr).forEach((td, c) => {
          const cell = row[c];
          if (!cell) return;
          this.ownedCells.set(this.stamp(td, "cid"), cell);
          this.indexBlocks(td, cell.paragraphs);
        });
      });
    });
  }

  private stamp(el: Element, key: "bid" | "rid" | "cid" = "bid"): number {
    const id = this.nextId++;
    (el as HTMLElement).dataset[key] = String(id);
    return id;
  }

  /** The rendered `<img>` carries no part name, and the writer cannot put its bytes back without it. */
  private imagePaths(): void {
    if (!this.page) return;
    const paths: string[] = [];
    const collect = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind === "table") b.rows.forEach((row) => row.forEach((c) => collect(c.paragraphs)));
        else if (b.kind === "paragraph") for (const inline of b.inlines) if (inline.image) paths.push(inline.image.path);
      }
    };
    collect(this.doc.blocks);
    [...this.page.querySelectorAll("img")].forEach((img, i) => {
      if (paths[i]) img.dataset.path = paths[i];
    });
  }

  private wire(): void {
    const page = this.page;
    if (!page) return;
    page.addEventListener("compositionstart", () => {
      this.composing = true;
      this.snapshot(this.caretPath());
    });
    page.addEventListener("compositionend", () => {
      this.composing = false;
      this.mutated();
    });
    page.addEventListener("beforeinput", (e) => {
      if (this.composing) return;
      const type = e.inputType;
      if (type === "insertText" || type === "insertFromPaste" || type === "insertFromDrop") this.mark();
      else if (type.startsWith("delete") || type === "insertParagraph" || type === "insertLineBreak") this.snapshot(this.caretPath());
    });
    page.addEventListener("input", (e) => {
      if (this.composing || (e as InputEvent).isComposing) return;
      this.mutated();
    });
    page.addEventListener("paste", (e) => {
      // Pasted HTML brings foreign styles and whole page layouts; a Word body wants the text.
      const text = (e as ClipboardEvent).clipboardData?.getData("text/plain");
      e.preventDefault();
      if (!text) return;
      this.snapshot(this.caretPath());
      document.execCommand("insertText", false, text);
    });
    page.addEventListener("keydown", (e) => {
      if (e.isComposing || this.composing || e.metaKey || e.ctrlKey || e.altKey) return;
      if (e.key !== "Tab") return;
      e.preventDefault();
      this.snapshot(this.caretPath());
      document.execCommand("insertHTML", false, '<span class="tab"></span>');
      this.mutated();
    });
    // Clicking a link should place the caret like any other text; ⌘-click still follows it.
    page.addEventListener("click", (e) => {
      const mark = closestClass(e.target as Node, "cmt-anchor");
      if (mark) {
        const id = Number(mark.dataset.cid);
        if (Number.isFinite(id)) this.panel?.openAt(id);
        return;
      }
      if (e.metaKey || e.ctrlKey) return;
      if (closestTag(e.target as Node, "A")) e.preventDefault();
    });
    // A click inside anchored text names the comment that holds it.
    for (const el of page.querySelectorAll<HTMLElement>(".cmt-in")) {
      el.addEventListener("click", () => {
        const id = Number(el.dataset.cids?.split(",")[0]);
        if (Number.isFinite(id)) this.panel?.openAt(id);
      });
    }
  }

  /** Remembered because a ribbon or panel click moves the focus out of the page. */
  private trackSelection(): void {
    const ends = this.rangeEnds();
    if (!ends) return;
    const painted = this.paintedCells;
    // 命令点亮的那一组格子还没被用户动过：页面里的那一条选区不能把命令算出的矩形缩回一个格。
    if (painted && this.picked(painted, ends.start) && this.picked(painted, ends.end)) return;
    this.paintedCells = null;
    this.clearPaint();
    this.pageSelection = ends;
  }

  /** Whether a path lands on one of the cells a selection command lit up. */
  private picked(painted: { prefix: string; cells: Set<string> }, at: DocPath): boolean {
    const cell = this.tableAt(at);
    return !!cell && cell.prefix.join() === painted.prefix && painted.cells.has(`${cell.where.row},${cell.where.col}`);
  }

  /** Re-derives the model from the DOM, then repairs anything structural the browser changed. */
  private mutated(): void {
    if (!this.page) return;
    const before = shape(this.doc.blocks);
    const anchorsBefore = countAnchors(this.doc.blocks);
    const caret = this.caretPath();
    const recorded = this.sync();
    // A bullet marker the browser deleted under Backspace has to come back from the model.
    const markers = markersOf(this.doc.blocks);
    const drawn = this.page.querySelectorAll(".bullet").length;
    // Losing an anchor means losing its comment: the body text it explained has nowhere to hang.
    const orphaned = this.pruneOrphans() > 0;
    if (before !== shape(this.doc.blocks) || recorded || markers !== drawn || orphaned || anchorsBefore !== countAnchors(this.doc.blocks)) {
      this.render();
      if (caret) this.setCaret(this.clamp(caret));
    }
    this.hooks.changed();
  }

  /**
   * Drops comments whose markers the reader deleted, the way Word does. Only ids the body
   * anchored are candidates, so a header/footer comment this page never shows survives.
   */
  private pruneOrphans(): number {
    const anchored = anchorIdsIn(this.doc.blocks);
    let removed = 0;
    for (const id of [...this.bodyAnchored]) {
      if (anchored.has(id)) continue;
      if (this.doc.removeComment(id)) removed += 1;
      this.bodyAnchored.delete(id);
    }
    return removed;
  }

  /**
   * Re-derives the model from the DOM. Returns whether tracking recorded anything: marks the page has
   * not shown yet are a change the caller has to render.
   */
  private sync(): boolean {
    if (!this.page) return false;
    wrapStrays(this.page);
    this.claimed = emptyClaims();
    this.carries = carryByKeys(this.doc.blocks);
    this.placedCarries = new Set();
    const before = this.trackRevisions ? this.bodyText() : undefined;
    const held = this.doc.blocks;
    const next: Block[] = [];
    const shown: Element[] = [];
    let previous: Paragraph | undefined;
    for (const el of [...this.page.children]) {
      if (el.tagName === "TABLE") {
        const known = this.ownedAt(el);
        next.push(this.tableOf(el as HTMLTableElement, known?.kind === "table" ? known : undefined));
        shown.push(el);
        previous = undefined;
        continue;
      }
      const known = this.ownedAt(el);
      if (el.classList.contains(CARRY_CLASS)) {
        // A container's element is not text to read: it either still stands for the block that was stamped
        // onto it, or it is a leftover the browser made and goes, so the model and the page keep one list.
        if (known?.kind === "carried") {
          next.push(known);
          shown.push(el);
        } else el.remove();
        previous = undefined;
        continue;
      }
      const paragraph: Paragraph = known?.kind === "paragraph" ? known : fresh(previous);
      paragraph.inlines = this.inlinesOf(el, paragraph.inlines);
      readParagraphRevisions(el, paragraph);
      next.push(paragraph);
      shown.push(el);
      previous = paragraph;
    }
    keepCarried(held, next, shown, (el, anchor) => this.page?.insertBefore(el, anchor ?? null));
    this.doc.blocks = next;
    const marked = before === undefined ? false : this.recordRevisions(before).marked;
    this.reindex();
    return marked;
  }

  /**
   * Rebuilds a table from the DOM. A row or cell that was already on the page keeps its own model object —
   * the height, the borders, the property set the file wrote and the model only carries — so a keystroke
   * inside one cell cannot flatten the grid around it. A row or cell the browser made has no such history
   * and starts empty, which is what a new one is.
   */
  private tableOf(el: HTMLTableElement, known: Table | undefined): Table {
    const table: Table = known ?? { kind: "table", columns: 1, rows: [] };
    table.rows = rowsOf(el).map((tr) => {
      let previous: Paragraph | undefined;
      // One record for the row, stamped on each of its cells — the same object, so the row reads as one
      // change and a cell that moves out of it cannot carry a signature that is not its own.
      const row = this.ownedRow(tr) ?? revisionRowOf(tr);
      return cellsOf(tr).map((td) => {
        wrapStrays(td);
        const cell: TableCell = this.ownedCell(td) ?? { paragraphs: [] };
        const held = cell.paragraphs;
        if (row) cell.row = row;
        else delete cell.row;
        const paragraphs: Block[] = [];
        const shown: Element[] = [];
        for (const p of [...td.children]) {
          if (p.tagName === "TABLE") {
            const inner = this.ownedAt(p);
            paragraphs.push(this.tableOf(p as HTMLTableElement, inner?.kind === "table" ? inner : undefined));
            shown.push(p);
            previous = undefined;
            continue;
          }
          const known = this.ownedAt(p);
          if (p.classList.contains(CARRY_CLASS)) {
            if (known?.kind === "carried") {
              paragraphs.push(known);
              shown.push(p);
            } else p.remove();
            previous = undefined;
            continue;
          }
          const paragraph: Paragraph = known?.kind === "paragraph" ? known : fresh(previous);
          paragraph.inlines = this.inlinesOf(p, paragraph.inlines);
          readParagraphRevisions(p, paragraph);
          paragraphs.push(paragraph);
          shown.push(p);
          previous = paragraph;
        }
        keepCarried(held, paragraphs, shown, (el, anchor) => td.insertBefore(el, anchor ?? null));
        cell.paragraphs = paragraphs;
        return cell;
      });
    });
    table.columns = Math.max(1, ...table.rows.map((r) => r.length));
    return table;
  }

  private ownedAt(el: Element): Block | undefined {
    return this.claim(this.owned, Number((el as HTMLElement).dataset.bid), this.claimed.blocks);
  }

  private ownedRow(el: Element): RowInfo | undefined {
    return this.claim(this.ownedRows, Number((el as HTMLElement).dataset.rid), this.claimed.rows);
  }

  private ownedCell(el: Element): TableCell | undefined {
    return this.claim(this.ownedCells, Number((el as HTMLElement).dataset.cid), this.claimed.cells);
  }

  /**
   * The record an element on the page was drawn from, found by the id stamped on it. A page can show the
   * same id twice — the browser copies a row when you duplicate one inside a table — and the second one is
   * a copy of the record, not the record: it keeps the width, the borders and the row's signature the copy
   * is drawn with, and editing it cannot rewrite the row or cell it was copied from.
   */
  private claim<T extends object>(pool: Map<number, T>, id: number, used: Set<number>): T | undefined {
    const known = pool.get(id);
    if (!known) return undefined;
    if (used.has(id)) {
      // The second element is one the browser made, and a thing the editor made has no identity to claim:
      // `w14:paraId` names exactly one paragraph in the whole document, so a copy that kept the number it
      // was cloned with would leave the file with two paragraphs sharing one and Word offering to repair it.
      // Widths, borders and property sets do travel with the copy — those are shapes, and shapes may repeat.
      const copy = { ...known } as T & { attrsXml?: string };
      delete copy.attrsXml;
      return copy;
    }
    used.add(id);
    return known;
  }

  private inlinesOf(el: Element, prev: Inline[] = []): Inline[] {
    const out: Inline[] = [];
    const walk = (node: Node, ctx: Context): void => {
      for (const child of node.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) {
          out.push({ ...ctx, text: cleanText(child.nodeValue ?? "") });
          continue;
        }
        if (!(child instanceof Element) || child.classList.contains("bullet")) continue;
        if (child.classList.contains("tab")) {
          out.push({ ...ctx, text: "", tab: true });
          continue;
        }
        // Read before descending: a reference badge shows a number, and that digit is not body text.
        if (child.classList.contains("cmt-anchor")) {
          const anchor = anchorOf(child);
          if (anchor) out.push({ text: "", anchor });
          continue;
        }
        // A structure the model only carries comes back by its key, not by its text: what it stands for is
        // bytes, and the paragraph it belongs to is the paragraph its element ended up in.
        if (child.classList.contains(RUN_CARRY_CLASS)) {
          const key = Number(child.getAttribute("data-carry"));
          const held = Number.isFinite(key) ? this.carries.get(key) : undefined;
          if (held && !this.placedCarries.has(key)) {
            this.placedCarries.add(key);
            out.push(held);
          } else child.remove();
          continue;
        }
        const tag = child.tagName.toLowerCase();
        if (tag === "br") {
          out.push({ ...ctx, text: "", break: true });
          continue;
        }
        if (tag === "img") {
          const image = imageOf(child);
          if (image) out.push({ ...ctx, text: "", image });
          continue;
        }
        const next: Context = { ...ctx, ...styleOf(child) };
        if (tag === "a") next.hyperlink = child.getAttribute("href") || undefined;
        const rev = revisionOf(child);
        if (rev) next.rev = rev;
        const fmt = formatChangeOf(child);
        if (fmt) next.revFormat = fmt;
        walk(child, next);
      }
    };
    walk(el, {});
    const merged = merge(out);
    keepCarries(prev, merged, this.placedCarries);
    return merged.every((i) => i.text.trim() === "" && !isMark(i)) ? [] : merged;
  }

  // --------------------------------------------------------------- caret

  /**
   * The block element that holds `node`, as a path into the model: the top-level block's index, then one
   * (row, cell, block) triple for every table the path steps through. A caret in a cell of a table nested
   * inside another table's cell therefore reads `[table, row, cell, block, row2, cell2, block2]` — the same
   * shape the model has, so no command can mistake the inner table for the outer one.
   */
  private pathOf(node: Node | null, offset: number): DocPath | null {
    if (!this.page || !node) return null;
    let top: Node | null = node;
    while (top && top.parentElement !== this.page) top = top.parentElement;
    if (!top) return null;
    const index = [...this.page.children].indexOf(top as Element);
    if (index < 0) return null;
    const el = top as Element;
    if (el.tagName !== "TABLE") return { path: [index], offset: modelOffset(el, node, offset) };
    const cells: number[][] = [];
    let holder: Element | null = null;
    let below: Element | null = node instanceof Element ? node : node.parentElement;
    for (let cur = below; cur && cur !== this.page; cur = cur.parentElement) {
      if (below && (cur.tagName === "TD" || cur.tagName === "TH")) {
        const tr = cur.parentElement;
        // The row's own table, found through the tbody the page draws it in — `closest` stops at the innermost
        // one, which is what a cell of a nested table belongs to.
        const table = cur.closest("table");
        const block = [...cur.children].indexOf(below);
        const row = tr && table ? rowsOf(table).indexOf(tr) : -1;
        const col = tr ? cellsOf(tr).indexOf(cur) : -1;
        if (block < 0 || row < 0 || col < 0) return null;
        cells.push([row, col, block]);
        holder ??= below;
        below = cur;
        continue;
      }
      below = cur;
    }
    if (!cells.length || !holder) return null;
    return { path: [index, ...cells.reverse().flat()], offset: modelOffset(holder, node, offset) };
  }

  caretPath(): DocPath | null {
    const sel = document.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    return this.pathOf(sel.anchorNode, sel.anchorOffset);
  }

  private rangeEnds(): { start: DocPath; end: DocPath } | null {
    const sel = document.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    const range = sel.getRangeAt(0);
    const start = this.pathOf(range.startContainer, range.startOffset);
    const end = this.pathOf(range.endContainer, range.endOffset);
    return start && end ? { start, end } : null;
  }

  setCaret(caret: DocPath): void {
    const el = this.elementAt(caret.path);
    if (!el) return;
    const spots = positions(el);
    const spot = spots[Math.min(caret.offset, spots.length - 1)];
    if (!spot) return;
    const range = document.createRange();
    range.setStart(spot.node, spot.offset);
    range.collapse(true);
    this.select(range, el);
  }

  private select(range: Range, focus: Element): void {
    const sel = document.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    if (focus instanceof HTMLElement) focus.focus();
  }

  /** The element the page draws for a path, walked table by table the way the model is walked. */
  private elementAt(path: number[]): Element | null {
    if (!this.page) return null;
    let el: Element | undefined = this.page.children[path[0]];
    for (let k = 1; el && k < path.length; k += 3) {
      const row: Element | undefined = rowsOf(el)[path[k]];
      el = row ? cellsOf(row)[path[k + 1]]?.children[path[k + 2]] : undefined;
    }
    return el ?? null;
  }

  /** A caret inside content the user just deleted lands on the nearest block that still exists. */
  private clamp(caret: DocPath): DocPath {
    const top = caret.path[0];
    if (top < this.doc.blocks.length) return caret;
    const last = this.doc.blocks.length - 1;
    return last < 0 ? { path: [0], offset: 0 } : { path: [last], offset: 0 };
  }

  // --------------------------------------------------------- undo/redo

  /** Typing coalesces into one entry per pause; everything else is its own step. */
  private mark(): void {
    const now = Date.now();
    if (now - this.burst < 900) {
      this.burst = now;
      return;
    }
    this.burst = now;
    this.snapshot(this.caretPath());
  }

  private snapshot(at: DocPath | null): void {
    this.burst = 0;
    this.undoStack.push({ blocks: structuredClone(this.doc.blocks), comments: structuredClone(this.doc.comments), caret: at });
    if (this.undoStack.length > 80) this.undoStack.shift();
    this.redoStack.length = 0;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undo(): boolean {
    const step = this.undoStack.pop();
    if (!step) return false;
    this.redoStack.push({ blocks: structuredClone(this.doc.blocks), comments: structuredClone(this.doc.comments), caret: this.caretPath() });
    this.doc.blocks = step.blocks;
    this.doc.comments = step.comments;
    this.render();
    if (step.caret) this.setCaret(this.clamp(step.caret));
    this.hooks.changed();
    return true;
  }

  redo(): boolean {
    const step = this.redoStack.pop();
    if (!step) return false;
    this.undoStack.push({ blocks: structuredClone(this.doc.blocks), comments: structuredClone(this.doc.comments), caret: this.caretPath() });
    this.doc.blocks = step.blocks;
    this.doc.comments = step.comments;
    this.render();
    if (step.caret) this.setCaret(this.clamp(step.caret));
    this.hooks.changed();
    return true;
  }

  // -------------------------------------------------------------- inserting

  /**
   * A plain grid at the caret, Word-default column widths (9026 usable twips
   * between A4 margins), one empty paragraph per cell. The caret lands in the
   * first cell, where Word puts it too. With the caret inside a cell the grid goes into that cell — a table
   * in a table — because that is the cell the user was working in, and a body-level insert would move the
   * new table out from under them.
   */
  insertTable(rows: number, cols: number): void {
    const caret = this.caretPath();
    const cell = this.tableAt(caret);
    const width = Math.floor(9026 / cols);
    const table: Table = {
      kind: "table",
      columns: cols,
      rows: Array.from({ length: rows }, () =>
        Array.from({ length: cols }, (): TableCell => ({ paragraphs: [{ kind: "paragraph", inlines: [] }], widthTwips: width })),
      ),
    };
    this.snapshot(caret);
    if (caret && cell) {
      const blocks = cell.table.rows[cell.where.row]?.[cell.where.col]?.paragraphs;
      if (blocks) {
        const at = Math.min(caret.path[caret.path.length - 1] + 1, blocks.length);
        blocks.splice(at, 0, table);
        this.render();
        this.setCaret({ path: [...cell.prefix, cell.where.row, cell.where.col, at, 0, 0, 0], offset: 0 });
        this.hooks.changed();
        return;
      }
    }
    const at = caret?.path[0] ?? this.doc.blocks.length;
    this.doc.blocks.splice(at, 0, table);
    this.render();
    this.setCaret({ path: [at, 0, 0, 0], offset: 0 });
    this.hooks.changed();
  }

  // -------------------------------------------------------- table structure

  /**
   * The block a path names: one top-level index, then a (row, cell, block) triple for each table the path
   * steps into. Resolution has the same shape as the model, so a nested table is reached through the cell
   * that holds it and never mistaken for the table around it.
   */
  private blockAt(path: number[]): Block | undefined {
    let block: Block | undefined = this.doc.blocks[path[0]];
    for (let k = 1; block && k < path.length; k += 3) {
      if (block.kind !== "table") return undefined;
      block = block.rows[path[k]]?.[path[k + 1]]?.paragraphs[path[k + 2]];
    }
    return block;
  }

  /**
   * The list a table sits in: the body's blocks, or the blocks of the cell that holds it. `prefix` names the
   * block inside that list, so a command that adds or removes a whole table has to reach the list too — and
   * for a table nested in a cell, the new table and the paragraph beside it belong to the cell, not to the
   * page around it.
   */
  private containerBlocks(prefix: number[]): Block[] | undefined {
    if (prefix.length === 1) return this.doc.blocks;
    const holder = this.blockAt(prefix.slice(0, prefix.length - 3));
    if (holder?.kind !== "table") return undefined;
    return holder.rows[prefix[prefix.length - 3]]?.[prefix[prefix.length - 2]]?.paragraphs;
  }

  /**
   * The table the caret sits in, with the cell it sits in. A path only reaches a cell when it names one, so a
   * caret in prose gets nothing back and every structure command refuses to run there. `prefix` names the
   * table itself — one segment in the body, four for a table nested in a cell — so a command that reshapes a
   * grid can put the caret back in the table it just changed instead of the one outside it.
   */
  private tableAt(caret: DocPath | null): { prefix: number[]; table: Table; where: CellWhere } | undefined {
    const path = caret?.path;
    if (!path || path.length < 4 || (path.length - 1) % 3 !== 0) return undefined;
    const prefix = path.slice(0, path.length - 3);
    const table = this.blockAt(prefix);
    if (table?.kind !== "table") return undefined;
    const where = { row: path[path.length - 3], col: path[path.length - 2] };
    // A path whose cell is no longer there is a stale caret, not a caret in this table: answering with the
    // table would let a command reshape a grid the user is not in.
    if (!table.rows[where.row]?.[where.col]) return undefined;
    return { prefix, table, where };
  }

  /** Both ends of a selection, when it starts and ends in cells of the same table. */
  private tableSpan(): { prefix: number[]; table: Table; from: CellWhere; to: CellWhere } | undefined {
    const ends = this.pageSelection;
    if (!ends) return undefined;
    const from = this.tableAt(ends.start);
    const to = this.tableAt(ends.end);
    if (!from || !to || from.prefix.join() !== to.prefix.join()) return undefined;
    return { prefix: from.prefix, table: from.table, from: from.where, to: to.where };
  }

  /**
   * Runs one structure command. Every command decides whether the shape can take it before it touches
   * anything, so a refusal costs the model nothing and only the undo entry taken for it has to come back.
   */
  private tableCommand(caret: DocPath | null, run: (table: Table, where: CellWhere) => TableEdit): CommandOutcome | undefined {
    const target = this.tableAt(caret);
    if (!target) return undefined;
    this.snapshot(caret);
    const edit = run(target.table, target.where);
    if (!edit.ok) {
      this.undoStack.pop();
      return edit;
    }
    return { ok: true, at: edit.at, prefix: target.prefix, table: target.table };
  }

  /** Repaints after a structural change and puts the caret on the cell the command named. */
  private settleTableEdit(prefix: number[], at: CellWhere): void {
    this.render();
    this.setCaret({ path: [...prefix, at.row, at.col, 0], offset: 0 });
    this.hooks.changed();
  }

  /** 修订: this editor's own signature on a change it is making right now. */
  private sign(kind: "ins" | "del"): Revision {
    this.revSeq += 1;
    return { kind, id: this.revSeq, author: this.authorName, date: stampNow() };
  }

  /** Puts a whole row's record on the row, and — for a deletion — strikes every run in it the same way. */
  private signRow(table: Table, index: number, rev: Revision): void {
    const row: RowInfo = { rev };
    table.rows[index]?.forEach((cell) => {
      cell.row = row;
      if (rev.kind !== "del") return;
      cell.paragraphs.forEach((block) => {
        if (block.kind === "paragraph") block.inlines.forEach((inline) => (inline.rev = rev));
      });
    });
  }

  /**
   * 插入行. Word's own shortcut puts the caret in the new row's first cell; with 修订 on, the row signs itself
   * as an insertion in its `w:trPr`, which is where a row's record belongs rather than around each of its runs.
   */
  insertRow(side: "above" | "below"): boolean {
    const caret = this.caretPath();
    const ran = this.tableCommand(caret, (table, where) => insertRowHere(table, where, side));
    if (!ran || !ran.ok) return false;
    if (this.trackRevisions) this.signRow(ran.table, ran.at.row, this.sign("ins"));
    this.settleTableEdit(ran.prefix, ran.at);
    return true;
  }

  /** 插入列: the grid gains a column and the table keeps the width it had. */
  insertColumn(side: "left" | "right"): boolean {
    const ran = this.tableCommand(this.caretPath(), (table, where) => insertColumnHere(table, where, side));
    if (!ran || !ran.ok) return false;
    this.settleTableEdit(ran.prefix, ran.at);
    return true;
  }

  /**
   * 插入单元格: one cell appears and the ones already there give way, to the right or downward. The caret ends on
   * the new empty cell. A shape the command cannot take — a merge in the column being shifted, a caret row with
   * no row under it — is refused with the reason, the way 合并单元格 reports it, rather than redrawn as something
   * nobody asked for.
   */
  insertCell(shift: "right" | "down"): { ok: boolean; reason?: string } {
    const ran = this.tableCommand(this.caretPath(), (table, where) => insertCellHere(table, where, shift));
    if (!ran) return { ok: false, reason: "光标不在表格里" };
    if (!ran.ok) return { ok: false, reason: ran.reason };
    this.settleTableEdit(ran.prefix, ran.at);
    return { ok: true };
  }

  /**
   * 复制表格: take the table the caret sits in — the one the paste commands will put back. The caller gets the
   * same object the editor keeps, because a paste clones it again; handing the copy to the shell is how a table
   * copied in one document reaches another document's editor, where the shell passes it as `source`.
   */
  copyTable(): { ok: boolean; reason?: string; rows?: number; cols?: number; table?: Table } {
    const target = this.tableAt(this.caretPath());
    if (!target) return { ok: false, reason: "光标不在表格里" };
    const copy = structuredClone(target.table);
    this.clipTable = copy;
    return { ok: true, rows: copy.rows.length, cols: gridColumns(copy), table: copy };
  }

  /**
   * 粘贴为嵌套表: the copied table lands inside the cell the caret sits in, below the paragraph the caret is
   * on. It is a copy, so the original stays where it was and the two never share a cell or a row record.
   */
  pasteAsNestedTable(source?: Table): { ok: boolean; reason?: string } {
    const caret = this.caretPath();
    const target = this.tableAt(caret);
    if (!target) return { ok: false, reason: "光标不在表格里" };
    const clip = source ?? this.clipTable;
    if (!clip) return { ok: false, reason: "剪贴板里没有表格" };
    const blocks = this.containerBlocks(target.prefix);
    const index = target.prefix[target.prefix.length - 1];
    if (!blocks || blocks[index] !== target.table) return { ok: false, reason: "这张表已经不在原来的位置" };
    this.snapshot(caret);
    const pasted = pasteAsNestedTableHere(target.table, target.where, clip, caret!.path[caret!.path.length - 1] + 1);
    if (!pasted.ok) {
      this.undoStack.pop();
      return { ok: false, reason: pasted.reason };
    }
    this.render();
    // 光标进新表的第一格：嵌套表是内容，落点就在内容里头。
    this.setCaret({
      path: [...target.prefix, target.where.row, target.where.col, pasted.block, 0, 0, 0],
      offset: 0,
    });
    this.hooks.changed();
    return { ok: true };
  }

  /** 粘贴为追加行: the copied table's rows join this table at the bottom, sized to its columns. */
  pasteAppendTable(source?: Table): { ok: boolean; reason?: string } {
    const caret = this.caretPath();
    const clip = source ?? this.clipTable;
    if (!clip) return { ok: false, reason: "剪贴板里没有表格" };
    const ran = this.tableCommand(caret, (table, where) => {
      const pasted = pasteAppendTableHere(table, where, clip);
      return pasted.ok ? { ok: true, at: pasted.at } : pasted;
    });
    if (!ran) return { ok: false, reason: "光标不在表格里" };
    if (!ran.ok) return { ok: false, reason: ran.reason };
    this.settleTableEdit(ran.prefix, ran.at);
    return { ok: true };
  }

  /**
   * 删除行. With 修订 on the row stays in the file and says it was removed — `w:trPr/w:del` on the row, its
   * runs struck — because a row somebody deleted is still a row until the change is accepted.
   */
  deleteRow(): boolean {
    const caret = this.caretPath();
    const target = this.tableAt(caret);
    if (!target) return false;
    if (this.trackRevisions) {
      this.snapshot(caret);
      const rev = this.sign("del");
      this.signRow(target.table, target.where.row, rev);
      this.settleTableEdit(target.prefix, target.where);
      return true;
    }
    const ran = this.tableCommand(caret, deleteRowHere);
    if (!ran || !ran.ok) return false;
    this.settleTableEdit(ran.prefix, ran.at);
    return true;
  }

  /** 删除列. */
  deleteColumn(): boolean {
    const ran = this.tableCommand(this.caretPath(), deleteColumnHere);
    if (!ran || !ran.ok) return false;
    this.settleTableEdit(ran.prefix, ran.at);
    return true;
  }

  /**
   * 合并单元格: the rectangle the selection covers becomes one cell. A merge that crosses a cell already
   * merged is refused by the command, and the page says why rather than redrawing a shape nobody asked for.
   */
  mergeCells(): { ok: boolean; reason?: string } {
    const span = this.tableSpan();
    if (!span) return { ok: false, reason: "先把光标放在一个表格里" };
    const ran = this.tableCommand(this.pageSelection?.start ?? null, (table) => mergeHere(table, span.from, span.to));
    if (!ran) return { ok: false, reason: "先把光标放在一个表格里" };
    if (!ran.ok) return { ok: false, reason: ran.reason };
    this.settleTableEdit(ran.prefix, ran.at);
    return { ok: true };
  }

  /** 拆分单元格: a merged cell back into the grid it swallowed. */
  splitCell(): boolean {
    const ran = this.tableCommand(this.caretPath(), splitHere);
    if (!ran || !ran.ok) return false;
    this.settleTableEdit(ran.prefix, ran.at);
    return true;
  }

  /**
   * 拆分表格: the caret's row and the rows under it leave the table and become a second one, with an empty
   * paragraph between — Word puts the caret in the lower table's first cell, and so does this. The split
   * happens wherever the caret's table lives, so in a cell it is that cell that ends up holding a paragraph
   * and two tables while the table around it keeps the rows it had.
   */
  splitTable(): { ok: boolean; reason?: string } {
    const caret = this.caretPath();
    const target = this.tableAt(caret);
    if (!target) return { ok: false, reason: "光标不在表格里" };
    const blocks = this.containerBlocks(target.prefix);
    const index = target.prefix[target.prefix.length - 1];
    // The path the page gave has to still name this table in that list; a caret left over from an earlier
    // render would otherwise splice a table in beside the wrong one.
    if (!blocks || blocks[index] !== target.table) return { ok: false, reason: "这张表已经不在原来的位置" };
    this.snapshot(caret);
    const split = splitTableAt(target.table, target.where.row);
    if (!split.ok) {
      this.undoStack.pop();
      return split;
    }
    blocks.splice(index + 1, 0, { kind: "paragraph", inlines: [] }, split.tail);
    this.render();
    this.setCaret({ path: [...target.prefix.slice(0, target.prefix.length - 1), index + 2, 0, 0, 0], offset: 0 });
    this.hooks.changed();
    return { ok: true };
  }

  /** 选中整行: every cell the caret's row paints, a merged one taken whole. */
  selectRow(): TableSelection {
    return this.selectCells("row");
  }

  /** 选中整列: the grid column the caret's own cell starts at, through every row that reaches it. */
  selectColumn(): TableSelection {
    return this.selectCells("column");
  }

  /** 选中当前单元格: the caret's cell, expanded to the `w:gridSpan` and `w:vMerge` band it belongs to. */
  selectCell(): TableSelection {
    return this.selectCells("cell");
  }

  /** 选中整个表格: every cell the table paints, its own table only — a caret in a nested table selects that one. */
  selectTable(): TableSelection {
    return this.selectCells("table");
  }

  /**
   * The four 选择单元格 commands (`wps.KSelection.SelectRow`, `SelectColumn`, `SelectCell`, `SelectTable`). Only the
   * page's selection moves: the document keeps every byte, so there is no snapshot and no `hooks.changed()` —
   * selecting is not an edit, and the shell learns about it from the browser's own `selectionchange`. The two ends
   * land in `pageSelection` in the same shape a dragged selection takes, which is what lets 合并单元格 work on a
   * range the user reached from the toolbar instead of the mouse.
   */
  private selectCells(kind: "cell" | "row" | "column" | "table"): TableSelection {
    const caret = this.caretPath();
    const target = this.tableAt(caret);
    if (!target) return { ok: false, reason: "光标不在表格里" };
    const range = cellRangeFor(target.table, target.where, kind);
    if (!range.ok) return { ok: false, reason: range.reason };
    const covered = cellsInRect(target.table, range.range);
    const head = covered.length ? this.cellBoundary(target.prefix, covered[0]!, "first") : undefined;
    const tail = covered.length ? this.cellBoundary(target.prefix, covered[covered.length - 1]!, "last") : undefined;
    if (!head || !tail) return { ok: false, reason: "这一格里没有可选中的文字" };
    this.paintCells(target.prefix, covered);
    const from = this.pathOf(head, 0);
    const to = this.pathOf(tail, tail.childNodes.length);
    if (from && to) this.pageSelection = { start: from, end: to };
    // 一条 Range 铺得满整段（整格、整行）就让它亮出去，复制与键入替换跟着的都是这一段。
    // 整列铺不满：格与格不相邻，而可编辑页只留得下一条 Range（实测后 addRange 的那几条被丢掉），
    // 所以列上每一格靠 `.solo-cell-sel` 逐格点亮，页面里的选区退回光标自己那一格。
    const whole = cellsInSweep(target.table, covered[0]!, covered[covered.length - 1]!).length === covered.length;
    const shown = whole ? { start: head, end: tail } : this.cellEnds(target.prefix, target.where) ?? { start: head, end: tail };
    const span = document.createRange();
    span.setStart(shown.start, 0);
    span.setEnd(shown.end, shown.end.childNodes.length);
    this.paintedCells = { prefix: target.prefix.join(), cells: new Set(covered.map((c) => `${c.row},${c.col}`)) };
    this.select(span, shown.start);
    return { ok: true, selected: this.selectedCells(target.table) || covered.length };
  }

  /** Light the cells a selection command picked, and nothing else. The class is drawn on the page only: the
   * model is re-derived from the page's text and marks, and a render rebuilds these cells from it anyway. */
  private paintCells(prefix: number[], covered: CellWhere[]): void {
    this.clearPaint();
    for (const where of covered) {
      const block = this.elementAt([...prefix, where.row, where.col, 0]);
      block?.parentElement?.classList.add("solo-cell-sel");
    }
  }

  private clearPaint(): void {
    for (const td of this.page?.querySelectorAll("td.solo-cell-sel") ?? []) td.classList.remove("solo-cell-sel");
  }

  /** The two page elements one cell's own text runs between. */
  private cellEnds(prefix: number[], where: CellWhere): { start: Element; end: Element } | null {
    const start = this.cellBoundary(prefix, where, "first");
    const end = this.cellBoundary(prefix, where, "last");
    return start && end ? { start, end } : null;
  }

  /**
   * The block the page starts or stops a cell-range selection on: the first or last paragraph inside the cell.
   * A nested table is skipped because a selection has to end on something the page can put a caret in.
   */
  private cellBoundary(prefix: number[], where: CellWhere, side: "first" | "last"): Element | undefined {
    const blocks = this.containerBlocks([...prefix, where.row, where.col, 0]);
    if (!blocks?.length) return undefined;
    let at = side === "first" ? blocks.findIndex((b) => b.kind === "paragraph") : -1;
    if (at < 0 && side === "last") for (let i = blocks.length - 1; i >= 0; i -= 1) if (blocks[i].kind === "paragraph") { at = i; break; }
    if (at < 0) at = side === "first" ? 0 : blocks.length - 1;
    return this.elementAt([...prefix, where.row, where.col, at]) ?? undefined;
  }

  /**
   * What the shell needs to light the toolbar up: the table's size, whether the caret's cell is a merged one,
   * and how many cells the selection covers. Nothing here changes the document.
   */
  tableContext(): { inTable: boolean; rows: number; columns: number; merged: boolean; selected: number } {
    const target = this.tableAt(this.caretPath());
    if (!target) return { inTable: false, rows: 0, columns: 0, merged: false, selected: 0 };
    const rect = cellRect(target.table, target.where);
    return {
      inTable: true,
      rows: target.table.rows.length,
      columns: gridColumns(target.table),
      merged: rect !== undefined && (rect.c1 - rect.c0 > 1 || rect.r1 - rect.r0 > 1),
      selected: this.selectedCells(target.table) || 1,
    };
  }

  /** How many cells the selection covers: a merged cell counts once, wherever it sits in the grid. */
  private selectedCells(table: Table): number {
    const span = this.tableSpan();
    return span ? cellsInRect(table, { from: span.from, to: span.to }).length : 0;
  }

  /** An inline picture at the caret, sized to fit between the page margins. */
  async insertPicture(url: string, name: string): Promise<void> {
    const caret = this.caretPath();
    const paragraph = caret ? this.paragraphAt(caret.path) : undefined;
    if (!caret || !paragraph) return;
    const { width, height } = await imageSize(url);
    const section = this.doc.sections[0];
    const usable = section ? twipsToPx(section.pageWidthTwips - section.margins.left - section.margins.right) : 620;
    const scale = Math.min(1, usable / Math.max(1, width));
    const w = Math.max(24, Math.round(width * scale));
    const h = Math.max(24, Math.round(height * scale));
    const ext = /^data:image\/([a-z0-9.+-]+)/i.exec(url)?.[1] ?? "png";
    this.snapshot(caret);
    const image: NonNullable<Inline["image"]> = { src: url, width: w, height: h, name, path: this.mediaPath(ext) };
    const [head, tail] = splitAt(paragraph.inlines, caret.offset);
    paragraph.inlines = merge([...head, { text: "", image }, ...tail]);
    this.render();
    this.setCaret({ path: caret.path, offset: caret.offset + 1 });
    this.hooks.changed();
  }

  /** Next free media name; existing pictures keep theirs, so nothing is overwritten. */
  private mediaPath(ext: string): string {
    let max = 0;
    const scan = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind === "table") b.rows.forEach((row) => row.forEach((c) => scan(c.paragraphs)));
        else if (b.kind === "paragraph")
          for (const inline of b.inlines) {
            const m = /(\d+)\.[a-z0-9]+$/i.exec(inline.image?.path ?? "");
            if (m) max = Math.max(max, Number(m[1]));
          }
      }
    };
    scan(this.doc.blocks);
    return `word/media/solo-image${max + 1}.${IMAGE_EXT_BY_MIME[`image/${ext}`] ?? "png"}`;
  }

  /**
   * Starts a new page: an empty paragraph with `pageBreakBefore` right after the
   * caret's block, and the caret inside it.
   */
  insertPageBreak(): void {
    const path = this.caretPath()?.path;
    if (!path) return;
    this.snapshot(this.caretPath());
    const at = path[0];
    this.doc.blocks.splice(at + 1, 0, { kind: "paragraph", inlines: [], pageBreakBefore: true });
    this.render();
    this.setCaret({ path: [at + 1], offset: 0 });
    this.hooks.changed();
  }

  // ------------------------------------------------------------- toolbar

  private paragraphAt(path: number[]): Paragraph | undefined {
    const block = this.blockAt(path);
    return block?.kind === "paragraph" ? block : undefined;
  }

  /** What the caret's run and paragraph currently look like, so the toolbar can show its state. */
  markState(): { bold: boolean; italic: boolean; underline: boolean; strike: boolean; sizePt?: number; align?: string; list?: "bullet" | "number" } {
    const sel = document.getSelection();
    let el: Element | null = null;
    if (sel?.rangeCount) {
      const node = sel.getRangeAt(0).startContainer;
      el = node instanceof Element ? node : node.parentElement;
    }
    let carried: RunStyle = {};
    for (let scan = el; scan; scan = scan.parentElement) {
      carried = { ...styleOf(scan), ...carried };
      const tag = scan.tagName.toLowerCase();
      if (tag === "b" || tag === "strong") carried.bold = true;
      if (tag === "i" || tag === "em") carried.italic = true;
      if (tag === "u") carried.underline = true;
      if (scan.classList.contains("para")) break;
    }
    const path = this.caretPath()?.path;
    const paragraph = path ? this.paragraphAt(path) : undefined;
    const list = paragraph?.numPr ? listKind(this.doc, paragraph.numPr.numId) : undefined;
    return {
      bold: carried.bold ?? false,
      italic: carried.italic ?? false,
      underline: carried.underline ?? false,
      strike: carried.strike ?? false,
      sizePt: carried.sizePt,
      align: paragraph?.align,
      list,
    };
  }

  /** The numbering definitions the toolbar can offer come from the document itself. */
  listAvailability(): { bullet: boolean; number: boolean } {
    return { bullet: listNumId(this.doc, "bullet") !== undefined, number: listNumId(this.doc, "number") !== undefined };
  }

  applyStyle(patch: StylePatch): void {
    const sel = document.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (sel.getRangeAt(0).collapsed) {
      // Nothing to restyle yet: leave the format pending for whatever gets typed next.
      const command: Record<string, string | undefined> = {
        bold: "bold",
        italic: "italic",
        underline: "underline",
        strike: "strikeThrough",
        color: "foreColor",
        highlight: "hiliteColor",
      };
      const key = Object.keys(patch)[0];
      const value = patch[key as keyof StylePatch];
      if (key === "sizePt") return;
      document.execCommand("styleWithCSS", false, "true");
      if (key === "highlight") document.execCommand("hiliteColor", false, value === "" || value === undefined ? "transparent" : highlightColor(String(value)));
      else document.execCommand(command[key]!, false, typeof value === "string" ? value : undefined);
      this.hooks.changed();
      return;
    }
    const ends = this.rangeEnds();
    if (!ends) return;
    this.snapshot(ends.start);
    const sign = this.trackRevisions ? (before: Inline) => this.signFormatRevision(before) : undefined;
    if (ends.start.path.join() === ends.end.path.join()) {
      this.patchParagraph(this.paragraphAt(ends.start.path), ends.start.offset, ends.end.offset, patch, sign);
    } else {
      const start = this.paragraphAt(ends.start.path);
      if (start) start.inlines = restyle(start.inlines, ends.start.offset, Infinity, patch, sign);
      const end = this.paragraphAt(ends.end.path);
      if (end) end.inlines = restyle(end.inlines, 0, ends.end.offset, patch, sign);
      for (let i = ends.start.path[0] + 1; i < ends.end.path[0]; i += 1) {
        const block = this.doc.blocks[i];
        if (block?.kind === "paragraph") block.inlines = restyle(block.inlines, 0, Infinity, patch, sign);
      }
    }
    this.render();
    this.selectRange(ends.start, ends.end);
    this.hooks.changed();
  }

  private patchParagraph(paragraph: Paragraph | undefined, from: number, to: number, patch: StylePatch, sign?: (before: Inline) => FormatRevision | undefined): void {
    if (!paragraph) return;
    paragraph.inlines = restyle(paragraph.inlines, from, to, patch, sign);
  }

  private selectRange(start: DocPath, end: DocPath): void {
    const startEl = this.elementAt(start.path);
    const endEl = this.elementAt(end.path);
    if (!startEl || !endEl) return;
    const a = positions(startEl)[Math.min(start.offset, positions(startEl).length - 1)];
    const b = positions(endEl)[Math.min(end.offset, positions(endEl).length - 1)];
    if (!a || !b) return;
    const range = document.createRange();
    range.setStart(a.node, a.offset);
    range.setEnd(b.node, b.offset);
    this.select(range, startEl);
  }

  /** Paragraph-level properties go straight from the toolbar into the model. */
  setParagraph(patch: { align?: string; list?: "bullet" | "number" | "none" }): void {
    const path = this.caretPath()?.path;
    const paragraph = path ? this.paragraphAt(path) : undefined;
    if (!paragraph) return;
    this.snapshot(this.caretPath());
    const was = paraStyleOf(paragraph);
    // `syncParaSlots` replaces the array, so this reference keeps the set as the paragraph came in with it.
    const wasSlots = paragraph.pPrSlots;
    if (patch.align !== undefined) paragraph.align = patch.align || undefined;
    if (patch.list !== undefined) {
      const numId = listNumId(this.doc, patch.list);
      paragraph.numPr = patch.list === "none" || numId === undefined ? undefined : { numId, ilvl: paragraph.numPr?.ilvl ?? 0 };
    }
    // The paragraph's `w:pPr` travels as text, so the child these two buttons own moves with them.
    syncParaSlots(paragraph);
    // An alignment click on a paragraph that already sits there changes nothing, and a change that is
    // nothing should not be put in front of a reviewer.
    if (this.trackRevisions && JSON.stringify(was) !== JSON.stringify(paraStyleOf(paragraph))) {
      paragraph.revFormat = this.signParagraphRevision(paragraph, was, wasSlots);
    }
    const caret = this.caretPath();
    this.render();
    if (caret) this.setCaret(this.clamp(caret));
    this.hooks.changed();
  }

  // -------------------------------------------------------------- revisions

  /** The insertions and deletions still waiting for a decision, in body order. */
  revisions(): Revision[] {
    return this.doc.revisions();
  }

  /** The formatting changes (`w:rPrChange`) still waiting. A pane that lists every kind shows both. */
  formatRevisions(): FormatRevision[] {
    return this.doc.formatRevisions();
  }

  /** The paragraph-mark changes (`w:pPr/w:rPr/w:ins` / `w:del`) still waiting — WPS lists these as 段落改动. */
  markRevisions(): Revision[] {
    return this.doc.markRevisions();
  }

  /** The paragraph property changes (`w:pPrChange`) still waiting. */
  paragraphFormatRevisions(): ParagraphFormatRevision[] {
    return this.doc.paragraphFormatRevisions();
  }

  /** Every change still waiting, in document order and numbered from 1 — one row per 审阅窗格 entry. */
  revisionEntries(): RevisionEntry[] {
    return this.doc.revisionEntries();
  }

  /** 接受第 N 条修订, counted the way `revisionEntries()` lists them. Undoable. */
  acceptRevisionAt(index: number): boolean {
    return this.resolveAt("accept", index);
  }

  /** 拒绝第 N 条修订, counted the way `revisionEntries()` lists them. Undoable. */
  rejectRevisionAt(index: number): boolean {
    return this.resolveAt("reject", index);
  }

  /** 按作者接受修订: everything this name signed, nothing anybody else did. One undo step. */
  acceptRevisionsByAuthor(author: string): number {
    return this.resolveByAuthor("accept", author);
  }

  /** 按作者拒绝修订: everything this name signed, nothing anybody else did. One undo step. */
  rejectRevisionsByAuthor(author: string): number {
    return this.resolveByAuthor("reject", author);
  }

  private resolveAt(outcome: "accept" | "reject", index: number): boolean {
    const entry = this.doc.revisionEntries()[index - 1];
    if (!entry) return false;
    return this.resolveRevisions(outcome, entry.id) > 0;
  }

  private resolveByAuthor(outcome: "accept" | "reject", author: string): number {
    if (!this.doc.revisionEntries().some((e) => e.author === author)) return 0;
    // One snapshot for the batch, so 撤销 puts the whole author's decision back at once.
    this.snapshot(this.caretPath());
    const done = this.doc.resolveRevisionsByAuthor(outcome, author);
    this.pruneOrphans();
    this.render();
    this.hooks.changed();
    return done;
  }

  /**
   * 定位到第 N 条修订: selects what the change covers and brings it into view, the way clicking a row in
   * WPS' 审阅窗格 jumps to the change it lists. A break and a paragraph's property change have no
   * characters of their own to select, so the caret goes to the end of the paragraph that carries them.
   */
  revealRevision(index: number): boolean {
    if (!this.page) return false;
    const entry = this.doc.revisionEntries()[index - 1];
    if (!entry) return false;
    const marks = [...this.page.querySelectorAll<HTMLElement>(
      `[data-rev-id="${entry.id}"], [data-rev-format="${entry.id}"], [data-rev-mark-id="${entry.id}"], [data-para-rev="${entry.id}"], [data-row-rev-id="${entry.id}"]`,
    )];
    if (!marks.length) return false;
    this.activeRevision = entry.id;
    for (const el of marks) el.classList.add("solo-rev-hit");
    const sel = document.getSelection();
    if (sel) {
      const range = document.createRange();
      const first = marks[0];
      const last = marks[marks.length - 1];
      if (first.classList.contains("para")) {
        range.selectNodeContents(last);
        range.collapse(false);
      } else {
        range.setStartBefore(first.firstChild ?? first);
        range.setEndAfter(last.lastChild ?? last);
      }
      sel.removeAllRanges();
      sel.addRange(range);
    }
    marks[0].scrollIntoView({ block: "center" });
    this.page.focus({ preventScroll: true });
    return true;
  }

  /**
   * WPS's 下一条修订 / 上一条修订: walks the rows in the order `revisionEntries()` lists them, wrapping at
   * both ends, and answers the row it landed on. The first 下一条 from no row at all is the first one.
   */
  stepRevision(step: 1 | -1): number | undefined {
    const list = this.doc.revisionEntries();
    if (!list.length) return undefined;
    const current = list.findIndex((e) => e.id === this.activeRevision);
    const at = current < 0 ? (step === 1 ? 0 : list.length - 1) : (current + step + list.length) % list.length;
    const target = list[at]!;
    this.revealRevision(target.index);
    return target.index;
  }

  /**
   * Which row of the list the selection sits in — WPS' `RevisionIndexByBegin`. A pane uses it to
   * highlight the change under the caret, and a ribbon to decide whether 接受/拒绝这一条 has anything
   * to act on. Where two changes overlap it answers the innermost one, the smallest range there.
   */
  revisionIndexAtCaret(): number | undefined {
    const caret = this.caretPath();
    if (!caret) return undefined;
    const p = this.paragraphAt(caret.path);
    return p ? this.doc.revisionIndexAt(p, caret.offset) : undefined;
  }

  /** 接受修订: every insertion stands, every deletion takes effect, and revised formatting is kept. Undoable. */
  acceptRevisions(id?: number): number {
    return this.resolveRevisions("accept", id);
  }

  /** 拒绝修订: the deleted text comes back, the inserted text goes, and formatting returns to what it was. Undoable. */
  rejectRevisions(id?: number): number {
    return this.resolveRevisions("reject", id);
  }

  private resolveRevisions(outcome: "accept" | "reject", id?: number): number {
    // Every kind of pending change is listed in one collection, and each row carries the `w:id` it would
    // decide — runs, formatting, paragraph marks, and a table row's own record included. Consulting a
    // subset of the lists would refuse to decide a kind it does not know and report that nothing needed
    // deciding, which is what this guard is for.
    const waits = id === undefined || this.doc.revisionEntries().some((e) => e.id === id);
    if (!waits) return 0;
    this.snapshot(this.caretPath());
    const touched = this.doc.resolveRevisions(outcome, id);
    // Accepting a deletion can take a comment marker with it, which is the same loss typing causes.
    this.pruneOrphans();
    this.render();
    this.hooks.changed();
    return touched;
  }

  /**
   * 显示修订 / 显示最终状态 / 显示原文 — Word's view selector, and it changes only what is on
   * screen: the marks stay in the model and in the file, so looking is never deciding.
   */
  setRevisionView(view: RevisionView): void {
    this.revisionView = view;
    this.applyRevisionView();
  }

  get revisionShown(): RevisionView {
    return this.revisionView;
  }

  private applyRevisionView(): void {
    if (!this.page) return;
    this.page.classList.toggle("rev-final", this.revisionView === "final");
    this.page.classList.toggle("rev-original", this.revisionView === "original");
  }

  /** 修订 / 追踪修订: with this on, edits are recorded as tracked changes instead of applied. */
  setTrackChanges(on: boolean): void {
    this.trackRevisions = on;
  }

  get trackingChanges(): boolean {
    return this.trackRevisions;
  }

  /**
   * The record a formatting edit leaves behind while tracking is on. Revising a range nobody has
   * decided yet keeps the record already on it — its signature and the formatting from before that
   * change — so one reviewer never re-signs another's, and refusing undoes both edits at once.
   */
  private signFormatRevision(before: Inline): FormatRevision {
    const mine = before.revFormat?.author === this.authorName ? before.revFormat : undefined;
    if (mine) return mine;
    this.revSeq += 1;
    // The record says what the run's whole property set was, not only the nine fields the model can name, so a
    // later refusal puts back the `w:kern` and the `w:lang` the toolbar never touched.
    return {
      kind: "format",
      id: this.revSeq,
      author: this.authorName,
      date: stampNow(),
      prev: runStyleOf(before),
      ...(before.rPrSlots?.length ? { prevSlots: before.rPrSlots } : {}),
    };
  }

  /**
   * The same for a paragraph: `prev` is the properties from before this edit, given in by the caller
   * because the paragraph has already been changed by the time the record is hung on it — and so is its
   * `prevSlots`, since the set the paragraph owned before is what a reviewer gets to refuse.
   */
  private signParagraphRevision(p: Paragraph, prev: ParaStyle, prevSlots?: PropSlot[]): ParagraphFormatRevision {
    const mine = p.revFormat?.author === this.authorName ? p.revFormat : undefined;
    if (mine) return mine;
    this.revSeq += 1;
    return { kind: "format", id: this.revSeq, author: this.authorName, date: stampNow(), prev, ...(prevSlots?.length ? { prevSlots } : {}) };
  }

  /**
   * Every paragraph the body shows, table cells included, in reading order and grouped by the stretch
   * of paragraphs a break could join them with: one run of body paragraphs, and each table cell.
   */
  private bodyText(): BodyText {
    const spans: { paragraph: Paragraph; length: number; group: number }[] = [];
    let group = 0;
    let at = 0;
    while (at < this.doc.blocks.length) {
      const block = this.doc.blocks[at]!;
      if (block.kind === "paragraph") {
        while (at < this.doc.blocks.length && this.doc.blocks[at]!.kind === "paragraph") {
          const paragraph = this.doc.blocks[at] as Paragraph;
          spans.push({ paragraph, length: charLength(paragraph.inlines), group });
          at += 1;
        }
        group += 1;
        continue;
      }
      // A container the model only carries has no text a break could move, so it ends the stretch it sits
      // in and takes a group number with it — otherwise the stretch after it reads as its replacement.
      if (block.kind === "carried") {
        group += 1;
        at += 1;
        continue;
      }
      for (const row of block.rows) {
        for (const cell of row) {
          for (const paragraph of cell.paragraphs) {
            if (paragraph.kind !== "paragraph") continue;
            spans.push({ paragraph, length: charLength(paragraph.inlines), group });
          }
          group += 1;
        }
      }
      at += 1;
    }
    return { runs: spans.flatMap((s) => s.paragraph.inlines), spans };
  }

  /**
   * 追踪修订: turn the edit the page just made into tracked changes rather than applying it. The diff
   * runs over the whole body instead of one paragraph, because pressing Enter moves characters between
   * paragraphs without adding or removing any, and a move is not a change anyone should sign.
   */
  private recordRevisions(before: BodyText): { marked: boolean; struck: number } {
    const now = this.bodyText();
    const oldText = flatOf(before.runs);
    const newText = flatOf(now.runs);
    if (!now.spans.length) return { marked: false, struck: 0 };
    // No character moved, so any difference the page shows is the boundary between paragraphs having
    // moved: Enter or Backspace, which is a revision of its own kind and is recorded as one.
    if (oldText === newText) return { marked: this.recordBreaks(before, now), struck: 0 };

    let at = 0;
    while (at < oldText.length && at < newText.length && oldText[at] === newText[at]) at += 1;
    let back = 0;
    while (back < oldText.length - at && back < newText.length - at && oldText[oldText.length - 1 - back] === newText[newText.length - 1 - back]) back += 1;

    const [, restOld] = splitAt(before.runs, at);
    const [removed] = splitAt(restOld, oldText.length - at - back);
    const [head, restNew] = splitAt(now.runs, at);
    const [inserted, tail] = splitAt(restNew, newText.length - at - back);

    const date = stampNow();
    let seq = [...before.runs, ...now.runs].reduce((n, i) => Math.max(n, i.rev?.id ?? 0), this.revSeq);
    const struck: Inline[] = [];
    let delId = 0;
    for (const run of removed) {
      // Text the file already signed keeps its own author and id — nothing here rewrites someone else's
      // revision. An insertion that gets deleted is cancelled outright: those characters were never in
      // the original, so there is nothing for anyone to decide about them.
      if (run.rev) {
        if (run.rev.kind === "del") struck.push({ ...run });
        continue;
      }
      if (delId === 0) delId = ++seq;
      struck.push({ ...run, rev: { kind: "del", id: delId, author: this.authorName, date } });
    }
    const insId = inserted.length ? ++seq : 0;
    const ins = inserted.map((run) => ({ ...run, rev: { kind: "ins" as const, id: insId, author: this.authorName, date } }));
    const struckChars = charLength(struck);

    // The pieces each paragraph gets are cut from one stream: what came before the change, then the struck
    // characters, then what was typed, then what followed. A paragraph that starts at or after the change
    // point picks the stream up `struckChars` later than it looks on the page, because the deleted
    // characters are still in the file and belong to the paragraph they were taken from.
    const bounds: { obj: Paragraph; group: number; rest: number; cut: boolean }[] = [];
    let pos = 0;
    for (let i = 0; i < now.spans.length; i += 1) {
      const span = now.spans[i]!;
      bounds.push({ obj: span.paragraph, group: span.group, rest: pos < at || (pos === at && i === 0) ? pos : pos + struckChars, cut: false });
      pos += span.length;
    }

    // A selection that reached over a break deleted that break along with its characters, and the page shows
    // one merged paragraph where the file must keep both halves while the break waits for a decision: the
    // paragraph the walk dropped comes back with its own style and its own pending records, taking the struck
    // text that belongs above it. Which breaks were eaten reads off the file's coordinates — the removed
    // region is `at`..`oldText.length - back`, and a pure insertion eats none. Characters the file had
    // already signed as an insertion cancel out instead of being struck, so they take their offset with them.
    const keptBefore = (into: number): number => {
      let over = 0;
      let kept = 0;
      for (const run of removed) {
        const length = inlineLength(run);
        if (into <= over + length) return run.rev && run.rev.kind !== "del" ? kept : kept + (into - over);
        over += length;
        if (!run.rev || run.rev.kind === "del") kept += length;
      }
      return kept;
    };
    const survivors = new Set(now.spans.map((s) => s.paragraph));
    const cuts = swallowedBreaks(before, at, oldText.length - back)
      // Whichever half the page kept as its element, the other half is the one being restored.
      .map((b) => ({ obj: survivors.has(b.prior) ? b.upper : b.prior, group: b.group, rest: at + keptBefore(b.at - at), cut: true }))
      .filter((c) => !survivors.has(c.obj))
      // A cut landing where the page already breaks is no change: that boundary survived the edit.
      .filter((c) => !bounds.some((x) => x.rest === c.rest));

    const cut = new Set<number>();
    const ordered = [...bounds, ...cuts].sort((a, b) => a.rest - b.rest || Number(a.cut) - Number(b.cut));
    const total = newText.length + struckChars;
    const plan: { paragraph: Paragraph; length: number; group: number }[] = [];
    const restorals: { carrier: Paragraph; group: number }[] = [];
    for (let i = 0; i < ordered.length; i += 1) {
      const bound = ordered[i]!;
      const length = (ordered[i + 1]?.rest ?? total) - bound.rest;
      const above = ordered[i - 1];
      if (!bound.cut || !above) {
        plan.push({ paragraph: bound.obj, length, group: bound.group });
        continue;
      }
      const waits = above.obj.markRev;
      if (waits && waits.author === this.authorName && waits.kind === "ins") {
        // This session inserted the break and the edit ate it: no change, so no boundary comes back.
        delete above.obj.markRev;
        const piece = plan[plan.length - 1];
        if (piece) piece.length += length;
        continue;
      }
      plan.push({ paragraph: bound.obj, length, group: bound.group });
      restorals.push({ carrier: above.obj, group: bound.group });
    }

    // A break somebody else already signed keeps that signature: the record is the file's, not ours to
    // overwrite, and both halves still come back so there is something for the decision to act on.
    for (const back of restorals) {
      cut.add(back.group);
      if (back.carrier.markRev) continue;
      back.carrier.markRev = { kind: "del", id: ++seq, author: this.authorName, date };
    }
    this.revSeq = seq;

    let rest = [...head, ...struck, ...ins, ...tail];
    for (const entry of plan) {
      const [slice, more] = splitAt(rest, entry.length);
      entry.paragraph.inlines = merge(slice);
      rest = more;
    }
    if (cut.size) {
      const pieces = new Map<number, Paragraph[]>();
      for (const entry of plan) {
        if (!cut.has(entry.group)) continue;
        const list = pieces.get(entry.group);
        if (list) list.push(entry.paragraph);
        else pieces.set(entry.group, [entry.paragraph]);
      }
      this.rebuildBlocks(pieces);
    }
    return { marked: true, struck: struckChars };
  }

  /**
   * 追踪修订: the breaks the page added or removed, which the character diff above cannot see because
   * no character moved. A break that appeared is one Enter made: the page already shows two paragraphs,
   * which is exactly the shape a break waiting for a decision has, so it only needs a signature. A break
   * that disappeared is one Backspace took: the page shows one joined paragraph, and a change nobody has
   * decided must still keep both halves, so the paragraph is opened back up where the boundary list says
   * it was closed. A break inserted and then removed again cancels out, the way typing then deleting does.
   */
  private recordBreaks(before: BodyText, now: BodyText): boolean {
    const was = groupBreaks(before.spans);
    const is = groupBreaks(now.spans);
    // A paragraph that crossed a table wall left the stretch its breaks belong to. The page decides
    // where it went, and nobody was asked to sign a move.
    if (!sameList([...was.keys()], [...is.keys()])) return false;
    const changed = [...was.keys()].filter((g) => !sameList(was.get(g)!, is.get(g)!));
    if (!changed.length) return false;

    const date = stampNow();
    const holders = [...before.spans, ...now.spans].map((s) => s.paragraph);
    let seq = Math.max(
      this.revSeq,
      ...[...before.runs, ...now.runs].map((i) => Math.max(i.rev?.id ?? 0, i.revFormat?.id ?? 0)),
      ...holders.map((p) => Math.max(p.markRev?.id ?? 0, p.revFormat?.id ?? 0)),
    );

    const pieces = new Map<number, Paragraph[]>();
    for (const g of changed) {
      const bs = spansInGroup(before.spans, g);
      const ns = spansInGroup(now.spans, g);
      const total = ns.reduce((n, s) => n + s.length, 0);
      const endingWas = endsAt(bs);
      const endingNow = endsAt(ns);

      // Every break either reading has, in the order they sit in the text, and what to do about each.
      // Two empty paragraphs at one offset are two breaks, so the count of them there tells them apart,
      // and a break that went is cancelled by the record on the paragraph that used to end there.
      const kept: { at: number; kind?: "ins" | "del" }[] = [];
      for (const end of mergedEnds(was.get(g)!, is.get(g)!, total)) {
        if (!end.kind) {
          kept.push({ at: end.at });
          continue;
        }
        const carrier = end.kind === "del" ? endingWas(end.at, end.occ) : endingNow(end.at, end.occ);
        const waits = carrier?.markRev;
        if (waits && waits.author === this.authorName && waits.kind !== end.kind) {
          // The same break made and then unmade in one session is no change: the record goes, and where
          // the page has already joined the halves again the break does not come back.
          delete carrier!.markRev;
          if (end.kind !== "del") kept.push({ at: end.at });
          continue;
        }
        kept.push({ at: end.at, kind: end.kind });
      }

      const fromNow = startsAt(ns);
      const fromWas = startsAt(bs);
      const out: Paragraph[] = [];
      let rest = ns.flatMap((s) => s.paragraph.inlines);
      let at = 0;
      for (const end of kept) {
        const [slice, more] = splitAt(rest, end.at - at);
        rest = more;
        const start = at;
        at = end.at;
        // A piece keeps the paragraph the page has at the offset it begins at. Where nothing begins there
        // any more, the paragraph the model had comes back with its own style and its own records — that
        // is how a merged-away half returns for a break whose deletion is still waiting to be decided.
        const target = fromNow(start) ?? fromWas(start) ?? fresh(out[out.length - 1]);
        target.inlines = merge(slice);
        out.push(target);
        // Standing as it was, or the end of the stretch, or already carrying somebody's record: nothing to sign.
        if (!end.kind || target.markRev) continue;
        target.markRev = { kind: end.kind, id: ++seq, author: this.authorName, date };
      }
      pieces.set(g, out);
    }
    this.revSeq = seq;
    this.rebuildBlocks(pieces);
    return true;
  }

  /** Puts the reconstructed stretches back where they came from: body runs in place, cells rewritten. */
  private rebuildBlocks(pieces: Map<number, Paragraph[]>): void {
    const rebuilt: Block[] = [];
    let group = 0;
    let at = 0;
    while (at < this.doc.blocks.length) {
      const block = this.doc.blocks[at];
      if (block.kind === "paragraph") {
        const run: Block[] = [];
        while (at < this.doc.blocks.length && this.doc.blocks[at].kind === "paragraph") run.push(this.doc.blocks[at++]);
        rebuilt.push(...(pieces.get(group) ?? run));
        group += 1;
        continue;
      }
      if (block.kind === "carried") {
        rebuilt.push(block);
        group += 1;
        at += 1;
        continue;
      }
      for (const row of block.rows) {
        for (const cell of row) {
          const out = pieces.get(group);
          if (out) cell.paragraphs = out;
          group += 1;
        }
      }
      rebuilt.push(block);
      at += 1;
    }
    this.doc.blocks = rebuilt;
  }

  // -------------------------------------------------------------- comments

  /** The name this session signs its own marks with — a comment and a recorded revision both use it. WPS uses the account; we have none of our own. */
  setCommentAuthor(name: string, initials?: string): void {
    const trimmed = name.trim();
    if (trimmed === "") return;
    this.authorName = trimmed;
    this.authorInitials = (initials ?? trimmed.slice(0, 1)).trim().slice(0, 2) || this.authorInitials;
  }

  listComments(): CommentEntry[] {
    const anchored = anchorIdsIn(this.doc.blocks);
    return this.doc.comments.map((comment, i) => ({
      id: comment.id,
      ordinal: i + 1,
      author: comment.author,
      initials: comment.initials,
      date: comment.date,
      paragraphs: commentTexts(comment),
      runs: comment.paragraphs,
      editable: comment.editable,
      anchored: anchored.has(comment.id),
      resolved: comment.resolved === true,
      parentId: comment.parentId,
      replies: this.doc.comments.filter((c) => c.parentId === comment.id).map((c) => c.id),
    }));
  }

  /** The comments the caret sits inside, read back from the rendered anchored ranges. */
  commentIdsAtCaret(): number[] {
    const sel = document.getSelection();
    if (!sel || sel.rangeCount === 0) return [];
    const node = sel.getRangeAt(0).startContainer;
    let el: Element | null = node instanceof Element ? node : node.parentElement;
    for (let scan = el; scan && scan !== this.page; scan = scan.parentElement) {
      if (scan.classList.contains("cmt-anchor")) {
        const id = Number((scan as HTMLElement).dataset.cid);
        return Number.isFinite(id) ? [id] : [];
      }
      const cids = (scan as HTMLElement).dataset.cids;
      if (cids) return cids.split(",").map(Number).filter(Number.isFinite);
    }
    return [];
  }

  /**
   * Comments the selection, or the caret's position when nothing is selected — Word anchors a
   * zero-length range there too. Closing markers are placed first so the start offset still
   * counts the same characters.
   */
  newComment(text: string): number | undefined {
    const ends = this.rangeEnds() ?? this.pageSelection;
    if (!ends) return undefined;
    const id = this.doc.nextCommentId();
    this.snapshot(ends.start);
    if (!this.placeAnchors(ends.end.path, ends.end.offset, [
      { id, kind: "end" },
      { id, kind: "ref" },
    ])) {
      return undefined;
    }
    if (!this.placeAnchors(ends.start.path, ends.start.offset, [{ id, kind: "start" }])) return undefined;
    this.doc.addComment(id, text, this.authorName, this.authorInitials);
    this.bodyAnchored.add(id);
    this.render();
    this.revealComment(id);
    this.panel?.openAt(id);
    this.hooks.changed();
    return id;
  }

  /** The panel's own "新建": an empty comment at the selection, ready for typing into. */
  newCommentFromPanel(): number | undefined {
    return this.newComment("");
  }

  /** Replaces a comment's body. A rich one the model cannot rebuild refuses, and says so. */
  setComment(id: number, text: string): boolean {
    const found = this.doc.comments.find((c) => c.id === id);
    if (!found || !found.editable) return false;
    const body = text.replace(/\r\n?/g, "\n");
    if (body === commentTexts(found).join("\n")) return false;
    this.snapshot(this.caretPath());
    return this.doc.setCommentBody(id, body);
  }

  /**
   * Replaces a comment's body with the runs the panel's field holds, marks included: a comment
   * written in bold stays bold when its text is edited, and a change of formatting alone counts
   * as a change. A field whose shape already matches the comment commits nothing, so a click in
   * the panel without an edit does not make the document modified.
   */
  setCommentRuns(id: number, paragraphs: Inline[][]): boolean {
    const found = this.doc.comments.find((c) => c.id === id);
    if (!found || !found.editable) return false;
    if (bodyShape(paragraphs) === bodyShape(found.paragraphs)) return false;
    this.snapshot(this.caretPath());
    return this.doc.setCommentRuns(id, paragraphs);
  }

  /**
   * The comment field back into body paragraphs. Each block the browser created by pressing
   * Enter is the next `w:p`; runs typed before the first of those belong to the first one.
   */
  commentParagraphsOf(el: Element): Inline[][] {
    const blocks = [...el.children].filter((c) => /^(div|p|li|ul|ol|table)$/i.test(c.tagName));
    if (!blocks.length) return [this.inlinesOf(el)];
    const loose: Node[] = [];
    for (const node of el.childNodes) {
      if (node === blocks[0]) break;
      loose.push(node);
    }
    const out = blocks.map((b) => this.inlinesOf(b));
    if (loose.length) {
      const head = document.createElement("div");
      // Moving the loose nodes out of the live field is fine: the commit re-renders the panel.
      for (const node of loose) head.append(node);
      out.unshift(this.inlinesOf(head));
    }
    return out;
  }

  deleteComment(id: number): void {
    this.snapshot(this.caretPath());
    this.dropAnchors(id);
    this.doc.removeComment(id);
    this.bodyAnchored.delete(id);
    this.render();
    this.hooks.changed();
  }

  /** Shows the anchored text: scrolled into view, selected, and lit until the next render. */
  revealComment(id: number): boolean {
    if (!this.page) return false;
    let marks = [...this.page.querySelectorAll<HTMLElement>(`.cmt-anchor[data-cid="${id}"]`)];
    if (!marks.length) {
      // A reply carries no markers of its own — Word threads it onto its parent's range —
      // so revealing one jumps to the range that holds the thread.
      const parent = this.doc.comments.find((c) => c.id === id)?.parentId;
      if (parent === undefined || parent === id) return false;
      return this.revealComment(parent);
    }
    this.activeComment = id;
    const ranges = [...this.page.querySelectorAll<HTMLElement>(".cmt-in")].filter((el) => (el.dataset.cids ?? "").split(",").includes(String(id)));
    for (const el of [...marks, ...ranges]) el.classList.add("solo-cmt-anchor-hit");
    const sel = document.getSelection();
    if (sel) {
      const range = document.createRange();
      if (ranges.length) {
        const first = ranges[0];
        const last = ranges[ranges.length - 1];
        range.setStartBefore(first.firstChild ?? first);
        range.setEndAfter(last.lastChild ?? last);
      } else {
        // A zero-length range has no text to select: the caret goes where the mark is.
        range.setStartAfter(marks[marks.length - 1]);
        range.collapse(true);
      }
      sel.removeAllRanges();
      sel.addRange(range);
    }
    marks[0].scrollIntoView({ block: "center" });
    this.page.focus({ preventScroll: true });
    this.panel?.setActive(id);
    return true;
  }

  /** Ribbon toggle for the panel; with no argument it flips whatever the reader is looking at. */
  showComments(show = !(this.panel?.isOpen() ?? false)): void {
    this.panel?.setOpen(show);
  }

  /**
   * `w15:commentEx/@done`. WPS and Word keep a resolved comment visible and greyed, so this
   * only flips the state — nothing in the body or the body's anchors changes.
   */
  setCommentResolved(id: number, resolved: boolean): boolean {
    const found = this.doc.comments.find((c) => c.id === id);
    if (!found || found.resolved === resolved) return false;
    this.snapshot(this.caretPath());
    const changed = this.doc.setCommentResolved(id, resolved);
    this.render();
    this.hooks.changed();
    return changed;
  }

  /**
   * Answers an existing comment. A reply shares its parent's anchored range — Word writes no
   * second set of markers for a thread — so this adds a `w:comment` and a `w15:paraIdParent`.
   */
  replyComment(parentId: number, text = ""): number | undefined {
    const parent = this.doc.comments.find((c) => c.id === parentId);
    if (!parent) return undefined;
    const id = this.doc.nextCommentId();
    this.snapshot(this.caretPath());
    this.doc.addComment(id, text, this.authorName, this.authorInitials, parentId);
    this.render();
    this.hooks.changed();
    return id;
  }

  /** WPS's "删除所有批注": every comment and every anchor goes, the body text stays. */
  deleteAllComments(): number {
    const ids = this.doc.comments.map((c) => c.id);
    if (!ids.length) return 0;
    this.snapshot(this.caretPath());
    for (const id of ids) {
      this.dropAnchors(id);
      this.bodyAnchored.delete(id);
    }
    this.doc.comments = [];
    this.render();
    this.hooks.changed();
    return ids.length;
  }

  /**
   * WPS's "&Next Comment" / "&Previous Comment": steps through the list in the order the
   * reference marks are numbered, wrapping at both ends.
   */
  stepComment(step: 1 | -1): number | undefined {
    const list = this.listComments();
    if (!list.length) return undefined;
    const current = list.findIndex((entry) => entry.id === this.activeComment);
    const at = current < 0 ? (step === 1 ? 0 : list.length - 1) : (current + step + list.length) % list.length;
    const target = list[at];
    this.revealComment(target.id);
    return target.id;
  }

  /**
   * WPS exposes `PrintComments` and `ShowCommentShading`: printing normally drops the
   * annotations, and this puts them back. Only the print stylesheet reads the flag, so the
   * screen appearance is untouched.
   */
  setPrintComments(show: boolean): void {
    document.documentElement.classList.toggle("solo-print-comments", show);
  }

  private placeAnchors(path: number[], offset: number, anchors: CommentAnchor[]): boolean {
    const paragraph = this.paragraphAt(path);
    if (!paragraph) return false;
    const [head, tail] = splitAt(paragraph.inlines, offset);
    paragraph.inlines = merge([...head, ...anchors.map((anchor) => ({ text: "", anchor })), ...tail]);
    return true;
  }

  private dropAnchors(id: number): void {
    const walk = (blocks: Block[]): void => {
      for (const block of blocks) {
        if (block.kind === "paragraph") block.inlines = block.inlines.filter((i) => i.anchor?.id !== id);
        else if (block.kind === "table") for (const row of block.rows) for (const cell of row) walk(cell.paragraphs);
      }
    };
    walk(this.doc.blocks);
  }

  /** Fold pending DOM text into the model before a save reads it. */
  commit(): void {
    this.hfEditing?.blur();
    this.sync();
  }

  // ------------------------------------------------------- find / replace

  /** Every paragraph in body order, table cells included and a nested table walked through its own cells. */
  private paragraphPaths(): number[][] {
    const out: number[][] = [];
    const walk = (blocks: Block[], prefix: number[]): void => {
      blocks.forEach((block, i) => {
        const path = [...prefix, i];
        if (block.kind === "paragraph") out.push(path);
        else if (block.kind === "table") block.rows.forEach((row, r) => row.forEach((cell, c) => walk(cell.paragraphs, [...path, r, c])));
      });
    };
    walk(this.doc.blocks, []);
    return out;
  }

  matches(query: string): DocPath[] {
    if (!query) return [];
    const out: DocPath[] = [];
    for (const path of this.paragraphPaths()) {
      const paragraph = this.paragraphAt(path);
      if (!paragraph) continue;
      const text = paragraph.inlines.map((i) => i.text).join("");
      for (let at = text.indexOf(query); at >= 0; at = text.indexOf(query, at + query.length)) out.push({ path, offset: at });
    }
    return out;
  }

  focus(at: DocPath, length: number): void {
    const el = this.elementAt(at.path);
    if (!el) return;
    this.selectRange(at, { path: at.path, offset: at.offset + length });
  }

  replaceAt(at: DocPath, length: number, text: string): boolean {
    const paragraph = this.paragraphAt(at.path);
    if (!paragraph) return false;
    this.snapshot(at);
    const before = this.trackRevisions ? this.bodyText() : undefined;
    paragraph.inlines = spliceText(paragraph.inlines, at.offset, at.offset + length, text);
    const struck = before === undefined ? 0 : this.recordRevisions(before).struck;
    this.render();
    // The struck text sits where the old text was, so the caret steps past it to stay after the change.
    this.setCaret({ path: at.path, offset: at.offset + struck + text.length });
    this.hooks.changed();
    return true;
  }
}

/** Paragraphs inherit the shape of the one before them — Word's rule for Enter mid-document. */
function fresh(previous: Paragraph | undefined): Paragraph {
  if (!previous) return { kind: "paragraph", inlines: [] };
  const empty = previous.inlines.every((i) => i.text.trim() === "" && !isMark(i));
  return {
    kind: "paragraph",
    styleId: previous.styleId,
    styleName: previous.styleName,
    align: previous.align,
    numPr: empty ? undefined : previous.numPr,
    indentLeft: previous.indentLeft,
    spacingBefore: previous.spacingBefore,
    spacingAfter: previous.spacingAfter,
    lineSpacing: previous.lineSpacing,
    inlines: [],
  };
}

function listKind(doc: WordDocument, numId: number): "bullet" | "number" {
  return doc.levelOf(numId, 0)?.numFmt === "bullet" ? "bullet" : "number";
}/**
 * Where the breaks inside each stretch of paragraphs sit, as offsets in that stretch's own text. The
 * break that ends the last paragraph of a stretch is left out: it closes the body or the cell, and no
 * edit joins it away, so counting it would invent a revision at every table.
 */
function groupBreaks(spans: { length: number; group: number }[]): Map<number, number[]> {
  const out = new Map<number, number[]>();
  let offset = 0;
  let previous: number | undefined;
  for (let i = 0; i < spans.length; i++) {
    const span = spans[i];
    if (previous !== undefined && span.group !== previous) offset = 0;
    previous = span.group;
    const list = out.get(span.group) ?? [];
    out.set(span.group, list);
    offset += span.length;
    if (i + 1 < spans.length && spans[i + 1].group === span.group) list.push(offset);
  }
  return out;
}

const spansInGroup = <T extends { group: number }>(spans: T[], group: number): T[] =>
  spans.filter((s) => s.group === group);

/**
 * The paragraph breaks a text edit ate along with its characters: those that sat strictly inside the
 * replaced region, in the order they sat in the text, in the whole-body coordinates the diff uses. A
 * break at either edge survived, because the selection stopped there; one across a table wall belongs
 * to a stretch that moved, and a move is nobody's revision.
 */
function swallowedBreaks(
  before: BodyText,
  at: number,
  regionEnd: number,
): { at: number; group: number; upper: Paragraph; prior: Paragraph }[] {
  const out: { at: number; group: number; upper: Paragraph; prior: Paragraph }[] = [];
  const spans = before.spans;
  let offset = 0;
  for (let i = 0; i + 1 < spans.length; i += 1) {
    const upper = spans[i]!;
    offset += upper.length;
    const lower = spans[i + 1]!;
    if (lower.group !== upper.group) continue;
    if (at < offset && offset < regionEnd) out.push({ at: offset, group: upper.group, upper: upper.paragraph, prior: lower.paragraph });
  }
  return out;
}

/**
 * Every break either reading has, in the order they sit in the text, and what to do about each: a break
 * only the page has was inserted, one only the model has was removed, one both have stands as it was.
 * Two empty paragraphs at the same offset are two breaks, so the count of them there tells them apart.
 * The end of the stretch comes last and is nobody's revision.
 */
function mergedEnds(
  was: number[],
  is: number[],
  total: number,
): { at: number; occ: number; kind?: "ins" | "del" }[] {
  const key = (at: number, n: number): string => `${at}#${n}`;
  const tag = (list: number[]): Map<string, number> => {
    const seen = new Map<number, number>();
    const out = new Map<string, number>();
    for (const at of list) {
      const n = seen.get(at) ?? 0;
      seen.set(at, n + 1);
      out.set(key(at, n), at);
    }
    return out;
  };
  const before = tag(was);
  const after = tag(is);
  const offsets = [...new Set([...before.keys(), ...after.keys()])].sort(
    (x, y) => Number(x.slice(0, x.indexOf("#"))) - Number(y.slice(0, y.indexOf("#"))) || Number(x.slice(x.indexOf("#") + 1)) - Number(y.slice(y.indexOf("#") + 1)),
  );
  const ends = offsets.map((k) => ({
    at: Number(k.slice(0, k.indexOf("#"))),
    occ: Number(k.slice(k.indexOf("#") + 1)),
    kind: !before.has(k) ? ("ins" as const) : !after.has(k) ? ("del" as const) : undefined,
  }));
  ends.push({ at: total, occ: 0, kind: undefined });
  return ends;
}

/** The paragraph that ends at an offset, counting which of them it is: two empty paragraphs end at one
 *  offset, and only the one the edit touched carries the record a cancelling edit takes back. */
function endsAt(spans: { paragraph: Paragraph; length: number }[]): (at: number, occ: number) => Paragraph | undefined {
  const by = new Map<string, Paragraph>();
  const seen = new Map<number, number>();
  let offset = 0;
  for (const span of spans) {
    offset += span.length;
    const n = seen.get(offset) ?? 0;
    seen.set(offset, n + 1);
    by.set(`${offset}#${n}`, span.paragraph);
  }
  return (at: number, occ: number) => by.get(`${at}#${occ}`);
}

/** The paragraph objects that begin at an offset, handed out one per visit: two empty paragraphs begin
 *  at the same offset and are not the same paragraph, and each of them has to be found again. */
function startsAt(spans: { paragraph: Paragraph; length: number }[]): (at: number) => Paragraph | undefined {
  const by = new Map<number, Paragraph[]>();
  let offset = 0;
  for (const span of spans) {
    const list = by.get(offset);
    if (list) list.push(span.paragraph);
    else by.set(offset, [span.paragraph]);
    offset += span.length;
  }
  const used = new Map<number, number>();
  return (at: number) => {
    const list = by.get(at);
    if (!list) return undefined;
    const n = used.get(at) ?? 0;
    if (n >= list.length) return undefined;
    used.set(at, n + 1);
    return list[n];
  };
}

function sameList<T>(a: T[], b: T[]): boolean {
  return a.length === b.length && a.every((x, i) => x === b[i]);
}

/** Reuses the document's own numbering definitions; there is nothing to invent one from. */
function listNumId(doc: WordDocument, kind: "bullet" | "number" | "none"): number | undefined {
  if (kind === "none") return undefined;
  for (const entry of doc.numberingTable()) {
    if (listKind(doc, entry.numId) === kind) return entry.numId;
  }
  return undefined;
}

/** Cuts a run list at a flat caret offset, keeping styles on both sides. */
function splitAt(inlines: Inline[], at: number): [Inline[], Inline[]] {
  if (at <= 0) return [[], inlines.map((i) => ({ ...i }))];
  const left: Inline[] = [];
  const right: Inline[] = [];
  let seen = 0;
  for (const inline of inlines) {
    const length = inlineLength(inline);
    if (seen >= at) {
      right.push({ ...inline });
      continue;
    }
    if (seen + length <= at) {
      left.push({ ...inline });
      seen += length;
      continue;
    }
    const cut = at - seen;
    if (inline.text !== "") {
      left.push({ ...inline, text: inline.text.slice(0, cut) });
      right.push({ ...inline, text: inline.text.slice(cut) });
    } else {
      left.push({ ...inline });
    }
    seen = at;
  }
  return [left, right];
}

function patched(inline: Inline, patch: StylePatch): Inline {
  const out: Inline = { ...inline };
  for (const [key, value] of Object.entries(patch) as [keyof StylePatch, RunStyle[keyof RunStyle]][]) {
    if (value === undefined) continue;
    if (typeof value === "string" && value === "") delete out[key];
    else (out as unknown as Record<string, unknown>)[key] = value;
  }
  return out;
}

function restyle(inlines: Inline[], from: number, to: number, patch: StylePatch, sign?: (before: Inline) => FormatRevision | undefined): Inline[] {
  const [head, rest] = splitAt(inlines, from);
  const [middle, tail] = splitAt(rest, to === Infinity ? Infinity : to - from);
  const changed = middle.map((i) => {
    const next = patched(i, patch);
    // Nothing actually changed, so nothing gets signed: a toolbar toggle that lands on the formatting
    // the run already has must not leave a revision nobody asked for.
    if (!sign || sameRunProps(i, next)) return next;
    const record = sign(i);
    return record ? { ...next, revFormat: record } : next;
  });
  return merge([...head, ...changed, ...tail]);
}

/** Replaces a flat character range with new text, which takes the style it lands in. */
function spliceText(inlines: Inline[], from: number, to: number, text: string): Inline[] {
  const [head, rest] = splitAt(inlines, from);
  const [middle, tail] = splitAt(rest, Math.max(0, to - from));
  const source = middle.find((i) => i.text !== "") ?? head[head.length - 1];
  const inserted: Inline[] = [];
  if (text !== "") inserted.push({ ...(source ? pickTextStyle(source) : {}), text });
  return merge([...head, ...inserted, ...tail]);
}

/**
 * The style typed text takes from the run it lands in. The property set comes along with it, because what the
 * cursor sits in is a run with a font, a language and a style the nine modelled fields do not add up to.
 */
function pickTextStyle(inline: Inline): Partial<Inline> {
  const style: Partial<Inline> = {};
  for (const key of CARRIED) {
    if (key === "break" || key === "tab") continue;
    const value = inline[key];
    if (value !== undefined) (style as Record<string, unknown>)[key] = value;
  }
  if (inline.rPrSlots) {
    style.rPrSlots = inline.rPrSlots;
    style.asRead = inline.asRead;
  }
  return style;
}

function shape(blocks: Block[]): string {
  return blocks.map((b) => (b.kind === "paragraph" ? "p" : b.kind === "table" ? `t${b.rows.length}x${b.columns}` : "c")).join(",");
}

function forEachInline(blocks: Block[], doIt: (inline: Inline) => void): void {
  for (const block of blocks) {
    if (block.kind === "paragraph") {
      for (const inline of block.inlines) doIt(inline);
      continue;
    }
    if (block.kind !== "table") continue;
    for (const row of block.rows) for (const cell of row) forEachInline(cell.paragraphs, doIt);
  }
}

/** Which comments the body still holds a marker for. */
function anchorIdsIn(blocks: Block[]): Set<number> {
  const out = new Set<number>();
  forEachInline(blocks, (i) => {
    if (i.anchor) out.add(i.anchor.id);
  });
  return out;
}

function countAnchors(blocks: Block[]): number {
  let n = 0;
  forEachInline(blocks, (i) => {
    if (i.anchor) n += 1;
  });
  return n;
}

function markersOf(blocks: Block[]): number {
  let n = 0;
  for (const b of blocks) {
    if (b.kind === "table") b.rows.forEach((row) => row.forEach((c) => (n += markersOf(c.paragraphs))));
    else if (b.kind === "paragraph" && b.numPr) n += 1;
  }
  return n;
}

/** The element a carried container is drawn by: read-only, and named after what it holds. */
function carryEl(block: Carried): HTMLElement {
  const el = document.createElement("div");
  el.className = CARRY_CLASS;
  el.contentEditable = "false";
  el.dataset.tag = block.tag;
  el.setAttribute("style", CARRY_STYLE);
  el.textContent = block.text;
  return el;
}

/**
 * A container the page lost, put back in the list and on the page at the place it held. Nothing on the page
 * can edit one, so its element only goes missing because the browser discarded it while reflowing around a
 * keystroke — and a structure the file wrote must not vanish for that. Both halves get it back because the
 * model and the page are stamped against each other by position: a block with no element under it would
 * slide every block after it onto the wrong paragraph.
 */
function keepCarried(held: Block[], next: Block[], shown: Element[], put: (el: HTMLElement, before: Element | undefined) => void): void {
  held.forEach((block, i) => {
    if (block.kind !== "carried" || next.includes(block)) return;
    const following = held.slice(i + 1).find((b) => next.includes(b));
    const at = following === undefined ? next.length : next.indexOf(following);
    const el = carryEl(block);
    put(el, shown[at]);
    next.splice(at, 0, block);
    shown.splice(at, 0, el);
  });
}

/** Every carried inline structure in the model, by the key the page draws into its marker. */
function carryByKeys(blocks: Block[]): Map<number, Inline> {
  const out = new Map<number, Inline>();
  forEachInline(blocks, (inline) => {
    if (inline.carry) out.set(inline.carry.key, inline);
  });
  return out;
}

/**
 * A carry whose marker element the browser swallowed — a zero-width span merged away by typing beside it —
 * goes back into the paragraph it came from. Its place there is the end rather than where it was, because
 * nothing about a rebuilt paragraph keeps a zero-width offset: a bookmark at the wrong end of a sentence is
 * still the document's bookmark.
 */
function keepCarries(prev: Inline[], next: Inline[], placed: Set<number>): void {
  for (const inline of prev) {
    if (!inline.carry || placed.has(inline.carry.key) || next.includes(inline)) continue;
    next.push(inline);
  }
}

/** Enter and Backspace can leave bare text under a block element; it belongs to a paragraph. */
function wrapStrays(el: Element): void {
  for (const child of [...el.childNodes]) {
    if (child.nodeType === Node.TEXT_NODE && (child.nodeValue ?? "").trim() !== "") {
      const div = document.createElement("div");
      div.className = "para";
      el.insertBefore(div, child);
      div.append(child);
    }
  }
}

/** Natural pixel size of an image, with a fallback when the browser cannot decode it. */
function imageSize(url: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 320, height: img.naturalHeight || 240 });
    img.onerror = () => resolve({ width: 320, height: 240 });
    img.src = url;
  });
}

function closestTag(node: Node | null, tag: string): HTMLElement | null {
  let el: Element | null = node instanceof Element ? node : node?.parentElement ?? null;
  while (el && el.tagName !== tag) el = el.parentElement;
  return el as HTMLElement | null;
}

function closestClass(node: Node | null, name: string): HTMLElement | null {
  let el: Element | null = node instanceof Element ? node : node?.parentElement ?? null;
  while (el && !el.classList.contains(name)) el = el.parentElement;
  return el as HTMLElement | null;
}

function imageOf(el: Element): Inline["image"] {
  const img = el as HTMLImageElement;
  const path = el instanceof HTMLElement ? el.dataset.path : undefined;
  if (!path) return undefined;
  return {
    src: img.getAttribute("src") ?? "",
    width: Number(img.getAttribute("width")) || img.naturalWidth || 160,
    height: Number(img.getAttribute("height")) || img.naturalHeight || 90,
    name: img.getAttribute("alt") ?? "picture",
    path,
  };
}

/** The anchor a rendered marker element stands for, or undefined when it carries none. */
function anchorOf(el: Element): CommentAnchor | undefined {
  const id = Number(el instanceof HTMLElement ? el.dataset.cid : NaN);
  const kind = el instanceof HTMLElement ? el.dataset.kind : undefined;
  if (!Number.isFinite(id) || (kind !== "start" && kind !== "end" && kind !== "ref")) return undefined;
  return { id, kind };
}

function modelChars(node: Node): number {
  if (node.nodeType === Node.TEXT_NODE) return (node.nodeValue ?? "").length;
  if (!(node instanceof Element) || node.classList.contains("bullet")) return 0;
  if (node.tagName === "BR" || node.tagName === "IMG" || node.classList.contains("tab") || node.classList.contains("cmt-anchor")) return 1;
  let n = 0;
  for (const child of node.childNodes) n += modelChars(child);
  return n;
}

function modelOffset(root: Element, node: Node, offset: number): number {
  let count = 0;
  const walk = (parent: Node): boolean => {
    if (parent === node) {
      const kids = [...parent.childNodes];
      for (let i = 0; i < Math.min(offset, kids.length); i += 1) count += modelChars(kids[i]);
      return true;
    }
    for (const child of parent.childNodes) {
      if (child === node && child.nodeType === Node.TEXT_NODE) {
        count += offset;
        return true;
      }
      if (child.nodeType === Node.TEXT_NODE) {
        count += modelChars(child);
        continue;
      }
      if (!(child instanceof Element)) continue;
      if (child.classList.contains("bullet")) continue;
      if (child.contains(node)) return walk(child);
      count += modelChars(child);
    }
    return false;
  };
  walk(root);
  return count;
}

/** One DOM position per model character, so a caret offset survives a re-render exactly. */
function positions(root: Element): { node: Node; offset: number }[] {
  const out: { node: Node; offset: number }[] = [];
  const walk = (parent: Node): void => {
    for (const child of parent.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.nodeValue ?? "";
        for (let i = 0; i < text.length; i += 1) out.push({ node: child, offset: i });
        continue;
      }
      if (!(child instanceof Element) || child.classList.contains("bullet")) continue;
      if (child.tagName === "BR" || child.tagName === "IMG" || child.classList.contains("tab") || child.classList.contains("cmt-anchor")) {
        out.push({ node: parent, offset: [...parent.childNodes].indexOf(child) });
        continue;
      }
      walk(child);
    }
  };
  walk(root);
  const last = lastTextNode(root);
  if (last) out.push({ node: last, offset: (last.nodeValue ?? "").length });
  return out;
}

function lastTextNode(root: Element): Text | null {
  let found: Text | null = null;
  const walk = (parent: Node): void => {
    for (const child of parent.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) found = child as Text;
      // The badge's own digit is a label, not a place to put the caret.
      else if (child instanceof Element && !child.classList.contains("bullet") && !child.classList.contains("cmt-anchor")) walk(child);
    }
  };
  walk(root);
  return found;
}
