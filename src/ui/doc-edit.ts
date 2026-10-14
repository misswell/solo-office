import type { Block, Carried, CarriedInline, FormatRevision, HfSegment, Inline, ParaStyle, PageGeometry, PageSetup, Paragraph, ParagraphFormatRevision, PropSlot, Revision, RevisionEntry, RowInfo, RunStyle, Table, TableCell, WordDocument } from "../core/docx.js";
import type { CommentAnchor, DocComment } from "../core/comments.js";
import { bodyShape, commentTexts } from "../core/comments.js";
import { DEFAULT_SECTION, RUN_STYLE_KEYS, countTexts, ownParaStyle, ownRunStyle, paraStyleOf, paragraphTextsOf, runStyleOf, sectionEndsOf, stampNow, syncParaSlots, twipsToPx } from "../core/docx.js";
import type { CellWhere, TableEdit } from "../core/doc-table-ops.js";
import { cellRangeFor, cellRect, cellsInRect, cellsInSweep, deleteColumnHere, deleteRowHere, gridColumns, insertCellHere, insertColumnHere, insertRowHere, mergeHere, pasteAppendTableHere, pasteAsNestedTableHere, splitHere, splitTableAt } from "../core/doc-table-ops.js";
import { IMAGE_EXT_BY_MIME } from "../core/media.js";
import { CARRY_CLASS, CARRY_STYLE, PAGE_BORDER_CLASS, RUN_CARRY_CLASS, highlightColor, renderWord } from "../core/render.js";
import { flatOf, inlineLength, isMark, keepRunIdentity } from "../core/run-identity.js";
import { CommentPanel } from "./doc-comments.js";
import { lineCaretAnchor, lineNumberAt, measurePageLayout, pageAtCaret, pageCaretAnchor, pagePositionAt, rowColumnAt, sectionCaretAnchor, type CaretLine, type PageAnchor, type PageMeasure, type PagePlacement } from "./doc-layout.js";
import { ReviewPanel } from "./doc-review.js";
import { OutlinePanel } from "./doc-outline.js";
import { openPageSetup as showPageSetupDialog } from "./doc-page-setup.js";
import { openWordCount, type WordCountReadout } from "./doc-word-count.js";
import { openGoTo } from "./doc-goto.js";

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

/**
 * The names this page can bring back off a CSS colour: the seven the toolbar offers, plus §17.17.64's
 * cancellation, which draws `transparent` and has to be read as a sentence rather than as an absence. It is a
 * fallback only — the renderer stamps the document's own name as `data-hl`, and every name outside this list
 * (`blue`, `darkGray`, a face the ribbon has never offered) comes back from that attribute instead of from a
 * reverse lookup that could not have named it.
 */
const HIGHLIGHT_NAMES = ["yellow", "green", "cyan", "magenta", "red", "darkYellow", "lightGray", "none"];

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

/** The zero-width space the pending cancellation stands in for: a character no run may ever keep. */
const ZWSP = "\u200b";

/**
 * How WPS spells a soft return: `w:type="textWrapping"`, not a bare `w:br`. Measured over the 86 files this
 * machine's WPS wrote — 132 breaks state `type="textWrapping"` with no `w:clear` and 9 add `w:clear="all"`, while
 * **0 are spelled without `w:type`**. A break the page created itself (Shift+Enter, which Blink inserts as a bare
 * `<br>`) therefore comes back with this token; a break the file authored keeps the token it arrived with.
 */
const TEXTWRAPPING_BR = '<w:br w:type="textWrapping"/>';

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
  // Which of the two the page hands back is the fresher claim: the colour we can name out of the inline style is
  // the one Blink wrote *this time*, while `data-hl` is an attribute it copies along when it splits a span. At a
  // collapsed caret inside a highlighted run, 工具栏 changing the colour rewrites the style (`background-color:
  // rgb(0, 255, 0)`) and leaves the copied name (`yellow`) behind — reading the stamp first made the change land
  // as no change at all, in the model and in the saved file. So the colour wins when this page can name it, and
  // the stamp is the fallback for the names CSS cannot carry back (§17.17.64 has 17 members, the ribbon offers 8,
  // and `blue` round-trips through CSS only as the stamp).
  const bg = inline.backgroundColor || inline.background;
  const fromCss = bg ? HIGHLIGHT_BY_CSS.get(cssValue("background", bg)) : undefined;
  const hl = fromCss || el.getAttribute("data-hl");
  if (hl) s.highlight = hl;
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

type Context = RunStyle & Pick<Inline, "hyperlink" | "linkAttrs"> & { rev?: Revision; revFormat?: FormatRevision };

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

/** The run fields an edit carries from the run it splits or joins. Derived from `RUN_STYLE_KEYS` so a new
 *  property cannot be read, drawn and saved while a split quietly drops it. */
const CARRIED: (keyof Inline)[] = [...RUN_STYLE_KEYS, "hyperlink", "linkAttrs", "break", "brXml", "tab"];

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

const charLength = (inlines: Inline[]): number => inlines.reduce((n, i) => n + inlineLength(i), 0);

/**
 * 光标（`node`/`offset` 处那枚**折叠** Range）画出来的矩形**中线**，px。
 * Chrome 给折叠 Range 一枚零宽、整行高的矩形；中线而不是上沿是给 `pageAtCaret` 用的：光标矩形的上沿是字体盒
 * 的上沿，与量出来的行盒跨度差零点几像素且方向不定，正好落在相邻两枚相接的缝上（`doc-layout.ts` 文件头）。
 * 一枚矩形都拿不到时退回它的包围盒，包围盒也是空的话退回所在元素的盒子（这种退路上只有盒子中心可用）。
 * 都拿不到时 `undefined`——那时状态栏宁可保留上一次的读数，也不报一个凭空算出来的页。
 */
function caretMidOf(node: Node, offset: number): number | undefined {
  const range = document.createRange();
  try {
    range.setStart(node, offset);
    range.collapse(true);
  } catch {
    return undefined;
  }
  const own = range.getClientRects()[0];
  if (own) return own.top + own.height / 2;
  const box = range.getBoundingClientRect();
  if (box.width || box.height || box.top || box.left) return box.top + box.height / 2;
  const el = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
  const rect = el?.getBoundingClientRect();
  return rect ? rect.top + rect.height / 2 : undefined;
}

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
  /**
   * An undo step holds the text, the comments and the page. `blocks` carries a paragraph's own section break,
   * but the body's `w:sectPr` and the geometry the page is drawn from sit outside it, so a step that cloned
   * only the text would step a margin change back into a page that never moved.
   */
  private undoStack: { blocks: Block[]; comments: DocComment[]; caret: DocPath | null; page: PageGeometry }[] = [];
  private redoStack: { blocks: Block[]; comments: DocComment[]; caret: DocPath | null; page: PageGeometry }[] = [];
  private burst = 0;
  private composing = false;
  /**
   * 悬光标待用的那一次格式，现在举着「突出显示」和「字号」两族——这两族 Blink 都交不出来。高亮：折叠光标处
   * `execCommand("hiliteColor", false, "transparent")` 返回 `true` 却是空操作，随后打的字照旧坐在原有的/借来的底色里
   * （第 2 问）；它带得动颜色，但那是另一张嘴：新字被包进它自己的 span，我们举着的 span 盖不住它（第 7 问：待用的「无」
   * 之后改点「绿」，打出来是绿）。字号：`fontSize` 只吃 1–7 那七个枚举，22 磅根本没有对应的命令，悬光点点字号从前是
   * 什么都不发生。所以这条通道由编辑器自己举一个 span，把待用的 CSS 写在它身上：同族再点换掉自己那一项（最后说话
   * 的那次算），跨族往上叠（先「黄」后「22」，打出来的字两样都有）。span 里那个零宽空格是撑位置的——空的 span Blink
   * 直接跨过去，字落在左边的文本节点里（第 6 问）。占位字符绝不能吃进文档：每次回折模型前由 `settlePending` 摘掉。
   */
  private pendingFormat: HTMLElement | null = null;
  /**
   * 占位 span 上那一次点过的格式，攒着的样子（不是最后一次的）：同族再点换掉自己那一项，别族点到往上叠。
   * span 是它的投影，`settlePending` 一收就清空。
   */
  private pendingFormatPatch: StylePatch = {};
  private panel: CommentPanel | null = null;
  /** 审阅窗格: the list of every change still waiting, docked beside the page. */
  private review: ReviewPanel | null = null;
  /** 导航窗格: the headings the document's outline levels name, docked on the left. */
  private outline: OutlinePanel | null = null;
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
    private readonly hooks: { changed(): void; moved?(): void },
  ) {
    for (const id of anchorIdsIn(this.doc.blocks)) this.bodyAnchored.add(id);
    this.revSeq = this.doc.maxRevisionId();
    injectRevisionStyles();
    this.panel = new CommentPanel(this.host, this);
    this.review = new ReviewPanel(this.host, this);
    this.outline = new OutlinePanel(this.host, this);
    document.addEventListener("selectionchange", () => this.trackSelection());
    this.render();
  }

  render(): void {
    this.host.innerHTML = renderWord(this.doc);
    this.page = this.host.querySelector<HTMLElement>(".page");
    this.panel?.mount();
    this.review?.mount();
    this.outline?.mount();
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
    for (const btn of this.host.querySelectorAll<HTMLElement>("[data-hf-page]")) {
      // The band opens for editing on a click anywhere in it; this button is a command, not the caret.
      btn.addEventListener("mousedown", (e) => e.preventDefault());
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        this.insertPageNumber(btn.dataset.hfPage === "header" ? "header" : "footer");
      });
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
        this.commitHF(el, which);
      },
      { once: true },
    );
  }

  /**
   * The band's paragraphs as the model holds them: every direct child of the band is one paragraph — what
   * the page writes is `.hf-p`, and what the browser adds when the user presses Enter is a bare `div`, which
   * is the same paragraph by a different name and must not be read as nothing — and inside it the text the
   * page shows, with a field coming back by its key rather than by the label drawn for it. The label is for
   * reading; the field's bytes are for the file, and the number it cached is neither.
   */
  private readHF(el: HTMLElement, which: "header" | "footer"): { segs: HfSegment[][]; caret: number } {
    const held = (which === "header" ? this.doc.header : this.doc.footer)?.segments ?? [];
    const atoms = new Map<number, CarriedInline>();
    for (const segs of held) for (const s of segs) if (s.carry) atoms.set(s.carry.key, s.carry);
    const placed = new Set<number>();
    const caret = document.getSelection()?.anchorNode ?? null;
    const segs: HfSegment[][] = [];
    let caretAt = -1;
    for (const child of [...el.children]) {
      if (child.classList.contains("hf-bar")) continue;
      const out: HfSegment[] = [];
      let text = "";
      const flush = (): void => {
        if (text === "") return;
        out.push({ text });
        text = "";
      };
      for (const node of child.childNodes) {
        if (node.nodeType === Node.TEXT_NODE) {
          text += (node.textContent ?? "").replace(/\u00a0/g, " ");
          continue;
        }
        const c = node as HTMLElement;
        if (c.classList?.contains(RUN_CARRY_CLASS)) {
          const key = Number(c.getAttribute("data-carry"));
          const atom = Number.isFinite(key) ? atoms.get(key) : undefined;
          if (!atom || placed.has(key)) continue;
          flush();
          placed.add(key);
          out.push({ carry: atom });
          continue;
        }
        text += (c.textContent ?? "").replace(/\u00a0/g, " ");
      }
      flush();
      if (caret && child.contains(caret)) caretAt = segs.length;
      segs.push(out);
    }
    return { segs, caret: caretAt };
  }

  /** Writes the band back into the model when what it holds has changed. */
  private commitHF(el: HTMLElement, which: "header" | "footer"): void {
    const before = which === "header" ? this.doc.header : this.doc.footer;
    if (!before) return;
    const { segs } = this.readHF(el, which);
    if (JSON.stringify(segs) === JSON.stringify(before.segments ?? before.paragraphs.map((t) => [{ text: t }]))) return;
    if (!this.doc.setHeaderSegments(which, segs)) return;
    this.hooks.changed();
  }

  /**
   * Puts a page number in the band, keeping whatever the user had just typed there. The page is drawn again
   * afterwards and the band opened, because the field is one thing the caret cannot be put inside of.
   */
  private insertPageNumber(which: "header" | "footer"): void {
    const el = this.host.querySelector<HTMLElement>(`.hf-${which}`);
    if (!el) return;
    const { caret } = this.readHF(el, which);
    this.commitHF(el, which);
    if (!this.doc.insertPageNumber(which, caret)) {
      this.render();
      return;
    }
    this.render();
    const again = this.host.querySelector<HTMLElement>(`.hf-${which}`);
    if (again) this.editHF(again);
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
      const dir = e.shiftKey ? -1 : 1;
      // 列表项里 Tab 那一族说的是层级：普通 Tab 只在段落开头降一级（多缩一层），Shift+Tab 随时升一级。
      // 表格里不碰——那面 Tab 是走格子的事。而那一级文件自己没说就什么都不动：`w:lvl` 是文件写的形状，
      // 页面不能替它编一条它没写的（补出的那条只有第 0 级，所以在一张新纸上 Tab 仍是插入制表符）。
      const caret = this.caretPath();
      const atStart = dir < 0 || (document.getSelection()?.isCollapsed === true && caret?.offset === 0);
      if (atStart && !this.tableAt(caret) && this.shiftListLevel(dir)) {
        e.preventDefault();
        return;
      }
      e.preventDefault();
      this.snapshot(caret);
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
    // 待用的格式只是光标底下的一件事：光标走到它外面，它就没被用过，占位符不能留在页面里。
    const pending = this.pendingFormat;
    if (pending?.isConnected) {
      const sel = document.getSelection();
      const at = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
      if (at && !pending.contains(at.startContainer)) this.settlePending();
    }
    const ends = this.rangeEnds();
    if (!ends) return;
    // The ribbon reads the caret, so it is told whenever the caret moves inside the page — selecting is not an
    // edit and does not go through `changed()`, and a ribbon that only refreshes after an edit keeps showing
    // the state of the last click rather than of the run the reader is pointing at.
    this.hooks.moved?.();
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

  /**
   * The CSS the pending format stands in for, in the shape `styleOf` reads back off the page. 高亮 needs both faces:
   * the background for the names CSS can carry, `data-hl` for the ones it cannot (§17.17.64 has 17 members, CSS
   * cannot name `blue`). 字号 needs only the size, and the writer spells the pair the file keeps (`w:sz`/`w:szCs`)
   * off the one field. A key this model has no place for holds nothing.
   */
  private pendingStyle(patch: StylePatch): string {
    const parts: string[] = [];
    if (patch.highlight !== undefined) parts.push(`background:${highlightColor(patch.highlight || "none")};`);
    if (patch.sizePt !== undefined && Number.isFinite(patch.sizePt)) parts.push(`font-size:${patch.sizePt}pt;`);
    return parts.join("");
  }

  /**
   * Stand the pending format up under the caret. Blink's own pending format is not usable here for these two
   * families: `hiliteColor` with `transparent` returns true and does nothing at a collapsed caret, and when it *does*
   * carry a colour it wraps the new character in a span of its own — which beats whatever the caret was sitting in,
   * so clicking 黄 then 无 types yellow, the earlier of the two clicks. 字号 has no command at all (`fontSize` takes
   * the 1-7 enum, so 22 磅 was simply dropped). Carrying the pending format ourselves makes the last click within a
   * family the one that counts, and keeps clicks across families stacked — Word leaves a pending highlight standing
   * when you then pick a size, so the next character gets both. The placeholder is what makes the span usable: Blink
   * steps over an empty one and lands the character in the text node to the left.
   */
  private armPending(patch: StylePatch): void {
    const next: StylePatch = { ...this.pendingFormatPatch, ...patch };
    const style = this.pendingStyle(next);
    this.settlePending();
    if (!style) return;
    const sel = document.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const at = sel.getRangeAt(0);
    if (!at.collapsed || !this.page?.contains(at.startContainer)) return;
    at.collapse(true);
    const span = document.createElement("span");
    if (next.highlight !== undefined) span.setAttribute("data-hl", next.highlight || "none");
    span.setAttribute("style", style);
    const holder = document.createTextNode(ZWSP);
    span.append(holder);
    at.insertNode(span);
    const caret = document.createRange();
    caret.setStart(holder, ZWSP.length);
    caret.setEnd(holder, ZWSP.length);
    sel.removeAllRanges();
    sel.addRange(caret);
    this.pendingFormat = span;
    this.pendingFormatPatch = next;
  }

  /**
   * The placeholder is not text: settle it before every fold back into the model and whenever the caret walks away.
   * A span that still holds only it was never used and goes; one the characters landed in keeps them and loses the
   * placeholder — the format now lives in the model and in what the renderer puts back, so this one pass is enough.
   * Either way the pending patch is spent: what gets typed next starts from the run the caret sits in again.
   */
  private settlePending(): void {
    const span = this.pendingFormat;
    this.pendingFormat = null;
    this.pendingFormatPatch = {};
    if (!span?.isConnected) return;
    if ((span.textContent ?? "") === ZWSP) {
      span.remove();
      return;
    }
    const holder = span.firstChild;
    if (holder instanceof Text && holder.textContent?.startsWith(ZWSP)) holder.deleteData(0, ZWSP.length);
  }

  /** Re-derives the model from the DOM, then repairs anything structural the browser changed. */
  private mutated(): void {
    if (!this.page) return;
    this.settlePending();
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
      // A page border is drawn into the paper box rather than written on it. It holds no text and is not a
      // block, so reading it back would turn the file's `w:pgBorders` into an empty paragraph in the body.
      if (el.classList.contains(PAGE_BORDER_CLASS)) continue;
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
      const paragraph: Paragraph = known?.kind === "paragraph" ? known : fresh(previous, this.doc);
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
          const paragraph: Paragraph = known?.kind === "paragraph" ? known : fresh(previous, this.doc);
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
          // Whatever the file wrote on the `w:br` comes back off the marker unchanged, so a keystroke in a
          // page-break paragraph does not quietly turn it into a plain line break. A `br` with no marker is one
          // this editor just made (Shift+Enter), and it gets WPS's own spelling rather than a bare `w:br`.
          out.push({ ...ctx, text: "", break: true, brXml: child.getAttribute("data-brk") || TEXTWRAPPING_BR });
          continue;
        }
        if (tag === "img") {
          const image = imageOf(child);
          if (image) out.push({ ...ctx, text: "", image });
          continue;
        }
        const next: Context = { ...ctx, ...styleOf(child) };
        if (tag === "a") {
          next.hyperlink = child.getAttribute("href") || undefined;
          // The attributes the link came in with ride on the anchor and have to come back off it, or a
          // keystroke inside the link drops them the same way it used to drop the link itself.
          next.linkAttrs = child.getAttribute("data-link-attrs") || undefined;
        }
        const rev = revisionOf(child);
        if (rev) next.rev = rev;
        const fmt = formatChangeOf(child);
        if (fmt) next.revFormat = fmt;
        walk(child, next);
      }
    };
    walk(el, {});
    const merged = merge(out);
    // Before anything is appended from the old list: the alignment below is between what the page shows and what
    // it was drawn from, and a carry the browser swallowed has no place in either.
    keepRunIdentity(merged, prev);
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
    if (node === this.page) return this.pathOfPageEdge(offset);
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

  /**
   * 端点落在编辑页本身上的那种选区：全选（⌘A）铺出来的 Range，起止容器就是这一张连续的纸，下标是块的下标，
   * 不是某个段里的字符数。把它们折回首/末那个块的边界，全选才和拖出来的选区一样能寻址——格式命令与
   * 「所选文字」都走 `rangeEnds()`，此前它们在全选时直接拿不到端点。
   */
  private pathOfPageEdge(offset: number): DocPath | null {
    if (!this.page) return null;
    const blocks = [...this.page.children];
    if (!blocks.length) return null;
    // 下标 k（k < 块数）指「第 k 块之前」，下标等于块数才指「最后一块之后」。
    const atEnd = offset >= blocks.length;
    const block = blocks[atEnd ? blocks.length - 1 : offset];
    if (!block) return null;
    const para = atEnd
      ? [...block.querySelectorAll(".para")].pop() ?? (block.classList.contains("para") ? block : undefined)
      : block.classList.contains("para")
        ? block
        : block.querySelector(".para");
    if (!para) return null;
    return this.pathOf(para, atEnd ? para.childNodes.length : 0);
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

  /**
   * 段落级那三条命令（左中右与两端对齐、项目符号／编号、增加／减少缩进量）要改的段落：折叠光标就只光标那一段，
   * 有选区就是选区盖住的**每一段**。跨段的取材与 `applyStyle` 改文字格式时是同一句 `rangeEnds()`、同一条走法
   * （首段、末段、顶层两者之间的那些块），所以「这一段被加粗了却没被居中」这种分裂不可能出现。
   */
  private selectionParagraphs(): Paragraph[] {
    const ends = this.rangeEnds();
    if (!ends) {
      const at = this.caretPath();
      const only = at ? this.paragraphAt(at.path) : undefined;
      return only ? [only] : [];
    }
    const first = this.paragraphAt(ends.start.path);
    if (ends.start.path.join() === ends.end.path.join()) return first ? [first] : [];
    const out: Paragraph[] = [];
    if (first) out.push(first);
    for (let i = ends.start.path[0] + 1; i < ends.end.path[0]; i += 1) {
      const block = this.doc.blocks[i];
      if (block?.kind === "paragraph") out.push(block);
    }
    const last = this.paragraphAt(ends.end.path);
    if (last && last !== first) out.push(last);
    return out;
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
    this.undoStack.push({
      blocks: structuredClone(this.doc.blocks),
      comments: structuredClone(this.doc.comments),
      caret: at,
      page: this.doc.pageGeometryState(),
    });
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
    this.redoStack.push({ blocks: structuredClone(this.doc.blocks), comments: structuredClone(this.doc.comments), caret: this.caretPath(), page: this.doc.pageGeometryState() });
    this.doc.blocks = step.blocks;
    this.doc.comments = step.comments;
    this.doc.restorePageGeometry(step.page);
    this.render();
    if (step.caret) this.setCaret(this.clamp(step.caret));
    this.hooks.changed();
    return true;
  }

  redo(): boolean {
    const step = this.redoStack.pop();
    if (!step) return false;
    this.undoStack.push({ blocks: structuredClone(this.doc.blocks), comments: structuredClone(this.doc.comments), caret: this.caretPath(), page: this.doc.pageGeometryState() });
    this.doc.blocks = step.blocks;
    this.doc.comments = step.comments;
    this.doc.restorePageGeometry(step.page);
    this.render();
    if (step.caret) this.setCaret(this.clamp(step.caret));
    this.hooks.changed();
    return true;
  }

  /** Ribbon command: 「页面设置」 for the page this document is drawn on. */
  openPageSetup(): void {
    showPageSetupDialog(this.doc.sections[0] ?? DEFAULT_SECTION, (set) => this.applyPageSetup(set));
  }

  /**
   * The numbers 「字数统计」 reports, read off the model as it stands right now. Both faces of the command
   * (the dialog and the 状态栏那一格) go through here, so they cannot disagree. 页数／行数 不在这一份里：那是
   * 排版之后的答案（WPS 自己也是用 `ComputeStatistics(wdStatisticPages)` 那一族问的），走 `layout()`。
   */
  counts(): WordCountReadout {
    return this.doc.wordCounts();
  }

  /**
   * 「字数统计」里 页数／行数 那两行的来路：把这张连续的纸按**每一节自己的**页高与上下边距量一遍，行盒从**画出来的盒子**里读
   * （口径、以及它为什么不能当成 WPS 引擎那个数，都写在 `doc-layout.ts` 文件头）。纸没画上、或某一节说的页高连边距
   * 都不够留时返回 `undefined`——那一行整个不报，而不是报一个量不出来的数。
   */
  layout(): PageMeasure | undefined {
    if (!this.page) return undefined;
    return measurePageLayout(this.page, this.doc);
  }

  /**
   * 状态栏 页码／页数 那两格（WPS 的 `KxWpsStatusPageNumberInfoCommand` 与 `KxWpsStatusPageAreaInfoCommand`）：
   * 光标（选区的活动端）落在第几页、一共几页。总页数与「字数统计」的 页数 出自同一趟 `pageAtCaret` 测量，
   * 光标→页的走法写在 `doc-layout.ts` 文件头。
   *
   * 光标不在纸面上时返回 `undefined`——点了批注面板的输入框、点了页眉页脚带，活动端就不在正文里了；
   * 那两格该保留上一次的读数（WPS 真机在那种时候怎么表现**未取证**），由状态栏自己决定，这里不替它编一个页号。
   */
  caretPage(): PagePlacement | undefined {
    if (!this.page) return undefined;
    const sel = document.getSelection();
    const node = sel?.focusNode ?? null;
    if (!node || !this.page.contains(node)) return undefined;
    const mid = caretMidOf(node, sel!.focusOffset);
    if (mid === undefined) return undefined;
    return pageAtCaret(this.page, this.doc, mid);
  }

  /**
   * 状态栏 行／列 那两格（WPS 的 `KxWpsStatusRowInfoCommand` 与 `KxWpsStatusColumnInfoCommand`，kuip 里
   * 靠 `rightSibling`/`leftSibling` 配对）：行是光标在**这一页里**排第几（口径见 `doc-layout.ts` 的
   * `rowInPage`），列是它在这一行的第几个字（段落才数得出；表格/只读容器不给列）。
   * 取点、判「光标在不在纸面上」与 `caretPage()` 同一条：两格与页码必须出自同一处光标。
   */
  caretLine(): CaretLine | undefined {
    if (!this.page) return undefined;
    const sel = document.getSelection();
    const node = sel?.focusNode ?? null;
    if (!node || !this.page.contains(node)) return undefined;
    const mid = caretMidOf(node, sel!.focusOffset);
    if (mid === undefined) return undefined;
    return rowColumnAt(this.page, this.doc, { mid, node, offset: sel!.focusOffset });
  }

  /**
   * 「定位」「行」目标那两行读数：光标在第几行（**绝对行号**——从文首连数、换页不重数，口径与走法见
   * `doc-layout.ts` 的 `lineNumberAt`）与全文共几行（`measurePageLayout` 的 `lines`，与「字数统计」那一页
   * 同一份测量）。与状态栏「行」格不同：那一格数的是**页内**序数（`rowInPage`），这里数的是全篇序数；
   * WPS 真机「行」按哪种数**未取证**。取点、判「光标在不在纸面上」与 `caretPage()` 同一条。
   */
  caretLineNumber(): { line: number; lines: number } | undefined {
    if (!this.page) return undefined;
    const sel = document.getSelection();
    const node = sel?.focusNode ?? null;
    if (!node || !this.page.contains(node)) return undefined;
    const mid = caretMidOf(node, sel!.focusOffset);
    if (mid === undefined) return undefined;
    const line = lineNumberAt(this.page, this.doc, { mid, node, offset: sel!.focusOffset });
    const lines = this.layout()?.lines;
    if (line === undefined || lines === undefined) return undefined;
    return { line, lines };
  }

  /**
   * 状态栏 节 那一格（WPS 的 `KxWpsStatusSectionInfoCommand`）：光标在第几节、一共几节。与 页码/行/列 那几格不同，
   * 这一对**不量几何**——节的边界是文件自己声明的（`w:pPr/w:sectPr` 的段落结束它那一节，`sectionEndsOf` 从模型
   * 数出来），块 `b` 属于第 `1 + #{e < b}` 节：带断点的段落自己属于**它结束的那一节**。节首一块的头是
   * `sectionCaretAnchor`（「第 N 节在哪里」问的是同一张表）。
   *
   * 光标不在纸面上时返回 `undefined`，与 `caretPage()` 同一条口径（由状态栏保留上一次读数）。
   */
  caretSection(): { section: number; sections: number } | undefined {
    if (!this.page) return undefined;
    const sel = document.getSelection();
    const node = sel?.focusNode ?? null;
    if (!node || !this.page.contains(node)) return undefined;
    const at = this.pathOf(node, sel!.focusOffset);
    if (!at) return undefined;
    const ends = sectionEndsOf(this.doc.blocks);
    const b = at.path[0]!;
    if (b >= this.doc.blocks.length) return undefined;
    return { section: 1 + ends.filter((e) => e < b).length, sections: ends.length + 1 };
  }

  /**
   * 状态栏 纵向位置 那一格（WPS 的 `KxWpsPagePositionInfoCommand`，kuip 452 行，detail「@Vertical Page
   * Position. Click to Open "Go To" dialog.」）：光标那行离**纸沿**多远（px，从纸的上沿往下量）。
   * 取点、判「光标在不在纸面上」与 `caretPage()` 同一条：这一格与「第 X 页」必须出自同一处光标。
   * 纸沿怎么反推、量到哪条线，口径都写在 `doc-layout.ts` 的 `pagePositionAt` 上。厘米与小数位是显示层
   * （状态栏）的事，这里报 px。
   */
  caretPosition(): number | undefined {
    if (!this.page) return undefined;
    const sel = document.getSelection();
    const node = sel?.focusNode ?? null;
    if (!node || !this.page.contains(node)) return undefined;
    const mid = caretMidOf(node, sel!.focusOffset);
    if (mid === undefined) return undefined;
    return pagePositionAt(this.page, this.doc, { mid, node, offset: sel!.focusOffset });
  }

  /**
   * 选区里那几段被切出来的文字，按 `counts()` **同一份口径**数（`countTexts` 只有一个实现，不会两处算法）。
   * 起止用的是页面上真选区映射进模型的那两条路径（`rangeEnds()`，与格式命令改选区时同一条寻址），所以「选中」
   * 在模型里说的位置与在页面上说的是同一处。跨段时两端各取自己那半段，中间那些块整段计入（表格连格子里的段落下钻）。
   * 折叠光标返回 `undefined`：那时没有「所选文字」这一栏，而不是报一排 0。
   */
  selectionCounts(): WordCountReadout | undefined {
    const ends = this.rangeEnds();
    if (!ends) return undefined;
    const slice = (path: number[], from: number, to: number): string[] => {
      const p = this.paragraphAt(path);
      if (!p) return [];
      return [p.inlines.map((i) => i.text ?? "").join("").slice(from, to)];
    };
    if (ends.start.path.join() === ends.end.path.join()) {
      return ends.start.offset === ends.end.offset ? undefined : countTexts(slice(ends.start.path, ends.start.offset, ends.end.offset));
    }
    // 中间那些块按顶层顺序走（与 `applyStyle` 改跨段选区那条同一口径）：起点与终点落在**同一张表**里时，
    // 表内那几段中间段落不在这一步里——两端各算自己那半段。这条限制是照 `applyStyle` 抄的，两处一致比一处更聪明要重要。
    const between = this.doc.blocks.slice(ends.start.path[0] + 1, ends.end.path[0]);
    return countTexts([...slice(ends.start.path, ends.start.offset, Infinity), ...paragraphTextsOf(between), ...slice(ends.end.path, 0, ends.end.offset)]);
  }

  /** Ribbon command: 「字数统计」. Reading a count changes nothing, so this takes no undo step and does not call `changed()`. */
  showWordCount(): void {
    openWordCount(this.counts(), this.selectionCounts(), this.layout());
  }

  /**
   * 「定位」（⌘G 与状态栏 页码／页数／节 那三格，WPS 的 `KxLegacyTriggerCommand id="GoTo"`，`doc-goto.ts` 头注有
   * 实据）：先量一次当前页（`caretPage()`，与状态栏同一份），再开对话框。纸没画上、或页高量不出来时**不开**——
   * 一个没有页数的定位对话框只能骗人。光标不在纸面上时种子取第 1 页／第 1 节：状态栏那种情形保留上一次读数，
   * 编辑器没有那一份存档，这里不假装有。`target` 是打开时预选的目标（页／节／行／书签／批注／表格／图形／标题），由点哪一格决定；
   * ⌘G 走默认的「页」。打开这一刻顺手把书签名单（`bookmarks()`）、批注条数（`doc.comments` 的条数，与
   * 纸上编号同一份数组）与总行数/光标行号（`caretLineNumber`，绝对行号）量一份交给对话框——都是这一刻的
   * 模型读数，对话框是模态的，期间模型不会变；批注号的种子从光标所在的那条读（`commentIdsAtCaret` 给的 id
   * 反查下标 + 1），落点不指着任何一条时归 1；行号的种子是光标所在那一行的绝对行号，光标不在纸面上时归 1；
   * 表格张数与表格号的种子（`tableCount()`／`tableNumberAtCaret()`）都是这一刻从模型数的，光标不在任何
   * 表格里时归 1；图形个数（`graphicCount()`）同理，图形没有「光标所在的那一枚」，种子永远归 1；
   * 标题个数与标题号的种子（`headingCount()`／`headingNumberAtCaret()`）数的是导航窗格那趟大纲走查，
   * 光标不在任何一枚标题里时归 1。
   */
  showGoTo(target: "page" | "section" | "line" | "bookmark" | "comment" | "table" | "graphic" | "heading" = "page"): void {
    const at = this.caretPage();
    const pages = at?.pages ?? this.layout()?.pages;
    if (pages === undefined) return;
    const cs = this.caretSection();
    const sections = cs?.sections ?? sectionEndsOf(this.doc.blocks).length + 1;
    const cl = this.caretLineNumber();
    const first = this.commentIdsAtCaret()[0];
    const at0 = first === undefined ? -1 : this.doc.comments.findIndex((c) => c.id === first);
    openGoTo(
      { page: at?.page ?? 1, pages, section: cs?.section ?? 1, sections, line: cl?.line ?? 1, lines: this.layout()?.lines ?? 0, bookmarks: this.bookmarks().map((b) => b.name), comments: this.doc.comments.length, comment: at0 >= 0 ? at0 + 1 : 1, table: this.tableNumberAtCaret() ?? 1, tables: this.tableCount(), graphic: 1, graphics: this.graphicCount(), heading: this.headingNumberAtCaret() ?? 1, headings: this.headingCount(), target },
      { page: (n) => this.goToPage(n), section: (n) => this.goToSection(n), line: (n) => this.goToLine(n), bookmark: (name) => this.goToBookmark(name), comment: (n) => this.goToComment(n), table: (n) => this.goToTable(n), graphic: (n) => this.goToGraphic(n), heading: (n) => this.goToHeading(n) },
    );
  }

  /**
   * 定位到第 `page` 页：`pageCaretAnchor` 直接给出这一页第一枚单位的文档位置，落光标（折叠 Range），
   * 再把那一块滚进视野中上段。不经过视口取点：`caretRangeFromPoint` 只认视口里的点，「先滚后量」要求
   * 滚动与量数同步，对不齐就会落到不知道哪里去；按「枚→块」的来路直接指出位置，与滚动、视口都无关。
   * 先落光标再滚动：选区记的是文档位置，不跟着视口走。落完 `focus({ preventScroll: true })`（与
   * `revealOutline` 同一条，别让聚焦又把刚滚好的视口挪走）。越界的页号在 `pageCaretAnchor` 那层就不给位置，
   * 这里连选区都不动。
   */
  goToPage(page: number): boolean {
    if (!this.page) return false;
    const anchor = pageCaretAnchor(this.page, this.doc, page);
    return anchor ? this.landAt(anchor) : false;
  }

  /**
   * 定位到第 `section` 节：落点是**节首一块的头**（`sectionCaretAnchor`，节的边界从模型数出来，见
   * `sectionEndsOf`）。落光标与滚动同 `goToPage`——两条命令共用一个落点动作 `landAt`，不许两处各写一份。
   */
  goToSection(section: number): boolean {
    if (!this.page) return false;
    const ends = sectionEndsOf(this.doc.blocks);
    const anchor = sectionCaretAnchor(this.page, ends, ends.length + 1, section);
    return anchor ? this.landAt(anchor) : false;
  }

  /**
   * 定位到第 `line` 行：行号是**绝对**的（从文首连数、换页不重数），与「行」那一格的种子同一条口径；
   * 落点落在这一行头一个字符上（`lineCaretAnchor`，行首按上沿二分量出来，与「列」同一趟取点）。
   * 落光标与滚动同 `goToPage`——同一枚落点动作 `landAt`。越界的行号在 `lineCaretAnchor` 那层就不给位置，
   * 返回 false（对话框的 1..N 来自同一刻量出来的总行数，这是第二道闸门）。WPS 真机「行」按绝对还是按页
   * 重数**未取证**。
   */
  goToLine(line: number): boolean {
    if (!this.page) return false;
    const anchor = lineCaretAnchor(this.page, this.doc, line);
    return anchor ? this.landAt(anchor) : false;
  }

  /**
   * 正文里的书签，按文档顺序：`w:bookmarkStart` 是随 run/段落携带的结构（`CarriedInline`），这里只收
   * `forEachInline` 那一趟能走到的地方——正文块与表内（含嵌套表）。**页眉页脚的书签不进列表**：那是另一
   * 个 story，Word 的「定位」列的是正文书签（真机**未取证**）。含 `_` 前缀的机器书签（`_Toc`／`_GoBack`）
   * 不过滤：WPS 的「隐藏书签」开关真机**未取证**，而 /tmp 那批 86 份 WPS 写的件里 86/86 都带这类名字，
   * 过滤掉等于把列表做成空的。同名取第一枚（Word 自己不许同名；这 86 份里 0 处重复，真机遇重名落哪一枚
   * **未取证**）。表格行与行之间的 `w:bookmarkStart`（`betweenRows`）不在这一趟里——纸上没有它的元素
   * （见 `goToBookmark` 同一条），语料那 86 份实测该位置只出现过 `w:bookmarkEnd`。
   */
  bookmarks(): { name: string; key: number }[] {
    const out: { name: string; key: number }[] = [];
    const seen = new Set<string>();
    forEachInline(this.doc.blocks, (inline) => {
      const carry = inline.carry;
      if (carry?.tag !== "bookmarkStart" || carry.name === undefined || seen.has(carry.name)) return;
      seen.add(carry.name);
      out.push({ name: carry.name, key: carry.key });
    });
    return out;
  }

  /**
   * 定位到书签 `name`：第一枚同名书签的渲染键找回它在纸上的 `run-carry` 元素（`data-carry` 是页上唯一
   * 的把手，见 `carries` 的注释），光标落在它前面那一位——书签零宽，落前落后是同一处。落点动作与
   * `goToPage`／`goToSection` 共用 `landAt`；滚动要的「块」取元素在纸上的直接祖先（书签在表格单元格里
   * 时，块就是那张表）。元素不在纸上——行间书签这类没有元素的位置，或段落已被重整而标记被吞——不落，
   * 返回 false；名字不在列表里同样 false（对话框的名单来自 `bookmarks()`，这是第二道闸门）。
   */
  goToBookmark(name: string): boolean {
    if (!this.page) return false;
    const hit = this.bookmarks().find((b) => b.name === name);
    if (!hit) return false;
    const span = this.page.querySelector<HTMLElement>(`.${RUN_CARRY_CLASS}[data-carry="${hit.key}"]`);
    const parent = span?.parentNode;
    if (!span || !parent) return false;
    let block: HTMLElement | null = span.parentElement;
    while (block && block.parentElement !== this.page) block = block.parentElement;
    if (!block) return false;
    return this.landAt({ node: parent, offset: [...parent.childNodes].indexOf(span), block });
  }

  /**
   * 定位到第 `n` 条批注：号码与纸上的批注编号**同源**——`render.ts` 那枚 `numbers` 映射就是 `this.doc.comments`
   * 的下标 + 1（回复也占一个号，虽然它自己没有标记、纸上不显号），这里数的是同一个数组，「打哪个号」与
   * 「纸上标着几号」不是两份算法。落点走审阅窗格那条「定位」`revealComment`：锚着文字的那条选中那段文字；
   * 回复没有标记，跳到父批注那一段；一个标记都找不回来的（锚在页眉带里、标记被重整吞掉）落不了点、返回
   * false。号码越界同样 false（对话框的 1..N 来自同一刻的 `doc.comments.length`，这是第二道闸门）。WPS
   * 真机上「批注」目标数的是不是同一份下标、回复算不算一枚**未取证**。
   */
  goToComment(n: number): boolean {
    if (!this.page) return false;
    if (!Number.isInteger(n) || n < 1 || n > this.doc.comments.length) return false;
    return this.revealComment(this.doc.comments[n - 1]!.id);
  }

  /**
   * 本文的表格张数：正文块与表内按**文档顺序**数，嵌套表也算一张——父表计入它自己的位置，其格里的表紧随
   * 其后（与「书签」名单的递归同一条：表内按行→格→段；口径见 `doc-goto.ts` 头注「表格」那一段）。WPS 真机
   * 数不数嵌套表**未取证**；这里钉的是自己那条回路：对话框那格的「本文共 N 张表格」与 `goToTable` 的第 n 张
   * 出自同一趟走查。
   */
  tableCount(): number {
    let n = 0;
    const walk = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind !== "table") continue;
        n += 1;
        for (const row of b.rows) for (const cell of row) walk(cell.paragraphs);
      }
    };
    walk(this.doc.blocks);
    return n;
  }

  /**
   * 本文的图形个数：数的是**行内图片**（`Inline.image`）——正文块与表内按文档顺序数，嵌套表也算（与 `tableCount`
   * 同一趟走查的口径；`imagePaths` 给纸上每枚 `<img>` 盖 `data-path` 用的正是这个序）。浮动形状不进这一数：模型里
   * 还没有它们的位置，WPS 真机把浮动形状算不算一枚图形**未取证**；这里钉的是自己那条回路：对话框那格的「本文共
   * N 个图形」与 `goToGraphic` 的第 n 个出自同一趟走查。
   */
  graphicCount(): number {
    let n = 0;
    const walk = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind === "table") b.rows.forEach((row) => row.forEach((c) => walk(c.paragraphs)));
        else if (b.kind === "paragraph") for (const inline of b.inlines) if (inline.image) n += 1;
      }
    };
    walk(this.doc.blocks);
    return n;
  }

  /**
   * 本文的标题个数：导航窗格那份名单的长度（`outlineEntries()`——模型里带大纲级别的段落，表内也算）。
   * 这里不多走一趟：窗格列的就是纸上的标记，数与纸同一份读数的来源（`outlineMarks()` 与名单一一对号，
   * 见 `revealOutline` 那段注）。WPS 真机数不数表内标题**未取证**；这里钉的是自己那条回路：对话框那格的
   * 「本文共 N 个标题」与 `goToHeading` 的第 n 枚出自同一份大纲走查。
   */
  headingCount(): number {
    return this.outlineEntries().length;
  }

  /** 第 `n` 张表的引用与模型坐标前缀（`tableCount` 数的那一趟里的第 n 张）；数不到返回 undefined。 */
  private nthTable(n: number): { table: Table; prefix: number[] } | undefined {
    if (!Number.isInteger(n) || n < 1) return undefined;
    let left = n;
    let found: { table: Table; prefix: number[] } | undefined;
    const walk = (blocks: Block[], prefix: number[]): void => {
      for (let i = 0; i < blocks.length && !found; i += 1) {
        const b = blocks[i]!;
        if (b.kind !== "table") continue;
        left -= 1;
        if (left === 0) {
          found = { table: b, prefix: [...prefix, i] };
          return;
        }
        for (let r = 0; r < b.rows.length && !found; r += 1) {
          const row = b.rows[r]!;
          for (let c = 0; c < row.length && !found; c += 1) walk(row[c]!.paragraphs, [...prefix, i, r, c]);
        }
      }
    };
    walk(this.doc.blocks, []);
    return found;
  }

  /**
   * 一张表里第一个落得下的位置：按行→格→块找头一段落，落点是它的头一个字符前（`offset: 0`）。格子头一层的块里
   * 又有表的，就再往里一层——与「表格」目标的落点约定同一条（`doc-goto.ts` 头注）。整张表一个段落都没有才数
   * 不出来（模型上不该有，写出侧也会补 `w:p`）。
   */
  private firstSpotIn(table: Table, prefix: number[]): DocPath | undefined {
    for (let r = 0; r < table.rows.length; r += 1) {
      const row = table.rows[r]!;
      for (let c = 0; c < row.length; c += 1) {
        const blocks = row[c]!.paragraphs;
        for (let i = 0; i < blocks.length; i += 1) {
          const b = blocks[i]!;
          if (b.kind === "paragraph") return { path: [...prefix, r, c, i], offset: 0 };
          if (b.kind !== "table") continue;
          const deeper = this.firstSpotIn(b, [...prefix, r, c, i]);
          if (deeper) return deeper;
        }
      }
    }
    return undefined;
  }

  /**
   * 定位到第 `n` 张表：光标落在 `firstSpotIn` 指出那个位置的头一个模型字符上——先按模型路径拿到纸上的元素，
   * 再取它上面第 0 个落点（空段落一个字符都没有时落在元素自身开头），滚动的那一块取它在纸上的顶层祖先
   * （表就是那张顶层的表），与 `goToBookmark`／`goToPage` 共用 `landAt`。号码越界、落点找不回纸上（表格刚被
   * 重整之类）都返回 false——对话框的 1..N 来自同一刻的 `tableCount()`，这是第二道闸门。WPS 真机数不数
   * 嵌套表、落点在格首还是别处**未取证**。
   */
  goToTable(n: number): boolean {
    if (!this.page) return false;
    const hit = this.nthTable(n);
    if (!hit) return false;
    const spot = this.firstSpotIn(hit.table, hit.prefix);
    if (!spot) return false;
    const el = this.elementAt(spot.path);
    if (!el) return false;
    let block: HTMLElement | null = el instanceof HTMLElement ? el : el.parentElement;
    while (block && block.parentElement !== this.page) block = block.parentElement;
    if (!block) return false;
    const at = positions(el)[0] ?? { node: el, offset: 0 };
    return this.landAt({ node: at.node, offset: at.offset, block });
  }

  /**
   * 定位到第 `n` 个图形：按 `graphicCount` 的同一趟走查找到第 n 枚行内图片——它的段落路径与它在那一段落里的
   * 次序 `k`。先按模型路径拿到纸上的段落元素，再取这一段里第 `k` 枚 `<img>`（段落元素里的 `<img>` 只有这一段
   * 自己的图片，格子里的嵌套表是兄弟元素、不在里面）。落点是那枚图片**前面那一处**：折叠 Range 落到 `img` 的
   * 父节点（那个段落元素）里 `img` 之前的下标上——图片是行内的一员，光标进不了它里面，只能落在它旁边。号码
   * 越界、落点找不回纸上都返回 false——对话框的 1..N 来自同一刻的 `graphicCount()`，这是第二道闸门。WPS 真机
   * 落点在图形前还是别处**未取证**。
   */
  goToGraphic(n: number): boolean {
    if (!this.page) return false;
    if (!Number.isInteger(n) || n < 1) return false;
    let left = n;
    let found: { path: number[]; k: number } | undefined;
    const walk = (blocks: Block[], prefix: number[]): void => {
      for (let i = 0; i < blocks.length && !found; i += 1) {
        const b = blocks[i]!;
        if (b.kind === "table") {
          for (let r = 0; r < b.rows.length && !found; r += 1) {
            const row = b.rows[r]!;
            for (let c = 0; c < row.length && !found; c += 1) walk(row[c]!.paragraphs, [...prefix, i, r, c]);
          }
        } else if (b.kind === "paragraph") {
          let k = 0;
          for (const inline of b.inlines) {
            if (!inline.image) continue;
            left -= 1;
            if (left === 0) {
              found = { path: [...prefix, i], k };
              break;
            }
            k += 1;
          }
        }
      }
    };
    walk(this.doc.blocks, []);
    if (!found) return false;
    const el = this.elementAt(found.path);
    if (!el) return false;
    const img = el.querySelectorAll("img")[found.k];
    if (!img || !img.parentNode) return false;
    let block: HTMLElement | null = el instanceof HTMLElement ? el : el.parentElement;
    while (block && block.parentElement !== this.page) block = block.parentElement;
    if (!block) return false;
    const offset = [...img.parentNode.childNodes].indexOf(img);
    return this.landAt({ node: img.parentNode, offset, block });
  }

  /**
   * 光标所在标题的号（按 `outlineEntries` 数出来的那个号）：从 `caretPath` 的模型路径取光标所在的那一段落
   * 元素，看它是不是纸上的大纲标记（`data-outline`——`render.ts` 给解析得出级别的段落盖上的一枚），是就回到
   * 它在 `outlineMarks()` 里的下标 + 1。光标不在任何一枚标题里返回 undefined（对话框的种子归 1）。号与落点
   * 出自同一份大纲读数——`outlineMarks()` 与 `outlineEntries()` 一一对号，这里是第三处读者，没有第二份算法。
   */
  private headingNumberAtCaret(): number | undefined {
    const at = this.caretPath();
    if (!at) return undefined;
    const el = this.elementAt(at.path);
    if (!el || !el.hasAttribute("data-outline")) return undefined;
    const index = this.outlineMarks().indexOf(el as HTMLElement);
    return index >= 0 ? index + 1 : undefined;
  }

  /**
   * 定位到第 `n` 枚标题：落点走导航窗格那条「定位」——`revealOutline` 按下标点，选中标题文字、光标停在
   * **标题末尾**、纸上那枚标记打上高亮并滚进视野。号码越界、纸上那枚标记找不回来都返回 false——对话框的
   * 1..N 来自同一刻的 `headingCount()`，这是第二道闸门。Word 真机 `GoTo` 的落点在标题**开头**、数不数表内
   * 标题**未取证**；这里钉的是自己那条回路：落点与导航窗格点一行是同一条路，末尾也是窗格那条路的末尾。
   */
  goToHeading(n: number): boolean {
    if (!Number.isInteger(n) || n < 1) return false;
    return this.revealOutline(n - 1);
  }

  /**
   * 光标所在表格的号（按 `tableCount` 数出来的那个号）：从 `caretPath` 的模型路径取**最内层**包着光标的
   * 那张表的前缀——光标在格子的段落里时是它直接所在的那张表，路径再深一层（格子里的表）时是那张嵌套表。
   * 光标不在任何表格里返回 undefined（对话框的种子归 1）。号与落点出自同一趟模型走查，不是第二份算法。
   */
  private tableNumberAtCaret(): number | undefined {
    const at = this.caretPath();
    if (!at || at.path.length < 4) return undefined;
    const target = at.path.slice(0, at.path.length - 3).join();
    let n = 0;
    let hit: number | undefined;
    const walk = (blocks: Block[], prefix: number[]): boolean => {
      for (let i = 0; i < blocks.length; i += 1) {
        const b = blocks[i]!;
        if (b.kind !== "table") continue;
        const here = [...prefix, i];
        n += 1;
        if (here.join() === target) {
          hit = n;
          return true;
        }
        for (let r = 0; r < b.rows.length; r += 1) {
          const row = b.rows[r]!;
          for (let c = 0; c < row.length; c += 1) {
            if (walk(row[c]!.paragraphs, [...here, r, c])) return true;
          }
        }
      }
      return false;
    };
    walk(this.doc.blocks, []);
    return hit;
  }

  /** 落点动作：折叠 Range 落到 `anchor` 指出的文档位置，块滚进视野中上段，焦点收进纸张。 */
  private landAt(anchor: PageAnchor): boolean {
    const sel = document.getSelection();
    if (!sel) return false;
    const range = document.createRange();
    range.setStart(anchor.node, anchor.offset);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    // 滚动走最近的滚动祖先（外壳里是 `#view`），把这一块的上沿带到视口上三分之一处。
    for (let p = anchor.block.parentElement; p; p = p.parentElement) {
      const cs = getComputedStyle(p);
      if (cs.overflowY === "auto" || cs.overflowY === "scroll") {
        const box = p.getBoundingClientRect();
        p.scrollBy(0, anchor.block.getBoundingClientRect().top - (box.top + p.clientHeight / 3));
        break;
      }
    }
    this.page!.focus({ preventScroll: true });
    return true;
  }

  /**
   * 「页面设置」: the page's margins, band distances, binding width and direction. The section the numbers land in
   * is the one the page is drawn from, and the step back restores it — the body's own `w:sectPr` bytes included,
   * which a text-only undo step would leave where they were.
   *
   * 光标在开窗前的位置上不动：对话框关窗时已经把押住的选区放回来了（`doc-page-setup.ts` 的 `restore`），
   * 所以这里读到的就是开窗前那一个。`render()` 整张纸重建、旧节点全断，落点按**模型路径**找回来
   * （与 `undo`/`redo` 同一条），不拿 DOM 节点。WPS 真机按「确定」之后光标停在哪 未取证；本版钉的是
   * 「不动」——同一个字符位置。
   */
  applyPageSetup(set: PageSetup): void {
    const caret = this.caretPath();
    this.snapshot(caret);
    this.doc.setPageSetup(set);
    this.render();
    if (caret) this.setCaret(this.clamp(caret));
    this.hooks.changed();
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
    const section = this.doc.sections[0] ?? DEFAULT_SECTION;
    const usable = twipsToPx(section.pageWidthTwips - section.margins.left - (section.margins.gutter ?? 0) - section.margins.right);
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
   * 三条断命令共用的一条腿：在**光标处**放一枚 `w:br`。段落还是那一段，不因为要断就多出一个空段落（从前这条
   * `splice` 的是「下一段带 `pageBreakBefore`」，那是另一条命令的形状）。`mutated()` 那一步会把 `data-brk` 里的
   * 原串读回 `brXml`，之后这段里再怎么打字都还按原来那一型写回文件。
   */
  private insertBreakAtCaret(brXml: string): void {
    const page = this.page;
    const at = this.caretPath();
    if (!page || !at) return;
    page.focus();
    this.setCaret(at);
    const sel = document.getSelection();
    const range = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
    if (!range || !page.contains(range.startContainer)) return;
    this.snapshot(at);
    // `execCommand("insertHTML")` 放不下这枚标记：光标在段落末尾时 Blink 会把那枚 `<br>` 当成换行处理，
    // 段落被劈开而标记自己消失（实测段落 4 → 5、`w:br` 0 枚）。所以直接按 DOM 插，再让 `mutated()` 把它读回模型。
    const br = document.createElement("br");
    br.setAttribute("data-brk", brXml);
    range.insertNode(br);
    range.setStartAfter(br);
    range.collapse(true);
    sel?.removeAllRanges();
    sel?.addRange(range);
    this.mutated();
    this.hooks.changed();
  }

  /**
   * WPS 的 `InsertPageBreak`（`office6/res/wpscommon.kuip:219`，`shortcutHint="Ctrl+Enter"`，快捷键表 1115/1116
   * 与 1304/1305 给了 `Ctrl+Return`／`Meta+Return`）：`w:br w:type="page"`。
   */
  insertPageBreak(): void {
    this.insertBreakAtCaret('<w:br w:type="page"/>');
  }

  /**
   * WPS 的 `InsertTextWrappingBreak`（`wpscommon.kuip` aid=80156，`shortcutHint="Shift+Enter"`，快捷键表另有
   * `Shift+Return`；ribbon 上跟分页符、分栏符挂在同一个「断行」下拉里）：`w:br w:type="textWrapping"`。
   * 平时按 Shift+Enter 走的是浏览器自己的软换行，这一条是给菜单/工具栏那颗用的。
   */
  insertLineBreak(): void {
    this.insertBreakAtCaret(TEXTWRAPPING_BR);
  }

  /**
   * WPS 的 `InsertColumnBreak`（快捷键表 `key="Ctrl+Shift+Return"`，ribbon id 在插入页与页面布局页各一枚）：
   * `w:br w:type="column"`。命令声明行（aid、按钮文案）没在 `wpscommon.kuip` 里找到，所以这一条按**快捷键表 +
   * ribbon 挂载**接，文案沿用 WPS 家族里已经取证的「断行」说法；语料 86 份里 `column` 0 枚，形状是规范给的。
   */
  insertColumnBreak(): void {
    this.insertBreakAtCaret('<w:br w:type="column"/>');
  }

  // ------------------------------------------------------------- toolbar

  private paragraphAt(path: number[]): Paragraph | undefined {
    const block = this.blockAt(path);
    return block?.kind === "paragraph" ? block : undefined;
  }

  /**
   * What the caret's run and paragraph currently look like, so the toolbar can show its state.
   *
   * `highlight` keeps the three states #88 separated, because they are three different sentences the file
   * could have written: `undefined` is 「没提过」, `"none"` is a run that *says* it has no background, and a
   * colour name is a background. A control that shows 「没提过」 the same way as 「说了没有」 answers the reader
   * with a silence where the document spoke.
   */
  markState(): { bold: boolean; italic: boolean; underline: boolean; strike: boolean; sizePt?: number; highlight?: string; align?: string; list?: "bullet" | "number" } {
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
    // A paragraph whose numbering the file cancelled with `w:numId w:val="0"` is not in a list, and one that
    // only borrows the list from its style is — the button follows what the page draws.
    const list = paragraph?.numPr && paragraph.numPr.numId !== 0 ? listKind(this.doc, paragraph.numPr.numId) : undefined;
    return {
      bold: carried.bold ?? false,
      italic: carried.italic ?? false,
      underline: carried.underline ?? false,
      strike: carried.strike ?? false,
      sizePt: carried.sizePt,
      highlight: carried.highlight,
      align: paragraph?.align,
      list,
    };
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
      };
      const key = Object.keys(patch)[0];
      const value = patch[key as keyof StylePatch];
      // 高亮与字号这两族的待用格式由编辑器自己举（`armPending`）：Blink 交不出「无」，也没有写得出磅数的命令。
      // 余下的（粗体/斜体/下划线/删除线/墨色）实测待用有效，仍交给它。
      if (key === "highlight" || key === "sizePt") this.armPending(patch);
      else {
        document.execCommand("styleWithCSS", false, "true");
        document.execCommand(command[key]!, false, typeof value === "string" ? value : undefined);
      }
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
    const targets = this.selectionParagraphs();
    if (!targets.length) return;
    this.snapshot(this.caretPath());
    // 一次点击只讨一个 numId：选中五段按「项目符号」是**一个**列表接着排，不是五段各自起一个新的列表。
    const numId = patch.list && patch.list !== "none" ? this.doc.ensureListNumbering(patch.list) : undefined;
    for (const paragraph of targets) {
      const was = paraStyleOf(paragraph);
      // `syncParaSlots` replaces the array, so this reference keeps the set as the paragraph came in with it.
      const wasSlots = paragraph.pPrSlots;
      if (patch.align !== undefined) paragraph.align = patch.align || undefined;
      if (patch.list !== undefined) {
        if (patch.list === "none") {
          // 取消编号是文件自己写的形状：语料里 517 段带着 `w:numPr` 而 `w:numId w:val="0"`。清空字段在这里不够，
          // 因为编号现在会从样式链上借来（48 段就是这么进列表的）——只有写下那个 0 才盖得住样式的那一条。
          // 本来就不在列表里的段落什么都不写。
          if (paragraph.numPr && paragraph.numPr.numId !== 0) paragraph.numPr = { numId: 0, ilvl: paragraph.numPr.ilvl };
        } else if (numId !== undefined) {
          // WPS lets this click through on a page that has no numbering part at all, so the definition is the
          // document's own when it names the shape and the model's when it does not (`ensureListNumbering`).
          paragraph.numPr = { numId, ilvl: paragraph.numPr?.ilvl ?? 0 };
        }
      }
      // The paragraph's `w:pPr` travels as text, so the child these two buttons own moves with them.
      syncParaSlots(paragraph);
      // An alignment click on a paragraph that already sits there changes nothing, and a change that is
      // nothing should not be put in front of a reviewer.
      if (this.trackRevisions && JSON.stringify(was) !== JSON.stringify(paraStyleOf(paragraph))) {
        paragraph.revFormat = this.signParagraphRevision(paragraph, was, wasSlots);
      }
    }
    const caret = this.caretPath();
    this.render();
    if (caret) this.setCaret(this.clamp(caret));
    this.hooks.changed();
  }

  /**
   * WPS 列表项里 Tab／Shift+Tab 改的那个数：`w:numPr/w:ilvl`。只在**这一级由定义自己说过**时动
   * （`w:lvl` 是文件的形状，页面不能凭空造一条），也不动 `w:numId w:val="0"` 那种取消——那一段根本不在列表里。
   * 段落从样式链借来的那一条也算它的（读取时已经落在 `paragraph.numPr` 上），改完由段落自己写下这一笔。
   */
  shiftListLevel(dir: 1 | -1): boolean {
    const path = this.caretPath()?.path;
    const paragraph = path ? this.paragraphAt(path) : undefined;
    const numPr = paragraph?.numPr;
    if (!paragraph || !numPr || numPr.numId === 0) return false;
    const to = numPr.ilvl + dir;
    if (to < 0 || this.doc.levelOf(numPr.numId, to) === undefined) return false;
    this.snapshot(this.caretPath());
    const was = paraStyleOf(paragraph);
    const wasSlots = paragraph.pPrSlots;
    paragraph.numPr = { numId: numPr.numId, ilvl: to };
    syncParaSlots(paragraph);
    if (this.trackRevisions && JSON.stringify(was) !== JSON.stringify(paraStyleOf(paragraph))) {
      paragraph.revFormat = this.signParagraphRevision(paragraph, was, wasSlots);
    }
    const caret = this.caretPath();
    this.render();
    if (caret) this.setCaret(this.clamp(caret));
    this.hooks.changed();
    return true;
  }

  /**
   * WPS 主页 ribbon 段落组那两条 增加缩进量／减少缩进量（`IncreaseIndent` aid=3472 `idMso="IndentIncreaseWord"`、
   * `DecreaseIndent` aid=3473 `idMso="IndentDecreaseWord"`，右键菜单里同名的两条是 80089/80088，Touch Bar 图标的
   * 中文标题写着「增加缩进量」「减少缩进量」）改的那个数。一档是多少、按字符还是按缇走，都由 `WordDocument.nextIndent`
   * 按件自己的写法决定（取证见那里）；已经在边上时它返回 `undefined`，这一击就不进撤销栈。
   */
  shiftParagraphIndent(dir: 1 | -1): boolean {
    const targets = this.selectionParagraphs();
    // 每一段按**它自己**的写法走一档（`nextIndent` 认件里是 `w:left` 还是 `w:leftChars`）；已经在边上的那一段
    // 整段不动，全都在边上时这一击不进撤销栈——与只有一段时同一个判定。
    if (!targets.some((p) => this.doc.nextIndent(p, dir))) return false;
    this.snapshot(this.caretPath());
    for (const paragraph of targets) {
      const next = this.doc.nextIndent(paragraph, dir);
      if (!next) continue;
      const was = paraStyleOf(paragraph);
      const wasSlots = paragraph.pPrSlots;
      paragraph.indentLeft = next.indentLeft;
      if (next.indentLeftChars !== undefined) paragraph.indentLeftChars = next.indentLeftChars;
      syncParaSlots(paragraph);
      if (this.trackRevisions && JSON.stringify(was) !== JSON.stringify(paraStyleOf(paragraph))) {
        paragraph.revFormat = this.signParagraphRevision(paragraph, was, wasSlots);
      }
    }
    const caret = this.caretPath();
    this.render();
    if (caret) this.setCaret(this.clamp(caret));
    this.hooks.changed();
    return true;
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
    // One change at a time is lit: without this, stepping 下一条 accumulates a flash on every row visited
    // since the last render, so the page shows the whole walk rather than where the reader now is.
    for (const stale of this.page.querySelectorAll<HTMLElement>(".solo-rev-hit")) stale.classList.remove("solo-rev-hit");
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
    this.review?.revealFromPage(entry.id);
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

  /** Ribbon toggle for 审阅窗格; with no argument it flips whatever the reader is looking at. */
  showReviewPane(show = !(this.review?.isOpen() ?? false)): void {
    this.review?.setOpen(show);
  }

  get reviewPaneOpen(): boolean {
    return this.review?.isOpen() ?? false;
  }

  /** What the 导航窗格 lists: the headings the document's outline levels name, in reading order. */
  outlineEntries(): { level: number; text: string }[] {
    return this.doc.outline();
  }

  /**
   * The page's heading marks in document order, index for index with `outlineEntries()`: `renderBlocks` writes
   * one `[data-outline]` per paragraph that resolves to a level and walks the model the same way `outline()`
   * does, table cells included. So a row number is an address in the page, and the pane needs no scheme of its
   * own for finding the paragraph it names.
   */
  outlineMarks(): HTMLElement[] {
    return [...(this.page?.querySelectorAll<HTMLElement>("[data-outline]") ?? [])];
  }

  /**
   * 定位到第 N 个标题: lights the heading and brings it into view, the way clicking a row in WPS' 导航窗格 jumps
   * to the section it names. The caret goes to the end of the heading, which is where the reader means to
   * continue from.
   */
  revealOutline(index: number): boolean {
    if (!this.page) return false;
    const mark = this.outlineMarks()[index];
    if (!mark) return false;
    // One place at a time is lit: the page already uses this class for the change the review pane last named.
    for (const stale of this.page.querySelectorAll<HTMLElement>(".solo-rev-hit")) stale.classList.remove("solo-rev-hit");
    mark.classList.add("solo-rev-hit");
    const sel = document.getSelection();
    if (sel) {
      const range = document.createRange();
      range.selectNodeContents(mark);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    }
    mark.scrollIntoView({ block: "center" });
    this.page.focus({ preventScroll: true });
    this.outline?.revealFromPage(index);
    return true;
  }

  /** Ribbon toggle for 导航窗格; with no argument it flips whatever the reader is looking at. */
  showOutlinePane(show = !(this.outline?.isOpen() ?? false)): void {
    this.outline?.setOpen(show);
  }

  get outlinePaneOpen(): boolean {
    return this.outline?.isOpen() ?? false;
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
    // The record says what the run's whole property set was, not only the fields the model can name, so a
    // later refusal puts back the `w:kern` and the `w:lang` the toolbar never touched. Where the run carried no
    // `w:rPr` of its own, what it showed came down the style ladder, and §17.13.5.19 records none of that: the
    // record says nothing, and a refusal puts the run back to borrowing.
    const stated = runStyleOf(before);
    const prev = before.rPrSlots?.length ? stated : ownRunStyle(before.asRead, stated);
    return {
      kind: "format",
      id: this.revSeq,
      author: this.authorName,
      date: stampNow(),
      prev,
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
    // The reviewer gets the set the paragraph wrote. When it carried a `w:pPr` the record carries that text; when
    // it carried none, the properties on it came down the style chain and §17.13.5.19 records none of them, so
    // the record says nothing and a refusal puts the paragraph back to borrowing.
    const recorded = prevSlots?.length ? prev : ownParaStyle(p.asRead, prev);
    return { kind: "format", id: this.revSeq, author: this.authorName, date: stampNow(), prev: recorded, ...(prevSlots?.length ? { prevSlots } : {}) };
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

/**
 * Paragraphs inherit the shape of the one before them — Word's rule for Enter mid-document.
 *
 * Passing `doc` is what makes it an Enter: only a paragraph the page just split asks the previous paragraph's
 * style for its §17.7.2.24 next style, and that switch is the model's own rule (`WordDocument.nextSplit`). A
 * paragraph that comes back because a break's deletion is still waiting to be decided takes the shape it had, not
 * a new style, so that caller leaves `doc` out.
 */
function fresh(previous: Paragraph | undefined, doc?: WordDocument): Paragraph {
  if (!previous) return { kind: "paragraph", inlines: [] };
  const split = doc?.nextSplit(previous);
  if (split) return split;
  const empty = previous.inlines.every((i) => i.text.trim() === "" && !isMark(i));
  return {
    kind: "paragraph",
    styleId: previous.styleId,
    styleName: previous.styleName,
    align: previous.align,
    numPr: empty ? undefined : previous.numPr,
    indentLeft: previous.indentLeft,
    indentRight: previous.indentRight,
    indentFirstLine: previous.indentFirstLine,
    indentFirstLineChars: previous.indentFirstLineChars,
    indentHanging: previous.indentHanging,
    spacingBefore: previous.spacingBefore,
    spacingAfter: previous.spacingAfter,
    lineSpacing: previous.lineSpacing,
    lineSpacingRule: previous.lineSpacingRule,
    // The level travels with the style for the same reason the style does, and it has to: a paragraph this
    // editor makes has no `w:pPr` slots to fall back on, so a heading split in two whose new half states no
    // level would save as body text.
    outlineLevel: previous.outlineLevel,
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
  if (inline.rPrSlots) style.rPrSlots = inline.rPrSlots;
  // The read snapshot travels with the set it was taken against. A run whose bold came from the paragraph's
  // style shows bold here, and the new run has to know it did not state it — otherwise the first save after
  // typing writes a `w:b` into a run whose file never had one, and the text stops following its style.
  if (inline.asRead) style.asRead = inline.asRead;
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
    else if (b.kind === "paragraph" && b.numPr && b.numPr.numId !== 0) n += 1;
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
