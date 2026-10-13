import type { Block, Carried, CarriedInline, Inline, PageBorder, Paragraph, RowInfo, Section, Table, TableBorder, WordDocument } from "./docx.js";
import { DEFAULT_SECTION, PAGE_BORDER_SIDES, twipsToPx } from "./docx.js";
import type { CommentAnchor } from "./comments.js";
import type { Presentation, Shape, SlideParagraph } from "./pptx.js";

const esc = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

export function renderInline(inline: Inline): string {
  if (inline.image) {
    return `<img src="${inline.image.src}" width="${Math.round(inline.image.width)}" height="${Math.round(inline.image.height)}" alt="${esc(inline.image.name)}">`;
  }
  if (inline.break) return "<br>";
  if (inline.tab) return '<span class="tab"></span>';
  let css = "";
  if (inline.bold) css += "font-weight:600;";
  if (inline.italic) css += "font-style:italic;";
  if (inline.strike) css += "text-decoration:line-through;";
  if (inline.underline) css += "text-decoration:underline;";
  if (inline.sizePt) css += `font-size:${inline.sizePt}pt;`;
  if (inline.color) css += `color:${inline.color};`;
  if (inline.font) css += `font-family:${esc(inline.font.replace(/[;\\]/g, ""))};`;
  if (inline.highlight) css += `background:${highlightColor(inline.highlight)};`;
  if (inline.vertAlign) css += `vertical-align:${inline.vertAlign === "superscript" ? "super" : "sub"};font-size:.75em;`;
  let html = inline.text === "" ? "" : `<span${css ? ` style="${css}"` : ""}>${esc(inline.text)}</span>`;
  if (html === "" && !css) html = esc(inline.text);
  // A link to a place in this document keeps its window to itself: `target="_blank"` on a `#name` href hands
  // the app's own page address to the system browser, where nobody has the bookmark open. The attributes the
  // file put on the link element have no meaning on the page, so they ride along as data for the editor to
  // hand back on the next keystroke — dropped here, they are gone from the saved file.
  if (inline.hyperlink)
    html =
      `<a href="${esc(inline.hyperlink)}"${inline.hyperlink.startsWith("#") ? "" : ' target="_blank" rel="noreferrer"'}` +
      `${inline.linkAttrs ? ` data-link-attrs="${esc(inline.linkAttrs)}"` : ""}>${html || esc(inline.text)}</a>`;
  // A tracked change is marked by class, never by inline style: the editor reads inline styles
  // back into the run, and "underlined because inserted" must not become "underlined because bold".
  const rev = inline.rev;
  if (rev) {
    const what = rev.move === "from" ? "移出" : rev.move === "to" ? "移入" : rev.kind === "del" ? "删除" : "插入";
    const who = `${rev.author}${rev.date ? ` · ${rev.date.slice(0, 10)}` : ""} ${what}`;
    html =
      `<span class="rev rev-${rev.kind}" data-rev="${rev.kind}" data-rev-id="${rev.id}"` +
      ` data-rev-author="${esc(rev.author)}"${rev.date ? ` data-rev-date="${esc(rev.date)}"` : ""}` +
      `${rev.move ? ` data-rev-move="${rev.move}"` : ""}` +
      ` title="${esc(who)}">${html || esc(inline.text)}</span>`;
  }
  // A formatting revision changes how a run looks, not what it says: the run keeps its text and shows
  // its new properties, so the record rides along as data on the run's own span for the editor to read
  // back. It is not a `w:ins`/`w:del` mark and must not be given one class or the edit will read it as
  // an insertion.
  const fmt = inline.revFormat;
  if (fmt) {
    const who = `${fmt.author}${fmt.date ? ` · ${fmt.date.slice(0, 10)}` : ""} 修改格式`;
    html =
      `<span class="rev-fmt" data-rev-format="${fmt.id}" data-rev-format-author="${esc(fmt.author)}"` +
      `${fmt.date ? ` data-rev-format-date="${esc(fmt.date)}"` : ""}` +
      ` data-rev-prev="${encodeURIComponent(JSON.stringify(fmt.prev))}"` +
      ` title="${esc(who)}">${html || esc(inline.text)}</span>`;
  }
  return html;
}

export function highlightColor(name: string): string {
  const map: Record<string, string> = {
    yellow: "#ffff00",
    green: "#00ff00",
    cyan: "#00ffff",
    magenta: "#ff00ff",
    red: "#ff0000",
    darkYellow: "#808000",
    lightGray: "#d4d4d4",
  };
  return map[name] ?? name;
}

const ALIGN: Record<string, string> = { center: "center", right: "right", both: "justify", distribute: "justify" };

/**
 * Comment rendering state for one pass over the body: which anchored ranges are currently
 * open, and the number each comment shows as its reference mark. Anchors are in-stream, so
 * the set is threaded across paragraphs and table cells exactly as the ranges are.
 */
interface CommentState {
  open: Set<number>;
  numbers: Map<number, number>;
  /** Ids whose thread is resolved, so the mark renders greyed rather than hidden. */
  done: Set<number>;
}

function anchorTag(anchor: CommentAnchor, numbers: Map<number, number>, done: Set<number>): string {
  const label = numbers.get(anchor.id);
  const text = anchor.kind === "ref" && label !== undefined ? String(label) : "";
  // WPS and Word keep a resolved thread in place and grey it out; hiding it would lose the
  // reader's context for the surrounding discussion.
  const resolved = done.has(anchor.id) ? " cmt-done" : "";
  return `<span class="cmt-anchor${anchor.kind === "ref" ? " cmt-ref" : ""}${resolved}" data-cid="${anchor.id}" data-kind="${anchor.kind}" contenteditable="false">${text}</span>`;
}

/** The class an inline carry is drawn with, and the only way the page hands the same one back. */
export const RUN_CARRY_CLASS = "run-carry";

/**
 * A run-level structure the model carries rather than understands. `data-carry` is its identity across a
 * re-render, the same job `data-bid` does for blocks. What shows is what a reader would see inside it — the
 * result of a page field, the date in an inline content control — and nothing at all for a bookmark, which
 * has no visible content to give.
 */
function carryTag(carry: CarriedInline): string {
  return `<span class="${RUN_CARRY_CLASS}" data-carry="${carry.key}" data-tag="${esc(carry.tag)}" contenteditable="false">${esc(carry.show)}</span>`;
}

/**
 * A page field inside a header/footer band, drawn as a labelled box. A page number that reads like the text
 * beside it would be typed over and deleted by accident; `页码` says it is one thing, and the caret cannot
 * enter it. Inline style, because the page's stylesheet is not this module's to change — and the label, not
 * the number the field cached, because this product has no pagination to say which page it would be.
 */
function hfFieldTag(carry: CarriedInline): string {
  return `<span class="${RUN_CARRY_CLASS}" data-carry="${carry.key}" data-tag="${esc(carry.tag)}" contenteditable="false" style="border:1px dashed #9aa3b0;border-radius:3px;padding:0 5px;margin:0 2px;background:#f2f5fa;color:#3b4453">${esc(carry.show)}</span>`;
}

/** The band's own row of commands, and the button that matches the page's other small buttons. */
const HF_BAR_STYLE = "display:flex;gap:6px;margin-top:4px";
const HF_BTN_STYLE = "font:inherit;font-size:11px;padding:2px 8px;border:1px solid #c3c9d2;border-radius:5px;background:#f2f4f7;color:#26292e;cursor:pointer";

/** Runs between an open range's markers are wrapped, so the anchored text reads as one block. */
function wordRuns(inlines: Inline[], state: CommentState): string {
  const out: string[] = [];
  let i = 0;
  while (i < inlines.length) {
    const inline = inlines[i];
    if (inline.carry) {
      out.push(carryTag(inline.carry));
      i += 1;
      continue;
    }
    if (inline.anchor) {
      if (inline.anchor.kind === "start") state.open.add(inline.anchor.id);
      else if (inline.anchor.kind === "end") state.open.delete(inline.anchor.id);
      out.push(anchorTag(inline.anchor, state.numbers, state.done));
      i += 1;
      continue;
    }
    const from = i;
    while (i < inlines.length && !inlines[i].anchor && !inlines[i].carry) i += 1;
    const stretch = inlines.slice(from, i).map(renderInline).join("");
    // `data-cids` says which comments hold this stretch anchored: the UI maps a caret in the
    // text back to a comment with it, and one range can sit inside several anchored ranges.
    const ids = [...state.open].join(",");
    out.push(ids === "" ? stretch : `<span class="cmt-in" data-cids="${ids}">${stretch}</span>`);
  }
  return out.join("");
}

/**
 * The paragraph's first-line indent as one CSS value, from whichever of the two spellings the file put its
 * weight behind: the twips value when it states a non-zero one, otherwise the character count (WPS writes 首行
 * 缩进 2 字符 as `w:firstLineChars="200"`), and a `w:hanging` pulls the other way. Both are stated in 1,147 of
 * the corpus's elements and their ratio is always a font size, so choosing the twips value there draws what
 * WPS drew; the 88 that state a positive `w:firstLine` next to `w:firstLineChars="0"` are the reason the
 * character form does not simply win, as §17.3.1.14 has it.
 *
 * A `w:firstLineChars` with no twips counterpart — 236 elements — is drawn in `em`, which makes it follow this
 * page's own font size rather than the document's. 未取证 how wide WPS makes it, since the size WPS measures
 * the character against is the paragraph mark's, and the model does not carry that as a number.
 */
const firstLineIndent = (p: Paragraph): string => {
  const twips = p.indentFirstLine ?? (p.indentHanging === undefined ? undefined : -p.indentHanging);
  if (twips) return `${twipsToPx(twips)}px`;
  const chars = p.indentFirstLineChars ?? 0;
  return chars > 0 ? `${(chars / 100).toFixed(2)}em` : "";
};

/**
 * The paragraph's line height as one CSS value, in the unit the file wrote it in. `w:line` counts 240ths of a
 * line under `w:auto` (13,396 of the corpus's elements) and twips under `w:exact`/`w:atLeast` (6,802 of them),
 * so drawing every one of them as a multiplier put 固定值 22 磅 at 1.83 倍行距 — a third of the heights in the
 * files WPS itself wrote. A stated `w:line="0"` draws no height at all; the corpus states one once, and no unit
 * turns a zero into a line.
 *
 * `exact` is a height, so it is simply that height; glyphs taller than it are cut, which is what 固定值 does.
 * `atLeast` is a floor, and CSS has no line height that says "this much, or whatever the text needs", so it is
 * drawn as the greater of the stated height and `1em` — the least that cannot clip the paragraph's own glyphs.
 * 未取证 how much taller WPS makes such a line when the text needs more than `1em`, since that is a question
 * about font metrics and this page has none.
 */
const lineHeight = (p: Paragraph): string => {
  if (!p.lineSpacing) return "";
  if (p.lineSpacingRule === "exact") return `${twipsToPx(p.lineSpacing)}px`;
  if (p.lineSpacingRule === "atLeast") return `max(${twipsToPx(p.lineSpacing)}px, 1em)`;
  return (p.lineSpacing / 240).toFixed(2);
};

function renderParagraph(p: Paragraph, doc: WordDocument, counters: Map<string, number>, state: CommentState): string {
  const style: string[] = [];
  if (p.align && ALIGN[p.align]) style.push(`text-align:${ALIGN[p.align]}`);
  if (p.indentLeft) style.push(`margin-left:${twipsToPx(p.indentLeft)}px`);
  if (p.indentRight) style.push(`margin-right:${twipsToPx(p.indentRight)}px`);
  const first = firstLineIndent(p);
  if (first) style.push(`text-indent:${first}`);
  if (p.spacingBefore !== undefined) style.push(`margin-top:${twipsToPx(p.spacingBefore)}px`);
  if (p.spacingAfter !== undefined) style.push(`margin-bottom:${twipsToPx(p.spacingAfter)}px`);
  const lh = lineHeight(p);
  if (lh) style.push(`line-height:${lh}`);
  if (p.pageBreakBefore) style.push("page-break-before:always");
  const heading = /^Heading([1-9])$/i.exec(p.styleId ?? "");
  const tag = heading ? `h${heading[1]}` : "div";
  let prefix = "";
  if (p.numPr) {
    const level = doc.levelOf(p.numPr.numId, p.numPr.ilvl);
    // Each numId is its own list, so a bullet list does not consume the numbers of a numbered one.
    const key = `${p.numPr.numId}:${p.numPr.ilvl}`;
    const seen = counters.get(key) ?? (level?.start ?? 1) - 1;
    counters.set(key, seen + 1);
    const n = seen + 1;
    prefix =
      !level || level.numFmt === "bullet"
        ? `<span class="bullet">${esc(level?.lvlText?.replace(/%\d/g, "") || "•")}</span>`
        : `<span class="bullet">${ordinal(n, level.numFmt)}</span>`;
    style.push(`padding-left:${8 + p.numPr.ilvl * 24}px`);
  }
  const body = wordRuns(p.inlines, state) || "&nbsp;";
  // A paragraph-level revision rides on the paragraph element itself: the break that ends it is not
  // text a run can carry. Class and data attributes only, never inline style, so the editor cannot
  // read a mark back as formatting — and the properties record has to survive the round trip through
  // the page, because the editor re-derives the model from the DOM on every keystroke.
  const mark = p.markRev;
  const fmt = p.revFormat;
  const classes = `para${p.styleName ? " s-" + esc(p.styleName.replace(/\s+/g, "-")) : ""}${mark ? ` rev-mark-${mark.kind}` : ""}${fmt ? " para-rev" : ""}`;
  const carried =
    (mark
      ? ` data-rev-mark="${mark.kind}" data-rev-mark-id="${mark.id}" data-rev-mark-author="${esc(mark.author)}"${mark.date ? ` data-rev-mark-date="${esc(mark.date)}"` : ""}${mark.move ? ` data-rev-mark-move="${mark.move}"` : ""}`
      : "") +
    (fmt
      ? ` data-para-rev="${fmt.id}" data-para-rev-author="${esc(fmt.author)}"${fmt.date ? ` data-para-rev-date="${esc(fmt.date)}"` : ""}` +
        ` data-para-prev="${encodeURIComponent(JSON.stringify(fmt.prev))}"`
      : "");
  const shown = mark ? `${mark.author}${mark.date ? ` · ${mark.date.slice(0, 10)}` : ""} ${mark.move ? (mark.move === "from" ? "移出段落标记" : "移入段落标记") : mark.kind === "del" ? "删除段落标记" : "插入段落标记"}` : fmt ? `${fmt.author}${fmt.date ? ` · ${fmt.date.slice(0, 10)}` : ""} 修改段落格式` : "";
  return `<${tag} class="${classes}"${carried}${shown ? ` title="${esc(shown)}"` : ""}${style.length ? ` style="${style.join(";")}"` : ""}>${prefix}${body}</${tag}>`;
}

function ordinal(n: number, fmt: string): string {
  switch (fmt) {
    case "lowerLetter":
      return String.fromCharCode(96 + ((n - 1) % 26)) + ".";
    case "upperLetter":
      return String.fromCharCode(64 + ((n - 1) % 26)) + ".";
    case "lowerRoman":
      return toRoman(n).toLowerCase() + ".";
    case "upperRoman":
      return toRoman(n) + ".";
    case "none":
      return "";
    default:
      return n + ".";
  }
}

function toRoman(n: number): string {
  const table: [number, string][] = [[10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let out = "";
  for (const [v, s] of table) while (n >= v) ((out += s), (n -= v));
  return out;
}

/** The stroke names Word uses for a border, and how CSS draws the same line. */
const STROKE: Record<string, string> = {
  single: "solid",
  thick: "solid",
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  dotDash: "dashed",
  dash: "dashed",
  dashSmallGap: "dashed",
  dashDotStroked: "dashed",
  dashDot: "dashed",
  dashDotDot: "dashed",
  threeDashSmallGap: "dashed",
  hairline: "solid",
  inset: "inset",
  outset: "outset",
  groove: "groove",
  ridge: "ridge",
};

/** The stroke tokens that name no line, which is what they say. Everything else draws a line. */
const NO_LINE: ReadonlySet<string> = new Set(["none", "hidden", "nil"]);

/**
 * Whether a `w:val` asks for a visible stroke. The file's own answer, not the renderer's fallback: an unknown
 * token still draws, as solid.
 */
export const drawsLine = (val?: string): boolean => val !== undefined && !NO_LINE.has(val);

/**
 * One edge as CSS — a table's rule, a cell's, or the page's own border, which is why the edge is taken by
 * structure rather than by name. A token `STROKE` has no entry for draws a solid line of the width asked for,
 * the least-wrong reading of a stroke the file invented. 未取证: how each of these looks in WPS' own page view
 * was never observed.
 */
function borderCss(property: string, edge?: TableBorder): string {
  if (!edge || !drawsLine(edge.val)) return "";
  const width = Math.max(1, Math.round((edge.size ?? 8) / 6)); // eighths of a point → CSS px
  const color = !edge.color || edge.color === "auto" ? "currentColor" : "#" + edge.color.replace(/^#/, "").toUpperCase();
  return `${property}:${width}px ${STROKE[edge.val] ?? "solid"} ${color};`;
}

const isSet = (edge?: TableBorder): edge is TableBorder => edge !== undefined;

/** `w:tblW`, counted in its own type: twips, or fiftieths of a percent. */
function tableWidth(t: Table): string {
  if (t.widthTwips === undefined) return "";
  if (t.widthType === "pct") return `width:${t.widthTwips / 50}%;`;
  if (t.widthType === "auto") return "";
  return `width:${twipsToPx(t.widthTwips)}px;`;
}

/** `w:jc` and `w:tblInd`: where the table sits inside the text column, and how far it leans right. */
function tableAlign(t: Table): string {
  const indent = t.indentTwips ? `margin-left:${twipsToPx(t.indentTwips)}px;` : "";
  if (t.align === "center") return `margin-left:auto;margin-right:auto;`;
  if (t.align === "right") return `margin-left:auto;`;
  return indent;
}

/** `w:vAlign` on a cell, in the words CSS uses for the same thing. */
const VALIGN: Record<string, string> = { top: "top", center: "center", bottom: "bottom" };

/** `w:textDirection`: the vertical spellings are the ones a page can actually draw. */
function writingMode(direction: string): string {
  if (direction.startsWith("tbRl")) return "vertical-rl";
  if (direction.startsWith("tbLr")) return "vertical-lr";
  return "horizontal-tb";
}

/** A row's own height and hiding. `atLeast` is a floor; `exact` is the band the row may not leave. */
function rowStyle(row?: RowInfo): string {
  if (!row) return "";
  let css = row.hidden ? "display:none;" : "";
  if (row.heightTwips) {
    const px = twipsToPx(row.heightTwips);
    css += row.heightRule === "exact" ? `height:${px}px;max-height:${px}px;` : `min-height:${px}px;`;
  }
  return css;
}

/**
 * What each cell of a table draws as: how many rows it covers, and whether it is the part a merge above
 * already drew. `w:vMerge` says only "restart" or "continue", so the height comes from counting the
 * continues that line up under a restart in the same grid column.
 *
 * The covered cell stays in the page, hidden, rather than being left out. The editor re-derives the model
 * from the cells it finds, and a row that has lost one of them would lose its `w:tcPr` with it — a vertical
 * merge that survives viewing but vanishes on save is the failure this is here to avoid.
 */
function mergeOf(t: Table): { rows: number; hidden: boolean }[][] {
  const startOf = t.rows.map((cells) => {
    let column = 0;
    return cells.map((c) => {
      const at = column;
      column += Math.max(1, c.spanCols ?? 1);
      return at;
    });
  });
  const out = t.rows.map((cells) => cells.map((c) => ({ rows: 1, hidden: c.spanRows === 0 })));
  t.rows.forEach((cells, r) => {
    cells.forEach((c, index) => {
      if (c.spanRows !== 1) return;
      let spanned = 1;
      for (let below = r + 1; below < t.rows.length; below++) {
        const continued = t.rows[below].findIndex((x, i) => x.spanRows === 0 && startOf[below][i] === startOf[r][index]);
        if (continued < 0) break;
        spanned += 1;
      }
      out[r][index].rows = spanned;
    });
  });
  return out;
}

function renderTable(t: Table, doc: WordDocument, state: CommentState): string {
  const table = t.borders;
  const merges = mergeOf(t);
  const css = [
    // A table that draws its own lines wants them shared with the cells, which is what collapsing does;
    // one that has none keeps the page's own spacing.
    table ? "border-collapse:collapse;" : "",
    t.layout === "fixed" ? "table-layout:fixed;" : "",
    tableWidth(t),
    tableAlign(t),
    borderCss("border-top", table?.top),
    borderCss("border-left", table?.left ?? table?.start),
    borderCss("border-bottom", table?.bottom),
    borderCss("border-right", table?.right ?? table?.end),
  ].join("");
  const rows = t.rows
    .map((cells, r) => {
      // A row's own revision rides on the row, as a class plus data attributes the editor reads back —
      // never as inline style, so a mark cannot be mistaken for formatting when the page is re-read.
      const row = cells[0]?.row;
      const rev = row?.rev;
      const carried = rev
        ? ` class="rev-row rev-${rev.kind}" data-row-rev="${rev.kind}" data-row-rev-id="${rev.id}" data-row-rev-author="${esc(rev.author)}"${rev.date ? ` data-row-rev-date="${esc(rev.date)}"` : ""}`
        : "";
      const shown = rev ? `${rev.author}${rev.date ? ` · ${rev.date.slice(0, 10)}` : ""} ${rev.kind === "del" ? "删除行" : "插入行"}` : "";
      const height = rowStyle(row);
      return `<tr${carried}${height ? ` style="${height}"` : ""}${shown ? ` title="${esc(shown)}"` : ""}>${cells
        .map((c, c2) => {
          const own = c.borders;
          const merge = merges[r][c2]!;
          const style = [
            merge.hidden ? "display:none;" : "",
            c.widthTwips ? `width:${c.widthType === "pct" ? `${c.widthTwips / 50}%` : `${twipsToPx(c.widthTwips)}px`};` : "",
            c.shading ? `background:${c.shading};` : "",
            c.valign ? `vertical-align:${VALIGN[c.valign] ?? c.valign};` : "",
            c.noWrap ? "white-space:nowrap;" : "",
            c.textDirection ? `writing-mode:${writingMode(c.textDirection)};` : "",
            // The table's inside lines belong to the seam, so they draw on the cell that faces it, and a
            // border the cell set for itself wins over the one the table set for the grid.
            borderCss("border-top", [own?.top, r > 0 ? table?.insideH : undefined].find(isSet)),
            borderCss("border-bottom", [own?.bottom, r < t.rows.length - 1 ? table?.insideH : undefined].find(isSet)),
            borderCss("border-left", [own?.left ?? own?.start, c2 > 0 ? table?.insideV : undefined].find(isSet)),
            borderCss("border-right", [own?.right ?? own?.end, c2 < cells.length - 1 ? table?.insideV : undefined].find(isSet)),
          ].join("");
          const spans = [c.spanCols && c.spanCols > 1 ? ` colspan="${c.spanCols}"` : "", merge.rows > 1 ? ` rowspan="${merge.rows}"` : ""].join("");
          return `<td${style ? ` style="${style}"` : ""}${spans}>${renderBlocks(c.paragraphs, doc, new Map(), state)}</td>`;
        })
        .join("")}</tr>`;
    })
    .join("");
  return `<table class="doc-table"${css ? ` style="${css}"` : ""}><tbody>${rows}</tbody></table>`;
}

function renderBlocks(blocks: Block[], doc: WordDocument, counters: Map<string, number> = new Map(), state: CommentState = { open: new Set(), numbers: new Map(), done: new Set() }): string {
  return blocks
    .map((b) => (b.kind === "paragraph" ? renderParagraph(b, doc, counters, state) : b.kind === "table" ? renderTable(b, doc, state) : renderCarried(b)))
    .join("");
}

/** How a carried container is drawn. The editor re-creates one with this when the page loses it. */
export const CARRY_STYLE = "border:1px dashed #b3b3b3;border-radius:2px;padding:2px 6px;margin:2px 0;color:#555";

/** The class a carried container's element wears: the page's one signal that a block is not editable. */
export const CARRY_CLASS = "doc-carry";

/**
 * A block the model keeps as the source text it came in with: what is inside it is shown, and it is
 * read-only — the caret cannot enter, and a save writes the bytes back rather than what the page drew. The
 * box is drawn with inline style because the page's own stylesheet is not this module's to change; `data-tag`
 * names what was kept, and nothing else about the element is read back as formatting.
 */
function renderCarried(b: Carried): string {
  const body = esc(b.text) || "（此处内容暂不可编辑）";
  return `<div class="${CARRY_CLASS}" contenteditable="false" data-tag="${esc(b.tag)}" style="${CARRY_STYLE}">${body}</div>`;
}

/** The class a page border's element wears: drawn inside the paper box, and never a block of text. */
export const PAGE_BORDER_CLASS = "doc-page-border";

/**
 * The line `w:pgBorders` asks for, drawn as one absolutely positioned element per stated side inside the paper
 * box. Sides are separate because each names its own stroke and its own gap: a document with a thick bottom and
 * no right side has to draw that, and one rectangle with four CSS borders could not — nor could a box whose
 * missing side still holds the gap of its neighbour.
 *
 * The gap is `w:space` in points, measured against what `w:offsetFrom` names: the paper's edge, or the body,
 * whose edge on a side is that side's margin (plus the binding width on the left, as the padding counts it).
 * Each line spans the paper in the other direction, so four sides meet at the corners the way a frame does.
 */
function renderPageBorders(sec: Section): string {
  const borders = sec.pageBorders;
  if (!borders) return "";
  const gutter = sec.margins.gutter ?? 0;
  const inset = (side: (typeof PAGE_BORDER_SIDES)[number]): number => {
    const space = ((borders.sides[side]?.space ?? 0) * 4) / 3;
    if (borders.offsetFrom !== "text") return space;
    const margin =
      side === "top" ? sec.margins.top : side === "bottom" ? sec.margins.bottom : side === "left" ? sec.margins.left + gutter : sec.margins.right;
    return space + twipsToPx(margin);
  };
  const cross = (side: (typeof PAGE_BORDER_SIDES)[number]): string => (side === "top" || side === "bottom" ? "left:0;right:0" : "top:0;bottom:0");
  return PAGE_BORDER_SIDES.map((side) => {
    const css = borderCss(`border-${side}`, borders.sides[side]);
    return css
      ? `<div class="${PAGE_BORDER_CLASS}" data-side="${side}" contenteditable="false" aria-hidden="true" style="position:absolute;pointer-events:none;${cross(side)};${side}:${Math.round(inset(side))}px;${css}"></div>`
      : "";
  }).join("");
}

export function renderWord(doc: WordDocument): string {
  const sec = doc.sections[0] ?? DEFAULT_SECTION;
  const width = twipsToPx(sec.pageWidthTwips);
  // `w:gutter` is binding width: ECMA-376 §17.3.2.13 adds it to the margin on the binding side, and a section
  // that says nothing has none. This model holds no RTL section, so the binding side is the left.
  const gutter = sec.margins.gutter ?? 0;
  const style = `--page-w:${width}px;--pad-t:${twipsToPx(sec.margins.top)}px;--pad-r:${twipsToPx(sec.margins.right)}px;--pad-b:${twipsToPx(sec.margins.bottom)}px;--pad-l:${twipsToPx(sec.margins.left + gutter)}px`;
  // The distance a band's text is held from the body's, taken from the section's own `w:header`/`w:footer`.
  // Those count from the **page edge**, so what separates a band from the body text is the margin it sits in
  // minus that distance: a header pinned near the edge has a wide gap under it, one pushed down toward the
  // text has almost none. 未取证: WPS' own scrolling view was not observed placing the band, and here the band
  // hangs outside the paper box, so the declared distance can only show as spacing rather than as placement
  // inside the margin. A section that states neither keeps the stylesheet's fixed gap — inventing Word's
  // docDefaults here would draw a number nobody wrote.
  const bandGap = (margin: number, distance: number | undefined): number | undefined =>
    distance === undefined ? undefined : Math.max(0, twipsToPx(margin - distance));
  const hf = (which: "header" | "footer"): string => {
    const area = which === "header" ? doc.header : doc.footer;
    if (!area) return "";
    // `.hf-p` is `white-space: pre-wrap`, so a tab the part holds stays a tab on screen and comes back as
    // one when the line is read out of the page again. A paragraph that holds a field is drawn piece by
    // piece, because the field is not text and the order the two sit in is the file's, not the reader's.
    const paras =
      area.paragraphs
        .map((_, i) => {
          const segs = area.segments?.[i] ?? [{ text: area.paragraphs[i] ?? "" }];
          return `<div class="hf-p">${segs.map((s) => (s.carry ? hfFieldTag(s.carry) : esc(s.text ?? ""))).join("") || "<br>"}</div>`;
        })
        .join("") || '<div class="hf-p"><br></div>';
    const name = which === "header" ? "页眉" : "页脚";
    const tip = area.editable ? `点击编辑${name}` : `此${name}含表格或图片等富内容，暂不支持编辑`;
    // The command the band carries, drawn outside the paragraphs so reading them back never picks it up as
    // text. Inline styles again: the page's stylesheet is not this module's to change.
    const bar = area.editable
      ? `<div class="hf-bar" contenteditable="false" style="${HF_BAR_STYLE}"><button class="hf-page" data-hf-page="${which}" style="${HF_BTN_STYLE}" title="在此${name}插入页码域">插入页码</button></div>`
      : "";
    // `hf-bottom` is the rule that puts a footer's padding under its text instead of above it. The band
    // travels with the page's own box variables: it sits beside `.page`, not inside it, so without them the
    // stylesheet answers with its 794px fallback and a landscape section gets a portrait header — same page,
    // two widths. The horizontal padding is overridden to the section's margins so the band's text starts at
    // the same column as the body's.
    const gap = bandGap(which === "header" ? sec.margins.top : sec.margins.bottom, which === "header" ? sec.margins.header : sec.margins.footer);
    const spacing = gap === undefined ? "" : `;padding-${which === "header" ? "bottom" : "top"}:${gap}px`;
    return `<div class="hf hf-${which}${which === "footer" ? " hf-bottom" : ""}" data-hf="${which}" style="${style};padding-left:var(--pad-l);padding-right:var(--pad-r)${spacing}" title="${tip}">${paras}${bar}</div>`;
  };
  const borders = renderPageBorders(sec);
  // The border elements come last, inside the paper box: the editor addresses a block by its position among
  // `.page`'s children, and a leading decoration would move every one of those positions by a step.
  return `${hf("header")}<div class="page" style="${style}${borders ? ";position:relative" : ""}">${renderBlocks(doc.blocks, doc, new Map(), {
    open: new Set(),
    numbers: new Map(doc.comments.map((c, i) => [c.id, i + 1])),
    done: new Set(doc.comments.filter((c) => c.resolved === true).map((c) => c.id)),
  })}${borders}</div>${hf("footer")}`;
}

function renderSlideParagraph(p: SlideParagraph): string {
  const css: string[] = [];
  if (p.align && ALIGN[p.align]) css.push(`text-align:${ALIGN[p.align]}`);
  const runs = p.runs.map(renderInline).join("");
  return `<div class="sp-p" style="${css.join(";")};padding-left:${p.level * 18}px">${p.bullet ? `<span class="bullet">${esc(p.bullet)}</span>` : ""}${runs || "<br>"}</div>`;
}

function renderShape(s: Shape, pres: Presentation, idx?: number): string {
  // Inline position:absolute keeps exported standalone HTML faithful without a stylesheet.
  const common = `position:absolute;left:${s.x}px;top:${s.y}px;width:${s.width}px;height:${s.height}px`;
  const rot = s.rotation ? `transform:rotate(${s.rotation}deg);` : "";
  // Only top-level shapes carry an identity: the editor selects and drags those.
  const pick = idx === undefined ? "" : ` data-idx="${idx}"`;
  if (s.kind === "picture") {
    return s.src ? `<img class="shape"${pick} style="${common};${rot}object-fit:fill" src="${s.src}" alt="${esc(s.name)}">` : "";
  }
  if (s.kind === "table") {
    const html = (s.rows ?? [])
      .map((r) => `<tr>${r.map((c) => `<td${c.fill ? ` style="background:${c.fill}"` : ""}>${c.runs.map(renderInline).join("") || esc(c.text)}</td>`).join("")}</tr>`)
      .join("");
    return `<div class="shape st-table"${pick} style="${common};${rot}"><table>${html}</table></div>`;
  }
  if (s.kind === "group") {
    return `<div class="shape"${pick} style="${common};${rot}">${(s.children ?? []).map((c) => renderShape(c, pres)).join("")}</div>`;
  }
  const style = [common, rot];
  if (s.geometry === "line") {
    // A line is a zero-height box in the model; its stroke is drawn on that line.
    const width = Math.max(1, s.lineWidth ?? 1);
    const color = s.lineColor ?? "#000000";
    return `<div class="shape"${pick} style="${style.join(";")}"><div style="position:absolute;left:0;right:0;top:0;transform:translateY(-50%);border-top:${width}px solid ${color}"></div></div>`;
  }
  if (s.fill) style.push(`background:${s.fill}`);
  if (s.lineColor) style.push(`border:${Math.max(1, s.lineWidth ?? 1)}px solid ${s.lineColor}`);
  if (s.geometry && s.geometry !== "rect") style.push(`clip-path:${clipFor(s.geometry)}`);
  const text = (s.paragraphs ?? []).map(renderSlideParagraph).join("");
  return `<div class="shape"${pick} style="${style.join(";")}"><div class="shape-text">${text}</div></div>`;
}

function clipFor(geom: string): string {
  switch (geom) {
    case "ellipse":
      return "ellipse(50% 50%)";
    case "roundRect":
      return "inset(0 round 12px)";
    case "triangle":
      return "polygon(50% 0,100% 100%,0 100%)";
    case "parallelogram":
      return "polygon(12% 0,100% 0,88% 100%,0 100%)";
    case "diamond":
      return "polygon(50% 0,100% 50%,50% 100%,0 50%)";
    case "rightArrow":
      return "polygon(0 30%,60% 30%,60% 0,100% 50%,60% 100%,60% 70%,0 70%)";
    case "star5":
      return "polygon(50% 0%,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)";
    case "chevron":
      return "polygon(25% 0,100% 0,75% 100%,0 100%)";
    default:
      return "";
  }
}

export function renderSlideHtml(pres: Presentation, index: number): string {
  const slide = pres.slides[index];
  if (!slide) return "";
  const bg = slide.background ? `background:${slide.background};` : "";
  return `<div class="slide" style="width:${pres.slideWidthPx}px;height:${pres.slideHeightPx}px;${bg}">${slide.shapes
    .map((s, i) => renderShape(s, pres, i))
    .join("")}</div>`;
}
