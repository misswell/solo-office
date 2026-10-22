import { childSpans, elementContent } from "./opc.js";
import type { Block, Paragraph, Table, TableCell } from "./docx.js";

/**
 * The table commands WPS' 表格工具栏 has and this editor did not: insert or delete a row or a column,
 * merge cells into one, split one back into the grid it came from.
 *
 * They work on the model in place and say where the caret belongs afterwards. Nothing here renders or signs
 * a change — a row inserted while 修订 is on is recorded by the editor, not here, because the author and the
 * `w:id` belong to whoever is holding the pen.
 *
 * The geometry these preserve is the one ECMA-376 describes: `w:gridSpan` says how many grid columns a cell
 * covers and its row simply holds fewer `w:tc` elements, while `w:vMerge` keeps a cell in every row it covers
 * and marks all but the first as a continuation. So a cell's grid column is never its index in the row: it is
 * the sum of the spans before it, and every command here derives positions from that.
 */

/** The usable width between A4 margins, the band a new grid is sized from. */
const USABLE_TWIPS = 9026;

/** A cell addressed the way the page and the model agree: its row, and its place among that row's cells. */
export interface CellWhere {
  row: number;
  col: number;
}

/** What a command reports: the caret it asked for, or why the shape on the page could not take it. */
export type TableEdit = { ok: true; at: CellWhere } | { ok: false; reason: string };

/** 插入单元格的方向: the cells already there move right out of the new cell's way, or down. */
export type CellShift = "right" | "down";

const refused = (reason: string): TableEdit => ({ ok: false, reason });
const landed = (row: number, col: number): TableEdit => ({ ok: true, at: { row, col } });

/** How many grid columns a cell covers: `w:gridSpan`, one for an ordinary cell. */
const spanOf = (cell: TableCell): number => Math.max(1, cell.spanCols ?? 1);

const blank = (): Paragraph => ({ kind: "paragraph", inlines: [] });

/** Whether a block holds anything a reader would miss: a carried container counts, whatever its text. */
const hasText = (block: Block): boolean =>
  block.kind !== "paragraph" || block.inlines.some((i) => i.text !== "" || i.image !== undefined || i.break === true);

/** The grid column each of a row's cells starts at, in the order they sit. */
export function columnStarts(row: TableCell[]): number[] {
  let at = 0;
  return row.map((cell) => {
    const start = at;
    at += spanOf(cell);
    return start;
  });
}

/** The widest row wins: a table's grid is as wide as the row that covers the most columns. */
export function gridColumns(t: Table): number {
  return Math.max(1, ...t.rows.map((row) => row.reduce((n, cell) => n + spanOf(cell), 0)));
}

/** The place in a row of the cell that covers grid column `at`, or -1 when the row does not reach it. */
export function cellCovering(row: TableCell[], at: number): number {
  const starts = columnStarts(row);
  return starts.findIndex((start, i) => start <= at && start + spanOf(row[i]) > at);
}

/** The rectangle a cell covers — the rows of a vertical merge and the grid columns of a horizontal one. */
export function cellRect(t: Table, where: CellWhere) {
  const row = t.rows[where.row];
  const cell = row?.[where.col];
  if (!cell) return undefined;
  const starts = columnStarts(row);
  const start = starts[where.col];
  let bottom = where.row;
  if (cell.spanRows === 1) {
    for (let below = where.row + 1; below < t.rows.length; below++) {
      const at = columnStarts(t.rows[below]).indexOf(start);
      if (at < 0 || t.rows[below][at].spanRows !== 0) break;
      bottom = below;
    }
  }
  return { c0: start, c1: start + spanOf(cell), r0: where.row, r1: bottom + 1 };
}

/**
 * The cell a vertical merge band starts at. `w:vMerge` keeps a `w:tc` in every row the band covers, so the page
 * can put the caret on a continuation; the band is only named by its first row. Every rectangle below — the one
 * 合并单元格 takes, the one 选中整行 lights up — has to start there, or a selection would cover half a cell.
 */
export function bandAnchor(t: Table, where: CellWhere): CellWhere {
  const row = t.rows[where.row] ?? [];
  const cell = row[where.col];
  if (!cell || cell.spanRows !== 0) return where;
  const start = columnStarts(row)[where.col];
  const at = (r: number): number => (t.rows[r] ? cellCovering(t.rows[r], start) : -1);
  let top = where.row;
  while (top > 0) {
    const above = at(top - 1);
    if (above < 0 || t.rows[top - 1][above].spanRows !== 0) break;
    top -= 1;
  }
  const col = at(top - 1);
  // 走到带顶就停：上面那格是起点（`spanRows` 1）才接着往上挪一格，再往上就没有这一带的份了。
  if (top > 0 && col >= 0 && t.rows[top - 1][col].spanRows === 1) return { row: top - 1, col };
  return { row: top, col: at(top) < 0 ? where.col : at(top) };
}

/** The rectangle a cell really paints — its `w:vMerge` band and its `w:gridSpan` both taken whole. */
function paintedRect(t: Table, where: CellWhere): Rect | undefined {
  return cellRect(t, bandAnchor(t, where));
}

/** A run of cells as the page selects them: the two corner cells of the rectangle they cover. */
export interface CellRange {
  from: CellWhere;
  to: CellWhere;
}

/** What 选中整行/整列/整格/整个表格 answers: the range to light up, or why this shape does not have one. */
export type CellRangeOutcome = { ok: true; range: CellRange } | { ok: false; reason: string };

type Rect = NonNullable<ReturnType<typeof cellRect>>;

/**
 * 选中整格 / 整行 / 整列 / 整个表格 (`wps.KSelection.SelectCell`, `SelectRow`, `SelectColumn`, `SelectTable`):
 * which cells light up, in the same (row, col) rectangle `mergeHere` already understands, so the highlight and
 * any command that later acts on the selection mean the same set of cells.
 *
 * A merged cell is always taken whole: `w:gridSpan` and `w:vMerge` make the painted cell bigger than the row's
 * own slice of it, and a selection stopping at the row's edge would light up half of it. Which cells WPS itself
 * paints for these four commands is **未取证** — the strings only prove the commands exist.
 */
export function cellRangeFor(
  t: Table,
  where: CellWhere,
  kind: "cell" | "row" | "column" | "table",
): CellRangeOutcome {
  const row = t.rows[where.row];
  if (!row?.[where.col]) return { ok: false, reason: "这一格已经不在表格里" };
  const own = paintedRect(t, where);
  if (!own) return { ok: false, reason: "这一格已经不在表格里" };
  if (kind === "cell") return rectRange(t, [own]);
  if (kind === "row") return rectRange(t, row.map((_, i) => paintedRect(t, { row: where.row, col: i })).filter(isRect));
  // 整个表格: 每一格里画出来的矩形并起来。带（`w:vMerge`）在行之间本来就会互相比邻，所以整张表总能凑成一个矩形，
  // 没有"选不出矩形"这一说 —— 与整列不同，那一条只在光标那一列的宽度上要求对齐。
  if (kind === "table")
    return rectRange(t, t.rows.flatMap((r, at) => r.map((_, i) => paintedRect(t, { row: at, col: i })).filter(isRect)));
  // 整列: the grid column the caret's own cell starts at, through every row that reaches it. Taking the cell's
  // left seam rather than "the column under the caret's text" is what makes a wide cell answer for its own span.
  const column = own.c0;
  const strip: Rect[] = [];
  for (let r = 0; r < t.rows.length; r += 1) {
    const at = cellCovering(t.rows[r], column);
    if (at < 0) continue;
    const rect = paintedRect(t, { row: r, col: at });
    if (rect) strip.push(rect);
  }
  if (!strip.length) return { ok: false, reason: "这一列上没有可选中的单元格" };
  // 矩形选不出来就拒绝：这一列上只要有格子横向合并过，它的网格就比光标那格宽，
  // 而 from/to 只能表达一个轴对齐的矩形 —— 于是会把没人选中的左邻格一起点亮。与 `insertCellDown` 同一取舍。
  if (strip.some((r) => r.c0 !== own.c0 || r.c1 !== own.c1))
    return { ok: false, reason: "这一列上有横向合并的单元格，整列选不出一个矩形" };
  const top = Math.min(...strip.map((r) => r.r0));
  const bottom = Math.max(...strip.map((r) => r.r1)) - 1;
  const from = cellCovering(t.rows[top] ?? [], column);
  const to = cellCovering(t.rows[bottom] ?? [], column);
  if (from < 0 || to < 0) return { ok: false, reason: "这一列上没有可选中的单元格" };
  return { ok: true, range: { from: { row: top, col: from }, to: { row: bottom, col: to } } };
}

const isRect = (r: Rect | undefined): r is Rect => r !== undefined;

/** The rectangle around a set of painted cells, addressed by the two cells that hold its corners. */
function rectRange(t: Table, rects: Rect[]): CellRangeOutcome {
  if (!rects.length) return { ok: false, reason: "这一格已经不在表格里" };
  const r0 = Math.min(...rects.map((r) => r.r0));
  const r1 = Math.max(...rects.map((r) => r.r1)) - 1;
  const c0 = Math.min(...rects.map((r) => r.c0));
  const c1 = Math.max(...rects.map((r) => r.c1));
  const from = cellCovering(t.rows[r0] ?? [], c0);
  const to = cellCovering(t.rows[r1] ?? [], c1 - 1);
  if (from < 0 || to < 0) return { ok: false, reason: "这一格已经不在表格里" };
  return { ok: true, range: { from: { row: r0, col: from }, to: { row: r1, col: to } } };
}

/**
 * Every cell a range paints, in document order and one entry per cell: a vertical band is listed at its top
 * only, because that is the single `w:tc` whose text the page shows — the continuation halves are drawn
 * `display:none`, so a selection cannot be put on them and a count that listed them would double the cell.
 *
 * This is the set the highlight lights up and the set 合并单元格 acts on, so both read it from here rather than
 * counting the two corners, which name a rectangle and not the cells inside it.
 */
export function cellsInRect(t: Table, range: CellRange): CellWhere[] {
  const a = paintedRect(t, range.from);
  const b = paintedRect(t, range.to);
  if (!a || !b) return [];
  const c0 = Math.min(a.c0, b.c0);
  const c1 = Math.max(a.c1, b.c1);
  return cellsBetween(t, Math.min(range.from.row, range.to.row), Math.max(range.from.row, range.to.row), (row, i) => {
    const starts = columnStarts(row);
    return starts[i] < c1 && starts[i] + spanOf(row[i]) > c0;
  });
}

/** Every painted cell of a table, in document order — a vertical band once, at its top. */
export function allCells(t: Table): CellWhere[] {
  return cellsBetween(t, 0, t.rows.length - 1, () => true);
}

/** A cell's own four edges, in the order the border commands name them. */
export type CellSide = "top" | "left" | "bottom" | "right";

/**
 * 「外侧框线」在格选区里要写哪几条边：逐格只取朝向选区外的那几条，块内相邻的缝一条不写。
 * 选区按整格算——竖合并的一带由起头那一格答，合并格伸出选区的那一段也算在块里，因为选区切不开一格。
 * WPS 自己在这一击写出的字节未取证；这一条按 §17.4.77 的 `w:tcBorders` 与「外侧」的字面义立据。
 */
export function cellOuterSides(t: Table, cells: CellWhere[]): { where: CellWhere; sides: CellSide[] }[] {
  const covered = new Set<string>();
  const rects = new Map<CellWhere, NonNullable<ReturnType<typeof cellRect>>>();
  for (const where of cells) {
    const rect = cellRect(t, where);
    if (!rect) continue;
    rects.set(where, rect);
    for (let r = rect.r0; r < rect.r1; r += 1)
      for (let c = rect.c0; c < rect.c1; c += 1) covered.add(`${r},${c}`);
  }
  const inside = (r: number, c: number): boolean => covered.has(`${r},${c}`);
  const columns = gridColumns(t);
  return cells.flatMap((where) => {
    const rect = rects.get(where);
    if (!rect) return [];
    const sides: CellSide[] = [];
    // 顶/底按这一格横跨的每一列看：只要有一列上方不在选区里，这条边就有一段露在外面。
    const exposedRow = (r: number): boolean => {
      for (let c = rect.c0; c < rect.c1; c += 1) if (!inside(r, c)) return true;
      return false;
    };
    // 左/右按这一格纵跨的每一行看，理由同上。
    const exposedColumn = (c: number): boolean => {
      for (let r = rect.r0; r < rect.r1; r += 1) if (!inside(r, c)) return true;
      return false;
    };
    if (rect.r0 === 0 || exposedRow(rect.r0 - 1)) sides.push("top");
    if (rect.r1 >= t.rows.length || exposedRow(rect.r1)) sides.push("bottom");
    if (rect.c0 === 0 || exposedColumn(rect.c0 - 1)) sides.push("left");
    if (rect.c1 >= columns || exposedColumn(rect.c1)) sides.push("right");
    return [{ where, sides }];
  });
}

/**
 * Every cell the page draws between two cells in its own order — exactly what one Range would light up if it
 * were pulled from the first to the last. Put against `cellsInRect` it answers whether a single Range can paint
 * a selection: a row can, a column cannot (the cells do not sit next to each other), and an editable page keeps
 * only one Range, so what a Range cannot reach has to be lit cell by cell.
 */
export function cellsInSweep(t: Table, from: CellWhere, to: CellWhere): CellWhere[] {
  const down = from.row <= to.row;
  const r0 = down ? from.row : to.row;
  const r1 = down ? to.row : from.row;
  const at0 = down ? from.col : to.col;
  const at1 = down ? to.col : from.col;
  return cellsBetween(t, r0, r1, (_row, i, r) => i >= (r0 === r ? at0 : 0) && i <= (r1 === r ? at1 : _row.length));
}

/** The band-anchored cells a predicate picks out, walked the way the page walks them. */
function cellsBetween(t: Table, r0: number, r1: number, keep: (row: TableCell[], i: number, r: number) => boolean): CellWhere[] {
  const out: CellWhere[] = [];
  const seen = new Set<string>();
  for (let r = r0; r <= r1; r++) {
    const row = t.rows[r] ?? [];
    row.forEach((_, i) => {
      if (!keep(row, i, r)) return;
      const anchor = bandAnchor(t, { row: r, col: i });
      const key = `${anchor.row},${anchor.col}`;
      if (seen.has(key)) return;
      seen.add(key);
      out.push(anchor);
    });
  }
  return out;
}

/** The grid columns a table's rows use, twips each — from the file's grid, or measured off its cells. */
function gridWidths(t: Table): number[] {
  const columns = gridColumns(t);
  if (t.columnWidths?.length === columns) return [...t.columnWidths];
  const carried = t.gridXml ? childSpans(elementContent(t.gridXml)) : [];
  if (carried.length === columns)
    return carried.map((k) => Number(/\bw:w="(-?\d+(?:\.\d+)?)"/.exec(k.xml)?.[1] ?? 0));
  const first = t.rows[0] ?? [];
  const starts = columnStarts(first);
  return Array.from({ length: columns }, (_, column) => {
    const at = first.findIndex((cell, i) => starts[i] <= column && starts[i] + spanOf(cell) > column);
    return at < 0 ? 0 : Math.round((first[at].widthTwips ?? 0) / spanOf(first[at]));
  });
}

/** The carried `w:gridCol` elements of the file's own grid, indexed by grid column. */
function gridCarried(t: Table): (string | undefined)[] {
  const kids = t.gridXml ? childSpans(elementContent(t.gridXml)).filter((k) => k.tag === "gridCol") : [];
  return kids.length === gridColumns(t) ? kids.map((k) => k.xml) : [];
}

/**
 * What the table's grid adds up to right now: the file's own `w:gridCol` set first, then the model's mirror of
 * it. Not `gridWidths` — a command reaches here after it has already moved the row's cells, and that function
 * falls back to measuring the columns off the rows, which would report the grid the command is *making*.
 */
function statedGridSum(t: Table): number {
  const kids = t.gridXml ? childSpans(elementContent(t.gridXml)).filter((k) => k.tag === "gridCol") : [];
  if (kids.length)
    return Math.round(kids.reduce((sum, k) => sum + Number(/\bw:w="(-?\d+(?:\.\d+)?)"/.exec(k.xml)?.[1] ?? 0), 0));
  return Math.round((t.columnWidths ?? []).reduce((sum, w) => sum + w, 0));
}

/**
 * The table's grid, re-laid out: a carried `w:gridCol` is reused wherever a column survived unchanged.
 *
 * The grid is the table's width: when a command leaves it summing to something else, the table's own
 * `w:tblW` is restated to the new sum. A command that keeps the sum — splitting a column in two, inserting a
 * cell to the right — costs that property no bytes at all.
 */
function setGrid(t: Table, widths: number[], carried: (string | undefined)[]): void {
  const before = statedGridSum(t);
  const after = Math.round(widths.reduce((sum, w) => sum + w, 0));
  t.columnWidths = widths.map((w) => Math.round(w));
  t.columns = widths.length;
  if (t.gridXml) {
    const prefix = /<([A-Za-z_][-\w.]*):gridCol/.exec(t.gridXml)?.[1] ?? "w";
    t.gridXml = withChildren(
      t.gridXml,
      widths.map((w, i) => carried[i] ?? `<${prefix}:gridCol w:w="${Math.round(w)}"/>`),
    );
  }
  if (before !== after) restateTableWidth(t, after);
}

/**
 * 网格换和之后，把表自己那句 `w:tblW w:w="…"` 改到新的和上，只动 `w:w` 这一个值，元素其余部分原样留着。
 *
 * 授权口径 46 件里 58 枚 `w:tbl` 全带 `w:tblPr/w:tblW w:type="dxa"`，而且 58/58 的 `w:w` 正等于本表
 * `w:tblGrid/w:gridCol` 之和（`.scratch/table-width-audience.py`）——件写「整表多宽」就是照网格算一遍。
 * 多一条列还留着旧的数，写出的是 WPS 不写的那一种形状；而页面按 `w:tblW` 画这只盒
 * （`render.ts:tableWidth`），纸上的表会被压回插入前那条宽度，看着像没插进去。
 *
 * 只改**件自己说过 `dxa`** 的那一枚：`pct` 量的是纸面可用宽度的百分比、`auto` 与 `nil` 说的都是「不按宽度摆」，
 * 三者都不由网格的和来定；`w:type` 没写就是写出器认的 `dxa`。本来没说过宽度的表不给它造一枚，与写出器那句
 * 「没声明就什么都不写」同一条口径。
 */
function restateTableWidth(t: Table, twips: number): void {
  const carried = t.tblPrXml ? childSpans(elementContent(t.tblPrXml)).find((k) => k.tag === "tblW")?.xml : undefined;
  const unit = carried ? /\bw:type="(\w+)"/.exec(carried)?.[1] ?? "dxa" : t.widthType ?? "dxa";
  if (unit !== "dxa") return;
  if (t.widthTwips === undefined && !carried) return;
  t.widthTwips = twips;
  if (!carried || !t.tblPrXml) return;
  const next = /\bw:w="[^"]*"/.test(carried)
    ? carried.replace(/\bw:w="[^"]*"/, `w:w="${twips}"`)
    : `<${/^<([A-Za-z_][-\w.]*):/.exec(carried)?.[1] ?? "w"}:tblW w:w="${twips}"/>`;
  t.tblPrXml = t.tblPrXml.replace(carried, next);
}

// ------------------------------------------------------- carried property sets

/** `CT_TcPrBase`'s children in the order the schema gives them, so a rewritten one lands in its place. */
const TC_PR_ORDER = [
  "cnfStyle",
  "tcW",
  "gridSpan",
  "hMerge",
  "vMerge",
  "tcBorders",
  "shd",
  "noWrap",
  "tcMar",
  "textDirection",
  "tcFitText",
  "vAlign",
  "hideMark",
  "cellIns",
  "cellDel",
  "cellMerge",
  "tcPrChange",
];

const rankOf = (tag: string): number => {
  const at = TC_PR_ORDER.indexOf(tag);
  return at < 0 ? TC_PR_ORDER.length : at;
};

/**
 * The same element, carrying these children: its own open tag is reused, so a prefix declared on it survives.
 * An element left with nothing is dropped, which is what an empty `w:tcPr` amounts to.
 */
function withChildren(xml: string, kids: string[]): string {
  const gt = xml.indexOf(">");
  if (gt < 0 || !kids.length) return "";
  const head = xml.slice(0, gt + 1);
  const open = head.endsWith("/>") ? `${head.slice(0, -2)}>` : head;
  const name = /^<([^/\s>]+)/.exec(open)?.[1];
  return name ? `${open}${kids.join("")}</${name}>` : xml;
}

/** What a structural command owns inside a cell: its width, its column span, its place in a vertical merge. */
interface CellShape {
  widthTwips?: number;
  spanCols?: number;
  spanRows?: number;
}

/**
 * Puts a structural change onto a cell: the typed fields the page draws, and — when the file wrote this cell's
 * `w:tcPr` itself — the three children these commands own (`w:tcW`, `w:gridSpan`, `w:vMerge`). Every other
 * child comes back byte for byte, so a cell's borders, shading, `w:tcMar` and a pending `w:tcPrChange` survive
 * sitting next to a merge. `spanRows` is touched when the key is present, `undefined` meaning "no merge".
 *
 * A width the file counted in `pct` is left as it was: percentages measure against the text column, not
 * against `w:tblGrid`, so twips arithmetic would silently change what the cell says.
 */
function applyCellProps(cell: TableCell, props: CellShape): void {
  if (props.widthTwips !== undefined) cell.widthTwips = props.widthTwips;
  if (props.spanCols !== undefined) cell.spanCols = props.spanCols > 1 ? props.spanCols : undefined;
  if ("spanRows" in props) cell.spanRows = props.spanRows;
  if (!cell.tcPrXml) return;

  const kids = childSpans(elementContent(cell.tcPrXml));
  const prefix = /<([A-Za-z_][-\w.]*):tcW[\s/>]/.exec(cell.tcPrXml)?.[1] ?? "w";
  const owned = new Map<string, string>();

  const carriedWidth = kids.find((k) => k.tag === "tcW")?.xml ?? "";
  const unit = /\bw:type="(\w+)"/.exec(carriedWidth)?.[1];
  if (props.widthTwips !== undefined && (!carriedWidth || unit === undefined || unit === "dxa" || unit === "nil"))
    owned.set("tcW", `<${prefix}:tcW w:w="${Math.round(props.widthTwips)}" w:type="${unit ?? "dxa"}"/>`);
  owned.set("gridSpan", cell.spanCols ? `<${prefix}:gridSpan w:val="${cell.spanCols}"/>` : "");
  owned.set(
    "vMerge",
    cell.spanRows === undefined ? "" : cell.spanRows === 1 ? `<${prefix}:vMerge w:val="restart"/>` : `<${prefix}:vMerge/>`,
  );

  let kept: { tag: string; xml: string }[] = [];
  for (const k of kids) {
    const next = owned.get(k.tag);
    if (next === undefined || next === k.xml) kept.push(k);
  }
  for (const [tag, xml] of [...owned]) {
    if (xml === "" || kept.some((k) => k.tag === tag)) continue;
    const at = kept.findIndex((k) => rankOf(k.tag) > rankOf(tag));
    kept = at < 0 ? [...kept, { tag, xml }] : [...kept.slice(0, at), { tag, xml }, ...kept.slice(at)];
  }
  cell.tcPrXml = withChildren(cell.tcPrXml, kept.map((k) => k.xml));
}

const makeCell = (widthTwips: number, row?: TableCell["row"]): TableCell => ({
  paragraphs: [blank()],
  widthTwips,
  ...(row ? { row } : {}),
});

/** A cell's paragraphs with the empty ones dropped — what a merge collects into its anchor. */
const contentOf = (cell: TableCell): Block[] => cell.paragraphs.filter(hasText);

// --------------------------------------------------------------- the commands

/**
 * A new row before `at`. Its shape copies the row it sits next to, because every grid column has to be
 * covered by something; its cells come empty. Where a merge crosses the boundary — the row below continues a
 * merge at the same grid column — the new cell continues it too, so inserting inside a merged band makes it
 * taller instead of breaking it. A new row claims no height, no header repeat and is not hidden: those are
 * one row's own promises, and copying them would keep on hiding a row nobody asked to hide.
 */
function insertRow(t: Table, at: number): TableEdit {
  if (!t.rows.length) return refused("这里没有表格");
  if (at < 0 || at > t.rows.length) return refused("这一行不在表格里");
  const template = t.rows[at - 1] ?? t.rows[at] ?? t.rows[0];
  const widths = gridWidths(t);
  const starts = columnStarts(template);
  const below = t.rows[at];
  const belowStarts = below ? columnStarts(below) : [];
  const row = template.map((cell, i) => {
    // A new row is not the row above: it gets that row's shape and width, but none of its promises, and its
    // own `RowInfo` — two rows sharing one record would let a decision on one sign the other.
    const made = makeCell(widths[starts[i]] ?? 0);
    if (cell.spanCols) made.spanCols = cell.spanCols;
    const belowAt = below ? belowStarts.indexOf(starts[i]) : -1;
    if (cell.spanRows === 0 || (belowAt >= 0 && below![belowAt].spanRows === 0)) made.spanRows = 0;
    return made;
  });
  t.rows.splice(at, 0, row);
  t.columns = gridColumns(t);
  return landed(at, 0);
}

/**
 * A new grid column before `at`. The column it goes in front of is split in two so the table keeps the width
 * it had; at the right edge a new column matches its neighbour instead, and the table grows by one. Which cell
 * a seam touches decides what happens to it: a cell that *starts* on the seam hands the new column its width
 * and keeps the rest of the host column, while a cell the seam runs *through* gains a column of span and no
 * width at all, because the column it now covers is one it already owned. A vertical merge is not touched — it
 * occupies a column, not a place between two.
 */
function insertColumn(t: Table, at: number, caret: CellWhere): TableEdit {
  if (!t.rows.length) return refused("这里没有表格");
  const columns = gridColumns(t);
  if (at < 0 || at > columns) return refused("这一列不在表格里");
  const widths = gridWidths(t);
  const host = at < columns ? at : Math.max(0, columns - 1);
  const hostWidth = widths[host] ?? Math.floor(USABLE_TWIPS / Math.max(1, columns));
  const halved = at < columns;
  const added = halved ? Math.floor(hostWidth / 2) : Math.max(1, Math.round(hostWidth));
  const carried = [...gridCarried(t)];
  if (halved) {
    widths[host] = hostWidth - added;
    carried.splice(host, 1, undefined);
  }
  widths.splice(at, 0, added);
  carried.splice(at, 0, undefined);

  let land = caret.col;
  t.rows.forEach((row, r) => {
    const starts = columnStarts(row);
    const exact = starts.indexOf(at);
    if (exact >= 0) {
      const cell = row[exact];
      if (halved && cell.widthTwips !== undefined) applyCellProps(cell, { widthTwips: Math.max(0, cell.widthTwips - added) });
      row.splice(exact, 0, makeCell(added, cell.row));
      if (r === caret.row) land = exact;
      return;
    }
    const covering = starts.findIndex((start, i) => start < at && start + spanOf(row[i]) > at);
    if (covering >= 0) {
      applyCellProps(row[covering], { spanCols: spanOf(row[covering]) + 1 });
      if (r === caret.row) land = covering;
    } else if (!halved) {
      row.push(makeCell(added, row[0]?.row));
      if (r === caret.row) land = row.length - 1;
    }
  });
  setGrid(t, widths, carried);
  return landed(caret.row, land);
}

/**
 * A row out of the grid. A vertical merge whose start sat in it is picked up by the first row below that still
 * continues it, so no merge ends up with continuations nobody anchored.
 */
function deleteRow(t: Table, index: number): TableEdit {
  if (index < 0 || index >= t.rows.length) return refused("这一行不在表格里");
  if (t.rows.length === 1) return refused("表格至少得有一行");
  t.rows.splice(index, 1);
  repairMerges(t);
  t.columns = gridColumns(t);
  return landed(Math.min(index, t.rows.length - 1), 0);
}

/** Turns a continuation whose anchor is gone into the start of the merge it now belongs to. */
export function repairMerges(t: Table): number {
  let promoted = 0;
  t.rows.forEach((row, r) => {
    const starts = columnStarts(row);
    const above = r === 0 ? [] : columnStarts(t.rows[r - 1]);
    row.forEach((cell, i) => {
      if (cell.spanRows !== 0) return;
      const at = above.indexOf(starts[i]);
      const anchor = at < 0 ? undefined : t.rows[r - 1][at];
      if (anchor && (anchor.spanRows === 0 || anchor.spanRows === 1)) return;
      applyCellProps(cell, { spanRows: 1 });
      promoted += 1;
    });
  });
  return promoted;
}

/**
 * A grid column out of the table. A cell that covered only it goes with it; a cell that spanned it loses one
 * column of span and one of width, so each row keeps tiling what is left of the grid.
 */
export function deleteColumn(t: Table, at: number, caret: CellWhere): TableEdit {
  const columns = gridColumns(t);
  if (at < 0 || at >= columns) return refused("这一列不在表格里");
  if (columns === 1) return refused("表格至少得有一列");
  const widths = gridWidths(t);
  const removed = widths[at] ?? 0;
  const carried = [...gridCarried(t)];
  let land = caret.col;
  t.rows.forEach((row, r) => {
    const starts = columnStarts(row);
    const exact = starts.indexOf(at);
    const covering = cellCovering(row, at);
    if (covering < 0) return;
    const cell = row[covering];
    if (covering === exact && spanOf(cell) === 1) {
      row.splice(exact, 1);
      if (r === caret.row) land = Math.min(caret.col, Math.max(0, row.length - 1));
      return;
    }
    applyCellProps(cell, {
      spanCols: spanOf(cell) - 1,
      ...(cell.widthTwips !== undefined ? { widthTwips: Math.max(0, cell.widthTwips - removed) } : {}),
    });
    if (r === caret.row) land = covering;
  });
  widths.splice(at, 1);
  carried.splice(at, 1);
  setGrid(t, widths, carried);
  return landed(caret.row, land);
}

/**
 * One cell out of its row, with the columns it covered given to the row's last cell so the row still tiles
 * the grid — what Word calls "delete cells, shift cells left". A row that would be left with no cell keeps
 * every cell it has, which the caller reports as the change not happening.
 */
export function removeCell(row: TableCell[], at: number): boolean {
  if (at < 0 || at >= row.length || row.length === 1) return false;
  const gone = row.splice(at, 1)[0];
  const tail = row[row.length - 1];
  applyCellProps(tail, {
    spanCols: spanOf(tail) + spanOf(gone),
    ...(tail.widthTwips !== undefined && gone.widthTwips !== undefined
      ? { widthTwips: tail.widthTwips + gone.widthTwips }
      : {}),
  });
  return true;
}

/**
 * One cell out of a rectangle of them. The top-left cell keeps the merged body — every paragraph the others
 * held, in reading order — and each row under it keeps a continuation cell, which is how OOXML says a vertical
 * merge occupies a row. A rectangle that cuts through an existing merge is refused rather than guessed at: the
 * cells on its edge belong to a merge that started outside the selection, and pulling them apart would leave a
 * row that does not tile the grid.
 */
function mergeCells(t: Table, from: CellWhere, to: CellWhere): TableEdit {
  const a = cellRect(t, from);
  const b = cellRect(t, to);
  if (!a || !b) return refused("选区不在同一张表里");
  const r0 = Math.min(a.r0, b.r0);
  const r1 = Math.max(a.r1, b.r1);
  const c0 = Math.min(a.c0, b.c0);
  const c1 = Math.max(a.c1, b.c1);
  if (r1 - r0 === 1 && c1 - c0 === 1) return refused("已经是一个单元格了");

  for (let r = r0; r < r1; r++) {
    const row = t.rows[r];
    if (!row) return refused("选区超出了表格");
    const starts = columnStarts(row);
    for (let i = 0; i < row.length; i++) {
      const end = starts[i] + spanOf(row[i]);
      if ((starts[i] < c0 && end > c0) || (starts[i] < c1 && end > c1))
        return refused("选区压到了已有合并的边界");
    }
  }

  const widths = gridWidths(t);
  const merged = widths.slice(c0, c1).reduce((n, w) => n + w, 0);
  const picked = t.rows.slice(r0, r1).map((row) => {
    const starts = columnStarts(row);
    return row.filter((cell, i) => starts[i] >= c0 && starts[i] < c1);
  });
  const anchor = picked[0][0];
  const body = picked.flatMap((cells, r) => (r === 0 ? cells.slice(1) : cells).flatMap(contentOf));

  applyCellProps(anchor, {
    spanCols: c1 - c0,
    ...(merged > 0 ? { widthTwips: merged } : {}),
    spanRows: r1 - r0 > 1 ? 1 : undefined,
  });
  if (body.length) anchor.paragraphs = [...anchor.paragraphs.filter(hasText), ...body];

  picked.forEach((cells, r) => {
    const row = t.rows[r0 + r];
    if (r > 0) {
      applyCellProps(cells[0], {
        spanCols: c1 - c0,
        ...(merged > 0 ? { widthTwips: merged } : {}),
        spanRows: 0,
      });
      cells[0].paragraphs = [blank()];
    }
    const dropped = new Set(cells.slice(1));
    t.rows[r0 + r] = row.filter((cell) => !dropped.has(cell));
  });
  t.columns = gridColumns(t);
  return landed(r0, t.rows[r0].indexOf(anchor));
}

/**
 * A merged cell back into as many cells as it covers. Continuations in the rows below lose their `w:vMerge` and
 * become ordinary cells, and each row regains the columns its cell had swallowed; the width is shared out
 * across them, which is what a grid of that shape was worth before it was merged.
 */
function splitCell(t: Table, where: CellWhere): TableEdit {
  const rect = cellRect(t, where);
  if (!rect) return refused("这里没有单元格");
  const columns = rect.c1 - rect.c0;
  const rows = rect.r1 - rect.r0;
  if (columns === 1 && rows === 1) return refused("这个单元格没有合并");
  const widths = gridWidths(t);
  const each = widths.slice(rect.c0, rect.c1).reduce((n, w) => n + w, 0);
  const share = columns ? Math.floor(each / columns) : 0;

  for (let r = rect.r0; r < rect.r1; r++) {
    const row = t.rows[r];
    const at = columnStarts(row).indexOf(rect.c0);
    if (at < 0) continue;
    const cell = row[at];
    const covered = spanOf(cell);
    applyCellProps(cell, { spanCols: 1, spanRows: undefined, ...(share > 0 ? { widthTwips: share } : {}) });
    const extra = Array.from({ length: covered > 1 ? columns - 1 : 0 }, (): TableCell => makeCell(share, cell.row));
    row.splice(at + 1, 0, ...extra);
  }
  t.columns = gridColumns(t);
  return landed(where.row, where.col);
}

/**
 * 插入单元格, 单元格右移: an empty cell appears to the **left** of the caret cell and that row's cells give way
 * — `wdInsertCellsShiftRight`, and one of the four options behind `wpscore`'s `InsertCells`.
 *
 * Only one row asked for a cell, but the grid is shared: it gains a column, so every other row has to keep
 * tiling it without visibly changing. The cell that owned the host column takes the new one over, gaining a
 * `w:gridSpan` and no width, because the column it now covers was already one of its own. That is also why the
 * seam is the caret cell's *own* first grid column — a seam that cut through a cell would widen one nobody
 * selected.
 *
 * The table keeps the width it had, the same way 插入列 buys a column: the host column is halved, the caret cell
 * hands over half of its width, and the empty cell is given that half.
 */
function insertCellRight(t: Table, where: CellWhere): TableEdit {
  const rect = cellRect(t, where);
  if (!rect) return refused("这里没有单元格");
  const seam = rect.c0;
  const widths = gridWidths(t);
  const hostWidth = widths[seam] ?? 0;
  const added = Math.max(1, Math.floor(hostWidth / 2));
  const carried = [...gridCarried(t)];
  widths[seam] = Math.max(0, hostWidth - added);
  carried.splice(seam, 1, undefined);
  widths.splice(seam, 0, added);
  carried.splice(seam, 0, undefined);

  const row = t.rows[where.row];
  const caret = row[where.col];
  if (caret.widthTwips !== undefined) {
    applyCellProps(caret, { widthTwips: Math.max(0, caret.widthTwips - added) });
  }
  row.splice(where.col, 0, makeCell(added, caret.row));

  t.rows.forEach((r, index) => {
    if (index === where.row) return;
    const covering = cellCovering(r, seam);
    if (covering >= 0) applyCellProps(r[covering], { spanCols: spanOf(r[covering]) + 1 });
  });
  setGrid(t, widths, carried);
  return landed(where.row, where.col);
}

/**
 * 插入单元格, 单元格下移: an empty cell appears **above** the caret cell — `wdInsertCellsShiftDown` — which is
 * the caret cell's own place taken and its contents, and every cell below it in the same grid column, moved one
 * row down.
 *
 * The cells move, their properties do not: each receiving cell keeps the `w:tcPr` it already had (its borders,
 * shading, width), because the command adds a cell to the column, not a row to the table. Below the last row
 * there is nowhere to shift to, so the content that would leave the table is joined to the cell it lands in, in
 * reading order, and nothing is dropped.
 *
 * A column that carries a merge cannot be shifted inside: a horizontally merged cell has no cell starting on
 * this grid column to move into, and a vertically merged band belongs to rows above and below the caret, so
 * moving one link of it detaches the rest. Both refuse and say so.
 */
function insertCellDown(t: Table, where: CellWhere): TableEdit {
  const rect = cellRect(t, where);
  if (!rect) return refused("这里没有单元格");
  const last = t.rows.length - 1;
  if (where.row >= last) return refused("下面没有可让位的单元格");
  const seam = rect.c0;
  const places: number[] = [];
  for (let r = where.row; r <= last; r++) {
    const at = columnStarts(t.rows[r]).indexOf(seam);
    if (at < 0) return refused("这一列上有横向合并的单元格，下移会让它装不进来");
    if (t.rows[r][at].spanRows !== undefined) return refused("这一列上有纵向合并的单元格，下移会把它挪位");
    places.push(at);
  }

  const offset = (r: number): number => places[r - where.row];
  const overflow = t.rows[last][offset(last)].paragraphs;
  for (let r = last; r > where.row; r--) t.rows[r][offset(r)].paragraphs = t.rows[r - 1][offset(r - 1)].paragraphs;
  const landedCell = t.rows[last][offset(last)];
  landedCell.paragraphs = [...landedCell.paragraphs, ...overflow.filter(hasText)];
  t.rows[where.row][offset(where.row)].paragraphs = [blank()];
  return landed(where.row, places[0]);
}

// ------------------------------------------------------------------ entry points

/** 在光标所在行的上方/下方插入一行。 */
export function insertRowHere(t: Table, where: CellWhere, side: "above" | "below"): TableEdit {
  if (!t.rows[where.row]) return refused("这一行不在表格里");
  return insertRow(t, side === "above" ? where.row : where.row + 1);
}

/** 在光标所在格的左/右插入一列；光标留在这一行里被改动的那个格上。 */
export function insertColumnHere(t: Table, where: CellWhere, side: "left" | "right"): TableEdit {
  const rect = cellRect(t, where);
  if (!rect) return refused("这里没有单元格");
  return insertColumn(t, side === "left" ? rect.c0 : rect.c1, where);
}

/** 删除光标所在的整行。 */
export function deleteRowHere(t: Table, where: CellWhere): TableEdit {
  if (!t.rows[where.row]) return refused("这一行不在表格里");
  return deleteRow(t, where.row);
}

/** 删除光标所在格占的第一条网格列。 */
export function deleteColumnHere(t: Table, where: CellWhere): TableEdit {
  const rect = cellRect(t, where);
  if (!rect) return refused("这里没有单元格");
  return deleteColumn(t, rect.c0, where);
}

/** 把两点外接矩形盖住的格合成一个。 */
export function mergeHere(t: Table, from: CellWhere, to: CellWhere = from): TableEdit {
  return mergeCells(t, from, to);
}

/** 拆开光标所在的合并格。 */
export function splitHere(t: Table, where: CellWhere): TableEdit {
  return splitCell(t, where);
}

/**
 * 插入单元格: which way the existing cells give way. `right` puts the new cell to the left of the caret cell and
 * shifts that row; `down` puts it above and shifts that grid column. The caret ends on the new, empty cell —
 * which is the difference from 插入列: a column is added to the whole table and every row takes a cell, while a
 * cell is added to one row or one column and the rest of the table only absorbs the grid it needs.
 */
export function insertCellHere(t: Table, where: CellWhere, shift: CellShift): TableEdit {
  return shift === "right" ? insertCellRight(t, where) : insertCellDown(t, where);
}

// ------------------------------------------------- 粘贴为嵌套表 / 粘贴为追加行

/** What 粘贴为嵌套表 reports: where the new table landed inside the cell, or why the cell would not take it. */
export type NestedOutcome = { ok: true; nested: Table; block: number } | { ok: false; reason: string };

/** What 粘贴为追加行 reports: the rows that joined the table, or why the grid could not hold them. */
export type AppendOutcome = { ok: true; at: CellWhere; rows: TableCell[][] } | { ok: false; reason: string };

/**
 * A copy is new content, so it does not travel with the original's own identity: `w14:paraId` names one
 * paragraph in one document, and the comment links key on it. Duplicating the id would put two paragraphs
 * under one name; leaving a copy with no id is what a *new* paragraph is anyway, and the writer never
 * invents one. Row and cell attributes carry `w:rsid*` save marks, which documents share freely, so those
 * stay exactly as the file wrote them.
 */
const dropCopyIdentity = (blocks: Block[]): void => {
  for (const block of blocks) {
    if (block.kind === "paragraph") delete block.attrsXml;
    else if (block.kind === "table") for (const row of block.rows) for (const cell of row) dropCopyIdentity(cell.paragraphs);
  }
};

/**
 * Whether this table still holds a change nobody has decided.
 *
 * A paste takes the copied table *as it is*, and a pending record is not content — it is a claim about the
 * document, signed with a `w:id` that is unique across the whole part and, for a cell or a table, written
 * inside the carried `w:tcPr`/`w:tblPr` text rather than in a field the model owns. Copying it would put two
 * records under one `w:id`, and a reviewer deciding one would decide the shape of the other. The commands
 * below refuse that case until something can hand a pasted record a fresh id and rewrite it where it lives.
 */
export function carriesUndecidedRevision(t: Table): boolean {
  const pending = (records: { id: number }[] | undefined): boolean => !!records?.length;
  const block = (b: Block): boolean => {
    if (b.kind === "paragraph")
      return !!b.markRev || !!b.revFormat || b.inlines.some((i) => !!i.rev || !!i.revFormat);
    if (b.kind !== "table") return false;
    return tableHas(b);
  };
  const tableHas = (x: Table): boolean =>
    !!x.revFormat ||
    x.rows.some(
      (row) =>
        !!row[0]?.row?.rev ||
        !!row[0]?.row?.revFormat ||
        row.some((cell) => pending(cell.revs) || cell.paragraphs.some(block)),
    );
  return tableHas(t);
}

/**
 * 粘贴为嵌套表 (`wps.KSelection.PasteAsNestedTable`, `EditPasteAsNestedTable` in the command id table): the
 * copied table becomes a table *inside* the cell the insertion point is in, the shape Word documents for —
 * "复制的表格将嵌套在插入点所在单元格中".
 *
 * The copy keeps its own `w:tblGrid`, widths and properties: nested in a cell it is a table of its own and
 * sizes itself there, which is the difference from 追加行. `afterBlock` is the block the insertion point sits
 * on, so the new table lands where the user was rather than at the bottom of the cell.
 *
 * The two shapes a cell may not hold — a table as its last block, and two tables side by side — are closed
 * here with an empty paragraph instead of being left for the writer to paper over, so what the page shows is
 * what a save writes and a reload reads back.
 */
export function pasteAsNestedTableHere(t: Table, where: CellWhere, source: Table, afterBlock?: number): NestedOutcome {
  const cell = t.rows[where.row]?.[where.col];
  if (!cell) return { ok: false, reason: "这里没有单元格" };
  if (!source.rows.some((row) => row.length > 0)) return { ok: false, reason: "复制的表里没有行" };
  if (carriesUndecidedRevision(source))
    return { ok: false, reason: "复制的表里还有待判的修订，粘出来的副本会与原处共用同一个 w:id" };

  const nested = structuredClone(source);
  nested.columns = gridColumns(nested);
  for (const row of nested.rows) for (const cell of row) dropCopyIdentity(cell.paragraphs);
  const at = Math.max(0, Math.min(afterBlock ?? cell.paragraphs.length, cell.paragraphs.length));
  const placed: Block[] = [...cell.paragraphs.slice(0, at), nested, ...cell.paragraphs.slice(at)];
  // 一个格子不能以表收尾，两张表也不能并排 —— 写出器 `cellBlocksXml` 的同一两条规则在这里就补齐，
  // 于是按下命令看到的画面、存出的件与存完重读回来的块序是同一份，而不是差着 writer 补的那一段。
  const blocks: Block[] = [];
  placed.forEach((block, i) => {
    if (block.kind === "table" && i > 0 && placed[i - 1].kind === "table") blocks.push(blank());
    blocks.push(block);
  });
  const index = blocks.indexOf(nested);
  if (index === blocks.length - 1) blocks.push(blank());
  // 就地换掉格子里的块序，而不是把数组换成新的：拿着这个列表引用的调用方（页面、历史）不该被换掉。
  cell.paragraphs.splice(0, cell.paragraphs.length, ...blocks);
  return { ok: true, nested, block: index };
}

/**
 * 粘贴为追加行 (`wps.KSelection.PasteAppendTable`, `EditPasteAppendTable`): the copied table's rows join the
 * table the insertion point is in, below it. Word's own description is the rule this follows — the pasted
 * table "将以和上表等列宽的方式追加在下面", i.e. it is adjusted to the columns of the table it lands on.
 *
 * So the destination's `w:tblGrid` is what sizes the new rows: each copied cell is re-tiled onto the grid
 * columns it now covers, and where a copied row has fewer columns than the destination, the last cell of the
 * row absorbs the rest rather than the row leaving a hole. Nothing on the destination changes — no grid
 * column, no `w:tcPr` of a cell that was already there.
 */
export function pasteAppendTableHere(t: Table, where: CellWhere, source: Table): AppendOutcome {
  if (!t.rows[where.row]) return { ok: false, reason: "这一行不在表格里" };
  if (carriesUndecidedRevision(source))
    return { ok: false, reason: "复制的表里还有待判的修订，粘出来的副本会与原处共用同一个 w:id" };
  const rows = source.rows.filter((row) => row.length > 0);
  if (!rows.length) return { ok: false, reason: "复制的表里没有行" };

  const widths = gridWidths(t);
  const columns = widths.length;
  const tooWide = rows.find((row) => row.reduce((n, cell) => n + spanOf(cell), 0) > columns);
  if (tooWide)
    return {
      ok: false,
      reason: `复制的表一行有 ${tooWide.reduce((n, c) => n + spanOf(c), 0)} 列，这张表只有 ${columns} 列，追加会装不下`,
    };

  const appended = rows.map((row) => {
    // 整行一起克隆：同一行的格子共用同一份 `w:trPr` 记录，分开克隆会把这份共用拆成各格自己的。
    const cells = structuredClone(row);
    cells.forEach((cell) => dropCopyIdentity(cell.paragraphs));
    const spare = columns - cells.reduce((n, cell) => n + spanOf(cell), 0);
    const last = cells[cells.length - 1];
    if (spare > 0) applyCellProps(last, { spanCols: spanOf(last) + spare });
    let at = 0;
    for (const cell of cells) {
      const covered = spanOf(cell);
      // 按目标网格重新给宽；`pct` 是相对正文栏量的，网格列宽对它没有发言权。
      if (cell.widthType !== "pct") applyCellProps(cell, { widthTwips: widths.slice(at, at + covered).reduce((a, b) => a + b, 0) });
      at += covered;
    }
    // 行属性跟着行走，但那是这张表里的另一行了：换一份新的记录，别让两张表的行互相影响。
    const info = cells[0]?.row;
    if (info) {
      const fresh = structuredClone(info);
      cells.forEach((cell) => (cell.row = fresh));
    }
    return cells;
  });

  const first = t.rows.length;
  t.rows.push(...appended);
  // 源表里一段竖合并被从中间截走时，续格在新表里没有上头：与删行、拆表同一处理方式，把它升为起点。
  repairMerges(t);
  t.columns = gridColumns(t);
  return { ok: true, at: { row: first, col: 0 }, rows: appended };
}

/** What 拆分表格 reports: the lower half as its own table, or why the grid would not come apart. */
export type SplitOutcome = { ok: true; tail: Table } | { ok: false; reason: string };

/**
 * 拆分表格: from `atRow` down, the rows become a second table — the command whose name `wpscore` spells
 * `SplitTable`, in the same string table as `InsertCells`, `SelectRow` and `wps.KSelection.PasteAsNestedTable`.
 *
 * The lower half is a table of its own, so it gets the same properties the file wrote for the grid it came
 * from: `w:tblPr`, the `w:tblGrid`, the column widths. Rows and cells carry their own pending records with
 * them, which is why a split is allowed while a row or a cell is undecided. A *table's* record is the
 * exception: one `w:tblPrChange` with one `w:id` cannot describe two tables, so that case refuses rather
 * than signing the same change onto both halves.
 *
 * A vertical merge that was cut in half comes out the other side as a merge of the half it now lives in —
 * `repairMerges` promotes the orphaned continuations, the same way deleting a row does.
 */
export function splitTableAt(t: Table, atRow: number): SplitOutcome {
  if (atRow <= 0) return { ok: false, reason: atRow === 0 ? "第一行上面没有可拆的半张表" : "这一行不在表格里" };
  if (atRow >= t.rows.length) return { ok: false, reason: "最后一行下面没有可拆的半张表" };
  if (t.revFormat) return { ok: false, reason: "这张表上有一条待判的属性改动，拆开会把它同时记到两张表上" };
  const tail: Table = {
    kind: "table",
    columns: t.columns,
    rows: t.rows.splice(atRow),
    styleId: t.styleId,
    widthTwips: t.widthTwips,
    widthType: t.widthType,
    indentTwips: t.indentTwips,
    align: t.align,
    layout: t.layout,
    borders: t.borders,
    columnWidths: t.columnWidths ? [...t.columnWidths] : undefined,
    gridXml: t.gridXml,
    tblPrXml: t.tblPrXml,
  };
  t.columns = gridColumns(t);
  tail.columns = gridColumns(tail);
  repairMerges(t);
  repairMerges(tail);
  return { ok: true, tail };
}

/** What 删除表格 reports: the slot the table left, now an empty paragraph, or why nothing was there. */
export type DeleteTableOutcome = { ok: true; at: number } | { ok: false; reason: string };

/**
 * 删除表格 (`CT_TableTools.kuip:36`: `ksoCmd="DeleteTable"`, `idMso="TableDelete"`, the fourth leg of the `Delete`
 * menu at :32-36): the table at `index` leaves the block list it sits in — the body's, or one cell's — and one
 * empty paragraph takes its place.
 *
 * The paragraph is not decoration. A list with nothing left in it has no line for the caret to sit on, and two
 * tables with nothing between them are read back as one table — the same two rules `cellBlocksXml` spells for a
 * cell and `pasteAsNestedTableHere` keeps when it lays a copy down. Replacing the block instead of splicing it
 * out also leaves every other index where the page had it, so the path a command reports after this one still
 * names the same blocks.
 *
 * A tracked deletion ends here too: 修订 signs each row `w:trPr/w:del` and leaves the table standing until the
 * reviewer accepts it, and `WordDocument.resolveRows` — which has to drop a table whose last row went — puts
 * this same empty paragraph in the slot. What WPS's own 删除表格 leaves behind is 未取证（装机件那一条点不到）,
 * so the shape here is the one the format owes rather than a copy of a screen nobody has seen.
 */
export function deleteTableBlock(blocks: Block[], index: number): DeleteTableOutcome {
  if (blocks[index]?.kind !== "table") return { ok: false, reason: "这里没有表格" };
  blocks.splice(index, 1, blank());
  return { ok: true, at: index };
}
