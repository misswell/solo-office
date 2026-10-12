/**
 * Pure helpers for grid editing: reference geometry, structural formula
 * rewriting, and clipboard text. No workbook state lives here.
 */

import { colToIndex, indexToCol, type Value } from "./formula.js";

export interface Rect {
  c1: number;
  r1: number;
  c2: number;
  r2: number;
}

export const rect = (c1: number, r1: number, c2: number, r2: number): Rect => ({
  c1: Math.min(c1, c2),
  r1: Math.min(r1, r2),
  c2: Math.max(c1, c2),
  r2: Math.max(r1, r2),
});

export const refOf = (c: number, r: number): string => `${indexToCol(c)}${r + 1}`;

export const parseRef = (ref: string): { c: number; r: number } => {
  const m = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/.exec(ref.trim());
  if (!m) throw new Error(`无效的单元格引用: ${ref}`);
  return { c: colToIndex(m[1]), r: Number(m[2]) - 1 };
};

/** "A1:B3" — or a lone "B2" — to an ordered rectangle. */
export const parseRange = (spec: string): Rect => {
  const [from, to] = spec.split(":");
  const a = parseRef(from);
  if (!to) return rect(a.c, a.r, a.c, a.r);
  const b = parseRef(to);
  return rect(a.c, a.r, b.c, b.r);
};

export const rangeSpec = (box: Rect): string =>
  box.c1 === box.c2 && box.r1 === box.r2 ? refOf(box.c1, box.r1) : `${refOf(box.c1, box.r1)}:${refOf(box.c2, box.r2)}`;

/** Row-major walk over a rectangle. */
export function* eachCell(box: Rect): Generator<{ c: number; r: number }> {
  for (let r = box.r1; r <= box.r2; r++) for (let c = box.c1; c <= box.c2; c++) yield { c, r };
}

export const inRect = (box: Rect, c: number, r: number): boolean => c >= box.c1 && c <= box.c2 && r >= box.r1 && r <= box.r2;

export const expand = (box: Rect, dC1: number, dR1: number, dC2: number, dR2: number): Rect => ({
  c1: box.c1 + dC1,
  r1: box.r1 + dR1,
  c2: box.c2 + dC2,
  r2: box.r2 + dR2,
});

// ------------------------------------------------------- structural shifting

export type Axis = "row" | "col";

/**
 * A reference or range with an optional sheet qualifier. Letters and digits are
 * captured separately so `$` marks can be carried through untouched. Sheet names
 * are matched as unicode words so `汇总!B2` qualifies like `'数据 2'!B2` does.
 */
const REF =
  /(?:(?:'([^']+)'|([\p{L}\p{N}_.][\p{L}\p{N}_.]*))!)?(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})(?::(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6}))?(?![0-9A-Za-z_(])/gu;

const outsideStrings = (formula: string, fn: (chunk: string) => string): string =>
  formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) => (i % 2 === 1 ? part : fn(part)))
    .join("");

/**
 * Rewrite references for a structural edit. `count > 0` deletes the `count`
 * lines starting at `at`; `count < 0` inserts `-count` lines there.
 *
 * A range that straddles a deleted band shrinks to its surviving part, a range
 * entirely inside it becomes `#REF!`, and a lone reference inside the band does
 * too. `$` anchors are preserved but do not suppress the shift: they pin a
 * reference against *copying*, while an inserted row still moves the cell that
 * every kind of reference points at.
 *
 * `scope.edited` names the sheet whose lines moved and `scope.from` the sheet
 * the formula lives on (the same sheet unless stated): an unqualified `B2` means
 * "B2 over here", so editing 汇总 rewrites 汇总's own formulas and any `汇总!B2`
 * elsewhere, and touches nothing that points at a third sheet.
 */
export function shiftRefs(
  formula: string,
  axis: Axis,
  at: number,
  count: number,
  scope?: { edited: string; from?: string },
): string {
  const drop = count > 0 ? count : 0;
  const delta = -count;
  const edited = scope?.edited.toLowerCase();
  const local = scope?.from?.toLowerCase() ?? edited;
  return outsideStrings(formula, (chunk) =>
    chunk.replace(
      REF,
      (
        match,
        quoted: string | undefined,
        bare: string | undefined,
        cMark1: string,
        col1: string,
        rMark1: string,
        row1: string,
        cMark2: string | undefined,
        col2: string | undefined,
        rMark2: string | undefined,
        row2: string | undefined,
      ) => {
        const qualifier = quoted ?? bare;
        if (scope) {
          const owner = qualifier?.toLowerCase() ?? local;
          if (owner !== edited) return match;
        }
        const prefix = qualifier === undefined ? "" : quoted !== undefined ? `'${quoted}'!` : `${bare}!`;
        const isRange = col2 !== undefined && row2 !== undefined;
        const c1 = colToIndex(col1);
        const r1 = Number(row1) - 1;
        const c2 = isRange ? colToIndex(col2 as string) : c1;
        const r2 = isRange ? Number(row2 as string) - 1 : r1;
        const line = (v: number, which: "start" | "end"): number | undefined => {
          if (drop) {
            if (v >= at && v < at + drop) return which === "start" ? at : at - 1;
            return v >= at ? v + delta : v;
          }
          return v >= at ? v + delta : v;
        };
        const inside = (v: number): boolean => drop > 0 && v >= at && v < at + drop;
        const target = axis === "row" ? [r1, r2] : [c1, c2];
        if (inside(target[0]) && inside(target[1])) return `${prefix}#REF!`;
        const moved = line(target[0], "start");
        const movedEnd = isRange ? line(target[1], "end") : undefined;
        if (moved === undefined || moved < 0) return `${prefix}#REF!`;
        if (!isRange) {
          const col = axis === "col" ? moved : c1;
          const row = axis === "row" ? moved : r1;
          return `${prefix}${cMark1}${indexToCol(col)}${rMark1}${row + 1}`;
        }
        if (movedEnd === undefined || movedEnd < moved) return `${prefix}#REF!`;
        const endCol = axis === "col" ? movedEnd : c2;
        const endRow = axis === "row" ? movedEnd : r2;
        const startCol = axis === "col" ? moved : c1;
        const startRow = axis === "row" ? moved : r1;
        return `${prefix}${cMark1}${indexToCol(startCol)}${rMark1}${startRow + 1}:${cMark2}${indexToCol(endCol)}${rMark2}${endRow + 1}`;
      },
    ),
  );
}


// ---------------------------------------------------------------- clipboard

/**
 * Clipboard/paste text: tab between columns, newline between rows. Quoted
 * fields (as produced by other spreadsheets) are honoured.
 */
export function parseGrid(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else quoted = false;
      } else field += ch;
      continue;
    }
    if (ch === '"' && field === "") {
      quoted = true;
      continue;
    }
    if (ch === "\t") {
      row.push(field);
      field = "";
      continue;
    }
    if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      continue;
    }
    field += ch;
  }
  row.push(field);
  rows.push(row);
  while (rows.length && rows[rows.length - 1].every((f) => f === "")) rows.pop();
  if (!rows.length) return [];
  const width = Math.max(...rows.map((r) => r.length));
  for (const r of rows) while (r.length < width) r.push("");
  if (rows[rows.length - 1].every((f) => f === "")) rows.pop();
  return rows.length ? rows : [];
}

/** Excel's 1900 epoch with the historical leap-year bug, as numfmt reads it back. */
const EPOCH_1900 = Date.UTC(1899, 11, 30);

export const looksNumeric = (s: string): boolean => /^-?(\d+(\.\d+)?|\.\d+)([eE][-+]?\d+)?$/.test(s.trim());

const asDateSerial = (s: string): number | undefined => {
  const m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(s);
  if (!m) return undefined;
  const d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  if (Number.isNaN(d.getTime())) return undefined;
  return Math.round((d.getTime() - EPOCH_1900) / 86400000);
};

/** Typed input to stored value: formulas keep their text, literals are inferred. */
export function parseInput(input: string): { constant: Value; formula?: string } {
  if (input.startsWith("=")) return { constant: null, formula: input.slice(1) };
  if (input === "") return { constant: null };
  if (looksNumeric(input)) return { constant: Number(input) };
  const date = asDateSerial(input.trim());
  if (date !== undefined) return { constant: date };
  const upper = input.toUpperCase();
  if (upper === "TRUE") return { constant: true };
  if (upper === "FALSE") return { constant: false };
  return { constant: input };
}
