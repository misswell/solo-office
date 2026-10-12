/**
 * REA/Ghidra etxmlrw FUN_00095ecc confirms per-column dual AND/OR conditions,
 * and FUN_0001ea64 distinguishes blanks/nonblanks in Excel 2003 SpreadsheetML.
 * Modern OOXML encoding/row visibility below are independent implementations;
 * these legacy observations do not establish complete WPS XLSX behavior.
 * Evidence: .scratch/filter-evidence-report.md (2026-10-11).
 */
import { isError, type Value } from "./formula.js";
import { asArray, children } from "./opc.js";
import { parseRange, rangeSpec, rect, type Axis, type Rect } from "./sheet-ops.js";

export type FilterOperator = "equal" | "notEqual" | "lessThan" | "lessThanOrEqual" | "greaterThan" | "greaterThanOrEqual";
export type FilterRule = { kind: "values"; values: string[]; blank: boolean }
  | { kind: "custom"; and: boolean; tests: { operator: FilterOperator; value: string }[] }
  | { kind: "opaque" };
export interface FilterColumn { id: number; rule: FilterRule; sourceId?: number; edited?: boolean }
export interface AutoFilter { ref: string; columns: FilterColumn[]; unsupported?: string }
export const FILTER_OPERATORS: FilterOperator[] = ["equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"];
export const filterFingerprint = (filter: AutoFilter | undefined): string => JSON.stringify(filter ?? null);
const yes = (v: unknown): boolean => v === "1" || v === "true" || v === true;
const scalar = (v: any): string => String(v?.["#text"] ?? v ?? "");

export function filterRange(spec: string): Rect {
  if (!/^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(spec)) throw new Error("筛选区域必须是 A1:B10 形式");
  const box = parseRange(spec);
  if (box.c2 > 16383 || box.r2 > 1048575) throw new Error("筛选区域超出工作表边界");
  return box;
}

/** Parse only supported SpreadsheetML criteria; unknown criteria remain owned by the source XML. */
export function readAutoFilter(node: any): AutoFilter | undefined {
  if (!node) return undefined;
  let ref: string;
  try { ref = rangeSpec(filterRange(scalar(node["@_ref"]))); } catch { return undefined; }
  const box = filterRange(ref), columns: FilterColumn[] = [];
  let unsupported: string | undefined;
  const seen = new Set<number>();
  for (const col of asArray(node.filterColumn)) {
    const id = Number(col["@_colId"]);
    let rule: FilterRule = { kind: "opaque" };
    // Ask the parser's own view of its children, not `Object.keys`: a node with two distinct child tags
    // carries `#order` besides them, and counting that bookkeeping as a criterion made every column that
    // holds an extension next to its filters read as an unsupported one.
    const kinds = children(col).filter((k) => k.name !== "extLst").map((k) => k.name);
    if (Number.isInteger(id) && id >= 0 && id <= box.c2 - box.c1 && !seen.has(id) && kinds.length === 1) {
      if (col.filters && !col.filters["@_calendarType"] && childNamesAre(col.filters, ["filter"])) {
        rule = { kind: "values", values: asArray(col.filters.filter).map((f) => scalar(f["@_val"])), blank: yes(col.filters["@_blank"]) };
      } else if (col.customFilters && childNamesAre(col.customFilters, ["customFilter"])) {
        const tests = asArray(col.customFilters.customFilter).map((f) => ({ operator: scalar(f["@_operator"] || "equal") as FilterOperator, value: scalar(f["@_val"]) }));
        if (tests.length >= 1 && tests.length <= 2 && tests.every((t) => FILTER_OPERATORS.includes(t.operator))) rule = { kind: "custom", and: yes(col.customFilters["@_and"]), tests };
      }
    }
    seen.add(id);
    // "暂未支持" is about a criterion we would misread, not about bytes we do not own. A `filterColumn`
    // holding only `extLst` or attributes like `hiddenButton` states no condition at all: matching everything
    // is exactly what `opaque` means, and the unowned bytes travel back intact, so it must not freeze the
    // other columns of the filter.
    if (rule.kind === "opaque" && kinds.length) unsupported = "此文件含暂未支持的筛选条件，保留原始筛选与行可见性";
    columns.push({ id, sourceId: id, rule });
  }
  return { ref, columns, unsupported };
}

/** Nothing but these child tags (attributes are not children here): an extension we do not own says no criterion. */
const childNamesAre = (node: any, allowed: string[]): boolean => children(node).every((k) => allowed.includes(k.name));

export function filterText(value: Value): string {
  return value === null ? "" : typeof value === "boolean" ? value ? "TRUE" : "FALSE" : String(value);
}
export function filterMatches(value: Value, rule: FilterRule): boolean {
  const text = filterText(value);
  if (rule.kind === "opaque") return true;
  if (rule.kind === "values") return text === "" ? rule.blank : rule.values.some((v) => v.toLowerCase() === text.toLowerCase());
  const results = rule.tests.map((t) => {
    const expected = Number(t.value), numeric = typeof value === "number" && t.value.trim() !== "" && Number.isFinite(expected);
    if (t.operator === "equal" || t.operator === "notEqual") {
      const equal = numeric ? value === expected : wildcard(t.value).test(text);
      return t.operator === "equal" ? equal : !equal;
    }
    if (text === "" || isError(value)) return false;
    const cmp = numeric ? (value as number) - expected : text.toLowerCase() < t.value.toLowerCase() ? -1 : text.toLowerCase() > t.value.toLowerCase() ? 1 : 0;
    return t.operator === "lessThan" ? cmp < 0 : t.operator === "lessThanOrEqual" ? cmp <= 0 : t.operator === "greaterThan" ? cmp > 0 : cmp >= 0;
  });
  return rule.and ? results.every(Boolean) : results.some(Boolean);
}
function wildcard(pattern: string): RegExp {
  let expression = "";
  const escape = (c: string): string => c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === "~" && i + 1 < pattern.length && ["*", "?", "~"].includes(pattern[i + 1])) expression += escape(pattern[++i]);
    else expression += c === "*" ? ".*" : c === "?" ? "." : escape(c);
  }
  return new RegExp(`^${expression}$`, "isu");
}

/** Expand the active cell's current region, stopping at wholly blank row/column separators. */
export function filterDataRegion(at: { c: number; r: number }, bounds: Rect, occupied: (c: number, r: number) => boolean): Rect {
  if (!occupied(at.c, at.r)) throw new Error("请选择数据区域中的单元格");
  let box = rect(at.c, at.r, at.c, at.r), changed = true;
  while (changed) {
    const before = JSON.stringify(box);
    const colHas = (c: number): boolean => { for (let r = box.r1; r <= box.r2; r++) if (occupied(c, r)) return true; return false; };
    const rowHas = (r: number): boolean => { for (let c = box.c1; c <= box.c2; c++) if (occupied(c, r)) return true; return false; };
    while (box.c1 > 0 && colHas(box.c1 - 1)) box.c1--;
    while (box.c2 < bounds.c2 && colHas(box.c2 + 1)) box.c2++;
    while (box.r1 > 0 && rowHas(box.r1 - 1)) box.r1--;
    while (box.r2 < bounds.r2 && rowHas(box.r2 + 1)) box.r2++;
    changed = before !== JSON.stringify(box);
  }
  return box;
}

export function validateFilterRule(rule: FilterRule): void {
  if (rule.kind === "opaque") throw new Error("不支持修改此筛选类型");
  if (rule.kind === "values") {
    if (!Array.isArray(rule.values) || rule.values.some((v) => typeof v !== "string") || typeof rule.blank !== "boolean") throw new Error("无效筛选值");
  } else if (!rule.tests.length || rule.tests.length > 2 || !rule.tests.every((t) => FILTER_OPERATORS.includes(t.operator) && typeof t.value === "string")) throw new Error("筛选需要一至两个有效条件");
}

/** Structural changes preserve column identity; deleting the header removes the filter. */
export function shiftAutoFilter(filter: AutoFilter, axis: Axis, at: number, count: number): AutoFilter | undefined {
  const box = filterRange(filter.ref), drop = Math.max(0, count), delta = -count;
  if (axis === "row" && drop && at <= box.r1 && box.r1 < at + drop) return undefined;
  const start = axis === "row" ? box.r1 : box.c1, end = axis === "row" ? box.r2 : box.c2;
  const remaining = drop ? [start, end].map((n, i) => n < at ? n : n >= at + drop ? n + delta : i ? at - 1 : at) : [start >= at ? start + delta : start, end >= at ? end + delta : end];
  if (remaining[0] > remaining[1]) return undefined;
  const nextBox = axis === "row" ? rect(box.c1, remaining[0], box.c2, remaining[1]) : rect(remaining[0], box.r1, remaining[1], box.r2);
  const columns = filter.columns.flatMap((col) => {
    if (axis === "row") return [structuredClone(col)];
    const absolute = box.c1 + col.id;
    if (drop && absolute >= at && absolute < at + drop) return [];
    const id = (absolute >= at ? absolute + delta : absolute) - nextBox.c1;
    return [{ ...structuredClone(col), id }];
  });
  return { ...filter, ref: rangeSpec(nextBox), columns };
}
