import type { Node, Value } from "./formula.js";
import { Evaluator, isError, matchCriteria, toNumber } from "./formula.js";

/**
 * Statistical extension library: population spread, distribution shapes,
 * linear-regression measures and the paired-criteria plurals. Function names
 * and semantics follow Excel/WPS; the WPS side is evidenced by the KF_* classes
 * in etcore (see .scratch/rea-notes.md §8: KF_CorRel, KF_GeoMean, KF_Countifs, …).
 *
 * Range semantics: blanks and text are ignored; booleans in ranges count as
 * 1/0; an error anywhere in a touched range propagates.
 */

export type Fn = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string) => Value;

const err = (code: string): never => {
  throw new Error(code);
};

const collect = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string, from = 0): Value[] => {
  const out: Value[] = [];
  for (let i = from; i < node.args.length; i++) out.push(...ev.list(node.args[i], sheet));
  return out;
};

const numeric = (vals: Value[]): number[] => {
  const out: number[] = [];
  for (const v of vals) {
    if (typeof v === "number") {
      out.push(v);
      continue;
    }
    if (typeof v === "boolean") {
      out.push(v ? 1 : 0);
      continue;
    }
    // Error values are strings, so this check must precede the string branch.
    if (isError(v)) err(v); // an error anywhere in a touched range propagates
    if (typeof v === "string") {
      if (v === "") continue;
      const n = Number(v);
      if (Number.isFinite(n)) out.push(n);
      continue;
    }
  }
  return out;
};

const numArg = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string, i: number): number => {
  const v = node.args[i] ? ev.sub(node.args[i], sheet) : null;
  const n = toNumber(v);
  if (typeof n !== "number") throw new Error(n);
  return n;
};

const mean = (xs: number[]): number => xs.reduce((a, b) => a + b, 0) / xs.length;

const sumSqDev = (xs: number[]): number => {
  const m = mean(xs);
  return xs.reduce((a, b) => a + (b - m) ** 2, 0);
};

/** Regression/correlation pairs: rows where both sides are real numbers. */
function paired(ev: Evaluator, node: Node & { t: "fn" }, sheet: string): { xs: number[]; ys: number[] } {
  const ys = ev.list(node.args[0], sheet);
  const xs = ev.list(node.args[1], sheet);
  // An error anywhere in either range propagates, like Excel's CORREL family.
  for (const v of [...ys, ...xs]) if (isError(v)) err(v);
  const out: { xs: number[]; ys: number[] } = { xs: [], ys: [] };
  const n = Math.min(xs.length, ys.length);
  for (let i = 0; i < n; i++) {
    const a = xs[i];
    const b = ys[i];
    if (typeof a !== "number" || typeof b !== "number") continue;
    out.xs.push(a);
    out.ys.push(b);
  }
  return out;
}

function pearson(xs: number[], ys: number[]): number {
  const mx = mean(xs);
  const my = mean(ys);
  const den = Math.sqrt(sumSqDev(xs) * sumSqDev(ys));
  if (!den) err("#DIV/0!");
  return xs.reduce((a, x, i) => a + (x - mx) * (ys[i] - my), 0) / den;
}

const pairedOr = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string, min: number): { xs: number[]; ys: number[] } => {
  const { xs, ys } = paired(ev, node, sheet);
  if (xs.length !== ys.length) err("#N/A");
  if (xs.length < min) err(min > 1 ? "#DIV/0!" : "#N/A");
  return { xs, ys };
};

/** Shared walker for the *_IFS family: sum/avg/max/min ranges plus criteria pairs. */
function ifsWalk(ev: Evaluator, node: Node & { t: "fn" }, sheet: string): { target: Value[]; picked: number[] } {
  if (node.args.length < 3 || (node.args.length - 1) % 2) err("#VALUE!");
  const target = ev.list(node.args[0], sheet);
  const ranges: Value[][] = [];
  const crits: Value[] = [];
  for (let i = 1; i < node.args.length; i += 2) {
    ranges.push(ev.list(node.args[i], sheet));
    crits.push(ev.sub(node.args[i + 1], sheet));
  }
  const size = target.length;
  if (ranges.some((r) => r.length !== size)) err("#VALUE!");
  const picked: number[] = [];
  for (let row = 0; row < size; row++) {
    if (!ranges.every((r, k) => matchCriteria(r[row], crits[k]))) continue;
    const v = target[row];
    if (isError(v)) err(v); // matched cells in the target range propagate errors
    if (typeof v === "number") picked.push(v);
  }
  return { target, picked };
}

function rankOf(ev: Evaluator, n: Node & { t: "fn" }, sheet: string, average: boolean): Value {
  const x = numArg(ev, n, sheet, 0);
  const v = numeric(ev.list(n.args[1], sheet));
  if (!v.some((b) => b === x)) err("#N/A");
  const asc = n.args[2] ? numArg(ev, n, sheet, 2) !== 0 : false;
  const below = asc ? v.filter((b) => b < x).length : v.filter((b) => b > x).length;
  const same = v.filter((b) => b === x).length;
  return average ? below + (same + 1) / 2 : below + 1;
}

const percentileInc = (v: number[], k: number): number => {
  const s = [...v].sort((a, b) => a - b);
  const idx = k * (s.length - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  return s[lo] + (idx - lo) * (s[hi] - s[lo]);
};

const QUARTILE = (ev: Evaluator, n: Node & { t: "fn" }, s: string): Value => {
  const v = numeric(ev.list(n.args[0], s));
  const q = numArg(ev, n, s, 1);
  if (!v.length || q < 0 || q > 4 || !Number.isInteger(q)) err("#NUM!");
  return percentileInc(v, q / 4);
};
const STDEVP = (ev: Evaluator, n: Node & { t: "fn" }, s: string): Value => {
  const v = numeric(collect(ev, n, s));
  if (!v.length) err("#DIV/0!");
  return Math.sqrt(sumSqDev(v) / v.length);
};
const VARP = (ev: Evaluator, n: Node & { t: "fn" }, s: string): Value => {
  const v = numeric(collect(ev, n, s));
  if (!v.length) err("#DIV/0!");
  return sumSqDev(v) / v.length;
};
const MODE = (ev: Evaluator, n: Node & { t: "fn" }, s: string): Value => {
  const v = numeric(collect(ev, n, s));
  const counts = new Map<number, number>();
  for (const x of v) counts.set(x, (counts.get(x) ?? 0) + 1);
  let best = 1;
  for (const c of counts.values()) best = Math.max(best, c);
  if (best < 2) err("#N/A");
  let out = Infinity;
  for (const [x, c] of counts) if (c === best && x < out) out = x;
  return out;
};
const PERCENTILE = (ev: Evaluator, n: Node & { t: "fn" }, s: string): Value => {
  const v = numeric(ev.list(n.args[0], s));
  const k = numArg(ev, n, s, 1);
  if (!v.length || k < 0 || k > 1) err("#NUM!");
  return percentileInc(v, k);
};
const CORREL = (ev: Evaluator, n: Node & { t: "fn" }, s: string): Value => {
  const { xs, ys } = pairedOr(ev, n, s, 1);
  return pearson(xs, ys);
};
const COVARP = (ev: Evaluator, n: Node & { t: "fn" }, s: string): Value => {
  const { xs, ys } = pairedOr(ev, n, s, 1);
  const mx = mean(xs);
  const my = mean(ys);
  return xs.reduce((a, x, i) => a + (x - mx) * (ys[i] - my), 0) / xs.length;
};
const SLOPE = (ev: Evaluator, n: Node & { t: "fn" }, s: string): Value => {
  const { xs, ys } = pairedOr(ev, n, s, 2);
  const den = sumSqDev(xs);
  if (!den) err("#DIV/0!");
  const mx = mean(xs);
  const my = mean(ys);
  return xs.reduce((a, x, i) => a + (x - mx) * (ys[i] - my), 0) / den;
};

export const FUNCTIONS: Record<string, Fn> = {
  STDEVP,
  "STDEV.P": STDEVP,
  VARP,
  "VAR.P": VARP,
  MODE,
  "MODE.SNGL": MODE,
  GEOMEAN: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    if (!v.length || v.some((x) => x <= 0)) err("#NUM!");
    return Math.exp(v.reduce((a, b) => a + Math.log(b), 0) / v.length);
  },
  HARMEAN: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    if (!v.length || v.some((x) => x <= 0)) err("#NUM!");
    return v.length / v.reduce((a, b) => a + 1 / b, 0);
  },
  PERCENTILE,
  "PERCENTILE.INC": PERCENTILE,
  QUARTILE,
  "QUARTILE.INC": QUARTILE,
  DEVSQ: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    if (!v.length) err("#NUM!");
    return sumSqDev(v);
  },
  AVEDEV: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    if (!v.length) err("#NUM!");
    const m = mean(v);
    return v.reduce((a, b) => a + Math.abs(b - m), 0) / v.length;
  },
  AVERAGEA: (ev, n, s) => {
    const vals = collect(ev, n, s).filter((v) => v !== null && v !== undefined && !isError(v));
    if (!vals.length) err("#DIV/0!");
    const xs = vals.map((v) => (typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : 0));
    return mean(xs);
  },
  CORREL,
  PEARSON: CORREL,
  COVAR: COVARP,
  "COVARIANCE.P": COVARP,
  "COVARIANCE.S": (ev, n, s) => {
    const { xs, ys } = pairedOr(ev, n, s, 2);
    const mx = mean(xs);
    const my = mean(ys);
    return xs.reduce((a, x, i) => a + (x - mx) * (ys[i] - my), 0) / (xs.length - 1);
  },
  SLOPE,
  INTERCEPT: (ev, n, s) => {
    const { xs, ys } = pairedOr(ev, n, s, 2);
    const den = sumSqDev(xs);
    if (!den) err("#DIV/0!");
    const mx = mean(xs);
    const my = mean(ys);
    const slope = xs.reduce((a, x, i) => a + (x - mx) * (ys[i] - my), 0) / den;
    return my - slope * mx;
  },
  RSQ: (ev, n, s) => {
    const { xs, ys } = pairedOr(ev, n, s, 2);
    return pearson(xs, ys) ** 2;
  },
  COUNTIFS: (ev, n, s) => {
    if (n.args.length < 2 || n.args.length % 2) err("#VALUE!");
    const ranges: Value[][] = [];
    const crits: Value[] = [];
    for (let i = 0; i < n.args.length; i += 2) {
      ranges.push(ev.list(n.args[i], s));
      crits.push(ev.sub(n.args[i + 1], s));
    }
    const size = ranges[0].length;
    if (ranges.some((r) => r.length !== size)) err("#VALUE!");
    let count = 0;
    for (let row = 0; row < size; row++) {
      if (ranges.every((r, k) => matchCriteria(r[row], crits[k]))) count++;
    }
    return count;
  },
  SUMIFS: (ev, n, s) => {
    const { picked } = ifsWalk(ev, n, s);
    return picked.reduce((a, b) => a + b, 0);
  },
  AVERAGEIFS: (ev, n, s) => {
    const { picked } = ifsWalk(ev, n, s);
    if (!picked.length) err("#DIV/0!");
    return mean(picked);
  },
  MAXIFS: (ev, n, s) => {
    const { picked } = ifsWalk(ev, n, s);
    return picked.length ? Math.max(...picked) : 0;
  },
  MINIFS: (ev, n, s) => {
    const { picked } = ifsWalk(ev, n, s);
    return picked.length ? Math.min(...picked) : 0;
  },
  "RANK.EQ": (ev, n, s) => rankOf(ev, n, s, false),
  "RANK.AVG": (ev, n, s) => rankOf(ev, n, s, true),
};
