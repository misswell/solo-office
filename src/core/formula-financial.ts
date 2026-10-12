import type { Node, Value } from "./formula.js";
import { Evaluator, isError, toNumber } from "./formula.js";

/**
 * Financial extension library: the time-value-of-money set (PMT/IPMT/PPMT/
 * PV/FV/NPER/RATE/NPV/IRR/MIRR/CUMIPMT/CUMPRINC) and depreciation (SLN/SYD/DDB).
 * WPS's side is evidenced by the KF_Pmt / KF_IPmt / KF_Cumipmt / KF_Irr …
 * implementation classes in etcore (see .scratch/rea-notes.md §8).
 *
 * Sign convention matches Excel: money the user pays out is negative.
 * Convergence: RATE and IRR solve by bracketed bisection over (-0.9999, 10);
 * a missing bracket (no sign change, no root) surfaces as #NUM!.
 */

export type Fn = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string) => Value;

const err = (code: string): never => {
  throw new Error(code);
};

const numArg = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string, i: number): number => {
  const v = node.args[i] ? ev.sub(node.args[i], sheet) : null;
  const n = toNumber(v);
  if (typeof n !== "number") throw new Error(n);
  return n;
};

const listArg = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string, i: number): Value[] =>
  node.args[i] ? ev.list(node.args[i], sheet) : [];

/** Annuity payment. `paysStart` is Excel's type flag (1 = payments at period start). */
export function pmtOf(rate: number, nper: number, pv: number, fv: number, paysStart: number): number {
  if (rate === 0) return -(pv + fv) / nper;
  const k = (1 + rate) ** nper;
  return -(pv * k + fv) * rate / ((k - 1) * (1 + rate * paysStart));
}

export function fvOf(rate: number, nper: number, pmt: number, pv: number, paysStart: number): number {
  if (rate === 0) return -(pv + pmt * nper);
  const k = (1 + rate) ** nper;
  return -(pv * k + pmt * (1 + rate * paysStart) * (k - 1) / rate);
}

export function pvOf(rate: number, nper: number, pmt: number, fv: number, paysStart: number): number {
  if (rate === 0) return -(fv + pmt * nper);
  const k = (1 + rate) ** nper;
  return -(fv + pmt * (1 + rate * paysStart) * (k - 1) / rate) / k;
}

/** Interest portion of the payment for one period. */
export function ipmtOf(rate: number, per: number, nper: number, pv: number, fv: number, paysStart: number): number {
  if (per < 1 || per > nper) err("#NUM!");
  if (per === 1) return paysStart ? 0 : -pv * rate;
  const prior = fvOf(rate, per - (paysStart ? 2 : 1), pmtOf(rate, nper, pv, fv, paysStart), pv, paysStart);
  const i = prior * rate;
  return paysStart ? i / (1 + rate) : i;
}

/** The annuity equation whose root is RATE. */
const rateEquation = (r: number, nper: number, pmt: number, pv: number, fv: number, paysStart: number): number => {
  if (Math.abs(r) < 1e-12) return pv + pmt * nper + fv;
  const k = (1 + r) ** nper;
  return pv * k + pmt * (1 + r * paysStart) * (k - 1) / r + fv;
};

/** Bracketed bisection; a root outside the bracket surfaces as #NUM!. */
function bisect(f: (x: number) => number, lo: number, hi: number): number {
  let flo = f(lo);
  const fhi = f(hi);
  if (flo * fhi > 0) err("#NUM!");
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2;
    const fm = f(mid);
    if (fm === 0) return mid;
    if (flo * fm < 0) hi = mid;
    else {
      lo = mid;
      flo = fm;
    }
  }
  return (lo + hi) / 2;
}

const NPV = (rate: number, vals: number[]): number => vals.reduce((a, v, i) => a + v / (1 + rate) ** (i + 1), 0);

export const FUNCTIONS: Record<string, Fn> = {
  PMT: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const nper = numArg(ev, n, s, 1);
    const pv = numArg(ev, n, s, 2);
    const fv = n.args[3] ? numArg(ev, n, s, 3) : 0;
    const typ = n.args[4] ? numArg(ev, n, s, 4) : 0;
    return pmtOf(rate, nper, pv, fv, typ);
  },
  FV: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const nper = numArg(ev, n, s, 1);
    const pmt = numArg(ev, n, s, 2);
    const pv = n.args[3] ? numArg(ev, n, s, 3) : 0;
    const typ = n.args[4] ? numArg(ev, n, s, 4) : 0;
    return fvOf(rate, nper, pmt, pv, typ);
  },
  PV: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const nper = numArg(ev, n, s, 1);
    const pmt = numArg(ev, n, s, 2);
    const fv = n.args[3] ? numArg(ev, n, s, 3) : 0;
    const typ = n.args[4] ? numArg(ev, n, s, 4) : 0;
    return pvOf(rate, nper, pmt, fv, typ);
  },
  NPER: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const pmt = numArg(ev, n, s, 1);
    const pv = numArg(ev, n, s, 2);
    const fv = n.args[3] ? numArg(ev, n, s, 3) : 0;
    const typ = n.args[4] ? numArg(ev, n, s, 4) : 0;
    if (rate === 0) {
      if (pmt === 0) err("#DIV/0!");
      return -(pv + fv) / pmt;
    }
    const adj = pmt * (1 + rate * typ);
    const ratio = (adj - fv * rate) / (adj + pv * rate);
    if (ratio <= 0) err("#NUM!");
    return Math.log(ratio) / Math.log(1 + rate);
  },
  RATE: (ev, n, s) => {
    const nper = numArg(ev, n, s, 0);
    const pmt = numArg(ev, n, s, 1);
    const pv = numArg(ev, n, s, 2);
    const fv = n.args[3] ? numArg(ev, n, s, 3) : 0;
    const typ = n.args[4] ? numArg(ev, n, s, 4) : 0;
    const guess = n.args[5] ? numArg(ev, n, s, 5) : 0.1;
    void guess; // bisection is guess-independent; the argument is accepted for compatibility
    return bisect((r) => rateEquation(r, nper, pmt, pv, fv, typ), -0.999999, 10);
  },
  NPV: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const vals: number[] = [];
    for (let i = 1; i < n.args.length; i++) {
      for (const v of ev.list(n.args[i], s)) {
        if (v === null || typeof v === "boolean") continue;
        if (isError(v)) err(v);
        if (typeof v === "number") vals.push(v);
      }
    }
    return NPV(rate, vals);
  },
  IRR: (ev, n, s) => {
    const vals: number[] = [];
    for (const v of listArg(ev, n, s, 0)) {
      if (isError(v)) err(v); // errors in the cashflow propagate
      if (v === null || typeof v === "boolean" || typeof v === "string") continue;
      vals.push(Number(v));
    }
    if (!vals.some((v) => v > 0) || !vals.some((v) => v < 0)) err("#NUM!");
    const f = (r: number): number => vals.reduce((a, v, i) => a + v / (1 + r) ** i, 0);
    return bisect(f, -0.9999, 10);
  },
  MIRR: (ev, n, s) => {
    const vals: number[] = [];
    for (const v of listArg(ev, n, s, 0)) {
      if (isError(v)) err(v);
      if (typeof v !== "number") continue;
      vals.push(v);
    }
    const finance = numArg(ev, n, s, 1);
    const reinvest = numArg(ev, n, s, 2);
    const count = vals.length;
    if (count < 2) err("#DIV/0!");
    const pos = vals.reduce((a, v, i) => a + (v > 0 ? v * (1 + reinvest) ** (count - 1 - i) : 0), 0);
    const neg = vals.reduce((a, v, i) => a + (v < 0 ? v / (1 + finance) ** i : 0), 0);
    if (pos === 0 || neg === 0) err("#DIV/0!");
    return (pos / -neg) ** (1 / (count - 1)) - 1;
  },
  IPMT: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const per = numArg(ev, n, s, 1);
    const nper = numArg(ev, n, s, 2);
    const pv = numArg(ev, n, s, 3);
    const fv = n.args[4] ? numArg(ev, n, s, 4) : 0;
    const typ = n.args[5] ? numArg(ev, n, s, 5) : 0;
    return ipmtOf(rate, per, nper, pv, fv, typ);
  },
  PPMT: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const per = numArg(ev, n, s, 1);
    const nper = numArg(ev, n, s, 2);
    const pv = numArg(ev, n, s, 3);
    const fv = n.args[4] ? numArg(ev, n, s, 4) : 0;
    const typ = n.args[5] ? numArg(ev, n, s, 5) : 0;
    return pmtOf(rate, nper, pv, fv, typ) - ipmtOf(rate, per, nper, pv, fv, typ);
  },
  CUMIPMT: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const nper = numArg(ev, n, s, 1);
    const pv = numArg(ev, n, s, 2);
    const start = numArg(ev, n, s, 3);
    const end = numArg(ev, n, s, 4);
    const typ = n.args[5] ? numArg(ev, n, s, 5) : 0;
    if (start < 1 || end < start || rate <= 0 || nper <= 0 || pv <= 0 || (typ !== 0 && typ !== 1)) err("#NUM!");
    let total = 0;
    for (let k = Math.ceil(start); k <= Math.floor(end); k++) total += ipmtOf(rate, k, nper, pv, 0, typ);
    return total;
  },
  CUMPRINC: (ev, n, s) => {
    const rate = numArg(ev, n, s, 0);
    const nper = numArg(ev, n, s, 1);
    const pv = numArg(ev, n, s, 2);
    const start = numArg(ev, n, s, 3);
    const end = numArg(ev, n, s, 4);
    const typ = n.args[5] ? numArg(ev, n, s, 5) : 0;
    if (start < 1 || end < start || rate <= 0 || nper <= 0 || pv <= 0 || (typ !== 0 && typ !== 1)) err("#NUM!");
    let total = 0;
    for (let k = Math.ceil(start); k <= Math.floor(end); k++) {
      total += pmtOf(rate, nper, pv, 0, typ) - ipmtOf(rate, k, nper, pv, 0, typ);
    }
    return total;
  },
  SLN: (ev, n, s) => {
    const cost = numArg(ev, n, s, 0);
    const salvage = numArg(ev, n, s, 1);
    const life = numArg(ev, n, s, 2);
    if (life === 0) err("#DIV/0!");
    return (cost - salvage) / life;
  },
  SYD: (ev, n, s) => {
    const cost = numArg(ev, n, s, 0);
    const salvage = numArg(ev, n, s, 1);
    const life = numArg(ev, n, s, 2);
    const per = numArg(ev, n, s, 3);
    if (life <= 0 || per < 1 || per > life) err("#NUM!");
    return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1));
  },
  DDB: (ev, n, s) => {
    const cost = numArg(ev, n, s, 0);
    const salvage = numArg(ev, n, s, 1);
    const life = numArg(ev, n, s, 2);
    const period = numArg(ev, n, s, 3);
    const factor = n.args[4] ? numArg(ev, n, s, 4) : 2;
    if (cost < 0 || salvage < 0 || life <= 0 || period < 1 || factor <= 0) err("#NUM!");
    const rate = factor / life;
    if (rate > 1) err("#NUM!");
    let accum = 0;
    let dep = 0;
    for (let k = 1; k <= Math.floor(period); k++) {
      dep = Math.min((cost - accum) * rate, Math.max(0, cost - salvage - accum));
      accum += dep;
    }
    return dep;
  },
};
