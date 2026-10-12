import type { Node, Value } from "./formula.js";
import { Evaluator, isError, toNumber } from "./formula.js";

/**
 * Math and information extension library: EXP/LN/LOG/LOG10, CEILING/FLOOR/
 * MROUND multiples, parity tests and the #N/A family. Names and semantics
 * follow Excel/WPS (etcore evidence: KF_Exp, KF_Ceiling, KF_Floor, KF_Error, …).
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

export const FUNCTIONS: Record<string, Fn> = {
  EXP: (ev, n, s) => Math.exp(numArg(ev, n, s, 0)),
  LN: (ev, n, s) => {
    const v = numArg(ev, n, s, 0);
    if (v <= 0) err("#NUM!");
    return Math.log(v);
  },
  LOG: (ev, n, s) => {
    const v = numArg(ev, n, s, 0);
    const base = n.args[1] ? numArg(ev, n, s, 1) : 10;
    if (v <= 0 || base <= 0) err("#NUM!");
    if (base === 1) err("#DIV/0!");
    return Math.log(v) / Math.log(base);
  },
  LOG10: (ev, n, s) => {
    const v = numArg(ev, n, s, 0);
    if (v <= 0) err("#NUM!");
    return Math.log10(v);
  },
  CEILING: (ev, n, s) => {
    const v = numArg(ev, n, s, 0);
    const sig = n.args[1] ? numArg(ev, n, s, 1) : 1;
    if (sig === 0) return 0;
    if (v > 0 && sig < 0) err("#NUM!");
    return Math.ceil(v / sig) * sig;
  },
  FLOOR: (ev, n, s) => {
    const v = numArg(ev, n, s, 0);
    const sig = n.args[1] ? numArg(ev, n, s, 1) : 1;
    if (sig === 0) err("#DIV/0!");
    if (v > 0 && sig < 0) err("#NUM!");
    return Math.floor(v / sig) * sig;
  },
  MROUND: (ev, n, s) => {
    const v = numArg(ev, n, s, 0);
    const m = numArg(ev, n, s, 1);
    if (m === 0) return 0;
    if (v > 0 !== m > 0) err("#NUM!");
    return Math.round(v / m) * m;
  },
  ISEVEN: (ev, n, s) => Math.trunc(numArg(ev, n, s, 0)) % 2 === 0,
  ISODD: (ev, n, s) => Math.trunc(numArg(ev, n, s, 0)) % 2 !== 0,
  IFNA: (ev, n, s) => {
    const v = n.args[0] ? ev.sub(n.args[0], s) : null;
    return v === "#N/A" ? (n.args[1] ? ev.sub(n.args[1], s) : null) : v;
  },
  ISNA: (ev, n, s) => (n.args[0] ? ev.sub(n.args[0], s) : null) === "#N/A",
  NA: () => "#N/A",
};
