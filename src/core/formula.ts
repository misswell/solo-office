/**
 * Formula language: tokenizer, precedence-climbing parser, and evaluator.
 *
 * The token taxonomy mirrors the boundary observed in WPS's xlsx reader, which
 * calls CreateBoolToken / CreateDblToken / CreateErrorToken / CreateFunctionToken
 * / CreateHandleToken into the solver factory exported as _CreateSolverSystem.
 */

export type ErrorValue =
  | "#DIV/0!"
  | "#VALUE!"
  | "#REF!"
  | "#NAME?"
  | "#NUM!"
  | "#N/A"
  | "#NULL!"
  | "#CYCLE!";

export type Value = number | string | boolean | ErrorValue | null;

/** A bare range node evaluates to its cells; scalar contexts take the top-left one. */
export type Evaluated = Value | Value[];

export type Node =
  | { t: "num"; v: number }
  | { t: "str"; v: string }
  | { t: "bool"; v: boolean }
  | { t: "err"; v: ErrorValue }
  | { t: "ref"; sheet?: string; col: number; row: number }
  | { t: "range"; sheet?: string; from: { col: number; row: number }; to: { col: number; row: number } }
  | { t: "name"; v: string }
  | { t: "fn"; v: string; args: Node[] }
  | { t: "bin"; op: string; l: Node; r: Node }
  | { t: "un"; op: string; e: Node };

export const ERRORS: ErrorValue[] = [
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#NULL!",
  "#CYCLE!",
];

export const colToIndex = (col: string): number => {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};

export const indexToCol = (index: number): string => {
  let s = "";
  let n = index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - r) / 26);
  }
  return s;
};

export const refToString = (col: number, row: number): string => `${indexToCol(col)}${row + 1}`;

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/;

/**
 * A sheet qualifier ends at the first `!`, which is illegal inside a sheet name,
 * so unquoted and CJK names split the same way as `'quoted'!`.
 */
function splitSheet(token: string): { sheet?: string; rest: string } {
  const bang = token.indexOf("!");
  if (bang < 0) return { rest: token };
  const raw = token.slice(0, bang);
  const quoted = /^'(.*)'$/.exec(raw);
  return { sheet: quoted ? quoted[1] : raw, rest: token.slice(bang + 1) };
}

// ---------------------------------------------------------------- tokenizer

type Tok =
  | { k: "num"; v: number }
  | { k: "str"; v: string }
  | { k: "err"; v: ErrorValue }
  | { k: "cell"; v: string }
  | { k: "name"; v: string }
  | { k: "op"; v: string }
  | { k: "eof" };

const OPS = ["<=", ">=", "<>", "+", "-", "*", "/", "^", "&", "=", "<", ">", "(", ")", ",", ":", "%"];

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (c === " " || c === "\t" || c === "\n" || c === "\r") {
      i++;
      continue;
    }
    if (c === '"') {
      let s = "";
      i++;
      while (i < src.length) {
        if (src[i] === '"') {
          if (src[i + 1] === '"') {
            s += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += src[i++];
      }
      out.push({ k: "str", v: s });
      continue;
    }
    if (c === "'") {
      let s = "";
      i++;
      while (i < src.length && src[i] !== "'") s += src[i++];
      i++;
      if (src[i] === "!") {
        i++;
        const tail = readCellOrRangeTail(src, i);
        i += tail.length;
        out.push({ k: "cell", v: `'${s}'!${tail}` });
        continue;
      }
      out.push({ k: "name", v: s });
      continue;
    }
    if (c === "#") {
      let s = c;
      i++;
      while (i < src.length && /[A-Z0-9\/!?,.]/i.test(src[i])) s += src[i++];
      const err = ERRORS.find((e) => e.toUpperCase() === s.toUpperCase());
      out.push(err ? { k: "err", v: err } : { k: "name", v: s });
      continue;
    }
    if (/[0-9]/.test(c) || (c === "." && /[0-9]/.test(src[i + 1] ?? ""))) {
      let s = "";
      while (i < src.length && /[0-9.eE+\-]/.test(src[i])) {
        if (/[+\-]/.test(src[i]) && !/[eE]/.test(src[i - 1] ?? "")) break;
        s += src[i++];
      }
      out.push({ k: "num", v: Number(s) });
      continue;
    }
    if (/[\p{L}_$]/u.test(c)) {
      let s = "";
      // A `!` continues an unquoted sheet qualifier (e.g. 数据2!B2); names keep every other letter.
      while (i < src.length && (/[\p{L}\p{N}_$.]/u.test(src[i]) || (src[i] === "!" && /[\p{L}$.]/u.test(src[i + 1] ?? "")))) s += src[i++];
      if (src[i] === "(") {
        out.push({ k: "name", v: s });
        continue;
      }
      const { sheet, rest } = splitSheet(s);
      if (sheet && CELL_RE.test(rest)) {
        out.push({ k: "cell", v: s });
        continue;
      }
      if (CELL_RE.test(s)) {
        out.push({ k: "cell", v: s });
        continue;
      }
      out.push({ k: "name", v: s });
      continue;
    }
    const op = OPS.find((o) => src.startsWith(o, i));
    if (!op) throw new Error(`公式中无法识别的字符: ${c}`);
    i += op.length;
    out.push({ k: "op", v: op });
  }
  out.push({ k: "eof" });
  return out;
}

function readCellOrRangeTail(src: string, from: number): string {
  let s = "";
  let i = from;
  while (i < src.length && /[A-Za-z0-9$]/.test(src[i])) s += src[i++];
  return s;
}

// ------------------------------------------------------------------- parser

class ParseError extends Error {}

export function parse(formula: string): Node {
  const src = formula.startsWith("=") ? formula.slice(1) : formula;
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const eatOp = (v: string) => {
    const t = toks[p];
    if (t.k === "op" && t.v === v) {
      p++;
      return true;
    }
    return false;
  };

  const parseComparison = (): Node => {
    let l = parseConcat();
    for (;;) {
      const t = peek();
      if (t.k === "op" && ["=", "<", ">", "<=", ">=", "<>"].includes(t.v)) {
        p++;
        l = { t: "bin", op: t.v, l, r: parseConcat() };
      } else return l;
    }
  };

  const parseConcat = (): Node => {
    let l = parseAdd();
    while (eatOp("&")) l = { t: "bin", op: "&", l, r: parseAdd() };
    return l;
  };

  const parseAdd = (): Node => {
    let l = parseMul();
    for (;;) {
      const t = peek();
      if (t.k === "op" && (t.v === "+" || t.v === "-")) {
        p++;
        l = { t: "bin", op: t.v, l, r: parseMul() };
      } else return l;
    }
  };

  const parseMul = (): Node => {
    let l = parsePower();
    for (;;) {
      const t = peek();
      if (t.k === "op" && (t.v === "*" || t.v === "/")) {
        p++;
        l = { t: "bin", op: t.v, l, r: parsePower() };
      } else if (t.k === "op" && t.v === ":") {
        // implicit intersection is out of scope; range handled in parsePrimary
        return l;
      } else return l;
    }
  };

  /** `^` is right-associative and binds looser than sign or `%`, so -2^2 is 4. */
  const parsePower = (): Node => {
    const l = parseUnary();
    if (peek().k === "op" && (peek() as any).v === "^") {
      p++;
      return { t: "bin", op: "^", l, r: parsePower() };
    }
    return l;
  };

  const parseUnary = (): Node => {
    const t = peek();
    if (t.k === "op" && (t.v === "-" || t.v === "+")) {
      p++;
      return { t: "un", op: t.v, e: parseUnary() };
    }
    return parsePostfix();
  };

  const parsePostfix = (): Node => {
    let e = parsePrimary();
    while (eatOp("%")) e = { t: "un", op: "%", e };
    return e;
  };

  const cellNode = (raw: string): { col: number; row: number; sheet?: string } => {
    const { sheet, rest } = splitSheet(raw);
    const m = CELL_RE.exec(rest);
    if (!m) throw new ParseError(`无效引用 ${raw}`);
    return { sheet, col: colToIndex(m[1]), row: Number(m[2]) - 1 };
  };

  const parsePrimary = (): Node => {
    const t = peek();
    if (t.k === "op" && t.v === "(") {
      p++;
      const e = parseComparison();
      if (!eatOp(")")) throw new ParseError("缺少右括号");
      return e;
    }
    if (t.k === "num") {
      p++;
      return { t: "num", v: t.v };
    }
    if (t.k === "str") {
      p++;
      return { t: "str", v: t.v };
    }
    if (t.k === "err") {
      p++;
      return { t: "err", v: t.v };
    }
    if (t.k === "cell") {
      p++;
      const c = cellNode(t.v);
      if (eatOp(":")) {
        const nt = peek();
        if (nt.k !== "cell") throw new ParseError("范围右端不是单元格");
        p++;
        const d = cellNode(nt.v);
        return { t: "range", sheet: c.sheet, from: { col: c.col, row: c.row }, to: { col: d.col, row: d.row } };
      }
      return { t: "ref", sheet: c.sheet, col: c.col, row: c.row };
    }
    if (t.k === "name") {
      p++;
      if (peek().k === "op" && (peek() as any).v === "(") {
        p++;
        const args: Node[] = [];
        if (!eatOp(")")) {
          for (;;) {
            args.push(parseComparison());
            if (eatOp(",")) continue;
            if (!eatOp(")")) throw new ParseError("函数参数缺少右括号");
            break;
          }
        }
        return { t: "fn", v: t.v.toUpperCase(), args };
      }
      const up = t.v.toUpperCase();
      if (up === "TRUE" || up === "FALSE") return { t: "bool", v: up === "TRUE" };
      return { t: "name", v: t.v };
    }
    throw new ParseError(`意外的记号 ${JSON.stringify(t)}`);
  };

  const ast = parseComparison();
  if (peek().k !== "eof") throw new ParseError(`公式尾部有多余内容 ${JSON.stringify(peek())}`);
  return ast;
}

// -------------------------------------------------------------- coercion

/** Drops a multi-cell range to its top-left value, the way implicit intersection does. */
export const scalar = (v: Evaluated): Value => (Array.isArray(v) ? (v[0] ?? null) : v);

export const isError = (v: Value): v is ErrorValue =>
  typeof v === "string" && (ERRORS as string[]).includes(v);

export function toNumber(v: Value, strict = true): number | ErrorValue {
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (v === null || v === undefined) return strict ? "#VALUE!" : 0;
  if (isError(v)) return v;
  const n = Number(v.trim());
  return Number.isFinite(n) ? n : "#VALUE!";
}

export const toText = (v: Value): string => {
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return formatGeneral(v);
  return v;
};

/** Excel's General format: up to 11 significant digits, exponent past that. */
export function formatGeneral(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e11) return String(n);
  const s = n.toPrecision(11);
  return String(Number(s));
}

// ------------------------------------------------------------ evaluation

export interface CellSource {
  /** Raw stored value of a cell (already resolved for constants). */
  constant(sheet: string, col: number, row: number): Value;
  formula(sheet: string, col: number, row: number): string | undefined;
  name(scope: string): Node | undefined;
}

export class Evaluator {
  private readonly memo = new Map<string, Value>();
  private readonly visiting = new Set<string>();

  constructor(
    private readonly src: CellSource,
    private readonly sheet: string,
  ) {}

  cellValue(sheet: string, col: number, row: number): Value {
    const key = `${sheet}!${col}:${row}`;
    const hit = this.memo.get(key);
    if (hit !== undefined) return hit;
    const f = this.src.formula(sheet, col, row);
    if (f === undefined) {
      const v = this.src.constant(sheet, col, row);
      this.memo.set(key, v);
      return v;
    }
    if (this.visiting.has(key)) return "#CYCLE!";
    this.visiting.add(key);
    let v: Value;
    try {
      v = scalar(this.eval(parse(f), sheet));
    } catch (e) {
      v = e instanceof Error && (ERRORS as string[]).includes(e.message) ? (e.message as ErrorValue) : "#VALUE!";
    } finally {
      this.visiting.delete(key);
    }
    this.memo.set(key, v);
    return v;
  }

  eval(node: Node, sheet: string): Evaluated {
    switch (node.t) {
      case "num":
        return node.v;
      case "str":
        return node.v;
      case "bool":
        return node.v;
      case "err":
        return node.v;
      case "ref":
        return this.cellValue(node.sheet ?? sheet, node.col, node.row);
      case "range": {
        const cells = this.expand(node, sheet);
        return cells.length === 1 ? cells[0] : cells;
      }
      case "name": {
        const def = this.src.name(node.v);
        if (!def) return "#NAME?";
        return this.eval(def, sheet);
      }
      case "un": {
        const v = scalar(this.eval(node.e, sheet));
        if (isError(v)) return v;
        if (node.op === "%") {
          const n = toNumber(v);
          return typeof n === "number" ? n / 100 : n;
        }
        const n = toNumber(v);
        if (typeof n !== "number") return n;
        return node.op === "-" ? -n : n;
      }
      case "bin":
        return this.binary(node, sheet);
      case "fn":
        return this.call(node, sheet);
    }
  }

  expand(node: Node & { t: "range" }, sheet: string): Value[] {
    const s = node.sheet ?? sheet;
    const c1 = Math.min(node.from.col, node.to.col);
    const c2 = Math.max(node.from.col, node.to.col);
    const r1 = Math.min(node.from.row, node.to.row);
    const r2 = Math.max(node.from.row, node.to.row);
    const out: Value[] = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(this.cellValue(s, c, r));
    return out;
  }

  private flatten(v: Evaluated): Value[] {
    return Array.isArray(v) ? v : [v];
  }

  private numbers(v: Value | Value[], ignoreErrors = false): number[] | ErrorValue {
    const out: number[] = [];
    for (const item of this.flatten(v)) {
      if (item === null) continue;
      if (isError(item)) {
        if (ignoreErrors) continue;
        return item;
      }
      if (typeof item === "boolean") {
        out.push(item ? 1 : 0);
        continue;
      }
      if (typeof item === "string") {
        if (item === "") continue;
        const n = Number(item);
        if (Number.isFinite(n)) out.push(n);
        else if (!ignoreErrors) return "#VALUE!";
        continue;
      }
      out.push(item);
    }
    return out;
  }

  private binary(node: Node & { t: "bin" }, sheet: string): Value {
    const l = scalar(this.eval(node.l, sheet));
    if (isError(l)) return l;
    const r = scalar(this.eval(node.r, sheet));
    if (isError(r)) return r;
    if (node.op === "&") return toText(l) + toText(r);
    if (["=", "<>", "<", ">", "<=", ">="].includes(node.op)) {
      const cmp = compare(l, r);
      if (cmp === null) return "#VALUE!";
      switch (node.op) {
        case "=":
          return cmp === 0;
        case "<>":
          return cmp !== 0;
        case "<":
          return cmp < 0;
        case ">":
          return cmp > 0;
        case "<=":
          return cmp <= 0;
        default:
          return cmp >= 0;
      }
    }
    const a = toNumber(l);
    if (typeof a !== "number") return a;
    const b = toNumber(r);
    if (typeof b !== "number") return b;
    switch (node.op) {
      case "+":
        return a + b;
      case "-":
        return a - b;
      case "*":
        return a * b;
      case "/":
        return b === 0 ? "#DIV/0!" : a / b;
      case "^":
        return a ** b;
      default:
        return "#VALUE!";
    }
  }

  private call(node: Node & { t: "fn" }, sheet: string): Value {
    const arg = (i: number): Evaluated => {
      const a = node.args[i];
      if (!a) throw new Error("#N/A");
      return this.eval(a, sheet);
    };
    const numArg = (i: number): number | ErrorValue => {
      const v = arg(i);
      const flat = this.flatten(v);
      const n = toNumber(flat[0] ?? null);
      return n;
    };
    const nums = (from = 0): number[] | ErrorValue => {
      const all: Value[] = [];
      for (let i = from; i < node.args.length; i++) all.push(...this.flatten(this.eval(node.args[i], sheet)));
      return this.numbers(all);
    };
    const F = FUNCTIONS[node.v];
    if (!F) return "#NAME?";
    try {
      return F(this, node, sheet);
    } catch (e) {
      if (e instanceof Error && (ERRORS as string[]).includes(e.message)) return e.message as ErrorValue;
      throw e;
    }
  }

  /** Exposed for the function library. */
  sub(node: Node, sheet: string): Value {
    return scalar(this.eval(node, sheet));
  }
  list(node: Node, sheet: string): Value[] {
    const v = this.eval(node, sheet);
    return this.flatten(v);
  }
}

function compare(l: Value, r: Value): number | null {
  if (typeof l === "number" || typeof r === "number") {
    const a = toNumber(l, false);
    const b = toNumber(r, false);
    if (typeof a !== "number" || typeof b !== "number") return null;
    return a === b ? 0 : a < b ? -1 : 1;
  }
  if (typeof l === "boolean" || typeof r === "boolean") {
    const a = l === true ? 1 : 0;
    const b = r === true ? 1 : 0;
    return a === b ? 0 : a < b ? -1 : 1;
  }
  const a = String(l ?? "");
  const b = String(r ?? "");
  return a === b ? 0 : a < b ? -1 : 1;
}

// -------------------------------------------------------- function library

export type Fn = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string) => Value;

export const collect = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string, from = 0): Value[] => {
  const out: Value[] = [];
  for (let i = from; i < node.args.length; i++) out.push(...ev.list(node.args[i], sheet));
  return out;
};
export const numeric = (vals: Value[]): number[] =>
  vals.filter((v): v is number => typeof v === "number" || typeof v === "boolean" ? true : typeof v === "string" && v !== "" && Number.isFinite(Number(v)))
    .map((v) => (typeof v === "boolean" ? (v ? 1 : 0) : Number(v)));
export const first = (ev: Evaluator, node: Node & { t: "fn" }, sheet: string, i = 0): Value =>
  node.args[i] ? (ev.list(node.args[i], sheet)[0] ?? null) : null;

const serialOf = (d: Date): number => Math.round((d.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
const dateOf = (serial: number): Date => new Date(Math.round(serial) * 86400000 + Date.UTC(1899, 11, 30));

// Core table; the statistical, financial and math/info extension libraries
// register alongside it (each WPS function has its own KF_* class in etcore —
// see .scratch/rea-notes.md §8).
const BASE_FUNCTIONS: Record<string, Fn> = {
  SUM: (ev, n, s) => numeric(collect(ev, n, s)).reduce((a, b) => a + b, 0),
  PRODUCT: (ev, n, s) => numeric(collect(ev, n, s)).reduce((a, b) => a * b, 1),
  AVERAGE: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    if (!v.length) throw new Error("#DIV/0!");
    return v.reduce((a, b) => a + b, 0) / v.length;
  },
  MIN: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    return v.length ? Math.min(...v) : 0;
  },
  MAX: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    return v.length ? Math.max(...v) : 0;
  },
  COUNT: (ev, n, s) => numeric(collect(ev, n, s)).length,
  COUNTA: (ev, n, s) => collect(ev, n, s).filter((v) => v !== null && v !== "").length,
  COUNTBLANK: (ev, n, s) => collect(ev, n, s).filter((v) => v === null || v === "").length,
  COUNTIF: (ev, n, s) => {
    const range = ev.list(n.args[0], s);
    const crit = first(ev, n, s, 1);
    return range.filter((v) => matchCriteria(v, crit)).length;
  },
  SUMIF: (ev, n, s) => {
    const range = ev.list(n.args[0], s);
    const crit = first(ev, n, s, 1);
    const sum = n.args[2] ? ev.list(n.args[2], s) : range;
    return range.reduce<number>((acc, v, i) => (matchCriteria(v, crit) ? acc + (numeric([sum[i] ?? null])[0] ?? 0) : acc), 0);
  },
  AVERAGEIF: (ev, n, s) => {
    const range = ev.list(n.args[0], s);
    const crit = first(ev, n, s, 1);
    const vals = (n.args[2] ? ev.list(n.args[2], s) : range).filter((_, i) => matchCriteria(range[i], crit));
    const nums = numeric(vals);
    if (!nums.length) throw new Error("#DIV/0!");
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  },
  IF: (ev, n, s) => {
    const c = first(ev, n, s, 0);
    if (isError(c)) return c;
    const truthy = c === true || (typeof c === "number" && c !== 0) || (typeof c === "string" && c.toUpperCase() === "TRUE");
    return truthy ? (n.args[1] ? first(ev, n, s, 1) : true) : n.args[2] ? first(ev, n, s, 2) : false;
  },
  IFERROR: (ev, n, s) => {
    const v = first(ev, n, s, 0);
    return isError(v) ? first(ev, n, s, 1) : v;
  },
  AND: (ev, n, s) => collect(ev, n, s).every((v) => v === true || v === 1),
  OR: (ev, n, s) => collect(ev, n, s).some((v) => v === true || v === 1),
  NOT: (ev, n, s) => !first(ev, n, s, 0),
  ABS: (ev, n, s) => {
    const v = numeric([first(ev, n, s, 0)])[0];
    if (v === undefined) throw new Error("#VALUE!");
    return Math.abs(v);
  },
  SIGN: (ev, n, s) => Math.sign(numeric([first(ev, n, s, 0)])[0] ?? NaN),
  INT: (ev, n, s) => Math.floor(numeric([first(ev, n, s, 0)])[0] ?? NaN),
  TRUNC: (ev, n, s) => Math.trunc(numeric([first(ev, n, s, 0)])[0] ?? NaN),
  ROUND: (ev, n, s) => {
    const v = numeric([first(ev, n, s, 0)])[0] ?? NaN;
    const d = numeric([first(ev, n, s, 1)])[0] ?? 0;
    const f = 10 ** d;
    return Math.sign(v) * Math.round(Math.abs(v) * f) / f;
  },
  ROUNDUP: (ev, n, s) => {
    const v = numeric([first(ev, n, s, 0)])[0] ?? NaN;
    const d = numeric([first(ev, n, s, 1)])[0] ?? 0;
    const f = 10 ** d;
    return Math.sign(v) * Math.ceil(Math.abs(v) * f) / f;
  },
  ROUNDDOWN: (ev, n, s) => {
    const v = numeric([first(ev, n, s, 0)])[0] ?? NaN;
    const d = numeric([first(ev, n, s, 1)])[0] ?? 0;
    const f = 10 ** d;
    return Math.sign(v) * Math.floor(Math.abs(v) * f) / f;
  },
  MOD: (ev, n, s) => {
    const a = numeric([first(ev, n, s, 0)])[0] ?? NaN;
    const b = numeric([first(ev, n, s, 1)])[0] ?? NaN;
    if (b === 0) throw new Error("#DIV/0!");
    return a - b * Math.floor(a / b);
  },
  POWER: (ev, n, s) => {
    const a = numeric([first(ev, n, s, 0)])[0] ?? NaN;
    return a ** (numeric([first(ev, n, s, 1)])[0] ?? NaN);
  },
  SQRT: (ev, n, s) => {
    const v = numeric([first(ev, n, s, 0)])[0] ?? NaN;
    if (v < 0) throw new Error("#NUM!");
    return Math.sqrt(v);
  },
  SUMPRODUCT: (ev, n, s) => {
    const cols = n.args.map((a) => numeric(ev.list(a, s)));
    const len = Math.max(...cols.map((c) => c.length));
    return Array.from({ length: len }, (_, i) => cols.reduce((acc, c) => acc * (c[i] ?? 0), 1)).reduce((a, b) => a + b, 0);
  },
  CONCATENATE: (ev, n, s) => collect(ev, n, s).map(toText).join(""),
  CONCAT: (ev, n, s) => collect(ev, n, s).map(toText).join(""),
  TEXTJOIN: (ev, n, s) => {
    const delim = toText(first(ev, n, s, 0));
    const skip = first(ev, n, s, 1) === true;
    return collect(ev, n, s, 2)
      .filter((v) => !(skip && (v === null || v === "")))
      .map(toText)
      .join(delim);
  },
  LEFT: (ev, n, s) => toText(first(ev, n, s, 0)).slice(0, numeric([first(ev, n, s, 1)])[0] ?? 1),
  RIGHT: (ev, n, s) => {
    const t = toText(first(ev, n, s, 0));
    const k = numeric([first(ev, n, s, 1)])[0] ?? 1;
    return t.slice(Math.max(0, t.length - k));
  },
  MID: (ev, n, s) => {
    const t = toText(first(ev, n, s, 0));
    const start = numeric([first(ev, n, s, 1)])[0] ?? 1;
    const len = numeric([first(ev, n, s, 2)])[0] ?? 0;
    return t.slice(start - 1, start - 1 + len);
  },
  LEN: (ev, n, s) => toText(first(ev, n, s, 0)).length,
  UPPER: (ev, n, s) => toText(first(ev, n, s, 0)).toUpperCase(),
  LOWER: (ev, n, s) => toText(first(ev, n, s, 0)).toLowerCase(),
  PROPER: (ev, n, s) => toText(first(ev, n, s, 0)).replace(/\w\S*/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase()),
  TRIM: (ev, n, s) => toText(first(ev, n, s, 0)).trim().replace(/\s+/g, " "),
  REPLACE: (ev, n, s) => {
    const t = toText(first(ev, n, s, 0));
    const start = numeric([first(ev, n, s, 1)])[0] ?? 1;
    const len = numeric([first(ev, n, s, 2)])[0] ?? 0;
    return t.slice(0, start - 1) + toText(first(ev, n, s, 3)) + t.slice(start - 1 + len);
  },
  FIND: (ev, n, s) => {
    const needle = toText(first(ev, n, s, 0));
    const hay = toText(first(ev, n, s, 1));
    const at = hay.indexOf(needle, (numeric([first(ev, n, s, 2)])[0] ?? 1) - 1);
    if (at < 0) throw new Error("#VALUE!");
    return at + 1;
  },
  SEARCH: (ev, n, s) => {
    const needle = toText(first(ev, n, s, 0)).toLowerCase();
    const hay = toText(first(ev, n, s, 1)).toLowerCase();
    const at = hay.indexOf(needle);
    if (at < 0) throw new Error("#VALUE!");
    return at + 1;
  },
  ISBLANK: (ev, n, s) => first(ev, n, s, 0) === null || first(ev, n, s, 0) === "",
  ISERROR: (ev, n, s) => isError(first(ev, n, s, 0)),
  ISNUMBER: (ev, n, s) => typeof first(ev, n, s, 0) === "number",
  ISTEXT: (ev, n, s) => typeof first(ev, n, s, 0) === "string",
  IFS: (ev, n, s) => {
    for (let i = 0; i + 1 < n.args.length; i += 2) {
      if (scalar(ev.eval(n.args[i], s)) === true) return scalar(ev.eval(n.args[i + 1], s));
    }
    throw new Error("#N/A");
  },
  TODAY: () => serialOf(new Date(new Date().toISOString().slice(0, 10))),
  NOW: () => (Date.now() - Date.UTC(1899, 11, 30)) / 86400000,
  DATE: (ev, n, s) => {
    const [y, m, d] = [0, 1, 2].map((i) => numeric([first(ev, n, s, i)])[0] ?? 0);
    return serialOf(new Date(Date.UTC(y, m - 1, d)));
  },
  YEAR: (ev, n, s) => dateOf(numeric([first(ev, n, s, 0)])[0] ?? 0).getUTCFullYear(),
  MONTH: (ev, n, s) => dateOf(numeric([first(ev, n, s, 0)])[0] ?? 0).getUTCMonth() + 1,
  DAY: (ev, n, s) => dateOf(numeric([first(ev, n, s, 0)])[0] ?? 0).getUTCDate(),
  SUMSQ: (ev, n, s) => numeric(collect(ev, n, s)).reduce((a, b) => a + b * b, 0),
  STDEV: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    if (v.length < 2) throw new Error("#DIV/0!");
    const m = v.reduce((a, b) => a + b, 0) / v.length;
    return Math.sqrt(v.reduce((a, b) => a + (b - m) ** 2, 0) / (v.length - 1));
  },
  VAR: (ev, n, s) => {
    const v = numeric(collect(ev, n, s));
    if (v.length < 2) throw new Error("#DIV/0!");
    const m = v.reduce((a, b) => a + b, 0) / v.length;
    return v.reduce((a, b) => a + (b - m) ** 2, 0) / (v.length - 1);
  },
  MEDIAN: (ev, n, s) => {
    const v = numeric(collect(ev, n, s)).sort((a, b) => a - b);
    if (!v.length) throw new Error("#NUM!");
    const mid = v.length >> 1;
    return v.length % 2 ? v[mid] : (v[mid - 1] + v[mid]) / 2;
  },
  LARGE: (ev, n, s) => {
    const v = numeric(ev.list(n.args[0], s)).sort((a, b) => b - a);
    return v[(numeric([first(ev, n, s, 1)])[0] ?? 1) - 1] ?? throw0();
  },
  SMALL: (ev, n, s) => {
    const v = numeric(ev.list(n.args[0], s)).sort((a, b) => a - b);
    return v[(numeric([first(ev, n, s, 1)])[0] ?? 1) - 1] ?? throw0();
  },
  RANK: (ev, n, s) => {
    const x = numeric([first(ev, n, s, 0)])[0] ?? 0;
    const v = numeric(ev.list(n.args[1], s));
    return v.filter((b) => b > x).length + 1;
  },
};

function throw0(): never {
  throw new Error("#NUM!");
}

import { FUNCTIONS as STAT_FUNCTIONS } from "./formula-stats.js";
import { FUNCTIONS as FINANCIAL_FUNCTIONS } from "./formula-financial.js";
import { FUNCTIONS as MATH_FUNCTIONS } from "./formula-math.js";
const FUNCTIONS: Record<string, Fn> = { ...BASE_FUNCTIONS, ...STAT_FUNCTIONS, ...FINANCIAL_FUNCTIONS, ...MATH_FUNCTIONS };

/** Excel criteria semantics: plain equality plus comparison-operator and wildcard forms. */
export function matchCriteria(value: Value, criteria: Value): boolean {
  if (typeof criteria !== "string") return compare(value, criteria) === 0;
  const m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(criteria);
  if (m) {
    const rhs = coerceLiteral(m[2]);
    const c = compare(value, rhs);
    if (c === null) return false;
    switch (m[1]) {
      case "<":
        return c < 0;
      case ">":
        return c > 0;
      case "<=":
        return c <= 0;
      case ">=":
        return c >= 0;
      case "=":
        return c === 0;
      default:
        return c !== 0;
    }
  }
  if (/[*?]/.test(criteria)) {
    const re = new RegExp("^" + criteria.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
    return re.test(toText(value));
  }
  const rhs = coerceLiteral(criteria);
  return compare(value, rhs) === 0;
}

function coerceLiteral(s: string): Value {
  if (s === "") return null;
  if (/^-?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(s)) return Number(s);
  if (s.toUpperCase() === "TRUE") return true;
  if (s.toUpperCase() === "FALSE") return false;
  return s;
}
