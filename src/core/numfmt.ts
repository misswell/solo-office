/** ECMA-376 number format codes: built-in table plus a practical subset of custom codes. */

export const BUILTIN: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: "$#,##0_);($#,##0)",
  6: "$#,##0_);[Red]($#,##0)",
  7: "$#,##0.00_);($#,##0.00)",
  8: "$#,##0.00_);[Red]($#,##0.00)",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  41: '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ ',
  42: '_$* #,##0.00_ ;_$* -#,##0.00_ ;_$* "-"??_ ;_(@_)',
  43: '_ * #,##0_ ;_ * -#,##0_ ;_ * "-"??_ ;_ @_ ',
  44: '_$* #,##0_ ;_$* -#,##0_ ;_$* "-"??_ ;_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

const DATE_TOKEN = /(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm)/i;

export function isDateFormat(code: string): boolean {
  return DATE_TOKEN.test(code.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, ""));
}

const MONTHS_LONG = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const DAYS_LONG = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const DAYS_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

/** Excel serial (1900 system, with the historical leap-year bug) to UTC date parts. */
export function serialToParts(serial: number): { y: number; mo: number; d: number; h: number; mi: number; s: number; wd: number } {
  const ms = Math.round(serial * 86400000) + Date.UTC(1899, 11, 30);
  const dt = new Date(ms);
  return {
    y: dt.getUTCFullYear(),
    mo: dt.getUTCMonth() + 1,
    d: dt.getUTCDate(),
    h: dt.getUTCHours(),
    mi: dt.getUTCMinutes(),
    s: dt.getUTCSeconds(),
    wd: dt.getUTCDay(),
  };
}

function fixed(n: number, digits: number, thousands: boolean): string {
  const neg = n < 0;
  let s = Math.abs(n).toFixed(digits);
  if (thousands) {
    const [i, f] = s.split(".");
    s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (f ? "." + f : "");
  }
  return (neg ? "-" : "") + s;
}

/**
 * `General` shows the number itself: no padding, no grouping, at most eleven
 * characters, and scientific notation once it no longer fits.
 */
export function general(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  const a = Math.abs(n);
  if (a !== 0 && (a >= 1e11 || a < 1e-9)) return scientific(n);
  if (Number.isInteger(n)) return String(n);
  const plain = String(n);
  if (plain.length <= 11) return plain;
  const intDigits = Math.max(1, Math.floor(Math.log10(a)) + 1);
  const frac = Math.max(0, 10 - (n < 0 ? 1 : 0) - intDigits);
  return trimZeros(n.toFixed(frac));
}

const trimZeros = (s: string): string => (s.includes(".") ? s.replace(/\.?0+$/, "") : s);

function scientific(n: number): string {
  const [mantissa, exponent] = n.toExponential(9).split("e");
  const e = Number(exponent);
  return `${trimZeros(mantissa).toUpperCase()}E${e < 0 ? "-" : "+"}${String(Math.abs(e)).padStart(2, "0")}`;
}

/** `m` reads as minutes only when it touches an hour or second token; otherwise it is a month. */
function isMinutesAt(src: string, start: number, len: number): boolean {
  const before = src.slice(0, start).replace(/[^A-Za-z]/g, "").slice(-1);
  const after = src.slice(start + len).replace(/[^A-Za-z]/g, "").slice(0, 1);
  return /[hs]/i.test(before) || /[hs]/i.test(after);
}

function renderDate(code: string, serial: number): string {
  const p = { ...serialToParts(serial), eh: Math.floor(serial * 24) };
  const clock12 = /AM\/PM|am\/pm/.test(code);
  const hour = clock12 ? p.h % 12 || 12 : p.h;
  return code.replace(
    /\[h+\]|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm/g,
    (tok, offset: number) => {
      const minutes = isMinutesAt(code, offset, tok.length);
      switch (tok.toLowerCase()) {
        case "yyyy":
          return String(p.y);
        case "yy":
          return String(p.y % 100).padStart(2, "0");
        case "mmmm":
          return MONTHS_LONG[p.mo - 1];
        case "mmm":
          return MONTHS_SHORT[p.mo - 1];
        case "mm":
          return String(minutes ? p.mi : p.mo).padStart(2, "0");
        case "m":
          return String(minutes ? p.mi : p.mo);
        case "dddd":
          return DAYS_LONG[p.wd];
        case "ddd":
          return DAYS_SHORT[p.wd];
        case "dd":
          return String(p.d).padStart(2, "0");
        case "d":
          return String(p.d);
        case "hh":
          return String(hour).padStart(2, "0");
        case "h":
          return String(hour);
        case "ss":
          return String(p.s).padStart(2, "0");
        case "s":
          return String(p.s);
        default: {
          if (tok[0] === "[") return String(p.eh).padStart(tok.length - 2, "0");
          const meridiem = p.h < 12 ? "AM" : "PM";
          return tok[0] === tok[0].toLowerCase() ? meridiem.toLowerCase() : meridiem;
        }
      }
    },
  );
}

/**
 * Apply one format code. Section syntax `pos;neg;zero;text` is supported, as are
 * literals in quotes, `$`, `%`, `,` thousands and `;` separators.
 */
export function format(value: number | string | boolean | null, code: string): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") {
    if (value.startsWith("#")) return value;
    return code.includes("@") ? code.replace(/"@[^"]*"/g, "").replace(/@/g, value) : value;
  }
  const sections = splitSections(code);
  const idx = value < 0 ? 1 : value === 0 ? 2 : 0;
  const sec = sections[Math.min(idx, sections.length - 1)] ?? sections[0];
  if (sec === undefined) return String(value);
  if (/^\s*(general|@)\s*$/i.test(sec)) return general(value);
  if (isDateFormat(sec)) return renderDate(sec, value);

  const bare = sec.replace(/"[^"]*"/g, "").replace(/\\/g, "");
  const thousands = /,/.test(bare.replace(/\d+,\d+/, "")) || /#,\d|0,\d/.test(bare);
  const percent = bare.includes("%");
  const currency = bare.includes("$");
  const decimalMatch = /\.(0+|#+)/.exec(bare);
  const digits = decimalMatch ? decimalMatch[1].length : /0\.(0+)/.test(bare) ? RegExp.$1.length : bare.includes(".") ? 2 : 0;
  const intPattern = /(#,##0|#+0|0+)/.exec(bare.replace(/,/g, ""));
  const minInt = intPattern ? intPattern[1].replace(/[^0]/g, "").length : 1;

  let n = value;
  if (percent) n *= 100;
  let out = fixed(n, digits, thousands);
  if (minInt > 1) {
    const [i, f] = out.split(".");
    const pad = i.startsWith("-") ? "-" + i.slice(1).padStart(minInt, "0") : i.padStart(minInt, "0");
    out = f ? `${pad}.${f}` : pad;
  }
  if (currency) out = "$" + out;
  if (percent) out += "%";
  const negWrap = /^\(.*\)$/.test(sec.replace(/"[^"]*"/g, ""));
  if (negWrap && value < 0) out = "(" + out.replace("-", "") + ")";
  const red = /\[Red\]/i.test(sec);
  return red && value < 0 ? `\u0001${out}\u0002` : out;
}

function splitSections(code: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let cur = "";
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') {
      let j = i + 1;
      while (j < code.length && code[j] !== '"') j++;
      cur += code.slice(i, j + 1);
      i = j;
      continue;
    }
    if (c === "[") depth++;
    if (c === "]") depth--;
    if (c === ";" && depth === 0) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += c;
  }
  out.push(cur);
  return out;
}
