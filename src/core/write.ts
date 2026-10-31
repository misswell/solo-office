import { chartPartXml } from "./chart.js";
import { chartDataFingerprint } from "./chart-data.js";
import { filterFingerprint, type AutoFilter } from "./filter.js";
import { strToU8, zipSync } from "fflate";
import { ATTRS, buildXml, parseOrderedXml, rawAttrs, rawChildren, rawFind, rawFindAll, relsPathOf, tagLocalOf, tagOf, text as bytesText, type OpcPackage } from "./opc.js";
import { isError, type Value } from "./formula.js";
import { BUILTIN } from "./numfmt.js";
import { refOf } from "./sheet-ops.js";
import { bytesOfDataUrl } from "./media.js";
import { chartGeometryFingerprint, drawingAnchorAtPx, imageBoxPx, colsFingerprint, condFingerprint, imagesFingerprint, type Border, type Cell, type CellStyle, type ChartModel, type CondFormat, type CondRule, type Run, type Sheet, type SheetImage, type Spreadsheet } from "./xlsx.js";

/**
 * Serializes the model back to SpreadsheetML. Style records are written in the
 * model's own order, so a `s=` index in the sheet still means the same style
 * after a read/write round trip.
 */

const XML_ESC = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "application/vnd.openxmlformats-officedocument.spreadsheetml";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

const BUILTIN_IDS = new Map<string, number>();
for (const [id, code] of Object.entries(BUILTIN)) {
  if (!BUILTIN_IDS.has(code)) BUILTIN_IDS.set(code, Number(id));
}

interface StyleTables {
  fonts: string[];
  fills: string[];
  borders: string[];
  numFmts: Map<string, number>;
  xf: (style: CellStyle) => string;
}

const fontXmlOf = (s: CellStyle): string => {
  const parts = [`<sz val="${s.fontSize || 11}"/>`];
  if (s.fontColor) parts.push(`<color rgb="FF${s.fontColor.toUpperCase()}"/>`);
  parts.push(`<name val="${XML_ESC(s.fontName || "Calibri")}"/>`);
  const tag = (on: boolean, name: string): void => {
    if (on) parts.unshift(`<${name}/>`);
  };
  tag(s.underline, "u");
  tag(s.italic, "i");
  tag(s.bold, "b");
  return `<font>${parts.join("")}</font>`;
};

const fillXmlOf = (s: CellStyle): string | undefined =>
  s.bgColor
    ? `<fill><patternFill patternType="solid"><fgColor rgb="FF${s.bgColor.toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`
    : undefined;

const BORDER_SIDES = ["left", "right", "top", "bottom"] as const;

/** A side keeps the file's own line token and ink; a side with neither is written as "no line". */
const borderXmlOf = (b: Border): string =>
  `<border>${BORDER_SIDES.map((side) => {
    const line = b[side];
    if (!line) return `<${side}/>`;
    const style = XML_ESC(line.style || "thin");
    return `<${side} style="${style}">${line.color ? `<color rgb="FF${XML_ESC(line.color.toUpperCase())}"/>` : ""}</${side}>`;
  }).join("")}<diagonal/></border>`;

/** The `xf` record tying a style to font/fill/border/numFmt records already in the file. */
const xfTagOf = (s: CellStyle, ids: { numFmt: number; font: number; fill: number; border: number }): string => {
  const attrs = [
    `numFmtId="${ids.numFmt}"`,
    `fontId="${ids.font}"`,
    `fillId="${ids.fill}"`,
    `borderId="${ids.border}"`,
    `xfId="0"`,
    ids.numFmt ? `applyNumberFormat="1"` : "",
    ids.font ? `applyFont="1"` : "",
    ids.fill ? `applyFill="1"` : "",
    ids.border ? `applyBorder="1"` : "",
  ].filter(Boolean);
  const alignment: string[] = [];
  if (s.align) alignment.push(`horizontal="${XML_ESC(s.align)}"`);
  if (s.valign) alignment.push(`vertical="${XML_ESC(s.valign)}"`);
  if (s.wrap) alignment.push(`wrapText="1"`);
  if (alignment.length) attrs.push(`applyAlignment="1"`);
  const body = alignment.length ? `<alignment ${alignment.join(" ")}/>` : "";
  return `<xf ${attrs.join(" ")}>${body}</xf>`;
};

function buildStyles(styles: CellStyle[]): StyleTables {
  const fonts = [`<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>`];
  const fontIds = new Map<string, number>([[`<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>`, 0]]);
  // OOXML reserves fill 0 = none and fill 1 = gray125.
  const fills = [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`];
  const fillIds = new Map<string, number>();
  const borders = [`<border><left/><right/><top/><bottom/><diagonal/></border>`];
  const borderIds = new Map<string, number>([[keyOfBorder(undefined), 0]]);
  const numFmts = new Map<string, number>();
  let nextCustom = 164;

  const fontOf = (s: CellStyle): number => {
    const xml = fontXmlOf(s);
    let id = fontIds.get(xml);
    if (id === undefined) {
      id = fonts.push(xml) - 1;
      fontIds.set(xml, id);
    }
    return id;
  };

  const fillOf = (s: CellStyle): number => {
    const xml = fillXmlOf(s);
    if (!xml) return 0;
    let id = fillIds.get(xml);
    if (id === undefined) {
      id = fills.push(xml) - 1;
      fillIds.set(xml, id);
    }
    return id;
  };

  const borderOf = (s: CellStyle): number => {
    const key = keyOfBorder(s.border);
    let id = borderIds.get(key);
    if (id === undefined) {
      id = borders.push(borderXmlOf(s.border!)) - 1;
      borderIds.set(key, id);
    }
    return id;
  };

  const numFmtOf = (s: CellStyle): number => {
    const code = s.numFmt;
    if (!code || code === "General") return 0;
    const fixed = BUILTIN_IDS.get(code);
    if (fixed !== undefined) return fixed;
    let id = numFmts.get(code);
    if (id === undefined) {
      id = nextCustom++;
      numFmts.set(code, id);
    }
    return id;
  };

  const xf = (s: CellStyle): string => xfTagOf(s, { numFmt: numFmtOf(s), font: fontOf(s), fill: fillOf(s), border: borderOf(s) });

  return { fonts, fills, borders, numFmts, xf };
}

const keyOfBorder = (b: Border | undefined): string =>
  b
    ? JSON.stringify(BORDER_SIDES.map((side) => {
        const line = b[side];
        return line ? [line.style || "thin", (line.color ?? "").toUpperCase()] : null;
      }))
    : "";

/** A differential format record for an editor-created conditional rule. */
function dxfXmlOf(st: { bgColor?: string; fontColor?: string; bold?: boolean; italic?: boolean; strike?: boolean; underline?: boolean }): any {
  const fontKids: any[] = [];
  if (st.bold) fontKids.push(el("b"));
  if (st.italic) fontKids.push(el("i"));
  if (st.strike) fontKids.push(el("strike"));
  if (st.underline) fontKids.push(el("u"));
  if (st.fontColor) fontKids.push(el("color", { "@_rgb": `FF${st.fontColor}` }));
  const kids: any[] = [];
  if (fontKids.length) kids.push(el("font", undefined, fontKids));
  if (st.bgColor) kids.push(el("fill", undefined, [el("patternFill", undefined, [el("bgColor", { "@_rgb": `FF${st.bgColor}` })])]));
  return el("dxf", undefined, kids);
}

function stylesXml(styles: CellStyle[]): string {
  const t = buildStyles(styles);
  const numFmtTags = [...t.numFmts].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${XML_ESC(code)}"/>`).join("");
  const cellXfs = styles.map((s) => t.xf(s)).join("");
  return `${DECL}<styleSheet xmlns="${NS}">` +
    (numFmtTags ? `<numFmts count="${t.numFmts.size}">${numFmtTags}</numFmts>` : "") +
    `<fonts count="${t.fonts.length}">${t.fonts.join("")}</fonts>` +
    `<fills count="${t.fills.length}">${t.fills.join("")}</fills>` +
    `<borders count="${t.borders.length}">${t.borders.join("")}</borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="${styles.length}">${cellXfs}</cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
    `</styleSheet>`;
}

/** Styles the editor created land after the source's own records, in ids above its maxima. */
class StyleAppender {
  readonly fonts: string[] = [];
  readonly fills: string[] = [];
  readonly borders: string[] = [];
  readonly numFmts: { id: number; code: string }[] = [];
  readonly xfs: string[] = [];
  private fontIds = new Map<string, number>();
  private fillIds = new Map<string, number>();
  private borderIds = new Map<string, number>();
  private numFmtIds = new Map<string, number>();

  constructor(private base: { fonts: number; fills: number; borders: number; numFmt: number }) {}

  private font(s: CellStyle): number {
    const xml = fontXmlOf(s);
    let id = this.fontIds.get(xml);
    if (id === undefined) {
      id = this.base.fonts + this.fonts.length;
      this.fonts.push(xml);
      this.fontIds.set(xml, id);
    }
    return id;
  }

  private fill(s: CellStyle): number {
    const xml = fillXmlOf(s);
    if (!xml) return 0;
    let id = this.fillIds.get(xml);
    if (id === undefined) {
      id = this.base.fills + this.fills.length;
      this.fills.push(xml);
      this.fillIds.set(xml, id);
    }
    return id;
  }

  private border(s: CellStyle): number {
    if (!s.border) return 0;
    const key = keyOfBorder(s.border);
    let id = this.borderIds.get(key);
    if (id === undefined) {
      id = this.base.borders + this.borders.length;
      this.borders.push(borderXmlOf(s.border));
      this.borderIds.set(key, id);
    }
    return id;
  }

  private numFmt(s: CellStyle): number {
    const code = s.numFmt;
    if (!code || code === "General") return 0;
    const fixed = BUILTIN_IDS.get(code);
    if (fixed !== undefined) return fixed;
    let id = this.numFmtIds.get(code);
    if (id === undefined) {
      id = this.base.numFmt + this.numFmts.length;
      this.numFmts.push({ id, code });
      this.numFmtIds.set(code, id);
    }
    return id;
  }

  xf(s: CellStyle): string {
    const tag = xfTagOf(s, { numFmt: this.numFmt(s), font: this.font(s), fill: this.fill(s), border: this.border(s) });
    this.xfs.push(tag);
    return tag;
  }
}

/**
 * Where each model style lands in the saved `xl/styles.xml`. Records that came from
 * the source file keep their index and their bytes; only styles the editor created are
 * appended. Theme colours, tints and scheme fonts the model does not track — nearly
 * every WPS workbook has them — therefore survive an edited save untouched.
 */
interface StylePlan {
  /** `s=` index for a cell; `undefined` or 0 leaves the attribute off (xf 0). */
  sIndexOf: (cell: Cell) => number | undefined;
  xml: string;
}

function stylePlan(wb: Spreadsheet): StylePlan {
  const styles = wb.styleTable();
  const source = wb.source()?.text("xl/styles.xml");
  const canonical = (): StylePlan => {
    const dxfs = wb.editorDxfTable();
    const xml = dxfs.length
      ? stylesXml(styles).replace("</styleSheet>", `<dxfs count="${dxfs.length}">${dxfs.map((st) => buildXml([dxfXmlOf(st)])).join("")}</dxfs></styleSheet>`)
      : stylesXml(styles);
    return { sIndexOf: (cell) => cell.styleIndex, xml };
  };
  if (!source) return canonical();
  const tree = parseOrderedXml(source);
  const root = rawFind(tree, "styleSheet");
  const kids = root ? rawChildren(root) : undefined;
  if (!kids) return canonical();
  // Editor-created conditional styles join the source's dxfs, in ids above its own.
  const editorDxfs = wb.editorDxfTable();
  if (editorDxfs.length) {
    const node = childOf(kids, "dxfs");
    const records = editorDxfs.map((st) => dxfXmlOf(st));
    if (node) {
      rawChildren(node).push(...records);
      const attrs = rawAttrs(node);
      if (attrs["@_count"] !== undefined) attrs["@_count"] = String(Number(attrs["@_count"]) + records.length);
    } else {
      place(kids, "dxfs", el("dxfs", { "@_count": String(records.length) }, records), STYLESHEET_ORDER);
    }
  }
  const aliases = wb.styleAliases();
  const used = new Set<number>();
  for (const sheet of wb.sheets) for (const cell of sheet.cells.values()) used.add(cell.styleIndex);
  const extras: number[] = [];
  for (let i = 0; i < styles.length; i++) if (aliases[i] === undefined && used.has(i)) extras.push(i);
  if (!extras.length) return { sIndexOf: (cell) => aliases[cell.styleIndex], xml: source };

  const countOf = (tag: string, child: string): number => {
    const node = childOf(kids, tag);
    return node ? rawChildren(node).filter((entry) => tagLocalOf(entry) === child).length : 0;
  };
  let maxNumFmt = 163;
  const numFmtsNode = childOf(kids, "numFmts");
  for (const nf of numFmtsNode ? rawChildren(numFmtsNode) : []) {
    if (tagLocalOf(nf) === "numFmt") maxNumFmt = Math.max(maxNumFmt, Number(rawAttrs(nf)["@_numFmtId"] ?? 0));
  }
  const cellXfsNode = childOf(kids, "cellXfs");
  const base = cellXfsNode ? rawChildren(cellXfsNode).filter((entry) => tagLocalOf(entry) === "xf").length : 0;
  const append = new StyleAppender({
    fonts: countOf("fonts", "font"),
    fills: countOf("fills", "fill"),
    borders: countOf("borders", "border"),
    numFmt: maxNumFmt + 1,
  });
  const at = new Map<number, number>();
  extras.forEach((id, k) => {
    append.xf(styles[id]);
    at.set(id, base + k);
  });

  const insert = (tag: string, records: any[]): void => {
    if (!records.length) return;
    const node = childOf(kids, tag);
    if (!node) {
      place(kids, tag, el(tag, { "@_count": String(records.length) }, records), STYLESHEET_ORDER);
      return;
    }
    rawChildren(node).push(...records);
    const attrs = rawAttrs(node);
    if (attrs["@_count"] !== undefined) attrs["@_count"] = String(Number(attrs["@_count"]) + records.length);
  };
  const recordsOf = (xmls: string[]): any[] => {
    const wrap = rawFind(parseOrderedXml(`<r>${xmls.join("")}</r>`), "r");
    return wrap ? rawChildren(wrap) : [];
  };
  insert("numFmts", append.numFmts.map((nf) => el("numFmt", { "@_numFmtId": String(nf.id), "@_formatCode": nf.code })));
  insert("fonts", recordsOf(append.fonts));
  insert("fills", recordsOf(append.fills));
  insert("borders", recordsOf(append.borders));
  insert("cellXfs", recordsOf(append.xfs));
  return { sIndexOf: (cell) => aliases[cell.styleIndex] ?? at.get(cell.styleIndex), xml: serializeTree(tree) };
}

/** Strings become shared-string entries, keeping rich runs when a cell has them. */
function sharedStringsXml(cells: Iterable<Cell>): { xml: string; count: number; indexOf: (cell: Cell) => number | undefined } {
  const table: { key: string; runs: Run[]; text: string }[] = [];
  const ids = new Map<string, number>();
  const map = new Map<Cell, number>();
  for (const cell of cells) {
    const value = typeof cell.constant === "string" ? cell.constant : undefined;
    if (value === undefined || value === "") continue;
    const runs = cell.runs?.length ? cell.runs : [{ text: value }];
    const key = JSON.stringify(runs);
    let id = ids.get(key);
    if (id === undefined) {
      id = table.push({ key, runs, text: value }) - 1;
      ids.set(key, id);
    }
    map.set(cell, id);
  }
  const body = table
    .map((entry) =>
      entry.runs.length === 1
        ? `<si><t${/^\s|\s$/.test(entry.runs[0].text) ? ' xml:space="preserve"' : ""}>${XML_ESC(entry.runs[0].text)}</t></si>`
        : `<si>${entry.runs
            .map((r) => {
              const pr: string[] = [];
              if (r.bold) pr.push("<b/>");
              if (r.italic) pr.push("<i/>");
              if (r.underline) pr.push("<u/>");
              if (r.size) pr.push(`<sz val="${r.size}"/>`);
              if (r.color) pr.push(`<color rgb="FF${r.color.toUpperCase()}"/>`);
              return `<r>${pr.length ? `<rPr>${pr.join("")}</rPr>` : ""}<t>${XML_ESC(r.text)}</t></r>`;
            })
            .join("")}</si>`,
    )
    .join("");
  const xml = `${DECL}<sst xmlns="${NS}" count="${table.length}" uniqueCount="${table.length}">${body}</sst>`;
  return { xml, count: table.length, indexOf: (cell) => map.get(cell) };
}

// ----------------------------------------------------------- sheet part nodes

/** Ordered node, the shape `parseOrderedXml` reads and `buildXml` writes. */
const el = (tag: string, attrs?: Record<string, string>, children: any[] = []): any =>
  attrs && Object.keys(attrs).length ? { [tag]: children, [ATTRS]: attrs } : { [tag]: children };

const textNode = (value: string): any => ({ "#text": value });

/** Stored value doubles as the cached result, so other readers show it without recomputing. */
const cacheOf = (v: Value): { attrs?: Record<string, string>; value?: string } => {
  if (v === null || v === "") return {};
  if (typeof v === "number") return { value: jsonNumber(v) };
  if (typeof v === "boolean") return { attrs: { "@_t": "b" }, value: v ? "1" : "0" };
  if (isError(v)) return { attrs: { "@_t": "e" }, value: v };
  return { attrs: { "@_t": "str" }, value: String(v) };
};

/** What a sheet write needs besides the sheet itself: where strings and styles land in the file. */
interface SheetWriter {
  indexOf: (cell: Cell) => number | undefined;
  sIndexOf: (cell: Cell) => number | undefined;
  /** Relationship id of a freshly created drawing part, when the model added one. */
  drawingRid?: string;
  /** Relationship id of the VML part that carries the sheet's comment shapes. */
  legacyRid?: string;
}

function cellNode(cell: Cell, w: SheetWriter): any | undefined {
  const attrs: Record<string, string> = { "@_r": cell.ref };
  const style = w.sIndexOf(cell);
  if (style) attrs["@_s"] = String(style);
  if (cell.formula) {
    const cached = cacheOf(cell.constant);
    if (cached.attrs) Object.assign(attrs, cached.attrs);
    const children = [el("f", undefined, [textNode(cell.formula)])];
    if (cached.value !== undefined) children.push(el("v", undefined, [textNode(cached.value)]));
    return el("c", attrs, children);
  }
  const shared = w.indexOf(cell);
  if (shared !== undefined) return el("c", { ...attrs, "@_t": "s" }, [el("v", undefined, [textNode(String(shared))])]);
  const v = cell.constant;
  if (v === null || v === "") return style ? el("c", attrs) : undefined;
  if (typeof v === "number") return el("c", attrs, [el("v", undefined, [textNode(jsonNumber(v))])]);
  if (typeof v === "boolean") return el("c", { ...attrs, "@_t": "b" }, [el("v", undefined, [textNode(v ? "1" : "0")])]);
  if (isError(v)) return el("c", { ...attrs, "@_t": "e" }, [el("v", undefined, [textNode(v)])]);
  return undefined;
}

const jsonNumber = (n: number): string => (Number.isFinite(n) ? String(n) : "0");

function paneNode(sheet: Sheet): any {
  const f = sheet.freeze!;
  const active = f.rows && f.cols ? "bottomRight" : f.rows ? "bottomLeft" : "topRight";
  return el("pane", {
    ...(f.cols ? { "@_xSplit": String(f.cols) } : {}),
    ...(f.rows ? { "@_ySplit": String(f.rows) } : {}),
    "@_topLeftCell": refOf(f.cols, f.rows),
    "@_activePane": active,
    "@_state": "frozen",
  });
}

interface SourceCol {
  min: number;
  max: number;
  attrs: Record<string, string>;
}

/** Column entries carry outline levels and styles this model does not track; those survive a rebuild. */
function colsNode(sheet: Sheet, source: SourceCol[] = []): any | undefined {
  const entries: { col: number; width?: number; hidden?: boolean }[] = [];
  for (const [col, width] of sheet.colWidths) entries.push({ col, width });
  for (const col of sheet.hiddenCols) if (!sheet.colWidths.has(col)) entries.push({ col, hidden: true });
  if (!entries.length) return undefined;
  entries.sort((a, b) => a.col - b.col);
  return el(
    "cols",
    undefined,
    entries.map(({ col, width, hidden }) => {
      const attrs: Record<string, string> = { "@_min": String(col + 1), "@_max": String(col + 1) };
      const from = source.find((s) => col + 1 >= s.min && col + 1 <= s.max);
      if (from) {
        for (const [k, v] of Object.entries(from.attrs)) {
          if (!["@_min", "@_max", "@_width", "@_hidden", "@_customWidth"].includes(k)) attrs[k] = v;
        }
      }
      if (width !== undefined) {
        attrs["@_width"] = String(width);
        attrs["@_customWidth"] = "1";
      } else if (hidden) {
        attrs["@_hidden"] = "1";
        attrs["@_customWidth"] = "1";
      }
      return el("col", attrs);
    }),
  );
}

const applyRowHeight = (attrs: Record<string, string>, height: number | undefined): void => {
  if (height !== undefined) {
    attrs["@_ht"] = String(height);
    attrs["@_customHeight"] = "1";
  } else {
    delete attrs["@_ht"];
    delete attrs["@_customHeight"];
  }
};
const applyRowHidden = (attrs: Record<string, string>, sheet: Sheet, row: number): void => {
  const hidden = !!(sheet.hiddenRows?.has(row) || sheet.filteredRows?.has(row));
  const original = attrs["@_hidden"] === "1" || attrs["@_hidden"] === "true";
  if (hidden === original) return;
  if (hidden) attrs["@_hidden"] = "1"; else delete attrs["@_hidden"];
};

/**
 * Rows from the model. A row the source file already had keeps its element — and
 * with it the hidden/outline/spans attributes this model does not surface — with
 * its cell children replaced.
 */
function rowsOf(sheet: Sheet, w: SheetWriter, origin: Map<number, any> | undefined): any[] {
  const byRow = new Map<number, Cell[]>();
  for (const cell of sheet.cells.values()) {
    const list = byRow.get(cell.row) ?? [];
    list.push(cell);
    byRow.set(cell.row, list);
  }
  for (const row of [...sheet.rowHeights.keys(), ...(sheet.hiddenRows ?? []), ...(sheet.filteredRows ?? [])]) if (!byRow.has(row)) byRow.set(row, []);
  const out: any[] = [];
  for (const [rowNo, cells] of byRow) {
    const children = cells
      .sort((a, b) => a.col - b.col)
      .map((cell) => cellNode(cell, w))
      .filter(Boolean);
    const source = origin?.get(rowNo);
    if (!source) {
      const attrs: Record<string, string> = { "@_r": String(rowNo + 1) };
      applyRowHeight(attrs, sheet.rowHeights.get(rowNo));
      applyRowHidden(attrs, sheet, rowNo);
      out.push(el("row", attrs, children));
      continue;
    }
    const kids = rawChildren(source);
    const kept = kids.filter((child) => tagLocalOf(child) !== "c");
    kids.length = 0;
    kids.push(...children, ...kept);
    applyRowHeight(rawAttrs(source), sheet.rowHeights.get(rowNo));
    applyRowHidden(rawAttrs(source), sheet, rowNo);
    out.push(source);
  }
  if (origin) {
    for (const [rowNo, node] of origin) {
      if (byRow.has(rowNo)) continue;
      // A row the model emptied keeps its element only while its height matters;
      // empty rows that never carried state leave the file.
      if (rawChildren(node).some((child) => tagLocalOf(child) === "c")) {
        if (!sheet.rowHeights.has(rowNo)) continue;
        const kids = rawChildren(node);
        for (let i = kids.length - 1; i >= 0; i--) if (tagLocalOf(kids[i]) === "c") kids.splice(i, 1);
      }
      applyRowHeight(rawAttrs(node), sheet.rowHeights.get(rowNo));
      applyRowHidden(rawAttrs(node), sheet, rowNo);
      out.push(node);
    }
  }
  out.sort((a, b) => Number(rawAttrs(a)["@_r"] ?? 0) - Number(rawAttrs(b)["@_r"] ?? 0));
  return out;
}

/** CT_Worksheet child order; the writer weaves its elements back into the source's sequence. */
const WORKSHEET_ORDER = [
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
  "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
  "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
  "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags",
  "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems",
  "tableParts", "extLst",
];

/** CT_Stylesheet child order, for the sections a save carries over from the source. */
const STYLESHEET_ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];

const rankIn = (order: string[], tag: string): number => {
  const i = order.indexOf(tag);
  return i < 0 ? order.length + 1 : i;
};

/** Pretty-printed whitespace is not an element: it never bounds where a new one goes. */
const entryRank = (entry: any, order: string[]): number => {
  const tag = tagLocalOf(entry);
  return tag === undefined ? -1 : rankIn(order, tag);
};

/** Replaces an element in place, at the position the schema orders it. */
function place(kids: any[], tag: string, node: any | undefined, order: string[] = WORKSHEET_ORDER): void {
  for (let i = kids.length - 1; i >= 0; i--) if (tagLocalOf(kids[i]) === tag) kids.splice(i, 1);
  if (!node) return;
  const rank = rankIn(order, tag);
  const at = kids.findIndex((kid) => entryRank(kid, order) > rank);
  kids.splice(at < 0 ? kids.length : at, 0, node);
}

/**
 * Serializes a patched source tree. The parser keeps the source's `<?xml?>` as a tree
 * entry; dropping it here is what keeps the declaration from being written twice.
 */
const serializeTree = (tree: any[]): string => DECL + buildXml(tree.filter((entry) => tagOf(entry) !== "?xml"));

function childOf(kids: any[], tag: string): any | undefined {
  return kids.find((kid) => tagLocalOf(kid) === tag);
}

function sheetChildren(sheet: Sheet, w: SheetWriter, origin?: Map<number, any>): any[] {
  const kids: any[] = [];
  if (sheet.autoFilter?.columns.length) kids.push(el("sheetPr", { "@_filterMode": "1" }));
  if (sheet.freeze && (sheet.freeze.rows || sheet.freeze.cols)) {
    kids.push(el("sheetViews", undefined, [el("sheetView", { "@_workbookViewId": "0" }, [paneNode(sheet)])]));
  }
  const cols = colsNode(sheet);
  if (cols) kids.push(cols);
  kids.push(el("sheetData", undefined, rowsOf(sheet, w, origin)));
  if (sheet.autoFilter) kids.push(autoFilterNode(sheet.autoFilter));
  if (sheet.merges.length) {
    kids.push(el("mergeCells", { "@_count": String(sheet.merges.length) }, sheet.merges.map((m) => el("mergeCell", { "@_ref": m }))));
  }
  if (w.drawingRid) kids.push(el("drawing", { "@_r:id": w.drawingRid }));
  if (w.legacyRid) kids.push(el("legacyDrawing", { "@_r:id": w.legacyRid }));
  for (const cf of sheet.cond) kids.push(condNode(cf));
  return kids;
}

/** Replace owned criteria, keeping source root/column attributes and unknown siblings. */
function autoFilterNode(filter: AutoFilter, source?: any, worksheetPrefix = ""): any {
  const oldTag = source && tagOf(source), prefix = oldTag?.includes(":") ? oldTag.slice(0, oldTag.indexOf(":")) : worksheetPrefix;
  const t = (name: string, attrs?: Record<string, string>, kids?: any[]): any => el(prefix ? `${prefix}:${name}` : name, attrs, kids);
  const original = source ? rawChildren(source) : [];
  const oldColumns = original.filter((c) => tagLocalOf(c) === "filterColumn");
  const columns = filter.columns.map((col) => {
    const old = oldColumns.find((c) => Number(rawAttrs(c)["@_colId"]) === (col.sourceId ?? col.id));
    if (old && !col.edited) { rawAttrs(old)["@_colId"] = String(col.id); return old; }
    const rule = col.rule;
    const criteria = rule.kind === "values" ? t("filters", rule.blank ? { "@_blank": "1" } : undefined, rule.values.map((v) => t("filter", { "@_val": v })))
      : rule.kind === "custom" ? t("customFilters", rule.and ? { "@_and": "1" } : undefined, rule.tests.map((v) => t("customFilter", { "@_operator": v.operator, "@_val": v.value }))) : undefined;
    const extra = old ? rawChildren(old).filter((c) => tagLocalOf(c) === "extLst") : [];
    return t("filterColumn", { ...(old ? rawAttrs(old) : {}), "@_colId": String(col.id) }, [...(criteria ? [criteria] : []), ...extra]);
  });
  return t("autoFilter", { ...(source ? rawAttrs(source) : {}), "@_ref": filter.ref }, [...columns, ...original.filter((c) => tagLocalOf(c) !== "filterColumn")]);
}

/** One `conditionalFormatting` element; rules the model cannot express stay raw. */
function condNode(cf: CondFormat): any {
  return el("conditionalFormatting", { "@_sqref": cf.ref }, cf.rules.map(ruleNode));
}

function ruleNode(r: CondRule): any {
  // A rule from a source file rides its own XML, exactly as it was read.
  if (r.raw) {
    const tree = parseOrderedXml(r.raw);
    return tree.find((entry: any) => tagLocalOf(entry) === "cfRule") ?? tree[0];
  }
  const attrs: Record<string, string> = { "@_type": String(r.type), "@_priority": String(r.priority) };
  if (r.stopIfTrue) attrs["@_stopIfTrue"] = "1";
  const formulaTag = (text: string): any => el("formula", undefined, [textNode(text)]);
  if (r.type === "cellIs") {
    attrs["@_operator"] = r.op ?? "greaterThan";
    if (r.dxfId !== undefined) attrs["@_dxfId"] = String(r.dxfId);
    return el("cfRule", attrs, [formulaTag(String(r.formula ?? ""))]);
  }
  if (r.type === "expression") return el("cfRule", attrs, [formulaTag(String(r.formula ?? ""))]);
  if (r.type === "dataBar") {
    const body = el("dataBar", undefined, [el("cfvo", { "@_type": "min" }), el("cfvo", { "@_type": "max" }), el("color", { "@_rgb": `FF${r.colors?.[0] ?? "638EC6"}` })]);
    return el("cfRule", attrs, [body]);
  }
  if (r.type === "colorScale") {
    const colors = r.colors ?? [];
    const cfvo = (type: string, val?: string): any => el("cfvo", val ? { "@_type": type, "@_val": val } : { "@_type": type });
    const points = colors.length >= 3 ? [cfvo("min"), cfvo("percentile", "50"), cfvo("max")] : [cfvo("min"), cfvo("max")];
    const body = el("colorScale", undefined, [...points, ...colors.map((c) => el("color", { "@_rgb": `FF${c}` }))]);
    return el("cfRule", attrs, [body]);
  }
  return el("cfRule", attrs);
}

function sheetXml(sheet: Sheet, w: SheetWriter): string {
  return DECL + buildXml([el("worksheet", { "@_xmlns": NS, "@_xmlns:r": NS_R }, sheetChildren(sheet, w))]);
}

/**
 * Rewrites a source sheet in place: the elements the model owns are replaced and
 * everything else — conditional formats, margins, drawings, table references —
 * stays exactly where the source file had it.
 */
function patchSheetXml(source: string, sheet: Sheet, w: SheetWriter): string {
  const tree = parseOrderedXml(source);
  const root = rawFind(tree, "worksheet");
  const kids = root ? rawChildren(root) : undefined;
  if (!kids) return sheetXml(sheet, w);

  if (sheet.filterAtRead !== filterFingerprint(sheet.autoFilter)) {
    const full = tagOf(root!)!, prefix = full.includes(":") ? full.slice(0, full.indexOf(":")) : "";
    place(kids, "autoFilter", sheet.autoFilter ? autoFilterNode(sheet.autoFilter, childOf(kids, "autoFilter"), prefix) : undefined);
    let props = childOf(kids, "sheetPr");
    if (!props && sheet.autoFilter?.columns.length) { props = el(prefix ? `${prefix}:sheetPr` : "sheetPr"); place(kids, "sheetPr", props); }
    if (props) {
      if (sheet.autoFilter?.columns.length) rawAttrs(props)["@_filterMode"] = "1";
      else delete rawAttrs(props)["@_filterMode"];
    }
  }

  const frozen = !!(sheet.freeze && (sheet.freeze.rows || sheet.freeze.cols));
  if (rawFind(kids, "pane") !== undefined || frozen) {
    const views = childOf(kids, "sheetViews");
    const view = views ? rawFind(rawChildren(views), "sheetView") : undefined;
    if (!view) {
      place(kids, "sheetViews", frozen ? el("sheetViews", undefined, [el("sheetView", { "@_workbookViewId": "0" }, [paneNode(sheet)])]) : undefined);
    } else {
      // Only the pane is the model's; zoom, gridlines and selection stay.
      const viewKids = rawChildren(view);
      for (let i = viewKids.length - 1; i >= 0; i--) if (tagLocalOf(viewKids[i]) === "pane") viewKids.splice(i, 1);
      if (frozen) viewKids.unshift(paneNode(sheet));
    }
  }
  if (sheet.colsAtRead === undefined || sheet.colsAtRead !== colsFingerprint(sheet)) {
    const colsEl = childOf(kids, "cols");
    const source: SourceCol[] = colsEl
      ? rawChildren(colsEl)
          .filter((child) => tagLocalOf(child) === "col")
          .map((child) => ({ min: Number(rawAttrs(child)["@_min"] ?? 0), max: Number(rawAttrs(child)["@_max"] ?? 0), attrs: rawAttrs(child) }))
      : [];
    place(kids, "cols", colsNode(sheet, source));
  }
  if (sheet.condAtRead === undefined || condFingerprint(sheet) !== sheet.condAtRead) {
    // The model owns the rules now: the source's set is replaced by what it holds
    // (rules the model cannot evaluate ride their raw nodes back verbatim).
    for (let i = kids.length - 1; i >= 0; i--) if (tagLocalOf(kids[i]) === "conditionalFormatting") kids.splice(i, 1);
    if (sheet.cond.length) {
      // place() replaces same-tag elements, and a sheet carries one element per
      // sqref — so the group is inserted as a run at the schema position.
      const rank = rankIn(WORKSHEET_ORDER, "conditionalFormatting");
      let at = kids.findIndex((kid) => entryRank(kid, WORKSHEET_ORDER) > rank);
      if (at < 0) at = kids.length;
      for (const cf of sheet.cond) kids.splice(at++, 0, condNode(cf));
    }
  }

  const dataEl = childOf(kids, "sheetData");
  const origin = new Map<number, any>();
  for (const row of dataEl ? rawChildren(dataEl) : []) {
    if (tagLocalOf(row) !== "row") continue;
    const r = Number(rawAttrs(row)["@_r"]);
    if (Number.isFinite(r)) origin.set(r - 1, row);
  }
  place(kids, "sheetData", el("sheetData", undefined, rowsOf(sheet, w, origin)));
  place(
    kids,
    "mergeCells",
    sheet.merges.length
      ? el("mergeCells", { "@_count": String(sheet.merges.length) }, sheet.merges.map((m) => el("mergeCell", { "@_ref": m })))
      : undefined,
  );
  if (w.drawingRid) place(kids, "drawing", el("drawing", { "@_r:id": w.drawingRid }));
  if (w.legacyRid) place(kids, "legacyDrawing", el("legacyDrawing", { "@_r:id": w.legacyRid }));
  return serializeTree(tree);
}

// ------------------------------------------------------------- workbook parts

/** Relationship ids are re-derived for the parts a save owns; the rest keep theirs. */
function workbookParts(wb: Spreadsheet, hasSst: boolean): { workbook: string; rels: string } {
  const kept: { id: string; type: string; target: string; targetMode: string }[] = [];
  for (const [, r] of wb.source()?.rels("xl/workbook.xml") ?? []) {
    if (/(worksheet|styles|sharedStrings|calcChain)$/.test(r.type)) continue;
    kept.push(r);
  }
  const highest = kept.reduce((max, r) => Math.max(max, Number(/(\d+)$/.exec(r.id)?.[1] ?? 0)), 0);
  let next = highest + 1;
  const sheetIds = wb.sheets.map(() => `rId${next++}`);
  const stylesId = `rId${next++}`;
  const sstId = hasSst ? `rId${next++}` : undefined;

  const names = wb.definedNameTable().filter(([, text]) => !!text);
  const definedNames = names.length
    ? `<definedNames>${names.map(([name, text]) => `<definedName name="${XML_ESC(name)}">${XML_ESC(text)}</definedName>`).join("")}</definedNames>`
    : "";
  const workbook =
    `${DECL}<workbook xmlns="${NS}" xmlns:r="${NS_R}">` +
    `<sheets>${wb.sheets.map((s, i) => `<sheet name="${XML_ESC(s.name)}" sheetId="${i + 1}"${s.hidden ? ` state="${s.hidden}"` : ""} r:id="${sheetIds[i]}"/>`).join("")}</sheets>` +
    definedNames +
    `<bookViews><workbookView activeTab="${wb.activeSheet}"/></bookViews>` +
    // Cached results in the sheets may predate an edit; let the opening app recompute them.
    `<calcPr calcId="0" fullCalcOnLoad="1"/>` +
    `</workbook>`;

  const sheets = sheetIds.map((id, i) => `<Relationship Id="${id}" Type="${NS_R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("");
  const styles = `<Relationship Id="${stylesId}" Type="${NS_R}/styles" Target="styles.xml"/>`;
  const strings = sstId ? `<Relationship Id="${sstId}" Type="${NS_R}/sharedStrings" Target="sharedStrings.xml"/>` : "";
  const others = kept
    .map((r) => `<Relationship Id="${XML_ESC(r.id)}" Type="${XML_ESC(r.type)}" Target="${XML_ESC(r.target)}"${r.targetMode === "External" ? ' TargetMode="External"' : ""}/>`)
    .join("");
  return { workbook, rels: `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets}${styles}${strings}${others}</Relationships>` };
}

const packageRelsXml = (): string =>
  `${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

const packageRelsAreUsable = (bytes: Uint8Array | undefined): boolean =>
  bytes !== undefined && /relationships\/officeDocument"[^>]*Target="\/?xl\/workbook\.xml"/.test(bytesText(bytes));

const MEDIA_TYPES: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  svg: "image/svg+xml",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
};

/** Content types keep the source's declarations for every part that survives and restate the owned ones. */
function contentTypesXml(wb: Spreadsheet, files: Record<string, Uint8Array>, owned: Set<string>): string {
  const src = wb.source();
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  if (src) {
    for (const [ext, type] of src.defaultTypes) defaults.set(ext, type);
    for (const [part, type] of src.overrideTypes) {
      const path = part.replace(/^\/+/, "");
      if (files[path] && !owned.has(path)) overrides.set("/" + path, type);
    }
  }
  if (!defaults.has("rels")) defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml");
  if (!defaults.has("xml")) defaults.set("xml", "application/xml");
  for (const path of Object.keys(files)) {
    const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
    if (defaults.has(ext) || overrides.has("/" + path)) continue;
    if (ext === "vml") defaults.set(ext, "application/vnd.openxmlformats-officedocument.vmlDrawing");
    if (ext in MEDIA_TYPES) defaults.set(ext, MEDIA_TYPES[ext]);
  }
  overrides.set("/xl/workbook.xml", `${CT}.sheet.main+xml`);
  overrides.set("/xl/styles.xml", `${CT}.styles+xml`);
  if (files["xl/sharedStrings.xml"]) overrides.set("/xl/sharedStrings.xml", `${CT}.sharedStrings+xml`);
  for (const path of Object.keys(files)) {
    if (/^xl\/drawings\/drawing\d+\.xml$/.test(path)) overrides.set(`/${path}`, DRAWING_CT);
    if (/^xl\/charts\/chart\d+\.xml$/.test(path)) overrides.set(`/${path}`, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
    if (/^xl\/comments\d+\.xml$/.test(path)) overrides.set(`/${path}`, `${CT}.comments+xml`);
  }
  wb.sheets.forEach((_, i) => overrides.set(`/xl/worksheets/sheet${i + 1}.xml`, `${CT}.worksheet+xml`));
  const defs = [...defaults].map(([ext, type]) => `<Default Extension="${XML_ESC(ext)}" ContentType="${XML_ESC(type)}"/>`).join("");
  const ovs = [...overrides].map(([part, type]) => `<Override PartName="${XML_ESC(part)}" ContentType="${XML_ESC(type)}"/>`).join("");
  return `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defs}${ovs}</Types>`;
}

// ------------------------------------------------------------------- drawings

const DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMG_REL = `${NS_R}/image`;
const DRAW_REL = `${NS_R}/drawing`;
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

/**
 * Gives a sheet's model pictures a place in the package. A source drawing keeps
 * its shape and chart anchors; picture anchors are rebuilt from the model, so a
 * picture the editor deleted disappears and one it added appears. A sheet with
 * no usable source drawing gets a fresh part, a relationship and a `<drawing>`
 * element (returned as `sheetRid`).
 */
function planDrawing(files: Record<string, Uint8Array>, wb: Spreadsheet, sheet: Sheet, sheetPath: string): { sheetRid?: string } {
  const src = wb.source();
  const srcRels = src && sheet.srcPath ? src.rels(sheet.srcPath) : undefined;
  const existing = srcRels ? [...srcRels.values()].find((r) => r.type.endsWith("/drawing")) : undefined;
  if (!sheet.images.length && !sheet.charts.length && !existing) return {};
  // Chart content has its own owner: title/colors/data edits never require an anchor rewrite.
  for (const chart of sheet.charts) {
    if (chart.editUnsupported && chart.xml) { files[chart.part] = strToU8(chart.xml); continue; }
    const live = wb.chartForDisplay(sheet.name, chart);
    if (!chart.editUnsupported && (chart.xml || (chart.dataAtRead !== undefined && chartDataFingerprint(live.series) !== chart.dataAtRead))) files[chart.part] = strToU8(chartPartXml(live, sheet.name));
  }
  // An untouched drawing keeps its exact bytes — a chart or a form control in
  // it must survive an edited save without a serialization round trip.
  if (existing && sheet.imagesAtRead !== undefined && imagesFingerprint(sheet) === sheet.imagesAtRead) return {};
  if (existing) {
    const drawingPath = src!.resolve(sheet.srcPath!, existing);
    if (src!.parts.has(drawingPath)) {
      rewriteDrawing(files, src!, drawingPath, sheet);
      for (const img of sheet.images) files[img.path] = bytesOfDataUrl(img.src);

      return {};
    }
  }
  let n = 0;
  for (const p of Object.keys(files)) {
    const m = /^xl\/drawings\/drawing(\d+)\.xml$/.exec(p);
    if (m) n = Math.max(n, Number(m[1]));
  }
  const part = `xl/drawings/drawing${n + 1}.xml`;
  let ridNo = 0;
  let idNo = 1;
  const rels: string[] = [];
  const anchors: any[] = [];
  for (const img of sheet.images) {
    const rid = `rId${++ridNo}`;
    rels.push(`<Relationship Id="${rid}" Type="${IMG_REL}" Target="${XML_ESC(img.path.replace(/^xl\//, "../"))}"/>`);
    anchors.push(anchorNode(img, rid, idNo++));
  }
  for (const chart of sheet.charts) {
    const rid = `rId${++ridNo}`;
    rels.push(`<Relationship Id="${rid}" Type="${NS_R}/chart" Target="${XML_ESC(chart.part.replace(/^xl\//, "../"))}"/>`);
    anchors.push(chartFrameNode(chart, rid, idNo++));

  }
  for (const img of sheet.images) files[img.path] = bytesOfDataUrl(img.src);
  files[part] = strToU8(wsDrXml(anchors));
  files[relsPathOf(part)] = strToU8(`${DECL}<Relationships xmlns="${PKG_NS}">${rels.join("")}</Relationships>`);
  const relsPath = relsPathOf(sheetPath);
  // The sheet may already carry relationships (hyperlinks) from the source, even
  // under a renamed part; a fresh drawing relationship joins them, not replaces.
  const carried = files[relsPath] ?? (src && sheet.srcPath ? src.parts.get(relsPathOf(sheet.srcPath)) : undefined);
  const rid = `rId${maxRelId(carried ? decode(carried) : "") + 1}`;
  const tag = `<Relationship Id="${rid}" Type="${DRAW_REL}" Target="../drawings/drawing${n + 1}.xml"/>`;
  files[relsPath] = strToU8(relsWith(carried ? decode(carried) : "", tag));
  return { sheetRid: rid };
}

/** Replaces the picture anchors of one source drawing with what the model holds now. */
function rewriteDrawing(files: Record<string, Uint8Array>, src: OpcPackage, drawingPath: string, sheet: Sheet): void {
  const source = src.text(drawingPath);
  if (!source) return;
  const tree = parseOrderedXml(source);
  const root = rawFind(tree, "wsDr");
  if (!root) return;
  const full = tagOf(root) ?? "xdr:wsDr";
  const prefix = full.includes(":") ? full.slice(0, full.indexOf(":")) : "";
  const imagesByAnchor = new Map(sheet.images.filter((i) => i.drawingAnchorId).map((i) => [i.drawingAnchorId!, i]));
  const chartsByAnchor = new Map(sheet.charts.filter((c) => c.drawingAnchorId).map((c) => [c.drawingAnchorId!, c]));
  const kids = rawChildren(root);
  const seen = new Set<string>();
  const anchorIndexes = new Map<string, number>();
  const modeledAnchors = new Set(sheet.drawingAnchorsAtRead ?? []);
  for (const entry of [...kids]) {
    const local = tagLocalOf(entry);
    if (local !== "oneCellAnchor" && local !== "twoCellAnchor" && local !== "absoluteAnchor") continue;
    const anchorIndex = anchorIndexes.get(local) ?? 0;
    anchorIndexes.set(local, anchorIndex + 1);
    const anchorId = `${local}:${anchorIndex}`;
    if (!modeledAnchors.has(anchorId)) continue;
    const picture = rawChildren(entry).find((x) => tagLocalOf(x) === "pic");
    if (picture) {
      if (imagesByAnchor.has(anchorId)) seen.add(anchorId);
      else kids.splice(kids.indexOf(entry), 1);
      continue;
    }
    // A chart frame keeps its place while its part survives in the model.
    const graphicFrame = rawChildren(entry).find((x) => tagLocalOf(x) === "graphicFrame");
    const chartRef = graphicFrame ? rawFindAll(rawChildren(graphicFrame), "chart")[0] : undefined;
    const chartAttrs = chartRef ? rawAttrs(chartRef) : undefined;
    const chartId = chartAttrs ? (chartAttrs["@_r:id"] ?? chartAttrs["@_id"]) : undefined;
    if (chartId) {
      const chart = chartsByAnchor.get(anchorId);
      if (chart) {
        if (chart.geometryAtRead !== undefined && chartGeometryFingerprint(chart) !== chart.geometryAtRead) {
          const generated = chartFrameNode(chart, String(chartId), 0, prefix);
          const frame = rawChildren(entry).find((x) => tagLocalOf(x) === "graphicFrame");
          const oldTag = tagOf(entry)!;
          const kept = rawChildren(entry).filter((x) => !["from", "to", "ext", "pos"].includes(tagLocalOf(x) ?? ""));
          const geometry = rawChildren(generated).filter((x) => ["from", "ext"].includes(tagLocalOf(x) ?? ""));
          if (local === "twoCellAnchor") {
            // Keep the original move/resize-with-cells policy and editAs attribute.
            const box = imageBoxPx(sheet, chart), to = drawingAnchorAtPx(sheet, box.x + box.w, box.y + box.h);
            const t = (name: string, kids: any[]): any => el(prefix ? `${prefix}:${name}` : name, undefined, kids);
            const v = (name: string, value: number): any => t(name, [textNode(String(value))]);
            entry[oldTag] = [geometry[0], t("to", [v("col", to.col), v("colOff", to.colOff), v("row", to.row), v("rowOff", to.rowOff)]), ...kept];
          } else {
            entry[oldTag] = [...geometry, ...kept];
          }
          if (frame) {
            const xfrm = rawFind(rawChildren(frame), "xfrm");
            const ext = xfrm && rawFind(rawChildren(xfrm), "ext");
            if (ext) Object.assign(rawAttrs(ext), { "@_cx": String(chart.cx), "@_cy": String(chart.cy) });
          }
        }
        seen.add(anchorId);
        continue;
      }
      kids.splice(kids.indexOf(entry), 1);
    }
  }
  const fresh = sheet.images.filter((img) => !img.drawingAnchorId || !seen.has(img.drawingAnchorId));
  const freshCharts = sheet.charts.filter((c) => c.xml && (!c.drawingAnchorId || !seen.has(c.drawingAnchorId)));
  if (fresh.length || freshCharts.length) {
    const relsPath = relsPathOf(drawingPath);
    const current = files[relsPath] ?? src.parts.get(relsPath);
    let next = maxRelId(current ? decode(current) : "") + 1;
    let id = 1000;
    const add: string[] = [];
    for (const img of fresh) {
      const rid = `rId${next++}`;
      add.push(`<Relationship Id="${rid}" Type="${IMG_REL}" Target="${XML_ESC(img.path.replace(/^xl\//, "../"))}"/>`);
      kids.push(anchorNode(img, rid, id++, prefix));
    }
    for (const chart of freshCharts) {
      const rid = `rId${next++}`;
      add.push(`<Relationship Id="${rid}" Type="${NS_R}/chart" Target="${XML_ESC(chart.part.replace(/^xl\//, "../"))}"/>`);
      kids.push(chartFrameNode(chart, rid, id++, prefix));

    }
    files[relsPath] = strToU8(relsWith(current ? decode(current) : "", add.join("")));
  }
  const attrs = rawAttrs(root);
  if (!attrs["@_xmlns:a"]) attrs["@_xmlns:a"] = NS_A;
  if (!attrs["@_xmlns:r"]) attrs["@_xmlns:r"] = NS_R;
  files[drawingPath] = strToU8(serializeTree(tree));
}

/** Builds a oneCellAnchor ordered node; `prefix` matches the host file's own. */
function anchorNode(img: SheetImage, rid: string, id: number, prefix = "xdr"): any {
  const t = (name: string, kids: any[] = [], attrs?: Record<string, string>): any => el(prefix ? `${prefix}:${name}` : name, attrs, kids);
  const v = (name: string, value: string | number): any => t(name, [textNode(String(value))]);
  return t("oneCellAnchor", [
    t("from", [v("col", img.col), v("colOff", img.colOff), v("row", img.row), v("rowOff", img.rowOff)]),
    t("ext", [], { "@_cx": String(img.cx), "@_cy": String(img.cy) }),
    t("pic", [
      t("nvPicPr", [t("cNvPr", [], { "@_id": String(id), "@_name": img.name || "图片" }), t("cNvPicPr")]),
      t("blipFill", [t2("a:blip", { "@_r:embed": rid }), t2("a:stretch", undefined, [t2("a:fillRect")])]),
      t("spPr", [
        t2("a:xfrm", undefined, [t2("a:off", { "@_x": "0", "@_y": "0" }), t2("a:ext", { "@_cx": String(img.cx), "@_cy": String(img.cy) })]),
        t2("a:prstGeom", { "@_prst": "rect" }, [t2("a:avLst")]),
      ]),
    ]),
    t("clientData"),
  ]);
}

const t2 = (tag: string, attrs?: Record<string, string>, kids: any[] = []): any => el(tag, attrs, kids);

/** The graphicFrame anchor that places a chart on the sheet. */
function chartFrameNode(chart: ChartModel, rid: string, id: number, prefix = "xdr"): any {
  const t = (name: string, kids: any[] = [], attrs?: Record<string, string>): any => el(prefix ? `${prefix}:${name}` : name, attrs, kids);
  const v = (name: string, value: string | number): any => t(name, [textNode(String(value))]);
  return t("oneCellAnchor", [
    t("from", [v("col", chart.col), v("colOff", chart.colOff), v("row", chart.row), v("rowOff", chart.rowOff)]),
    t("ext", [], { "@_cx": String(chart.cx), "@_cy": String(chart.cy) }),
    t("graphicFrame", [
      t("nvGraphicFramePr", [t("cNvPr", [], { "@_id": String(id), "@_name": chart.title || "图表" }), t("cNvGraphicFramePr")]),
      t("xfrm", [t2("a:off", { "@_x": "0", "@_y": "0" }), t2("a:ext", { "@_cx": String(chart.cx), "@_cy": String(chart.cy) })]),
      t2("a:graphic", undefined, [
        t2("a:graphicData", { "@_uri": "http://schemas.openxmlformats.org/drawingml/2006/chart" }, [
          {
            "c:chart": [],
            ":@": { "@_xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart", "@_xmlns:r": NS_R, "@_r:id": rid },
          },
        ]),
      ]),
    ]),
    t("clientData"),
  ]);
}

const wsDrXml = (anchors: any[]): string =>
  `${DECL}<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">${anchors.map((a) => buildXml([a])).join("")}</xdr:wsDr>`;

const maxRelId = (rels: string): number => [...rels.matchAll(/Id="rId(\d+)"/g)].reduce((max, m) => Math.max(max, Number(m[1])), 0);

/** Appends relationship tags to a rels document, creating it when absent. */
function relsWith(rels: string, tags: string): string {
  if (!rels.trim()) return `${DECL}<Relationships xmlns="${PKG_NS}">${tags}</Relationships>`;
  return rels.replace(/<\/Relationships>\s*$/, `${tags}</Relationships>`);
}

// -------------------------------------------------------------------- comments

/** Writes the notes of one sheet: a comments part, VML shapes and both rels. */
function planComments(files: Record<string, Uint8Array>, wb: Spreadsheet, sheet: Sheet, sheetPath: string): { legacyRid?: string } {
  if (!sheet.comments.length) return {};
  const src = wb.source();
  let n = 0;
  for (const p of Object.keys(files)) {
    const m = /^xl\/comments(\d+)\.xml$/.exec(p);
    if (m) n = Math.max(n, Number(m[1]));
  }
  const commentsPath = `xl/comments${n + 1}.xml`;
  const vmlPath = `xl/drawings/vmlDrawing${n + 1}.vml`;
  files[commentsPath] = strToU8(commentsXml(sheet.comments));
  files[vmlPath] = strToU8(commentsVml(sheet.comments));
  const relsPath = relsPathOf(sheetPath);
  // A renamed sheet part's own relationships may not be in place yet; carry them.
  const carried = files[relsPath] ?? (src && sheet.srcPath ? src.parts.get(relsPathOf(sheet.srcPath)) : undefined);
  const base = maxRelId(carried ? decode(carried) : "");
  const tags =
    `<Relationship Id="rId${base + 1}" Type="${NS_R}/comments" Target="../comments${n + 1}.xml"/>` +
    `<Relationship Id="rId${base + 2}" Type="${NS_R}/vmlDrawing" Target="../drawings/vmlDrawing${n + 1}.vml"/>`;
  files[relsPath] = strToU8(relsWith(carried ? decode(carried) : "", tags));
  return { legacyRid: `rId${base + 2}` };
}

const commentsXml = (comments: { ref: string; text: string; author?: string }[]): string => {
  const authors = [...new Set(comments.map((c) => c.author || "批注"))];
  const list = comments
    .map((c) => {
      const authorId = Math.max(0, authors.indexOf(c.author || "批注"));
      return `<comment ref="${XML_ESC(c.ref)}" authorId="${authorId}"><text><r><t>${XML_ESC(c.text)}</t></r></text></comment>`;
    })
    .join("");
  return `${DECL}<comments xmlns="${NS}"><authors>${authors.map((a) => `<author>${XML_ESC(a)}</author>`).join("")}</authors><commentList>${list}</commentList></comments>`;
};

const commentsVml = (comments: { ref: string; text: string }[]): string => {
  const shapes = comments
    .map((c, i) => {
      const m = /([A-Z]+)([0-9]+)/.exec(c.ref);
      const col = m ? colToIndexSafe(m[1]) : 0;
      const row = m ? Number(m[2]) - 1 : 0;
      return (
        `<v:shape id="_x0000_s10${25 + i}" type="#_x0000_t202" style='position:absolute;margin-left:120pt;margin-top:60pt;width:150pt;height:80pt;z-index:${i + 1};visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto">` +
        `<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>` +
        `<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>` +
        `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1,15,${row},2,3,31,${row + 4}</x:Anchor>` +
        `<x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`
      );
    })
    .join("");
  return (
    `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>` +
    `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
    shapes +
    `</xml>`
  );
};

const colToIndexSafe = (letters: string): number => letters.split("").reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;

export function writeXlsx(wb: Spreadsheet): Uint8Array {
  const src = wb.source();
  const files: Record<string, Uint8Array> = {};
  if (src) for (const [path, bytes] of src.parts) files[path] = bytes;
  // A workbook nobody edited saves back byte-for-byte: the model covers cells and
  // styles, not themes, tables or drawings, and rebuilding those would drop them.
  if (src && wb.revision === 0) return zipSync(files, { level: 6 });
  wb.refreshFilters();

  const sheetPaths = wb.sheets.map((_, i) => `xl/worksheets/sheet${i + 1}.xml`);
  const owned = new Set<string>([
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    ...sheetPaths,
  ]);
  for (const path of Object.keys(files)) {
    // Directory entries (`xl/worksheets/_rels/`) are zip bookkeeping, not parts.
    if (path.endsWith("/")) continue;
    if (owned.has(path) || /^xl\/worksheets\/[^/]+\.xml$/.test(path)) delete files[path];
    else if (/^xl\/worksheets\/_rels\//.test(path) && !sheetPaths.some((p) => relsPathOf(p) === path)) delete files[path];
  }

  const styleTables = stylePlan(wb);
  const sst = sharedStringsXml(wb.sheets.flatMap((s) => [...s.cells.values()]));
  const hasSst = sst.count > 0;
  const base: SheetWriter = { indexOf: sst.indexOf, sIndexOf: styleTables.sIndexOf };
  wb.sheets.forEach((sheet, i) => {
    const plan = planDrawing(files, wb, sheet, sheetPaths[i]);
    const notes = planComments(files, wb, sheet, sheetPaths[i]);
    const w: SheetWriter = { ...base, drawingRid: plan.sheetRid, legacyRid: notes.legacyRid };
    const source = sheet.srcPath ? src?.text(sheet.srcPath) : undefined;
    files[sheetPaths[i]] = strToU8(source ? patchSheetXml(source, sheet, w) : sheetXml(sheet, w));
    // A renamed sheet part keeps its relationships, or the drawing it referenced
    // would lose its target.
    const relsAt = sheet.srcPath ? relsPathOf(sheet.srcPath) : undefined;
    const relsDst = relsPathOf(sheetPaths[i]);
    if (src && relsAt && relsAt !== relsDst && !files[relsDst] && src.parts.has(relsAt)) files[relsDst] = src.parts.get(relsAt)!;
  });

  files["xl/styles.xml"] = strToU8(styleTables.xml);
  if (hasSst) files["xl/sharedStrings.xml"] = strToU8(sst.xml);

  const plan = workbookParts(wb, hasSst);
  files["xl/workbook.xml"] = strToU8(plan.workbook);
  files["xl/_rels/workbook.xml.rels"] = strToU8(plan.rels);

  const sourceRels = src?.bytes("_rels/.rels");
  files["_rels/.rels"] = packageRelsAreUsable(sourceRels) ? sourceRels! : strToU8(packageRelsXml());

  files["[Content_Types].xml"] = strToU8(contentTypesXml(wb, files, owned));
  return zipSync(files, { level: 6 });
}
