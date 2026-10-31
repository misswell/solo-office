import { ATTRS, OpcPackage, asArray, decodeEntities, parseOrderedXml, rawAttrs, rawChildren, rawFind, tagLocalOf, text } from "./opc.js";
import { dataUrl } from "./media.js";
import { chartPartXml, rewriteChartSourceReferences } from "./chart.js";
import { chartRange, chartRef, chartSeriesForRange, chartDataFingerprint, chartSourceRange, shiftChartRef } from "./chart-data.js";
import { Evaluator, colToIndex, indexToCol, refToString, type CellSource, type Node, type Value } from "./formula.js";
import { BUILTIN, format } from "./numfmt.js";
import { eachCell, parseGrid, parseInput, parseRange, parseRef, rangeSpec, rect, refOf, shiftRefs, type Axis, type Rect } from "./sheet-ops.js";
import { filterFingerprint, filterMatches, filterRange, readAutoFilter, shiftAutoFilter, validateFilterRule, type AutoFilter, type FilterRule } from "./filter.js";

// ---- default geometry, shared by the renderer and the drawing anchors ----

/** One CSS pixel in EMU, the OOXML drawing unit. */
export const EMU_PX = 9525;
/** A column the file gives no width for behaves like Excel's default 8.43 characters. */
export const DEFAULT_COL_CHARS = 8.43;
/** A row with no stored height is 15 points, which is 20 CSS pixels. */
export const DEFAULT_ROW_PT = 15;

export const colWidthPx = (sheet: Sheet, col: number): number =>
  sheet.hiddenCols.has(col) ? 0 : Math.round((sheet.colWidths.get(col) ?? sheet.defaultColWidth ?? DEFAULT_COL_CHARS) * 7.2);

export const rowHeightPx = (sheet: Sheet, row: number): number =>
  sheet.hiddenRows?.has(row) || sheet.filteredRows?.has(row)
    ? 0
    : Math.round(((sheet.rowHeights.get(row) ?? sheet.defaultRowHeight ?? DEFAULT_ROW_PT) * 4) / 3);

function spanEmu(sheet: Sheet, line: number, axis: "col" | "row"): number {
  return Math.round(axis === "col" ? colWidthPx(sheet, line) : rowHeightPx(sheet, line)) * EMU_PX;
}

function cellOffsetEmu(sheet: Sheet, line: number, off: number, axis: "col" | "row"): number {
  let total = off;
  for (let i = 0; i < line; i++) total += spanEmu(sheet, i, axis);
  return total;
}

/** Where a sheet picture or chart sits in CSS pixels, so the renderer and the anchors agree. */
export function imageBoxPx(
  sheet: Sheet,
  box: Pick<SheetImage, "col" | "row" | "colOff" | "rowOff" | "cx" | "cy">,
): { x: number; y: number; w: number; h: number } {
  return {
    x: cellOffsetEmu(sheet, box.col, box.colOff, "col") / EMU_PX,
    y: cellOffsetEmu(sheet, box.row, box.rowOff, "row") / EMU_PX,
    w: box.cx / EMU_PX,
    h: box.cy / EMU_PX,
  };
}

/** Inverse of imageBoxPx's position, for a dragged object on a nonuniform grid. */
export function drawingAnchorAtPx(sheet: Sheet, x: number, y: number): { col: number; row: number; colOff: number; rowOff: number } {
  if (!Number.isFinite(x) || !Number.isFinite(y)) throw new Error("无效绘图位置");
  const locate = (px: number, axis: "col" | "row"): { line: number; off: number } => {
    let remaining = Math.max(0, px), line = 0;
    const max = axis === "col" ? 16383 : 1048575;
    while (line < max) {
      const size = axis === "col" ? colWidthPx(sheet, line) : rowHeightPx(sheet, line);
      if (size > 0 && remaining < size) break;
      remaining -= size;
      line++;
    }
    return { line, off: Math.round(remaining * EMU_PX) };
  };
  const col = locate(x, "col"), row = locate(y, "row");
  return { col: col.line, row: row.line, colOff: col.off, rowOff: row.off };
}

export const chartGeometryFingerprint = (chart: Pick<ChartModel, "col" | "row" | "colOff" | "rowOff" | "cx" | "cy">): string => JSON.stringify([chart.col, chart.row, chart.colOff, chart.rowOff, chart.cx, chart.cy]);

export interface Run {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  size?: number;
  color?: string;
}

export interface Cell {
  ref: string;
  col: number;
  row: number;
  /** Stored constant (already resolved away from shared-string indexes). */
  constant: Value;
  formula?: string;
  styleIndex: number;
  runs?: Run[];
}

export interface Sheet {
  name: string;
  cells: Map<string, Cell>;
  merges: string[];
  colWidths: Map<number, number>;
  hiddenCols: Set<number>;
  /** Manual row visibility and derived filter visibility have separate owners. */
  hiddenRows?: Set<number>;
  filteredRows?: Set<number>;
  autoFilter?: AutoFilter;
  filterAtRead?: string;
  rowHeights: Map<number, number>;
  /** `sheetFormatPr@defaultRowHeight` in points; absent means the OOXML default of 15. */
  defaultRowHeight?: number;
  /** `sheetFormatPr@defaultColWidth` in characters; absent means the OOXML default of 8.43. */
  defaultColWidth?: number;
  /** Floating pictures anchored to cells, as read or as the editor inserted them. */
  images: SheetImage[];
  /** Charts anchored to cells, as read or as the editor inserted them. */
  charts: ChartModel[];
  /** Conditional formats, grouped by their `sqref`. */
  cond: CondFormat[];
  /** Cell notes (批注), keyed by their cell ref in list order. */
  comments: SheetComment[];
  freeze?: { rows: number; cols: number };
  /** `hidden`/`veryHidden` as the source workbook had it; a save must not reveal it. */
  hidden?: "hidden" | "veryHidden";
  /** Part the sheet was read from; a save patches that XML instead of rebuilding it. */
  srcPath?: string;
  /** Column state as read; a save can keep the source `cols` element while it is unchanged. */
  colsAtRead?: string;
  /** Picture fingerprint as read; a save leaves the source drawing part alone while it holds. */
  imagesAtRead?: string;
  /** Only these source anchors were modeled; opaque sibling drawings must survive edits. */
  drawingAnchorsAtRead?: readonly string[];
  /** Conditional-format fingerprint as read; a save keeps the source rules while it holds. */
  condAtRead?: string;
}

/** A floating picture: anchor cell, offset inside it, display size, media part. */
export interface SheetImage {
  name: string;
  /** Identity of this source anchor, independent of shared media part paths. */
  drawingAnchorId?: string;
  /** Top-left anchor: cell plus EMU offsets measured inside that cell. */
  col: number;
  row: number;
  colOff: number;
  rowOff: number;
  /** Display size in EMU (9525 per CSS pixel). */
  cx: number;
  cy: number;
  /** Package path of the media bytes (`xl/media/imageN.ext`). */
  path: string;
  /** Data URL for the renderer; never written back to the package. */
  src: string;
}

/** The handful of dxf properties a conditional format can restyle. */
export interface CondStyle {
  fontColor?: string;
  bgColor?: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: boolean;
}

/** One conditional-formatting rule, as read or as the editor created it. */
export interface CondRule {
  type: "colorScale" | "dataBar" | "cellIs" | "expression" | (string & {});
  priority: number;
  /** cellIs only: greaterThan, lessThan, equal, notEqual, between… */
  op?: string;
  /** cellIs/expression operand or condition, stored without the leading `=`. */
  formula?: string;
  /** colorScale: 2–3 colors low→high; dataBar: exactly one. */
  colors?: string[];
  /** Index into the workbook dxf table (cellIs rules from a source file). */
  dxfId?: number;
  stopIfTrue?: boolean;
  /**
   * The rule exactly as read (its XML text), when it came from a source file —
   * icon sets, top10, text rules and other kinds the model cannot express travel
   * back verbatim instead of being dropped.
   */
  raw?: string;
}

/** A cell note: the text a reviewer left on one cell. */
export interface SheetComment {
  ref: string;
  text: string;
  author?: string;
}

export interface CondFormat {
  /** `sqref` — a range or space-separated list the rule applies to. */
  ref: string;
  rules: CondRule[];
}

export interface ChartSeries {
  name: string;
  nameRef?: string;
  catRef?: string;
  valRef?: string;
  catNumeric?: boolean;
  cats: string[];
  vals: (number | null)[];
}

/** One chart as read from a drawing, or as the editor inserted it. */
export interface ChartModel {
  /** Identity of this source anchor, independent of shared chart part paths. */
  drawingAnchorId?: string;
  kind: "bar" | "line" | "pie" | "area" | "scatter" | "doughnut" | "unknown";
  title: string;
  /** bar only: `col` (vertical) or `bar` (horizontal). */
  barDir?: "col" | "bar";
  grouping?: string;
  series: ChartSeries[];
  colors?: string[];
  showLegend?: boolean;
  sourceXml?: string;
  dataAtRead?: string;
  editUnsupported?: string;
  editFields?: ("title" | "type" | "source" | "colors" | "legend")[];
  geometryAtRead?: string;
  /** Anchor cell plus EMU offsets, same geometry as a picture. */
  col: number;
  row: number;
  colOff: number;
  rowOff: number;
  cx: number;
  cy: number;
  /** Package path of the chart part; read charts keep their exact bytes. */
  part: string;
  /** The chart XML the editor generated after an explicit chart edit. */
  xml?: string;
  /** Source range the data came from, for the references a new part carries. */
  srcRef?: string;
}

/** The width/hidden state a `cols` element carries, reduced to what the model can lose. */
export function colsFingerprint(sheet: Sheet): string {
  return JSON.stringify([
    [...sheet.colWidths.entries()].sort((a, b) => a[0] - b[0]),
    [...sheet.hiddenCols].sort((a, b) => a - b),
  ]);
}

/**
 * A theme colour slot's ink, with the `tint` Excel stores next to it applied. OOXML tints move a colour's
 * HSL lightness by that fraction (`tint<0` darker, `tint>0` lighter); the hue and saturation are kept.
 */
export function schemeColor(scheme: (string | undefined)[] | undefined, index: number, tint: number): string | undefined {
  const base = scheme?.[index];
  if (!base) return undefined;
  if (!tint) return base;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(base.slice(i, i + 2), 16) / 255);
  const max = Math.max(r!, g!, b!), min = Math.min(r!, g!, b!);
  let h = 0;
  const l = (max + min) / 2;
  const d = max - min;
  const s = d === 0 ? 0 : l > 0.5 ? d / (2 - max - min) : d / (max + min);
  if (d !== 0) {
    if (max === r) h = (g! - b!) / d + (g! < b! ? 6 : 0);
    else if (max === g) h = (b! - r!) / d + 2;
    else h = (r! - g!) / d + 4;
    h /= 6;
  }
  const l2 = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const channel = (p: number, q: number, t: number): number => {
    const tt = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };
  const q = l2 < 0.5 ? l2 * (1 + s) : l2 + s - l2 * s;
  const p = 2 * l2 - q;
  const out = s === 0 ? [l2, l2, l2] : [channel(p, q, h + 1 / 3), channel(p, q, h), channel(p, q, h - 1 / 3)];
  return out.map((v) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** One side of a cell's border: the file's `ST_BorderStyle` token and the ink it carries. */
export interface BorderLine {
  /** `thin`, `medium`, `double`, `dashed`, … — the token the writer puts back in `@_style`. */
  style: string;
  /** Six hex digits without `#`; absent means the app's default ink (a side that names no colour). */
  color?: string;
}

export interface Border {
  top?: BorderLine;
  right?: BorderLine;
  bottom?: BorderLine;
  left?: BorderLine;
}

export interface CellStyle {
  numFmt: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  fontSize: number;
  fontName?: string;
  /** Six hex digits without `#`, the way OOXML stores RGB. */
  fontColor?: string;
  bgColor?: string;
  align?: string;
  valign?: string;
  wrap: boolean;
  border?: Border;
}

export const DEFAULT_STYLE: CellStyle = { numFmt: "General", bold: false, italic: false, underline: false, fontSize: 11, wrap: false };

export interface WorkbookSnapshot {
  sheets: Sheet[];
  styles: CellStyle[];
  names: [string, Node][];
  activeSheet: number;
}

export class Spreadsheet implements CellSource {
  sheets: Sheet[] = [];
  private styles: CellStyle[] = [DEFAULT_STYLE];
  /** Differential formats from styles.xml, resolved to the properties CF can restyle. */
  private dxfs: CondStyle[] = [];
  private styleIds = new Map<string, number>();
  /** Source `cellXfs` index each model style came from; absent for styles the editor made. */
  private xfAlias: (number | undefined)[] = [];
  /** The source record's full meaning, to tell apart xfs the model reads as equal. */
  private xfShape: (string | undefined)[] = [];
  private definedNames = new Map<string, Node>();
  private readonly sharedStrings: { text: string; runs: Run[] }[] = [];
  private readonly sharedFormulas = new Map<string, { formula: string; col: number; row: number }>();
  private src?: OpcPackage;
  activeSheet = 0;
  private evaluators = new Map<string, Evaluator>();
  private rev = 0;
  private filtersRevision = -1;

  static open(pkg: OpcPackage): Spreadsheet {
    const wb = new Spreadsheet();
    wb.read(pkg);
    return wb;
  }

  static openBytes(data: Uint8Array): Spreadsheet {
    return Spreadsheet.open(OpcPackage.from(data));
  }

  /** A new one-sheet workbook with no source package; a save writes the canonical parts. */
  static blank(sheetName = "Sheet1"): Spreadsheet {
    const wb = new Spreadsheet();
    wb.sheets.push({ name: sheetName, cells: new Map(), merges: [], colWidths: new Map(), hiddenCols: new Set(), rowHeights: new Map(), images: [], charts: [], cond: [], comments: [] });
    return wb;
  }

  /** The package this workbook was read from; a save preserves the parts the model does not own. */
  source(): OpcPackage | undefined {
    return this.src;
  }

  // ------------------------------------------------------------------ reading

  private read(pkg: OpcPackage): void {
    this.src = pkg;
    const wbx = pkg.xml("xl/workbook.xml");
    const rels = pkg.rels("xl/workbook.xml");

    if (pkg.parts.has("xl/sharedStrings.xml")) {
      for (const si of asArray(pkg.xml("xl/sharedStrings.xml")?.sst?.si)) {
        this.sharedStrings.push(readSharedString(si));
      }
    }

    // xl/styles.xml roots as <styleSheet>, not <styles>.
    const st = pkg.parts.has("xl/styles.xml") ? pkg.xml("xl/styles.xml") : {};
    const custom = new Map<number, string>();
    for (const nf of asArray(st.styleSheet?.numFmts?.numFmt)) {
      custom.set(Number(nf["@_numFmtId"]), nf["@_formatCode"]);
    }
    // `@_rgb` is a literal; `@_theme` names a slot in the workbook's own clrScheme, with an optional tint.
    // Font and fill inks share this rule with borders, which is why the scheme is read once up here.
    const scheme = (() => {
      const path = [...pkg.parts.keys()].find((p) => /^xl\/theme\/[^/]+\.xml$/i.test(p));
      const root = path ? parseOrderedXml(pkg.text(path) ?? "") : undefined;
      const clr = root ? rawFind(root, "clrScheme") : undefined;
      if (!clr) return undefined;
      const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
      // Preserve-order entries keep attributes under `ATTRS`, not at the top level.
      const hex = (entry: any): string | undefined => {
        const attrs = entry?.[ATTRS];
        // `sysClr@val` is a system-colour name; its concrete ink is `lastClr`.
        const v = attrs?.["@_lastClr"] ?? attrs?.["@_val"];
        return typeof v === "string" && /^[0-9A-Fa-f]{6}$/.test(v) ? v.toUpperCase() : undefined;
      };
      const found = new Map<string, string>();
      for (const child of rawChildren(clr)) {
        const tag = tagLocalOf(child);
        const color = tag ? hex(rawChildren(child)[0]) : undefined;
        if (tag && color) found.set(tag, color);
      }
      return order.map((name) => found.get(name));
    })();
    const inkOf = (c: any): string | undefined => {
      const rgb = c?.["@_rgb"];
      if (typeof rgb === "string" && /^[0-9A-Fa-f]{6,8}$/.test(rgb)) return rgb.slice(-6).toUpperCase();
      if (c?.["@_theme"] !== undefined) return schemeColor(scheme, Number(c["@_theme"]), Number(c["@_tint"] ?? 0));
      return undefined;
    };
    const fonts = asArray(st.styleSheet?.fonts?.font).map((f: any) => ({
      bold: f.b !== undefined && f["@_val"] !== "0",
      italic: f.i !== undefined && f["@_val"] !== "0",
      underline: f.u !== undefined && f["@_val"] !== "none",
      size: f.sz ? Number(f.sz["@_val"]) : 11,
      name: f.name?.["@_val"] ? String(f.name["@_val"]) : undefined,
      color: inkOf(f.color),
    }));
    const fills = asArray(st.styleSheet?.fills?.fill).map((f: any) => {
      const pf = f.patternFill;
      const type = pf?.["@_patternType"];
      if (!pf || type === "none") return undefined;
      // Solid fills carry the colour as a child element: <fgColor rgb="FFFFF2CC"/> or <fgColor theme="4"/>.
      return inkOf(pf.fgColor) ?? inkOf(pf.bgColor);
    });
    const borders = asArray(st.styleSheet?.borders?.border);
    const rawFonts = asArray(st.styleSheet?.fonts?.font);
    const rawFills = asArray(st.styleSheet?.fills?.fill);
    // A side's own `@_style` is what draws it: `<left/>` and `style="none"` are both "no line" (the earlier
    // rule read an absent `style` as a drawn side, so a bottom-only rule also boxed the cell's other three).
    const borderLine = (bd: any, which: string): BorderLine | undefined => {
      const el = bd?.[which];
      const style = el?.["@_style"];
      if (!el || typeof style !== "string" || style === "none") return undefined;
      const color = inkOf(el.color);
      return color ? { style, color } : { style };
    };
    // The file's xf indices are not ours: identical records collapse into one entry.
    // Each entry remembers the record it came from, so a save can put a cell back
    // on the source `xf` instead of rebuilding a lookalike without its theme colours.
    const xfMap = asArray(st.styleSheet?.cellXfs?.xf).map((xf: any, xfIndex: number) => {      const fmtId = Number(xf["@_numFmtId"] ?? 0);
      const fontId = Number(xf["@_fontId"] ?? 0);
      const fillId = Number(xf["@_fillId"] ?? 0);
      const borderId = Number(xf["@_borderId"] ?? 0);
      const font = fonts[fontId];
      const al = xf.alignment;
      const bd = borders[borderId];
      const sides = ["left", "right", "top", "bottom"] as const;
      const lines = sides.map((side) => borderLine(bd, side));
      const border: Border | undefined =
        borderId > 0 && lines.some(Boolean)
          ? Object.fromEntries(sides.map((side, i) => [side, lines[i]])) as Border
          : undefined;
      const style: CellStyle = {
        numFmt: custom.get(fmtId) ?? BUILTIN[fmtId] ?? "General",
        bold: font?.bold ?? false,
        italic: font?.italic ?? false,
        underline: font?.underline ?? false,
        fontSize: font?.size ?? 11,
        fontName: font?.name,
        fontColor: font?.color,
        bgColor: fillId > 0 ? fills[fillId] : undefined,
        align: al?.["@_horizontal"],
        valign: al?.["@_vertical"],
        wrap: al?.["@_wrapText"] === "1",
        border,
      };
      // What the record means beyond what the model tracks: theme colours, tints,
      // scheme fonts, alignment details.
      const shape = JSON.stringify([rawFonts[fontId], rawFills[fillId], borders[borderId], fmtId, al ?? null]);
      if (xfIndex === 0) {
        // The workbook's own default cell style becomes the model's: cells with no
        // `s=` mean it, and so do cells the user creates in this worksheet.
        this.styles[0] = style;
        this.styleIds.set(this.keyOf(style), 0);
        this.xfAlias[0] = 0;
        this.xfShape[0] = shape;
        return 0;
      }
      const id = this.intern(style);
      if (this.xfAlias[id] === undefined) {
        this.xfAlias[id] = xfIndex;
        this.xfShape[id] = shape;
        return id;
      }
      if (this.xfShape[id] === shape) return id;
      // Two source records the model cannot tell apart: keep one entry per record
      // so cells that used either get their own back.
      const sep = this.styles.length;
      this.styles.push(style);
      this.xfAlias[sep] = xfIndex;
      this.xfShape[sep] = shape;
      return sep;
    });

    // Differential formats back the conditional-format rules a cell can carry.
    this.dxfs = asArray(st.styleSheet?.dxfs?.dxf).map((d: any): CondStyle => {
      const out: CondStyle = {};
      const f = d.font;
      if (f) {
        const color = f.color?.["@_rgb"] ?? f.color?.["@_val"];
        if (typeof color === "string" && color.length >= 6) out.fontColor = String(color).slice(-6).toUpperCase();
        if (f.b !== undefined) out.bold = true;
        if (f.i !== undefined) out.italic = true;
        if (f.strike !== undefined) out.strike = true;
        if (f.u !== undefined) out.underline = true;
      }
      const fill = d.fill?.patternFill;
      if (fill) {
        const rgb = fill.bgColor?.["@_rgb"] ?? fill.fgColor?.["@_rgb"] ?? fill.bgColor?.["@_val"] ?? fill.fgColor?.["@_val"];
        if (typeof rgb === "string" && rgb.length >= 6) out.bgColor = String(rgb).slice(-6).toUpperCase();
      }
      return out;
    });

    for (const dn of asArray(wbx.workbook?.definedNames?.definedName)) {
      const node = parseName(dn["#text"]);
      if (node) this.definedNames.set(String(dn["@_name"]).toUpperCase(), node);
    }

    for (const s of asArray(wbx.workbook?.sheets?.sheet)) {
      const rel = rels.get(s["@_id"]);
      const state = s["@_state"];
      const sheet: Sheet = {
        name: s["@_name"],
        cells: new Map(),
        merges: [],
        colWidths: new Map(),
        hiddenCols: new Set(),
        rowHeights: new Map(),
        images: [],
        charts: [],
        comments: [],
        cond: [],
        hidden: state === "hidden" || state === "veryHidden" ? state : undefined,
      };
      const path = rel ? pkg.resolve("xl/workbook.xml", rel) : `xl/worksheets/sheet${this.sheets.length + 1}.xml`;
      // Registered before its parts are read: chart references the drawing
      // carries resolve against this very sheet.
      this.sheets.push(sheet);
      if (pkg.parts.has(path)) {
        sheet.srcPath = path;
        this.readSheet(pkg, path, sheet, xfMap);
        this.readImages(pkg, path, sheet);
      }
    }
    const view = wbx.workbook?.views?.workbookView;
    if (view?.["@_activeTab"]) this.activeSheet = Number(view["@_activeTab"]);
    this.refreshFilters();
    // OOXML has one hidden flag. Rows rejected by supported criteria are filter-owned;
    // other hidden rows remain manual. Ambiguous manual+filtered overlap is not recoverable.
    for (const sheet of this.sheets) for (const row of sheet.filteredRows ?? []) sheet.hiddenRows?.delete(row);
  }

  private readSheet(pkg: OpcPackage, path: string, sheet: Sheet, xfMap: number[]): void {
    const x = pkg.xml(path);
    sheet.autoFilter = readAutoFilter(x.worksheet?.autoFilter);
    sheet.filterAtRead = filterFingerprint(sheet.autoFilter);
    sheet.hiddenRows = new Set();
    // The sheet's own default row height / column width; every sheet in the authorized corpus has one.
    const fmt = x.worksheet?.sheetFormatPr;
    if (fmt?.["@_defaultRowHeight"] !== undefined) sheet.defaultRowHeight = Number(fmt["@_defaultRowHeight"]);
    if (fmt?.["@_defaultColWidth"] !== undefined) sheet.defaultColWidth = Number(fmt["@_defaultColWidth"]);
    for (const c of asArray(x.worksheet?.cols?.col)) {
      const min = Number(c["@_min"]);
      const max = Number(c["@_max"]);
      const hidden = c["@_hidden"] === "1";
      for (let i = min; i <= max; i++) {
        if (hidden) sheet.hiddenCols.add(i - 1);
        else if (c["@_width"]) sheet.colWidths.set(i - 1, Number(c["@_width"]));
      }
    }
    sheet.colsAtRead = colsFingerprint(sheet);
    const pane = x.worksheet?.sheetViews?.sheetView?.pane;
    if (pane) {
      sheet.freeze = {
        rows: pane["@_ySplit"] ? Number(pane["@_ySplit"]) : 0,
        cols: pane["@_xSplit"] ? Number(pane["@_xSplit"]) : 0,
      };
    }
    for (const row of asArray(x.worksheet?.sheetData?.row)) {
      const rowIndex = Number(row["@_r"]) - 1;
      if (["1", "true", true].includes(row["@_hidden"])) sheet.hiddenRows.add(rowIndex);
      if (row["@_ht"]) sheet.rowHeights.set(rowIndex, Number(row["@_ht"]));
      for (const c of asArray(row.c)) {
        const ref = c["@_r"] ?? `${indexToCol(Number(c["@_i"] ?? 0))}${rowIndex + 1}`;
        const col = colToIndex(ref.replace(/[0-9]/g, ""));
        const type = c["@_t"];
        const styleIndex = Number(c["@_s"] ?? 0);
        let constant: Value = null;
        let runs: Run[] | undefined;
        const rawV = c.v?.["#text"] ?? (typeof c.v === "string" || typeof c.v === "number" ? String(c.v) : undefined);
        if (type === "s") {
          const entry = this.sharedStrings[Number(rawV)];
          constant = entry?.text ?? "";
          runs = entry?.runs;
        } else if (type === "inlineStr") {
          const r = readSharedString(c.is);
          constant = r.text;
          runs = r.runs;
        } else if (type === "b") constant = rawV === "1";
        else if (type === "e") constant = (rawV ?? "#VALUE!") as Value;
        else if (type === "str") constant = rawV === undefined ? "" : decodeEntities(String(rawV));
        else if (rawV !== undefined && rawV !== "") constant = Number(rawV);

        let formula: string | undefined;
        const f = c.f;
        const fText = typeof f === "string" ? f : typeof f?.["#text"] === "string" ? f["#text"] : undefined;
        const shared = f?.["@_t"] === "shared" ? String(f["@_si"]) : undefined;
        if (fText) {
          formula = fText;
          if (shared !== undefined) this.sharedFormulas.set(shared, { formula, col, row: rowIndex });
        } else if (shared !== undefined) {
          const master = this.sharedFormulas.get(shared);
          if (master) formula = translate(master.formula, master.col, master.row, col, rowIndex);
        }
        sheet.cells.set(ref, { ref, col, row: rowIndex, constant, formula, styleIndex: xfMap[styleIndex] ?? 0, runs });
      }
    }
    for (const mc of asArray(x.worksheet?.mergeCells?.mergeCell)) sheet.merges.push(mc["@_ref"]);
    // The cfRule elements in raw text, matched to parsed rules by document order.
    const ruleTags = [...String(pkg.text(path) ?? "").matchAll(/<cfRule\b[\s\S]*?<\/cfRule>|<cfRule\b[^>]*\/>/g)].map((m) => m[0]);
    for (const cf of asArray(x.worksheet?.conditionalFormatting)) {
      const ref = cf["@_sqref"];
      if (!ref) continue;
      const rules: CondRule[] = [];
      for (const r of asArray(cf.cfRule)) {
        const type = String(r["@_type"] ?? "");
        const rule: CondRule = { type: type as CondRule["type"], priority: Number(r["@_priority"] ?? 1) || 1 };
        if (r["@_operator"]) rule.op = String(r["@_operator"]);
        if (r["@_dxfId"] !== undefined && r["@_dxfId"] !== "") rule.dxfId = Number(r["@_dxfId"]);
        if (r["@_stopIfTrue"] === "1") rule.stopIfTrue = true;
        if (type === "colorScale") {
          const colors = asArray(r.colorScale?.color)
            .map((c: any) => String(c["@_rgb"] ?? "").slice(-6).toUpperCase())
            .filter((c: string) => /^[0-9A-F]{6}$/i.test(c));
          if (colors.length >= 2) rule.colors = colors;
        } else if (type === "dataBar") {
          const c = r.dataBar?.color?.["@_rgb"] ?? r.dataBar?.color?.["@_val"];
          if (typeof c === "string" && c.length >= 6) rule.colors = [String(c).slice(-6).toUpperCase()];
        } else if (type === "cellIs" || type === "expression") {
          const t = r.formula?.["#text"];
          if (t !== undefined && t !== null && String(t) !== "") rule.formula = String(t);
        }
        // Everything — including kinds the model cannot evaluate — rides along.
        if (ruleTags.length) rule.raw = ruleTags.shift();
        rules.push(rule);
      }
      if (rules.length) sheet.cond.push({ ref: String(ref), rules });
    }
    for (const comRel of [...pkg.rels(path).values()].filter((r) => r.type.endsWith("/comments"))) {
      const comPath = pkg.resolve(path, comRel);
      if (!pkg.parts.has(comPath)) continue;
      const root = pkg.xml(comPath)?.comments;
      const authors = asArray(root?.authors?.author).map((a: any) => (typeof a === "string" ? a : String(a?.["#text"] ?? "")));
      for (const c of asArray(root?.commentList?.comment)) {
        const ref = String(c["@_ref"] ?? "");
        if (!ref) continue;
        const authorId = Number(c["@_authorId"] ?? 0) || 0;
        let text = "";
        for (const t of asArray(c.text?.r?.t)) text += typeof t === "string" ? t : String(t?.["#text"] ?? "");
        if (!text && c.text) {
          const walk = (o: any): void => {
            if (!o || typeof o !== "object") return;
            for (const [k, v] of Object.entries(o)) {
              if (k === "t") text += typeof v === "string" ? v : String((v as { "#text"?: unknown })?.["#text"] ?? "");
              else if (v && typeof v === "object") walk(v);
            }
          };
          walk(c.text);
        }
        sheet.comments.push({ ref, text: decodeEntities(text), author: authors[authorId] });
      }
    }
    sheet.condAtRead = condFingerprint(sheet);
  }

  /**
   * Floating pictures of one sheet, from its drawing part. Anchors are normalized
   * to oneCellAnchor form: a twoCellAnchor keeps its top-left cell and derives its
   * extent from the column widths and row heights the sheet itself declares.
   */
  private readImages(pkg: OpcPackage, sheetPath: string, sheet: Sheet): void {
    const rel = [...pkg.rels(sheetPath).values()].find((r) => r.type.endsWith("/drawing"));
    if (!rel) return;
    const drawingPath = pkg.resolve(sheetPath, rel);
    if (!pkg.parts.has(drawingPath)) return;
    const drawRels = pkg.rels(drawingPath);
    const root = pkg.xml(drawingPath)?.wsDr;
    if (!root) return;
    const cellOf = (which: any): { col: number; row: number; colOff: number; rowOff: number } | undefined => {
      if (!which) return undefined;
      const num = (v: any): number => Number(v?.["#text"] ?? v ?? 0) || 0;
      return { col: num(which.col), row: num(which.row), colOff: num(which.colOff), rowOff: num(which.rowOff) };
    };
    const extentOf = (anchor: any, key: string, from: { col: number; row: number; colOff: number; rowOff: number }): { cx: number; cy: number } => {
      let cx = Number(anchor.ext?.["@_cx"] ?? 0);
      let cy = Number(anchor.ext?.["@_cy"] ?? 0);
      if (key === "twoCellAnchor") {
        const to = cellOf(anchor.to);
        if (to) {
          cx = Math.max(EMU_PX, cellOffsetEmu(sheet, to.col, to.colOff, "col") - cellOffsetEmu(sheet, from.col, from.colOff, "col"));
          cy = Math.max(EMU_PX, cellOffsetEmu(sheet, to.row, to.rowOff, "row") - cellOffsetEmu(sheet, from.row, from.rowOff, "row"));
        }
      }
      return { cx, cy };
    };
    const modeledAnchors: string[] = [];
    for (const key of ["oneCellAnchor", "twoCellAnchor"] as const) {
      for (const [anchorIndex, anchor] of asArray(root[key]).entries()) {
        const from = cellOf(anchor.from);
        if (!from) continue;
        const { cx, cy } = extentOf(anchor, key, from);
        const frame = anchor.graphicFrame;
        if (frame) {
          const chartRef = frame.graphic?.graphicData?.chart?.["@_r:id"] ?? frame.graphic?.graphicData?.chart?.["@_id"];
          const chartRel = chartRef ? drawRels.get(String(chartRef)) : undefined;
          const chartPath = chartRel ? pkg.resolve(drawingPath, chartRel) : undefined;
          if (!chartPath || !pkg.parts.has(chartPath)) continue;
          const chart = parseChart(pkg, chartPath, sheet.name, (f) => this.resolveFormulaRange(f, sheet.name));
          if (!chart || !cx || !cy) continue;
          chart.col = from.col;
          chart.row = from.row;
          chart.colOff = from.colOff;
          chart.rowOff = from.rowOff;
          chart.cx = cx;
          chart.cy = cy;
          chart.geometryAtRead = chartGeometryFingerprint(chart);
          chart.drawingAnchorId = `${key}:${anchorIndex}`;
          sheet.charts.push(chart);
          modeledAnchors.push(`${key}:${anchorIndex}`);
          continue;
        }
        const pic = anchor.pic;
        const embed = pic?.blipFill?.blip?.["@_embed"] ?? pic?.blipFill?.blip?.["@_r:embed"];
        const mediaRel = embed ? drawRels.get(String(embed)) : undefined;
        const mediaPath = mediaRel ? pkg.resolve(drawingPath, mediaRel) : undefined;
        if (!pic || !mediaPath || !pkg.parts.has(mediaPath)) continue;
        if (!cx || !cy) continue;
        const mime = mediaPath.endsWith(".png") ? "image/png" : mediaPath.endsWith(".gif") ? "image/gif" : mediaPath.endsWith(".svg") ? "image/svg+xml" : "image/jpeg";
        sheet.images.push({
          drawingAnchorId: `${key}:${anchorIndex}`,
          name: pic.nvPicPr?.cNvPr?.["@_name"] ?? "图片",
          col: from.col,
          row: from.row,
          colOff: from.colOff,
          rowOff: from.rowOff,
          cx,
          cy,
          path: mediaPath,
          src: dataUrl(mime, pkg.parts.get(mediaPath)!),
        });
        modeledAnchors.push(`${key}:${anchorIndex}`);
      }
    }
    sheet.imagesAtRead = imagesFingerprint(sheet);
    sheet.drawingAnchorsAtRead = modeledAnchors;
  }
  // --------------------------------------------------------- CellSource API

  constant(sheetName: string, col: number, row: number): Value {
    return this.sheet(sheetName)?.cells.get(refToString(col, row))?.constant ?? null;
  }

  formula(sheetName: string, col: number, row: number): string | undefined {
    return this.sheet(sheetName)?.cells.get(refToString(col, row))?.formula;
  }

  name(scope: string): Node | undefined {
    return this.definedNames.get(scope.toUpperCase());
  }

  sheet(nameOrIndex: string | number): Sheet | undefined {
    return typeof nameOrIndex === "number"
      ? this.sheets[nameOrIndex]
      : this.sheets.find((s) => s.name.toLowerCase() === String(nameOrIndex).toLowerCase());
  }

  /** Computed value: formulas evaluated on demand, constants passed through. */
  value(nameOrIndex: string | number, ref: string): Value {
    const s = this.sheet(nameOrIndex);
    if (!s) return null;
    const cell = s.cells.get(ref);
    if (!cell) return null;
    if (!cell.formula) return cell.constant;
    return this.currentEvaluator().cellValue(s.name, cell.col, cell.row);
  }

  private currentEvaluator(): Evaluator {
    const key = this.sheets[this.activeSheet]?.name ?? "";
    let ev = this.evaluators.get(key);
    if (!ev) {
      ev = new Evaluator(this, this.sheets[this.activeSheet]?.name ?? "Sheet1");
      this.evaluators.set(key, ev);
    }
    return ev;
  }

  display(nameOrIndex: string | number, ref: string): { text: string; colored: boolean } {
    const s = this.sheet(nameOrIndex);
    const cell = s?.cells.get(ref);
    if (!s || !cell) return { text: "", colored: false };
    const v = this.value(s.name, ref);
    const style = this.styles[cell.styleIndex] ?? DEFAULT_STYLE;
    if (typeof v === "number") return { text: format(v, style.numFmt ?? "General"), colored: false };
    if (typeof v === "boolean") return { text: v ? "TRUE" : "FALSE", colored: false };
    return { text: v === null ? "" : String(v), colored: false };
  }

  /** The formula bar shows the expression, not its result. */
  editOf(nameOrIndex: string | number, ref: string): string {
    const cell = this.sheet(nameOrIndex)?.cells.get(ref);
    if (!cell) return "";
    if (cell.formula) return "=" + cell.formula;
    return this.display(nameOrIndex, ref).text;
  }

  styleOf(cell: Cell): CellStyle {
    return this.styles[cell.styleIndex] ?? DEFAULT_STYLE;
  }

  styleAt(nameOrIndex: string | number, c: number, r: number): CellStyle {
    const cell = this.sheet(nameOrIndex)?.cells.get(refOf(c, r));
    return cell ? this.styleOf(cell) : DEFAULT_STYLE;
  }

  bounds(nameOrIndex: string | number): { maxRow: number; maxCol: number } {
    const s = this.sheet(nameOrIndex);
    let maxRow = 0;
    let maxCol = 0;
    for (const c of s?.cells.values() ?? []) {
      maxRow = Math.max(maxRow, c.row);
      maxCol = Math.max(maxCol, c.col);
    }
    for (const m of s?.merges ?? []) {
      const box = parseRange(m);
      maxRow = Math.max(maxRow, box.r2);
      maxCol = Math.max(maxCol, box.c2);
    }
    return { maxRow, maxCol };
  }

  // ------------------------------------------------------------------ editing

  /** Write a literal or formula and drop cached results so the next read recomputes. */
  setCell(nameOrIndex: string | number, ref: string, input: string): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    const at = parseRef(ref);
    const cell = this.touch(s, at);
    const parsed = parseInput(input);
    cell.constant = parsed.constant;
    cell.formula = parsed.formula;
    cell.runs = undefined;
    this.fresh();
  }

  /** Put clipboard text down with the anchor as its top-left cell. */
  pasteText(nameOrIndex: string | number, c: number, r: number, text: string): Rect | undefined {
    const s = this.sheet(nameOrIndex);
    if (!s) return undefined;
    const grid = parseGrid(text);
    if (!grid.length) return undefined;
    for (const [ri, row] of grid.entries()) {
      for (const [ci, field] of row.entries()) {
        const cell = this.touch(s, { c: c + ci, r: r + ri });
        const parsed = parseInput(field);
        cell.constant = parsed.constant;
        cell.formula = parsed.formula;
        cell.runs = undefined;
      }
    }
    this.fresh();
    return rect(c, r, c + grid[0].length - 1, r + grid.length - 1);
  }

  /**
   * Fill-handle semantics: project the source range into the target band. Text
   * and booleans copy; numbers continue their series; formulas shift their
   * relative references, exactly like a shared formula does.
   */
  fill(nameOrIndex: string | number, from: Rect, to: Rect): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    const growDown = to.r2 - from.r2;
    const growUp = from.r1 - to.r1;
    const growRight = to.c2 - from.c2;
    const growLeft = from.c1 - to.c1;
    if (growDown <= 0 && growUp <= 0 && growRight <= 0 && growLeft <= 0) return;
    const vertical = Math.max(growDown, growUp) >= Math.max(growRight, growLeft);
    const height = from.r2 - from.r1 + 1;
    const width = from.c2 - from.c1 + 1;
    for (let r = to.r1; r <= to.r2; r++) {
      for (let c = to.c1; c <= to.c2; c++) {
        if (c >= from.c1 && c <= from.c2 && r >= from.r1 && r <= from.r2) continue;
        const sc = vertical ? c : from.c1 + (((c - from.c1) % width) + width) % width;
        const sr = vertical ? from.r1 + (((r - from.r1) % height) + height) % height : r;
        const src = s.cells.get(refOf(sc, sr));
        const ref = refOf(c, r);
        if (!src || (src.constant === null && src.formula === undefined)) {
          s.cells.delete(ref);
          continue;
        }
        const cell = this.touch(s, { c, r });
        cell.runs = src.runs?.map((x) => ({ ...x }));
        cell.styleIndex = src.styleIndex;
        if (src.formula) {
          cell.formula = shiftFormulaRefs(src.formula, c - sc, r - sr);
          cell.constant = src.constant;
        } else if (typeof src.constant === "number") {
          cell.formula = undefined;
          cell.constant = this.seriesValue(s, vertical, sc, sr, height, width, growDown > 0 || growRight > 0, vertical ? r : c, from);
        } else {
          cell.formula = undefined;
          cell.constant = src.constant;
        }
      }
    }
    this.fresh();
  }

  /** One step past the edge of a numeric series, in the direction the handle moved. */
  private seriesValue(s: Sheet, vertical: boolean, sc: number, sr: number, height: number, width: number, after: boolean, line: number, from: Rect): number {
    const at = (i: number): number | undefined => {
      const cell = s.cells.get(refOf(vertical ? sc : from.c1 + i, vertical ? from.r1 + i : sr));
      return cell && typeof cell.constant === "number" ? cell.constant : undefined;
    };
    const count = vertical ? height : width;
    const last = at(count - 1);
    const first = at(0);
    if (last === undefined || first === undefined) return 0;
    const slope = count >= 2 && at(count - 2) !== undefined ? last - at(count - 2)! : 1;
    const distance = after ? line - (vertical ? from.r2 : from.c2) : (vertical ? from.r1 : from.c1) - line;
    return after ? last + slope * distance : first - slope * distance;
  }

  setStyle(nameOrIndex: string | number, box: Rect, patch: Partial<CellStyle>): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    for (const at of eachCell(box)) {
      const cell = this.touch(s, at);
      cell.styleIndex = this.intern({ ...this.styleOf(cell), ...patch });
    }
    this.fresh();
  }

  clear(nameOrIndex: string | number, box: Rect, what: "content" | "format" | "all" = "content"): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    for (const at of eachCell(box)) {
      const ref = refOf(at.c, at.r);
      const cell = s.cells.get(ref);
      if (!cell) continue;
      if (what !== "format") {
        cell.constant = null;
        cell.formula = undefined;
        cell.runs = undefined;
      }
      if (what !== "content") cell.styleIndex = 0;
      if (cell.constant === null && !cell.formula && cell.styleIndex === 0) s.cells.delete(ref);
    }
    this.fresh();
  }

  private touch(s: Sheet, at: { c: number; r: number }): Cell {
    const ref = refOf(at.c, at.r);
    let cell = s.cells.get(ref);
    if (!cell) {
      cell = { ref, col: at.c, row: at.r, constant: null, styleIndex: 0 };
      s.cells.set(ref, cell);
    }
    return cell;
  }

  /** Bumps on every mutation, so undo can tell a recorded state from a changed one. */
  get revision(): number {
    return this.rev;
  }

  private fresh(): void {
    this.rev++;
    this.evaluators = new Map();
  }

  refreshFilters(): void {
    if (this.filtersRevision === this.rev) return;
    for (const s of this.sheets) {
      const hidden = new Set<number>(), filter = s.autoFilter;
      if (filter && !filter.unsupported && filter.columns.length) {
        const box = filterRange(filter.ref);
        for (let row = box.r1 + 1; row <= box.r2; row++) {
          if (!filter.columns.every((col) => filterMatches(this.value(s.name, refOf(box.c1 + col.id, row)), col.rule))) hidden.add(row);
        }
      }
      s.filteredRows = hidden;
    }
    this.filtersRevision = this.rev;
  }

  enableAutoFilter(nameOrIndex: string | number, spec: string): void {
    const s = this.sheet(nameOrIndex);
    if (!s) throw new Error("未找到工作表");
    const ref = rangeSpec(filterRange(spec));
    if (s.autoFilter?.ref === ref) return;
    s.autoFilter = { ref, columns: [] };
    this.fresh();
  }

  setFilterColumn(nameOrIndex: string | number, id: number, rule?: FilterRule): void {
    const s = this.sheet(nameOrIndex), filter = s?.autoFilter;
    if (!s || !filter) throw new Error("请先启用筛选");
    if (filter.unsupported) throw new Error(filter.unsupported);
    const box = filterRange(filter.ref);
    if (!Number.isInteger(id) || id < 0 || id > box.c2 - box.c1) throw new Error("无效筛选列");
    if (rule) validateFilterRule(rule);
    const old = filter.columns.find((c) => c.id === id), next = filter.columns.filter((c) => c.id !== id);
    if (rule) next.push({ id, sourceId: old?.sourceId, edited: true, rule: structuredClone(rule) });
    next.sort((a, b) => a.id - b.id);
    const draft = { ...filter, columns: next };
    if (filterFingerprint(draft) === filterFingerprint(filter)) return;
    s.autoFilter = draft;
    this.fresh();
  }

  clearFilterCriteria(nameOrIndex: string | number): void {
    const s = this.sheet(nameOrIndex);
    if (!s?.autoFilter || !s.autoFilter.columns.length) return;
    s.autoFilter = { ref: s.autoFilter.ref, columns: [] };
    this.fresh();
  }

  removeAutoFilter(nameOrIndex: string | number): void {
    const s = this.sheet(nameOrIndex);
    if (!s?.autoFilter) return;
    s.autoFilter = undefined;
    this.fresh();
  }

  hideRow(nameOrIndex: string | number, row: number, hidden: boolean): void {
    const s = this.sheet(nameOrIndex);
    if (!s || !Number.isInteger(row) || row < 0 || row > 1048575) return;
    s.hiddenRows ??= new Set();
    if (s.hiddenRows.has(row) === hidden) return;
    if (hidden) s.hiddenRows.add(row); else s.hiddenRows.delete(row);
    this.fresh();
  }

  private keyOf(style: CellStyle): string {
    return JSON.stringify([
      style.numFmt,
      style.bold,
      style.italic,
      style.underline,
      style.fontSize,
      style.fontName ?? "",
      style.fontColor ?? "",
      style.bgColor ?? "",
      style.align ?? "",
      style.valign ?? "",
      style.wrap,
      style.border ? (["left", "right", "top", "bottom"] as const).map((side) => {
        const line = style.border![side];
        return line ? [line.style, line.color ?? ""] : null;
      }) : null,
    ]);
  }

  private intern(style: CellStyle): number {
    const key = this.keyOf(style);
    let id = this.styleIds.get(key);
    if (id === undefined) {
      id = this.styles.length;
      this.styles.push(style);
      this.styleIds.set(key, id);
    }
    return id;
  }

  /** Style records in the exact order a writer must emit them as `cellXfs`. */
  styleTable(): CellStyle[] {
    return this.styles;
  }

  /** Source `cellXfs` index for each model style; a save keeps those records verbatim. */
  styleAliases(): (number | undefined)[] {
    return this.xfAlias;
  }

  definedNameTable(): [string, string][] {
    const out: [string, string][] = [];
    for (const [name, node] of this.definedNames) {
      const text = nodeText(node);
      if (text !== undefined) out.push([name, text]);
    }
    return out;
  }

  // -------------------------------------------------------- structure changes

  insertRows(nameOrIndex: string | number, at: number, count = 1): void {
    this.splice(nameOrIndex, "row", at, -count);
  }

  deleteRows(nameOrIndex: string | number, at: number, count = 1): void {
    this.splice(nameOrIndex, "row", at, count);
  }

  insertCols(nameOrIndex: string | number, at: number, count = 1): void {
    this.splice(nameOrIndex, "col", at, -count);
  }

  deleteCols(nameOrIndex: string | number, at: number, count = 1): void {
    this.splice(nameOrIndex, "col", at, count);
  }

  /**
   * Insert (`count < 0`) or delete (`count > 0`) lines on one sheet. Only that
   * sheet's cells, merges and sizes move, but formulas on *every* sheet are
   * rewritten: a `汇总!B2` written elsewhere still points at the moved cells,
   * while a reference to a third sheet must stay put.
   */
  private splice(nameOrIndex: string | number, axis: Axis, at: number, count: number): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    const drop = count > 0 ? count : 0;
    const delta = -count;
    const next = new Map<string, Cell>();
    for (const cell of s.cells.values()) {
      const line = axis === "row" ? cell.row : cell.col;
      if (drop && line >= at && line < at + drop) continue;
      const moved = line >= at ? line + delta : line;
      if (axis === "row") cell.row = moved;
      else cell.col = moved;
      cell.ref = refOf(cell.col, cell.row);
      next.set(cell.ref, cell);
    }
    s.cells = next;
    s.merges = s.merges
      .map((m) => shiftMerge(m, axis, at, delta, drop))
      .filter((m): m is string => m !== undefined);
    const shiftKeys = <T>(src: Map<number, T>): Map<number, T> => {
      const out = new Map<number, T>();
      for (const [k, v] of src) {
        if (drop && k >= at && k < at + drop) continue;
        out.set(k >= at ? k + delta : k, v);
      }
      return out;
    };
    if (axis === "col") {
      s.colWidths = shiftKeys(s.colWidths);
      s.hiddenCols = new Set([...s.hiddenCols].filter((k) => !(drop && k >= at && k < at + drop)).map((k) => (k >= at ? k + delta : k)));
    } else {
      s.rowHeights = shiftKeys(s.rowHeights);
      if (s.hiddenRows) s.hiddenRows = new Set(shiftKeys(new Map([...s.hiddenRows].map((r) => [r, true]))).keys());
    }
    if (s.autoFilter) s.autoFilter = shiftAutoFilter(s.autoFilter, axis, at, count);
    for (const other of this.sheets) {
      for (const cell of other.cells.values()) {
        if (!cell.formula) continue;
        const rewritten = shiftRefs(cell.formula, axis, at, count, { edited: s.name, from: other.name });
        if (rewritten !== cell.formula) cell.formula = rewritten;
      }
    }
    this.rewriteChartReferences((ref, owner) => shiftChartRef(ref, owner, s.name, axis, at, count));
    for (const [name, node] of [...this.definedNames]) {
      const scope = node.t === "ref" || node.t === "range" ? node.sheet : undefined;
      if (scope !== undefined && !sameSheet(scope, s.name)) continue;
      const shifted = shiftNode(node, axis, at, delta, drop);
      if (shifted) this.definedNames.set(name, shifted);
      else this.definedNames.delete(name);
    }
    this.fresh();
  }

  merge(nameOrIndex: string | number, box: Rect): void {
    const s = this.sheet(nameOrIndex);
    if (!s || (box.c1 === box.c2 && box.r1 === box.r2)) return;
    const spec = rangeSpec(box);
    if (!s.merges.includes(spec)) s.merges.push(spec);
    this.touch(s, { c: box.c1, r: box.r1 });
    for (const at of eachCell(box)) {
      if (at.c === box.c1 && at.r === box.r1) continue;
      s.cells.delete(refOf(at.c, at.r));
    }
    for (const other of [...s.merges]) {
      if (other === spec) continue;
      if (rangesOverlap(parseRange(other), box)) s.merges = s.merges.filter((m) => m !== other);
    }
    this.fresh();
  }

  unmerge(nameOrIndex: string | number, box: Rect): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    s.merges = s.merges.filter((m) => !rangesOverlap(parseRange(m), box));
    this.fresh();
  }

  /** The merged block covering a position, if any. */
  mergeSpan(nameOrIndex: string | number, c: number, r: number): Rect | undefined {
    const s = this.sheet(nameOrIndex);
    for (const m of s?.merges ?? []) {
      const box = parseRange(m);
      if (inBox(box, c, r)) return box;
    }
    return undefined;
  }

  setColWidth(nameOrIndex: string | number, col: number, chars: number): void {
    const s = this.sheet(nameOrIndex);
    if (s) {
      s.colWidths.set(col, Math.max(1, Math.round(chars * 10) / 10));
      this.fresh();
    }
  }

  /**
   * Drops a picture anchored at a cell. The bytes travel as a data URL only inside
   * the session; the save writes them to `xl/media/` and links a drawing part.
   */
  addImage(nameOrIndex: string | number, at: { c: number; r: number }, src: string, widthPx: number, heightPx: number, name = "图片"): void {
    const s = this.sheet(nameOrIndex);
    if (!s || !src.startsWith("data:image/")) return;
    const mime = /^data:image\/([a-z0-9.+-]+)/i.exec(src)?.[1] ?? "png";
    const ext = mime === "jpeg" ? "jpg" : mime === "svg+xml" ? "svg" : mime;
    const taken = new Set<string>();
    for (const p of this.src?.parts.keys() ?? []) if (p.startsWith("xl/media/")) taken.add(p);
    for (const other of this.sheets) for (const img of other.images) taken.add(img.path);
    let n = 0;
    while (taken.has(`xl/media/image${n + 1}.${ext}`)) n++;
    s.images.push({
      name,
      col: at.c,
      row: at.r,
      colOff: 0,
      rowOff: 0,
      cx: Math.round(widthPx * EMU_PX),
      cy: Math.round(heightPx * EMU_PX),
      path: `xl/media/image${n + 1}.${ext}`,
      src,
    });
    this.fresh();
  }

  /** Removes one picture; its media part is dropped by the save when nothing references it. */
  removeImage(nameOrIndex: string | number, index: number): void {
    const s = this.sheet(nameOrIndex);
    if (s?.images[index]) {
      s.images.splice(index, 1);
      this.fresh();
    }
  }

  /**
   * Builds a chart from a rectangular block — first row names the series, first
   * column (when present) carries the categories — anchored beside the data.
   */
  addChart(kind: ChartModel["kind"], nameOrIndex: string | number, box: Rect, anchor: { c: number; r: number }, title: string): void {
    const s = this.sheet(nameOrIndex);
    if (!s || box.r2 <= box.r1) return;
    const series = chartSeriesForRange(s.name, box, (ref) => this.value(s.name, ref));
    const taken = new Set<string>();
    for (const p of this.src?.parts.keys() ?? []) if (p.startsWith("xl/charts/")) taken.add(p);
    for (const other of this.sheets) for (const ch of other.charts) taken.add(ch.part);
    let n = 0;
    while (taken.has(`xl/charts/chart${n + 1}.xml`)) n++;
    const model: ChartModel = {
      kind,
      title: title || series[0]?.name || "图表",
      barDir: "col",
      grouping: "clustered",
      series,
      col: anchor.c,
      row: anchor.r,
      colOff: 0,
      rowOff: 0,
      cx: Math.round(500 * EMU_PX),
      cy: Math.round(300 * EMU_PX),
      part: `xl/charts/chart${n + 1}.xml`,
    };
    model.srcRef = chartRef(s.name, rangeSpec(box));
    model.xml = chartPartXml(model, s.name);
    s.charts.push(model);
    this.fresh();
  }

  removeChart(nameOrIndex: string | number, index: number): void {
    const s = this.sheet(nameOrIndex);
    if (s?.charts[index]) {
      s.charts.splice(index, 1);
      this.fresh();
    }
  }

  setChartBox(nameOrIndex: string | number, index: number, box: { x: number; y: number; w: number; h: number }): void {
    const s = this.sheet(nameOrIndex), chart = s?.charts[index];
    if (!s || !chart) throw new Error("未找到图表");
    if (![box.x, box.y, box.w, box.h].every(Number.isFinite) || box.w <= 0 || box.h <= 0) throw new Error("无效图表尺寸");
    const next = { ...drawingAnchorAtPx(s, box.x, box.y), cx: Math.round(box.w * EMU_PX), cy: Math.round(box.h * EMU_PX) };
    if (chartGeometryFingerprint(next) === chartGeometryFingerprint(chart)) return;
    Object.assign(chart, next);
    this.fresh();
  }

  /** Build first, then commit: invalid edits never alter the workbook or undo history. */
  prepareChartEdit(nameOrIndex: string | number, index: number, edit: { kind: ChartModel["kind"]; barDir?: "col" | "bar"; title: string; source?: string; colors?: string[]; showLegend: boolean }): ChartModel {
    const s = this.sheet(nameOrIndex);
    const old = s?.charts[index];
    if (!s || !old) throw new Error("未找到图表");
    if (old.editUnsupported) throw new Error(old.editUnsupported);
    if (!["bar", "line", "area", "pie", "doughnut", "scatter"].includes(edit.kind)) throw new Error("不支持此图表类型");
    if (edit.colors && (!edit.colors.length || edit.colors.some((c) => !/^[0-9a-f]{6}$/i.test(c)))) throw new Error("无效图表颜色");
    const next: ChartModel = { ...this.chartForDisplay(s.name, old), kind: edit.kind, barDir: edit.barDir ?? "col", title: edit.title, showLegend: edit.showLegend, colors: edit.colors ? [...edit.colors] : old.colors };
    if (edit.source?.trim()) {
      const { sheet: sourceSheet, box } = chartRange(edit.source, s.name);
      if (!this.sheet(sourceSheet)) throw new Error(`找不到工作表：${sourceSheet}`);
      next.series = chartSeriesForRange(sourceSheet, box, (ref) => this.value(sourceSheet, ref));
      next.srcRef = chartRef(sourceSheet, rangeSpec(box));
    }
    if (edit.kind === "scatter" && next.series.some((ser) => !ser.cats.length || ser.cats.some((x) => !x.trim() || !Number.isFinite(Number(x))))) throw new Error("散点图的第一列需包含数字 X 坐标");
    next.editFields = [...new Set([...(old.editFields ?? []),
      ...(next.title !== old.title ? ["title" as const] : []),
      ...(next.kind !== old.kind || next.barDir !== old.barDir ? ["type" as const] : []),
      ...(JSON.stringify(next.series.map((x) => [x.nameRef, x.catRef, x.valRef])) !== JSON.stringify(old.series.map((x) => [x.nameRef, x.catRef, x.valRef])) ? ["source" as const] : []),
      ...(JSON.stringify(next.colors) !== JSON.stringify(old.colors) ? ["colors" as const] : []),
      ...(next.showLegend !== old.showLegend ? ["legend" as const] : []),
    ])];
    next.xml = chartPartXml(next, s.name);
    return next;
  }

  updateChart(nameOrIndex: string | number, index: number, next: ChartModel): void {
    const s = this.sheet(nameOrIndex);
    if (!s?.charts[index] || s.charts[index].part !== next.part) throw new Error("图表已改变，请重新选择");
    s.charts[index] = next;
    this.fresh();
  }

  private rewriteChartReferences(rewrite: (ref: string, owner: string) => string): void {
    for (const sheet of this.sheets) for (const chart of sheet.charts) {
      const rawRewrite = chart.editUnsupported && chart.sourceXml ? rewriteChartSourceReferences(chart.sourceXml, (ref) => rewrite(ref, sheet.name)) : undefined;
      let changed = false;
      chart.series = chart.series.map((ser) => {
        const next = { ...ser };
        for (const key of ["nameRef", "catRef", "valRef"] as const) {
          if (!ser[key]) continue;
          next[key] = rewrite(ser[key]!, sheet.name);
          changed ||= next[key] !== ser[key];
        }
        return next;
      });
      if (chart.editUnsupported) {
        if (rawRewrite) { chart.sourceXml = rawRewrite; chart.xml = rawRewrite; }
      } else if (changed) {
        chart.srcRef = chartSourceRange(chart.series, sheet.name);
        chart.xml = chartPartXml(chart, sheet.name);
      }
    }
  }

  /** References are authoritative when resolvable; unknown/external data retain caches. */
  chartForDisplay(defaultSheet: string, chart: ChartModel): ChartModel {
    return { ...chart, series: chart.series.map((s) => {
      const vals = s.valRef ? this.resolveFormulaRange(s.valRef, defaultSheet) : [];
      const cats = s.catRef ? this.resolveFormulaRange(s.catRef, defaultSheet) : [];
      const names = s.nameRef ? this.resolveFormulaRange(s.nameRef, defaultSheet) : [];
      return { ...s, name: s.nameRef === "#REF!" ? "#REF!" : names.length ? String(names[0] ?? "") : s.name, cats: s.catRef === "#REF!" ? [] : cats.length ? cats.map((v) => String(v ?? "")) : [...s.cats], vals: s.valRef === "#REF!" ? s.vals.map(() => null) : vals.length ? vals.map((v) => typeof v === "number" && Number.isFinite(v) ? v : null) : [...s.vals] };
    }) };
  }

  /** The dxf record a cellIs rule points at, resolved for rendering. */
  dxfOf(id: number): CondStyle | undefined {
    return this.dxfs[id] ?? this.editorDxfs[id - this.dxfs.length];
  }

  /** Resolves a chart's `f` reference (`'Sheet'!$B$2:$B$4`) to cell values in row order. */
  resolveFormulaRange(f: string, defaultSheet: string): (string | number | null)[] {
    try {
      const { sheet, box } = chartRange(f, defaultSheet);
      if (!this.sheet(sheet)) return [];
      return [...eachCell(box)].map(({ c, r }) => {
        const v = this.value(sheet, refOf(c, r));
        return typeof v === "number" || typeof v === "string" ? v : v === null ? null : String(v);
      });
    } catch { return []; }
  }

  /** Replaces whatever formats the box carried with one editor-created rule. */
  setCond(nameOrIndex: string | number, box: Rect, rule: { type: "colorScale"; colors: string[] } | { type: "dataBar"; color: string }): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    const ref = rangeSpec(box);
    const priority = Math.max(0, ...s.cond.flatMap((c) => c.rules.map((r) => r.priority))) + 1;
    s.cond = s.cond.filter((cf) => cf.ref !== ref);
    s.cond.push({ ref, rules: [{ type: rule.type, priority, colors: rule.type === "colorScale" ? [...rule.colors] : [rule.color] }] });
    this.fresh();
  }

  /** Sets (or replaces) the note on one cell. */
  setComment(nameOrIndex: string | number, ref: string, text: string, author = "批注"): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    s.comments = s.comments.filter((c) => c.ref !== ref);
    if (text.trim()) s.comments.push({ ref, text, author });
    this.fresh();
  }

  /** The dxf records the editor created for its own cellIs rules, for the writer. */
  editorDxfTable(): CondStyle[] {
    return this.editorDxfs;
  }

  private editorDxfs: CondStyle[] = [];

  /** A cellIs rule with the given look, its style interned as a new dxf record. */
  setCondCellIs(nameOrIndex: string | number, box: Rect, op: string, formula: string, style: CondStyle = { bgColor: "FFC7CE", fontColor: "9C0006" }): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    const key = JSON.stringify(style);
    let dxfId = this.dxfs.findIndex((d) => JSON.stringify(d) === key);
    if (dxfId < 0) {
      const at = this.editorDxfs.findIndex((d) => JSON.stringify(d) === key);
      if (at >= 0) dxfId = this.dxfs.length + at;
      else {
        this.editorDxfs.push(style);
        dxfId = this.dxfs.length + this.editorDxfs.length - 1;
      }
    }
    const ref = rangeSpec(box);
    const priority = Math.max(0, ...s.cond.flatMap((c) => c.rules.map((r) => r.priority))) + 1;
    s.cond = s.cond.filter((cf) => cf.ref !== ref);
    s.cond.push({ ref, rules: [{ type: "cellIs", priority, op, formula, dxfId }] });
    this.fresh();
  }

  /** Drops every rule whose range overlaps the box. */
  clearCond(nameOrIndex: string | number, box: Rect): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    const before = JSON.stringify(s.cond);
    s.cond = s.cond
      .map((cf) => ({ ref: cf.ref, rules: cf.rules.filter((r) => !rangesOverlap(parseRange(cf.ref), box)) }))
      .filter((cf) => cf.rules.length);
    if (JSON.stringify(s.cond) !== before) this.fresh();
  }

  setRowHeight(nameOrIndex: string | number, row: number, points: number): void {
    const s = this.sheet(nameOrIndex);
    if (s) {
      s.rowHeights.set(row, Math.max(1, Math.round(points * 10) / 10));
      this.fresh();
    }
  }

  hideCol(nameOrIndex: string | number, col: number, hidden: boolean): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    if (hidden) s.hiddenCols.add(col);
    else s.hiddenCols.delete(col);
    this.fresh();
  }

  setFreeze(nameOrIndex: string | number, rows: number, cols: number): void {
    const s = this.sheet(nameOrIndex);
    if (s) {
      s.freeze = rows || cols ? { rows, cols } : undefined;
      this.fresh();
    }
  }

  /** Reorder the rows of `box` by one column. Formulas travel with their cells. */
  sort(nameOrIndex: string | number, box: Rect, byCol: number, desc = false): void {
    const s = this.sheet(nameOrIndex);
    if (!s) return;
    const rows: { key: Value; cells: Cell[] }[] = [];
    for (let r = box.r1; r <= box.r2; r++) {
      const cells: Cell[] = [];
      for (let c = box.c1; c <= box.c2; c++) {
        const cell = s.cells.get(refOf(c, r));
        if (cell) cells.push(cell);
      }
      rows.push({ key: this.value(s.name, refOf(byCol, r)), cells });
    }
    rows.sort((x, y) => orderValues(x.key, y.key, desc));
    for (const row of rows) for (const cell of row.cells) s.cells.delete(cell.ref);
    for (const [slot, row] of rows.entries()) {
      for (const cell of row.cells) {
        cell.row = box.r1 + slot;
        cell.ref = refOf(cell.col, cell.row);
        s.cells.set(cell.ref, cell);
      }
    }
    this.fresh();
  }

  // -------------------------------------------------------------- sheet list

  addSheet(name?: string): Sheet {
    let i = this.sheets.length + 1;
    let label = name?.trim() || `Sheet${i}`;
    while (this.sheet(label)) label = `Sheet${++i}`;
    const sheet: Sheet = { name: label, cells: new Map(), merges: [], colWidths: new Map(), hiddenCols: new Set(), rowHeights: new Map(), images: [], charts: [], cond: [], comments: [] };
    this.sheets.push(sheet);
    this.fresh();
    return sheet;
  }

  renameSheet(index: number, name: string): void {
    const s = this.sheets[index];
    const clean = name.trim();
    if (!s || !clean || this.sheet(clean)) return;
    const old = s.name;
    s.name = clean;
    // Stored formulas qualify other sheets by name, so a rename has to follow.
    for (const sheet of this.sheets) {
      for (const cell of sheet.cells.values()) {
        if (cell.formula) cell.formula = renameSheetRef(cell.formula, old, clean);
      }
    }
    this.rewriteChartReferences((ref, owner) => {
      try { const parsed = chartRange(ref, owner); return parsed.sheet.toLowerCase() === old.toLowerCase() ? chartRef(clean, rangeSpec(parsed.box)) : ref; } catch { return ref; }
    });
    this.fresh();
  }

  removeSheet(index: number): boolean {
    if (this.sheets.length < 2 || !this.sheets[index]) return false;
    const removed = this.sheets[index].name;
    this.rewriteChartReferences((ref, owner) => {
      try { return chartRange(ref, owner).sheet.toLowerCase() === removed.toLowerCase() ? "#REF!" : ref; } catch { return ref; }
    });
    this.sheets.splice(index, 1);
    this.activeSheet = Math.min(this.activeSheet, this.sheets.length - 1);
    this.fresh();
    return true;
  }

  // ------------------------------------------------------- search and history

  /** Displayed text is what the analyst sees, so search runs over it. */
  find(nameOrIndex: string | number, needle: string, box?: Rect): string[] {
    const s = this.sheet(nameOrIndex);
    if (!s || !needle) return [];
    const target = needle.toLowerCase();
    const area = box ?? { c1: 0, r1: 0, c2: this.bounds(s.name).maxCol, r2: this.bounds(s.name).maxRow };
    const hits: string[] = [];
    for (const at of eachCell(area)) {
      const ref = refOf(at.c, at.r);
      if (this.display(s.name, ref).text.toLowerCase().includes(target)) hits.push(ref);
    }
    return hits;
  }

  replaceAll(nameOrIndex: string | number, needle: string, withText: string, box?: Rect): number {
    const s = this.sheet(nameOrIndex);
    if (!s || !needle) return 0;
    let n = 0;
    for (const cell of s.cells.values()) {
      if (box && !inBox(box, cell.col, cell.row)) continue;
      if (typeof cell.constant === "string" && cell.constant.includes(needle)) {
        cell.constant = cell.constant.split(needle).join(withText);
        cell.runs = undefined;
        n++;
      } else if (typeof cell.formula === "string" && cell.formula.includes(needle)) {
        cell.formula = cell.formula.split(needle).join(withText);
        n++;
      }
    }
    if (n) this.fresh();
    return n;
  }

  /** Numbers under a selection, for the status bar. */
  stats(nameOrIndex: string | number, box: Rect): { count: number; numeric: number; sum: number; avg: number; min: number; max: number } {
    const s = this.sheet(nameOrIndex);
    const out = { count: 0, numeric: 0, sum: 0, avg: 0, min: Infinity, max: -Infinity };
    if (!s) return out;
    for (const at of eachCell(box)) {
      const ref = refOf(at.c, at.r);
      const cell = s.cells.get(ref);
      if (!cell) continue;
      out.count++;
      const v = this.value(s.name, ref);
      const n = typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : typeof v === "string" && v !== "" && Number.isFinite(Number(v)) ? Number(v) : undefined;
      if (n === undefined) continue;
      out.numeric++;
      out.sum += n;
      out.min = Math.min(out.min, n);
      out.max = Math.max(out.max, n);
    }
    out.avg = out.numeric ? out.sum / out.numeric : 0;
    return out;
  }

  snapshot(): WorkbookSnapshot {
    return {
      sheets: this.sheets.map(cloneSheet),
      styles: this.styles.map((st) => ({ ...st, border: st.border ? { ...st.border } : undefined })),
      names: [...this.definedNames].map(([k, n]) => [k, structuredClone(n)]),
      activeSheet: this.activeSheet,
    };
  }

  restore(snap: WorkbookSnapshot): void {
    this.sheets = snap.sheets.map(cloneSheet);
    this.styles = snap.styles.map((st) => ({ ...st, border: st.border ? { ...st.border } : undefined }));
    this.styleIds = new Map(this.styles.map((st, i) => [this.keyOf(st), i]));
    this.definedNames = new Map(snap.names.map(([k, n]) => [k, structuredClone(n)]));
    this.activeSheet = snap.activeSheet;
    this.fresh();
  }
}

function cloneSheet(s: Sheet): Sheet {
  return {
    name: s.name,
    cells: new Map([...s.cells].map(([k, c]) => [k, { ...c, runs: c.runs?.map((r) => ({ ...r })) }])),
    merges: [...s.merges],
    colWidths: new Map(s.colWidths),
    hiddenCols: new Set(s.hiddenCols),
    hiddenRows: s.hiddenRows ? new Set(s.hiddenRows) : undefined,
    filteredRows: s.filteredRows ? new Set(s.filteredRows) : undefined,
    autoFilter: s.autoFilter ? structuredClone(s.autoFilter) : undefined,
    filterAtRead: s.filterAtRead,
    rowHeights: new Map(s.rowHeights),
    defaultRowHeight: s.defaultRowHeight,
    defaultColWidth: s.defaultColWidth,
    images: s.images.map((i) => ({ ...i })),
    charts: s.charts.map((c) => ({ ...c, colors: c.colors ? [...c.colors] : undefined, editFields: c.editFields ? [...c.editFields] : undefined, series: c.series.map((ser) => ({ ...ser, cats: [...ser.cats], vals: [...ser.vals] })) })),
    comments: s.comments.map((c) => ({ ...c })),
    cond: s.cond.map((cf) => ({ ref: cf.ref, rules: cf.rules.map((r) => ({ ...r, colors: r.colors ? [...r.colors] : undefined })) })),
    freeze: s.freeze ? { ...s.freeze } : undefined,
    hidden: s.hidden,
    srcPath: s.srcPath,
    colsAtRead: s.colsAtRead,
    imagesAtRead: s.imagesAtRead,
    drawingAnchorsAtRead: s.drawingAnchorsAtRead ? [...s.drawingAnchorsAtRead] : undefined,
  };
}

/** The drawing fingerprint a save compares against, so untouched charts keep their exact bytes. */
export function imagesFingerprint(sheet: Sheet): string {
  return JSON.stringify([sheet.images.map((i) => [i.col, i.row, i.colOff, i.rowOff, i.cx, i.cy, i.path]), sheet.charts.map((c) => [c.col, c.row, c.colOff, c.rowOff, c.cx, c.cy, c.part])]);
}

/** Reads a chart part (chartSpace) into the model, anchored by the caller. */
function parseChart(pkg: OpcPackage, path: string, sheetName: string, resolve: (f: string) => (string | number | null)[]): ChartModel | undefined {
  const chart = pkg.xml(path)?.chartSpace?.chart;
  if (!chart) return undefined;
  const titleText = (node: any): string => {
    const out: string[] = [];
    const walk = (o: any): void => {
      if (!o || typeof o !== "object") return;
      for (const [k, v] of Object.entries(o)) {
        if (k === "a:t" || k === "t") out.push(decodeEntities(typeof v === "string" ? v : String((v as { "#text"?: unknown })?.["#text"] ?? "")));
        else if (v && typeof v === "object") walk(v);
      }
    };
    walk(node);
    return out.join("");
  };
  const plot = chart.plotArea ?? {};
  const kinds: [string, ChartModel["kind"]][] = [
    ["barChart", "bar"],
    ["lineChart", "line"],
    ["pieChart", "pie"],
    ["doughnutChart", "doughnut"],
    ["areaChart", "area"],
    ["scatterChart", "scatter"],
  ];
  const cachePts = (node: any): Map<number, string> => {
    const pts = new Map<number, string>();
    for (const p of asArray(node?.pt)) {
      const raw = p.v;
      const s = typeof raw === "string" ? raw : String(raw?.["#text"] ?? "");
      pts.set(Number(p["@_idx"] ?? pts.size), s);
    }
    return pts;
  };
  const serOf = (ser: any, index: number): ChartSeries => {
    const cached = cachePts(ser.tx?.strRef?.strCache);
    const name = String(cached.get(0) ?? (typeof ser.tx?.v === "string" ? ser.tx.v : String(ser.tx?.v?.["#text"] ?? ""))) || String(resolve(textOfF(ser.tx))[0] ?? `系列${index + 1}`);
    const cat = ser.cat ?? ser.xVal;
    const val = ser.val ?? ser.yVal;
    const catRef = cat?.strRef ?? cat?.numRef;
    const catCache = cachePts(catRef?.strCache ?? catRef?.numCache ?? cat?.strLit ?? cat?.numLit);
    const valCache = cachePts(val?.numRef?.numCache ?? val?.numLit);
    let cats: string[];
    let vals: (number | null)[];
    if (valCache.size) {
      const count = Number(val?.numRef?.numCache?.ptCount?.["@_val"] ?? Math.max(0, ...valCache.keys()) + 1);
      vals = Array.from({ length: count }, (_, i) => {
        const raw = valCache.get(i);
        return raw === undefined || raw === "" ? null : Number.isFinite(Number(raw)) ? Number(raw) : null;
      });
      cats = Array.from({ length: Number((catRef?.strCache ?? catRef?.numCache ?? cat?.strLit ?? cat?.numLit)?.ptCount?.["@_val"] ?? count) }, (_, i) => catCache.get(i) ?? "");
      if (!catCache.size && textOfF(cat)) cats = resolve(textOfF(cat)).map((v) => String(v ?? ""));
    } else {
      // No cached values: evaluate the series' own references against the sheet.
      const rawVals = resolve(textOfF(val));
      vals = rawVals.map((v) => (typeof v === "number" ? v : typeof v === "string" && v !== "" && Number.isFinite(Number(v)) ? Number(v) : null));
      cats = resolve(textOfF(cat)).map((v) => String(v ?? ""));
    }
    return { name: decodeEntities(name), cats: cats.map(decodeEntities), vals, nameRef: textOfF(ser.tx) || undefined, catRef: textOfF(cat) || undefined, valRef: textOfF(val) || undefined, catNumeric: !!cat?.numRef || !!ser.xVal };
  };
  for (const [tag, kind] of kinds) {
    const node = asArray(plot[tag])[0];
    if (!node) continue;
    const series = asArray(node.ser).map(serOf);
    if (!series.length) continue;
    return {
      kind,
      title: titleText(chart.title),
      sourceXml: text(pkg.bytes(path)!),
      dataAtRead: chartDataFingerprint(series),
      srcRef: chartSourceRange(series, sheetName),
      showLegend: !!chart.legend,
      colors: (kind === "pie" || kind === "doughnut") && asArray(asArray(node.ser)[0]?.dPt).every((p: any) => p.spPr?.solidFill?.srgbClr?.["@_val"]) && asArray(asArray(node.ser)[0]?.dPt).length ? asArray(asArray(node.ser)[0].dPt).map((p: any) => String(p.spPr.solidFill.srgbClr["@_val"])) : series.every((_, i) => asArray(node.ser)[i]?.spPr?.solidFill?.srgbClr?.["@_val"]) ? asArray(node.ser).map((ser: any) => String(ser.spPr.solidFill.srgbClr["@_val"])) : undefined,
      editUnsupported: Object.keys(plot).filter((k) => /Chart$/.test(k)).length !== 1 || asArray(plot[tag]).length !== 1 || chart.pivotFmts || pkg.xml(path)?.chartSpace?.pivotSource || pkg.xml(path)?.chartSpace?.externalData || (node.grouping?.["@_val"] && !["clustered", "standard"].includes(node.grouping["@_val"])) ? "此组合、堆积或外部数据图表目前保留原样，暂不支持编辑" : undefined,
      barDir: node.barDir?.["@_val"] === "bar" ? "bar" : "col",
      grouping: node.grouping?.["@_val"],
      series,
      col: 0,
      row: 0,
      colOff: 0,
      rowOff: 0,
      cx: 0,
      cy: 0,
      part: path,
    };
  }
  return undefined;
}

const textOfF = (ref: any): string => {
  const f = ref?.strRef?.f ?? ref?.numRef?.f;
  return decodeEntities(typeof f === "string" ? f : String(f?.["#text"] ?? ""));
};

/** The conditional-format fingerprint a save compares against the sheet's own. */
export function condFingerprint(sheet: Sheet): string {
  return JSON.stringify(sheet.cond);
}

/** Color at fraction `t` (0..1) along 2–3 stops, as six hex digits. */
export function scaleColorAt(colors: string[], t: number): string {
  const rgb = (hex: string): [number, number, number] => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) || 0) as [number, number, number];
  const stops = colors.map(rgb);
  const x = Math.max(0, Math.min(1, t));
  const seg = stops.length > 2 ? Math.min(stops.length - 2, Math.floor(x * (stops.length - 1))) : 0;
  const local = stops.length > 2 ? x * (stops.length - 1) - seg : x;
  const a = stops[seg];
  const b = stops[seg + 1];
  const mix = a.map((v, i) => Math.round(v + (b[i] - v) * local));
  return mix.map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** A cellIs comparison the way Excel evaluates it: numeric when both sides are. */
export function cfCompare(v: Value, op: string, operand: Value): boolean {
  const num = (x: Value): number | undefined =>
    typeof x === "number" ? x : typeof x === "string" && x !== "" && Number.isFinite(Number(x)) ? Number(x) : undefined;
  const a = num(v);
  const b = num(operand);
  if (a !== undefined && b !== undefined) return compareOp(op, a, b);
  const s = String(v ?? "");
  const t = String(operand ?? "");
  return compareOp(op, s.toLowerCase(), t.toLowerCase());
}

function compareOp(op: string, a: string | number, b: string | number): boolean {
  switch (op) {
    case "greaterThan":
      return a > b;
    case "greaterThanOrEqual":
      return a >= b;
    case "lessThan":
      return a < b;
    case "lessThanOrEqual":
      return a <= b;
    case "notEqual":
      return a !== b;
    default:
      return a === b;
  }
}

const inBox = (box: Rect, c: number, r: number): boolean => c >= box.c1 && c <= box.c2 && r >= box.r1 && r <= box.r2;

const rangesOverlap = (a: Rect, b: Rect): boolean => Math.max(a.c1, b.c1) <= Math.min(a.c2, b.c2) && Math.max(a.r1, b.r1) <= Math.min(a.r2, b.r2);

/**
 * Move a merged block across a structural edit. On the edited axis a range that
 * straddles the deleted band shrinks to its surviving part; one entirely inside
 * the band goes away, as does a block that collapses to a single cell.
 */
function shiftMerge(spec: string, axis: Axis, at: number, delta: number, drop: number): string | undefined {
  const box = parseRange(spec);
  const inside = (v: number): boolean => drop > 0 && v >= at && v < at + drop;
  const move = (v: number): number => (v >= at ? v + delta : v);
  const span = (lo: number, hi: number): [number, number] | undefined => {
    if (inside(lo) && inside(hi)) return undefined;
    const a = inside(lo) ? at : move(lo);
    const b = inside(hi) ? at - 1 : move(hi);
    if (b < a || a < 0) return undefined;
    return [a, b];
  };
  if (axis === "row") {
    const rows = span(box.r1, box.r2);
    if (!rows || (rows[0] === rows[1] && box.c1 === box.c2)) return undefined;
    return `${refOf(box.c1, rows[0])}:${refOf(box.c2, rows[1])}`;
  }
  const cols = span(box.c1, box.c2);
  if (!cols || (cols[0] === cols[1] && box.r1 === box.r2)) return undefined;
  return `${refOf(cols[0], box.r1)}:${refOf(cols[1], box.r2)}`;
}

/** Blanks last in either direction, numbers before text, matching the default sort of a grid. */
function orderValues(a: Value, b: Value, desc = false): number {
  const an = a === null || a === "";
  const bn = b === null || b === "";
  if (an && bn) return 0;
  if (an) return 1;
  if (bn) return -1;
  const sign = desc ? -1 : 1;
  if (typeof a === "number" && typeof b === "number") return sign * (a - b);
  if (typeof a === "number" && typeof b !== "number") return -1;
  if (typeof a !== "number" && typeof b === "number") return 1;
  return sign * (String(a) < String(b) ? -1 : String(a) > String(b) ? 1 : 0);
}

/** Textual sheet qualifier rewrite for renames; quoted names keep their quotes. */
function renameSheetRef(formula: string, oldName: string, clean: string): string {
  const esc = oldName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return formula
    .replace(new RegExp(`'${esc}'!`, "g"), `'${clean}'!`)
    .replace(new RegExp(`(?<![A-Za-z0-9_.'"-])${esc}!`, "g"), `${clean}!`);
}

/** Sheet qualifiers are case-insensitive and may or may not carry quotes. */
const sameSheet = (a: string, b: string): boolean => a.replace(/^'+|'+$/g, "").toLowerCase() === b.toLowerCase();

function shiftNode(node: Node, axis: Axis, at: number, delta: number, drop: number): Node | undefined {
  const line = (v: number): number => {
    if (drop && v >= at && v < at + drop) return -1;
    return v >= at ? v + delta : v;
  };
  if (node.t === "ref") {
    const col = axis === "col" ? line(node.col) : node.col;
    const row = axis === "row" ? line(node.row) : node.row;
    return col < 0 || row < 0 ? undefined : { ...node, col, row };
  }
  if (node.t === "range") {
    const from = { col: axis === "col" ? line(node.from.col) : node.from.col, row: axis === "row" ? line(node.from.row) : node.from.row };
    const to = { col: axis === "col" ? line(node.to.col) : node.to.col, row: axis === "row" ? line(node.to.row) : node.to.row };
    return from.col < 0 || from.row < 0 || to.col < 0 || to.row < 0 ? undefined : { ...node, from, to };
  }
  return node;
}

/** Render a defined name's target back to the text a workbook stores. */
function nodeText(node: Node): string | undefined {
  if (node.t === "ref") return `${node.sheet ? `'${node.sheet}'!` : ""}${refOf(node.col, node.row)}`;
  if (node.t === "range") return `${node.sheet ? `'${node.sheet}'!` : ""}${refOf(node.from.col, node.from.row)}:${refOf(node.to.col, node.to.row)}`;
  return undefined;
}

function readSharedString(si: any): { text: string; runs: Run[] } {
  if (!si) return { text: "", runs: [] };
  if (typeof si === "string") return { text: decodeEntities(si), runs: [{ text: decodeEntities(si) }] };
  if (si.t !== undefined) {
    const t = decodeEntities(typeof si.t === "string" ? si.t : (si.t["#text"] ?? ""));
    return { text: t, runs: [{ text: t }] };
  }
  const runs = asArray(si.r).map((r: any) => {
    const rPr = r.rPr ?? {};
    return {
      text: decodeEntities(typeof r.t === "string" ? r.t : (r.t?.["#text"] ?? "")),
      bold: rPr.b !== undefined,
      italic: rPr.i !== undefined,
      underline: rPr.u !== undefined,
      size: rPr.sz ? Number(rPr.sz["@_val"]) : undefined,
      color: rPr.color?.["@_rgb"] ? String(rPr.color["@_rgb"]).slice(-6) : undefined,
    } as Run;
  });
  return { text: runs.map((r) => r.text).join(""), runs };
}

const CELL_IN_FORMULA = /(?<![A-Za-z0-9_!$'])(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})(?![0-9A-Za-z_(])/g;

/**
 * Shift relative references from a shared formula's anchor cell to a target cell,
 * leaving `$`-anchored parts and string literals untouched.
 */
function translate(formula: string, fromCol: number, fromRow: number, toCol: number, toRow: number): string {  const dCol = toCol - fromCol;
  const dRow = toRow - fromRow;
  if (!dCol && !dRow) return formula;
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) =>
      i % 2 === 1
        ? part
        : part.replace(CELL_IN_FORMULA, (_m, absC, col, absR, row) => {
            const c = absC === "$" ? colToIndex(col) : colToIndex(col) + dCol;
            const r = absR === "$" ? Number(row) - 1 : Number(row) - 1 + dRow;
            if (c < 0 || r < 0) return "#REF!";
            return `${absC}${indexToCol(c)}${absR}${r + 1}`;
          }),
    )
    .join("");
}

/** Relative-reference shift for the fill handle: the same rule shared formulas use. */
export function shiftFormulaRefs(formula: string, dCol: number, dRow: number): string {
  return translate(formula, 0, 0, dCol, dRow);
}

function parseName(ref: string | undefined): Node | undefined {
  if (!ref) return undefined;
  const spec = ref.trim();
  const bang = spec.lastIndexOf("!");
  const sheet = bang < 0 ? undefined : spec.slice(0, bang).replace(/^'+|'+$/g, "");
  const body = bang < 0 ? spec : spec.slice(bang + 1);
  const single = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/.exec(body);
  if (single) return { t: "ref", sheet, col: colToIndex(single[1]), row: Number(single[2]) - 1 };
  const range = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*):\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/.exec(body);
  if (range) {
    return {
      t: "range",
      sheet,
      from: { col: colToIndex(range[1]), row: Number(range[2]) - 1 },
      to: { col: colToIndex(range[3]), row: Number(range[4]) - 1 },
    };
  }
  return undefined;
}
