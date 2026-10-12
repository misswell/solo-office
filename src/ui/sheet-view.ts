import type { Cell, CellStyle, CondStyle, Sheet, SheetImage, Spreadsheet } from "../core/xlsx.js";
import { DEFAULT_COL_CHARS, DEFAULT_ROW_PT, cfCompare, colWidthPx, imageBoxPx, rowHeightPx, scaleColorAt, shiftFormulaRefs, type CondRule, type ChartModel } from "../core/xlsx.js";
import { openChartEditor } from "./chart-edit.js";
import { openFilterEditor } from "./filter-edit.js";
import { filterDataRegion, filterRange, validateFilterRule } from "../core/filter.js";
import { renderChartSvg } from "../core/chart.js";
import { indexToCol, isError } from "../core/formula.js";
import { eachCell, parseRange, rect, refOf, type Axis, type Rect } from "../core/sheet-ops.js";

export interface SheetViewHooks {
  /** The model changed: the shell takes an undo snapshot and refreshes its chrome. */
  edited(): void;
  /** Selection or active cell moved: refresh address box, formula bar, toolbar, stats. */
  selected(): void;
  /** Right-click inside the grid. */
  menu(view: SheetView, x: number, y: number): void;
}

const ROW_COUNT = 60;
const COL_COUNT = 18;
/** Defaults match the core geometry, so pictures anchored in cells land exactly on the grid. */
const COL_W = Math.round(DEFAULT_COL_CHARS * 7.2);
const ROW_H = Math.round((DEFAULT_ROW_PT * 4) / 3);
const HEAD_W = 46;
const HEAD_H = 22;

const colPx = (s: Sheet, c: number): number => colWidthPx(s, c);
/** Row heights are stored in points; CSS gets whole pixels so our own arithmetic matches layout. */
const rowPx = (s: Sheet, r: number): number => rowHeightPx(s, r);
const css = (v: string | undefined): string | undefined => (v ? `#${v.replace(/^#/, "")}` : undefined);

export class SheetView {
  readonly wrap = document.createElement("div");
  /** Visible column indexes and rendered row count, needed by the header hit-tests. */
  cols: number[] = [];
  rowCount = ROW_COUNT;
  sel: Rect = rect(0, 0, 0, 0);
  active = { c: 0, r: 0 };
  /** Document zoom; pointer deltas divide by it so drags stay true to the grid. */
  zoom = 1;

  private readonly frame = document.createElement("div");
  private readonly fillBox = document.createElement("i");
  private readonly pics = document.createElement("div");
  private readonly grips = document.createElement("div");
  private editor: HTMLElement | null = null;
  selectedChart = -1;
  private chartDrag: { index: number; element: HTMLElement; pointerId: number; startX: number; startY: number; scale: number; resize: boolean; original: { x: number; y: number; w: number; h: number }; preview: { x: number; y: number; w: number; h: number } } | null = null;
  private dragging = false;
  filling: Rect | null = null;
  sizing: { axis: Axis; line: number; start: number; origin: number } | null = null;

  constructor(
    host: HTMLElement,
    readonly wb: Spreadsheet,
    public sheetIndex: number,
    readonly hooks: SheetViewHooks,
  ) {
    this.wrap.className = "sheet-wrap";
    this.wrap.tabIndex = 0;
    this.frame.className = "sel-frame";
    this.fillBox.className = "fill-box";
    this.pics.className = "sheet-pics";
    host.replaceChildren(this.wrap);
    wireSheet(this);
  }

  get sheet(): Sheet {
    return this.wb.sheets[this.sheetIndex];
  }

  get isEditing(): boolean {
    return this.editor !== null;
  }

  // -------------------------------------------------------------------- paint

  render(): void {
    this.cancelChartDrag();
    const s = this.sheet;
    this.wb.refreshFilters();
    const bounds = this.wb.bounds(s.name);
    this.buildCond();
    this.cols = [];
    for (let c = 0; c <= Math.max(COL_COUNT - 1, bounds.maxCol + 3); c++) if (!s.hiddenCols.has(c)) this.cols.push(c);
    this.rowCount = Math.max(ROW_COUNT, bounds.maxRow + 5);

    const span = new Map<string, { rs: number; cs: number }>();
    const covered = new Set<string>();
    for (const m of s.merges) {
      const box = parseRange(m);
      span.set(refOf(box.c1, box.r1), { rs: box.r2 - box.r1 + 1, cs: box.c2 - box.c1 + 1 });
      for (const at of eachCell(box)) if (at.c !== box.c1 || at.r !== box.r1) covered.add(refOf(at.c, at.r));
    }

    const table = document.createElement("table");
    table.className = "grid";
    const group = document.createElement("colgroup");
    group.append(colEl(HEAD_W));
    for (const c of this.cols) group.append(colEl(colPx(s, c)));
    table.append(group);

    const fz = s.freeze ?? { rows: 0, cols: 0 };
    const head = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "corner";
    head.append(corner);
    for (const c of this.cols) {
      const th = document.createElement("th");
      th.className = "colh";
      th.textContent = indexToCol(c);
      th.dataset.col = String(c);
      if (this.colIndex(c) < fz.cols) {
        th.classList.add("fz");
        th.style.left = `${colLeft(s, this.cols, c)}px`;
      }
      head.append(th);
    }
    const thead = document.createElement("thead");
    thead.append(head);
    table.append(thead);

    const body = document.createElement("tbody");
    for (let r = 0; r < this.rowCount; r++) {
      if (s.hiddenRows?.has(r) || s.filteredRows?.has(r)) continue;
      const tr = document.createElement("tr");
      tr.style.height = `${rowPx(s, r)}px`;
      const rh = document.createElement("th");
      rh.className = "rowh";
      rh.textContent = String(r + 1);
      rh.dataset.row = String(r);
      tr.append(rh);
      for (const c of this.cols) {
        const ref = refOf(c, r);
        if (covered.has(ref)) continue;
        const td = document.createElement("td");
        td.dataset.ref = ref;
        td.dataset.col = String(c);
        td.dataset.row = String(r);
        const sp = span.get(ref);
        if (sp) {
          td.rowSpan = sp.rs;
          td.colSpan = sp.cs;
        }
        if (this.colIndex(c) < fz.cols) {
          td.classList.add("fz");
          td.style.left = `${colLeft(s, this.cols, c)}px`;
        }
        const cell = s.cells.get(ref);
        if (cell) this.paint(td, cell);
        else applyStyle(td, this.wb.styleAt(s.name, c, r));
        if (s.autoFilter) {
          const box = filterRange(s.autoFilter.ref);
          if (r === box.r1 && c >= box.c1 && c <= box.c2) {
            const id = c - box.c1, arrow = document.createElement("button");
            arrow.className = "filter-arrow" + (s.autoFilter.columns.some((col) => col.id === id) ? " filter-active" : "");
            arrow.textContent = "▾"; arrow.dataset.filterCol = String(id);
            arrow.setAttribute("aria-label", `筛选 ${indexToCol(c)} 列`);
            arrow.onpointerdown = (e) => e.stopPropagation(); arrow.onkeydown = (e) => e.stopPropagation();
            arrow.onclick = (e) => { e.stopPropagation(); this.editFilterColumn(id); };
            td.classList.add("filter-header"); td.append(arrow);
          }
        }
        this.applyCond(td, ref);
        const note = s.comments.find((cm) => cm.ref === ref);
        if (note) {
          td.classList.add("has-note");
          td.title = `${note.author ? note.author + "：" : ""}${note.text}`;
        }
        tr.append(td);
      }
      if (r < fz.rows) {
        const top = `${headTop(s, r)}px`;
        rh.classList.add("fz");
        rh.style.top = top;
        for (const td of tr.querySelectorAll<HTMLElement>("td")) {
          td.classList.add("fz");
          td.style.top = top;
        }
      }
      body.append(tr);
    }
    table.append(body);

    this.pics.replaceChildren(
      ...s.images.map((img, i) => {
        const box = imageBoxPx(s, img);
        const el = document.createElement("img");
        el.className = "sheet-pic";
        el.dataset.idx = String(i);
        el.src = img.src;
        el.alt = img.name;
        el.style.left = `${box.x}px`;
        el.style.top = `${box.y}px`;
        el.style.width = `${box.w}px`;
        el.style.height = `${box.h}px`;
        return el;
      }),
      ...s.charts.map((chart, i) => {
        const box = imageBoxPx(s, chart);
        const el = document.createElement("div");
        el.className = "sheet-chart" + (this.selectedChart === i ? " chart-selected" : "");
        el.tabIndex = 0;
        el.setAttribute("role", "button");
        el.setAttribute("aria-label", `图表：${chart.title || "无标题"}，双击编辑`);
        el.title = "双击编辑图表";
        el.onpointerdown = (e) => this.startChartDrag(i, el, e);
        el.onpointermove = (e) => this.moveChartDrag(e);
        el.onpointerup = (e) => this.endChartDrag(e);
        el.onpointercancel = () => this.cancelChartDrag();
        el.onlostpointercapture = () => { if (this.chartDrag?.element === el) this.cancelChartDrag(); };
        el.ondblclick = () => this.editSelectedChart();
        el.onkeydown = (e) => this.chartKey(i, e);
        el.dataset.idx = String(i);
        el.style.left = `${box.x}px`;
        el.style.top = `${box.y}px`;
        el.style.width = `${box.w}px`;
        el.style.height = `${box.h}px`;
        el.innerHTML = renderChartSvg(this.wb.chartForDisplay(s.name, chart), Math.round(box.w), Math.round(box.h));
        const grip = document.createElement("i");
        grip.className = "chart-resize";
        grip.tabIndex = 0;
        grip.setAttribute("role", "button");
        grip.setAttribute("aria-label", "调整图表大小，方向键微调");
        grip.title = "拖动调整大小";
        el.append(grip);
        return el;
      }),
    );

    // One drag handle per visible line: column widths ride the header edge,
    // row heights the row edge. The grips live in wrap coordinates.
    const grips = document.createDocumentFragment();
    for (const c of this.cols) {
      const g = document.createElement("i");
      g.className = "grip grip-col";
      g.dataset.col = String(c);
      g.style.left = `${colLeft(s, this.cols, c) + colPx(s, c) - 3}px`;
      g.style.top = "0px";
      g.style.height = `${HEAD_H}px`;
      grips.append(g);
    }
    for (let r = 0; r < this.rowCount; r++) {
      if (s.hiddenRows?.has(r) || s.filteredRows?.has(r)) continue;
      const g = document.createElement("i");
      g.className = "grip grip-row";
      g.dataset.row = String(r);
      g.style.left = "0px";
      g.style.top = `${headTop(s, r) + rowPx(s, r) - 3}px`;
      g.style.width = `${HEAD_W}px`;
      grips.append(g);
    }
    this.grips.replaceChildren(grips);

    this.wrap.replaceChildren(table, this.pics, this.grips, this.frame, this.fillBox);
    this.paintSelection();
    this.hooks.selected();
  }

  // ------------------------------------------------------- conditional formats

  /** Cell-ref → conditional style overlay, rebuilt per render from the sheet's rules. */
  private cfMap = new Map<string, CondStyle & { bar?: string; icon?: { color: string; glyph: string } }>();

  private buildCond(): void {
    this.cfMap = new Map();
    const s = this.sheet;
    for (const cf of s.cond) {
      let box: Rect;
      try {
        box = parseRange(cf.ref);
      } catch {
        continue;
      }
      for (const rule of [...cf.rules].sort((a, b) => a.priority - b.priority)) {
        if (rule.type === "colorScale" || rule.type === "dataBar") this.applyGradient(rule, box);
        else if (rule.type === "cellIs") this.applyCellIs(rule, box);
        else if (rule.type === "iconSet") this.applyIconSet(rule, box);
        else if (rule.type === "top10" || rule.type === "aboveAverage" || rule.type === "belowAverage" || rule.type === "containsText" || rule.type === "notContainsText") this.applyRuleStyle(rule, box);
      }
    }
  }

  /** Color scales and data bars interpolate over the numeric values of their range. */
  private applyGradient(rule: CondRule, box: Rect): void {
    const numbers: number[] = [];
    for (const at of eachCell(box)) {
      const v = this.wb.value(this.sheet.name, refOf(at.c, at.r));
      if (typeof v === "number") numbers.push(v);
    }
    if (!numbers.length) return;
    const min = Math.min(...numbers);
    const max = Math.max(...numbers);
    for (const at of eachCell(box)) {
      const ref = refOf(at.c, at.r);
      const v = this.wb.value(this.sheet.name, ref);
      if (typeof v !== "number") continue;
      const t = max === min ? 0.5 : (v - min) / (max - min);
      const entry: CondStyle & { bar?: string } = this.cfMap.get(ref) ?? {};
      if (rule.type === "colorScale") entry.bgColor = scaleColorAt(rule.colors ?? [], t);
      else {
        const pct = Math.round(t * 100);
        entry.bar = `linear-gradient(to right, #${rule.colors?.[0] ?? "638EC6"} ${pct}%, transparent ${pct}%)`;
      }
      this.cfMap.set(ref, entry);
    }
  }

  /** cellIs rules compare each value against the stored operand, styled by its dxf. */
  private applyCellIs(rule: CondRule, box: Rect): void {
    if (rule.formula === undefined) return;
    const num = Number(rule.formula);
    const operand: number | string = rule.formula.trim() !== "" && Number.isFinite(num) ? num : rule.formula;
    const dxf = rule.dxfId !== undefined ? this.wb.dxfOf(rule.dxfId) : undefined;
    for (const at of eachCell(box)) {
      const ref = refOf(at.c, at.r);
      const v = this.wb.value(this.sheet.name, ref);
      if (v === null || v === "") continue;
      if (!cfCompare(v, rule.op ?? "greaterThan", operand)) continue;
      const entry: CondStyle & { bar?: string } = this.cfMap.get(ref) ?? {};
      if (dxf?.bgColor) entry.bgColor = dxf.bgColor;
      else entry.bgColor = "FFEB9C"; // the classic light-yellow a bare rule implies
      if (dxf?.fontColor) entry.fontColor = dxf.fontColor;
      if (dxf?.bold) entry.bold = true;
      if (dxf?.italic) entry.italic = true;
      if (dxf?.strike) entry.strike = true;
      if (dxf?.underline) entry.underline = true;
      this.cfMap.set(ref, entry);
    }
  }

  private applyCond(td: HTMLElement, ref: string): void {
    const cf = this.cfMap.get(ref);
    if (!cf) return;
    if (cf.bar) td.style.background = cf.bar;
    else if (cf.bgColor) td.style.background = `#${cf.bgColor}`;
    if (cf.fontColor) td.style.color = `#${cf.fontColor}`;
    if (cf.bold) td.style.fontWeight = "600";
    if (cf.italic) td.style.fontStyle = "italic";
    if (cf.strike) td.style.textDecoration = "line-through";
    if (cf.underline) td.style.textDecoration = "underline";
    if (cf.icon) {
      const dot = document.createElement("span");
      dot.style.color = cf.icon.color;
      dot.textContent = `${cf.icon.glyph} `;
      td.prepend(dot);
    }
  }

  /** Icon sets draw a ranked marker from the rule's own thresholds. */
  private applyIconSet(rule: CondRule, box: Rect): void {
    const numbers: number[] = [];
    for (const at of eachCell(box)) {
      const v = this.wb.value(this.sheet.name, refOf(at.c, at.r));
      if (typeof v === "number") numbers.push(v);
    }
    if (!numbers.length) return;
    const min = Math.min(...numbers);
    const max = Math.max(...numbers);
    const raw = rule.raw ?? "";
    const setName = /iconSet="([^"]+)"/.exec(String(raw))?.[1] ?? "3TrafficLights1";
    const marks: { color: string; glyph: string }[] = setName.includes("Arrow")
      ? [{ color: "#d6604d", glyph: "↓" }, { color: "#e6b800", glyph: "→" }, { color: "#4d9e4d", glyph: "↑" }]
      : setName.includes("Symbols") || setName.includes("Sign")
        ? [{ color: "#d6604d", glyph: "✗" }, { color: "#e6b800", glyph: "!" }, { color: "#4d9e4d", glyph: "✓" }]
        : [{ color: "#d6604d", glyph: "●" }, { color: "#e6b800", glyph: "●" }, { color: "#4d9e4d", glyph: "●" }];
    const buckets = setName.startsWith("4") || setName.startsWith("5") ? Math.min(Number(setName[0]), 5) : 3;
    for (const at of eachCell(box)) {
      const ref = refOf(at.c, at.r);
      const v = this.wb.value(this.sheet.name, ref);
      if (typeof v !== "number") continue;
      const t = max === min ? 0.5 : (v - min) / (max - min);
      const idx = Math.min(marks.length - 1, Math.floor(t * buckets));
      const entry = this.cfMap.get(ref) ?? {};
      entry.icon = marks[idx];
      this.cfMap.set(ref, entry);
    }
  }

  /** Rank, average and text rules share one shape: pick cells, then style them. */
  private applyRuleStyle(rule: CondRule, box: Rect): void {
    const raw = String(rule.raw ?? "");
    const all: { ref: string; v: number }[] = [];
    for (const at of eachCell(box)) {
      const v = this.wb.value(this.sheet.name, refOf(at.c, at.r));
      if (typeof v === "number") all.push({ ref: refOf(at.c, at.r), v });
    }
    const styleOf = (ref: string): CondStyle & { bar?: string; icon?: { color: string; glyph: string } } => {
      const dxf = rule.dxfId !== undefined ? this.wb.dxfOf(rule.dxfId) : undefined;
      const entry = this.cfMap.get(ref) ?? {};
      entry.bgColor = dxf?.bgColor ?? "FFEB9C";
      if (dxf?.fontColor) entry.fontColor = dxf.fontColor;
      this.cfMap.set(ref, entry);
      return entry;
    };
    const pick = (refs: string[]): void => {
      for (const ref of refs) styleOf(ref);
    };
    if (rule.type === "top10") {
      const rank = Math.max(1, Number(/rank="(\d+)"/.exec(raw)?.[1] ?? 10));
      const percent = /percent="1"/.test(raw);
      const bottom = /bottom="1"/.test(raw);
      const sorted = [...all].sort((a, b) => (bottom ? a.v - b.v : b.v - a.v));
      const n = percent ? Math.max(1, Math.round((all.length * rank) / 100)) : rank;
      pick(sorted.slice(0, Math.min(n, all.length)).map((x) => x.ref));
      return;
    }
    if (rule.type === "aboveAverage" || rule.type === "belowAverage") {
      const mean = all.reduce((a, x) => a + x.v, 0) / Math.max(1, all.length);
      pick(all.filter((x) => (rule.type === "aboveAverage" ? x.v > mean : x.v < mean)).map((x) => x.ref));
      return;
    }
    // containsText / notContainsText: the needle rides the rule's formula.
    const needle = /"([^"]*)"/.exec(raw)?.[1] ?? "";
    if (!needle) return;
    const refs: string[] = [];
    for (const at of eachCell(box)) {
      const ref = refOf(at.c, at.r);
      const v = this.wb.value(this.sheet.name, ref);
      const text = v === null ? "" : String(v);
      const has = text.toLowerCase().includes(needle.toLowerCase());
      if (rule.type === "containsText" ? has : !has && text !== "") refs.push(ref);
    }
    pick(refs);
  }

  /** Inserts a chart from the selection: first row names, first column categories. */
  insertChart(kind: ChartModel["kind"]): void {
    if (this.sel.c1 === this.sel.c2 && this.sel.r1 === this.sel.r2) return;
    const anchor = { c: this.sel.c2 + 1, r: this.sel.r1 };
    this.hooks.edited();
    this.wb.addChart(kind, this.sheet.name, this.sel, anchor, "");
    this.selectedChart = this.sheet.charts.length - 1;
    this.render();
  }

  selectChart(index: number): void {
    this.selectedChart = index;
    for (const el of this.pics.querySelectorAll<HTMLElement>(".sheet-chart")) el.classList.toggle("chart-selected", Number(el.dataset.idx) === index);
    this.hooks.selected();
  }

  private startChartDrag(index: number, element: HTMLElement, e: PointerEvent): void {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    this.cancelChartDrag();
    this.selectChart(index);
    const grip = (e.target as Element)?.closest?.<HTMLElement>(".chart-resize");
    (grip || element).focus({ preventScroll: true });
    try { element.setPointerCapture(e.pointerId); } catch { /* synthetic smoke pointer */ }
    const original = imageBoxPx(this.sheet, this.sheet.charts[index]);
    this.chartDrag = { index, element, pointerId: e.pointerId, startX: e.clientX, startY: e.clientY, scale: this.wrap.offsetWidth ? this.wrap.getBoundingClientRect().width / this.wrap.offsetWidth || 1 : 1, resize: !!grip, original, preview: { ...original } };
  }

  private moveChartDrag(e: PointerEvent): void {
    const drag = this.chartDrag;
    if (!drag || e.pointerId !== drag.pointerId) return;
    e.stopPropagation();
    const dx = (e.clientX - drag.startX) / drag.scale, dy = (e.clientY - drag.startY) / drag.scale;
    drag.preview = drag.resize
      ? { ...drag.original, w: Math.max(160, drag.original.w + dx), h: Math.max(100, drag.original.h + dy) }
      : { ...drag.original, x: Math.max(0, drag.original.x + dx), y: Math.max(0, drag.original.y + dy) };
    Object.assign(drag.element.style, { left: `${drag.preview.x}px`, top: `${drag.preview.y}px`, width: `${drag.preview.w}px`, height: `${drag.preview.h}px` });
  }

  private endChartDrag(e: PointerEvent): void {
    const drag = this.chartDrag;
    if (!drag || e.pointerId !== drag.pointerId) return;
    this.moveChartDrag(e);
    this.chartDrag = null;
    try { drag.element.releasePointerCapture(e.pointerId); } catch { /* no captured pointer */ }
    const changed = Object.keys(drag.original).some((k) => Math.abs(drag.preview[k as keyof typeof drag.preview] - drag.original[k as keyof typeof drag.original]) >= 1);
    if (!changed) {
      Object.assign(drag.element.style, { left: `${drag.original.x}px`, top: `${drag.original.y}px`, width: `${drag.original.w}px`, height: `${drag.original.h}px` });
      return;
    }
    this.hooks.edited();
    this.wb.setChartBox(this.sheet.name, drag.index, drag.preview);
    this.render();
  }

  private cancelChartDrag(): void {
    const drag = this.chartDrag;
    this.chartDrag = null;
    if (!drag) return;
    try { drag.element.releasePointerCapture(drag.pointerId); } catch { /* no captured pointer */ }
    Object.assign(drag.element.style, { left: `${drag.original.x}px`, top: `${drag.original.y}px`, width: `${drag.original.w}px`, height: `${drag.original.h}px` });
  }

  private chartKey(index: number, e: KeyboardEvent): void {
    if (e.key === "Escape") { e.stopPropagation(); this.cancelChartDrag(); return; }
    if (e.key === "Enter") { e.stopPropagation(); this.selectChart(index); this.editSelectedChart(); return; }
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault(); e.stopPropagation();
      this.hooks.edited(); this.wb.removeChart(this.sheet.name, index); this.selectedChart = -1; this.render(); return;
    }
    if (!["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(e.key)) return;
    e.preventDefault(); e.stopPropagation();
    const step = e.shiftKey ? 10 : 1, box = imageBoxPx(this.sheet, this.sheet.charts[index]);
    const dx = e.key === "ArrowLeft" ? -step : e.key === "ArrowRight" ? step : 0;
    const dy = e.key === "ArrowUp" ? -step : e.key === "ArrowDown" ? step : 0;
    const resize = !!(e.target as Element).closest(".chart-resize");
    const next = resize ? { ...box, w: Math.max(160, box.w + dx), h: Math.max(100, box.h + dy) } : { ...box, x: Math.max(0, box.x + dx), y: Math.max(0, box.y + dy) };
    if (next.x === box.x && next.y === box.y && next.w === box.w && next.h === box.h) return;
    this.hooks.edited(); this.wb.setChartBox(this.sheet.name, index, next); this.render();
    const el = this.pics.querySelector<HTMLElement>(`.sheet-chart[data-idx="${index}"]`);
    (resize ? el?.querySelector<HTMLElement>(".chart-resize") : el)?.focus({ preventScroll: true });
  }

  editSelectedChart(): void {
    this.cancelChartDrag();
    if (this.selectedChart < 0 || !this.sheet.charts[this.selectedChart]) return;
    const index = this.selectedChart;
    openChartEditor(this.wrap, this.wb, this.sheet.name, index, (next) => {
      this.hooks.edited();
      this.wb.updateChart(this.sheet.name, index, next);
      this.render();
    }, () => {
      this.hooks.edited();
      this.wb.removeChart(this.sheet.name, index);
      this.selectedChart = -1;
      this.render();
    });
  }

  /** Applies an editor cellIs rule (with its own dxf) over the selection. */
  setCellIsUI(op: string, formula: string): void {
    this.hooks.edited();
    this.wb.setCondCellIs(this.sheet.name, this.sel, op, formula);
    this.render();
  }

  setColorScale(): void {    this.hooks.edited();
    this.wb.setCond(this.sheet.name, this.sel, { type: "colorScale", colors: ["F8696B", "FFEB84", "63BE7B"] });
    this.render();
  }

  setDataBar(): void {
    this.hooks.edited();
    this.wb.setCond(this.sheet.name, this.sel, { type: "dataBar", color: "638EC6" });
    this.render();
  }

  // ------------------------------------------------------------------- notes

  hasComment(): boolean {
    return this.sheet.comments.some((c) => c.ref === refOf(this.active.c, this.active.r));
  }

  activeComment(): string | undefined {
    return this.sheet.comments.find((c) => c.ref === refOf(this.active.c, this.active.r))?.text;
  }

  setComment(text: string): void {
    this.hooks.edited();
    this.wb.setComment(this.sheet.name, refOf(this.active.c, this.active.r), text);
    this.render();
  }

  removeComment(): void {
    this.setComment("");
  }

  clearCond(): void {
    this.hooks.edited();
    this.wb.clearCond(this.sheet.name, this.sel);
    this.render();
  }

  private paint(td: HTMLElement, cell: Cell): void {    const name = this.sheet.name;
    const value = this.wb.value(name, cell.ref);
    if (cell.formula) td.classList.add("fx");
    if (isError(value)) {
      td.textContent = String(value);
      td.classList.add("err");
    } else if (typeof value === "number") {
      td.textContent = this.wb.display(name, cell.ref).text;
      td.classList.add("num");
    } else if (typeof value === "boolean") td.textContent = value ? "TRUE" : "FALSE";
    else td.textContent = value === null ? "" : String(value);

    if (cell.runs?.length) {
      td.replaceChildren();
      for (const run of cell.runs) {
        const s = document.createElement("span");
        s.textContent = run.text;
        if (run.bold) s.style.fontWeight = "600";
        if (run.italic) s.style.fontStyle = "italic";
        if (run.underline) s.style.textDecoration = "underline";
        if (run.color) s.style.color = css(run.color) ?? "";
        if (run.size) s.style.fontSize = `${run.size}px`;
        td.append(s);
      }
    }
    applyStyle(td, this.wb.styleOf(cell));
  }

  /** Selection tint plus one frame around the whole range, the way a grid marks it. */
  private paintSelection(): void {
    for (const td of this.wrap.querySelectorAll<HTMLTableCellElement>("td")) {
      const c = Number(td.dataset.col);
      const r = Number(td.dataset.row);
      td.classList.toggle("sel", c >= this.sel.c1 && c <= this.sel.c2 && r >= this.sel.r1 && r <= this.sel.r2);
      td.classList.toggle("anchor", c === this.active.c && r === this.active.r);
    }
    const from = this.cellEl(this.sel.c1, this.sel.r1);
    const to = this.cellEl(this.sel.c2, this.sel.r2);
    if (!from || !to) {
      this.frame.style.display = "none";
      this.fillBox.style.display = "none";
      return;
    }
    const a = from.getBoundingClientRect();
    const b = to.getBoundingClientRect();
    const o = this.wrap.getBoundingClientRect();
    // Rects are in zoomed screen pixels; children of the zoomed wrap lay out in
    // unzoomed px, so divide the offsets back before positioning.
    const z = this.zoom || 1;
    this.frame.style.display = "block";
    this.frame.style.left = `${(a.left - o.left) / z}px`;
    this.frame.style.top = `${(a.top - o.top) / z}px`;
    this.frame.style.width = `${(b.right - a.left) / z}px`;
    this.frame.style.height = `${(b.bottom - a.top) / z}px`;
    // The fill handle rides the bottom-right corner of the selection.
    this.fillBox.style.display = "block";
    this.fillBox.style.left = `${(b.right - o.left) / z - 4}px`;
    this.fillBox.style.top = `${(b.bottom - o.top) / z - 4}px`;
  }

  cellEl(c: number, r: number): HTMLElement | null {
    return this.wrap.querySelector<HTMLElement>(`td[data-ref="${refOf(c, r)}"]`);
  }

  // ---------------------------------------------------------------- selection

  select(box: Rect, active?: { c: number; r: number }): void {
    this.sel = box;
    this.active = active ?? { c: box.c1, r: box.r1 };
    this.paintSelection();
    this.hooks.selected();
  }

  selectAll(): void {
    this.select(rect(this.cols[0] ?? 0, 0, this.cols[this.cols.length - 1] ?? 0, this.rowCount - 1));
  }

  toggleFilter(): void {
    if (this.sheet.autoFilter) { this.hooks.edited(); this.wb.removeAutoFilter(this.sheet.name); this.render(); return; }
    const b = this.wb.bounds(this.sheet.name);
    let box = this.sel;
    const occupied = (c: number, r: number): boolean => { const cell = this.sheet.cells.get(refOf(c, r)); return !!cell && (cell.formula !== undefined || cell.constant !== null && cell.constant !== ""); };
    if (box.c1 === box.c2 && box.r1 === box.r2 && occupied(this.active.c, this.active.r)) box = filterDataRegion(this.active, rect(0, 0, b.maxCol, b.maxRow), occupied);
    this.hooks.edited(); this.wb.enableAutoFilter(this.sheet.name, `${refOf(box.c1, box.r1)}:${refOf(box.c2, box.r2)}`); this.render();
  }

  clearFilters(): void {
    if (!this.sheet.autoFilter?.columns.length) return;
    this.hooks.edited(); this.wb.clearFilterCriteria(this.sheet.name); this.render();
  }

  private editFilterColumn(id: number): void {
    openFilterEditor(this.wrap, this.wb, this.sheet.name, id, (rule) => {
      if (rule) validateFilterRule(rule);
      if (this.sheet.autoFilter?.unsupported) throw new Error(this.sheet.autoFilter.unsupported);
      this.hooks.edited(); this.wb.setFilterColumn(this.sheet.name, id, rule);
      this.wb.refreshFilters();
      if (this.sheet.filteredRows?.has(this.active.r)) { const b = filterRange(this.sheet.autoFilter!.ref); this.active = { c: b.c1 + id, r: b.r1 }; this.sel = rect(this.active.c, this.active.r, this.active.c, this.active.r); }
      this.render();
    }, (desc) => {
      const b = filterRange(this.sheet.autoFilter!.ref);
      if (b.r1 === b.r2) return;
      this.hooks.edited(); this.wb.sort(this.sheet.name, rect(b.c1, b.r1 + 1, b.c2, b.r2), b.c1 + id, desc); this.render();
    });
  }

  selectRange(spec: string): void {
    const box = parseRange(spec);
    this.select(box, { c: box.c1, r: box.r1 });
    this.scrollTo(box.c1, box.r1);
  }

  colIndex(c: number): number {
    return Math.max(0, this.cols.indexOf(c));
  }

  stepBy(dc: number, dr: number): void {
    this.move(dc, dr, false);
  }

  /** Recompute the selection frame, e.g. after the window resizes. */
  refreshFrame(): void {
    this.paintSelection();
  }

  private move(dc: number, dr: number, extend: boolean): void {
    const i = Math.max(0, Math.min(this.cols.length - 1, this.colIndex(this.active.c) + dc));
    const visible = Array.from({ length: this.rowCount }, (_, r) => r).filter((r) => !this.sheet.hiddenRows?.has(r) && !this.sheet.filteredRows?.has(r));
    let index = visible.indexOf(this.active.r);
    if (index < 0) index = Math.max(0, visible.findIndex((r) => r >= this.active.r));
    const r = visible[Math.max(0, Math.min(visible.length - 1, index + dr))] ?? 0;
    const at = { c: this.cols[i] ?? 0, r };
    this.active = at;
    this.sel = extend ? rect(this.sel.c1, this.sel.r1, at.c, at.r) : rect(at.c, r, at.c, r);
    this.cellEl(at.c, r)?.scrollIntoView({ block: "nearest", inline: "nearest" });
    this.paintSelection();
    this.hooks.selected();
  }

  scrollTo(c: number, r: number): void {
    this.cellEl(c, r)?.scrollIntoView({ block: "center", inline: "center" });
  }

  focusGrid(): void {
    this.wrap.focus({ preventScroll: true });
  }

  // ------------------------------------------------------------------ editing

  beginEdit(seed?: string): void {
    this.commitEdit();
    const td = this.cellEl(this.active.c, this.active.r);
    if (!td) return;
    const ed = document.createElement("div");
    ed.className = "cell-editor";
    ed.contentEditable = "true";
    ed.spellcheck = false;
    ed.textContent = seed ?? this.wb.editOf(this.sheet.name, refOf(this.active.c, this.active.r));
    const a = td.getBoundingClientRect();
    const o = this.wrap.getBoundingClientRect();
    const z = this.zoom || 1;
    ed.style.left = `${(a.left - o.left) / z}px`;
    ed.style.top = `${(a.top - o.top) / z}px`;
    ed.style.minWidth = `${a.width / z}px`;
    ed.style.height = `${a.height / z}px`;
    ed.style.fontSize = `${this.activeStyle().fontSize}px`;
    this.wrap.append(ed);
    this.editor = ed;
    placeCaretEnd(ed);
    ed.addEventListener("keydown", (e) => {
      const ke = e as KeyboardEvent;
      if (ke.isComposing) return;
      if (ke.key === "Enter" || ke.key === "Tab") {
        ke.preventDefault();
        const dc = ke.key === "Tab" ? (ke.shiftKey ? -1 : 1) : 0;
        const dr = ke.key === "Enter" ? (ke.shiftKey ? -1 : 1) : 0;
        this.commitEdit();
        this.move(dc, dr, false);
      } else if (ke.key === "Escape") {
        ke.preventDefault();
        this.cancelEdit();
      }
    });
    ed.addEventListener("blur", () => this.commitEdit());
    ed.focus();
  }

  /** Commit without waiting for a blur cycle, so typing over a cell stays in step. */
  commitEdit(): void {
    const ed = this.editor;
    if (!ed) return;
    this.editor = null;
    const text = ed.textContent ?? "";
    ed.remove();
    this.write(this.active.c, this.active.r, text);
  }

  cancelEdit(): void {
    this.editor?.remove();
    this.editor = null;
    this.focusGrid();
  }

  write(c: number, r: number, text: string): void {
    const ref = refOf(c, r);
    if (this.wb.editOf(this.sheet.name, ref) === text) return;
    this.hooks.edited();
    this.wb.setCell(this.sheet.name, ref, text);
    this.render();
  }

  // ----------------------------------------------------------------- commands

  activeStyle(): CellStyle {
    return this.wb.styleAt(this.sheet.name, this.active.c, this.active.r);
  }

  private patch(p: Partial<CellStyle>): void {
    this.hooks.edited();
    this.wb.setStyle(this.sheet.name, this.sel, p);
    this.render();
  }

  toggle(key: "bold" | "italic" | "underline" | "wrap"): void {
    this.patch({ [key]: !this.activeStyle()[key] } as Partial<CellStyle>);
  }

  setFontSize(px: number): void {
    this.patch({ fontSize: px });
  }

  setFontColor(color: string): void {
    this.patch({ fontColor: color.replace(/^#/, "").toUpperCase() });
  }

  setFill(color: string): void {
    this.patch({ bgColor: color.replace(/^#/, "").toUpperCase() });
  }

  setAlign(align: string): void {
    this.patch({ align });
  }

  setValign(valign: string): void {
    this.patch({ valign });
  }

  setNumFmt(numFmt: string): void {
    this.patch({ numFmt });
  }

  /** Add or drop one decimal place on whatever format the cell already carries. */
  stepDecimals(dir: 1 | -1): void {
    const fmt = this.activeStyle().numFmt;
    if (fmt === "General" || fmt === "0") return this.setNumFmt(dir > 0 ? "0.0" : "0");
    const m = /^([#,0]+)(?:\.(0*))([%;]?)$/.exec(fmt);
    if (!m) return;
    const places = Math.max(0, (m[2]?.length ?? 0) + dir);
    this.setNumFmt(`${m[1]}${places ? `.${"0".repeat(places)}` : ""}${m[3] ?? ""}`);
  }

  setBorder(kind: "box" | "all" | "none"): void {
    this.hooks.edited();
    for (const at of eachCell(this.sel)) {
      const cur = this.wb.styleAt(this.sheet.name, at.c, at.r).border ?? { top: false, right: false, bottom: false, left: false };
      const next = { ...cur };
      if (kind === "none") next.top = next.right = next.bottom = next.left = false;
      else if (kind === "all") next.top = next.right = next.bottom = next.left = true;
      else {
        if (at.r === this.sel.r1) next.top = true;
        if (at.r === this.sel.r2) next.bottom = true;
        if (at.c === this.sel.c1) next.left = true;
        if (at.c === this.sel.c2) next.right = true;
      }
      this.wb.setStyle(this.sheet.name, rect(at.c, at.r, at.c, at.r), { border: next });
    }
    this.render();
  }

  mergeSelection(): void {
    if (this.sel.c1 === this.sel.c2 && this.sel.r1 === this.sel.r2) return;
    this.hooks.edited();
    this.wb.merge(this.sheet.name, this.sel);
    this.render();
  }

  unmergeSelection(): void {
    this.hooks.edited();
    this.wb.unmerge(this.sheet.name, this.sel);
    this.render();
  }

  clear(what: "content" | "format" | "all"): void {
    this.hooks.edited();
    this.wb.clear(this.sheet.name, this.sel, what);
    this.render();
  }

  /** Lines are inserted or removed at the top edge of the selection, as many as it covers. */
  line(axis: Axis, insert: boolean): void {
    const at = axis === "row" ? this.sel.r1 : this.sel.c1;
    const n = axis === "row" ? this.sel.r2 - this.sel.r1 + 1 : this.sel.c2 - this.sel.c1 + 1;
    this.hooks.edited();
    if (insert) (axis === "row" ? this.wb.insertRows : this.wb.insertCols).call(this.wb, this.sheet.name, at, n);
    else (axis === "row" ? this.wb.deleteRows : this.wb.deleteCols).call(this.wb, this.sheet.name, at, n);
    this.render();
  }

  sortSelection(desc: boolean): void {
    this.hooks.edited();
    this.wb.sort(this.sheet.name, this.sel, this.active.c, desc);
    this.render();
  }

  freezeHere(): void {
    this.hooks.edited();
    this.wb.setFreeze(this.sheet.name, this.active.r, this.colIndex(this.active.c));
    this.render();
  }

  unfreeze(): void {
    this.hooks.edited();
    this.wb.setFreeze(this.sheet.name, 0, 0);
    this.render();
  }

  setColWidth(c: number, chars: number): void {
    this.hooks.edited();
    this.wb.setColWidth(this.sheet.name, c, chars);
    this.render();
  }

  setRowHeight(r: number, points: number): void {
    this.hooks.edited();
    this.wb.setRowHeight(this.sheet.name, r, points);
    this.render();
  }

  /** Applies the zoom the toolbar drives; pointer math divides client deltas by it. */
  setZoom(z: number): void {
    this.zoom = Math.min(2, Math.max(0.5, Math.round(z * 100) / 100));
    this.wrap.style.zoom = String(this.zoom);
    this.hooks.selected();
  }

  /** Drops a picture anchored at the active cell; bytes ride as a data URL until save. */
  insertPicture(src: string, width: number, height: number, name = "图片"): void {
    this.hooks.edited();
    this.wb.addImage(this.sheet.name, { c: this.active.c, r: this.active.r }, src, width, height, name);
    this.render();
  }

  /** Fill-handle release: project the origin selection into the dragged-out band. */
  applyFill(): void {
    const origin = this.filling;
    this.filling = null;
    if (!origin) return;
    const grown = this.sel;
    this.sel = origin;
    if (grown.c1 === origin.c1 && grown.c2 === origin.c2 && grown.r1 === origin.r1 && grown.r2 === origin.r2) return;
    this.hooks.edited();
    this.wb.fill(this.sheet.name, origin, grown);
    this.render();
  }

  hideCol(): void {
    this.hooks.edited();
    for (let c = this.sel.c1; c <= this.sel.c2; c++) this.wb.hideCol(this.sheet.name, c, true);
    this.render();
  }

  // ---------------------------------------------------------------- clipboard

  /** Copy hands over what was typed, so a formula survives a copy/paste round trip. */
  copyText(box = this.sel): string {
    const lines: string[] = [];
    for (let r = box.r1; r <= box.r2; r++) {
      const row: string[] = [];
      for (let c = box.c1; c <= box.c2; c++) row.push(this.wb.editOf(this.sheet.name, refOf(c, r)));
      lines.push(row.join("\t"));
    }
    return lines.join("\n");
  }

  paste(text: string): void {
    this.hooks.edited();
    const area = this.wb.pasteText(this.sheet.name, this.active.c, this.active.r, text);
    this.render();
    if (area) this.select(area, { c: area.c1, r: area.r1 });
  }

  cut(): void {
    this.clear("content");
  }

  // -------------------------------------------------------------- interaction

  pointerDown(c: number, r: number, extend: boolean): void {
    this.selectChart(-1);
    if (extend) {
      this.sel = rect(this.sel.c1, this.sel.r1, c, r);
      this.active = { c, r };
      this.paintSelection();
      this.hooks.selected();
    } else this.select(rect(c, r, c, r), { c, r });
    this.dragging = true;
    this.focusGrid();
  }

  dragTo(c: number, r: number): void {
    if (!this.dragging) return;
    this.sel = rect(this.sel.c1, this.sel.r1, c, r);
    this.active = { c, r };
    this.paintSelection();
    this.hooks.selected();
  }

  pointerUp(): void {
    this.dragging = false;
  }

  // ------------------------------------------------------------- fill handle

  startFill(): void {
    this.filling = { ...this.sel };
  }

  /** The preview band is the origin selection grown to the hovered cell. */
  fillTo(c: number, r: number): void {
    const from = this.filling;
    if (!from) return;
    this.sel = rect(Math.min(from.c1, c), Math.min(from.r1, r), Math.max(from.c2, c), Math.max(from.r2, r));
    this.paintSelection();
    this.hooks.selected();
  }

  // ------------------------------------------------------------- line sizing

  startSize(axis: Axis, line: number, clientPos: number): void {
    this.sizing = { axis, line, start: clientPos, origin: axis === "col" ? colPx(this.sheet, line) : rowPx(this.sheet, line) };
  }

  /** Live-resizes; the undo entry was taken when the grip went down. */
  sizeTo(clientPos: number): void {
    const s = this.sizing;
    if (!s) return;
    const px = Math.max(s.axis === "col" ? 24 : 14, s.origin + (clientPos - s.start) / this.zoom);
    if (s.axis === "col") this.wb.setColWidth(this.sheet.name, s.line, px / 7.2);
    else this.wb.setRowHeight(this.sheet.name, s.line, (px * 3) / 4);
    this.render();
  }

  endSize(): void {
    if (!this.sizing) return;
    this.sizing = null;
    this.hooks.selected();
  }

  wholeColumn(c: number): void {
    this.select(rect(c, 0, c, this.rowCount - 1), { c, r: 0 });
  }

  wholeRow(r: number): void {
    this.select(rect(this.cols[0] ?? 0, r, this.cols[this.cols.length - 1] ?? 0, r), { c: this.cols[0] ?? 0, r });
  }

  onKeyDown(e: KeyboardEvent): void {
    if (this.editor || e.isComposing) return;
    const meta = e.metaKey || e.ctrlKey;
    const step = (dc: number, dr: number): void => this.move(dc, dr, e.shiftKey);
    const k = e.key;

    if (meta) {
      if (k === "a" || k === "A") {
        e.preventDefault();
        return this.selectAll();
      }
      if (k === "Enter") {
        e.preventDefault();
        return this.beginEdit();
      }
      // ⌘C/X/V ride the copy/cut/paste events; ⌘Z/F belong to the shell.
      if (k !== "Home" && k !== "End") return;
    }
    if (e.altKey) return;

    switch (k) {
      case "ArrowUp":
        e.preventDefault();
        return step(0, -1);
      case "ArrowDown":
        e.preventDefault();
        return step(0, 1);
      case "ArrowLeft":
        e.preventDefault();
        return step(-1, 0);
      case "ArrowRight":
        e.preventDefault();
        return step(1, 0);
      case "Tab":
        e.preventDefault();
        return step(e.shiftKey ? -1 : 1, 0);
      case "Enter":
        e.preventDefault();
        return step(0, 1);
      case "F2":
        e.preventDefault();
        return this.beginEdit();
      case "Backspace":
      case "Delete":
        e.preventDefault();
        return this.clear("content");
      case "Home": {
        e.preventDefault();
        const first = this.cols[0] ?? 0;
        return meta ? this.select(rect(first, 0, first, 0)) : step(-this.cols.length, 0);
      }
      case "End": {
        e.preventDefault();
        const last = this.cols[this.cols.length - 1] ?? 0;
        return this.select(rect(last, this.active.r, last, this.active.r));
      }
      case "PageDown":
      case "PageUp":
        e.preventDefault();
        return step(0, (k === "PageDown" ? 1 : -1) * (this.rowCount - 6));
      default:
        if (k.length === 1) {
          e.preventDefault();
          this.beginEdit(k);
        }
    }
  }
}

function colEl(width: number): HTMLElement {
  const col = document.createElement("col");
  col.style.width = `${width}px`;
  return col;
}

function applyStyle(td: HTMLElement, st: CellStyle): void {
  if (st.bold) td.style.fontWeight = "600";
  if (st.italic) td.style.fontStyle = "italic";
  if (st.underline) td.style.textDecoration = "underline";
  if (st.fontSize && st.fontSize !== 11) td.style.fontSize = `${st.fontSize}px`;
  if (st.fontColor) td.style.color = css(st.fontColor) ?? "";
  if (st.bgColor) td.style.background = css(st.bgColor) ?? "";
  if (st.valign) td.style.verticalAlign = st.valign === "center" ? "middle" : st.valign;
  if (st.wrap) td.classList.add("wrap");
  const b = st.border;
  if (b) {
    const line = "1px solid #40464f";
    if (b.top) td.style.borderTop = line;
    if (b.bottom) td.style.borderBottom = line;
    if (b.left) td.style.borderLeft = line;
    if (b.right) td.style.borderRight = line;
  }
  // Numbers sit right unless the format says otherwise; explicit alignment wins.
  if (st.align === "left") td.style.textAlign = "left";
  else if (st.align === "center" || st.align === "centerContinuous") td.style.textAlign = "center";
  else if (st.align) td.style.textAlign = "right";
  else if (td.classList.contains("num")) td.style.textAlign = "right";
}

const headTop = (s: Sheet, r: number): number => {
  let y = HEAD_H;
  for (let i = 0; i < r; i++) y += rowPx(s, i);
  return y;
};

const colLeft = (s: Sheet, cols: number[], c: number): number => {
  let x = HEAD_W;
  for (const k of cols) {
    if (k >= c) break;
    x += colPx(s, k);
  }
  return x;
};

function placeCaretEnd(el: HTMLElement): void {
  const range = document.createRange();
  range.selectNodeContents(el);
  range.collapse(false);
  const sel = window.getSelection();
  sel?.removeAllRanges();
  sel?.addRange(range);
}

/** Mouse, keyboard and clipboard wiring lives here so the class stays about state. */
function capture(el: HTMLElement, pointerId: number): void {
  // Synthetic events carry pointerId 0 and no live pointer; capture is a nicety.
  try {
    el.setPointerCapture(pointerId);
  } catch {
    /* no pointer to capture */
  }
}

function wireSheet(view: SheetView): void {
  const wrap = view.wrap;
  const cellAt = (target: EventTarget | null): { c: number; r: number } | null => {
    const el = (target as HTMLElement | null)?.closest?.<HTMLElement>("td[data-col]");
    if (!el) return null;
    return { c: Number(el.dataset.col), r: Number(el.dataset.row) };
  };

  wrap.addEventListener("pointerdown", (e) => {
    const ev = e as PointerEvent;
    if (ev.button === 2) return;
    if ((ev.target as HTMLElement).closest?.(".fill-box")) {
      ev.preventDefault();
      capture(wrap, ev.pointerId);
      view.startFill();
      return;
    }
    const at = cellAt(ev.target);
    if (at) {
      ev.preventDefault();
      capture(wrap, ev.pointerId);
      view.pointerDown(at.c, at.r, ev.shiftKey);
      return;
    }
    const head = (ev.target as HTMLElement)?.closest?.<HTMLElement>("th[data-col], th[data-row]");
    if (!head) return;
    ev.preventDefault();
    if (head.dataset.col !== undefined) view.wholeColumn(Number(head.dataset.col));
    else view.wholeRow(Number(head.dataset.row));
    view.focusGrid();
  });

  // The fill handle and the header grips take their own drags before the grid sees them.
  wrap.addEventListener("pointerdown", (e) => {
    const ev = e as PointerEvent;
    const grip = (ev.target as HTMLElement).closest?.<HTMLElement>(".grip");
    if (!grip || ev.button !== 0) return;
    ev.preventDefault();
    ev.stopPropagation();
    capture(wrap, ev.pointerId);
    view.hooks.edited();
    if (grip.dataset.col !== undefined) view.startSize("col", Number(grip.dataset.col), ev.clientX);
    else view.startSize("row", Number(grip.dataset.row), ev.clientY);
  });

  // Pointer capture sends moves to the wrap, so hit-test by position instead.
  wrap.addEventListener("pointermove", (e) => {
    const ev = e as PointerEvent;
    const cellUnder = (target: EventTarget | null): HTMLElement | undefined =>
      (target as HTMLElement | null)?.closest?.<HTMLElement>("td[data-col]") ??
      document.elementFromPoint(ev.clientX, ev.clientY)?.closest<HTMLElement>("td[data-col]") ??
      undefined;
    if (view.filling) {
      const el = cellUnder(ev.target);
      if (el) view.fillTo(Number(el.dataset.col), Number(el.dataset.row));
      return;
    }
    if (view.sizing) {
      view.sizeTo(view.sizing.axis === "col" ? ev.clientX : ev.clientY);
      return;
    }
    const overHead = (ev.target as HTMLElement)?.closest?.<HTMLElement>("th[data-col], th[data-row]");
    if (overHead) {
      const grip =
        overHead.dataset.col !== undefined
          ? overHead.clientWidth - ev.offsetX <= 6
          : overHead.clientHeight - ev.offsetY <= 5;
      wrap.style.cursor = grip ? (overHead.dataset.col !== undefined ? "col-resize" : "row-resize") : "";
    } else wrap.style.cursor = "";
    const el = cellUnder(ev.target);
    if (el) view.dragTo(Number(el.dataset.col), Number(el.dataset.row));
  });
  wrap.addEventListener("pointerup", () => {
    if (view.filling) return view.applyFill();
    if (view.sizing) return view.endSize();
    view.pointerUp();
  });
  wrap.addEventListener("pointercancel", () => {
    view.filling = null;
    view.sizing = null;
    view.pointerUp();
  });
  wrap.addEventListener("dblclick", (e) => {
    if (cellAt((e as MouseEvent).target)) {
      e.preventDefault();
      view.beginEdit();
    }
  });
  wrap.addEventListener("contextmenu", (e) => {
    const ev = e as MouseEvent;
    ev.preventDefault();
    const at = cellAt(ev.target);
    if (at && !inBox(view.sel, at.c, at.r)) view.select(rect(at.c, at.r, at.c, at.r));
    view.hooks.menu(view, ev.clientX, ev.clientY);
  });

  // Keys inside the editor belong to the editor; they must not bubble into grid navigation.
  wrap.addEventListener("keydown", (e) => {
    if ((e.target as HTMLElement | null)?.closest?.(".cell-editor")) return;
    view.onKeyDown(e as KeyboardEvent);
  });
  wrap.addEventListener("copy", (e) => {
    const ce = e as ClipboardEvent;
    ce.clipboardData?.setData("text/plain", view.copyText());
    ce.preventDefault();
  });
  wrap.addEventListener("cut", (e) => {
    const ce = e as ClipboardEvent;
    ce.clipboardData?.setData("text/plain", view.copyText());
    ce.preventDefault();
    view.cut();
  });
  wrap.addEventListener("paste", (e) => {
    const ce = e as ClipboardEvent;
    const text = ce.clipboardData?.getData("text/plain") ?? "";
    if (!text) return;
    ce.preventDefault();
    view.paste(text);
  });
}

const inBox = (box: Rect, c: number, r: number): boolean => c >= box.c1 && c <= box.c2 && r >= box.r1 && r <= box.r2;
