import { Spreadsheet } from "../core/xlsx.js";
import { DEFAULT_SECTION, LINE_ONE, parseSizePt, sizeNameOf, twipsToPx, WordDocument, type PageSetup, type Section } from "../core/docx.js";
import { Presentation } from "../core/pptx.js";
import { OpcPackage } from "../core/opc.js";
import { writeXlsx } from "../core/write.js";
import { writeDocx } from "../core/write-docx.js";
import { History } from "../core/history.js";
import { CASE_OPTIONS, type CaseMode } from "../core/doc-case.js";
import { isError } from "../core/formula.js";
import { rangeSpec, rect, refOf, type Rect } from "../core/sheet-ops.js";
import { SheetView } from "./sheet-view.js";
import { DocEditor, type DocPath, type TableCommandAnswer } from "./doc-edit.js";
import type { RevisionView } from "../core/render.js";
import { type BorderLeg, type CellLeg, cellLegPatch, legPatch, tableLegPatch, type TableLeg } from "./doc-para-border.js";
import { DocStatusBar } from "./doc-status-bar.js";
import { showTableSortDialog } from "./doc-table-sort.js";
import { CUSTOM, PAPER, paperEdges, paperOf } from "./doc-page-setup.js";
import { currentUnit, onUnitChange, twipText, twipsFromUnit } from "./doc-units.js";
import { DeckEditor, type PresetGeometry } from "./deck-edit.js";
import { DeckShow } from "./deck-show.js";
import { dataUrl, IMAGE_MIME_BY_EXT } from "../core/media.js";
import { writePptx } from "../core/write-pptx.js";

interface Bridge {
  pickFile(): Promise<{ name: string; path: string; data: Uint8Array } | null>;
  pickImage(): Promise<{ name: string; data: Uint8Array } | null>;
  saveFile(name: string, data: Uint8Array | string): Promise<string | null>;
  writeFile(path: string, data: Uint8Array): Promise<string | null>;
  confirmUnsaved(name: string): Promise<"save" | "discard" | "cancel">;
  closeReply(proceed: boolean): void;
  /** Only the shell can enter macOS fullscreen without a user-gesture requirement. */
  setFullscreen?(on: boolean): void;
  /** Renders the current view to a PDF file; resolves to the path, null on cancel, or an error string. */
  exportPdf?(name: string, landscape: boolean): Promise<string | null>;
}

declare global {
  interface Window {
    soloOffice?: Bridge;
  }
}

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
const view = $<HTMLDivElement>("view");
const sheetbar = $<HTMLDivElement>("sheetbar");
const bar = $<HTMLDivElement>("bar");
const ribbon = $<HTMLDivElement>("ribbon");
const addr = $<HTMLDivElement>("addr");
const input = $<HTMLInputElement>("input");
const title = $<HTMLDivElement>("title");
const saveBtn = $<HTMLButtonElement>("save-btn");

type Opened =
  | { kind: "xlsx"; name: string; path?: string; wb: Spreadsheet; sheet: number }
  | { kind: "docx"; name: string; path?: string; doc: WordDocument }
  | { kind: "pptx"; name: string; path?: string; pres: Presentation; slide: number };

let current: Opened | null = null;
let grid: SheetView | null = null;
let editor: DocEditor | null = null;
/**
 * The 文字页脚的状态栏. It reports the open document's counts, so it travels with the editor: every path that
 * drops `editor` takes this down too, and it never lands on a spreadsheet's or a deck's page.
 */
let docStatus: DocStatusBar | null = null;
let deck: DeckEditor | null = null;
/** Non-null while a slideshow covers the screen; global keys yield to it. */
let shower: DeckShow | null = null;
let history = new History();
let dirty = false;
let lastCopy = "";

// ------------------------------------------------------------------ opening

async function openBytes(name: string, data: Uint8Array, path?: string): Promise<void> {
  shower?.close();
  let pkg: OpcPackage;
  try {
    pkg = OpcPackage.from(data);
  } catch (e) {
    fail("无法打开", `${name} 不是可读取的 OOXML 包：${(e as Error).message}`);
    return;
  }
  try {
    if (pkg.kind === "xlsx") {
      const wb = Spreadsheet.open(pkg);
      current = { kind: "xlsx", name, path, wb, sheet: wb.activeSheet };
      history = new History();
      dirty = false;
    } else if (pkg.kind === "docx") {
      current = { kind: "docx", name, path, doc: WordDocument.openBytes(data) };
    } else if (pkg.kind === "pptx") {
      current = { kind: "pptx", name, path, pres: Presentation.openBytes(data), slide: 0 };
    } else {
      fail("暂不支持", `${name} 里既没有 word/document.xml、xl/workbook.xml 也没有 ppt/presentation.xml。`);
      return;
    }
  } catch (e) {
    fail("解析失败", (e as Error).message);
    return;
  }
  title.textContent = name;
  saveBtn.disabled = false;
  mount();
}

function fail(heading: string, detail: string): void {
  // The document never made it into the editor, so nothing is open any more —
  // a later save must not write the file that was open before.
  current = null;
  grid = null;
  editor = null;
  dropDocStatusBar();
  deck = null;
  saveBtn.disabled = true;
  sheetUI(false);
  title.textContent = "没有打开的文档";
  view.innerHTML = emptyHtml(escapeHtml(heading), escapeHtml(detail));
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

const sheetUI = (on: boolean, barOn = on): void => {
  bar.classList.toggle("on", barOn);
  ribbon.classList.toggle("on", on);
  sheetbar.style.display = on && barOn ? "flex" : "none";
};

/** The 状态栏 reads the open document, so wherever the editor pointer is dropped the bar goes with it. */
function dropDocStatusBar(): void {
  docStatus?.detach();
  docStatus = null;
}

function mount(): void {
  if (!current) return emptyState();
  editor = null;
  dropDocStatusBar();
  deck = null;
  if (current.kind === "xlsx") return mountSheet(current);
  grid = null;
  if (current.kind === "docx") return mountDoc(current);
  sheetUI(false);
  return mountDeck(current);
}

function emptyHtml(headingHtml: string, detailHtml: string): string {
  return `<div class="empty"><div class="drop" id="drop">
      <h1>${headingHtml}</h1>
      <p>${detailHtml}</p>
      <div class="fmts">
        <button class="fmt x" data-new="xlsx">新建表格</button>
        <button class="fmt w" data-new="docx">新建文字</button>
        <button class="fmt p" data-new="pptx">新建演示</button>
      </div>
    </div></div>`;
}

function emptyState(): void {
  grid = null;
  editor = null;
  dropDocStatusBar();
  deck = null;
  sheetUI(false);
  view.innerHTML = emptyHtml("打开或新建文档", `拖进来，或者按 <kbd>⌘O</kbd> 打开 XLSX / DOCX / PPTX`);
}

// ---------------------------------------------------------------- xlsx shell

function mountSheet(app: Extract<Opened, { kind: "xlsx" }>): void {
  // The grid owns #view, so the find panel has to be parked aside and put back.
  const find = document.getElementById("find");
  sheetUI(true);
  buildRibbon("xlsx");
  grid = new SheetView(view, app.wb, app.sheet, {
    edited: () => {
      record();
      refreshChrome();
    },
    selected: () => refreshChrome(),
    menu: (v, x, y) => openMenu(v, x, y),
  });
  grid.render();
  if (find) view.append(find);
  grid.select(rect(0, 0, 0, 0));
  mountSheetTabs(app);
  refreshChrome();
  grid.focusGrid();
}

/** Snapshot the state the user is about to leave behind, before the mutation runs. */
function record(): void {
  if (!current || current.kind !== "xlsx") return;
  history.record(current.wb);
  dirty = true;
  title.textContent = current.name + " ●";
}

const round = (n: number): string => String(Number(n.toFixed(6)));

const statusEl = (): HTMLElement | null => document.getElementById("status");

const flash = (message: string): void => {
  if (statusEl()) statusEl()!.textContent = message;
};

/**
 * 文档态没有 `#status`（那是表格壳的底栏），而结构命令的拒绝得说得出理由——「下面没有可让位的单元格」这类句子
 * 是 `DocEditor` 算出来的，不是按钮猜的。给它一枚浮在纸面右下角的小条，四秒后退场。WPS 拒绝时画的是什么
 * **未取证**，这一枚是本产品自己的呈报面，不按 WPS 的形状引用。
 */
let noticeEl: HTMLElement | undefined;
let noticeTimer = 0;
const notify = (message: string): void => {
  if (statusEl()) {
    flash(message);
    return;
  }
  if (!noticeEl) {
    noticeEl = document.createElement("div");
    noticeEl.className = "doc-notice";
    noticeEl.setAttribute("role", "status");
    noticeEl.style.cssText =
      "position:fixed;right:18px;bottom:52px;z-index:40;max-width:380px;padding:7px 11px;border-radius:6px;" +
      'background:#fff;color:#1f2328;border:1px solid #d5dae1;font:12px/1.5 -apple-system,"PingFang SC",sans-serif;box-shadow:0 4px 16px #0002;pointer-events:none';
    document.body.append(noticeEl);
  }
  noticeEl.textContent = message;
  noticeEl.hidden = false;
  window.clearTimeout(noticeTimer);
  noticeTimer = window.setTimeout(() => {
    if (noticeEl) noticeEl.hidden = true;
  }, 4000);
};

function refreshChrome(): void {
  if (!current || current.kind !== "xlsx" || !grid) return;
  const wb = current.wb;
  const name = grid.sheet.name;
  const at = refOf(grid.active.c, grid.active.r);
  addr.textContent = rangeSpec(grid.sel);
  // Never clobber what the user is typing in the formula bar.
  if (document.activeElement !== input) input.value = wb.editOf(name, at);
  syncRibbon(grid.activeStyle());
  refreshHistoryButtons();

  const stats = wb.stats(name, grid.sel);
  const multi = grid.sel.c1 !== grid.sel.c2 || grid.sel.r1 !== grid.sel.r2;
  const active = wb.value(name, at);
  const bits: string[] = [`${name} · ${grid.sheet.cells.size} 个单元格`];
  if (multi && stats.numeric) {
    bits.push(`计数 ${stats.count}`, `求和 ${round(stats.sum)}`, `平均 ${round(stats.avg)}`, `最小 ${round(stats.min)}`, `最大 ${round(stats.max)}`);
  } else if (typeof active === "number") bits.push(`值 ${round(active)}`);
  else if (isError(active)) bits.push(String(active));
  flash(bits.join(" · "));
}

function refreshHistoryButtons(): void {
  if (!current || current.kind !== "xlsx") return;
  const toggle = (key: string, enabled: boolean): void => {
    const b = ctl[key] as HTMLButtonElement | undefined;
    if (b) b.disabled = !enabled;
  };
  toggle("undo", history.canUndo(current.wb));
  toggle("redo", history.canRedo(current.wb));
}

// ------------------------------------------------------------------- ribbon

const NUMFMTS: [string, string][] = [
  ["General", "常规"],
  ["0", "整数"],
  ["#,##0", "千位分隔"],
  ["#,##0.00", "千位两位"],
  ["0%", "百分比"],
  ["0.00%", "百分比两位"],
  ["¥#,##0.00", "货币"],
  ["0.00E+00", "科学计数"],
  ["yyyy-mm-dd", "日期"],
  ["hh:mm", "时间"],
  ["@", "文本"],
];

const SIZES = [9, 10, 11, 12, 14, 18, 24, 36];
/**
 * 字号下拉那 21 档：`CT_TableTools.kuip` 的 `KxWpsFontSizeComboBoxCommand` 逐字写着
 * `fontSizeList="@5;5.5;6.5;7.5;8;9;10;10.5;11;12;14;16;18;20;22;26;28;36;48;56;72"`，
 * 开始页（`CT_Home.kuip` 的 `<command idQ="FontSize"/>`）、段落（`CT_Paragraph.kuip`）、正文（`CT_TextTool.kuip`）、
 * 绘图工具与表格工具页的每一处声明都是同一串。旧清单里的 24 与 42 两档不在这一串里，已按 WPS 切掉。
 * 这一串管的是**下拉**；框本身是 ComboBox（类名如此），能打别的数，见 `parseSizePt` 与 `sizeBox`。
 */
const DOC_SIZES = [5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 26, 28, 36, 48, 56, 72];
const DECK_SIZES = [12, 14, 16, 18, 20, 24, 28, 32, 40, 44];

/**
 * 行距下拉那五档倍数。档位与次序按 Writer 自己声明的那一族命令：`wpsentries.kuip` 的
 * `CT_Home.RB_Paragraph.*.m_LineSpacing` 挂 `LineSpacing1_0 / 1_5 / 2_0 / 2_5 / 3_0 / LineSpacingMore` 六条，
 * `wpscommon.kuip` 把前三条的字面写成 `1.0`/`1.5`/`2.0`；后两条的中文在 `.qm` 那堵墙后面读不出，所以「2.5」
 * 「3.0」与「行距…」是照前三条的写法补的**推断字面**，不是取到的词。落下去的数是 `w:line = 倍数 × 240`
 * 配 `w:lineRule="auto"`（§17.3.1.27）。
 *
 * 别拿 `wppentries.kuip`（演示）里那条 `LineSpacingInHomeTab` 当依据：那一份还多摆 1.25 与 1.75，是 PPT 的档位，
 * 不是 Writer 的。
 */
const DOC_LINES = [1, 1.5, 2, 2.5, 3];

/**
 * 突出显示's palette. The names are what `w:highlight w:val=` can hold (§17.17.64 has 17 members); these eight are
 * what this ribbon can paint. 授权 46 份里 `w:highlight` **0 枚**（对照：同一趟在那批件上数到 1,615 枚 `w:u`），所以
 * 调色板缺不缺口今天没有语料能判——旧读数「86 份里只有 yellow 与 none 落到带文字的 run 上」属作废语料，待重证。
 * 这八条留着是因为写端要把用户选的色拼成 §17.17.64 的成员名，不是因为读到过它们。
 */
const DOC_HIGHLIGHTS: [string, string][] = [
  ["none", "高亮：无"],
  ["yellow", "黄"],
  ["green", "绿"],
  ["cyan", "青"],
  ["magenta", "品红"],
  ["red", "红"],
  ["darkYellow", "深黄"],
  ["lightGray", "浅灰"],
];

const ctl: Record<string, HTMLElement> = {};

/**
 * 「边框」那一格按选区派发的三张下拉——WPS 那边本来就**是一枚控件**：`RbTableOutsideBorder` 的十条腿
 * （`res/commands/wps/ongmani/CT_Home.kuip:272-287`，`:311`/`:362` 把它挂在**段落组**与表格工具两处）
 * 每一条都绑同一条 `ksoCmd="TableOutsideBorderPalette"`，落在哪一层由选区决定。
 * 段落那张摆九档：四条边 + 无框线 + 所有框线 + 外侧框线 + 内部横框线（= `w:between`）+ 那枚对话框。
 * 少的那两档是 `CrossBorder`（横竖框线）与 `VMiddleBorder`（内部竖框线）——`CT_PBdr` 只有横着的那一枚
 * `w:between`，竖着的「内部」在段落这一层没有对应的一句话（`w:bar` 说的是贯穿左边距的那一条，不是一层内部）。
 * 表那张是十条腿的十档（去掉说不出来的 `between/start/end` 与对角线）。
 * 格那张摆四边 + 无框线 + 所有框线 + 内部横竖（那两档照 WPS 派发去向表）：`CT_TcBorders` 没有 inside 槽；
 * 「外侧框线」说的是**选区的周边**（选 2×2 时中间那十字不该画），这一层还没实现，先不摆——摆出来就是给用户
 * 一条会落错的腿。
 *
 * 各档中文措辞在 `.qm`（UTF-16）里读不出，属**推断措辞**，不是取到的字面量。
 */
const BORDER_LEGS: Record<"para" | "table" | "cells", [string, string][]> = {
  para: [
    ["", "边框"],
    ["bottom", "下框线"],
    ["top", "上框线"],
    ["left", "左框线"],
    ["right", "右框线"],
    ["none", "无框线"],
    ["all", "所有框线"],
    ["outside", "外侧框线"],
    ["insideH", "内部横框线"],
    ["more", "边框与底纹…"],
  ],
  table: [
    ["", "边框"],
    ["bottom", "下框线"],
    ["top", "上框线"],
    ["left", "左框线"],
    ["right", "右框线"],
    ["none", "无框线"],
    ["all", "所有框线"],
    ["outside", "外侧框线"],
    ["cross", "横竖框线"],
    ["insideH", "内部横框线"],
    ["insideV", "内部竖框线"],
    // 最后这一档与段落那一层是**同一条命令**：`CT_Home.kuip:286` 把它挂在同一个下拉的最后一段，
    // 落点由选区决定（见 `openBorderDialog`），不是段落专属。
    ["more", "边框与底纹…"],
  ],
  cells: [
    ["", "边框"],
    ["bottom", "下框线"],
    ["top", "上框线"],
    ["left", "左框线"],
    ["right", "右框线"],
    ["none", "无框线"],
    ["all", "所有框线"],
    // #151：格选区与 WPS 那一枚控件说同一组话。外侧框线落在**选区的外圈**（`setCellOutside` 逐格算朝向块外的边），
    // 横竖框线仍由表来说（`CT_TcBorders` 没有「内部」那一枚孩子）。
    ["outside", "外侧框线"],
    ["cross", "横竖框线"],
    ["insideH", "内部横框线"],
    ["insideV", "内部竖框线"],
    ["more", "边框与底纹…"],
  ],
};

/**
 * 命令钮上的图标：24×24 的细线 SVG，全是自家画的（WPS 那套是装机件里的彩色位图，形状不同是已知差别，
 * 一份图形资源也不搬）。键是自家命令名，不是 WPS 的 `icon="..."` 名。
 * 图标里不放 `<text>`：按钮的 `textContent` 必须仍等于那行字——DOM 探针与烟测都按文字找按钮。
 */
const RIBBON_ICONS: Record<string, string> = {
  undo: "<path d='M9 14 4 9l5-5'/><path d='M4 9h10.5a5.5 5.5 0 0 1 0 11H9'/>",
  redo: "<path d='M15 14l5-5-5-5'/><path d='M20 9H9.5a5.5 5.5 0 0 0 0 11H15'/>",
  rowAbove: "<path d='M4 4h16'/><path d='M12 8v9'/><path d='M8.5 11.5 12 8l3.5 3.5'/>",
  rowBelow: "<path d='M4 20h16'/><path d='M12 16V7'/><path d='M8.5 12.5 12 16l3.5-3.5'/>",
  colLeft: "<path d='M4 4v16'/><path d='M8 12h9'/><path d='M13.5 8.5 17 12l-3.5 3.5'/>",
  colRight: "<path d='M20 4v16'/><path d='M16 12H7'/><path d='M10.5 8.5 7 12l3.5 3.5'/>",
  wrap: "<path d='M4 6h16M4 12h11a3 3 0 0 1 0 6h-2'/><path d='M15 15l-2 3 2 3'/>",
  bold: "<path d='M8 5h5a3 3 0 0 1 0 6H8zM8 11h6a3 3 0 0 1 0 6H8z'/>",
  italic: "<path d='M10 5h7M14 19H7M14 5l-4 14'/>",
  underline: "<path d='M7 5v6a5 5 0 0 0 10 0V5M6 20h12'/>",
  strike: "<path d='M5 12h14M8 8a4 4 0 0 1 8 0M16 16a4 4 0 0 1-8 0'/>",
  superscript: "<path d='M4 17 9 6l5 11M6 13h6'/><path d='M16 9h4v4h-4z'/>",
  subscript: "<path d='M4 15 9 4l5 11M6 11h6'/><path d='M16 16h4v4h-4z'/>",
  alignLeft: "<path d='M4 6h16M4 10h10M4 14h16M4 18h7'/>",
  alignCenter: "<path d='M4 6h16M7 10h10M4 14h16M8 18h8'/>",
  alignRight: "<path d='M4 6h16M10 10h10M4 14h16M13 18h7'/>",
  alignJustify: "<path d='M4 6h16M4 10h16M4 14h16M4 18h16'/>",
  valignTop: "<path d='M4 5h16'/><path d='M12 10v9'/><path d='M8.5 13.5 12 10l3.5 3.5'/>",
  valignCenter: "<path d='M4 12h4M16 12h4'/><path d='M12 4v16'/><path d='M9 7l3-3 3 3M9 17l3 3 3-3'/>",
  valignBottom: "<path d='M4 19h16'/><path d='M12 5v9'/><path d='M8.5 10.5 12 14l3.5-3.5'/>",
  borderNone: "<path d='M4 4h16v16H4z' stroke-dasharray='2 2'/><path d='M5 19 19 5'/>",
  borderBox: "<path d='M4 4h16v16H4z'/>",
  borderAll: "<path d='M4 4h16v16H4z'/><path d='M12 4v16M4 12h16'/>",
  merge: "<path d='M3 5h6v14H3zM15 5h6v14h-6z'/><path d='M9 12h6'/>",
  split: "<path d='M3 5h18v14H3z'/><path d='M12 5v14'/>",
  clearFormat: "<path d='M5 16l9-9a2 2 0 0 1 3 3l-9 9H5v-3z'/><path d='M4 21h16'/>",
  colorScale: "<path d='M4 18h4v-4H4zM10 18h4v-8h-4zM16 18h4V6h-4z'/>",
  dataBar: "<path d='M4 7h6v3H4zM4 12h10v3H4zM4 17h14v3H4z'/>",
  highlight: "<path d='M5 15 14 6l4 4-9 9H5v-4z'/><path d='M13 7l4 4M4 21h16'/>",
  find: "<path d='M10.5 4a6.5 6.5 0 1 1 0 13 6.5 6.5 0 0 1 0-13z'/><path d='M15.5 15.5 20 20'/>",
  filter: "<path d='M4 5h16l-6 7v6l-4 2v-8z'/>",
  sortAsc: "<path d='M4 7h10M4 12h7M4 17h4'/><path d='M18 18V8'/><path d='M15 11l3-3 3 3'/>",
  sortDesc: "<path d='M4 7h4M4 12h7M4 17h10'/><path d='M18 8v10'/><path d='M15 15l3 3 3-3'/>",
  freeze: "<path d='M4 4h16v16H4z'/><path d='M4 9h16M9 4v16'/>",
  unfreeze: "<path d='M4 4h16v16H4z'/><path d='M9 4v16M4 9h7M16 9h4M16 15h4'/>",
  chartBar: "<path d='M5 19V9M11 19V5M17 19v-6'/><path d='M3 21h18'/>",
  chartLine: "<path d='M4 16l5-5 4 3 6-7'/><path d='M3 21h18'/>",
  chartPie: "<path d='M12 3a9 9 0 1 0 9 9h-9z'/><path d='M12 3v9h9'/>",
  picture: "<path d='M3 5h18v14H3z'/><path d='M3 16l5-5 4 4 3-3 6 6'/><path d='M8.5 8.5h.01'/>",
  formatPainter: "<path d='M14 3l7 7-3 3-7-7z'/><path d='M11 10l-6 6v4h4l6-6'/>",
  selectAll: "<path d='M4 4h5M15 4h5M4 20h5M15 20h5M4 9v4M20 9v4'/><path d='M9 12h6'/>",
  indentInc: "<path d='M4 6h16M4 12h16M4 18h16'/><path d='M8 9l3 3-3 3'/>",
  indentDec: "<path d='M4 6h16M4 12h16M4 18h16'/><path d='M12 9l-3 3 3 3'/>",
  ruler: "<path d='M3 8h18v8H3z'/><path d='M7 8v4M11 8v4M15 8v4M19 8v4'/>",
  navPane: "<path d='M4 5h16v14H4z'/><path d='M9.5 5v14'/><path d='M12.5 9h4.5M12.5 12h4.5M12.5 15h3'/>",
  bullet: "<path d='M9 7h11M9 12h11M9 17h11'/><path d='M4.5 7h.01M4.5 12h.01M4.5 17h.01'/>",
  numberList: "<path d='M9 7h11M9 12h11M9 17h11'/><path d='M4 5h1v4M4 13h2l-2 3h2'/>",
  locate: "<path d='M12 3v4M12 17v4M3 12h4M17 12h4'/><path d='M12 8a4 4 0 1 1 0 8 4 4 0 0 1 0-8z'/>",
  pageSetup: "<path d='M5 3h14v18H5z'/><path d='M8 7h8M8 11h8M8 15h5'/>",
  header: "<path d='M5 3h14v18H5z'/><path d='M5 8h14'/>",
  footer: "<path d='M5 3h14v18H5z'/><path d='M5 16h14'/>",
  table: "<path d='M3 5h18v14H3z'/><path d='M3 10h18M3 15h18M9 5v14M15 5v14'/>",
  comment: "<path d='M4 5h16v11H9l-4 4v-4H4z'/><path d='M12 8v5M9.5 10.5h5'/>",
  trackChanges: "<path d='M4 9a8 8 0 0 1 14-3'/><path d='M18 4v5h-5'/><path d='M20 15a8 8 0 0 1-14 3'/><path d='M6 20v-5h5'/>",
  wordCount: "<path d='M4 6h10M4 12h8M4 18h6'/><path d='M15 17l2 2 3-4'/>",
  slideAdd: "<path d='M3 5h18v11H3z'/><path d='M12 16v4M9 20h6'/><path d='M12 7v6M9 10h6'/>",
  slideDelete: "<path d='M3 5h18v11H3z'/><path d='M12 16v4M9 20h6'/><path d='M9 8l6 6M15 8l-6 6'/>",
  slideUp: "<path d='M7 3h10v5H7z'/><path d='M12 21v-7'/><path d='M8.5 17.5 12 14l3.5 3.5'/>",
  slideDown: "<path d='M7 16h10v5H7z'/><path d='M12 3v7'/><path d='M8.5 6.5 12 10l3.5-3.5'/>",
  textBox: "<path d='M3 5h18v14H3z'/><path d='M8 9h8M12 9v6'/>",
  deleteShape: "<path d='M4 4h9v9H4z' stroke-dasharray='2 2'/><path d='M13 13l7 7M20 13l-7 7'/>",
  play: "<path d='M7 4l12 8-12 8z'/>",
  playHere: "<path d='M5 4l11 8-11 8z'/><path d='M19 5v14'/>",
  shapeRect: "<path d='M4 6h16v12H4z'/>",
  shapeRoundRect: "<path d='M8 6h8a4 4 0 0 1 4 4v4a4 4 0 0 1-4 4H8a4 4 0 0 1-4-4v-4a4 4 0 0 1 4-4z'/>",
  shapeEllipse: "<path d='M12 5c4.4 0 8 3.1 8 7s-3.6 7-8 7-8-3.1-8-7 3.6-7 8-7z'/>",
  shapeTriangle: "<path d='M12 4l9 16H3z'/>",
  shapeLine: "<path d='M4 20 20 4'/><path d='M4 20h.01M20 4h.01'/>",
  shapeDiamond: "<path d='M12 3l9 9-9 9-9-9z'/>",
  shapeParallelogram: "<path d='M7 6h14l-4 12H3z'/>",
  shapeRightArrow: "<path d='M3 9h11V5l7 7-7 7v-4H3z'/>",
  shapeStar: "<path d='M12 3.5l2.6 5.6 6.1.7-4.6 4.2 1.3 6-5.4-3.1-5.4 3.1 1.3-6L3.3 9.8l6.1-.7z'/>",
  shapeChevron: "<path d='M4 6h13l4 6-4 6H4l4-6z'/>",
};

/** 哪条命令挂哪枚图标。认不出的命令不挂——纯文字的小钮（小数位、上下对齐箭头那些）保持原样。 */
const RIBBON_ICON_IDS: Record<string, string> = {
  撤销: "undo",
  重做: "redo",
  "＋行": "rowAbove",
  "－行": "rowBelow",
  "＋列": "colLeft",
  "－列": "colRight",
  换行: "wrap",
  粗: "bold",
  斜: "italic",
  线: "underline",
  删: "strike",
  上标: "superscript",
  下标: "subscript",
  左: "alignLeft",
  中: "alignCenter",
  右: "alignRight",
  两: "alignJustify",
  上: "valignTop",
  下: "valignBottom",
  "↑": "valignTop",
  "↕": "valignCenter",
  "↓": "valignBottom",
  无框: "borderNone",
  外框: "borderBox",
  所有框: "borderAll",
  合并: "merge",
  拆分: "split",
  清格式: "clearFormat",
  清除格式: "clearFormat",
  色阶: "colorScale",
  数据条: "dataBar",
  突出显示: "highlight",
  查找: "find",
  查找替换: "find",
  筛选: "filter",
  升序: "sortAsc",
  降序: "sortDesc",
  冻结: "freeze",
  解冻: "unfreeze",
  柱状图: "chartBar",
  折线图: "chartLine",
  饼图: "chartPie",
  图片: "picture",
  格式刷: "formatPainter",
  全选: "selectAll",
  "缩进+": "indentInc",
  "缩进−": "indentDec",
  "• 符号": "bullet",
  "1. 编号": "numberList",
  制表位: "ruler",
  标尺: "ruler",
  导航窗格: "navPane",
  定位: "locate",
  页面设置: "pageSetup",
  页眉: "header",
  页脚: "footer",
  表格: "table",
  新建批注: "comment",
  修订: "trackChanges",
  字数统计: "wordCount",
  "＋新页": "slideAdd",
  删除页: "slideDelete",
  上移: "slideUp",
  下移: "slideDown",
  文本框: "textBox",
  删除形状: "deleteShape",
  放映: "play",
  本页放映: "playHere",
  矩形: "shapeRect",
  圆角: "shapeRoundRect",
  椭圆: "shapeEllipse",
  三角: "shapeTriangle",
  直线: "shapeLine",
  菱形: "shapeDiamond",
  平行四边形: "shapeParallelogram",
  右箭头: "shapeRightArrow",
  五角星: "shapeStar",
  燕尾形: "shapeChevron",
};

/**
 * 把一枚命令钮改成「图标在上、字在下」。`textContent` 仍是那行字：图标 span 里没有文字节点，
 * 所以按文字找按钮的探针、烟测与回显都照旧。键表里没有的命令原样返回。
 */
function decorateRibbonButton(b: HTMLButtonElement, label: string): void {
  const id = RIBBON_ICON_IDS[label];
  if (!id) return;
  const art = document.createElement("span");
  art.className = "ic";
  art.setAttribute("aria-hidden", "true");
  art.innerHTML = `<svg viewBox="0 0 24 24">${RIBBON_ICONS[id]}</svg>`;
  const text = document.createElement("span");
  text.className = "lb";
  text.textContent = label;
  b.replaceChildren(art, text);
  b.classList.add("ico");
  b.dataset.ic = id;
}

function buildRibbon(kind: "xlsx" | "docx" | "pptx"): void {
  if (ribbon.dataset.built === kind) return;
  ribbon.dataset.built = kind;
  ribbon.innerHTML = "";
  // 三页都按 WPS 自己的页签条分页（`wpsongmani.kui` / `etongmani.kui` / `wppongmani.kui` 的
  // `KRbTabWidgetCommand` 各声明一套）。
  ribbon.classList.add("tabbed");
  for (const key of Object.keys(ctl)) delete ctl[key];
  if (kind === "xlsx") buildSheetRibbon();
  else if (kind === "docx") buildDocRibbon();
  else buildDeckRibbon();
}

function buildSheetRibbon(): void {
  const shell = ribbonTabs(SHEET_RIBBON_TABS);
  // 每组底下一行组名，次序与下面建组的次序一一对应（WPS 的组名在 `.qm` 里读不出来，这一行是我们自己的措辞）。
  // 用队列而不是给每个调用点加参数：组一多，漏一个就错位一整页，队列少了会立刻少一条组名。
  const caps: Record<string, string[]> = {
    home: ["行列", "数字", "字体", "对齐", "单元格", "条件格式", "编辑"],
    insert: ["插图", "图表"],
    data: ["排序和筛选"],
    view: ["窗口"],
  };
  const group = (tab: string): HTMLElement => {
    const g = shell.group(tab);
    const cap = caps[tab]?.shift();
    if (cap) g.dataset.cap = cap;
    return g;
  };
  const button = (g: HTMLElement, key: string, label: string, tip: string, run: () => void): void => {
    const b = document.createElement("button");
    b.className = "rb";
    b.textContent = label;
    b.title = tip;
    decorateRibbonButton(b, label);
    b.onclick = () => {
      grid?.commitEdit();
      run();
    };
    if (key) ctl[key] = b;
    g.append(b);
  };
  const picker = (g: HTMLElement, key: string, tip: string, run: (value: string) => void): void => {
    const i = document.createElement("input");
    i.type = "color";
    i.className = "swatch";
    i.title = tip;
    i.value = "#000000";
    i.oninput = () => run(i.value);
    ctl[key] = i;
    g.append(i);
  };
  const chooser = (g: HTMLElement, key: string, tip: string, options: readonly [string, string][], run: (value: string) => void): void => {
    const s = document.createElement("select");
    s.className = "pick";
    s.title = tip;
    for (const [value, label] of options) {
      const o = document.createElement("option");
      o.value = value;
      o.textContent = label;
      s.append(o);
    }
    s.onchange = () => run(s.value);
    ctl[key] = s;
    g.append(s);
  };
  const use = (run: (v: SheetView) => void): void => {
    if (grid) run(grid);
  };

  const nav = group("home");
  button(nav, "undo", "撤销", "撤销 ⌘Z", undo);
  button(nav, "redo", "重做", "重做 ⌘⇧Z", redo);
  button(nav, "", "＋行", "在选区上方插入行", () => use((v) => v.line("row", true)));
  button(nav, "", "－行", "删除选中的行", () => use((v) => v.line("row", false)));
  button(nav, "", "＋列", "在选区左侧插入列", () => use((v) => v.line("col", true)));
  button(nav, "", "－列", "删除选中的列", () => use((v) => v.line("col", false)));

  const numbers = group("home");
  chooser(numbers, "numfmt", "数字格式", NUMFMTS, (v) => use((g) => g.setNumFmt(v)));
  button(numbers, "", "0.", "减少小数位", () => use((v) => v.stepDecimals(-1)));
  button(numbers, "", "0.0+", "增加小数位", () => use((v) => v.stepDecimals(1)));
  button(numbers, "wrap", "换行", "自动换行", () => use((v) => v.toggle("wrap")));

  const font = group("home");
  chooser(
    font,
    "size",
    "字号",
    SIZES.map((s): [string, string] => [String(s), String(s)]),
    (v) => use((g) => g.setFontSize(Number(v))),
  );
  button(font, "bold", "粗", "加粗 ⌘B", () => use((v) => v.toggle("bold")));
  button(font, "italic", "斜", "斜体 ⌘I", () => use((v) => v.toggle("italic")));
  button(font, "under", "线", "下划线 ⌘U", () => use((v) => v.toggle("underline")));
  picker(font, "color", "字体颜色", (v) => use((g) => g.setFontColor(v)));
  picker(font, "fill", "填充颜色", (v) => use((g) => g.setFill(v)));

  const align = group("home");
  button(align, "left", "左", "左对齐", () => use((v) => v.setAlign("left")));
  button(align, "center", "中", "水平居中", () => use((v) => v.setAlign("center")));
  button(align, "right", "右", "右对齐", () => use((v) => v.setAlign("right")));
  button(align, "vtop", "上", "顶端对齐", () => use((v) => v.setValign("top")));
  button(align, "vmid", "↕", "垂直居中", () => use((v) => v.setValign("center")));
  button(align, "vbot", "下", "底端对齐", () => use((v) => v.setValign("bottom")));

  const cells = group("home");
  button(cells, "", "无框", "清除边框", () => use((v) => v.setBorder("none")));
  button(cells, "", "外框", "绘制外边框", () => use((v) => v.setBorder("box")));
  button(cells, "", "所有框", "绘制全部边框", () => use((v) => v.setBorder("all")));
  button(cells, "merge", "合并", "合并单元格", () => use((v) => v.mergeSelection()));
  button(cells, "", "拆分", "取消合并", () => use((v) => v.unmergeSelection()));
  button(cells, "", "清格式", "只保留内容", () => use((v) => v.clear("format")));

  const pic = group("insert");
  button(pic, "", "图片", "在当前单元格插入本地图片", () => void pickPicture((url, name, w, h) => grid?.insertPicture(url, w, h, name)));

  const cf = group("home");
  button(cf, "", "色阶", "对选区应用红黄绿三色阶", () => use((v) => v.setColorScale()));
  button(cf, "", "数据条", "对选区应用蓝色数据条", () => use((v) => v.setDataBar()));
  const cfBtn = document.createElement("button");
  cfBtn.className = "rb";
  cfBtn.textContent = "突出显示";
  cfBtn.title = "对选区应用大于/小于/等于规则";
  cfBtn.onclick = () => openCondPanel();
  ctl.cfBtn = cfBtn;
  cf.append(cfBtn);
  button(cf, "", "清条件", "清除选区的条件格式", () => use((v) => v.clearCond()));

  const charts = group("insert");
  button(charts, "", "柱状图", "用选区数据插入柱状图", () => use((v) => v.insertChart("bar")));
  button(charts, "", "折线图", "用选区数据插入折线图", () => use((v) => v.insertChart("line")));
  button(charts, "", "饼图", "用选区数据插入饼图", () => use((v) => v.insertChart("pie")));

  button(charts, "", "编辑图表", "选择图表后编辑类型、配色、标题和数据区域；也可双击图表", () => use((v) => v.editSelectedChart()));

  // 筛选与排序归数据页（`CT_Data.kuip` 的 `RB_Sort_Filter`）；WPS 在开始页另有一份快捷入口，
  // 家里只留数据页这一处，不让同一条命令在两个页签上各摆一枚。
  const data = group("data");
  button(data, "filter", "筛选", "为选区或当前数据区域启用/关闭筛选", () => use((v) => v.toggleFilter()));
  button(data, "", "清除筛选", "显示全部记录并保留筛选按钮", () => use((v) => v.clearFilters()));
  button(data, "", "升序", "按当前列升序排列选区", () => use((v) => v.sortSelection(false)));
  button(data, "", "降序", "按当前列降序排列选区", () => use((v) => v.sortSelection(true)));

  // 冻结窗格与缩放归视图页：`CT_View.kuip` 的 `RB_Workbook_View`(FreezePanes…) 与 `RB_Zoom`(Zoom)。
  const view = group("view");
  button(view, "", "冻结", "在当前位置冻结窗格", () => use((v) => v.freezeHere()));
  button(view, "", "解冻", "解除冻结窗格", () => use((v) => v.unfreeze()));
  shell.bars["view"]!.append(zoomGroup());

  // 查找属开始页的 `RB_Editing`。
  const editing = group("home");
  button(editing, "", "查找", "查找与替换 ⌘F", openFind);

  shell.sync();
}

/** 字体框的下拉数据源。`<datalist>` 挂在 body 上、不占版面；聚焦时才按当前文档刷一遍，因为「这份文件用过哪些脸」
 *  要把正文几千个 run 数一遍，每敲一键都数不值。 */
const FACES_ID = "doc-faces";
function fillFaceList(): void {
  let list = document.getElementById(FACES_ID);
  if (!list) {
    list = document.createElement("datalist");
    list.id = FACES_ID;
    document.body.append(list);
  }
  list.replaceChildren(
    ...(editor?.faces() ?? []).map((name) => {
      const o = document.createElement("option");
      o.value = name;
      return o;
    }),
  );
}

/** 字号框的下拉数据源：WPS 那份 `fontSizeList` 的 21 档。静态清单，不在聚焦时重算——它不随文档变。 */
const SIZES_ID = "doc-sizes";
function sizeList(): void {
  if (document.getElementById(SIZES_ID)) return;
  const list = document.createElement("datalist");
  list.id = SIZES_ID;
  list.replaceChildren(
    ...DOC_SIZES.map((n) => {
      const o = document.createElement("option");
      o.value = String(n);
      return o;
    }),
  );
  document.body.append(list);
}

/** WPS 自己声明的那条 tab 条（`res/wpsongmani.kui:113-162`，`<KRbTabWidgetCommand>` 的孩子按序排）。 */
const DOC_RIBBON_TABS: [string, string][] = [
  ["home", "开始"],
  ["insert", "插入"],
  ["layout", "页面布局"],
  ["review", "审阅"],
  ["view", "视图"],
  // 上下文页：WPS 的 `CT_TableTools` 只在光标进表时出现（`res/commands/wps/ongmani/CT_TableTools.kuip`）。
  ["table", "表格工具"],
];

/**
 * 表格页的页签次序，照 `res/etongmani.kui` 的 `<KRbTabWidgetCommand id="tabwidget">`：
 * `CT_Home → CT_Insert → CT_Layout → CT_Formulas → CT_Data → CT_Review → CT_View → CT_Tools`（其后是加载项/安全/工作区）。
 * 组归哪页按 `res/commands/et/ongmani/CT_*.kuip` 的组 id：`RB_Font`/`RB_Number`/`RB_Alignment`/`RB_Border`/
 * `RB_SheetTool`/`RB_Editing`/`RB_ConditionalFormatting` 在开始页；`RB_Illustrations`（图片、图表）在插入页；
 * `RB_Sort_Filter` 在数据页；`RB_Workbook_View`（冻结窗格）与 `RB_Zoom` 在视图页。
 * 只摆有命令的四档：公式/审阅/工具页今天没有命令，摆一张空页只是给人一条按不动的路。
 * 中文页名与文字页一样是自家措辞——`.kuip` 只给英文 source（`text="@Home"`），界面字面在读不出来的 `.qm` 里。
 */
const SHEET_RIBBON_TABS: [string, string][] = [
  ["home", "开始"],
  ["insert", "插入"],
  ["data", "数据"],
  ["view", "视图"],
];

/**
 * 演示页的页签，照 `res/wppongmani.kui:86` 的 `<KRbTabWidgetCommand id="tabwidget">`：
 * `CT_Home → CT_Insert → CT_Design → CT_SlideTrans → CT_Animation → CT_SlideShow → CT_Review → CT_View → …`。
 * 只摆有命令的四档（设计/切换/动画/审阅今天没有命令，空页只是给人一条按不动的路）。
 * 组归哪页按 `res/commands/wpp/ongmani/CT_*.kuip` 的组 id：`RB_Slide`/`RB_Font`/`RB_Paragraph`/`RB_Drawing`
 * 在开始页；`RB_Text`/`RB_Shape`/`RB_Media` 在插入页；`RB_StartSlideShow` 在放映页；`RB_Zoom` 在视图页。
 * 中文页名与组名一样是自家措辞——`.kuip` 只给英文 source（`text="@Home"`），界面字面在读不出来的 `.qm` 里。
 */
const DECK_RIBBON_TABS: [string, string][] = [
  ["home", "开始"],
  ["insert", "插入"],
  ["show", "放映"],
  ["view", "视图"],
];

/** 表格工具页上的命令控件：光标出表时页签退场，这几枚也一律灰着（键盘走得到，不能留一条能按的命令）。 */
const TABLE_TOOL_KEYS = [
  "rowAbove", "rowBelow", "colLeft", "colRight", "insertCellKind", "delRow", "delCol", "delTable", "deleteCellKind",
  "mergeCells", "splitCell", "splitTable",
  "cellHeight", "cellWidth", "distributeRows", "distributeCols",
  "tblFace", "tblSize", "tblBold", "tblItalic", "tblUnder", "tblColor", "tblHl",
  "tblTop", "tblMid", "tblBot", "tblLeft", "tblCenter", "tblRight",
  "repeatHeader", "convertTable", "sortRows",
  "selectCell", "selectCol", "selectRow", "selectTable",
] as const;

/**
 * `CT_PageLayout.kuip:153-166` 那一组（`RB_Page_Setup`，`needsinglerowmorebutton="true"`）里 `m_Margins` 之后
 * 连排的四枚 `KxWpsPageMarginSingleSpinCommand`：`component` 依次 2 / 0 / 3 / 1，也就是显示序 上 → 左 → 下 → 右，
 * 每枚 `fixedSpinBoxWidth="84"`、`singlerowfold="1"`，字面 `text="@Top"/"@Left"/"@Bottom"/"@Right"`。
 * 所以页边距在 ribbon 上是**四格填数**，不是预设菜单。第三列是节里读的那条边，第四列是写回时改的那一项。
 */
const PAGE_MARGIN_SPINS: [string, string, "top" | "left" | "bottom" | "right", "topTwips" | "leftTwips" | "bottomTwips" | "rightTwips"][] = [
  ["marginTop", "上", "top", "topTwips"],
  ["marginLeft", "左", "left", "leftTwips"],
  ["marginBottom", "下", "bottom", "bottomTwips"],
  ["marginRight", "右", "right", "rightTwips"],
];

/**
 * `RB_Adjust` 那两枚自旋格（`CT_TableTools.kuip` 的组序里排在 `RB_Merge` 之后：`RB_Table` → `RB_RowsAndColumns` →
 * `RB_Merge` → `RB_Adjust` → …）。OOXML 说不出「一格」的尺寸：高度归**行**（`w:trPr/w:trHeight`，§17.3.2.77），
 * 宽度归**列**（`w:gridCol/w:w` 加上那一列每一格那句 `w:tcPr/w:tcW`）。所以这一族只有两格是真的，第三列说的是
 * 「回显问哪一层」；往哪几行/哪几列落笔、说不出单一数时怎么回绝，都定在 `doc-edit.ts` 那两条命令里。
 */
const CELL_SIZE_SPINS: [string, string, "heightTwips" | "widthTwips"][] = [
  ["cellHeight", "高度", "heightTwips"],
  ["cellWidth", "宽度", "widthTwips"],
];

/**
 * `CT_TableTools.kuip:156-170` 的 `RB_Alignment`：两行 `KRbLayoutListCommand layoutOnly="1"`，先竖向三枚
 * （`TableToolsAlignTop/VCenter/Bottom` = `ksoCmd="TableCellAlignVertical"` 的 index 0/1/2），再横向三枚
 * （`AlignLeft/HCenter/Right` = `ksoCmd="TableCellAlignHorizontal"` 的 index 0/1/2）。第三枚 `m_TableTextDirection`
 * （:169，子清单 :72 带 `dialogHint="1"`）**不摆**：`w:tcPr/w:textDirection` 在授权 46 份里 0 处，那一击落什么
 * 字面也未取证——与 #187 页面布局那一页的文字方向同一口径，宁缺一条不摆一条按下去说不出话的。
 * 落点两层（两行各挂一条 `ksoCmd` 说的就是这个）：竖向那句归格（`w:tcPr/w:vAlign`），横向那句归格里的每一段
 * （`w:pPr/w:jc`）。命令各走各的（`setCellVAlign` / `setCellAlign`），回显问同一份读数（`cellAlignState`）。
 * 字面只取到 `@Align Top` 这一族英文占位，中文措辞在 `.qm` 里读不出，所以按钮照开始页那三枚用同一副字形。
 */
const TABLE_ALIGN_BUTTONS: [string, string, string, string, "v" | "h"][] = [
  ["tblTop", "↑", "顶端对齐", "top", "v"],
  ["tblMid", "↕", "垂直居中", "center", "v"],
  ["tblBot", "↓", "底端对齐", "bottom", "v"],
  ["tblLeft", "左", "左对齐", "left", "h"],
  ["tblCenter", "中", "居中", "center", "h"],
  ["tblRight", "右", "右对齐", "right", "h"],
];

/** 纸张方向那一格的两档；WPS 只声明了 `PageOrientation` 这一枚命令，它的Widget 是下拉还是开关**未取证**，这里按可回显的两档摆。 */
const PAGE_ORIENTATIONS: [string, string][] = [
  ["portrait", "纵向"],
  ["landscape", "横向"],
];

/**
 * 一节现在的整份几何摊成一份 `PageSetup`，再只换调用方要改的那一两项：页边距那一格不该顺手改纸，
 * 换方向也不该顺手把两条边写成另一对（`paperTwips` 不给就是「纸还是那张，只是转方向」）。
 */
const pageSetupOf = (s: Section, patch: Partial<PageSetup>): PageSetup => ({
  topTwips: s.margins.top,
  rightTwips: s.margins.right,
  bottomTwips: s.margins.bottom,
  leftTwips: s.margins.left,
  headerTwips: s.margins.header,
  footerTwips: s.margins.footer,
  gutterTwips: s.margins.gutter,
  landscape: s.pageWidthTwips > s.pageHeightTwips,
  ...patch,
});

/**
 * 分页 ribbon 的壳：一条页签 + 每页一行组，右边一枚溢出箭头。文字页与表格页共用。
 * 装不下的那一档要给一枚看得见的箭头：不给，用户看到的是一枚被切一半的按钮，说不出「这一行右边还有东西、能滚」。
 * WPS 那边「单行放不下」确有声明，但形状不同：装机件 `res/commands/wps/ongmani/CT_*.kuip` 量到 515 个
 * `<K*ListCommand>` 组声明，其中 30 个带 `needsinglerowmorebutton="true"`（如 `CT_Home.kuip:105 RB_Font`、
 * `:367 RB_Paragraph`）——它声明在**组**上，这一枚挂在**整行**右端，是已知差别，不按同物引用。
 * `res/wpsongmani.kui:113` 的 tab 条上另有 `idKey="CommonLogicGroupShowArrow"`：名字带 ShowArrow，指的是档名条
 * 还是组行**未取证**。滚到最右那一枚自己退场——最后一组整个进来了，留着只会挡字。
 */
function ribbonTabs(list: readonly [string, string][]): {
  group: (tab: string) => HTMLElement;
  showTab: (key: string) => void;
  bars: Record<string, HTMLElement>;
  tabs: Record<string, HTMLButtonElement>;
  shown: () => string;
  sync: () => void;
} {
  const strip = document.createElement("div");
  strip.className = "rtabs";
  ribbon.append(strip);
  const bars: Record<string, HTMLElement> = {};
  const tabs: Record<string, HTMLButtonElement> = {};
  let shown = list[0]![0];
  const more = document.createElement("button");
  more.className = "rover";
  more.textContent = "›";
  more.title = "这一档右边还有命令，点一下滚过去";
  const sync = (): void => {
    const bar = bars[shown];
    const room = bar.scrollWidth - bar.clientWidth;
    more.hidden = room <= 1 || bar.scrollLeft >= room - 1;
  };
  more.onclick = () => {
    const bar = bars[shown];
    bar.scrollLeft += Math.max(120, Math.floor(bar.clientWidth * 0.8));
    sync();
  };
  // 这一枚是随重建生成的：`buildRibbon()` 每趟先 `innerHTML = ""`，监听不跟着收就会一次切档多留一个。
  overflowAbort?.abort();
  overflowAbort = new AbortController();
  window.addEventListener("resize", sync, { signal: overflowAbort.signal });
  const showTab = (key: string): void => {
    shown = key;
    for (const [k] of list) {
      bars[k]!.hidden = k !== key;
      tabs[k]!.classList.toggle("on", k === key);
      tabs[k]!.setAttribute("aria-selected", k === key ? "true" : "false");
    }
    bars[key]!.scrollLeft = 0;
    sync();
  };
  for (const [key, label] of list) {
    const bar = document.createElement("div");
    bar.className = "rbar";
    bar.dataset.tab = key;
    bar.hidden = true;
    bars[key] = bar;
    ribbon.append(bar);
    bar.addEventListener("scroll", sync, { passive: true });
    const t = document.createElement("button");
    t.className = "rtab";
    t.textContent = label;
    t.title = `${label}：这一页的命令`;
    t.onclick = () => showTab(key);
    tabs[key] = t;
    strip.append(t);
  }
  // 箭头排在各档之后（它自己不在 `.rbar` 里，所以不占任何一组的控件数）。
  ribbon.append(more);
  const group = (tab: string): HTMLElement => {
    const g = document.createElement("div");
    g.className = "grp";
    bars[tab]!.append(g);
    return g;
  };
  showTab(shown);
  return { group, showTab, bars, tabs, shown: () => shown, sync };
}

/** Word-style formatting: run styling over the selection, paragraph shape for the whole paragraph. */
function buildDocRibbon(): void {
  // WPS 把文档命令分在几页上：`res/wpsongmani.kui:114-122` 在 `<KRbTabWidgetCommand id="tabwidget">` 里按序声明
  // CT_Home → CT_Insert → CT_PageLayout → CT_References → CT_Review → CT_View，每支 `CT_*.kuip` 只列自己那页的组。
  // 自家这 11 组横着排到头，装机件里量到的是内容 2,664 px 对可视 1,280 px——「页面设置」往后的那几枚根本点不到。
  // 所以同一种分法：一条 tab 条 + 每页一行组。组归哪页按 WPS 那份组清单（`CT_View.kuip` 的 `RB_Show`/`RB_Zoom`
  // 挂着 `sp_ShowDocumentMap` 与 `ZoomToolBar`，导航窗格和缩放就归视图页）。引用页没挂：家里还没有它的命令。
  // 页名用自家的话——`.kuip` 只给英文 source（`text="@Home"`），中文界面字面在读不出来的 `.qm` 里，仍**未取证**。
  const shell = ribbonTabs(DOC_RIBBON_TABS);
  // 表格工具页是**上下文页**：`syncDocRibbon` 每趟按光标问它该不该露，露着的那一页在光标离开时要退回开始页。
  // 那个函数在模块那一层，所以开关得从这里交出去。
  tableTab = { button: shell.tabs.table!, shown: shell.shown, show: shell.showTab };
  // 组名按 WPS 的组归属写（`CT_Home.kuip` 的 RB_Clipboard/RB_Font/RB_Paragraph/RB_Style/RB_Editing 等），
  // 中文措辞是我们自己的；次序与下面建组的次序一一对应。
  const caps: Record<string, string[]> = {
    home: ["常用", "字体", "段落", "列表", "样式", "编辑"],
    insert: ["页面", "插图", "批注"],
    table: ["表格", "行与列", "合并与拆分", "单元格尺寸", "字体", "对齐", "数据", "选择"],
    review: ["校对", "批注", "修订", "更改"],
    layout: ["页面设置", "页眉页脚"],
    view: ["显示"],
  };
  const group = (tab: string): HTMLElement => {
    const g = shell.group(tab);
    const cap = caps[tab]?.shift();
    if (cap) g.dataset.cap = cap;
    return g;
  };
  const button = (g: HTMLElement, key: string, label: string, tip: string, run: () => void): void => {
    const b = document.createElement("button");
    b.className = "rb";
    b.textContent = label;
    b.title = tip;
    decorateRibbonButton(b, label);
    b.onclick = run;
    ctl[key] = b;
    g.append(b);
  };
  const chooser = (g: HTMLElement, key: string, tip: string, options: readonly [string, string][], run: (value: string) => void): void => {
    const s = document.createElement("select");
    s.className = "pick";
    s.title = tip;
    for (const [value, label] of options) {
      const o = document.createElement("option");
      o.value = value;
      o.textContent = label;
      s.append(o);
    }
    s.onchange = () => run(s.value);
    ctl[key] = s;
    g.append(s);
  };
  const swatch = (g: HTMLElement, key: string, tip: string, value: string, run: (value: string) => void): void => {
    const i = document.createElement("input");
    i.type = "color";
    i.className = "swatch";
    i.title = tip;
    i.value = value;
    i.oninput = () => run(i.value);
    ctl[key] = i;
    g.append(i);
  };
  /**
   * ribbon 上的数字格：`label` 在外头，所以控件本体是 `input` 自己（`ctl` 存它，回显与「正在打的那格别动」都问它）。
   * 算的是**落笔**那一刻（`change`：失焦或回车），不是每按一个字符——半个数就改页边距，会把件写成没人要说的尺寸。
   */
  const spin = (g: HTMLElement, key: string, label: string, tip: string, run: (text: string) => void): void => {
    const box = document.createElement("label");
    box.className = "rspin";
    const i = document.createElement("input");
    i.type = "text";
    i.inputMode = "decimal";
    i.autocomplete = "off";
    i.title = tip;
    i.onchange = () => run(i.value);
    // `data-echo` 记下这一格上一趟被刷成的字面（`echoSpin` 是唯一的写手）。空格子的初始值是 `""`，两者一起立，
    // 第一趟回显才认得出「框里还是我们那句」。
    i.dataset.echo = "";
    box.append(label, i);
    ctl[key] = i;
    g.append(box);
  };
  const style = (patch: Parameters<DocEditor["applyStyle"]>[0]): void => {
    editor?.applyStyle(patch);
    syncDocRibbon();
  };

  const nav = group("home");
  // WPS 开始页第一组是 `RB_Clipboard`（`CT_Home.kuip:23-30`：粘贴/剪切/复制/格式刷），格式刷就挂在这一组里
  // （`:25` 的 `<command idQ="FormatPainter"/>`，`wpsentries.kuip` 的 `CT_Home.RB_ClipboardWrapper.FormatPainter`）。
  // 家里这一组是「常用」的替身，把格式刷摆在撤销前面；粘贴/剪切/复制在右键菜单里已有入口。
  button(nav, "painter", "格式刷", "把光标处的字符格式刷到下一段选中的文字（Esc 放下）", () => {
    editor?.copyFormat();
    syncDocRibbon();
  });
  button(nav, "undo", "撤销", "撤销 ⌘Z", () => undo());
  button(nav, "redo", "重做", "重做 ⌘⇧Z", () => redo());
  button(nav, "", "全选", "选中整篇", () => selectAllDoc());

  const font = group("home");
  // 字体框与字号框在 WPS 里是一对：`CT_Home.kuip:110-111` 把 `FontNameInHomeTab` 与 `FontSize` 同挂在 `RB_Font` 的
  // `font_info_pair` 里，字体在前、`minimumWidth="130"`。表格工具页那一组（`CT_TableTools.kuip:133`）是同一对里的
  // 另一条声明 `FontNameInTableToolTab`，只把 `minimumWidth` 报成 112——两页各一枚框，共用的只有这张清单与落笔。
  // 它是可编辑的一格：打名字回车，或从下拉里挑。下拉列的是这份文件自己用过的脸——Electron 没有「本机装了哪些
  // 字体」的接口可问，WPS 那份装机字体清单在此仍未取证，所以只把文档说过的名字递给用户，不替它编一张表。
  const faceInput = (key: string, px: number): HTMLInputElement => {
    const box = document.createElement("input");
    box.type = "text";
    box.className = "facebox";
    box.title = "字体";
    box.placeholder = "字体";
    box.style.width = `${px}px`;
    box.setAttribute("list", FACES_ID);
    box.onfocus = () => fillFaceList();
    box.onchange = () => {
      const name = box.value.trim();
      // 清空那一格不当命令办：这一页说不出「这段取消脸」该往件里写哪一句（「没提过」只能靠原件那半边的字节带回）。
      // 旁边那一枚「清除格式」撤的是整组直接格式，不是「只把这一格清空、别的都留着」；真机 WPS 那一格被清空后
      // 写什么仍未取证。
      if (name) style({ font: name });
    };
    ctl[key] = box;
    return box;
  };
  const faceBox = faceInput("font", 130);
  font.append(faceBox);
  /**
   * 字号那一格在 WPS 里是 `KxWpsFontSizeComboBoxCommand`：下拉是 `fontSizeList` 那 21 档，框本身还能打字。
   * 所以这里不是 `<select>` 而是 `input[list]`——`datalist` 给 21 档、键入别的磅数照样落笔（按半点取整，
   * 见 `parseSizePt`）。「能打字」是照 ComboBox 这个类名推的，WPS 真机那一格的手感未取证。
   * 打进来的不是磅数就不落笔，框退回这份文件此刻说的那个数：替它编一个数就是往件里写字。
   */
  const sizeBox = (key: string): HTMLInputElement => {
    const box = document.createElement("input");
    box.type = "text";
    box.className = "sizebox";
    box.title = "字号";
    box.placeholder = "字号";
    box.inputMode = "decimal";
    box.autocomplete = "off";
    // 21 档是静态清单、不随文档变，所以建框时就把 datalist 摆好：等聚焦那一下再挂，`list` 属性在聚焦前指的
    // 是一个还不存在的 id（读端量 DOM 也只看到空下拉）。
    sizeList();
    box.setAttribute("list", SIZES_ID);
    box.onchange = () => {
      const pt = parseSizePt(box.value);
      if (pt === undefined) {
        box.value = sizeEcho();
        return;
      }
      style({ sizePt: pt });
      // 落完笔回显按同一句号数制（打得进「12」也显示「小四」），与光标回显一个口径。
      box.value = sizeNameOf(pt) ?? String(pt);
    };
    ctl[key] = box;
    return box;
  };
  font.append(sizeBox("size"));
  // 「更改大小写」是 `RB_Font` 里那枚折叠钮 `Font1`（`CT_Home.kuip:121`，`referenceCommand="ChangeCaseRainbow"`、
  // `idMso="ChangeCaseGallery"`、`ksoCmd="ChangeCase"`、`rbHotKey="7"`）。左邻是字号那几枚（`FontSize:111` 与放大/缩小
  // `:115`/`:117`），右邻是 `ClearFormatsRainbow:122`——这两枚同挂在 `:120-123` 那一枚 `KRbLayoutListCommand` 里。中文界面字面（「更改大小写」「全部大写」「切换大小写」「词首字母大写」「小写」）在
  // `mui/zh_CN/wps.qm` 里都有，但那张池子只证明字在包里有，不证明下拉就摆这五档、也不证明次序：档位按 Word 那五档，
  // 「小型大写字母」那一枚（字面同在 `wps.qm`，落的是 `w:smallCaps` 那种格式而不是改字面）本片不接。
  // 这一格与「转换成文本」同形状：选方向的一次性命令，落完笔回到占位，不当开关、不回显。
  chooser(font, "changeCase", "更改大小写（WPS：Change Case）", CASE_OPTIONS, (v) => {
    (ctl.changeCase as HTMLSelectElement).value = "";
    if (v) editor?.changeCase(v as CaseMode);
    syncDocRibbon();
  });
  // 组内次序照 WPS `CT_Home.kuip` 的 `RB_Font`（同一份里的 id 序）：字体名 → 字号 →（放大/缩小）→ 更改大小写 →
  // 清除格式 → 加粗 → 斜体 → 下划线 →（上标/下标、文字效果）→ 突出显示 → 字体颜色。
  // 清除格式是 `ClearFormatsRainbow`（`idMso="ClearFormatting"`、`ksoCmd="ClearFormats"`、`icon="eraser_kd"`、
  // `hotKey="E"`），与 `Font1` 同挂在 `:120-123` 那一条布局里、在它之后；中文界面字面「清除格式」见
  // `mui/zh_CN/wps.qm`（`.qm` 是 UTF-16BE，只用来确认这四个字，不搬任何资源）。
  button(font, "clearfmt", "清除格式", "清除所选文字的直接格式（字体/字号/粗细/斜体/下划线/删除线/颜色/突出显示/上下标/字距）；段落格式与段落样式不动", () => {
    editor?.clearFormat();
    syncDocRibbon();
  });
  button(font, "bold", "粗", "加粗 ⌘B", () => style({ bold: !editor?.markState().bold }));
  button(font, "italic", "斜", "斜体 ⌘I", () => style({ italic: !editor?.markState().italic }));
  button(font, "under", "线", "下划线 ⌘U", () => style({ underline: editor?.markState().underline ? "none" : "single" }));
  button(font, "strike", "删", "删除线", () => style({ strike: !editor?.markState().strike }));
  // 上标/下标在 WPS 的 `RB_Font` 里是**一枚折叠钮**：`CT_Home.kuip:136` 那枚 `Font4`（`customTip="Superscript"`、
  // `icon="math_superscript_kd"`、`singlerowfold="1"`）把 `Superscript`（:50）与 `Subscript`（:51）两枚折在一起，
  // 组内声明序是 `Underline`（:133）→ `Font2` 删除线（:135）→ `Font4` 上下标（:136）→ `FontEffects`（:142）→
  // `HighlightColor`（:147）——正是这一行现在的次序。折叠钮的那一半（下拉里的另一档、`FontEffects` 文字效果）家里
  // 没有对应的控件形状，这里按两枚平铺：折进去的两只腿就是这两条命令，不另造一枚假下拉。两枚都是开关：按已按下的
  // 那枚＝撤回（上标再按下标＝换成下标，那是另一枚的事）。落笔与回显走 `applyStyle`/`markState()` 同一条路。
  button(font, "sup", "上标", "上标", () => style({ vertAlign: editor?.markState().vertAlign === "superscript" ? "" : "superscript" }));
  button(font, "sub", "下标", "下标", () => style({ vertAlign: editor?.markState().vertAlign === "subscript" ? "" : "subscript" }));
  chooser(
    font,
    "highlight",
    "突出显示",
    DOC_HIGHLIGHTS,
    // No reset: the box is the caret's sentence, not a one-shot command. `syncDocRibbon` reads the run the
    // caret sits in, so what it shows after a pick is what the document now says there.
    (v) => {
      editor?.applyStyle({ highlight: v });
      syncDocRibbon();
    },
  );
  swatch(font, "color", "字体颜色", "#c00000", (v) => style({ color: v }));

  const align = group("home");
  button(align, "left", "左", "左对齐", () => editor?.setParagraph({ align: "" }));
  button(align, "center", "中", "居中", () => editor?.setParagraph({ align: "center" }));
  button(align, "right", "右", "右对齐", () => editor?.setParagraph({ align: "right" }));
  button(align, "both", "两", "两端对齐", () => editor?.setParagraph({ align: "both" }));
  // WPS 表格「对齐方式」的竖向那一列（`w:tcPr/w:vAlign`）：横向那一句落段上，这三枚落格上。授权 46 份里这一枚
  // 4,130 处，页面早就在画，缺的是改它的手。光标不在表里时三枚灰着——段落没有「格子」可竖直摆。
  // 这一页与「表格工具」那一页那六枚走同一条命令、同一份回显（`cellAlignState`），对着同一份文档不各说一句话。
  button(align, "vtop", "↑", "顶端对齐", () => tableAction(editor?.setCellVAlign("top")));
  button(align, "vmid", "↕", "垂直居中", () => tableAction(editor?.setCellVAlign("center")));
  button(align, "vbot", "↓", "底端对齐", () => tableAction(editor?.setCellVAlign("bottom")));
  // WPS 把这两条放在主页段落组里（`IncreaseIndent`/`DecreaseIndent`，图标 `indent_increase_kd`/`indent_decrease_kd`，
  // Touch Bar 图标的中文标题写的就是「增加缩进量」「减少缩进量」），右键菜单里也有同名的一对。
  button(align, "outdent", "缩进−", "减少缩进量", () => editor?.shiftParagraphIndent(-1));
  button(align, "indent", "缩进+", "增加缩进量", () => editor?.shiftParagraphIndent(1));
  // 行距那一格与上面两条同属 `CT_Home.RB_Paragraph`。五档倍数按 `wpsentries.kuip` 的 `m_LineSpacing` 摆，
  // 最后一条 `LineSpacingMore` 开「段落」对话框——段前/段后与 最小值/固定值 没有下拉档，只能从那里说。
  chooser(
    align,
    "linespacing",
    "行距",
    [...DOC_LINES.map((n): [string, string] => [String(n), n.toFixed(1)]), ["more", "行距…"]],
    (v) => {
      if (v === "more") {
        editor?.openParaSpacing();
        syncDocRibbon();
        return;
      }
      // 只有那五档是命令：回显会把「件里说的那一档」插成一个额外选项，而一段没写行距的段落那一格是空着的
      // （`selectedIndex = -1`，值是空串）。`Number("")` 是 0，落进 `w:line` 就是一句「这一段的行盒零高」，
      // 所以认不出来的字面不当命令办。
      const mult = DOC_LINES.find((n) => String(n) === v);
      if (mult === undefined) return;
      editor?.setParagraphSpacing({ line: Math.round(mult * LINE_ONE), lineRule: "auto" });
      syncDocRibbon();
    },
  );
  // 边框那一格与行距同属 `CT_Home.RB_Paragraph`（`res/wpsentries.kuip:669` 把 `RbTableOutsideBorder` 挂在段落组），
  // 下拉里腿的先后照 `res/commands/wps/ongmani/CT_Home.kuip:272-287` 的 `<separator/>` 分段摆。列子随选区换
  // （`BORDER_LEGS`、`syncDocRibbon`），这一处只是建时的初始那张——光标落在正文上时问的就是段落那一圈。
  chooser(
    align,
    "paraBorder",
    "边框",
    BORDER_LEGS.para,
    (v) => {
      // 这一格是**命令的菜单**，不是这一段的状态：落完放回「边框」那一档，下一次点同一条腿才还是发同一条命令
      // （`syncDocRibbon` 不认领这一格，所以没人替它回显，也不需要——一层的框是逐边的，说不出「哪一档」）。
      (ctl.paraBorder as HTMLSelectElement).value = "";
      if (!v || !editor) return;
      // 「边框与底纹…」这一腿自己按选区问层（#150）：段落、整张表、所选单元格三副样子，不在这里替它分。
      if (v === "more") {
        editor.openBorderDialog();
        return;
      }
      const scope = editor.borderScope();
      if (scope === "para") {
        editor.setParagraphBox(legPatch(v as BorderLeg));
        syncDocRibbon();
        return;
      }
      // 表那一侧的每一腿都落 `w:tblBorders`；格那一侧两条例外——内部横竖与横竖框线 `CT_TcBorders` 说不出这一句，
      // 照 WPS 一样由表来说；外侧框线落在选区外圈（`setCellOutside`，没有铺满整张表时）。
      if (scope === "table") editor.setTableBox(tableLegPatch(v as TableLeg));
      else if (v === "cross" || v === "insideH" || v === "insideV") editor.setTableBox(tableLegPatch(v));
      else if (v === "outside") editor.setCellOutside();
      else editor.setCellBox(cellLegPatch(v as CellLeg));
      syncDocRibbon();
    },
  );

  // WPS 在同一个段落组（`tb_Paragraph`，`CT_Home.kuip:337`）里把 `Tabs` 排在 `Sort` 之后，也就是那一簇的最后一条；
  // 它自己声明的就是这张对话框：`CT_Home.kuip:309` 的 `<KxLegacyTriggerCommand aid="780" customTip="Tabs"
  // dialogHint="1" idMso="Tabs" ksoCmd="Tabs" icon="ruler_format_tabs_kd" text="@Tabs"/>`——`dialogHint="1"` 说明
  // 这一条按下开的是**一张对话框**，不是一个开关。这张页面上没有可回显的「哪一档」：一段的制表位是一列位置，
  // 说不出单选，所以这一枚不认领 `ctl`（同查找替换那两枚）。
  // WPS 的 `m_ShowHideEditMark`（`CT_Home.kuip:337`，与制表位同一组）声明了三条子项：段落标记、段落布局、小十字。
  // 只接了第一条：段落标记是**视图档**（见 `showEditMarks`），后两条画成什么样未取证，不摆。
  // 次序照 `CT_Home.kuip:314-315`：`m_ShowHideEditMark` 在前、`Tabs` 在后，所以这一枚要压在制表位之前。
  button(align, "editmarks", "¶", "显示或隐藏编辑标记：段落末尾的 ¶（WPS：显示/隐藏编辑标记 · 段落标记）；只影响屏幕，不改文件", () => {
    showEditMarks = !showEditMarks;
    syncDocRibbon();
  });
  button(align, "", "制表位", "制表位：这一段那一列位置、对齐与前导符，加件级的默认制表位", () => editor?.openTabStops());

  const lists = group("home");
  button(lists, "bullet", "• 符号", "项目符号", () => editor?.setParagraph({ list: "bullet" }));
  button(lists, "number", "1. 编号", "编号列表", () => editor?.setParagraph({ list: "number" }));
  button(lists, "nolist", "取消列表", "变回普通段落", () => editor?.setParagraph({ list: "none" }));

  // 开始页「样式」：`RB_Style` 那一组，位置照 `CT_Home.kuip` 那份 idQ 序——在 `RB_Paragraph` 之后、`RB_Editing`
  // 之前（见下一条注释）。列的是**这份文件自己声明的**段落样式（`WordDocument.paragraphStyles`），不替它编名字：
  // 授权 46 份共 112 枚段落样式，每份 6～27 枚（`.scratch/para-style-gallery-audience.py`）。第一格是默认段落样式，
  // 选它就是撤掉 `w:pStyle`（「正文」）；清单在 `syncDocRibbon` 里按内容签名重建，换文档那一趟（`force`）一定刷。
  const styleGroup = group("home");
  const stylePick = document.createElement("select");
  stylePick.className = "pick";
  stylePick.title = "样式";
  stylePick.onchange = () => {
    editor?.setParagraphStyle(stylePick.value === "" ? undefined : stylePick.value);
    syncDocRibbon();
  };
  ctl.style = stylePick;
  styleGroup.append(stylePick);

  // 开始页最后一组是 `RB_Editing`（`CT_Home.kuip:441`，`groupTip="FindGroup"` `idMso="GroupEditing"`），
  // 它在 `CT_Home.kuip:585-600` 那条 idQ 序里排在 `RB_Style` 之后——正是「最后一组」这个位置。
  // 组成员里家里有命令的是 `sp_Find`（:415 `hotKey="FD" shortcutHint="Ctrl+F" text="@Find and &#10;Replace"`，
  // 下拉里 `FindRainbow → ReplaceRainbow → GoTo`）：查找与替换在家里同在一个面板上（`openFind` 那格既有「查找」也有
  // 「替换／全部替换」），所以摆两枚——查找替换、定位（`⌘G` 已是 `editor.showGoTo()`）。
  // 拼写检查族（`CheckSpellToolBarList_Home`）与翻译族家里没有命令面，不摆空命令；§122 下一步② 那条同因。
  const editing = group("home");
  button(editing, "", "查找替换", "查找与替换 ⌘F", openFind);
  button(editing, "", "定位", "跳到页 / 节 / 行 / 书签 / 批注 / 表格 / 图形 / 标题 ⌘G", () => editor?.showGoTo());

  // 表格工具的上下文页。WPS 的组序是 `<subcommand id="CT_TableTools">`（`CT_TableTools.kuip:180-198`）说死的：
  // RB_Table → RB_RowsAndColumns → RB_Merge → RB_Adjust → RB_Table_Font → RB_Alignment → RB_Data → TableSelectMenu。
  // 这里摆家里真做得了的组，组名用自家的话，**位置按上面那个组序**（缺的是中间那几组，不是摆好的这几组的位）：
  //   · 表格 ← RB_Table 的两枚（`TableProperties` 属性对话框、`ShowGridlines` 视图那一档，见下面那枚 `props`）；
  //   · 行与列 ← RB_RowsAndColumns 里的 `TableTools_Insert`（上方/下方插入行、左侧/右侧插入列、Cells…）
  //     与 `Delete` 那档四条（:33-36：单元格、列、行、表格）；
  //   · 合并与拆分 ← RB_Merge 的 Merge Cells / Split Cells / Split Table by Row（按列拆仍是 0 命令）；
  //   · 单元格尺寸 ← RB_Adjust 的高度/宽度两格（见 `CELL_SIZE_SPINS`）与「分布行」「分布列」那两枚（`DistributeRows` / `DistributeColumns`）；
  //   · 对齐 ← RB_Alignment 的两行六枚（见 `TABLE_ALIGN_BUTTONS`）；
  //   · 数据 ← RB_Data 的「标题行重复」与「转换成文本」；
  //   · 选择 ← TableSelectMenu 的四条。
  // RB_Table 的两枚都接上了；RB_Data 的公式（`FieldFormula`）与 RB_Table_Font 仍未接，这一轮不摆空按钮——
  // 宁缺一条，不摆一条按下去什么也不改的。
  // `RB_Table`（`CT_TableTools.kuip:5-9`，`groupTip="TableToolPropertyGroup"`）在 WPS 那条 `<subcommand>` 组序里排
  // **第一**，组里两枚：`TableProperties`（:7）与 `ShowGridlines`（:8，`idMso="TableShowGridlines"`）。两枚都摆，次序照它：
  //  · 「属性」按下开的是「表格属性」对话框（`doc-table-props.ts`）。WPS 那张对话框里分几页、摆哪些字段**未取证**，
  //    所以这里不按 Word 的四页抄，只摆件里真有受众、又已经有落笔处的那四句：对齐 `w:jc`（14 枚/4 件，
  //    全是 `center`）、左缩进 `w:tblInd`（58 枚/10 件）、固定列宽 `w:tblLayout`（58 枚/10 件，全是 `fixed`）、
  //    默认单元格边距 `w:tblCellMar`（58 枚/10 件，`.scratch/tblpr-children-audience.py` 的读数）。
  //    宽度、边框、底纹不在这张对话框里：它们各自已经有落笔处（`setTableWidth`、`#145/#150/#152` 那张对话框）。
  //  · 「查看网格线」说的是**屏幕上画不画导引线**这一句，文件里没有它（下面那枚 `showGridlines` 说清了为什么），
  //    家里画得出来也回得了显，所以摆。
  const props = group("table");
  button(props, "tableProps", "属性", "表与行的对齐、左缩进、固定列宽与默认单元格边距（WPS：Table Properties）", () => editor?.openTableProps());
  button(props, "gridlines", "查看网格线", "把文件里没说过线的那些格子边，在屏幕上画成导引线（WPS：View Gridlines）；导引线上不了纸", () => {
    showGridlines = !showGridlines;
    syncDocRibbon();
  });

  const tbl = group("table");
  /**
   * 结构命令的答案：拒绝时把**命令自己说的那一句**摆出来，然后无论成败都重读一遍 ribbon。那一条带子上每一枚
   * 命令都回这一种形状（`TableCommandAnswer`），所以这一处不需要替谁编一句兜底话——兜底只在命令确实没话说时给。
   */
  const tableAction = (out: TableCommandAnswer | undefined): void => {
    if (out && !out.ok) notify(out.reason ?? "这一步在当前的表格形状里做不了");
    syncDocRibbon();
  };
  button(tbl, "rowAbove", "上方插入行", "上方插入行（WPS：Rows to the Above）", () => tableAction(editor?.insertRow("above")));
  button(tbl, "rowBelow", "下方插入行", "下方插入行（WPS：Rows to the Below）", () => tableAction(editor?.insertRow("below")));
  button(tbl, "colLeft", "左侧插入列", "左侧插入列（WPS：Columns to the Left）", () => tableAction(editor?.insertColumn("left")));
  button(tbl, "colRight", "右侧插入列", "右侧插入列（WPS：Columns to the Right）", () => tableAction(editor?.insertColumn("right")));
  chooser(tbl, "insertCellKind", "插入单元格（WPS：Cells…）", [["", "插入单元格"], ["right", "右侧插入"], ["down", "下方插入"]], (v) => {
    (ctl.insertCellKind as HTMLSelectElement).value = "";
    if (v) tableAction(editor?.insertCell(v as "right" | "down"));
  });
  button(tbl, "delRow", "删除行", "删除行（WPS：Delete Rows）", () => tableAction(editor?.deleteRow()));
  button(tbl, "delCol", "删除列", "删除列（WPS：Delete Columns）", () => tableAction(editor?.deleteColumn()));
  // `CT_TableTools.kuip:32-36` 那一档四条现在接齐了：行、列、表格、单元格。第四条与「插入单元格」同形状——
  // WPS 那一枚带 `dialogHint="1"`（:33 `DeleteCells`），家里没有对话框，就照插入那一条做一枚选方向的格子。
  // 落下的形状按自家已经说出的那句话定（不跟踪=行内吸收，跟踪=一句 `w:cellDel` 待点头）；WPS 自己写什么字节仍未取证。
  button(tbl, "delTable", "删除表格", "删除表格（WPS：Delete Table）", () => tableAction(editor?.deleteTable()));
  chooser(tbl, "deleteCellKind", "删除单元格（WPS：Delete Cells）", [["", "删除单元格"], ["left", "右侧左移补上"], ["up", "下方上移补上"]], (v) => {
    (ctl.deleteCellKind as HTMLSelectElement).value = "";
    if (v) tableAction(editor?.deleteCell(v as "left" | "up"));
  });

  const merges = group("table");
  button(merges, "mergeCells", "合并单元格", "合并单元格（WPS：Merge Cells）", () => tableAction(editor?.mergeCells()));
  button(merges, "splitCell", "拆分单元格", "拆分单元格（WPS：Split Cells）", () => tableAction(editor?.splitCell()));
  button(merges, "splitTable", "按行拆分表格", "按行拆分表格（WPS：Split Table by Row）", () => tableAction(editor?.splitTable()));

  // `RB_Adjust` 那两格按自家已经说出的两层落：高度那一格给选区碰到的**每一行**（`w:trHeight` 是行说的），
  // 宽度那一格给选区碰到的**每一整列**（网格加那一列每格那句 `w:tcW`），不是光标那一格。落笔口径照页边距四格：
  // 按状态栏的测量单位算，失焦或回车才算；说不出数的字面（空的、半个、负）当场刷回件里那句，不留在框里装成生效过。
  const sizes = group("table");
  for (const [key, label, dim] of CELL_SIZE_SPINS) {
    spin(sizes, key, label, `${label}按状态栏的测量单位填；落笔（失焦或回车）才算`, (text) => {
      const v = Number(text.trim());
      if (text.trim() === "" || !Number.isFinite(v) || v < 0) {
        echoSpin(ctl[key] as HTMLInputElement, editor?.cellSizeState()?.[dim]);
        return;
      }
      const twips = twipsFromUnit(v, currentUnit());
      tableAction(dim === "heightTwips" ? editor?.setCellHeight(twips) : editor?.setCellWidth(twips));
      // 落笔之后自己读回件里**现在**那句（`tableAction` 已重跑一趟回显，但那一趟会护住「还没落笔」的框，
      // 而这一格的字面已经落笔了）：成了报刚写的数，被回绝报件里没动的那个数——两条命令共用这一条出口。
      echoSpin(ctl[key] as HTMLInputElement, editor?.cellSizeState()?.[dim]);
    });
  }
  // `RB_Adjust` 组里那两枚「分布」：`DistributeRows`（`CT_TableTools.kuip:53`，`idMso="TableRowsDistribute"`）与
  // `DistributeColumns`（:54，`idMso="TableColumnsDistribute"`），两枚都挂在 :58 那枚 `TableAutoAdjust` 清单里
  // （:64/:65 的 `idQ`），与上面两格同组，声明序是行在前、列在后。中文措辞仍读不出（`.qm` 那一族按 #195 的口径
  // 只有在场、没有可读的字面），所以按钮照「单元格尺寸」那两格用自家的话。落下的是**等宽/等高而那一维的和不变**：
  // 见 `doc-table-ops.ts:distributeColumns` 与 `:distributeRows`。
  button(sizes, "distributeRows", "分布行", "把选区碰到的那几行摆成等高，表的总高不变（WPS：Distribute Rows Evenly）", () =>
    tableAction(editor?.distributeRows()));
  button(sizes, "distributeCols", "分布列", "把选区碰到的那几列摆成等宽，表的总宽不变（WPS：Distribute Columns Evenly）", () =>
    tableAction(editor?.distributeColumns()));

  // `RB_Table_Font`（`CT_TableTools.kuip:129`：`groupTip="TableToolFontGroup"`、`ksoCmd="Font"`、`customTip="RB_Font"`、
  // `horizontalSpacing="2"`、`needsinglerowmorebutton="true"`）在这份文件的组序里正夹在 `RB_Adjust` 与 `RB_Alignment`
  // 中间（:129 早于 :156 那枚 `RB_Alignment`），所以上面那组自旋格之后、下面那六枚对齐之前就是它的位置。
  // 组里七枚的实据（两行）：第一行 `FontNameInTableToolTab`（:133，`ksoCmd="FontName"`、`minimumWidth="112"`、
  // `singlerowfoldgroup="font_info_pair"`）与 `command idQ="FontSize"`（:134，引 :102 那一枚）；第二行 `Bold`（:138→:112）、
  // `Italic`（:141→:113）、`Underline`（:144→:114）、`FontColor`（:147→:108）、`HighlightColor`（:150→:105），后三枚各占
  // 一格 `fixedwidgetWidth="40"`。摆不完的三条也是声明里的账：`CustomBoldWidth`/`OfficialPrintingCustomBoldWidth`
  // （:139-140）说的是加粗的笔画宽度，家里没有那一笔；`Underline`（:114）在 WPS 是一枚带 8 档线样式的 gallery
  // （外加 `MoreUnderlines` :123、`UnderLineColor` :124），家里按开始页那一枚的开关做，落 `w:u w:val="single"`。
  // 「删除线」在这个文件里声明了（:103 `Strikethought`）却**没进这一组**，「着重号」（:104 `EmphasisMark`）同样——
  // 所以这一组比开始页少一枚「删」，那是 WPS 的组形状，不是家里缺的一条命令。
  // 落笔与回显都走开始页那一条路（`applyStyle` → `w:rPr` 里的 `w:b/w:i/w:u/w:color/w:highlight/w:rFonts/w:sz`；
  // `markState()` 一份读数喂两页的框），WPS 那边也是同一枚 `ksoCmd` 挂在两处。
  const tableFont = group("table");
  tableFont.append(faceInput("tblFace", 112));
  tableFont.append(sizeBox("tblSize"));
  button(tableFont, "tblBold", "粗", "加粗 ⌘B", () => style({ bold: !editor?.markState().bold }));
  button(tableFont, "tblItalic", "斜", "斜体 ⌘I", () => style({ italic: !editor?.markState().italic }));
  button(tableFont, "tblUnder", "线", "下划线 ⌘U", () => style({ underline: editor?.markState().underline ? "none" : "single" }));
  swatch(tableFont, "tblColor", "字体颜色", "#c00000", (v) => style({ color: v }));
  chooser(tableFont, "tblHl", "突出显示", DOC_HIGHLIGHTS, (v) => {
    editor?.applyStyle({ highlight: v });
    syncDocRibbon();
  });

  // `RB_Alignment` 那一组的六枚按 WPS 自己的两行摆：先竖向三枚（落格），再横向三枚（落格里的段）。
  // 组的位置照 `<subcommand id="CT_TableTools">` 的组序——`RB_Adjust` 之后、`TableSelectMenu` 之前；
  // 它前面那一组 `RB_Table_Font`（字体）已经接上，就在上面。
  const aligns = group("table");
  for (const [key, label, tip, value, leg] of TABLE_ALIGN_BUTTONS) {
    button(aligns, key, label, tip, () =>
      tableAction(leg === "v"
        ? editor?.setCellVAlign(value as "top" | "center" | "bottom")
        : editor?.setCellAlign(value as "left" | "center" | "right")));
  }

  // RB_Data 那一组（`CT_TableTools.kuip:78-91`）家里做得了的两条：「标题行重复」`RepeatHeaderRows`（:88）与
  // 「转换成文本」`ConvertTabletoText`（:89，排在重复标题行之后、排序之前）。公式那一枚（`FieldFormula`）仍未接，
  // 不摆空按钮。这一格是选方向的，不是开关：转换本身没有回显可言（表一没就没了），所以格子每次都回到「转换成文本」。
  const data = group("table");
  button(data, "repeatHeader", "标题行重复", "让表格开头连续的几行在每一页重复（WPS：Repeat Header Rows）", () => tableAction(editor?.toggleHeaderRows()));
  chooser(data, "convertTable", "转换成文本（WPS：Convert Table to Text）", [["", "转换成文本"], ["tab", "制表符分隔"], ["comma", "逗号分隔"], ["paragraph", "段落标记分隔"]], (v) => {
    (ctl.convertTable as HTMLSelectElement).value = "";
    if (v) tableAction(editor?.convertTableToText(v as "tab" | "comma" | "paragraph"));
  });
  // RB_Data 的最后一条：`CT_TableTools.kuip:90` 的 `Sort`（命令本体在 `wpscommon.kuip`，`ksoCmd="Sort"`）。
  // WPS 这一枚开的是对话框，家里照样开一张（`doc-table-sort.ts`）；形状不成立时由命令把理由带回来。
  button(data, "sortRows", "排序", "按某一列重排表体行：标题行钉住，整行一起走（WPS：Sort）", () => {
    const state = editor?.tableSortState();
    if (!state) {
      notify("光标不在表格里");
      return;
    }
    showTableSortDialog(state, (choice) => tableAction(editor?.sortTableRows(choice)));
  });

  const picking = group("table");
  button(picking, "selectCell", "选择单元格", "选择单元格（WPS：Select Cell）", () => tableAction(editor?.selectCell()));
  button(picking, "selectCol", "选择列", "选择列（WPS：Select Column）", () => tableAction(editor?.selectColumn()));
  button(picking, "selectRow", "选择行", "选择行（WPS：Select Row）", () => tableAction(editor?.selectRow()));
  button(picking, "selectTable", "选择表格", "选择表格（WPS：Select Table）", () => tableAction(editor?.selectTable()));

  // 插入页按 WPS 自己的**组归属**与**显示序**拆。显示序实据是 `CT_Insert.kuip:256-270` 那一段 `<subcommand id="CT_Insert">`：
  // `RB_Pages` →(sep)→ `RB_Illustrations` →(sep)→ `RB_Symbol` →(sep)→ `RB_InsertComments` →(sep)→ `RB_Links` → `RB_Text` →
  // `RB_ComplexObj` → `RB_Media`。家里只有头两档与批注那一档有命令，所以摆三组、不摆空组。
  // 页眉页脚的归属是这一片纠正出来的：`CT_Insert.kuip:139-150` 里 `RB_HeaderAndFooter` 是 **`embedHint="1"` 内嵌在 `RB_Pages` 里**
  // 的一枚子清单（孩子 `PageNumbers`→`m_InsertPageNumbers`、`HeaderAndFooter`→`HeaderAndFooterInInsertTab`），
  // 而 `CT_PageLayout.kuip:86-91` 那一页也有自己的一组 `RB_PageLayoutHeaderAndFooter`（`m_HeaderAndFooter`）——
  // **两处都摆**是 WPS 的形状，不是我们重复。断行族同挂 `RB_Pages`（`InsertBreaks`，在 `RB_HeaderAndFooter` 之前）。
  // 页码那一枚家里先不给入口：`insertPageNumber` 是私有方法，只由带子上那枚 chip 触发（`doc-edit.ts:776`），
  // ribbon 上按下去说不出该往哪一条带子插。
  const pages = group("insert");
  // WPS 把三条断命令挂在同一个下拉里（`wpsentries.kuip`：`CT_Insert.RB_Pages.InsertBreaks.*` 与
  // `CT_PageLayout.RB_Page_Setup.m_Breaks.*`、`InsertMenu.InsertBreakMenu.*`）。按钮文案只有分页符取到过证据
  // （`InsertPageBreak`），换行符那条 kuip 里写的是 `text="@Text Wrapping Break"` 这个待译占位，中文措辞在 `.qm`
  // 里读不出，所以「换行符／分栏符」两个词是按 WPS 家族习惯的**推断措辞**，不是取到的字面量。
  button(pages, "pagebreak", "分页符", "从这里另起一页", () => editor?.insertPageBreak());
  button(pages, "linebreak", "换行符", "在这里软换行，不另起段落", () => editor?.insertLineBreak());
  button(pages, "columnbreak", "分栏符", "从这里另起一栏", () => editor?.insertColumnBreak());
  button(pages, "header", "页眉", "编辑页眉（无则新建）", () => editor?.openHF("header"));
  button(pages, "footer", "页脚", "编辑页脚（无则新建）", () => editor?.openHF("footer"));

  // 表格在 WPS 那边不归「表」而归插图那一组：`CT_Insert.kuip:157` 的 `RB_Illustrations`（`groupTip="CommonObjectGroup"`）
  // 孩子第一个是 `Table`，接着 `m_rInsertPicture`。
  const illus = group("insert");
  button(illus, "table", "表格", "插入表格", openTablePicker);
  button(illus, "picture", "图片", "插入本地图片", () => void pickPicture((url, name) => void editor?.insertPicture(url, name)));

  // `CT_Insert.kuip:111-114` 的 `RB_InsertComments`（`groupTip="InsertCommentGroup"`）里是 `m_InsertCommentInToolBar`：
  // 「新建批注」在插入那一页还有一个入口，与审阅那一页那条同一个动作。键表是单值的（一枚 `ctl` 键只能挂一个元素，
  // 回显才不会挂错），所以这一枚不占 `cmtNew` 那格——它是命令的入口，不是状态。
  const insertCmt = group("insert");
  button(insertCmt, "", "新建批注", "给选中的文字加一条批注，接着在面板里写正文", () => {
    editor?.newCommentFromPanel();
    syncDocRibbon();
  });

  // 「更新目录」（#248）：WPS 把这条命令摆在 TOC 的右键菜单与「引用」页的 `UpdateTable` 里（入口位置
  // **未取证**，`.kuip` 只给 idMso）。自家没有「引用」页，摆进插入页与目录同一族；件里没有目录时灰着
  // （`hasToc`）——按下去什么都不做的按钮比没有更糊涂。
  button(pages, "tocUpdate", "更新目录", "按当前分页重算目录里的页码缓存（只动缓存、不动域名与书签）", () => {
    editor?.updateToc();
    syncDocRibbon();
  });

  // 审阅那一簇按 WPS 自己声明的组序排（`office6/res/commands/wps/ongmani/CT_Review.kuip:199-207`：
  // `RB_Proofing` → `RB_Comments` → `RB_Review` → `RB_Changes`，组与组之间是 `<separator/>`）。这一条是实据，
  // 不是印象：WPS 的批注簇在修订簇**之前**，接受/拒绝单独成一组，字数统计挂在校对那一组里（`WordCountToolBar`，
  // 同一支文件 :56）。组的中文名仍是**未取证**——`.kuip` 只给英文 source（`@Proofing` / `@Comments` /
  // `@Track Changes` / `@Change`），`mui/zh_CN/*.qm` 是压缩字典，读不出字面。所以标签用自家的话，顺序按它的。
  const proofing = group("review");
  button(proofing, "wordCount", "字数统计", "字数统计：报出段落数、字数与字符数", () => editor?.showWordCount());

  // 批注: the panel used to be reachable only by clicking a mark on the paper — so a document whose comments were
  // scrolled off screen, or a reader who wanted to annotate the selection without hunting for a mark first, had
  // no way in (`showComments` / `newCommentFromPanel` / `stepComment` had zero callers in `src`).
  const comments = group("review");
  button(comments, "cmtNew", "新建批注", "给选中的文字加一条批注，接着在面板里写正文", () => {
    editor?.newCommentFromPanel();
    syncDocRibbon();
  });
  button(comments, "cmtPrev", "上一条批注", "跳到上一条批注", () => {
    editor?.stepComment(-1);
    syncDocRibbon();
  });
  button(comments, "cmtNext", "下一条批注", "跳到下一条批注", () => {
    editor?.stepComment(1);
    syncDocRibbon();
  });
  button(comments, "cmtShow", "显示批注", "显示或收起批注面板", () => {
    editor?.showComments();
    syncDocRibbon();
  });

  // 修订: what a reader needs to work through somebody else's changes — the toggle, the list of pending ones,
  // stepping through them, and which version of the document the paper shows.
  const review = group("review");
  button(review, "track", "修订", "修订：开启后编辑记为待判修订，而不是直接落定", () => {
    if (editor) editor.setTrackChanges(!editor.trackingChanges);
    syncDocRibbon();
  });
  button(review, "revPane", "审阅窗格", "列出全部待判修订：逐条判定、按作者判定、定位到那一处", () => {
    editor?.showReviewPane();
    syncDocRibbon();
  });
  button(review, "revPrev", "‹", "上一条修订", () => {
    editor?.stepRevision(-1);
    syncDocRibbon();
  });
  button(review, "revNext", "›", "下一条修订", () => {
    editor?.stepRevision(1);
    syncDocRibbon();
  });
  chooser(
    review,
    "revView",
    "显示对文档的修改",
    [
      ["marks", "显示修订"],
      ["final", "最终状态"],
      ["original", "原文"],
    ],
    (v) => {
      editor?.setRevisionView(v as RevisionView);
      syncDocRibbon();
    },
  );

  // 更改: WPS keeps deciding the changes in their own group, right after the ones that track and show them.
  const changes = group("review");
  button(changes, "revAccept", "接受全部", "接受修订：插入的留下，删除的生效", () => {
    editor?.acceptRevisions();
    syncDocRibbon();
  });
  button(changes, "revReject", "拒绝全部", "拒绝修订：删除的文字回来，插入的文字去掉", () => {
    editor?.rejectRevisions();
    syncDocRibbon();
  });

  // 页面布局页只管纸本身：`CT_PageLayout.kuip:196-203` 那条 `<subcommand id="CT_PageLayout">` 的显示序是
  // `RB_Page_Setup` →(sep)→ `RB_EditTheme` →(sep)→ `RB_Genko_Setting` →(sep)→ `RB_PageLayoutHeaderAndFooter`
  // →（后面那两行是门控的「展开」按钮，`idKey="LogicGroupShowExtensionButton"`）。页面设置归第一组。
  // 第二、三组（效果 / 结构：页眉底纹、页面边框、水印、行号、分栏、删除分节符…）家里还没有命令面，不摆空组。
  const layout = group("layout");
  button(layout, "pageSetup", "页面设置", "纸张大小与方向、页边距、页眉页脚距离、装订线", () => editor?.openPageSetup());
  // `RB_Page_Setup` 里不止这一枚按钮：`CT_PageLayout.kuip:153-166` 把四枚页边距自旋格（上→左→下→右，每枚
  // `fixedSpinBoxWidth="84"`）、`PageOrientation`、`m_PageSize` 平铺在同一组里。三样家里都有面——都走 `Doc.setPageSetup`
  // 那一条，与「页面设置」对话框同一条路，不另开一套数。同组那两枚**不摆**：分栏（授权 46 份的 55 枚 `w:cols` 全是
  // `w:num="1"`，0 受众）与文字方向（页级那一维家里没有面，`w:textDirection` 只建模到表格单元格；WPS 的页级拼法未取证）。
  for (const [key, label, edge, slot] of PAGE_MARGIN_SPINS) {
    spin(layout, key, label, `${label}页边距，按状态栏的测量单位填；落笔（失焦或回车）才算`, (text) => {
      const s = editor?.pageSection();
      if (!s) return;
      const v = Number(text.trim());
      // 说不出数（空的、半个、负）就不改件，也**当场**把这一格刷回件里那句：押着一个没人认的「abc」在框里，
      // 用户看不出那是被拒了还是已经生效；而「打了还没落笔的那一格别动」那一条护的是没落笔的字，落笔这一趟它已经落笔了。
      if (text.trim() === "" || !Number.isFinite(v) || v < 0) {
        echoSpin(ctl[key] as HTMLInputElement, s.margins[edge]);
        return;
      }
      const patch: Partial<PageSetup> = {};
      const twips = twipsFromUnit(v, currentUnit());
      patch[slot] = twips;
      editor?.applyPageSetup(pageSetupOf(s, patch));
      // 落笔这一格自己读回刚写进件里的那个数（同口径，`twipText`/`twipsFromUnit` 在这一档能精确往返）。
      // 这一步不刷，下一趟换测量单位时会把「4」当成用户没落笔的字护住，四格里就剩它一个不换口径。
      echoSpin(ctl[key] as HTMLInputElement, twips);
      syncDocRibbon();
    });
  }
  chooser(layout, "orient", "纸张方向：这一张纸是纵向还是横向", PAGE_ORIENTATIONS, (v) => {
    const s = editor?.pageSection();
    if (!s) return;
    const want = v === "landscape";
    if ((s.pageWidthTwips > s.pageHeightTwips) === want) return;
    // 只给方向、不给 `paperTwips`：模型的这条意思是「纸还是那张，只是转方向」，不由界面替文件重算两条边。
    editor?.applyPageSetup(pageSetupOf(s, { landscape: want }));
    syncDocRibbon();
  });
  chooser(layout, "paper", "纸张大小：选一张标准纸，按当前方向摊开", PAPER.map((p): [string, string] => [p.name, p.name]), (name) => {
    const s = editor?.pageSection();
    // 名字读不出纸（那一格显示「自定义尺寸」时）就不改件：这一格说的是「换纸」，不是「把我叫不出名的那张再写一遍」。
    const edges = s ? paperEdges(name, s.pageWidthTwips > s.pageHeightTwips) : undefined;
    if (!s || !edges) return;
    editor?.applyPageSetup(pageSetupOf(s, { paperTwips: edges }));
    syncDocRibbon();
  });

  // 最后一组 `RB_PageLayoutHeaderAndFooter`（:86-91，`groupTip="HeaderFooterGroup"`）里是
  // `PageLayout_InsertPageNumbers` + `m_HeaderAndFooter`。页眉页脚在插入那一页也有一份（上面 `RB_Pages` 里那两枚），
  // WPS 两处都摆——所以这里不是重复，是那两页各自的入口。这一组的按钮不占 `ctl` 键：`header` / `footer` 那格已给插入页那两枚。
  const layoutHF = group("layout");
  button(layoutHF, "", "页眉", "编辑页眉（无则新建）", () => editor?.openHF("header"));
  button(layoutHF, "", "页脚", "编辑页脚（无则新建）", () => editor?.openHF("footer"));

  // 视图页改的是纸周围的工作台，不是纸上的字：`CT_View.kuip` 声明的组序 `RB_Document_View` → `RB_Show` → `RB_Zoom`
  // 里，`sp_ShowDocumentMap` 挂在 `RB_Show`、`ZoomToolBar` 挂在 `RB_Zoom`——所以导航窗格与缩放同在一页。
  const view = group("view");
  button(view, "navPane", "导航窗格", "按大纲级别列出各级标题，点一行跳到那一处", () => {
    editor?.showOutlinePane();
    syncDocRibbon();
  });
  // 标尺紧跟导航窗格：`CT_View.kuip` 的 `RB_Show`（`:30`，`idMso="GroupViewShowHide"`）里 `ViewRulerWord`
  // （`:34`，`ksoCmd="Ruler"`、`customTip="Ruler"`、`hotKey="l"`、`text="@Ruler"`）就排在
  // `sp_ShowDocumentMap` 之后、`CommentOrTrackingChangeMarkup` 之前，缩放是下一组 `RB_Zoom` 的事。
  // 它说的是工作台上画不画标尺：件里没有这一句，所以不进撤销栈，也不把件弄脏。
  button(view, "ruler", "标尺", "显示或隐藏标尺（WPS 视图·显示组 ViewRulerWord）", () => {
    showRuler = !showRuler;
    syncDocRibbon();
  });
  // 缩放另起一组，不塞进导航窗格那一组里：WPS 那一页把 `ZoomToolBar` 单挂在 `RB_Zoom`。
  shell.bars["view"]!.append(zoomGroup());

  // 各档的组都进来了，这一趟量准箭头的隐现（之后 resize / 滚动会自己刷新）。
  shell.sync();
  // 测量单位换了，`RB_Page_Setup` 那四格的字面口径就跟着换：回显只有一条路，就是再跑一遍 `syncDocRibbon`。
  // 标尺上那排数字也是一句「按哪个单位读」，所以同一趟把刻度重画一遍（`refreshRuler` 只改视图，不碰件）。
  // 这一枚订阅归本趟重建持有，先收上一趟的，不然每开一个文档就多留一个监听。
  unitEchoOff?.();
  unitEchoOff = onUnitChange(() => {
    syncDocRibbon();
    editor?.refreshRuler();
  });
}

/** `RB_Page_Setup` 那几格的「单位换了就重算」监听：同一个持有法则是溢出箭头的 `resize` 监听（`overflowAbort`）。 */
let unitEchoOff: (() => void) | undefined;

/** 溢出箭头的 `resize` 监听归这一枚元素自己：重建前先 `abort()`，不然每切一次文档种类就多留一个监听。 */
let overflowAbort: AbortController | undefined;

/** 表格工具那一档（上下文页）的开关：露不露、此刻停在哪一页、怎么退，都是 `buildDocRibbon` 建出来的。 */
let tableTab: { button: HTMLButtonElement; shown: () => string; show: (key: string) => void } | undefined;

/**
 * 「查看网格线」那一档（`CT_TableTools.kuip:8` 的 `ShowGridlines`，`idMso="TableShowGridlines"`）说的是**屏幕上
 * 画不画导引线**，不是文件里的一句话，所以它既不进 `DocEditor`、也不进撤销栈、更不把件弄脏：
 * OOXML 里确实没有这一句——授权 46 份的 `word/settings.xml` 里带 Grid 的四族（`drawingGridVerticalSpacing` 44 枚、
 * `displayHorizontalDrawingGridEvery` 46 枚、`displayVerticalDrawingGridEvery` 44 枚、`drawingGridHorizontalSpacing` 6 枚）
 * 说的是**画布**那张网格，与表的格子边无关（同一趟量的，`.scratch/gridlines-audience.py`）。
 * 落点只有容器上那一枚 `doc-grid`：`render.ts` 把「文件没说过要画线」的那几条边并成一串 `--gl` 内阴影，
 * CSS 按这一枚挂不挂决定画不画；关掉之后那些边一条都不画——那才是文件说的样子。
 * WPS 那一档默认开还是关**未取证**（装机件里没量过这一档）；家里默认开，理由是关掉之后一张没说过线的表
 * 在屏幕上连地界都看不见，用户无从下手。
 */
let showGridlines = true;

/**
 * 标尺画不画（`CT_View.kuip` 的 `ViewRulerWord`，`ksoCmd="Ruler"`）。与「查看网格线」同一条性质：
 * 它挂在工作台上，件里没有这一句，`word/settings.xml` 里也没有——所以既不改字节也不进撤销栈。
 * WPS 默认开还是关未取证（真机卡在首次启动那一页）；家里默认开，理由与标尺本身一样：看不到正文栏在哪，
 * 段落缩进就只能猜。标尺自己怎么画在 `doc-ruler.ts`。
 */
let showRuler = true;

/**
 * 编辑标记画不画（WPS `m_ShowHideEditMark` 的第一条子项「段落标记」）。与「查看网格线」「标尺」同一条性质：
 * 它挂在工作台上，文件里没有「显示段落标记」这一位——所以不改字节、不进撤销栈、也不报脏。
 * 墨本身在 `index.html` 的 `#view.edit-marks .page .para::after`（零宽，不动铺排）；这里只管挂不挂那一枚类。
 * WPS 默认显示还是隐藏**未取证**；家里默认隐藏，与 WPS 新开一档的观感一致（排版标记不该默认糊在纸上）。
 */
let showEditMarks = false;

/**
 * ribbon 跟着文档读的那一趟。`force` 只在**换文档**那一趟给：上一份里打了还没落笔的那串字属于上一份文件，
 * 不搬进这一份（`buildRibbon` 同种类不重建，那几格是同一批元素，所以这一条得由这里说）。
 */
function syncDocRibbon(force = false): void {
  const state = editor?.markState();
  const mark = (key: string, want: boolean): void => {
    (ctl[key] as HTMLButtonElement | undefined)?.classList.toggle("on", want);
  };
  // 同一句话在两页各有一枚按钮时，回显一次点亮两枚：开始页与表格工具页那两枚 `Bold`/`Italic`/`Underline`
  // 在 WPS 是同一个 `ksoCmd` 挂在两处（`CT_Home.kuip:125/127/130` 与 `CT_TableTools.kuip:138/141/144`），
  // 对着同一处光标各说一句就是撒谎。
  const lit = (keys: string[], on: boolean): void => {
    for (const key of keys) mark(key, on);
  };
  lit(["bold", "tblBold"], state?.bold ?? false);
  lit(["italic", "tblItalic"], state?.italic ?? false);
  lit(["under", "tblUnder"], state?.underline ?? false);
  mark("strike", state?.strike ?? false);
  // 上下标两枚一起亮是不可能的：`markState()` 只交一个字段；明说 `baseline` 的那枚 run（授权 9 份作者件里 0 枚、
  // 模型另有那一档）两枚都不亮——纸上它没抬也没落，亮哪一枚都是替它说话。
  mark("sup", state?.vertAlign === "superscript");
  mark("sub", state?.vertAlign === "subscript");
  // 「查看网格线」那一枚回显的是**视图档**、不问文档：它不跟着光标变，也不参与「件改了没有」的判断。
  // 容器上那一枚 `doc-grid` 也在这一趟挂/摘——换文件、开关、编辑后重读走的都是这一条路，不留第二处写法。
  mark("gridlines", showGridlines);
  view.classList.toggle("doc-grid", showGridlines);
  // 标尺那一枚与「查看网格线」同一条：回显的是**视图档**，不问文档。画与不画落在编辑器那边
  // （`DocEditor.setRulerShown` 同值不动，所以每趟 `syncDocRibbon` 都可以走这一句）。
  mark("ruler", showRuler);
  editor?.setRulerShown(showRuler);
  // 编辑标记与上面两条同一条：回显问视图档、不问文档；类挂在同一个容器上，换文件/重排都走这一条路。
  mark("editmarks", showEditMarks);
  view.classList.toggle("edit-marks", showEditMarks);
  // 格式刷亮着＝还举着一趟没用完（按下后到刷上之前，按钮亮着，纸上的指针也换成 `copy`）。
  mark("painter", editor?.painterArmed() ?? false);
  // 「更新目录」只在件里真有目录（只带容器里有 PAGEREF）时可用：按下去什么都不做的按钮比没有更糊涂。
  const tocBtn = ctl.tocUpdate as HTMLButtonElement | undefined;
  if (tocBtn) tocBtn.disabled = !(editor?.hasToc() ?? false);
  // 「没说」与「明说 left」画的是一件事（§17.3.1.13 默认左；授权 46 份里 88 段明说 left，
  // `.scratch/authorized-jc-echo.ts`），两档都该亮「左」——只亮 `undefined` 那一档，停在那 88 段上时
  // 纸面靠左、按钮上一枚都不亮，说的是谎话。
  mark("left", state?.align === undefined || state?.align === "left");
  mark("center", state?.align === "center");
  mark("right", state?.align === "right");
  mark("both", state?.align === "both");
  // 对齐那两行回显的口径与命令的落点同一份：竖向问**格**（`w:vAlign`），横向问格里的**段**（`w:pPr/w:jc`），
  // 两处都问选区盖住的那几格（`cellAlignState`）。光标在表外时开始页那三枚灰着、两页六枚都不亮；在表里时
  // 格没写就是顶部（§17.4.84 的默认值）、段没写就是左（§17.3.1.13），说了哪一档就亮哪一枚。
  // 选到顶/中/底各一格时那一行一枚都不亮——亮着「顶端」再按下去会把三格改成同一档，那一枚亮的是谎话。
  // 开始页那三枚与表格工具页那六枚读的是同一个 `cellAlignState()`，对着同一份文档不各说一句。
  const cellState = editor?.cellAlignState();
  const inTable = cellState?.inTable ?? false;
  const vOn = (want: string): boolean => inTable && !cellState?.valignMixed && cellState?.valign === want;
  const hOn = (want: string): boolean => inTable && !cellState?.alignMixed && cellState?.align === want;
  lit(["vtop", "tblTop"], vOn("top"));
  lit(["vmid", "tblMid"], vOn("center"));
  lit(["vbot", "tblBot"], vOn("bottom"));
  lit(["tblLeft"], hOn("left"));
  lit(["tblCenter"], hOn("center"));
  lit(["tblRight"], hOn("right"));
  // 「标题行重复」的回显问的是**光标那一行在不在表头那一段里**——正是那一击会关掉的状态（命令按「选区末行」
  // 定标题块尽头）。光标不在表里时 `rowHeaderState()` 给 `undefined`，按钮灰着（`TABLE_TOOL_KEYS` 一并管）。
  mark("repeatHeader", editor?.rowHeaderState() === true);
  for (const key of ["vtop", "vmid", "vbot"]) {
    const b = ctl[key] as HTMLButtonElement | undefined;
    if (b) b.disabled = !inTable;
  }
  // 表格工具这一页只在光标进表时露出来。光标离开时正停在这一页，就退回开始页——留着一页没有页签的带子，
  // 用户看不出自己在哪，也回不去（页签已经退场）。
  if (tableTab) {
    tableTab.button.hidden = !inTable;
    if (!inTable && tableTab.shown() === "table") tableTab.show("home");
  }
  for (const key of TABLE_TOOL_KEYS) {
    const b = ctl[key] as HTMLButtonElement | HTMLSelectElement | HTMLInputElement | undefined;
    if (b) b.disabled = !inTable;
  }
  mark("bullet", state?.list === "bullet");
  mark("number", state?.list === "number");
  // 「样式」那一格既是清单又是回显。清单按**件里的段落样式**刷，签名相同就不重建（同一份文件里每趟选区变化
  // 都 replaceChildren 一次，会把用户拉开的下拉合上）；换文档时 `force` 一定重建。回显问 `markState().styleId`：
  // `undefined` 是「不点名」，落在默认样式那一格（空值）——与纸面同一个 `shown`（原文视图读记录里的 prev）。
  const styleSel = ctl.style as HTMLSelectElement | undefined;
  if (styleSel) {
    const list = editor?.paragraphStyles() ?? [];
    const ordered = [...list.filter((s) => s.isDefault), ...list.filter((s) => !s.isDefault)];
    const sig = JSON.stringify(ordered);
    if (force || styleSel.dataset.sig !== sig) {
      styleSel.dataset.sig = sig;
      const make = (value: string, label: string): HTMLOptionElement => {
        const o = document.createElement("option");
        o.value = value;
        o.textContent = label;
        return o;
      };
      styleSel.replaceChildren(
        ...(ordered.length ? ordered.map((s) => make(s.isDefault ? "" : s.id, s.name)) : [make("", "正文")]),
      );
    }
    echoChoice(styleSel, state?.styleId ?? "", ["", ...ordered.filter((s) => !s.isDefault).map((s) => s.id)], 1);
  }
  // 边框那一格的列子问的是「这一击落在哪一层」：光标在正文、在表上铺满整张、还是只圈了几格。所以它随选区
  // 重建，而不是随选区**回显**——一层的框是逐边的，说不出「现在挂着哪一档」。重建只在换层时发生，同一层里
  // 用户选到一半的那格不动。
  const border = ctl.paraBorder as HTMLSelectElement | undefined;
  if (border) {
    const scope = editor?.borderScope() ?? "para";
    if (border.dataset.scope !== scope) {
      border.dataset.scope = scope;
      border.replaceChildren(
        ...BORDER_LEGS[scope].map(([value, label]) => {
          const o = document.createElement("option");
          o.value = value;
          o.textContent = label;
          return o;
        }),
      );
      border.value = "";
      border.title = scope === "table" ? "表格边框" : scope === "cells" ? "单元格边框" : "边框";
    }
  }
  // 字体、字号、底色这三族在表里与表外各有一枚控件（`CT_Home.kuip:110-111` 与 `CT_TableTools.kuip:133-134/150`），
  // 两枚读的是同一份 `markState()`：光标停在表格里那一段时，开始页那一格与表格工具这一格得报同一句话。
  // 正在打的那一格不回显（只有换文档那一趟 `force` 除外）：
  // `syncDocRibbon` 随每次选区变化跑一遍，把用户打到一半的名字吃掉就不是他的名字了。
  // 回显给的是选区起点那一段（`markState` 一路按起点算），两张脸并存时给这段字真正穿的那一张。
  for (const key of ["font", "tblFace"]) {
    const box = ctl[key] as HTMLInputElement | undefined;
    if (box && (force || document.activeElement !== box)) box.value = state?.font ?? "";
  }
  // 字号框是可打字的 `input`，不是 `<select>`：说不出就打空（「没提过」），说得出来就摆文件那个数——
  // 号数制里的磅数用中文名（五号/小四/小初……，WPS 同一显示法：`event plan.dotm` 主标题 36 磅显「小初」、
  // 目录题 18 磅显「小二」），不在号数制里的照旧显数字；不摆最近的一档预设（那会给这段字安一个它没写过的
  // 尺寸）。`undefined` 与「空框」同一句话。正在打字的那一格不刷（与字体框同一条礼节）：刷了就吃掉用户
  // 打到一半的数。打名字也认（`parseSizePt` 按同一张表换回磅数）。
  for (const key of ["size", "tblSize"]) {
    const box = ctl[key] as HTMLInputElement | undefined;
    if (box && (force || document.activeElement !== box)) box.value = state?.sizePt === undefined ? "" : sizeNameOf(state.sizePt) ?? String(state.sizePt);
  }
  // Three sentences the file could have written — 没提过底色 / 提了「没有」 / 提了某个颜色名 — three readings of
  // the box. Leaving it blank is the first, and only the `none` option spells the second.
  for (const key of ["highlight", "tblHl"]) {
    const hl = ctl[key] as HTMLSelectElement | undefined;
    if (hl) echoChoice(hl, state?.highlight, DOC_HIGHLIGHTS.map(([v]) => v), DOC_HIGHLIGHTS.length);
  }
  // 「字体颜色」在两页各有一格（`CT_Home.kuip:108` 的 `FontColor` 与 `CT_TableTools.kuip:147` 那枚同 `ksoCmd`
  // 的 `FontColor`），读的是同一份 `markState().color`。它和字号/底色那两族差一步口径：`input[type=color]`
  // 说不出「没提过」——它只能摆一个颜色——所以这一格报的是「光标这段字现在穿的墨」：从红的走到没说的，格子跟着
  // 走回纸上那层墨，不留在上一句。留在上一句就是替这段字说一件它自己没说过的颜色。
  for (const key of ["color", "tblColor"]) {
    const box = ctl[key] as HTMLInputElement | undefined;
    if (box && state?.color && (force || document.activeElement !== box)) box.value = state.color;
  }
  // 行距那一格只认 `w:lineRule="auto"` 那一族：`w:line` 在 固定值/最小值 下数的是缇，把它当成倍数摆进这一格
  // 就是给 ribbon 编一个件里没说过的数（#132 之前正是这样把 固定值 22 磅画成 1.83 倍的）。所以这两档与「没提过」
  // 一律留空——那句话说的是另一件事，要看得去「段落」对话框。件里 7,810 枚 `w:line` 有 2,178 枚是 `auto`。
  const ls = ctl.linespacing as HTMLSelectElement | undefined;
  if (ls) {
    const mult = state?.lineRule === "auto" && state.line !== undefined ? state.line / LINE_ONE : undefined;
    echoChoice(ls, mult === undefined ? undefined : String(mult), DOC_LINES.map(String), DOC_LINES.filter((n) => n < (mult ?? 0)).length);
  }
  // `RB_Page_Setup` 那几格回显的是**这一节现在**的几何：没落笔的那一格不动（护的是打到一半的字，见下），
  // 单位换字面靠 `onUnitChange` 再跑一趟这里；两条都问 `editor.pageSection()`，与对话框同一个出口。
  const page = editor?.pageSection();
  if (page) {
    for (const [key, , edge] of PAGE_MARGIN_SPINS) {
      const box = ctl[key] as HTMLInputElement | undefined;
      // 护的是「打了还没落笔」那一格，不是「光标正站着」那一格：站着的格若框里还是上一趟刷给它的那句，
      // 就没有用户的字要护——换测量单位时四格都得跟着换口径，只跳过站着的那一格会留下一个说旧单位字面的格。
      if (box && (force || box.value === box.dataset.echo)) echoSpin(box, page.margins[edge]);
    }
    const orient = ctl.orient as HTMLSelectElement | undefined;
    if (orient) orient.value = page.pageWidthTwips > page.pageHeightTwips ? "landscape" : "portrait";
    const paper = ctl.paper as HTMLSelectElement | undefined;
    // 件里那一对边叫不出标准名时，那一格显示「自定义尺寸」这件纸自己说的样子，而不是列表里最近的一张。
    if (paper) echoChoice(paper, paperOf(page.pageWidthTwips, page.pageHeightTwips) ?? CUSTOM, PAPER.map((p) => p.name), PAPER.length);
  }
  // 「单元格尺寸」那两格与页边距四格同一条口径：件（这里是光标那格所在的行/列）没说的不替它编一个数，
  // 框里空着；光标出表时 `cellSizeState()` 说不出尺寸，两格一起清空——留着上一张表的数，配合那一页页签已经退场，
  // 摆的是「这张纸上一个数也没有，但 ribbon 说它 2 厘米」那种假话。
  const cellSize = editor?.cellSizeState();
  for (const [key, , dim] of CELL_SIZE_SPINS) {
    const box = ctl[key] as HTMLInputElement | undefined;
    if (box && (force || box.value === box.dataset.echo)) echoSpin(box, cellSize?.[dim]);
  }
  mark("track", editor?.trackingChanges ?? false);
  mark("revPane", editor?.reviewPaneOpen ?? false);
  mark("navPane", editor?.outlinePaneOpen ?? false);
  mark("cmtShow", editor?.commentsPaneOpen ?? false);
  const revView = ctl.revView as HTMLSelectElement | undefined;
  if (revView && editor) revView.value = editor.revisionShown;
  const undoBtn = ctl.undo as HTMLButtonElement | undefined;
  const redoBtn = ctl.redo as HTMLButtonElement | undefined;
  if (undoBtn) undoBtn.disabled = !editor?.canUndo;
  if (redoBtn) redoBtn.disabled = !editor?.canRedo;
}

/**
 * ribbon 上那几格数字唯一的写手：把缇数按当前测量单位写成字面放进格子，同时把这个字面记进 `data-echo`。
 * 回显拿 `data-echo` 问「框里还是不是上一趟刷的那句」，落笔与拒收都走这里写自己那一格——两处共用，
 * 才能保证「护没落笔的那格」和「换单位几格一起换口径」不互相打架。
 * `undefined` 是「件里没说这个数」，摆空框：页边距那四格永远有数（节总有 `w:pgMar`），
 * 「单元格尺寸」那两格没有——行没写 `w:trHeight`、列说不出单一宽度时，替它编一个数就是往件里写字。
 */
function echoSpin(box: HTMLInputElement, twips: number | undefined): void {
  const want = twips === undefined ? "" : twipText(twips);
  box.value = want;
  box.dataset.echo = want;
}

/** 字号框此刻该显示的那一句：光标所在那枚 run 说的磅数，号数制里的数用中文名（五号/小四……，WPS 同一显示法，
 *  `event plan.dotm` 两处实测对上），不在号数制里的照旧显数字；说不出就是空（「没提过」），不替它编一个数。 */
function sizeEcho(): string {
  const pt = editor?.markState().sizePt;
  return pt === undefined ? "" : sizeNameOf(pt) ?? String(pt);
}

/**
 * A chooser states the sentence the caret's run says, not the closest thing it happens to list.
 *
 * `undefined` deselects, which is how 「没提过」 reads; `value` among `presets` selects that option; a value the
 * control does not list gets one option of its own at `insertAt`, taken back out on the next sync. Showing the
 * nearest preset instead would put a property in the ribbon that is not in the file — of the authorized 46 份, 30
 * state a run size at all in the body or a band and all 30 state one outside the eight presets: 134 of their 5,662
 * `w:sz` are 9, 15, 22, 26 or 42 磅, none of which the list carries, so a 15pt caret used to be answered with a
 * preset the file never stated
 * (`.scratch/authorized-slot-census.py`; the 61-of-86 reading is the quarantined cohort's arithmetic).
 */
function echoChoice(s: HTMLSelectElement, value: string | undefined, presets: string[], insertAt: number): void {
  for (const o of [...s.options]) if (o.dataset.caret === "1") o.remove();
  if (value === undefined) {
    s.selectedIndex = -1;
    return;
  }
  if (!presets.includes(value)) {
    const o = document.createElement("option");
    o.value = value;
    o.textContent = value;
    o.dataset.caret = "1";
    s.insertBefore(o, s.options[insertAt] ?? null);
  }
  s.value = value;
}

/** Deck operations; text styles act on the selection being typed into. */
function buildDeckRibbon(): void {
  const shell = ribbonTabs(DECK_RIBBON_TABS);
  /** 组名是自家措辞：`.kuip` 只有 `text="@Slide"` 这类英文占位，中文在 `.qm` 里读不出（与表格页同一口径）。 */
  const group = (tab: string, cap: string): HTMLElement => {
    const g = shell.group(tab);
    g.dataset.cap = cap;
    return g;
  };
  const button = (g: HTMLElement, key: string, label: string, tip: string, run: () => void): void => {
    const b = document.createElement("button");
    b.className = "rb";
    b.textContent = label;
    b.title = tip;
    decorateRibbonButton(b, label);
    b.onclick = run;
    ctl[key] = b;
    g.append(b);
  };
  const style = (kind: "bold" | "italic" | "underline"): void => {
    deck?.applyStyle(kind);
    syncDeckRibbon();
  };
  const chooser = (g: HTMLElement, key: string, tip: string, options: readonly [string, string][], run: (value: string) => void): void => {
    const s = document.createElement("select");
    s.className = "pick";
    s.title = tip;
    for (const [value, label] of options) {
      const o = document.createElement("option");
      o.value = value;
      o.textContent = label;
      s.append(o);
    }
    s.onchange = () => run(s.value);
    ctl[key] = s;
    g.append(s);
  };
  const swatch = (g: HTMLElement, key: string, tip: string, value: string, run: (value: string) => void): void => {
    const i = document.createElement("input");
    i.type = "color";
    i.className = "swatch";
    i.title = tip;
    i.value = value;
    i.oninput = () => run(i.value);
    ctl[key] = i;
    g.append(i);
  };

  // 撤销/重做只在顶栏 QAT 那一处（`undo()` 已按 `current.kind` 派发到演示的历史），开始页不再摆第二份。
  const pages = group("home", "幻灯片");
  button(pages, "", "＋新页", "在当前页后新建一页", () => deck?.addSlide());
  button(pages, "", "删除页", "删除当前页", () => deck?.removeSlide());
  button(pages, "", "上移", "当前页前移一位", () => deck?.moveSlide(-1));
  button(pages, "", "下移", "当前页后移一位", () => deck?.moveSlide(1));

  const font = group("home", "字体");
  button(font, "bold", "粗", "加粗 ⌘B", () => style("bold"));
  button(font, "italic", "斜", "斜体 ⌘I", () => style("italic"));
  button(font, "under", "线", "下划线 ⌘U", () => style("underline"));
  chooser(font, "size", "字号", DECK_SIZES.map((s): [string, string] => [String(s), String(s)]), (v) => {
    deck?.applySize(Number(v));
    syncDeckRibbon();
  });
  swatch(font, "color", "字体颜色", "#c00000", (v) => {
    deck?.applyRunColor(v);
    syncDeckRibbon();
  });

  const align = group("home", "段落");
  button(align, "aleft", "左", "左对齐", () => deck?.applyAlign("left"));
  button(align, "acenter", "中", "居中", () => deck?.applyAlign("center"));
  button(align, "aright", "右", "右对齐", () => deck?.applyAlign("right"));
  button(align, "aboth", "两", "两端对齐", () => deck?.applyAlign("justify"));

  // 开始页 `RB_Drawing`（`CT_Home.kuip`：`ksoCmd="FormatObject"`，组里是 `DynamicShape` 的形状画廊与
  // `InsertTextboxHorizontal`）。形状画廊归这一组，插入页那一档就不再摆第二份。
  const draw = group("home", "绘图");
  for (const [label, geometry] of [
    ["矩形", "rect"],
    ["圆角", "roundRect"],
    ["椭圆", "ellipse"],
    ["三角", "triangle"],
    ["直线", "line"],
    ["菱形", "diamond"],
    ["平行四边形", "parallelogram"],
    ["右箭头", "rightArrow"],
    ["五角星", "star5"],
    ["燕尾形", "chevron"],
  ] as [string, PresetGeometry][]) {
    button(draw, "", label, `插入${label}`, () => deck?.insertPreset(geometry));
  }
  button(draw, "", "删除形状", "删除选中的形状、图片或表格", () => deck?.deleteShape());
  swatch(draw, "fill", "形状填充", "#dce6f1", (v) => {
    deck?.applyShapeFill(v);
    syncDeckRibbon();
  });
  swatch(draw, "line", "形状边框", "#4472c4", (v) => {
    deck?.applyShapeLine(v);
    syncDeckRibbon();
  });

  // 插入页：`RB_Text`（`InsertTextboxHorizontal`）与 `RB_Shape`（`text="@Graph And Image"`，组里第一枚就是
  // `InsertPictureFromFile`）。WPS 的 `RB_Media` 是联机音频/视频，这两条今天没有命令，不摆空组。
  button(group("insert", "文本"), "", "文本框", "在当前页插入文本框", () => deck?.insertTextBox());
  button(group("insert", "图形和图片"), "", "图片", "插入本地图片", pickDeckPicture);

  const show = group("show", "开始放映");
  button(show, "", "放映", "全屏放映，从头开始（F5）", () => startShow(false));
  button(show, "", "本页放映", "从当前页开始放映（⇧F5）", () => startShow(true));

  shell.bars["view"]!.append(zoomGroup());
  shell.sync();
}

// ------------------------------------------------------------- shared chrome

/** Document zoom; each editor applies it to its own surface, not to the chrome. */
let zoomLevel = 1;

function applyZoom(z: number): void {
  zoomLevel = Math.min(2.5, Math.max(0.4, Math.round(z * 100) / 100));
  if (current?.kind === "xlsx") grid?.setZoom(zoomLevel);
  else if (current?.kind === "docx") editor?.setZoom(zoomLevel);
  else deck?.setZoom(zoomLevel);
  const label = ctl.zoom as HTMLElement | undefined;
  if (label) label.textContent = `${Math.round(zoomLevel * 100)}%`;
}

function zoomGroup(): HTMLElement {
  const g = document.createElement("div");
  g.className = "grp";
  // 三页的视图档都用这一组，组名照 WPS 的 `RB_Zoom`（各家的 `CT_View.kuip`）。
  g.dataset.cap = "缩放";
  const mk = (label: string, tip: string, run: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.className = "rb";
    b.textContent = label;
    b.title = tip;
    b.onclick = run;
    g.append(b);
    return b;
  };
  mk("⊖", "缩小", () => applyZoom(zoomLevel - 0.1));
  const label = document.createElement("button");
  label.className = "rb";
  label.textContent = "100%";
  label.title = "恢复 100%";
  label.onclick = () => applyZoom(1);
  ctl.zoom = label;
  g.append(label);
  mk("⊕", "放大", () => applyZoom(zoomLevel + 0.1));
  return g;
}

/**
 * One image picker for all three editors: the shell picks a file in Electron;
 * a plain browser build falls back to a file input. Hands over a data URL plus
 * the decoded natural size.
 */
async function pickPicture(run: (url: string, name: string, width: number, height: number) => void | Promise<void>): Promise<void> {
  let name = "图片";
  let url = "";
  const bridge = window.soloOffice;
  if (bridge) {
    const got = await bridge.pickImage();
    if (!got) return;
    name = got.name;
    url = dataUrl(imageMimeOf(name), got.data);
  } else {
    const picked = await new Promise<File | undefined>((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "image/*";
      input.onchange = () => resolve(input.files?.[0]);
      input.click();
    });
    if (!picked) return;
    name = picked.name;
    url = await new Promise<string>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result ?? ""));
      reader.readAsDataURL(picked);
    });
  }
  const size = await naturalSize(url);
  await run(url, name.replace(/\.[^.]+$/, "") || "图片", size.width, size.height);
}

const imageMimeOf = (name: string): string =>
  IMAGE_MIME_BY_EXT[name.slice(name.lastIndexOf(".") + 1).toLowerCase()] ?? "image/png";

function naturalSize(url: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 320, height: img.naturalHeight || 240 });
    img.onerror = () => resolve({ width: 320, height: 240 });
    img.src = url;
  });
}

/** The highlight-rule popover: an operator, a number, and the selection. */
function openCondPanel(): void {
  document.querySelector(".cond-panel")?.remove();
  const at = (ctl.cfBtn as HTMLElement | undefined)?.getBoundingClientRect();
  const panel = document.createElement("div");
  panel.className = "menu cond-panel";
  const sel = document.createElement("select");
  sel.className = "pick";
  for (const [v, label] of [["greaterThan", "大于"], ["lessThan", "小于"], ["equal", "等于"], ["notEqual", "不等于"]] as [string, string][]) {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = label;
    sel.append(o);
  }
  const input = document.createElement("input");
  input.type = "text";
  input.placeholder = "数值";
  input.style.cssText = "width:76px;font:inherit;font-size:12px;padding:3px 6px;border:1px solid #c9ced6;border-radius:5px;background:#fff;color:var(--ink);";
  const go = document.createElement("button");
  go.className = "rb";
  go.textContent = "应用";
  go.onclick = () => {
    const value = input.value.trim();
    if (value === "") return;
    panel.remove();
    grid?.setCellIsUI(sel.value, value);
  };
  input.addEventListener("keydown", (e) => {
    if ((e as KeyboardEvent).key === "Enter") go.click();
  });
  panel.append(sel, input, go);
  panel.style.left = `${Math.max(8, at ? at.left : 120)}px`;
  panel.style.top = `${(at?.bottom ?? 96) + 6}px`;
  document.body.append(panel);
  input.focus();
  window.addEventListener("pointerdown", (e) => {
    if (!panel.contains(e.target as Node)) panel.remove();
  }, { once: true });
}

/** The WPS-style hover grid: click a cell to insert that many rows × columns. */
function openTablePicker(): void {
  document.querySelector(".gridpick")?.remove();
  const panel = document.createElement("div");
  panel.className = "menu gridpick";
  const info = document.createElement("div");
  info.className = "gridpick-info";
  info.textContent = "表格";
  const cells: HTMLButtonElement[] = [];
  for (let r = 1; r <= 8; r++) {
    for (let c = 1; c <= 10; c++) {
      const b = document.createElement("button");
      b.dataset.r = String(r);
      b.dataset.c = String(c);
      b.onmouseenter = () => {
        for (const x of cells) x.classList.toggle("hot", Number(x.dataset.r) <= r && Number(x.dataset.c) <= c);
        info.textContent = `${c} × ${r} 表格`;
      };
      b.onclick = () => {
        panel.remove();
        editor?.insertTable(r, c);
      };
      cells.push(b);
      panel.append(b);
    }
  }
  panel.append(info);
  const at = (ctl.table as HTMLElement | undefined)?.getBoundingClientRect();
  panel.style.left = `${Math.max(8, at ? at.left : 120)}px`;
  panel.style.top = `${(at?.bottom ?? 96) + 6}px`;
  document.body.append(panel);
  window.addEventListener("pointerdown", (e) => {
    if (!panel.contains(e.target as Node)) panel.remove();
  }, { once: true });
}

/** A local image: the shell picks it, the renderer decodes it into a data URL for the deck. */
async function pickDeckPicture(): Promise<void> {
  const bridge = window.soloOffice;
  if (bridge) {
    const picked = await bridge.pickImage();
    if (!picked) return;
    const ext = picked.name.slice(picked.name.lastIndexOf(".") + 1).toLowerCase();
    const mime = IMAGE_MIME_BY_EXT[ext] ?? "image/png";
    await deck?.insertPicture(dataUrl(mime, picked.data), picked.name.replace(/\.[^.]+$/, "") || "图片");
    return;
  }
  const picker = document.createElement("input");
  picker.type = "file";
  picker.accept = "image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml";
  picker.onchange = () => {
    const file = picker.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const url = String(reader.result ?? "");
      if (url.startsWith("data:image/")) void deck?.insertPicture(url, file.name.replace(/\.[^.]+$/, "") || "图片");
    };
    reader.readAsDataURL(file);
  };
  picker.click();
}

function syncDeckRibbon(): void {
  const state = deck?.styleState();
  const mark = (key: string, want: boolean): void => {
    (ctl[key] as HTMLButtonElement | undefined)?.classList.toggle("on", want);
  };
  mark("bold", state?.bold ?? false);
  mark("italic", state?.italic ?? false);
  mark("under", state?.underline ?? false);
  for (const key of ["bold", "italic", "under"] as const) {
    const b = ctl[key] as HTMLButtonElement | undefined;
    if (b) b.disabled = !state;
  }
  const shape = deck?.shapeStyleState();
  const size = ctl.size as HTMLSelectElement | undefined;
  const sizes = shape?.sizePt && DECK_SIZES.includes(shape.sizePt) ? shape.sizePt : DECK_SIZES[3];
  if (size) size.value = String(sizes);
  for (const key of ["size", "color", "fill", "line"] as const) {
    const el = ctl[key] as HTMLInputElement | HTMLSelectElement | undefined;
    if (el) el.disabled = !shape;
  }
  if (shape?.fill) (ctl.fill as HTMLInputElement | undefined)!.value = shape.fill.slice(0, 7);
  if (shape?.line) (ctl.line as HTMLInputElement | undefined)!.value = shape.line.slice(0, 7);
}

function selectAllDoc(): void {
  const page = view.querySelector<HTMLElement>(".page");
  if (!page) return;
  const range = document.createRange();
  range.selectNodeContents(page);
  const sel = document.getSelection();
  sel?.removeAllRanges();
  sel?.addRange(range);
  page.focus();
}

function syncRibbon(st: ReturnType<SheetView["activeStyle"]>): void {
  const mark = (key: string, want: boolean): void => {
    const b = ctl[key] as HTMLButtonElement | undefined;
    b?.classList.toggle("on", want);
  };
  mark("bold", st.bold);
  mark("italic", st.italic);
  mark("under", st.underline);
  mark("wrap", st.wrap);
  for (const key of ["left", "center", "right"]) mark(key, st.align === key);
  for (const [key, valign] of [["vtop", "top"], ["vmid", "center"], ["vbot", "bottom"]] as [string, string][]) mark(key, st.valign === valign);
  const fmt = ctl.numfmt as HTMLSelectElement | undefined;
  if (fmt) fmt.value = NUMFMTS.some(([code]) => code === st.numFmt) ? st.numFmt : "General";
  const size = ctl.size as HTMLSelectElement | undefined;
  if (size) size.value = String(SIZES.includes(st.fontSize) ? st.fontSize : SIZES[0]);
  const color = ctl.color as HTMLInputElement | undefined;
  if (color && st.fontColor) color.value = "#" + st.fontColor.replace(/^#/, "").slice(0, 6);
  const fill = ctl.fill as HTMLInputElement | undefined;
  if (fill && st.bgColor) fill.value = "#" + st.bgColor.replace(/^#/, "").slice(0, 6);
}

// ------------------------------------------------------------- undo / redo

function undo(): void {
  if (current?.kind === "docx") {
    editor?.undo();
    syncDocRibbon();
    return;
  }
  if (current?.kind === "pptx") {
    deck?.undo();
    syncDeckRibbon();
    return;
  }
  if (!current || current.kind !== "xlsx") return;
  grid?.cancelEdit();
  if (!history.undoStep(current.wb)) return;
  afterStep(current, current.wb);
}

function redo(): void {
  if (current?.kind === "docx") {
    editor?.redo();
    syncDocRibbon();
    return;
  }
  if (current?.kind === "pptx") {
    deck?.redo();
    syncDeckRibbon();
    return;
  }
  if (!current || current.kind !== "xlsx") return;
  grid?.cancelEdit();
  if (!history.redoStep(current.wb)) return;
  afterStep(current, current.wb);
}

function afterStep(app: Extract<Opened, { kind: "xlsx" }>, wb: Spreadsheet): void {
  const before = grid;
  const keep = before ? { sel: before.sel, active: before.active, sheet: before.sheetIndex } : null;
  dirty = true;
  app.sheet = Math.max(0, Math.min(app.sheet, wb.sheets.length - 1));
  wb.activeSheet = app.sheet;
  title.textContent = app.name + " ●";
  mountSheet(app);
  if (!keep || !grid || keep.sheet !== grid.sheetIndex) return;
  const lastCol = grid.cols[grid.cols.length - 1] ?? 0;
  const c = (v: number): number => Math.max(0, Math.min(lastCol, v));
  const r = (v: number): number => Math.max(0, Math.min(grid!.rowCount - 1, v));
  grid.select(rect(c(keep.sel.c1), r(keep.sel.r1), c(keep.sel.c2), r(keep.sel.r2)), { c: c(keep.active.c), r: r(keep.active.r) });
}

// ------------------------------------------------------------ context menu

interface MenuItem {
  label?: string;
  run?: () => void;
  sep?: boolean;
}

let menuEl: HTMLElement | null = null;

function closeMenu(): void {
  menuEl?.remove();
  menuEl = null;
}

const buildMenu = (items: MenuItem[]): HTMLElement => {
  const menu = document.createElement("div");
  menu.className = "menu";
  for (const item of items) {
    if (item.sep) {
      const hr = document.createElement("div");
      hr.className = "sep";
      menu.append(hr);
      continue;
    }
    const b = document.createElement("button");
    b.textContent = item.label ?? "";
    b.onclick = () => {
      closeMenu();
      item.run?.();
    };
    menu.append(b);
  }
  return menu;
};

function place(menu: HTMLElement, x: number, y: number, above = false): void {
  closeMenu();
  document.body.append(menu);
  menu.style.left = `${Math.max(4, Math.min(x, window.innerWidth - menu.offsetWidth - 8))}px`;
  menu.style.top = `${Math.max(4, Math.min(above ? y - menu.offsetHeight : y, window.innerHeight - menu.offsetHeight - 8))}px`;
  menuEl = menu;
  window.addEventListener("pointerdown", closeMenu, { once: true });
  window.addEventListener("blur", closeMenu, { once: true });
}

function openMenu(v: SheetView, x: number, y: number): void {
  place(
    buildMenu([
      { label: "复制", run: () => void (lastCopy = v.copyText()) },
      {
        label: "剪切",
        run: () => {
          lastCopy = v.copyText();
          v.cut();
        },
      },
      { label: "粘贴", run: () => lastCopy && v.paste(lastCopy) },
      { sep: true },
      { label: "在上方插入行", run: () => v.line("row", true) },
      { label: "删除所在行", run: () => v.line("row", false) },
      { label: "在左侧插入列", run: () => v.line("col", true) },
      { label: "删除所在列", run: () => v.line("col", false) },
      { sep: true },
      { label: "合并单元格", run: () => v.mergeSelection() },
      { label: "取消合并", run: () => v.unmergeSelection() },
      { label: "清除内容", run: () => v.clear("content") },
      { label: "清除格式", run: () => v.clear("format") },
      { sep: true },
      { label: "按此列升序", run: () => v.sortSelection(false) },
      { label: "按此列降序", run: () => v.sortSelection(true) },
      { label: "隐藏此列", run: () => v.hideCol() },
      { sep: true },
      { label: v.hasComment() ? "编辑批注" : "插入批注", run: () => openCommentInput(v, x, y) },
      ...(v.hasComment() ? [{ label: "删除批注", run: () => v.removeComment() }] : []),
    ]),
    x,
    y,
  );
}

/** A tiny input for the note the user is leaving on the active cell. */
function openCommentInput(v: SheetView, x: number, y: number): void {
  document.querySelector(".comment-panel")?.remove();
  const panel = document.createElement("div");
  panel.className = "menu comment-panel";
  const input = document.createElement("input");
  input.type = "text";
  input.placeholder = "批注内容";
  input.value = v.activeComment() ?? "";
  input.style.cssText = "width:190px;font:inherit;font-size:12px;padding:3px 6px;border:1px solid #c9ced6;border-radius:5px;background:#fff;color:var(--ink);";
  const go = () => {
    panel.remove();
    v.setComment(input.value.trim());
  };
  input.addEventListener("keydown", (e) => {
    if ((e as KeyboardEvent).key === "Enter") go();
  });
  const ok = document.createElement("button");
  ok.className = "rb";
  ok.textContent = "确定";
  ok.onclick = go;
  panel.append(input, ok);
  panel.style.left = `${x}px`;
  panel.style.top = `${y}px`;
  document.body.append(panel);
  input.focus();
  window.addEventListener("pointerdown", (e) => {
    if (!panel.contains(e.target as Node)) panel.remove();
  }, { once: true });
}

// ------------------------------------------------------------ find/replace

function openFind(): void {
  document.getElementById("find")?.remove();
  const panel = document.createElement("div");
  panel.id = "find";
  panel.className = "find";
  panel.innerHTML =
    `<label>查找<input id="f-q" /></label><label>替换为<input id="f-r" /></label>` +
    `<button id="f-next">查找下一个</button><button id="f-rep">替换</button><button id="f-all">全部替换</button><span id="f-msg"></span><button id="f-close">×</button>`;
  view.append(panel);
  const q = panel.querySelector<HTMLInputElement>("#f-q")!;
  const r = panel.querySelector<HTMLInputElement>("#f-r")!;
  const msg = panel.querySelector<HTMLElement>("#f-msg")!;
  q.focus();

  const say = (text: string): void => {
    msg.textContent = text;
    flash(text);
  };
  const hits = (): string[] => (current?.kind === "xlsx" && grid ? current.wb.find(grid.sheet.name, q.value) : []);
  const docHits = (): DocPath[] => (current?.kind === "docx" && editor ? editor.matches(q.value) : []);
  const next = (): void => {
    if (current?.kind === "docx" && editor) {
      const found = docHits();
      if (!found.length) return say("没有匹配项");
      editor.focus(found[0], q.value.length);
      return say(`共 ${found.length} 处`);
    }
    const found = hits();
    if (!found.length || !grid) return say("没有匹配项");
    grid.selectRange(found[0]);
    say(`共 ${found.length} 处，定位 ${found[0]}`);
  };
  panel.querySelector("#f-next")!.addEventListener("click", next);
  panel.querySelector("#f-rep")!.addEventListener("click", () => {
    if (current?.kind === "docx" && editor) {
      const found = docHits();
      if (!found.length) return say("没有匹配项");
      const caret = editor.caretPath();
      const at = found.find((f) => caret && f.path.join() === caret.path.join() && f.offset === caret.offset) ?? found[0];
      editor.replaceAt(at, q.value.length, r.value);
      say(`替换了 1 处，还有 ${Math.max(0, found.length - 1)} 处`);
      return;
    }
    if (!grid || current?.kind !== "xlsx") return;
    const at = refOf(grid.active.c, grid.active.r);
    if (!hits().includes(at)) {
      say("当前位置不是匹配项");
      return;
    }
    record();
    const cell: Rect = rect(grid.active.c, grid.active.r, grid.active.c, grid.active.r);
    current.wb.replaceAll(grid.sheet.name, q.value, r.value, cell);
    grid.render();
    next();
  });
  panel.querySelector("#f-all")!.addEventListener("click", () => {
    if (current?.kind === "docx" && editor) {
      let n = 0;
      for (let found = editor.matches(q.value); found.length; found = editor.matches(q.value)) {
        editor.replaceAt(found[0], q.value.length, r.value);
        n += 1;
        if (n > 500) break;
      }
      say(`替换了 ${n} 处`);
      return;
    }
    if (!grid || current?.kind !== "xlsx") return;
    record();
    const n = current.wb.replaceAll(grid.sheet.name, q.value, r.value);
    grid.render();
    say(`替换了 ${n} 处`);
  });
  panel.querySelector("#f-close")!.addEventListener("click", () => panel.remove());
  q.addEventListener("keydown", (e) => {
    if ((e as KeyboardEvent).key === "Enter") next();
  });
  q.addEventListener("input", () => {
    if (current?.kind === "docx" && editor) {
      const found = docHits().length;
      return say(found ? `找到 ${found} 处` : "没有匹配项");
    }
    const found = hits().length;
    say(found ? `找到 ${found} 处` : "没有匹配项");
  });
}

// --------------------------------------------------------------- sheet tabs

function mountSheetTabs(app: Extract<Opened, { kind: "xlsx" }>): void {
  sheetbar.replaceChildren();
  app.wb.sheets.forEach((s, i) => {
    const b = document.createElement("button");
    b.textContent = s.name;
    b.className = i === app.sheet ? "on" : "";
    b.onclick = () => {
      app.wb.activeSheet = i;
      app.sheet = i;
      mount();
    };
    b.ondblclick = () => renameTab(app, i, b);
    b.oncontextmenu = (e) => {
      e.preventDefault();
      tabMenu(app, i, e.clientX, e.clientY);
    };
    sheetbar.append(b);
  });
  const add = document.createElement("button");
  add.className = "add";
  add.textContent = "＋";
  add.title = "新建工作表";
  add.onclick = () => {
    record();
    app.sheet = app.wb.sheets.indexOf(app.wb.addSheet());
    mount();
  };
  sheetbar.append(add);
  const status = document.createElement("div");
  status.className = "status";
  status.id = "status";
  sheetbar.append(status);
}

function tabMenu(app: Extract<Opened, { kind: "xlsx" }>, index: number, x: number, y: number): void {
  place(
    buildMenu([
      { label: "重命名", run: () => renameTab(app, index, sheetbar.children[index] as HTMLElement) },
      {
        label: "移到最后",
        run: () => {
          record();
          app.wb.sheets.push(app.wb.sheets.splice(index, 1)[0]);
          app.sheet = app.wb.sheets.length - 1;
          mount();
        },
      },
      {
        label: "删除工作表",
        run: () => {
          record();
          if (!app.wb.removeSheet(index)) return flash("至少需要保留一张工作表");
          app.sheet = Math.max(0, Math.min(app.sheet, app.wb.sheets.length - 1));
          mount();
        },
      },
    ]),
    x,
    y,
    true,
  );
}

function renameTab(app: Extract<Opened, { kind: "xlsx" }>, index: number, tab: HTMLElement): void {
  const box = document.createElement("input");
  box.className = "rename";
  box.value = app.wb.sheets[index].name;
  tab.replaceChildren(box);
  box.focus();
  box.select();
  let done = false;
  const commit = (): void => {
    if (done) return;
    done = true;
    const name = box.value.trim();
    if (name && name !== app.wb.sheets[index].name) {
      record();
      app.wb.renameSheet(index, name);
    }
    mount();
  };
  box.addEventListener("blur", commit);
  box.addEventListener("keydown", (e) => {
    const ke = e as KeyboardEvent;
    if (ke.key === "Enter") commit();
    if (ke.key === "Escape") {
      done = true;
      mount();
    }
  });
}

// ---------------------------------------------------------------- formula bar

input.addEventListener("keydown", (e) => {
  const ke = e as KeyboardEvent;
  if (!grid || ke.isComposing) return;
  if (ke.key === "Escape") {
    ke.preventDefault();
    refreshChrome();
    grid.focusGrid();
    return;
  }
  if (ke.key !== "Enter" && ke.key !== "Tab") return;
  ke.preventDefault();
  const at = { ...grid.active };
  grid.write(at.c, at.r, input.value);
  grid.stepBy(ke.key === "Tab" ? (ke.shiftKey ? -1 : 1) : 0, ke.key === "Enter" ? (ke.shiftKey ? -1 : 1) : 0);
  grid.focusGrid();
});

// -------------------------------------------------------------------- docx

function mountDoc(app: Extract<Opened, { kind: "docx" }>): void {
  sheetUI(true, false);
  editor = new DocEditor(view, app.doc, {
    changed: () => {
      dirty = true;
      title.textContent = app.name + " ●";
      syncDocRibbon();
      // `changed` is the one signal that the document's content moved — typing, undo, a ribbon command all end
      // here — so it is where the 状态栏那一格 is re-read. 光标移动不算，走 `moved`，那时数没变。
      docStatus?.refresh();
    },
    // 选区移动不是编辑，走不到 `changed`：光标从粗字挪到不粗的字上，ribbon 得跟着改口；
    // 状态栏的 第 X 页 也是跟着光标走的，而字数在光标移动时并不变，所以这里只重读状态栏。
    moved: () => {
      syncDocRibbon();
      docStatus?.refresh();
    },
    // 换修订视图那一趟：纸重铺了（藏起来的字不再占行盒，页数与页缝都跟着变），文件却一个字节没动，
    // 所以既不到 `changed` 也不到 `moved`。状态栏那五格是从铺排现读的，没人报这一趟它就挂着换档前的数。
    relayout: () => {
      docStatus?.refresh();
    },
    // 窗格开合变了：件没改字节、纸没重铺，要重读的只有 ribbon 上那三档挂不挂勾。「显示批注」挂着勾时点它是收起
    // 还是再开一次，问的就是这个勾。
    panesChanged: () => {
      syncDocRibbon();
    },
  });
  docStatus = new DocStatusBar(editor);
  // The ribbon reads the editor, so it is built after one exists.
  buildRibbon("docx");
  // `force`：这一趟是换了一份文件，上一份里打了还没落笔的那串字（字体框、四格页边距）不搬过来。
  syncDocRibbon(true);
  // 光标放到文首（WPS 同款）：字体/字号两格的回显问的是光标，没有光标两格只能空着；状态栏的行列
  // 也是跟着光标走的。放光标走 `moved`，sync 自己会再来一趟，不必在这里补一句。
  editor.caretToDocumentStart();
  // 打开即按窗宽定初始缩放（#243「页宽」档）：缩放 = 可用宽 ÷ 纸宽（内容宽 − 滚动条预留），两位小数，
  // 钳到与 ribbon 缩放同一界（0.4–2.5）。WPS 打开 `event plan.dotm` 状态栏显 121%，同一档的形状
  // （窗口宽/纸宽恰合）；它的精确式子与「是否记住上次缩放」未取证（锁屏挡着受控实验）。
  // §184 复验：page-zoom-dom 两档（93%/159%）0 红、table-tools 141 例 ×zoom 断言修正后单跑全绿；
  // 聚合 iframe 的宽视口（1.24×）下该页仍有断言与 saveBytes 兜底的交互问题，挂账 §184。
  const paperW = twipsToPx((app.doc.sections[0] ?? DEFAULT_SECTION).pageWidthTwips);
  if (paperW > 0) applyZoom((view.clientWidth - 15) / paperW);
}

// -------------------------------------------------------------------- pptx

/** Full-screen show; the editor underneath keeps its state untouched. */
function startShow(fromCurrent: boolean): void {
  if (current?.kind !== "pptx" || !deck || shower) return;
  deck.commit();
  shower = new DeckShow(current.pres, fromCurrent ? deck.index : 0, () => (shower = null));
  shower.open();
}

function mountDeck(app: Extract<Opened, { kind: "pptx" }>): void {
  sheetUI(true, false);
  deck = new DeckEditor(view, app.pres, app.slide, {
    changed: () => {
      dirty = true;
      title.textContent = app.name + " ●";
      syncDeckRibbon();
    },
    selected: () => syncDeckRibbon(),
  });
  // The ribbon reads the editor, so it is built after one exists.
  buildRibbon("pptx");
  syncDeckRibbon();
}

// --------------------------------------------------------------- commands

async function openCommand(): Promise<void> {
  if (!(await confirmLeave())) return;
  const bridge = window.soloOffice;
  if (bridge) {
    const picked = await bridge.pickFile();
    if (picked) {
      await openBytes(picked.name, picked.data, picked.path);
    }
    return;
  }
  $("file-input").click();
}

/** Opens bytes the shell handed over (drop, `solo:load`), after the unsaved-changes prompt. */
async function openWithGuard(name: string, data: Uint8Array, path?: string): Promise<void> {
  if (!(await confirmLeave())) return;
  await openBytes(name, data, path);
}

async function newCommand(kind: "xlsx" | "docx" | "pptx"): Promise<void> {
  shower?.close();
  if (!(await confirmLeave())) return;
  if (kind === "xlsx") {
    history = new History();
    current = { kind: "xlsx", name: "工作簿1.xlsx", wb: Spreadsheet.blank(), sheet: 0 };
  } else if (kind === "docx") {
    current = { kind: "docx", name: "文档1.docx", doc: WordDocument.blank() };
  } else {
    current = { kind: "pptx", name: "演示文稿1.pptx", pres: Presentation.blank(), slide: 0 };
  }
  // A new document has never been written anywhere; save asks for a place first.
  dirty = false;
  title.textContent = current.name;
  saveBtn.disabled = false;
  mount();
}

/**
 * The user's answer to leaving a modified document: `true` means go ahead —
 * the file was saved, or the modifications were dropped on purpose.
 */
async function confirmLeave(): Promise<boolean> {
  if (!dirty || !current) return true;
  const bridge = window.soloOffice;
  const answer: "save" | "discard" | "cancel" = bridge
    ? await bridge.confirmUnsaved(displayName())
    : window.confirm(`${displayName()} 的修改尚未保存，放弃修改吗？`)
      ? "discard"
      : "cancel";
  if (answer === "cancel") return false;
  if (answer === "save") return saveCommand();
  return true;
}

/** `path` may be absent in the browser build, where a save can only download. */
async function saveCommand(forceDialog = false): Promise<boolean> {
  if (!current) return false;
  grid?.commitEdit();
  editor?.commit();
  deck?.commit();
  if (current.kind === "xlsx") {
    const data = writeXlsx(current.wb);
    const name = current.name.replace(/\.xls[xm]$/i, "") + ".xlsx";
    return writeOut(current, name, data, forceDialog, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  }
  if (current.kind === "docx") {
    const data = writeDocx(current.doc);
    const name = current.name.replace(/\.(docx|dotx)$/i, "") + ".docx";
    return writeOut(current, name, data, forceDialog, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  }
  const data = writePptx(current.pres);
  const name = current.name.replace(/\.(pptx|potx|ppsx)$/i, "") + ".pptx";
  return writeOut(current, name, data, forceDialog, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
}

async function writeOut(app: Opened, name: string, data: Uint8Array, forceDialog: boolean, mime: string): Promise<boolean> {
  const bridge = window.soloOffice;
  if (!bridge) {
    download(name, new Blob([new Uint8Array(data)], { type: mime }));
    markSaved(app, name);
    return true;
  }
  // A document that already lives in a file saves back into it; otherwise — or on
  // a write failure, a read-only file for instance — the Save-As dialog decides.
  if (!forceDialog && app.path) {
    const error = await bridge.writeFile(app.path, data);
    if (error === null) {
      markSaved(app, app.name);
      return true;
    }
    flash(`原位保存失败：${error}`);
  }
  const target = await bridge.saveFile(name, data);
  if (!target) return false;
  app.path = target;
  markSaved(app, baseName(target));
  return true;
}

const baseName = (path: string): string => path.split(/[\\/]/).pop() ?? path;

const displayName = (): string => (current ? baseName(current.path ?? current.name) : "文档");

function markSaved(app: Opened, name: string): void {
  app.name = name;
  dirty = false;
  title.textContent = `${name} · 已保存`;
}

function download(name: string, blob: Blob): void {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

// ------------------------------------------------------------------ wiring

$("open-btn").onclick = openCommand;
saveBtn.onclick = () => void saveCommand();
// 快速访问那一排的其余三枚（`QAT` 默认命令里的 PrintToPDF / Undo / Redo）走的是同一条命令，
// 不另开路径：PDF 那一枚就是 `solo:export-pdf` 那趟，撤销/重做按当前打开的是哪种件分派。
$("pdf-btn").onclick = () => window.dispatchEvent(new CustomEvent("solo:export-pdf"));
$("undo-btn").onclick = () => undo();
$("redo-btn").onclick = () => redo();
// The start screen's 新建 buttons live inside #view and are rewritten on every
// state change, so the click is caught by delegation instead of re-binding.
view.addEventListener("click", (e) => {
  const kind = (e.target as HTMLElement).closest?.<HTMLElement>("[data-new]")?.dataset.new;
  if (kind === "xlsx" || kind === "docx" || kind === "pptx") void newCommand(kind);
});
$("file-input").addEventListener("change", async (e) => {
  const f = (e.target as HTMLInputElement).files?.[0];
  if (!f) return;
  await openWithGuard(f.name, new Uint8Array(await f.arrayBuffer()));
  (e.target as HTMLInputElement).value = "";
});

document.addEventListener("dragover", (e) => {
  e.preventDefault();
  document.getElementById("drop")?.classList.add("hot");
});
document.addEventListener("dragleave", () => document.getElementById("drop")?.classList.remove("hot"));
document.addEventListener("drop", async (e) => {
  e.preventDefault();
  document.getElementById("drop")?.classList.remove("hot");
  const f = (e as DragEvent).dataTransfer?.files?.[0];
  if (f) await openWithGuard(f.name, new Uint8Array(await f.arrayBuffer()));
});

window.addEventListener("keydown", (e) => {
  const ke = e as KeyboardEvent;
  if (shower) return; // the slideshow owns the keyboard
  if (current?.kind !== "pptx" || !deck) return;
  const active = document.activeElement;
  if (active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement || active instanceof HTMLSelectElement) return;
  if (active instanceof HTMLElement && active.isContentEditable) return;
  if (ke.metaKey || ke.ctrlKey || ke.altKey) return;
  if (ke.key === "Delete" || ke.key === "Backspace") {
    ke.preventDefault();
    deck.deleteShape();
    return;
  }
  const step = ke.shiftKey ? 10 : 1;
  const by: Record<string, [number, number]> = {
    ArrowLeft: [-step, 0],
    ArrowRight: [step, 0],
    ArrowUp: [0, -step],
    ArrowDown: [0, step],
  };
  const delta = by[ke.key];
  if (delta) {
    ke.preventDefault();
    deck.nudge(delta[0], delta[1]);
  }
});

window.addEventListener("keydown", (e) => {
  const ke = e as KeyboardEvent;
  if (shower) return; // the slideshow owns the keyboard
  if (!(ke.metaKey || ke.ctrlKey)) return;
  const key = ke.key.toLowerCase();
  if (key === "o") {
    ke.preventDefault();
    void openCommand();
    return;
  }
  if (key === "s") {
    ke.preventDefault();
    void saveCommand(ke.shiftKey);
    return;
  }
  if (current?.kind === "pptx" && ke.isComposing === false) {
    if (key === "z") {
      ke.preventDefault();
      return ke.shiftKey ? redo() : undo();
    }
    if (key === "y") {
      ke.preventDefault();
      return redo();
    }
    if (key === "b" || key === "i" || key === "u") {
      ke.preventDefault();
      deck?.applyStyle(key === "b" ? "bold" : key === "i" ? "italic" : "underline");
      syncDeckRibbon();
      return;
    }
  }
  if (current?.kind === "docx" && ke.isComposing === false) {
    if (key === "z") {
      ke.preventDefault();
      return ke.shiftKey ? redo() : undo();
    }
    if (key === "y") {
      ke.preventDefault();
      return redo();
    }
    if (key === "a") {
      ke.preventDefault();
      selectAllDoc();
      return;
    }
    if (key === "b" || key === "i" || key === "u") {
      ke.preventDefault();
      const toggle = editor?.markState();
      if (key === "b") editor?.applyStyle({ bold: !toggle?.bold });
      if (key === "i") editor?.applyStyle({ italic: !toggle?.italic });
      if (key === "u") editor?.applyStyle({ underline: toggle?.underline ? "none" : "single" });
      syncDocRibbon();
      return;
    }
    if (key === "f") {
      ke.preventDefault();
      openFind();
      return;
    }
    // WPS 给这条命令的快捷键：`res/wpscommon.kuip` 的快捷键表里 `cmd="InsertColumnBreak"` 只给了
    // `Ctrl+Shift+Return` 一行。这张表在别处都是 Ctrl／Meta 成对出现（分页符、按行拆分表格都成对），所以 macOS
    // 上把 ⌘⇧↵ 一并认下——这一条是**按表的成对惯例推的**，不是取到的字面量。必须排在下面那条 ⌘↵ 之前判。
    if (key === "enter" && ke.shiftKey && (ke.metaKey || ke.ctrlKey)) {
      ke.preventDefault();
      editor?.insertColumnBreak();
      return;
    }
    // WPS 给这条命令的快捷键：`res/wpscommon.kuip:219` 写着 `shortcutHint="Ctrl+Enter"`，快捷键表在
    // 1115/1116 与 1304/1305 各给了一对 `Ctrl+Return`／`Meta+Return`——macOS 上就是 ⌘↵。
    if (key === "enter" && (ke.metaKey || ke.ctrlKey)) {
      ke.preventDefault();
      editor?.insertPageBreak();
      return;
    }
    // WPS 给这条命令的快捷键：`res/wpscommon.kuip:1170–1171`（另一处 1359–1360 同样一对）写着
    // `shortcut cmd="WordCount"` 的 `Ctrl+Shift+G` 与 `Meta+Shift+G`，macOS 上就是 ⌘⇧G。
    if (key === "g" && ke.shiftKey) {
      ke.preventDefault();
      editor?.showWordCount();
      return;
    }
    // WPS 给这条命令的快捷键：`res/wpscommon.kuip:225` 的 `KxLegacyTriggerCommand id="GoTo"` 写着
    // `hotKey="g"`、`shortcutHint="Ctrl+G"`，快捷键表 1127/1128（另一处 1316/1317 逐字重复一组）给了一对
    // `Ctrl+G`／`Meta+G`——macOS 上就是 ⌘G。与上面那条并列、排它之后：带 Shift 的 g 归字数统计。
    if (key === "g") {
      ke.preventDefault();
      editor?.showGoTo();
      return;
    }
  }
  if (current?.kind !== "xlsx" || ke.isComposing) return;
  if (key === "f") {
    ke.preventDefault();
    openFind();
    return;
  }
  if (key === "z") {
    ke.preventDefault();
    return ke.shiftKey ? redo() : undo();
  }
  if (key === "y") {
    ke.preventDefault();
    return redo();
  }
  // Formatting shortcuts only while the grid, not a text field, has focus.
  if (grid && !grid.isEditing && document.activeElement !== input) {
    if (key === "b") {
      ke.preventDefault();
      grid.toggle("bold");
    }
    if (key === "i") {
      ke.preventDefault();
      grid.toggle("italic");
    }
    if (key === "u") {
      ke.preventDefault();
      grid.toggle("underline");
    }
  }
});

window.addEventListener("resize", () => {
  if (current?.kind === "pptx") deck?.refit();
  else grid?.refreshFrame();
});

// 缩放入口（#243）：外部（含取证探针）要求某一档缩放，走 applyZoom 的完整管线（钳界、两位小数、
// ribbon 回显与各编辑器的重排）。与 ribbon 那对 ⊖⊕ 同一条命令。
window.addEventListener("solo:zoom", (e) => {
  const z = Number((e as CustomEvent<{ z?: number }>).detail?.z);
  if (Number.isFinite(z) && z > 0) applyZoom(z);
});

window.addEventListener("solo:load", async (e) => {
  const detail = (e as CustomEvent<{ name: string; path?: string; data: Uint8Array }>).detail;
  if (detail?.data) await openWithGuard(detail.name, detail.data, detail.path);
});

window.addEventListener("solo:new", async (e) => {
  const detail = (e as CustomEvent<{ kind?: string }>).detail;
  if (detail?.kind === "xlsx" || detail?.kind === "docx" || detail?.kind === "pptx") await newCommand(detail.kind);
});

window.addEventListener("solo:saveas", () => void saveCommand(true));

window.addEventListener("solo:save", () => void saveCommand());

window.addEventListener("solo:open", () => void openCommand());
window.addEventListener("solo:show", (e) => {
  const detail = (e as CustomEvent<{ start?: "first" | "current" }>).detail;
  startShow(detail?.start === "current");
});
window.addEventListener("solo:export-pdf", () => {
  void (async () => {
    if (!current) return;
    const base = displayName().replace(/\.(xlsx|docx|pptx)$/i, "");
    const result = await window.soloOffice?.exportPdf?.(`${base}.pdf`, current.kind === "pptx");
    if (typeof result === "string") flash(`导出 PDF 失败：${result}`);
    else if (result) flash(`已导出 ${result}`);
  })();
});

window.addEventListener("solo:closerequest", async () => {
  const proceed = await confirmLeave();
  window.soloOffice?.closeReply(proceed);
});

emptyState();
