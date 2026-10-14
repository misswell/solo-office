import { Spreadsheet } from "../core/xlsx.js";
import { WordDocument } from "../core/docx.js";
import { Presentation } from "../core/pptx.js";
import { OpcPackage } from "../core/opc.js";
import { writeXlsx } from "../core/write.js";
import { writeDocx } from "../core/write-docx.js";
import { History } from "../core/history.js";
import { isError } from "../core/formula.js";
import { rangeSpec, rect, refOf, type Rect } from "../core/sheet-ops.js";
import { SheetView } from "./sheet-view.js";
import { DocEditor, type DocPath, type RevisionView } from "./doc-edit.js";
import { DocStatusBar } from "./doc-status-bar.js";
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
const DOC_SIZES = [10, 10.5, 12, 14, 16, 18, 24, 36];
const DECK_SIZES = [12, 14, 16, 18, 20, 24, 28, 32, 40, 44];

/**
 * 突出显示's palette. The names are what `w:highlight w:val=` can hold (§17.17.64 has 17 members); these eight are
 * what this ribbon can paint. Measured over the 86 WPS documents, only `yellow` and `none` ever reach a run with
 * text in them, so the palette gap has no audience today and is left alone (`.scratch/ribbon-echo-measure.ts`).
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

function buildRibbon(kind: "xlsx" | "docx" | "pptx"): void {
  if (ribbon.dataset.built === kind) return;
  ribbon.dataset.built = kind;
  ribbon.innerHTML = "";
  for (const key of Object.keys(ctl)) delete ctl[key];
  if (kind === "xlsx") buildSheetRibbon();
  else if (kind === "docx") buildDocRibbon();
  else buildDeckRibbon();
}

function buildSheetRibbon(): void {
  const group = (): HTMLElement => {
    const g = document.createElement("div");
    g.className = "grp";
    ribbon.append(g);
    return g;
  };
  const button = (g: HTMLElement, key: string, label: string, tip: string, run: () => void): void => {
    const b = document.createElement("button");
    b.className = "rb";
    b.textContent = label;
    b.title = tip;
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
  const chooser = (g: HTMLElement, key: string, tip: string, options: [string, string][], run: (value: string) => void): void => {
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

  const nav = group();
  button(nav, "undo", "↶", "撤销 ⌘Z", undo);
  button(nav, "redo", "↷", "重做 ⌘⇧Z", redo);
  button(nav, "", "＋行", "在选区上方插入行", () => use((v) => v.line("row", true)));
  button(nav, "", "－行", "删除选中的行", () => use((v) => v.line("row", false)));
  button(nav, "", "＋列", "在选区左侧插入列", () => use((v) => v.line("col", true)));
  button(nav, "", "－列", "删除选中的列", () => use((v) => v.line("col", false)));

  const numbers = group();
  chooser(numbers, "numfmt", "数字格式", NUMFMTS, (v) => use((g) => g.setNumFmt(v)));
  button(numbers, "", "0.", "减少小数位", () => use((v) => v.stepDecimals(-1)));
  button(numbers, "", "0.0+", "增加小数位", () => use((v) => v.stepDecimals(1)));
  button(numbers, "wrap", "换行", "自动换行", () => use((v) => v.toggle("wrap")));

  const font = group();
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

  const align = group();
  button(align, "left", "左", "左对齐", () => use((v) => v.setAlign("left")));
  button(align, "center", "中", "水平居中", () => use((v) => v.setAlign("center")));
  button(align, "right", "右", "右对齐", () => use((v) => v.setAlign("right")));
  button(align, "vtop", "上", "顶端对齐", () => use((v) => v.setValign("top")));
  button(align, "vmid", "↕", "垂直居中", () => use((v) => v.setValign("center")));
  button(align, "vbot", "下", "底端对齐", () => use((v) => v.setValign("bottom")));

  const cells = group();
  button(cells, "filter", "筛选", "为选区或当前数据区域启用/关闭筛选", () => use((v) => v.toggleFilter()));
  button(cells, "", "清除筛选", "显示全部记录并保留筛选按钮", () => use((v) => v.clearFilters()));
  button(cells, "", "无框", "清除边框", () => use((v) => v.setBorder("none")));
  button(cells, "", "外框", "绘制外边框", () => use((v) => v.setBorder("box")));
  button(cells, "", "所有框", "绘制全部边框", () => use((v) => v.setBorder("all")));
  button(cells, "merge", "合并", "合并单元格", () => use((v) => v.mergeSelection()));
  button(cells, "", "拆分", "取消合并", () => use((v) => v.unmergeSelection()));
  button(cells, "", "清格式", "只保留内容", () => use((v) => v.clear("format")));

  const pic = group();
  button(pic, "", "图片", "在当前单元格插入本地图片", () => void pickPicture((url, name, w, h) => grid?.insertPicture(url, w, h, name)));

  const cf = group();
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

  const charts = group();
  button(charts, "", "柱状图", "用选区数据插入柱状图", () => use((v) => v.insertChart("bar")));
  button(charts, "", "折线图", "用选区数据插入折线图", () => use((v) => v.insertChart("line")));
  button(charts, "", "饼图", "用选区数据插入饼图", () => use((v) => v.insertChart("pie")));

  button(charts, "", "编辑图表", "选择图表后编辑类型、配色、标题和数据区域；也可双击图表", () => use((v) => v.editSelectedChart()));

  ribbon.append(zoomGroup());

  const data = group();
  button(data, "", "冻结", "在当前位置冻结窗格", () => use((v) => v.freezeHere()));
  button(data, "", "解冻", "解除冻结窗格", () => use((v) => v.unfreeze()));
  button(data, "", "升序", "按当前列升序排列选区", () => use((v) => v.sortSelection(false)));
  button(data, "", "降序", "按当前列降序排列选区", () => use((v) => v.sortSelection(true)));
  button(data, "", "查找", "查找与替换 ⌘F", openFind);
}

/** Word-style formatting: run styling over the selection, paragraph shape for the whole paragraph. */
function buildDocRibbon(): void {
  const group = (): HTMLElement => {
    const g = document.createElement("div");
    g.className = "grp";
    ribbon.append(g);
    return g;
  };
  const button = (g: HTMLElement, key: string, label: string, tip: string, run: () => void): void => {
    const b = document.createElement("button");
    b.className = "rb";
    b.textContent = label;
    b.title = tip;
    b.onclick = run;
    ctl[key] = b;
    g.append(b);
  };
  const chooser = (g: HTMLElement, key: string, tip: string, options: [string, string][], run: (value: string) => void): void => {
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
  const style = (patch: Parameters<DocEditor["applyStyle"]>[0]): void => {
    editor?.applyStyle(patch);
    syncDocRibbon();
  };

  const nav = group();
  button(nav, "undo", "↶", "撤销 ⌘Z", () => undo());
  button(nav, "redo", "↷", "重做 ⌘⇧Z", () => redo());
  button(nav, "", "全选", "选中整篇", () => selectAllDoc());

  const font = group();
  button(font, "bold", "粗", "加粗 ⌘B", () => style({ bold: !editor?.markState().bold }));
  button(font, "italic", "斜", "斜体 ⌘I", () => style({ italic: !editor?.markState().italic }));
  button(font, "under", "线", "下划线 ⌘U", () => style({ underline: !editor?.markState().underline }));
  button(font, "strike", "删", "删除线", () => style({ strike: !editor?.markState().strike }));
  chooser(font, "size", "字号", DOC_SIZES.map((s): [string, string] => [String(s), String(s)]), (v) => style({ sizePt: Number(v) }));
  swatch(font, "color", "字体颜色", "#c00000", (v) => style({ color: v }));
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

  const align = group();
  button(align, "left", "左", "左对齐", () => editor?.setParagraph({ align: "" }));
  button(align, "center", "中", "居中", () => editor?.setParagraph({ align: "center" }));
  button(align, "right", "右", "右对齐", () => editor?.setParagraph({ align: "right" }));
  button(align, "both", "两", "两端对齐", () => editor?.setParagraph({ align: "both" }));
  // WPS 把这两条放在主页段落组里（`IncreaseIndent`/`DecreaseIndent`，图标 `indent_increase_kd`/`indent_decrease_kd`，
  // Touch Bar 图标的中文标题写的就是「增加缩进量」「减少缩进量」），右键菜单里也有同名的一对。
  button(align, "outdent", "缩进−", "减少缩进量", () => editor?.shiftParagraphIndent(-1));
  button(align, "indent", "缩进+", "增加缩进量", () => editor?.shiftParagraphIndent(1));

  const lists = group();
  button(lists, "bullet", "• 符号", "项目符号", () => editor?.setParagraph({ list: "bullet" }));
  button(lists, "number", "1. 编号", "编号列表", () => editor?.setParagraph({ list: "number" }));
  button(lists, "nolist", "取消列表", "变回普通段落", () => editor?.setParagraph({ list: "none" }));

  const insert = group();
  button(insert, "table", "表格", "插入表格", openTablePicker);
  button(insert, "picture", "图片", "插入本地图片", () => void pickPicture((url, name) => void editor?.insertPicture(url, name)));
  // WPS 把三条断命令挂在同一个下拉里（`wpsentries.kuip`：`CT_Insert.RB_Pages.InsertBreaks.*` 与
  // `CT_PageLayout.RB_Page_Setup.m_Breaks.*`、`InsertMenu.InsertBreakMenu.*`）。按钮文案只有分页符取到过证据
  // （`InsertPageBreak`），换行符那条 kuip 里写的是 `text="@Text Wrapping Break"` 这个待译占位，中文措辞在 `.qm`
  // 里读不出，所以「换行符／分栏符」两个词是按 WPS 家族习惯的**推断措辞**，不是取到的字面量。
  button(insert, "pagebreak", "分页符", "从这里另起一页", () => editor?.insertPageBreak());
  button(insert, "linebreak", "换行符", "在这里软换行，不另起段落", () => editor?.insertLineBreak());
  button(insert, "columnbreak", "分栏符", "从这里另起一栏", () => editor?.insertColumnBreak());
  button(insert, "header", "页眉", "编辑页眉（无则新建）", () => editor?.openHF("header"));
  button(insert, "footer", "页脚", "编辑页脚（无则新建）", () => editor?.openHF("footer"));

  // 审阅: the commands a reader needs to work through somebody else's changes. WPS keeps these on their own
  // tab; this ribbon is one row, so they get their own group at the end, before the zoom control.
  const review = group();
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
  button(review, "revAccept", "接受全部", "接受修订：插入的留下，删除的生效", () => {
    editor?.acceptRevisions();
    syncDocRibbon();
  });
  button(review, "revReject", "拒绝全部", "拒绝修订：删除的文字回来，插入的文字去掉", () => {
    editor?.rejectRevisions();
    syncDocRibbon();
  });
  button(review, "wordCount", "字数统计", "字数统计：报出段落数、字数与字符数", () => editor?.showWordCount());
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

  // 布局: the page itself rather than the text on it. WPS keeps these on a 「页面」 tab; this ribbon is one row,
  // so the page's own settings get a group at the end, before the zoom control. The 导航窗格 belongs with them:
  // WPS puts it on the 视图 tab, and what it changes is the workspace around the page, not the text on it.
  const layout = group();
  button(layout, "pageSetup", "页面设置", "纸张大小与方向、页边距、页眉页脚距离、装订线", () => editor?.openPageSetup());
  button(layout, "navPane", "导航窗格", "按大纲级别列出各级标题，点一行跳到那一处", () => {
    editor?.showOutlinePane();
    syncDocRibbon();
  });

  ribbon.append(zoomGroup());
}

function syncDocRibbon(): void {
  const state = editor?.markState();
  const mark = (key: string, want: boolean): void => {
    (ctl[key] as HTMLButtonElement | undefined)?.classList.toggle("on", want);
  };
  mark("bold", state?.bold ?? false);
  mark("italic", state?.italic ?? false);
  mark("under", state?.underline ?? false);
  mark("strike", state?.strike ?? false);
  mark("left", state?.align === undefined);
  mark("center", state?.align === "center");
  mark("right", state?.align === "right");
  mark("both", state?.align === "both");
  mark("bullet", state?.list === "bullet");
  mark("number", state?.list === "number");
  const size = ctl.size as HTMLSelectElement | undefined;
  if (size) {
    const pt = state?.sizePt;
    echoChoice(size, pt === undefined ? undefined : String(pt), DOC_SIZES.map(String), DOC_SIZES.filter((n) => n < (pt ?? 0)).length);
  }
  const hl = ctl.highlight as HTMLSelectElement | undefined;
  // Three sentences the file could have written — 没提过底色 / 提了「没有」 / 提了某个颜色名 — three readings of
  // the box. Leaving it blank is the first, and only the `none` option spells the second.
  if (hl) echoChoice(hl, state?.highlight, DOC_HIGHLIGHTS.map(([v]) => v), DOC_HIGHLIGHTS.length);
  mark("track", editor?.trackingChanges ?? false);
  mark("revPane", editor?.reviewPaneOpen ?? false);
  mark("navPane", editor?.outlinePaneOpen ?? false);
  const revView = ctl.revView as HTMLSelectElement | undefined;
  if (revView && editor) revView.value = editor.revisionShown;
  const undoBtn = ctl.undo as HTMLButtonElement | undefined;
  const redoBtn = ctl.redo as HTMLButtonElement | undefined;
  if (undoBtn) undoBtn.disabled = !editor?.canUndo;
  if (redoBtn) redoBtn.disabled = !editor?.canRedo;
}

/**
 * A chooser states the sentence the caret's run says, not the closest thing it happens to list.
 *
 * `undefined` deselects, which is how 「没提过」 reads; `value` among `presets` selects that option; a value the
 * control does not list gets one option of its own at `insertAt`, taken back out on the next sync. Showing the
 * nearest preset instead would put a property in the ribbon that is not in the file — measured over the 86 WPS
 * documents, 61 of them hold runs whose size is none of the eight presets, so a 21pt caret used to be answered
 * with 16 on every stop (`.scratch/ribbon-echo-measure.ts`).
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
  const group = (): HTMLElement => {
    const g = document.createElement("div");
    g.className = "grp";
    ribbon.append(g);
    return g;
  };
  const button = (g: HTMLElement, key: string, label: string, tip: string, run: () => void): void => {
    const b = document.createElement("button");
    b.className = "rb";
    b.textContent = label;
    b.title = tip;
    b.onclick = run;
    ctl[key] = b;
    g.append(b);
  };
  const style = (kind: "bold" | "italic" | "underline"): void => {
    deck?.applyStyle(kind);
    syncDeckRibbon();
  };
  const chooser = (g: HTMLElement, key: string, tip: string, options: [string, string][], run: (value: string) => void): void => {
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

  const nav = group();
  button(nav, "undo", "↶", "撤销 ⌘Z", () => undo());
  button(nav, "redo", "↷", "重做 ⌘⇧Z", () => redo());

  const pages = group();
  button(pages, "", "＋新页", "在当前页后新建一页", () => deck?.addSlide());
  button(pages, "", "删除页", "删除当前页", () => deck?.removeSlide());
  button(pages, "", "上移", "当前页前移一位", () => deck?.moveSlide(-1));
  button(pages, "", "下移", "当前页后移一位", () => deck?.moveSlide(1));

  const insert = group();
  button(insert, "", "文本框", "在当前页插入文本框", () => deck?.insertTextBox());
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
    button(insert, "", label, `插入${label}`, () => deck?.insertPreset(geometry));
  }
  button(insert, "", "图片", "插入本地图片", pickDeckPicture);

  const shape = group();
  button(shape, "", "删除形状", "删除选中的形状、图片或表格", () => deck?.deleteShape());

  const align = group();
  button(align, "aleft", "左", "左对齐", () => deck?.applyAlign("left"));
  button(align, "acenter", "中", "居中", () => deck?.applyAlign("center"));
  button(align, "aright", "右", "右对齐", () => deck?.applyAlign("right"));
  button(align, "aboth", "两", "两端对齐", () => deck?.applyAlign("justify"));

  ribbon.append(zoomGroup());

  const show = group();
  button(show, "", "▶ 放映", "全屏放映，从头开始（F5）", () => startShow(false));
  button(show, "", "本页放映", "从当前页开始放映（⇧F5）", () => startShow(true));

  const font = group();
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
  swatch(font, "fill", "形状填充", "#dce6f1", (v) => {
    deck?.applyShapeFill(v);
    syncDeckRibbon();
  });
  swatch(font, "line", "形状边框", "#4472c4", (v) => {
    deck?.applyShapeLine(v);
    syncDeckRibbon();
  });
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
  input.style.cssText = "width:76px;font:inherit;font-size:12px;padding:3px 6px;border:1px solid #3a404a;border-radius:5px;background:#262b33;color:var(--ink);";
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
  const undoBtn = ctl.undo as HTMLButtonElement | undefined;
  const redoBtn = ctl.redo as HTMLButtonElement | undefined;
  if (undoBtn) undoBtn.disabled = !deck?.canUndo();
  if (redoBtn) redoBtn.disabled = !deck?.canRedo();
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
  input.style.cssText = "width:190px;font:inherit;font-size:12px;padding:3px 6px;border:1px solid #3a404a;border-radius:5px;background:#262b33;color:var(--ink);";
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
  });
  docStatus = new DocStatusBar(editor);
  // The ribbon reads the editor, so it is built after one exists.
  buildRibbon("docx");
  syncDocRibbon();
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
      if (key === "u") editor?.applyStyle({ underline: !toggle?.underline });
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
