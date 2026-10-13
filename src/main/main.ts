import { app, BrowserWindow, dialog, ipcMain, Menu, shell } from "electron";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { basename, join } from "node:path";

const HERE = __dirname;
const FILTERS = [{ name: "Office 文档", extensions: ["xlsx", "xlsm", "docx", "dotx", "pptx", "potx", "zip"] }];
const IMAGE_FILTERS = [{ name: "图片", extensions: ["png", "jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp", "svg"] }];
const NEW_KINDS = [
  { label: "新建表格", kind: "xlsx" },
  { label: "新建文字", kind: "docx" },
  { label: "新建演示", kind: "pptx" },
];

/** Every live window; menu commands go to the focused one. */
const windows = new Set<BrowserWindow>();
/** Windows that answered the close prompt, so the second close goes through. */
const closing = new WeakSet<BrowserWindow>();
/** The newest window — the fallback target when nothing has focus. */
let last: BrowserWindow | null = null;

/** `SOLO_SMOKE` answers dialogs from the environment, so the shell can run without a human. */
const SMOKE_DIR = process.env.SOLO_SMOKE ? join(app.getPath("temp"), "solo-office-smoke") : "";
/** Paths `solo:write` was called with, so a smoke run can report in-place saves. */
const smokeWrites: string[] = [];
/** Paths the Save-As dialog resolved to under smoke, so a run can report what it produced. */
const smokeSaves: string[] = [];
/** A 1×1 PNG, so an image insert can be exercised without a real dialog. */
const SMOKE_PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";

function createWindow(): BrowserWindow {
  const win = new BrowserWindow({
    show: !process.env.SOLO_SMOKE,
    width: 1280,
    height: 820,
    minWidth: 720,
    minHeight: 420,
    title: "Solo Office",
    titleBarStyle: "hiddenInset",
    backgroundColor: "#1c1f24",
    webPreferences: {
      preload: join(HERE, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  windows.add(win);
  last = win;
  void win.loadFile(join(HERE, "index.html"));
  win.on("close", (e) => {
    if (closing.has(win) || win.webContents.isDestroyed()) return;
    // The renderer owns the dirty flag; ask it, and close again once it agrees.
    e.preventDefault();
    win.webContents.send("solo:closerequest");
  });
  win.on("closed", () => windows.delete(win));
  win.webContents.setWindowOpenHandler(({ url }) => {
    void shell.openExternal(url);
    return { action: "deny" };
  });
  return win;
}

/** Menu commands ride to whichever document the user is looking at. */
function target(): BrowserWindow | undefined {
  return BrowserWindow.getFocusedWindow() ?? (last && !last.isDestroyed() ? last : undefined);
}

const send = (channel: string, detail?: unknown): void => target()?.webContents.send(channel, detail);

async function pickFile(parent?: BrowserWindow): Promise<{ name: string; path: string; data: Uint8Array } | null> {
  const options = { properties: ["openFile"] as ("openFile")[], filters: FILTERS };
  const result = parent ? await dialog.showOpenDialog(parent, options) : await dialog.showOpenDialog(options);
  if (result.canceled || !result.filePaths[0]) return null;
  const path = result.filePaths[0];
  app.addRecentDocument(path);
  const buf = await readFile(path);
  return { name: basename(path), path, data: new Uint8Array(buf) };
}

/** A local image for insertion; the renderer turns the bytes into a data URL. */
async function pickImage(parent?: BrowserWindow): Promise<{ name: string; data: Uint8Array } | null> {
  if (SMOKE_DIR) return { name: "smoke.png", data: new Uint8Array(Buffer.from(SMOKE_PNG, "base64")) };
  const options = { properties: ["openFile"] as ("openFile")[], filters: IMAGE_FILTERS };
  const result = parent ? await dialog.showOpenDialog(parent, options) : await dialog.showOpenDialog(options);
  if (result.canceled || !result.filePaths[0]) return null;
  const path = result.filePaths[0];
  const buf = await readFile(path);
  return { name: basename(path), data: new Uint8Array(buf) };
}

async function saveFileAs(parent: BrowserWindow | undefined, name: string, data: Uint8Array | string): Promise<string | null> {
  if (SMOKE_DIR) await mkdir(SMOKE_DIR, { recursive: true });
  const options = { defaultPath: name, filters: [{ name, extensions: [name.split(".").pop() ?? "*"] }] };
  const result = SMOKE_DIR
    ? { canceled: false, filePath: join(SMOKE_DIR, name) }
    : parent
      ? await dialog.showSaveDialog(parent, options)
      : await dialog.showSaveDialog(options);
  if (result.canceled || !result.filePath) return null;
  await writeFile(result.filePath, typeof data === "string" ? data : Buffer.from(data));
  if (SMOKE_DIR) smokeSaves.push(result.filePath);
  app.addRecentDocument(result.filePath);
  return result.filePath;
}

/** Answers the close/new/open prompt: the renderer saves first when the user picks 保存. */
async function confirmUnsaved(parent: BrowserWindow | undefined, name: string): Promise<"save" | "discard" | "cancel"> {
  if (SMOKE_DIR) return "discard";
  const options = {
    type: "warning" as const,
    buttons: ["保存", "不保存", "取消"],
    defaultId: 0,
    cancelId: 2,
    message: `${name} 的修改尚未保存。`,
    detail: "关闭或打开其他文件前，要先保存吗？",
  };
  const result = parent ? await dialog.showMessageBox(parent, options) : await dialog.showMessageBox(options);
  return result.response === 0 ? "save" : result.response === 1 ? "discard" : "cancel";
}

/** Renders the current view to PDF and asks where to put it; returns the path, null on cancel, or an error. */
async function exportPdf(parent: BrowserWindow | undefined, suggested: string, landscape: boolean): Promise<string | null | string> {
  try {
    const bytes = parent && !parent.isDestroyed() ? await parent.webContents.printToPDF({ printBackground: true, landscape, margins: { top: 0, bottom: 0, left: 0, right: 0 } }) : undefined;
    if (!bytes) return "窗口已关闭";
    const options = { defaultPath: suggested, filters: [{ name: "PDF", extensions: ["pdf"] }] };
    const result = SMOKE_DIR
      ? { canceled: false, filePath: join(SMOKE_DIR, suggested) }
      : parent
        ? await dialog.showSaveDialog(parent, options)
        : await dialog.showSaveDialog(options);
    if (result.canceled || !result.filePath) return null;
    await writeFile(result.filePath, bytes);
    if (SMOKE_DIR) smokeSaves.push(result.filePath);
    return result.filePath;
  } catch (e) {
    return (e as Error).message;
  }
}

function menu(): void {
  Menu.setApplicationMenu(
    Menu.buildFromTemplate([
      {
        label: "文件",
        submenu: [
          ...NEW_KINDS.map((k, i) => ({
            label: k.label,
            accelerator: i === 0 ? "CmdOrCtrl+N" : undefined,
            click: () => send("solo:new", { kind: k.kind }),
          })),
          { label: "新建窗口", accelerator: "Shift+CmdOrCtrl+N", click: () => createWindow() },
          { type: "separator" as const },
          { label: "打开…", accelerator: "CmdOrCtrl+O", click: () => send("solo:open") },
          { role: "recentDocuments", submenu: [{ role: "clearRecentDocuments" }] },
          { type: "separator" as const },
          { label: "保存", accelerator: "CmdOrCtrl+S", click: () => send("solo:save") },
          { label: "另存为…", accelerator: "Shift+CmdOrCtrl+S", click: () => send("solo:saveas") },
          { type: "separator" as const },
          { label: "打印…", accelerator: "CmdOrCtrl+P", click: () => target()?.webContents.print({ printBackground: true }) },
          { label: "导出为 PDF…", click: () => send("solo:export-pdf") },
          { type: "separator" as const },
          { role: "close" },
        ],
      },
      {
        label: "编辑",
        submenu: [
          { role: "undo" },
          { role: "redo" },
          { type: "separator" },
          { role: "cut" },
          { role: "copy" },
          { role: "paste" },
          { role: "selectAll" },
        ],
      },
      {
        label: "视图",
        submenu: [
          { role: "reload" },
          { role: "toggleDevTools" },
          { type: "separator" },
          { role: "resetZoom" },
          { role: "zoomIn" },
          { role: "zoomOut" },
          { type: "separator" },
          { role: "togglefullscreen" },
        ],
      },
      {
        label: "放映",
        submenu: [
          { label: "从头开始放映", accelerator: "F5", click: () => send("solo:show", { start: "first" }) },
          { label: "从当前页放映", accelerator: "Shift+F5", click: () => send("solo:show", { start: "current" }) },
        ],
      },
      {
        label: "帮助",
        submenu: [{ label: "关于 Solo Office", click: () => dialog.showMessageBox({ title: "Solo Office", message: "Solo Office", detail: "OOXML 表格 / 文字 / 演示，按 ECMA-376 与 ISO 29500 规范独立实现。" }) }],
      },
    ]),
  );
}

/** File bytes hand over to the focused window, or a fresh one when none fits. */
function deliver(path: string, data: Uint8Array, name: string): void {
  const to = target() ?? createWindow();
  to.webContents.send("solo:load", { name, path, data });
}

ipcMain.handle("solo:pick", (e) => pickFile(BrowserWindow.fromWebContents(e.sender) ?? undefined));
ipcMain.handle("solo:pick-image", (e) => pickImage(BrowserWindow.fromWebContents(e.sender) ?? undefined));
ipcMain.handle("solo:save", (e, arg: { name: string; data: Uint8Array | string }) => saveFileAs(BrowserWindow.fromWebContents(e.sender) ?? undefined, arg.name, arg.data));
ipcMain.handle("solo:confirm", (e, name: string) => confirmUnsaved(BrowserWindow.fromWebContents(e.sender) ?? undefined, name));
/** Writes over the file the document came from, without a dialog. */
ipcMain.handle("solo:write", async (_e, arg: { path: string; data: Uint8Array }) => {
  try {
    await writeFile(arg.path, Buffer.from(arg.data));
    app.addRecentDocument(arg.path);
    if (SMOKE_DIR) smokeWrites.push(arg.path);
    return null;
  } catch (e) {
    return (e as Error).message;
  }
});
/** Renders what the window shows into a PDF file the user names. */
ipcMain.handle("solo:pdf", (e, arg: { name: string; landscape: boolean }) => exportPdf(BrowserWindow.fromWebContents(e.sender) ?? undefined, arg.name, arg.landscape));
ipcMain.on("solo:close-reply", (e, proceed: boolean) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  if (!win || !proceed) return;
  // A smoke run has no one to click the Dock icon again, and macOS keeps an app
  // alive after its last window, so the reply ends the process outright.
  if (SMOKE_DIR) {
    app.exit(0);
    return;
  }
  closing.add(win);
  win.close();
});

/** The renderer asks for window fullscreen; a smoke run skips the OS animation. */
ipcMain.on("solo:fullscreen", (_e, on: boolean) => {
  if (!SMOKE_DIR) target()?.setFullScreen(Boolean(on));
});
function smoke(): void {
  try { require("node:fs").writeFileSync("/tmp/smoke-armed.txt", String(Date.now())); } catch {}
  last?.webContents.once("did-finish-load", () => {
    setTimeout(() => {
      const probe = `(async () => {
        const wait = (ms) => new Promise((r) => setTimeout(r, ms));
        const fire = (name, detail) => window.dispatchEvent(new CustomEvent(name, { detail }));
        const view = document.querySelector("#view");
        const title = document.getElementById("title")?.textContent;
        const page = view?.querySelector(".page");
        if (!page) return { title, view: view?.firstElementChild?.className, heading: view?.querySelector("h1, .para")?.textContent?.slice(0, 40) };
        const blocks = page.children.length;
        const last = page.children[blocks - 1];
        const range = document.createRange();
        range.selectNodeContents(last);
        range.collapse(false);
        const sel = document.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        last.focus();
        document.execCommand("insertText", false, "SMOKE");
        await wait(80);
        const marker = [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent.includes("符号"));
        const bullets = document.querySelectorAll("#view .page .bullet").length;
        // A mark wearing a face its own paragraph does not is what the level's w:rPr asks for. Only the count is
        // reported, so a document's font names never leave the page.
        const bulletWorn = [...document.querySelectorAll("#view .page .bullet")].filter(
          (s) => getComputedStyle(s).fontFamily !== getComputedStyle(s.parentElement ?? document.body).fontFamily,
        ).length;
        marker?.click();
        await wait(150);
        const after = document.querySelector("#view .page");
        const docPart = { title: document.getElementById("title")?.textContent, editable: after?.getAttribute("contenteditable"), blocksBefore: blocks, blocksAfter: after?.children.length, typed: after?.textContent?.includes("SMOKE"), bulletsBefore: bullets, bulletsWearingOwnFont: bulletWorn, bulletsAfter: after?.querySelectorAll(".bullet").length, ribbonGroups: document.querySelectorAll("#ribbon .grp").length };

        // Insertion round: a 2×3 table, a page break and a picture land in the body.
        const rb = (label) => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === label);
        rb("表格")?.click();
        await wait(120);
        const pick = document.querySelector(".gridpick button[data-c='3'][data-r='2']");
        pick?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
        await wait(200);
        const tableRows = document.querySelectorAll("#view .page table.doc-table tr").length;
        const tableCols = document.querySelector("#view .page table.doc-table tr")?.children.length;
        rb("分页符")?.click();
        await wait(150);
        const breaks = document.querySelectorAll('#view .page [style*="page-break"]').length;
        rb("图片")?.click();
        await wait(500);
        const docPics = document.querySelectorAll("#view .page img").length;
        // Header editing: the button creates one when the doc has none.
        rb("页眉")?.click();
        await wait(250);
        const hf = document.querySelector("#view .hf-header");
        hf?.focus();
        document.execCommand("insertText", false, "机密文件");
        await wait(120);
        hf?.blur?.();
        await wait(150);
        const hfText = document.querySelector("#view .hf-header")?.textContent;
        const docInsert = { tableRows, tableCols, breaks, docPics, hfText };

        // 导航窗格: the page leaves one [data-outline] mark per paragraph whose outline level resolves, and the
        // pane lists a row for each in the same order, so the two counts are the invariant the pane is built on.
        const olMarks = document.querySelectorAll("#view .page [data-outline]").length;
        rb("导航窗格")?.click();
        await wait(200);
        const olRows = document.querySelectorAll("#view .solo-ol-item").length;
        const olLevel = [...document.querySelectorAll("#view .solo-ol-item")].map((r) => r.querySelector(".solo-ol-lv")?.textContent);
        const rows = [...document.querySelectorAll("#view .solo-ol-item")];
        rows[rows.length - 1]?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
        await wait(200);
        const docOutline = {
          olMarks,
          olRows,
          paneOpen: !!document.querySelector("#view .solo-ol-panel:not([hidden])"),
          hit: document.querySelectorAll("#view .page .solo-rev-hit[data-outline]").length,
          active: document.querySelectorAll("#view .solo-ol-item.solo-ol-active").length,
          hitLevel: document.querySelector("#view .page .solo-rev-hit[data-outline]")?.getAttribute("data-outline"),
          lastRowLevel: olLevel[olLevel.length - 1],
          levels: [...new Set(olLevel)].length,
        };
        rb("导航窗格")?.click();
        await wait(150);
        docOutline.paneClosed = !document.querySelector("#view .solo-ol-panel:not([hidden])");

        fire("solo:new", { kind: "pptx" });
        await wait(150);
        const pptPart = { slides: document.querySelectorAll("#view .thumb").length, stage: !!document.querySelector("#view .stage"), title: document.getElementById("title")?.textContent };
        // Insert a text box, type into it, restyle it through the ribbon, then
        // insert a shape and a picture through the shell picker.
        const stage = () => document.querySelector("#view .stage .slide");
        const selectAll = (el) => { const r = document.createRange(); r.selectNodeContents(el); const s = getSelection(); s.removeAllRanges(); s.addRange(r); };
        rb("文本框")?.click();
        await wait(200);
        document.querySelector("#view .stage .slide .shape-text")?.focus();
        document.execCommand("insertText", false, "插入的文本框");
        await wait(120);
        selectAll(document.querySelector("#view .stage .slide .shape-text"));
        const sizePick = document.querySelector("#ribbon select.pick");
        if (sizePick) { sizePick.value = "32"; sizePick.dispatchEvent(new Event("change", { bubbles: true })); }
        await wait(120);
        selectAll(document.querySelector("#view .stage .slide .shape-text"));
        const colorPick = document.querySelector('#ribbon input.swatch[title="字体颜色"]');
        if (colorPick) { colorPick.value = "#c00000"; colorPick.dispatchEvent(new Event("input", { bubbles: true })); }
        await wait(120);
        const body6 = document.querySelector("#view .stage .slide .shape-text");
        const spans6 = [...(body6?.querySelectorAll("span") ?? [])].filter((s) => s.textContent.trim());
        const styled = spans6.pop() ?? body6?.querySelector(".sp-p");
        const styledText = styled ? { size: getComputedStyle(styled).fontSize, color: getComputedStyle(styled).color } : null;
        document.activeElement?.blur?.();
        await wait(150);
        rb("矩形")?.click();
        await wait(150);
        const rect = stage()?.querySelectorAll("[data-idx]")[1];
        rect?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        await wait(100);
        const fillPick = document.querySelector('#ribbon input.swatch[title="形状填充"]');
        if (fillPick) { fillPick.value = "#ffd966"; fillPick.dispatchEvent(new Event("input", { bubbles: true })); }
        await wait(150);
        const rectFill = getComputedStyle(stage()?.querySelectorAll("[data-idx]")[1]).backgroundColor;
        rb("图片")?.click();
        await wait(600);
        const pptInsert = {
          shapes: stage()?.querySelectorAll("[data-idx]").length,
          picture: !!stage()?.querySelector("img.shape"),
          text: document.querySelector("#view .stage .slide .shape-text")?.textContent,
          styledText,
          rectFill,
          selected: !!document.querySelector("#view .sel-box"),
        };
        // Slideshow: add two pages, show from the first, step once, black out,
        // then leave with Escape — the overlay must be gone before saving.
        rb("＋新页")?.click();
        await wait(150);
        rb("＋新页")?.click();
        await wait(150);
        const deckPages = document.querySelectorAll("#view .thumb").length;
        fire("solo:show", { start: "first" });
        await wait(250);
        const showStart = { cover: !!document.querySelector(".show"), count: document.querySelector(".show .show-count")?.textContent };
        window.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight" }));
        await wait(150);
        const showNext = document.querySelector(".show .show-count")?.textContent;
        window.dispatchEvent(new KeyboardEvent("keydown", { key: "b" }));
        await wait(120);
        const showBlack = document.querySelector(".show")?.classList.contains("black");
        window.dispatchEvent(new KeyboardEvent("keydown", { key: "b" }));
        await wait(120);
        window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
        await wait(150);
        const pptShow = { deckPages, start: showStart, next: showNext, black: showBlack, left: !document.querySelector(".show") };
        fire("solo:save");
        await wait(600);
        const pptSaved = { title: document.getElementById("title")?.textContent };

        fire("solo:new", { kind: "xlsx" });
        await wait(150);
        const cells = document.querySelectorAll(".sheet-wrap td").length;
        const typeCell = async (text) => {
          const cell = document.querySelectorAll(".sheet-wrap td[data-col]")[3];
          // A synthetic dblclick carries no preceding pointerdown, so the active
          // cell has to be set the way a real click would.
          cell?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
          await wait(40);
          cell?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
          await wait(60);
          const ed = document.querySelector(".cell-editor");
          if (ed) { ed.textContent = text; ed.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })); }
          await wait(80);
        };
        await typeCell("7");
        fire("solo:save");
        await wait(500);
        const firstSave = document.getElementById("title")?.textContent;
        await typeCell("8");
        const editedTitle = document.getElementById("title")?.textContent;
        fire("solo:save");
        await wait(500);
        const sheetPart = { cells, firstSave, editedTitle, savedTitle: document.getElementById("title")?.textContent };
        // Fill handle: the blank workbook has no merge, so D1 holds the typed
        // numbers. Select it, grip down, two rows below, release; series continues.
        const tap = (ref) => document.querySelector('.sheet-wrap td[data-ref="' + ref + '"]');
        tap("D1")?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        await wait(80);
        document.querySelector(".sheet-wrap .fill-box")?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        await wait(80);
        tap("D2")?.dispatchEvent(new PointerEvent("pointermove", { bubbles: true }));
        await wait(60);
        tap("D3")?.dispatchEvent(new PointerEvent("pointermove", { bubbles: true }));
        await wait(60);
        document.querySelector(".sheet-wrap")?.dispatchEvent(new PointerEvent("pointerup", { bubbles: true }));
        await wait(150);
        const sheetFill = { d1: tap("D1")?.textContent, d2: tap("D2")?.textContent, d3: tap("D3")?.textContent };
        rb("图片")?.click();
        await wait(500);
        const sheetPic = { pics: document.querySelectorAll(".sheet-wrap .sheet-pic").length };
        // Conditional format: a three-color scale over the filled series.
        tap("D1")?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        await wait(60);
        tap("D3")?.dispatchEvent(new PointerEvent("pointermove", { bubbles: true }));
        await wait(60);
        document.querySelector(".sheet-wrap")?.dispatchEvent(new PointerEvent("pointerup", { bubbles: true }));
        await wait(80);
        rb("色阶")?.click();
        await wait(150);
        const cfBg = (ref) => document.querySelector('.sheet-wrap td[data-ref="' + ref + '"]')?.style.background;
        const sheetCF = { d1: cfBg("D1"), d2: cfBg("D2"), d3: cfBg("D3") };
        // Highlight rule: values above 8.5 (only D3=10 qualifies) go light red.
        rb("突出显示")?.click();
        await wait(150);
        const condPanel = document.querySelector(".cond-panel");
        const opSel = condPanel?.querySelector("select");
        if (opSel) opSel.value = "greaterThan";
        const numInput = condPanel?.querySelector("input");
        if (numInput) { numInput.value = "8.5"; numInput.dispatchEvent(new Event("input", { bubbles: true })); }
        condPanel?.querySelector("button")?.click();
        await wait(150);
        const cfGt = document.querySelector('.sheet-wrap td[data-ref="D3"]')?.style.background;
        const sheetGT = { d3: cfGt };
        // Cell note: right-click E5, 插入批注, type, confirm — the corner marker shows.
        tap("E5")?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        await wait(60);
        tap("E5")?.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true }));
        await wait(150);
        [...document.querySelectorAll(".menu button")].find((b) => b.textContent === "插入批注")?.click();
        await wait(150);
        const noteInput = document.querySelector(".comment-panel input");
        if (noteInput) { noteInput.value = "请复核"; noteInput.dispatchEvent(new Event("input", { bubbles: true })); }
        document.querySelector(".comment-panel button")?.click();
        await wait(150);
        const noted = document.querySelector('.sheet-wrap td.has-note[data-ref="E5"]');
        const sheetNote = { marked: !!noted, tip: noted?.title };
        // Chart: re-select D1:D3 (the note step moved the selection to E5).
        tap("D1")?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        await wait(60);
        tap("D3")?.dispatchEvent(new PointerEvent("pointermove", { bubbles: true }));
        await wait(60);
        document.querySelector(".sheet-wrap")?.dispatchEvent(new PointerEvent("pointerup", { bubbles: true }));
        await wait(80);
        // The SVG lands beside the data.
        rb("柱状图")?.click();
        await wait(250);
        // Editing follows the same UI paths a user opens by double-clicking.
        document.querySelector(".sheet-chart")?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        document.querySelector(".sheet-chart")?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
        await wait(150);
        const chartDialog = !!document.querySelector("dialog.chart-editor[open]");
        const chartTitle = document.querySelector('[name="chart-title"]');
        const chartType = document.querySelector('[name="chart-type"]');
        const chartPalette = document.querySelector('[name="chart-palette"]');
        const chartSource = document.querySelector('[name="chart-source"]');
        if (chartSource) { chartSource.value = "missing!A1:B2"; chartSource.dispatchEvent(new Event("input", { bubbles: true })); }
        const invalidSourceBlocked = document.querySelector(".chart-apply")?.disabled === true;
        if (chartSource) { chartSource.value = "D1:D3"; chartSource.dispatchEvent(new Event("input", { bubbles: true })); }
        if (chartTitle) { chartTitle.value = "图表编辑探针"; chartTitle.dispatchEvent(new Event("input", { bubbles: true })); }
        if (chartType) { chartType.value = "line"; chartType.dispatchEvent(new Event("change", { bubbles: true })); }
        if (chartPalette) { chartPalette.value = "ocean"; chartPalette.dispatchEvent(new Event("change", { bubbles: true })); }
        document.querySelector(".chart-cancel")?.click();
        const cancelPreserved = !document.querySelector(".sheet-chart svg")?.textContent.includes("图表编辑探针");
        document.querySelector(".sheet-chart")?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
        await wait(100);
        const setChartField = (name, value) => { const el = document.querySelector('[name="' + name + '"]'); if (el) { el.value = value; el.dispatchEvent(new Event("input", { bubbles: true })); } };
        setChartField("chart-title", "图表编辑探针");
        setChartField("chart-type", "line");
        setChartField("chart-palette", "ocean");
        document.querySelector(".chart-apply")?.click();
        await wait(180);
        const appliedTitle = document.querySelector(".sheet-chart svg")?.textContent.includes("图表编辑探针");
        rb("↶")?.click();
        await wait(100);
        const undoTitle = !document.querySelector(".sheet-chart svg")?.textContent.includes("图表编辑探针");
        rb("↷")?.click();
        await wait(100);
        const redoTitle = document.querySelector(".sheet-chart svg")?.textContent.includes("图表编辑探针");
        // Saved chart caches must follow subsequent cell edits.
        tap("D2")?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        tap("D2")?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
        const chartDataEdit = document.querySelector(".cell-editor");
        if (chartDataEdit) { chartDataEdit.textContent = "42"; chartDataEdit.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })); }
        await wait(160);
        const sheetChartEdit = { chartDialog, invalidSourceBlocked, cancelPreserved, appliedTitle, undoTitle, redoTitle, dataEdited: tap("D2")?.textContent === "42", oceanColor: !!document.querySelector('.sheet-chart [stroke="#087F8C"]') };
        let draggedChart = document.querySelector(".sheet-chart");
        const originalLeft = parseFloat(draggedChart?.style.left || "0");
        const originalTop = parseFloat(draggedChart?.style.top || "0");
        draggedChart?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, clientX: 400, clientY: 200 }));
        draggedChart?.dispatchEvent(new PointerEvent("pointermove", { bubbles: true, clientX: 437, clientY: 223 }));
        draggedChart?.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, clientX: 437, clientY: 223 }));
        await wait(100);
        draggedChart = document.querySelector(".sheet-chart");
        const moved = parseFloat(draggedChart?.style.left || "0") > originalLeft && parseFloat(draggedChart?.style.top || "0") > originalTop;
        const originalWidth = parseFloat(draggedChart?.style.width || "0");
        const sizeGrip = draggedChart?.querySelector(".chart-resize");
        sizeGrip?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, clientX: 500, clientY: 350 }));
        draggedChart?.dispatchEvent(new PointerEvent("pointermove", { bubbles: true, clientX: 544, clientY: 375 }));
        draggedChart?.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, clientX: 544, clientY: 375 }));
        await wait(100);
        const resized = parseFloat(document.querySelector(".sheet-chart")?.style.width || "0") > originalWidth;
        rb("↶")?.click(); await wait(100);
        const undoSize = Math.abs(parseFloat(document.querySelector(".sheet-chart")?.style.width || "0") - originalWidth) < .01;
        rb("↷")?.click(); await wait(100);
        draggedChart = document.querySelector(".sheet-chart");
        const beforeCancel = draggedChart?.style.left;
        draggedChart?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0, clientX: 400, clientY: 200 }));
        draggedChart?.dispatchEvent(new PointerEvent("pointermove", { bubbles: true, clientX: 430, clientY: 230 }));
        draggedChart?.dispatchEvent(new PointerEvent("pointercancel", { bubbles: true }));
        const cancelDrag = document.querySelector(".sheet-chart")?.style.left === beforeCancel;
        draggedChart = document.querySelector(".sheet-chart");
        const beforeKey = parseFloat(draggedChart?.style.left || "0");
        draggedChart?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", shiftKey: true, bubbles: true }));
        const keyboardMove = Math.abs(parseFloat(document.querySelector(".sheet-chart")?.style.left || "0") - beforeKey - 10) < .01;
        const beforeKeyWidth = parseFloat(document.querySelector(".sheet-chart")?.style.width || "0");
        document.querySelector(".sheet-chart .chart-resize")?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
        const keyboardResize = Math.abs(parseFloat(document.querySelector(".sheet-chart")?.style.width || "0") - beforeKeyWidth - 1) < .01;
        const sheetChartGeometry = { moved, resized, undoSize, cancelDrag, keyboardMove, keyboardResize };
        const sheetChart = { charts: document.querySelectorAll(".sheet-wrap .sheet-chart").length, svg: !!document.querySelector(".sheet-wrap .sheet-chart svg") };
        // Zoom controls ride the ribbon.
        const zin = [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "⊕");
        zin?.click();
        zin?.click();
        await wait(150);
        const zoomLabel = document.querySelector("#ribbon .grp button[title='恢复 100%']")?.textContent;
        const sheetZoom = { zoomLabel };
        // Persist the fill, the picture and the zoom state, so the saved file can
        // be checked by an independent reader (openpyxl) after the run.
        fire("solo:save");
        await wait(600);
        fire("solo:export-pdf");
        await wait(700);

        fire("solo:new", { kind: "docx" });
        await wait(150);
        const blank = document.querySelector("#view .page");
        blank?.focus();
        document.execCommand("insertText", false, "close-guard");
        await wait(120);
        const docNewPart = { paras: blank?.children.length, editable: blank?.getAttribute("contenteditable"), title: document.getElementById("title")?.textContent, dirty: (document.getElementById("title")?.textContent ?? "").includes("●") };

        // A file that is not a package at all: the screen must fall back to the
        // start state, and its 新建 buttons must still work from there.
        fire("solo:load", { name: "坏文件.docx", data: new Uint8Array([1, 2, 3, 4]) });
        await wait(200);
        const errorState = { heading: document.querySelector("#view .empty h1")?.textContent, title: document.getElementById("title")?.textContent, news: document.querySelectorAll("#view [data-new]").length };
        const newXlsx = document.querySelector('#view [data-new="xlsx"]');
        newXlsx?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
        await wait(200);
        await typeCell("9");
        const recovered = { cells: document.querySelectorAll(".sheet-wrap td").length, title: document.getElementById("title")?.textContent };
        // The run ends on a modified document, so quitting goes through the unsaved prompt.
        // A PDF the export asked for shows up in saveAsWrites, next to the pptx/xlsx saves.
        return { docPart, docInsert, docOutline, pptPart, pptInsert, pptShow, pptSaved, sheetPart, sheetFill, sheetPic, sheetCF, sheetGT, sheetNote, sheetChart, sheetChartEdit, sheetChartGeometry, sheetZoom, docNewPart, errorState, recovered };
      })()`;
      void last?.webContents
        .executeJavaScript(probe)
        .then((r) => console.log("SOLO_SMOKE", JSON.stringify({ ...r, inPlaceWrites: smokeWrites, saveAsWrites: smokeSaves })))
        .catch((e: unknown) => console.log("SOLO_SMOKE_ERROR", String(e)))
        .finally(() => app.quit());
    }, 1500);
  });
}

void app.whenReady().then(() => {
  menu();
  createWindow();
  if (process.env.SOLO_SMOKE) smoke();
  // `electron .` carries the app directory as argv[1]; the packaged binary does not.
  const first = process.argv.slice(app.isPackaged ? 1 : 2).find((a) => !a.startsWith("-"));
  if (first) {
    readFile(first)
      .then((buf) => last?.webContents.once("did-finish-load", () => deliver(first, new Uint8Array(buf), basename(first))))
      .catch(() => undefined);
  }
  app.on("open-file", (_e, path) => {
    readFile(path)
      .then((buf) => deliver(path, new Uint8Array(buf), basename(path)))
      .catch(() => undefined);
  });
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
