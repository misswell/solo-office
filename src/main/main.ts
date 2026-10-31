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
  // 探针抛出来的那一条落在渲染进程里，`executeJavaScript` 只会回一句「去看看渲染控制台」——
  // 无人值守的烟测里那句话等于没有，所以把渲染侧的 error/warning 收进数组，失败时跟着打出来。
  const smokeConsole: string[] = [];
  last?.webContents.on("console-message", (_e, level, message, line, source) => {
    if (level >= 2) smokeConsole.push(`w${level} ${source}:${line} ${message}`);
  });
  last?.webContents.once("did-finish-load", () => {
    setTimeout(() => {
      // 探针整段是一枚模板字符串：里面的反引号会终止字符串（tsc 报 TS1005/TS1443），反斜杠会被字符串先吃掉一层
      // （`/gradient\(/` 出去就成了 `/gradient(/`，渲染侧只报一句 Invalid regular expression，烟测只剩 SOLO_SMOKE_ERROR）。
      // 所以这一段既不写反引号，正则也照字面写、不带转义。
      const probe = `(async () => {
        const wait = (ms) => new Promise((r) => setTimeout(r, ms));
        const fire = (name, detail) => window.dispatchEvent(new CustomEvent(name, { detail }));
        const view = document.querySelector("#view");
        const title = document.getElementById("title")?.textContent;
        const page = view?.querySelector(".page");
        if (!page) return { title, view: view?.firstElementChild?.className, heading: view?.querySelector("h1, .para")?.textContent?.slice(0, 40) };
        // #211/#212：带子里那一段说自己坐在哪（w:jc、w:framePr，再是横向缩进 w:ind）——这一腿一枚也不按，也不打字，
        // 所以在最开头那一跳（选区+插字）之前读：读的是壳刚画完、还没被本烟测牵动过的那张纸。第二趟开的是第一趟写回的那份，
        // 于是「两趟读数一字不差 + 盘上 footer 字节一字未改」就是装机壳自己把落点画出来、又原样带回的唯一证据。
        // 件由 .scratch/make-band-place-smoke.ts 现拼（照 tests/helpers/make-band-place.ts 的同一份形状，带子 11 段：
        // 前 8 段是 #211 的居中与框，后 3 段是 #212 的缩进——自己说的 right、链上借来的首行、自己那句 0 顶掉链上那一枚）。
        // 问三样：① 壳把那句落点写没写进行内样式（style 属性字面，与缩放无关）；② 写了之后盒子与那一行墨在带子
        // 内容盒里落到哪（框摆右的那一枚必须贴右沿，锚在纸上那一枚照旧走常规流；右缩进收的是盒子、首行缩进动的只是墨）；
        // ③ 正文上沿与纸缝数——这两格不动才谈得上「落点只上身在这一行，没牵动铺排」。量的全是距离与属性，
        // 一个字的文档内容也不出页面。
        // 几何按纸的单位算：rect 会跟着 zoom 放大，getComputedStyle 的 padding 却不放大（DOM 门禁那一腿踩过同一坑）。
        const zPaper = (el) => { const z = parseFloat(getComputedStyle(el).zoom); return z > 0 ? z : 1; };
        const round1 = (n) => Math.round(n * 10) / 10;
        const gapRead = (el, boxL, boxR, bandTop) => {
          const cs = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          const rg = document.createRange();
          rg.selectNodeContents(el);
          const lines = rg.getClientRects();
          // 墨的左右沿只在 Range 交回来的那些行矩形里取，不许拿外框当起点：外框是整栏宽，一当起点，
          // 居中与靠右的那几行会被自己的盒子吃掉，读回来永远是 0/0（这一趟第一次就踩了它）。
          let inkL = -1, inkR = -1;
          for (let q = 0; q < lines.length; q++) {
            if (lines[q].width < 0.5) continue;
            if (inkL < 0 || lines[q].left < inkL) inkL = lines[q].left;
            if (inkR < 0 || lines[q].right > inkR) inkR = lines[q].right;
          }
          if (inkL < 0) { inkL = r.left; inkR = r.right; }
          return { css: el.getAttribute("style") ?? "", align: cs.textAlign, pos: cs.position, boxW: round1(r.width / zz), gL: round1(r.left / zz - boxL), gR: round1(boxR - r.right / zz), iL: round1(inkL / zz - boxL), iR: round1(boxR - inkR / zz), field: el.querySelectorAll("[data-field]").length };
        };
        let bandLeg = null;
        const footBand = page.querySelector(".hf-footer");
        const zz = zPaper(page);
        if (footBand) {
          const bCss = getComputedStyle(footBand);
          const bRect = footBand.getBoundingClientRect();
          const boxL = bRect.left / zz + parseFloat(bCss.paddingLeft);
          const boxR = bRect.right / zz - parseFloat(bCss.paddingRight);
          // 正文里第一个进流的块：两只带子（纸上的 .hf 与藏起来的复件）都带 data-hf-part，是绝对摆的，不算正文；
          // STYLE/SCRIPT 也不占流。谁排在最前就问谁——本烟测后面那一腿会在正文前插一张表，按「第一段」量的第二趟
          // 读到的是表后面那一格（96 变 112 踩的就是这个，不是落点把正文推下去了）。
          const firstFlow = [...page.children].find((e) => !e.hasAttribute("data-hf-part") && e.tagName !== "STYLE" && e.tagName !== "SCRIPT");
          bandLeg = {
            paras: [...footBand.children].filter((e) => e.classList.contains("hf-p")).map((e) => gapRead(e, boxL, boxR, bRect.top / zz)),
            padL: round1(parseFloat(bCss.paddingLeft)), padR: round1(parseFloat(bCss.paddingRight)), boxW: round1(boxR - boxL), zoom: zz,
            seams: document.querySelectorAll("#view .page").length,
            // 正文第一块离**纸顶**多远（纸的单位）：件里 w:pgMar@w:top=1440 缇 = 96px。不减 padding——减了就永远得 0。
            flowTop: firstFlow ? round1(firstFlow.getBoundingClientRect().top / zz - page.getBoundingClientRect().top / zz) : null,
            flowKind: firstFlow ? String(firstFlow.className).split(" ")[0] : "",
            kids: [...page.children].map((e) => e.tagName + "." + String(e.className).split(" ")[0]).join("|"),
          };
        }
        // 编辑标记（WPS m_ShowHideEditMark 的第一条子项「段落标记」）：视图档，文件里没有这一位。
        // 问三样：① 点一下 ::after 的字面从 none 变 ¶、按钮点亮；② 页数与段/表的 rect 一个都不动（零宽墨）；
        // ③ 再点回去字面回到 none、脏标仍是 0（不进撤销栈、不改字节）。这一腿不碰正文，位置在打字那一腿之前。
        let editMarksLeg = null;
        {
          const emBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "¶");
          const emPara = view.querySelector(".page .para");
          const emTable = view.querySelector(".page table.doc-table");
          const emRead = () => {
            const pr = emPara ? emPara.getBoundingClientRect() : null;
            const tr = emTable ? emTable.getBoundingClientRect() : null;
            return {
              pages: view.querySelectorAll(".page").length,
              para: pr ? [round1(pr.width), round1(pr.height), round1(pr.top)] : null,
              table: tr ? [round1(tr.width), round1(tr.height), round1(tr.top)] : null,
              ink: emPara ? getComputedStyle(emPara, "::after").content : "",
              dirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0,
              lit: emBtn()?.classList.contains("on") ? 1 : 0,
              cls: view.classList.contains("edit-marks") ? 1 : 0,
            };
          };
          const emBefore = emRead();
          emBtn()?.click();
          await wait(240);
          const emOn = emRead();
          emBtn()?.click();
          await wait(240);
          const emOff = emRead();
          editMarksLeg = {
            present: emBtn() ? 1 : 0,
            paras: view.querySelectorAll(".page .para").length,
            before: emBefore,
            on: emOn,
            off: emOff,
            sameBox: JSON.stringify([emBefore.pages, emBefore.para, emBefore.table]) === JSON.stringify([emOn.pages, emOn.para, emOn.table]),
            ink: emBefore.ink === "none" && emOn.ink.indexOf("¶") >= 0 && emOff.ink === "none",
            noDirty: emBefore.dirty === 0 && emOn.dirty === 0 && emOff.dirty === 0,
            falseLegs: [
              emBefore.lit === 0 && emOn.lit === 1 && emOff.lit === 0 ? "" : "lit",
              emBefore.cls === 0 && emOn.cls === 1 && emOff.cls === 0 ? "" : "cls",
            ].filter(Boolean),
          };
        }
        // #235：带子里制表符的落点在装机壳里长什么样。同一只只读腿、同一跳里量：位是 BandPlace.tabStops
        // 合过样式链的那一份（渲染器写进 data-tab-stops），applyTabStops 把每一枚 span.tab 垫到该停的那一格；
        // 字面 \t 装在 span 里，所以每段的 textContent 里制表符还在（只报计数，不出文档内容——bandLeg 的老规矩）。
        // 复件（paintBandCopies 画出来的那份）由 applySheets 末尾那一趟补跑接管，问它与原件同落。
        let bandTabsLeg = null;
        {
          const fb2 = page.querySelector(".hf-footer");
          if (fb2) {
            const z2 = zPaper(page);
            const pCs = getComputedStyle(page);
            const origin2 = page.getBoundingClientRect().left / z2 + parseFloat(pCs.paddingLeft);
            const readSpan = (t) => { const r = t.getBoundingClientRect(); return { w: round1(r.width / z2), land: round1(r.left / z2 - origin2), leader: t.dataset.leader || "" }; };
            const readPara = (e) => ({ attr: e.getAttribute("data-tab-stops") || "", tabs: (e.textContent.match(/\\t/g) || []).length, spans: [...e.querySelectorAll("span.tab")].map(readSpan) });
            const paras2 = [...fb2.children].filter((e) => e.classList.contains("hf-p"));
            const clones2 = [...page.querySelectorAll(".hf-footer")].filter((e) => e !== fb2);
            bandTabsLeg = {
              padL: round1(parseFloat(pCs.paddingLeft)), zoom: z2,
              paras: paras2.map(readPara),
              cloneParas: clones2.length ? [...clones2[0].children].filter((e) => e.classList.contains("hf-p")).map(readPara) : [],
              clones: clones2.length,
            };
          }
        }
        // #213：正文里的表在栏里坐哪（w:jc 与 w:tblInd 上身之后浏览器算完的那一格）。同一条只读腿，
        // 在读带子这一跳之前、壳刚画完那张纸的时候量；件由 .scratch/make-table-place-smoke.ts 现拼，
        // 六张表照授权口径的形状排：过宽居中 / 居中不说缩进 / 居中+288 / 居中−687 / 只说缩进 / 什么都没说。
        // 问的是**矩形**（离正文内容盒左沿多远、表自己宽多少）加行内样式那串字面，不动一个字的文档内容。
        let tablePlaceLeg = null;
        {
          const pCss = getComputedStyle(page);
          const pRect = page.getBoundingClientRect();
          const cL = pRect.left / zz + parseFloat(pCss.paddingLeft || "0");
          const cR = pRect.right / zz - parseFloat(pCss.paddingRight || "0");
          tablePlaceLeg = {
            colW: round1(cR - cL),
            padL: round1(parseFloat(pCss.paddingLeft || "0")),
            tables: [...page.querySelectorAll("table")].map((el) => {
              const r = el.getBoundingClientRect();
              return { css: el.getAttribute("style") ?? "", gL: round1(r.left / zz - cL), gR: round1(cR - r.right / zz), w: round1(r.width / zz) };
            }),
          };
        }
        // #228 标尺：这只装机盒子里尺子与纸同框、视图页那枚开关真有面、开合不动件。只报矩形、计数与
        // 「脏点还在不在」，一个字的文档内容也不出页面。标尺说的每一格都来自纸自己（--page-w 减两条边距、
        // 光标段落量出来的盒子、data-tab-stops），所以这里问的是「尺与纸同框」「正文栏宽与纸写的那两只数对得上」
        // 与「关掉不占位、开回来又占那一格」。件里没有一句说过标尺，开合前后标题必须一字不差。
        let rulerLeg = null;
        {
          const ru = view.querySelector(":scope > .doc-ruler");
          const tit0 = document.getElementById("title")?.textContent ?? "";
          if (ru) {
            const pr = page.getBoundingClientRect();
            const rr = ru.getBoundingClientRect();
            const pcs = getComputedStyle(page);
            const paperW = parseFloat(page.style.getPropertyValue("--page-w")) || 0;
            const padL = parseFloat(pcs.paddingLeft || "0");
            const padR = parseFloat(pcs.paddingRight || "0");
            const rb = [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "标尺");
            rb?.click();
            await wait(120);
            const off = { hidden: ru.hidden, h: round1(ru.getBoundingClientRect().height), pageH: round1(page.getBoundingClientRect().height) };
            rb?.click();
            await wait(120);
            // 标尺上那三枚缩进标记要能拖：把光标放进第一段，看尺子报不报纸面几何与可拖光标。
            // 这里只读数、不按下——落笔那一趟由 DOM 闸门（ruler-dom）在真实 Chrome 里走。
            const firstPara = view.querySelector(".page .para");
            if (firstPara) {
              const r0 = document.createRange();
              r0.setStart(firstPara.firstChild ?? firstPara, 0);
              r0.collapse(true);
              const s0 = document.getSelection();
              s0?.removeAllRanges();
              s0?.addRange(r0);
              await wait(120);
            }
            const ind = ru.querySelector('[data-ruler-mark="left"]');
            rulerLeg = {
              sibling: ru.nextElementSibling === page,
              one: view.querySelectorAll(".doc-ruler").length,
              sameBox: Math.abs(rr.width - pr.width) < 0.8 && Math.abs(rr.left - pr.left) < 0.8,
              pw: round1(pr.width),
              rw: round1(rr.width),
              pl: round1(pr.left),
              rl: round1(rr.left),
              sl: round1(view.scrollLeft),
              cw: round1(view.clientWidth),
              poff: round1(page.offsetLeft),
              roff: round1(ru.offsetLeft),
              pml: pcs.marginLeft,
              pzoom: getComputedStyle(page).zoom,
              rzoom: getComputedStyle(ru).zoom,
              text: round1(parseFloat(ru.dataset.rulerText || "0")),
              textWant: round1(paperW - padL - padR),
              majors: ru.querySelectorAll(".ruler-tick.major").length,
              unit: (ru.querySelector(".ruler-unit")?.textContent ?? "").length > 0,
              sticky: getComputedStyle(ru).position,
              hasButton: !!rb,
              drag: ind ? { cursor: getComputedStyle(ind).cursor, geom: !!ru.dataset.dragGeom, marks: ru.querySelectorAll("[data-ruler-mark]").length } : null,
              off,
              back: { hidden: ru.hidden, h: round1(ru.getBoundingClientRect().height) },
              dirtyKept: (document.getElementById("title")?.textContent ?? "") === tit0,
            };
          } else rulerLeg = { missing: true };
        }
        // #230 开始页「格式刷」：这只装机盒子里按钮真有面、按下去举起来（按钮亮、纸上换指针）、Esc 放下，
        // 两下都不碰件。刷上那一步由 DOM 闸门（painter-dom）在真实 Chrome 里走。
        let painterLeg = null;
        {
          const pb = [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "格式刷");
          const titP = document.getElementById("title")?.textContent ?? "";
          const fp = view.querySelector(".page .para");
          if (fp) {
            const rp = document.createRange();
            rp.setStart(fp.firstChild ?? fp, 0);
            rp.collapse(true);
            const sp = document.getSelection();
            sp?.removeAllRanges();
            sp?.addRange(rp);
            await wait(80);
          }
          pb?.click();
          await wait(120);
          const armed = {
            buttonOn: pb ? pb.classList.contains("on") : false,
            painted: page ? page.classList.contains("fmt-painter") : false,
            cursor: page ? getComputedStyle(page).cursor : "",
          };
          document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
          await wait(120);
          const released = {
            buttonOn: pb ? pb.classList.contains("on") : false,
            painted: page ? page.classList.contains("fmt-painter") : false,
          };
          painterLeg = { hasButton: !!pb, armed, released, dirtyKept: (document.getElementById("title")?.textContent ?? "") === titP };
        }
        // #234 表行真跨页：件「行内切纸-装机.docx」里第一张表唯一一行允许断（两格各 150 段字，铺出来把这一行
        // 切成 5 截：4 道行内刀口、2 格各垫 4 枚），这一腿量的是**刚画完的那一版**：后面的打字/插表/撤销各腿
        // 会改正文（第一次装机时它在末尾跑，表 A 已被改得不可断——页数读到 2、格内一枚垫片也没有）。件是专门的
        // 夹具，认件名就跑，跑完才轮到别的腿。
        // 第二张表唯一一行带 w:cantSplit（两格各 25 段字，一页装得下）。问五件事：① 允许断的那一行在纸上被
        // 切成几截——格里的 .solo-sheet-gap 计数，且表里没有垫片行；② 每一道缝都有条带（条带数 == 格内垫片 +
        // 垫片行）；③ 不许断的那一行整枚落在同一张纸上、格里一枚垫片都没有；④ 垫片是段落的孩子不是格的孩子
        // （读回模型时 td 的孩子数不许变）；⑤ 存盘不脏。行数/文字/那句 w:cantSplit 由文件外的 python 反读。
        const rsDoc = (document.getElementById("title")?.textContent ?? "").includes("行内切纸");
        let rowSplitLeg = null;
        if (rsDoc) {
          // 起手等铺排停稳：文件是 solo:load 交过来的，画完正文才轮到 applySheets 量页，而量页本身分几趟落纸
          // （补垫片、再量一遍）。只等条带出现会读到中间那一版——两格垫片还差 1px、行高差 2px。连读三次同一套
          // 数（纸/条带/垫片/垫片总高）才算停稳。
          let rsSig = "";
          let rsSame = 0;
          for (let tick = 0; tick < 60 && rsSame < 3; tick++) {
            await wait(80);
            const now = [
              document.querySelectorAll("#view .page").length,
              document.querySelectorAll("#view .page .solo-sheet-strip").length,
              document.querySelectorAll("#view .page .solo-sheet-gap").length,
              Math.round([...document.querySelectorAll("#view .page .solo-sheet-gap")].reduce((n, g) => n + g.getBoundingClientRect().height, 0)),
            ].join("/");
            rsSame = now === rsSig ? rsSame + 1 : 0;
            rsSig = now;
          }
          const rsTables = [...document.querySelectorAll("#view .page table.doc-table")];
          const rsA = rsTables.find((t) => (t.textContent ?? "").includes("甲"));
          const rsB = rsTables.find((t) => (t.textContent ?? "").includes("丙"));
          const rsGaps = (el) => (el ? [...el.querySelectorAll(".solo-sheet-gap")] : []);
          const rsKids = (el) => (el ? [...el.querySelectorAll("td")].map((c) => c.children.length) : []);
          const rsRows = (el) =>
            el ? [...el.querySelectorAll("tr")].filter((tr) => !tr.classList.contains("solo-sheet-gap-row") && !tr.classList.contains("solo-repeat-header-row")).length : -1;
          const rsGapRows = (el) => (el ? el.querySelectorAll("tr.solo-sheet-gap-row").length : -1);
          const rsPages = () => [...document.querySelectorAll("#view .page")];
          const gapsA = rsGaps(rsA);
          const gapsB = rsGaps(rsB);
          const rowA = rsA
            ? [...rsA.querySelectorAll("tr")].find((tr) => !tr.classList.contains("solo-sheet-gap-row") && !tr.classList.contains("solo-repeat-header-row"))
            : null;
          const rowB = rsB ? [...rsB.querySelectorAll("tr")].find((tr) => !tr.classList.contains("solo-sheet-gap-row") && !tr.classList.contains("solo-repeat-header-row")) : null;
          const statusCell = [...document.querySelectorAll(".solo-doc-status-loc")].find((b) => (b.title || "").indexOf("页数") === 0);
          // 状态栏那一格照 WPS 写成「页面: 当前/总数」：总数在斜杠后面，拿第一个数字会把当前页当成页数。
          const stTxt = statusCell ? statusCell.textContent || "" : "";
          const stSlash = stTxt.indexOf("/");
          const stTotal = stSlash >= 0 ? Number(stTxt.slice(stSlash + 1)) : -1;
          const strips = document.querySelectorAll("#view .page .solo-sheet-strip").length;
          const spansAll = document.querySelectorAll("#view .page .solo-sheet-gap").length;
          const gapRowsAll = document.querySelectorAll("#view .page tr.solo-sheet-gap-row").length;
          const stripEls = [...document.querySelectorAll("#view .page .solo-sheet-strip")];
          // 每一道缝都得被一段空地推开：垫片落在**下页第一枚内容**之前，它的盒跨着条带（上沿在缝上方、下沿在缝下方）。
          // 拿垫片的上沿去和条带对位是错的——同一道缝的行内那一刀按格各垫一枚，各格的上沿本来就不在缝上。
          const cutEverySeam = stripEls.every((s) => {
            const mid = s.getBoundingClientRect().top + 1;
            return [...document.querySelectorAll("#view .page .solo-sheet-gap, #view .page tr.solo-sheet-gap-row")].some((g) => {
              const r = g.getBoundingClientRect();
              return r.top < mid && r.bottom > mid;
            });
          })
            ? 1
            : 0;
          fire("solo:save");
          await wait(900);
          const rsBox = (el) => {
            const b = el.getBoundingClientRect();
            const p0 = rsPages()[0] ? rsPages()[0].getBoundingClientRect().top : 0;
            return { t: Math.round(b.top - p0), b: Math.round(b.bottom - p0), h: Math.round(b.height) };
          };
          rowSplitLeg = {
            pageEls: rsPages().length, rowA: rowA ? rsBox(rowA) : null, rowB: rowB ? rsBox(rowB) : null,
            gapTops: gapsA.map((g) => Math.round(g.getBoundingClientRect().top - (rsPages()[0] ? rsPages()[0].getBoundingClientRect().top : 0))),
            gapHs: gapsA.map((g) => Math.round(g.getBoundingClientRect().height)),
            hasA: rsA ? 1 : 0, hasB: rsB ? 1 : 0,
            rowsA: rsRows(rsA), rowsB: rsRows(rsB),
            gapsA: gapsA.length, gapRowsA: rsGapRows(rsA), gapsB: gapsB.length, gapRowsB: rsGapRows(rsB),
            inParaA: gapsA.filter((g) => g.parentElement && g.parentElement.classList.contains("para")).length,
            inCellA: gapsA.filter((g) => g.parentElement && (g.parentElement.tagName === "TD" || g.parentElement.tagName === "TH")).length,
            kidsA: rsKids(rsA), kidsB: rsKids(rsB),
            wholeB: rowB
              ? stripEls.every((s) => {
                  const r = s.getBoundingClientRect();
                  const b = rowB.getBoundingClientRect();
                  return !(r.top > b.top + 1 && r.top < b.bottom - 1);
                })
                ? 1
                : 0
              : -1,
            pages: stTotal, strips: strips, spansAll: spansAll, gapRowsAll: gapRowsAll,
            cutEverySeam: cutEverySeam,
            savedDirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0,
          };
        }
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
        // Symbols whose computed size is not the size their own paragraph has: the level states it for its mark
        // (17.3.1.12), and when the level says nothing the paragraph mark's resolved size steps in (17.3.1.4).
        // A count only, never a number that could name a document's typography.
        const bulletSized = [...document.querySelectorAll("#view .page .bullet")].filter(
          (s) => getComputedStyle(s).fontSize !== getComputedStyle(s.parentElement ?? document.body).fontSize,
        ).length;
        // Symbols wearing an ink of their own rather than the page stylesheet's: the paragraph mark states a colour
        // for the break (17.3.1.4) and these files write that same one on the runs, so before this round only the
        // text wore it while the symbol stayed in the page stylesheet's --ink. A count only, never a value that
        // could name a document's colour.
        const bulletInked = [...document.querySelectorAll("#view .page .bullet")].filter(
          (s) => getComputedStyle(s).color !== getComputedStyle(s.parentElement ?? document.body).color,
        ).length;
        // 字体框这条腿在真实壳里走一遍 #140 修的那条原路：焦点从正文移进这一格时，Blink 交回来的选区被折成页外的
        // 一个点，命令从此静默落不下去。DOM 门禁跑的是 src/ui/index.html 生成的一份副本，所以这一腿要证明的是 shipping
        // 的这只盒子。挑的名字是探针自己编的：文档里的字体名照旧不出页面，这里只报「穿上这张脸的格子数」和回显对不对得上。
        const faceName = "SoloSmokeFace";
        const smokeText = () => {
          const root = document.querySelector("#view .page");
          if (!root) return null;
          const walk = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
          let n;
          while ((n = walk.nextNode())) if (n.textContent?.includes("SMOKE")) return n;
          return null;
        };
        const worn = () => [...document.querySelectorAll("#view .page span")].filter((s) => s.textContent?.includes("SMOKE") && getComputedStyle(s).fontFamily.includes(faceName)).length;
        const faceBox = document.querySelector("#ribbon input.facebox");
        const faceNode = smokeText();
        if (faceNode && faceBox) {
          const fr = document.createRange();
          fr.selectNodeContents(faceNode);
          const fs = document.getSelection();
          fs?.removeAllRanges();
          if (fs) fs.addRange(fr);
          document.dispatchEvent(new Event("selectionchange"));
          await wait(40);
          faceBox.focus();
          faceBox.value = faceName;
          faceBox.dispatchEvent(new Event("change", { bubbles: true }));
          await wait(150);
          // 回显走的是光标那一段：先把手从这一格上拿开（这一格被聚焦时不回显，那是给用户还没打完的名字留的），
          // 再把光标放进 SMOKE 那几个字里，让 syncDocRibbon 从模型读那一段现在穿的脸。
          faceBox.blur();
          const afterNode = smokeText();
          if (afterNode) {
            const cr = document.createRange();
            cr.setStart(afterNode, 0);
            cr.collapse(true);
            const cs = document.getSelection();
            cs?.removeAllRanges();
            if (cs) cs.addRange(cr);
            document.dispatchEvent(new Event("selectionchange"));
            await wait(40);
          }
        }
        const fontBoxSpans = worn();
        const fontBoxEcho = faceBox && faceBox.value === faceName ? 1 : 0;
        // 布局那一半：WPS 把字体框与字号框挂在同一个 font_info_pair 里（字体在前，minimumWidth=130），所以这一对
        // 必须同顶同高。量的是一只盒子的两条边，报的是「对不对齐」和宽度，不报文档里的任何东西。
        // 这一腿会咬：把 applyStyle 开头那条「读不出正文位置就先把选区放回页面」的恢复拿掉，真实壳里回的是
        // fontBoxSpans=0、fontBoxEcho=0（用户看到的就是那次彻底的静默），装回来才是 1、1。
        const sizeBox = document.querySelector("#ribbon input.sizebox");
        const faceRect = faceBox?.getBoundingClientRect();
        const sizeRect = sizeBox?.getBoundingClientRect();
        const ribbonPair = faceRect && sizeRect && Math.abs(faceRect.top - sizeRect.top) < 0.5 && Math.abs(faceRect.height - sizeRect.height) < 0.5 && Math.round(faceRect.width) === 130 ? 1 : 0;
        marker?.click();
        await wait(150);
        const after = document.querySelector("#view .page");
        const docPart = { title: document.getElementById("title")?.textContent, editable: after?.getAttribute("contenteditable"), blocksBefore: blocks, blocksAfter: after?.children.length, typed: after?.textContent?.includes("SMOKE"), bulletsBefore: bullets, bulletsWearingOwnFont: bulletWorn, bulletsOwnSize: bulletSized, bulletsOwnColor: bulletInked, bulletsAfter: after?.querySelectorAll(".bullet").length, ribbonGroups: document.querySelectorAll("#ribbon .grp").length, fontBoxSpans, fontBoxEcho, ribbonPair, homeCaps: [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] .grp")].map((g) => g.dataset.cap || "-").join("/"), tableCapGroups: document.querySelectorAll("#ribbon .rbar[data-tab='table'] .grp[data-cap]").length, homeIcons: document.querySelectorAll("#ribbon .rbar[data-tab='home'] button.rb.ico .ic svg path").length, homeIconText: [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button.rb.ico")].every((b) => b.textContent === b.querySelector(".lb")?.textContent), titleCenter: (function () { const h = document.querySelector("body > header").getBoundingClientRect(); const t = document.getElementById("title").getBoundingClientRect(); return Math.round(Math.abs((t.left + t.right) / 2 - (h.left + h.right) / 2) * 10) / 10; })(), qat: [...document.querySelectorAll("#qat button")].map((b) => b.id).join("/"), qatIcons: document.querySelectorAll("#qat button svg path").length };

        // #210：w:u 那一族在装机那只盒子里画成什么样——这一腿一枚也不按，所以别的装机件字节账一字不动。
        // （探针整段是一枚模板字符串，这一段里既不写反引号也不写反斜杠。）
        // 件由 .scratch/make-underline-smoke.ts 现拼（照 tests/helpers/make-run.ts 的同一份形状）。问四件事：
        // 说了不要的画不画、借来的线画不画（且不许被 run 自己认领）、单线与粗线各多粗、什么都没说的那枚别发明线。
        const uNeedle = (t) => [...document.querySelectorAll("#view .page span")].find((s) => (s.textContent ?? "").includes(t));
        const uRead = (t) => {
          const el = uNeedle(t);
          if (!el) return null;
          const s = getComputedStyle(el);
          return { line: s.textDecorationLine, thick: s.textDecorationThickness, size: s.fontSize, stamp: el.getAttribute("data-u") ?? "" };
        };
        const underLeg = { none: uRead("说明没有"), covers: uRead("这句话盖掉借来的线"), lent: uRead("链上借来的线"), single: uRead("自己点了线"), thick: uRead("粗线原样回去"), silent: uRead("什么都没说") };

        // #144：段落那一圈 w:pBdr 与那一层 w:shd 在**真实壳**里画不画、占不占流里的地方。DOM 门禁跑的是
        // src/ui/index.html 生成的一份副本，这一腿要证明装机那只盒子也一样。量的都是「画了几段」「高差几个 px」
        // 「页缝数」，不报文档里的颜色名与字体名。件由 .scratch/make-pbdr-smoke.ts 现拼，两件只差样式 SoloLever
        // 里那一圈（8px 线 + 80px 间距），所以两份对照才读得出那一圈进没进铺排：seams 该是 1（无）对 5（有）。
        await wait(200);
        const boxedParas = [...document.querySelectorAll("#view .page .para")];
        const byNeedle = (needle) => boxedParas.find((e) => (e.textContent ?? "").includes(needle));
        const cssOf = (el, prop) => (el ? getComputedStyle(el)[prop] : "");
        const solidOn = (el) => ["Top", "Left", "Bottom", "Right"].some((k) => cssOf(el, "border" + k + "Style") === "solid" && parseFloat(cssOf(el, "border" + k + "Width")) > 0);
        const paintedOn = (el) => {
          const c = cssOf(el, "backgroundColor");
          return c !== "" && c !== "rgba(0, 0, 0, 0)" && c !== "transparent";
        };
        const heightOf = (needle) => {
          const el = byNeedle(needle);
          return el ? el.getBoundingClientRect().height : -1;
        };
        const paraBox = {
          // 件里 44 段，这一时刻纸上是 45 段。两份的 lined 差 41＝40 段 lever＋SOLOBOX：多出来那一段自己也画了线，
          // 说明它穿的是正文最后那段那一圈（只在带框那件显形，没带框那件那一圈是空的）。
          paras: boxedParas.length,
          lined: boxedParas.filter(solidOn).length,
          grounded: boxedParas.filter(paintedOn).length,
          // 样式那条 1px 底线与 1 磅间距（sz=6、space=1）：高差只有画了才是 2.333px。
          boxLeg: Math.abs(heightOf("SOLOBOX") - heightOf("SOLOPLAIN") - 2.333) < 0.05 ? 1 : 0,
          gapPx: Math.abs(parseFloat(cssOf(byNeedle("SOLOBOX"), "paddingBottom")) - 1.333) < 0.02 ? 1 : 0,
          // none 是取消不是空白：退掉那条线的那一段没线，但样式那层地照旧；四边全 none 的既不画也不铺。
          refusedLine: solidOn(byNeedle("SOLOREFUSE")) ? 0 : 1,
          refusedGround: paintedOn(byNeedle("SOLOREFUSE")) ? 1 : 0,
          allNoneClean: !solidOn(byNeedle("SOLOALLNONE")) && !paintedOn(byNeedle("SOLOALLNONE")) ? 1 : 0,
          // 那 40 段各带一圈 8px 线 + 80px 间距：与四边全取消那一段比高，写那一圈的件是 88px，没写的件是 0。
          leverMinusNone: Math.round((heightOf("第1行的话") - heightOf("SOLOALLNONE")) * 1000) / 1000,
          // 页缝垫片数：画的纸是按量出来的铺排切的，那一圈进没进铺排在这里见分晓。
          seams: document.querySelectorAll("#view .page .solo-sheet-gap").length,
        };

        // #145：ribbon 那条腿在**装机这只盒子**里落不落笔，落完按壳自己的原位保存路径写回原文件，再跑一遍读不读得回来。
        // 同一份探针跑两趟比的是同一个数：第一趟 cmdBefore 0（那一段四边没线）→ 点完 cmdAfter 1；第二趟（不再重拼件，
        // 直接开上一趟写回的那份）cmdBefore 就该是 1——那是 shipped 的壳把 shipped 的壳写下的字节读了回来。
        // #153 起这一腿还问中间那一条（w:between）：SOLOPLAIN 那段量的是**不画**（下一段样式不同，没有「之间」那句话），
        // 40 段 SoloLever 的第一段量的是**画**（下一段同样式，带子 1px、四边仍 none）。第二趟这两条都该仍是 1：
        // 装机的壳把自己写下的 w:between 读回来、又按同一条规则画了回去。
        // 第二趟只读这几条：cmdGapped 量的是那一次点击带来的**增量**，已经有框的段落再点一次不该垫第二层，读 0 才对；
        // paraBox 里量样式那圈高差的几条（boxLeg、leverMinusNone、seams）也被这一趟自己加上去的框带偏，是刻意被污染的，
        // 要读那几条就重拼件跑第一趟。
        // 落笔与保存都走生产路径（那一格的 change → setParagraphBox；solo:save → writeOut → bridge.writeFile），探针不碰模型。
        // 每次量之前重新按文字找那一段：改格式会让纸上的段落重建，抓在手里的旧元素会变成流外的一份。
        const widthsOf = (el) => ["Top", "Left", "Bottom", "Right"].map((k) => cssOf(el, "border" + k + "Width")).join("/");
        const padOf = (el, k) => (el ? parseFloat(cssOf(el, "padding" + k)) : -1);
        const findPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((e) => (e.textContent ?? "").includes(needle));
        const boxSel = () => [...document.querySelectorAll("#ribbon select")].find((s) => s.title === "边框");
        const cmd0 = findPara("SOLOPLAIN");
        const cmdRange = document.createRange();
        cmdRange.selectNodeContents(cmd0 ?? page);
        cmdRange.collapse(false);
        const cmdPoint = document.getSelection();
        cmdPoint?.removeAllRanges();
        if (cmdPoint) cmdPoint.addRange(cmdRange);
        cmd0?.focus?.();
        await wait(150);
        const cmdBefore = solidOn(findPara("SOLOPLAIN")) ? 1 : 0;
        // 间距只比增量：那四格不是这一族软件独有的（页面样式给段落自己那点呼吸），绝对值说不清。
        const padBase = ["Top", "Left", "Bottom", "Right"].map((k) => padOf(findPara("SOLOPLAIN"), k));
        const cmdLeg = (() => {
          const s = boxSel();
          if (!s) return 0;
          s.value = "all";
          s.dispatchEvent(new Event("change", { bubbles: true }));
          return 1;
        })();
        await wait(280);
        const cmd1 = findPara("SOLOPLAIN");
        const cmdAfter = solidOn(cmd1) ? 1 : 0;
        const cmdWidths = widthsOf(cmd1);
        const cmdFourSides = cmdWidths === "1px/1px/1px/1px" ? 1 : 0;
        const cmdGapped = ["Top", "Left", "Bottom", "Right"].every((k, i) => Math.abs(padOf(cmd1, k) - padBase[i] - 1.333) < 0.02) ? 1 : 0;
        const cmdLined = [...document.querySelectorAll("#view .page .para")].filter(solidOn).length;
        const cmdDirty = (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0;
        // #153 的 w:between：那一条说的是**这一段与下一段之间**，所以「下一段不是同一个样式」时不该出现。
        // SOLOPLAIN 紧跟着的是 SOLOREFUSE（样式 SoloBox），样式不同 ⇒ 没有「之间」那句话 → 反向对照。
        // 名字读的是那一句「中间没画」：背景里没有那条带子才记 1。
        const cmdNoMiddle = /gradient/.test(cssOf(cmd1, "backgroundImage")) ? 0 : 1;
        // 正向对照落在 40 段 SoloLever 那一族的第一段上：紧跟着的第 2 行是同一个样式，那一条才画得出。
        // 它四边本来一条也没说，所以「中间那条」在这里是唯一的一条线——背景带而非 border，四边仍该是 none。
        const midSpanOf = (needle) => findPara(needle)?.querySelector("span")?.firstChild ?? null;
        const midSpan = midSpanOf("第1行的话");
        if (midSpan) {
          const mr = document.createRange();
          mr.setStart(midSpan, 1);
          mr.setEnd(midSpan, 1);
          const ms = document.getSelection();
          ms?.removeAllRanges();
          if (ms) ms.addRange(mr);
        }
        await wait(150);
        const midBase = padOf(findPara("第1行的话"), "Bottom");
        const midLeg = (() => {
          const s = boxSel();
          if (!s) return 0;
          s.value = "insideH";
          s.dispatchEvent(new Event("change", { bubbles: true }));
          return 1;
        })();
        await wait(280);
        const mid1 = findPara("第1行的话");
        const cmdMiddle = /gradient/.test(cssOf(mid1, "backgroundImage")) ? 1 : 0;
        const cmdMiddleBare = ["Top", "Left", "Bottom", "Right"].every((k) => cssOf(mid1, "border" + k + "Style") === "none") ? 1 : 0;
        // 间距 1 磅（1.333px）加那一条线自己占的 1px：padding-bottom 只有一条，取底边与中间那一条里大的那一个。
        const cmdMiddleGap = Math.abs(padOf(mid1, "Bottom") - midBase - 2.333) < 0.02 ? 1 : 0;
        const cmdMiddleW = cssOf(mid1, "backgroundSize");
        fire("solo:save");
        await wait(800);
        const paraBoxCmd = { cmdLeg, cmdBefore, cmdAfter, cmdFourSides, cmdGapped, cmdLined, cmdDirty, cmdWidths, cmdNoMiddle, midLeg, cmdMiddle, cmdMiddleBare, cmdMiddleGap, cmdMiddleW };

        // Insertion round: a 2×3 table, a page break and a picture land in the body.
        const rb = (label) => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === label);
        rb("表格")?.click();
        await wait(120);
        const pick = document.querySelector(".gridpick button[data-c='3'][data-r='2']");
        pick?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
        await wait(200);
        // #158：纸缝落在表行之间时画的是垫片**行**，它与正文行共用 tr 这个名。这两读的是模型里的表（2×3），
        // 不滤就缝每多一道行数虚增一枚、列数还可能读成垫片那一格。与 doc-edit.ts 的 rowsOf 同一口径。
        const tableRows = document.querySelectorAll("#view .page table.doc-table tr:not(.solo-sheet-gap-row)").length;
        const tableCols = document.querySelector("#view .page table.doc-table tr:not(.solo-sheet-gap-row)")?.children.length;

        // #146：同一枚「边框」格在**装机这只盒子**里按选区派发的那两层。件由 .scratch/make-tbl-smoke.ts 现拼：
        // 一张 2×3 表，表与格在件里一句框线都没说（表那一圈没有 w:tblBorders，格也没有 w:tcBorders），落笔走的是
        // 写出按字段拼那一枚孩子那一条路；第二趟不重拼件、直接开上一趟按壳自己的路径写回的那份，这时同一张表已经
        // 自带那一枚孩子，另一条路一起在壳里量到。拖选的两个落点是件里自己的字（甲与末），不是探针替模型手写的选区。
        // 读的五件事：列子换没换到表那一层的十一档、落点画没画在 table 元素那一圈、那一击之后选区还在不在表这一层
        // （本轮修的正是这条：缩回一格的话第二击就换了层）、同一条腿第二次落笔是不是空过（按历史问：撤销一步就退回
        // 落笔前那圈没线的样子，说明那一次没进历史——这颗点在这件上量不出空过，上面刚插了一张表，整篇从一开始就挂着），
        // 以及光标退回一格时列子是不是回到十一档。件里那句 w:tblBorders 由 .scratch/tbl-cmd-install-check.sh 在外面问。
        // #148 又加一件：撤销之后**页面上此刻那条选区**自己还铺不铺在甲格与末格之间（不是探针替它重选的），
        // 因为落点那一层是它的下游——历史只记一个点的话这里会塌成一格，第二击问的就是别的层。
        // #225 那一件自带三层待判的属性记录；#146/#149 按「甲」「末」找表落笔、按窗格第一行判定，
        // 会把这张表的记录接受掉、底纹改掉（第一趟实测）。这一件只让 #225 自己读，前面那几腿一律不落笔。
        const otEarly = (document.getElementById("title")?.textContent ?? "").includes("原文表层");
        const ringOf = (needle) => otEarly ? undefined : [...document.querySelectorAll("#view .page table.doc-table")].find((t) => (t.textContent ?? "").includes(needle));
        const cellOf = (needle) => [...(ringOf(needle)?.querySelectorAll("td") ?? [])].find((c) => (c.textContent ?? "").includes(needle));
        const spanOf = (needle) => cellOf(needle)?.querySelector("span")?.firstChild ?? null;
        const ringBottom = (el) => (el ? cssOf(el, "borderBottomStyle") + "/" + cssOf(el, "borderBottomWidth") : "");
        const boxPick = () => [...document.querySelectorAll("#ribbon select")].find((s) => s.dataset && s.dataset.scope !== undefined);
        const legCount = () => {
          const s = boxPick();
          return s ? [...s.options].length : -1;
        };
        const clickLeg = (value) => {
          if (otEarly) return 0;
          const s = boxPick();
          if (!s) return 0;
          s.value = value;
          s.dispatchEvent(new Event("change", { bubbles: true }));
          return 1;
        };
        const dirtyNow = () => ((document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0);
        const tblRing0 = ringOf("甲");
        const tblBefore = ringBottom(tblRing0);
        const tblBeforeSolid = tblBefore.startsWith("solid/") ? 1 : 0;
        const head = spanOf("甲");
        const tail = spanOf("末");
        const tblSelected = head && tail ? 1 : 0;
        if (head && tail) {
          const dr = document.createRange();
          dr.setStart(head, 0);
          dr.setEnd(tail, (tail.textContent ?? "").length);
          const ds = document.getSelection();
          ds?.removeAllRanges();
          if (ds) ds.addRange(dr);
        }
        await wait(240);
        const tblScope = boxPick()?.dataset?.scope ?? "";
        const tblLegs = legCount();
        const tblLegNames = boxPick() ? [...boxPick().options].map((o) => o.textContent).join("|") : "";
        const tblTitle = document.getElementById("title")?.textContent ?? "";
        const tblClicked = clickLeg("bottom");
        await wait(340);
        const tblAfter = ringBottom(ringOf("甲"));
        const tblAfterSolid = tblAfter.startsWith("solid/") ? 1 : 0;
        const tblScopeKept = (boxPick()?.dataset?.scope ?? "") === "table" ? 1 : 0;
        const tblDirty = dirtyNow();
        // 同一条腿第二次落笔：件里那句话这时已经是这一腿要说的一句话，该空过。空过在壳里怎么验？那颗点量不出来
        // ——它说的是「整篇有没有未保存的改动」，第一击之后本来就一直挂着。换成按历史问：那一次空过不进历史，
        // 所以撤销**一步**就该直接退回件里那句话之前的样子（那一圈没线），重做再画回来。
        const tblReclick = clickLeg("bottom");
        await wait(340);
        const tblReclickSame = ringBottom(ringOf("甲")) === tblAfter ? 1 : 0;
        const ribbonBtn = (label) => [...document.querySelectorAll("#ribbon button")].find((b) => b.textContent === label);
        const tblUndo = ribbonBtn("撤销") ? (ribbonBtn("撤销").click(), 1) : 0;
        await wait(340);
        const tblAfterUndo = ringBottom(ringOf("甲"));
        // #148：历史里那一条记的是两端，撤销该把那条跨格选区整条铺回来。量此刻页面里那条选区自己的两端
        // （各落在哪一格、塌没塌），落点那一层是它的下游——只读下游的话回显与选区谁先谁后就分不出来了。
        const tblScopeAfterUndo = boxPick()?.dataset?.scope ?? "";
        const tblSpanAfterUndo = (() => {
          const r = document.getSelection()?.getRangeAt(0);
          if (!r || r.collapsed) return 0;
          const inCell = (n) => {
            const el = n && n.nodeType === 1 ? n : n?.parentElement;
            return el && typeof el.closest === "function" ? el.closest("td") : null;
          };
          const a = inCell(r.startContainer);
          const z = inCell(r.endContainer);
          return a && z && a !== z && (a.textContent ?? "").includes("甲") && (z.textContent ?? "").includes("末") ? 1 : 0;
        })();
        const tblRedo = ribbonBtn("重做") ? (ribbonBtn("重做").click(), 1) : 0;
        await wait(340);
        const tblAfterRedo = ringBottom(ringOf("甲"));
        // 光标退回一格：那一层说的是格（十一档），落点从 table 那一圈换到那一格自己那一圈的顶缝。
        const soloCell = spanOf("甲");
        if (soloCell) {
          const cr = document.createRange();
          cr.setStart(soloCell, 1);
          cr.setEnd(soloCell, 1);
          const cs = document.getSelection();
          cs?.removeAllRanges();
          if (cs) cs.addRange(cr);
        }
        await wait(240);
        const cellScope = boxPick()?.dataset?.scope ?? "";
        const cellLegs = legCount();
        const cellClicked = clickLeg("top");
        await wait(340);
        const cellTopNow = cellOf("甲") ? cssOf(cellOf("甲"), "borderTopStyle") + "/" + cssOf(cellOf("甲"), "borderTopWidth") : "";
        // #150/#152：ribbon 最后那一档在**格这一层**开的是同一枚「边框与底纹」对话框——接上那两层各自的 w:shd 原位
        // 写回之后，这里连底纹那一行一起摆，确定之后落的还是格自己那一圈。
        const dlgOpened = clickLeg("more");
        await wait(240);
        const dlg = document.querySelector(".solo-pb-dialog");
        const dlgScope = dlg ? (dlg.querySelector(".solo-pb-scope b")?.textContent ?? "") : "";
        const dlgRows = dlg ? dlg.querySelectorAll(".solo-pb-sides > b").length : 0;
        const dlgGround = dlg && dlg.querySelector('[name="pb-shading"]') ? 1 : 0;
        const dlgEcho = cellOf("甲") ? (dlg?.querySelector('[name="pb-top"]')?.value ?? "") : "";
        const dlgPick = (() => {
          const s = dlg ? dlg.querySelector('[name="pb-left"]') : null;
          if (!s) return 0;
          s.value = "single";
          s.dispatchEvent(new Event("change", { bubbles: true }));
          return 1;
        })();
        const dlgApply = dlg ? (dlg.querySelector(".solo-pb-apply")?.click(), 1) : 0;
        await wait(340);
        const dlgGone = document.querySelector(".solo-pb-dialog") ? 0 : 1;
        const cellLeftNow = cellOf("甲") ? cssOf(cellOf("甲"), "borderLeftStyle") + "/" + cssOf(cellOf("甲"), "borderLeftWidth") : "";
        const dlgScopeKept = (boxPick()?.dataset?.scope ?? "") === "cells" ? 1 : 0;
        // #152：接上 w:tblPr/w:shd 与 w:tcPr/w:shd 的原位写回之后，底纹那一行在**格与表两层**都要摆得出来，
        // 且第一趟件里两层都没说过地时那一行留空（留空=没说，不是替它摆一个 auto 进去）。填了要纸上真铺得出色，
        // 没点到的那几格自己那层照旧透明（表那一圈的地从底下透上来，不算它们说了话）。
        // 每次落笔整篇重画，早先抓到的那两个 text node 早就脱离文档了，所以重新按字面找一遍再铺选区。
        const g1Opened = clickLeg("more");
        await wait(240);
        const g1 = document.querySelector(".solo-pb-dialog");
        const cellGroundRow = g1 && g1.querySelector('[name="pb-shading"]') ? 1 : 0;
        const cellGroundEcho = g1 ? String(g1.querySelector('[name="pb-shading"]')?.value ?? "?") : "?";
        const cellGroundPick = (() => {
          const f = g1 ? g1.querySelector('[name="pb-shading"]') : null;
          if (!f) return 0;
          f.value = "C00000";
          f.dispatchEvent(new Event("input", { bubbles: true }));
          return 1;
        })();
        const g1Submit = g1 ? (g1.querySelector("form")?.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true })), 1) : 0;
        await wait(340);
        const cellGroundNow = cellOf("甲") ? cssOf(cellOf("甲"), "backgroundColor") : "";
        const g1Gone = document.querySelector(".solo-pb-dialog") ? 0 : 1;
        const head2 = spanOf("甲");
        const tail2 = spanOf("末");
        if (head2 && tail2) {
          const rr = document.createRange();
          rr.setStart(head2, 0);
          rr.setEnd(tail2, (tail2.textContent ?? "").length);
          const rs = document.getSelection();
          rs?.removeAllRanges();
          if (rs) rs.addRange(rr);
        }
        await wait(240);
        const tblGroundScope = boxPick()?.dataset?.scope ?? "";
        const g2Opened = clickLeg("more");
        await wait(240);
        const g2 = document.querySelector(".solo-pb-dialog");
        const tblGroundRow = g2 && g2.querySelector('[name="pb-shading"]') ? 1 : 0;
        const tblGroundEcho = g2 ? String(g2.querySelector('[name="pb-shading"]')?.value ?? "?") : "?";
        const tblGroundPick = (() => {
          const f = g2 ? g2.querySelector('[name="pb-shading"]') : null;
          if (!f) return 0;
          f.value = "00B050";
          f.dispatchEvent(new Event("input", { bubbles: true }));
          return 1;
        })();
        const g2Submit = g2 ? (g2.querySelector("form")?.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true })), 1) : 0;
        await wait(340);
        const tblGroundNow = ringOf("甲") ? cssOf(ringOf("甲"), "backgroundColor") : "";
        const cellGroundStill = cellOf("甲") ? cssOf(cellOf("甲"), "backgroundColor") : "";
        const siblingGround = cellOf("末") ? cssOf(cellOf("末"), "backgroundColor") : "";
        const g2Gone = document.querySelector(".solo-pb-dialog") ? 0 : 1;
        // #149：开着修订落笔，那一圈的话不直接落定——件里留下一枚待判的记录（表那一层 w:tblPrChange、
        // 格那一层 w:tcPrChange）。这一腿在装机壳里按这个顺序问：先开「审阅窗格」数读回来的待判（第一趟 0，
        // 第二趟是上一趟按壳自己的路径写回、被壳读回来的那两条）；点上「修订」再落笔，表与格两层各留一条，
        // 待判期间纸上画的仍是刚点的那一笔（待判不是撤销）；只有第二趟接着判定——接受「表格属性」那一行之后
        // 记录没了线还在，拒绝「单元格属性」那一行之后那一格的底缝退回落笔之前。件里那两枚记录的 w:id/w:author
        // 与判定后的字节由 .scratch/tbl-cmd-install-check.sh 拿 libxml2 与 python 的标准库在外面问。
        const revTopOf = () => (ringOf("甲") ? cssOf(ringOf("甲"), "borderTopStyle") + "/" + cssOf(ringOf("甲"), "borderTopWidth") : "");
        const revCellBottomOf = () => (cellOf("甲") ? cssOf(cellOf("甲"), "borderBottomStyle") + "/" + cssOf(cellOf("甲"), "borderBottomWidth") : "");
        const revItems = () => [...document.querySelectorAll(".solo-rev-item")];
        const revKindList = () => revItems().map((it) => it.querySelector(".solo-rev-kind")?.textContent ?? "").join("|");
        const revRowOf = (kind) => revItems().find((it) => (it.querySelector(".solo-rev-kind")?.textContent ?? "") === kind);
        const revHit = (kind, label) => {
          if (otEarly) return 0;
          const btn = [...(revRowOf(kind)?.querySelectorAll("button") ?? [])].find((b) => b.textContent === label);
          if (!btn) return 0;
          btn.click();
          return 1;
        };
        // 窗格里那一行报的作者，与件里那句 w:author 必须是同一个名字——探针把这一个字交出去，外面按它问。
        const revAuthorOf = (kind) => {
          const spans = revRowOf(kind)?.querySelectorAll(".solo-rev-meta span");
          return spans && spans.length > 2 ? (spans[2].textContent ?? "") : "";
        };
        rb("审阅窗格")?.click();
        await wait(250);
        const revPaneOpened = document.querySelector(".solo-rev-panel:not([hidden])") ? 1 : 0;
        const revInRows = revItems().length;
        const revInKinds = revKindList();
        const revTrackBtn = rb("修订");
        revTrackBtn?.click();
        await wait(160);
        const revTracking = revTrackBtn?.classList.contains("on") ? 1 : 0;
        const revHead = spanOf("甲");
        const revTail = spanOf("末");
        if (revHead && revTail) {
          const vr = document.createRange();
          vr.setStart(revHead, 0);
          vr.setEnd(revTail, (revTail.textContent ?? "").length);
          const vs = document.getSelection();
          vs?.removeAllRanges();
          if (vs) vs.addRange(vr);
        }
        await wait(200);
        const revTableScope = boxPick()?.dataset?.scope ?? "";
        const revTopBefore = revTopOf();
        const revTableStroke = clickLeg("top");
        await wait(340);
        const revTopAfter = revTopOf();
        const revRowsTable = revItems().length;
        // 光标退回一格：签的是这一格自己的那一枚，落在这一格那一圈。
        const revAnchor = spanOf("甲");
        if (revAnchor) {
          const xr = document.createRange();
          xr.setStart(revAnchor, 1);
          xr.setEnd(revAnchor, 1);
          const xs = document.getSelection();
          xs?.removeAllRanges();
          if (xs) xs.addRange(xr);
        }
        await wait(200);
        const revCellScope = boxPick()?.dataset?.scope ?? "";
        const revCellBottomBefore = revCellBottomOf();
        const revCellStroke = clickLeg("bottom");
        await wait(340);
        const revCellBottomAfter = revCellBottomOf();
        const revRowsAll = revItems().length;
        const revKindsAll = revKindList();
        const revAuthor = revAuthorOf("表格属性");
        const revCellAuthor = revAuthorOf("单元格属性");
        const revCountText = document.querySelector(".solo-rev-count")?.textContent ?? "";
        const revDecide = revInRows > 0 ? 1 : 0;
        const revAcceptHit = revDecide ? revHit("表格属性", "接受") : 0;
        await wait(340);
        const revRowsAfterAccept = revItems().length;
        const revTopAfterAccept = revTopOf();
        const revRejectHit = revDecide ? revHit("单元格属性", "拒绝") : 0;
        await wait(340);
        const revRowsAfterReject = revItems().length;
        const revCellBottomAfterReject = revCellBottomOf();
        const revCellTopStill = cellOf("甲") ? cssOf(cellOf("甲"), "borderTopStyle") + "/" + cssOf(cellOf("甲"), "borderTopWidth") : "";
        const revCellGroundStill = cellOf("甲") ? cssOf(cellOf("甲"), "backgroundColor") : "";
        rb("审阅窗格")?.click();
        await wait(200);
        const revPaneClosed = document.querySelector(".solo-rev-panel:not([hidden])") ? 0 : 1;
        const revDirty = dirtyNow();
        // #190 的装机那一腿：「表格工具」这一页按光标露藏，命令落下去之后纸上那只盒要跟着网格走。
        // 只在探针自己那件（标题带 表格工具-装机）上跑：它会往表里真加一列，别的装机件各自有字节账，
        // 被这一腿牵动就成了假红/假绿。落笔在本节后面那句 solo:save 之前，所以盘上的件由壳自己写。
        // 这一段整篇是模板字符串里的 JS：不写类型、不写反引号、正则不带转义。
        const ttDoc = (document.getElementById("title")?.textContent ?? "").includes("表格工具-装机");
        let ttCmd = null;
        if (ttDoc) {
          // 上一腿（#149）把「修订」点开着交过来的，而这一腿要量的第一条落法是**没跟踪**那一种：表整张离开纸、
          // 位子上留一段可落的光标行。所以先按一次把跟踪关掉，量完这一条再按回去量第二种落法（表还站着，
          // 每一行自己说一句被删）。开关本身不写字节，关掉之后那两条宽度读数量的仍是未跟踪那一笔。
          const revBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const ttTrackBefore = revBtn()?.classList.contains("on") ? 1 : 0;
          if (ttTrackBefore) revBtn()?.click();
          await wait(200);
          const ttTrackOff = revBtn()?.classList.contains("on") ? 1 : 0;
          // 现问：这一腿每按一次命令纸就重铺一回，开头抓住的那只盒当场离开文档——照旧引用量到的是 0 宽与旧列数。
          const ttFind = () => [...document.querySelectorAll("#view .page table.doc-table")].find((x) => (x.textContent ?? "").includes("甲"));
          const ttBox = () => Math.round((ttFind()?.getBoundingClientRect().width ?? 0) * 10) / 10;
          // 两样分开报：CSS 上写的那一句是文件的数（8000 缇 ÷ 15.0001 = 533.3px），盒量到的是纸面真实那只盒。
          // 这两个数在这一件上**不该相等**：同一趟更早的 #146 那一腿在这张表上落过框线（它按表形状认件），
          // 文件说了边的格子里那条 1px 在 collapse 模型下半宽出地界，所以盒恒比写的数宽 0.5px。
          // 这一条量在 .scratch/tbl-ring-width-probe.html：边挂在 table 自身不涨宽，挂在外圈格子上才涨。
          // 「页面替每一格发明一条缝线」那条假象不靠这个差量证据，改由下面「查看网格线」那一腿的墨/导引互斥审计直接量。
          const ttWritten = () => {
            const el = ttFind();
            return el ? Math.round(parseFloat(el.style.width || "0") * 10) / 10 : 0;
          };
          const ttCols = () => ttFind()?.rows[0]?.cells.length ?? 0;
          const ttCell = ttFind()?.querySelector("tr")?.lastElementChild;
          const ttRange = document.createRange();
          ttRange.selectNodeContents(ttCell);
          ttRange.collapse(true);
          const ttSel = document.getSelection();
          ttSel?.removeAllRanges();
          if (ttSel) ttSel.addRange(ttRange);
          ttCell?.focus();
          document.dispatchEvent(new Event("selectionchange"));
          await wait(220);
          const ttTab = (label) => [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === label);
          const ttRowEls = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] button")];
          const ttTabHidden = ttTab("表格工具")?.hidden === true ? 1 : 0;
          const ttTabPresent = ttTab("表格工具") ? 1 : 0;
          // 露出来的是页签自己（renderer.ts 的 syncDocRibbon 里那句 tableTab.button.hidden = !inTable）。
          // 点一下页签之后那一行才摆出命令——这两处各报一次：光标进表时那一行的藏露，和按页签之后它的藏露。
          // 「上下文页自己跳上前台」是 WPS 那边未取证的一条，所以这里不照它下断言，只量产品自己说得出话的那两条。
          // 这一段整篇在模板字符串里：注释不能带反引号。
          const ttBarHidden = document.querySelector("#ribbon .rbar[data-tab='table']")?.hidden === true ? 1 : 0;
          ttTab("表格工具")?.click();
          await wait(200);
          const ttBarOpen = document.querySelector("#ribbon .rbar[data-tab='table']")?.hidden === true ? 1 : 0;
          const ttBtns = ttRowEls().map((b) => b.textContent).join("/");
          const ttPxBefore = ttBox();
          const ttWrittenBefore = ttWritten();
          const ttColsBefore = ttCols();
          const ttHit = ttRowEls().find((b) => b.textContent === "右侧插入列");
          const ttDisabled = ttHit?.disabled === true ? 1 : 0;
          ttHit?.click();
          await wait(360);
          const ttNotice = [...document.querySelectorAll(".doc-notice")].map((n) => (n.hidden ? "" : n.textContent ?? "")).join("");
          const ttPxAfter = ttBox();
          const ttWrittenAfter = ttWritten();
          const ttColsAfter = ttCols();
          // 「删除表格」（CT_TableTools.kuip:36 的 DeleteTable）在装机这一趟按两种落法各一次：先把光标放回这只表，
          // 跟踪关着时按下之后纸上这只表要没了、⌘Z 之后要回来；跟踪开着时按下表不得没，改成每一行说一句被删。
          // 量的是命令接在产品自己的 ribbon 上、落点算得对，与 DOM 页里那一趟模拟各说各话。
          // 盘上的字节由本节末尾那句 solo:save 写，这两次都撤销干净所以不动那一件。
          const ttAgainCell = ttFind()?.querySelector("tr")?.lastElementChild;
          if (ttAgainCell) {
            const ttAgainRange = document.createRange();
            ttAgainRange.selectNodeContents(ttAgainCell);
            ttAgainRange.collapse(true);
            const ttAgainSel = document.getSelection();
            ttAgainSel?.removeAllRanges();
            if (ttAgainSel) ttAgainSel.addRange(ttAgainRange);
            ttAgainCell.focus();
            document.dispatchEvent(new Event("selectionchange"));
          }
          await wait(220);
          // 只数**这一件那一张**：#view 里可能同时摆着别处的纸，全局表数说不出这一次按没按下。
          const ttJiaSel = () => [...document.querySelectorAll("#view .page table.doc-table")].filter((x) => (x.textContent ?? "").includes("甲")).length;
          // 段数只数**表外**的那些：这一趟之前别的装机腿（#158 那一枚 3×2 网格 picker）在同一件纸上插过空表，
          // 全局段数把那些格子里的行也算进来，删表留下的那一段就被淹在里头。
          const ttOutLines = () => [...document.querySelectorAll("#view .page .para")].filter((p) => !p.closest("table")).length;
          const ttNoticeSel = () => [...document.querySelectorAll(".doc-notice")].map((n) => (n.hidden ? "" : n.textContent ?? "")).join("");
          const ttDelBtn = ttRowEls().find((b) => b.textContent === "删除表格");
          const ttDelDisabled = ttDelBtn?.disabled === true ? 1 : 0;
          const ttPages = document.querySelectorAll("#view .page").length;
          // 先把这一页到底摆着什么读出来再说别的：纸上如果有第二只表，全局的表数与段数就说不出这一次按没按下。
          const ttPageKids = [...(document.querySelector("#view .page")?.children ?? [])]
            .map((c) => String(c.tagName) + "." + (c.className || ""))
            .join("|")
            .slice(0, 260);
          const ttTablesInfo = [...document.querySelectorAll("#view .page table.doc-table")]
            .map((t) => "[摆" + (t.parentElement?.className || t.parentElement?.tagName || "") + " " + t.rows.length + "行" + (t.rows[0] ? t.rows[0].cells.length : 0) + "格 " + (t.textContent || "").slice(0, 12) + "]")
            .join("");
          const ttJiaBefore = ttJiaSel();
          const ttJiaLinesBefore = ttFind()?.querySelectorAll(".para").length ?? 0;
          const ttOutBefore = ttOutLines();
          ttDelBtn?.click();
          await wait(360);
          const ttJiaGone = ttJiaSel();
          const ttOutAfter = ttOutLines();
          const ttDelNotice = ttNoticeSel();
          window.dispatchEvent(new KeyboardEvent("keydown", { key: "z", metaKey: true, bubbles: true }));
          await wait(360);
          const ttJiaBack = ttJiaSel();
          const ttOutBack = ttOutLines();
          // 第二种落法：修订开着时按同一条命令，表不得离开纸——每一行自己说一句被删（纸上一行一枚划掉的行），
          // 审核时逐行点头或摇头，全部点头之后收到的才是上面那一种形状。按下之前光标得重新放回这只表：
          // 撤销把纸重铺了一回，上一趟抓住的那只格当场离开文档。
          revBtn()?.click();
          await wait(200);
          const ttTrackOn = revBtn()?.classList.contains("on") ? 1 : 0;
          const ttAgain2 = ttFind()?.querySelector("tr")?.lastElementChild;
          if (ttAgain2) {
            const t2r = document.createRange();
            t2r.selectNodeContents(ttAgain2);
            t2r.collapse(true);
            const t2s = document.getSelection();
            t2s?.removeAllRanges();
            if (t2s) t2s.addRange(t2r);
            ttAgain2.focus();
            document.dispatchEvent(new Event("selectionchange"));
          }
          await wait(220);
          const ttDel2 = ttRowEls().find((b) => b.textContent === "删除表格");
          const ttDel2Disabled = ttDel2?.disabled === true ? 1 : 0;
          ttDel2?.click();
          await wait(360);
          const ttTrackTable = ttJiaSel();
          const ttStruck = ttFind()?.querySelectorAll("tr.rev-row.rev-del").length ?? 0;
          const ttTrackNotice = ttNoticeSel();
          window.dispatchEvent(new KeyboardEvent("keydown", { key: "z", metaKey: true, bubbles: true }));
          await wait(360);
          const ttTrackUndoTable = ttJiaSel();
          const ttStruckBack = ttFind()?.querySelectorAll("tr.rev-row.rev-del").length ?? 0;
          // 交还给后面的腿：跟踪保持关着——这一趟不该在盘上留下待判记录。
          revBtn()?.click();
          await wait(160);
          const ttTrackEnd = revBtn()?.classList.contains("on") ? 1 : 0;
          // 「删除单元格」（CT_TableTools.kuip:33 的 DeleteCells，Delete 那一档四条里的第一条）在装机这一趟量三条：
          // 最末一行选「上移」是命令自己算出来的回绝（小条要说那一句、这一行少不得一格）；第一行选「左移」按得下去，
          // 那一行少一格、被删那格的字离开这一行，撤销回得来；开着修订按同一条时那一格不许没——它只是自己说一句
          // 被删，「审阅窗格」里多出一行「删除单元格」，撤销之后那一行跟着退。三条都按完就撤销，不动盘上那件的字节账。
          // 下拉在产品自己那一页上，按 title 找；每按一次纸重铺一回，光标每次现问，不照旧引用。
          const ttPut = (el) => {
            if (!el) return;
            const r = document.createRange();
            r.selectNodeContents(el);
            r.collapse(true);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            el.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
          };
          const ttUndoKey = () => window.dispatchEvent(new KeyboardEvent("keydown", { key: "z", metaKey: true, bubbles: true }));
          const ttCellChooser = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] select")]
            .find((x) => (x.title || "").indexOf("删除单元格") === 0);
          const ttChooseCell = (value) => {
            const s = ttCellChooser();
            if (!s) return 0;
            s.value = value;
            s.dispatchEvent(new Event("change"));
            return 1;
          };
          const ttRow0 = () => ttFind()?.rows[0] ?? null;
          const ttRow0Cells = () => ttRow0()?.cells.length ?? 0;
          const ttRow0Text = () => ttRow0()?.textContent ?? "";
          const ttBottomRow = () => { const t = ttFind(); return t ? t.rows[t.rows.length - 1] : null; };
          const ttPaneOpen = () => (document.querySelector(".solo-rev-panel:not([hidden])") ? 1 : 0);
          const ttCellDelItems = () => [...document.querySelectorAll(".solo-rev-item")]
            .filter((it) => (it.querySelector(".solo-rev-kind")?.textContent ?? "") === "删除单元格").length;
          const ttChooserHit = ttCellChooser() ? 1 : 0;
          const ttChooserDisabled = ttCellChooser()?.disabled === true ? 1 : 0;
          const ttNoticeBeforeUp = ttNoticeSel();
          const ttBottom = () => { const r = ttBottomRow(); return r ? r.cells.length : 0; };
          const ttBottomBefore = ttBottom();
          ttPut(ttBottomRow()?.cells[0]);
          await wait(220);
          const ttUpPick = ttChooseCell("up");
          await wait(360);
          const ttUpNotice = ttNoticeSel();
          const ttBottomAfterUp = ttBottom();
          // 小条是四秒后退场的浮层：上一击那句拒绝还没退就按下一击，「这一击没被回绝」就成了一条假绿。等它退干净再按。
          await wait(4200);
          const ttLeftTarget = ttFind()?.rows[0]?.cells[1] ?? null;
          const ttTargetText = ttLeftTarget?.textContent ?? "";
          ttPut(ttLeftTarget);
          await wait(220);
          const ttRow0Before = ttRow0Cells();
          const ttTextBefore = ttRow0Text();
          const ttLeftPick = ttChooseCell("left");
          await wait(360);
          const ttRow0AfterLeft = ttRow0Cells();
          const ttTextAfterLeft = ttRow0Text();
          const ttLeftNotice = ttNoticeSel();
          ttUndoKey();
          await wait(360);
          const ttRow0Back = ttRow0Cells();
          const ttTextBack = ttRow0Text();
          revBtn()?.click();
          await wait(200);
          const ttCellTrackOn = revBtn()?.classList.contains("on") ? 1 : 0;
          if (!ttPaneOpen()) rb("审阅窗格")?.click();
          await wait(260);
          const ttPaneRead = ttPaneOpen();
          const ttCellDelBefore = ttCellDelItems();
          ttPut(ttFind()?.rows[0]?.cells[1]);
          await wait(220);
          const ttSignPick = ttChooseCell("left");
          await wait(360);
          const ttRow0Signed = ttRow0Cells();
          const ttTextSigned = ttRow0Text();
          const ttCellDelSigned = ttCellDelItems();
          const ttSignedNotice = ttNoticeSel();
          ttUndoKey();
          await wait(360);
          const ttCellDelAfterUndo = ttCellDelItems();
          const ttRow0AfterUndo = ttRow0Cells();
          revBtn()?.click();
          await wait(160);
          const ttCellTrackEnd = revBtn()?.classList.contains("on") ? 1 : 0;
          // 「单元格尺寸」那两枚自旋格（RB_Adjust）在装机这一趟量四条：件没说一句行高时框里空着（不替它编数）；
          // 打下去之后那一行自己说一句 w:trHeight、纸上那一行长高；宽度那一格动的是整列（网格加这一列每一格），
          // 盒宽跟着涨、只有一列宽的那些格才动；光标退到表外时两格都空、按下去被回绝。四条按完全部撤销，
          // 不动这一件的字节账。字面按状态栏的测量单位（默认厘米）：2.54 厘米 = 1440 缇、5.292 厘米 = 3000 缇。
          // 这一段整篇在模板字符串里：不写类型、不写反引号、注释不带转义。
          const ttSpin = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] .rspin")]
            .find((x) => (x.textContent || "").trim().indexOf(label) === 0)?.querySelector("input");
          const ttType = (input, v) => {
            if (!input) return 0;
            input.focus();
            input.value = v;
            input.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          const ttRowH = () => { const r = ttFind()?.rows[1]; return r ? Math.round(r.getBoundingClientRect().height * 10) / 10 : -1; };
          const ttColW = (i) => { const c = ttFind()?.rows[1]?.cells[i]; return c ? Math.round(c.getBoundingClientRect().width * 10) / 10 : -1; };
          ttPut(ttFind()?.rows[1]?.cells[0]);
          await wait(240);
          const ttSizeHit = (ttSpin("高度") ? 1 : 0) + (ttSpin("宽度") ? 1 : 0);
          const ttHeightEcho0 = ttSpin("高度")?.value ?? "?";
          const ttWidthEcho0 = ttSpin("宽度")?.value ?? "?";
          const ttRowHBefore = ttRowH();
          const ttHeightPress = ttType(ttSpin("高度"), "2.54");
          await wait(360);
          const ttRowHAfter = ttRowH();
          const ttHeightEcho1 = ttSpin("高度")?.value ?? "?";
          ttUndoKey();
          await wait(360);
          const ttRowHUndone = ttRowH();
          const ttHeightEcho2 = ttSpin("高度")?.value ?? "?";
          const ttBoxBeforeW = ttBox();
          const ttCol0BeforeW = ttColW(0);
          const ttCol1BeforeW = ttColW(1);
          const ttWidthPress = ttType(ttSpin("宽度"), "5.292");
          await wait(360);
          const ttBoxAfterW = ttBox();
          const ttCol0AfterW = ttColW(0);
          const ttCol1AfterW = ttColW(1);
          const ttWidthEcho1 = ttSpin("宽度")?.value ?? "?";
          ttUndoKey();
          await wait(360);
          const ttBoxUndoneW = ttBox();
          const ttSizeOut = [...document.querySelectorAll("#view .page .para")].find((el) => !el.closest("table"));
          ttPut(ttSizeOut);
          await wait(240);
          const ttHeightEchoOut = ttSpin("高度")?.value ?? "?";
          const ttWidthEchoOut = ttSpin("宽度")?.value ?? "?";
          const ttOutPress = ttType(ttSpin("高度"), "2.54");
          await wait(360);
          const ttOutNotice = ttNoticeSel();
          const ttHeightEchoRefused = ttSpin("高度")?.value ?? "?";
          const ttRowHRefused = ttRowH();
          const ttBoxRefused = ttBox();
          // 「对齐」那一组六枚（CT_TableTools.kuip:156-170 的 RB_Alignment，WPS 自己摆成两行）在装机这一趟量五条：
          // 起手回显、落笔、回绝、撤销、两页同源。落点两层是这一组的要害：竖向那三枚归**格**（w:tcPr/w:vAlign），
          // 横向那三枚归格里的**每一段**（w:pPr/w:jc）——所以纸上跟着动的是两样东西（td 的 vertical-align 与
          // 那一段的 text-align），件里跟着动的是两句话，各按各的落笔都量一次。
          // 头一步按产品**自己的回显**挑一格：第一行**从右往左**第一枚「↑ 与 左 都亮」的。为什么从右往左——
          // 这一腿最后那一击不落撤销，要交本节末尾那句 solo:save 写回盘上，让第二趟把它读回来；而上面 #149 那一腿
          // 在第二趟会按一次「单元格属性 → 拒绝」，那句 w:tcPrChange 记的是**整套 tcPr**，拒绝就是把那一套整个
          // 退回「改前」——连本腿上一趟在跟踪关着时临时补上的那句 w:vAlign 一起退掉（装机的壳与 .scratch/
          // valign-readback-dom.ts 各量到一次：竖向那句随整套退了、段里那句 w:jc 还在，因为 jc 归段不归格；
          // WPS/Word 对「记录之后非跟踪补上的那一句」该退到哪一档 未取证）。所以跨趟那句证据只能压在
          // **修订那一腿不签名的格子**上：它锚在甲、范围到末，而每趟「右侧插入列」刚插出来的最右那一格谁也不签。
          // 第一趟落第 3 格（新插的空格），第二趟落第 4 格（这一趟新插的），它左边那一格（第 3 格）就是上一趟留下的：
          // 装机的壳在第二趟开头把上一趟写回的那两句读回来、照它亮 ↕ 与 中，那才是
          // 「装机的壳把装机的壳写下的那两句话读回来并照它亮」（与上面宽度那一腿的第二趟起手读数同一条口径）。
          // 空格也落得下去：它带一段空段（纸上有占位），按 ↕ 与 中 之后件里多出 w:vAlign center 与 w:jc center 各一枚。
          // 「没说」与「明说左」在纸上是同一张纸（render.ts 把两档都画成左；授权 46 份里光标停得进的 2,910 段有
          // 88 段明说 left，.scratch/authorized-jc-echo.ts），所以起手亮的那枚「左」读的是并成一档的两样。
          // 小条是四秒后退场的浮层：上一腿（尺寸）在表外按下的那句回绝还挂着，不先等它退干净，
          // 「这一击没被回绝」就成了假绿。这一段整篇在模板字符串里：不写类型、不写反引号、注释不带转义。
          await wait(4200);
          const ttAlignKeys = ["↑", "↕", "↓", "左", "中", "右"];
          const ttAlignBtn = (label) => ttRowEls().find((b) => b.textContent === label);
          const ttHomeBtn = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button")]
            .find((b) => b.textContent === label);
          const ttLitIn = (btn) => (btn && btn.classList.contains("on") ? 1 : 0);
          const ttAlignLit = () => ttAlignKeys.filter((k) => ttLitIn(ttAlignBtn(k)) === 1).join("");
          const ttHomeLit = () => ["↑", "↕", "↓"].filter((k) => ttLitIn(ttHomeBtn(k)) === 1).join("");
          const ttRow0N = () => ttFind()?.rows[0]?.cells.length ?? 0;
          const ttRow0Cell = (i) => (i < 0 || i >= ttRow0N() ? null : ttFind()?.rows[0]?.cells[i] ?? null);
          let ttAlignCol = -1;
          const ttAlignTd = () => (ttAlignCol < 0 ? null : ttRow0Cell(ttAlignCol));
          const ttAlignPara = () => ttAlignTd()?.querySelector(".para") ?? ttAlignTd();
          const ttValign = () => cssOf(ttAlignTd(), "verticalAlign");
          const ttTextAlign = () => cssOf(ttAlignPara(), "textAlign");
          const ttPress = (label) => {
            const b = ttAlignBtn(label);
            if (!b) return 0;
            b.click();
            return 1;
          };
          for (let i = ttRow0N() - 1; i >= 0; i--) {
            const el = ttRow0Cell(i);
            if (!el) continue;
            ttPut(el);
            await wait(240);
            if (ttLitIn(ttAlignBtn("↑")) === 1 && ttLitIn(ttAlignBtn("左")) === 1) {
              ttAlignCol = i;
              break;
            }
          }
          const ttAlignTrack = revBtn()?.classList.contains("on") ? 1 : 0;
          const ttAlignTabShow = ttTab("表格工具") ? 1 : 0;
          ttTab("表格工具")?.click();
          await wait(200);
          const ttAlignGroups = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] .grp")];
          const ttAlignGroupIdx = ttAlignGroups.findIndex((g) => [...g.querySelectorAll("button")].map((b) => b.textContent).join("|") === "↑|↕|↓|左|中|右");
          const ttAlignGroupN = ttAlignGroupIdx < 0 ? 0 : ttAlignGroups[ttAlignGroupIdx].querySelectorAll("button").length;
          const ttAlignSpinN = ttAlignGroupIdx < 0 ? -1 : ttAlignGroups[ttAlignGroupIdx].querySelectorAll(".rspin").length;
          const ttAlignBarShown = document.querySelector("#ribbon .rbar[data-tab='table']")?.hidden === true ? 0 : 1;
          const ttAlignCellText = ttAlignTd()?.textContent ?? "";
          const ttStartLit = ttAlignLit();
          const ttStartHome = ttHomeLit();
          const ttValign0 = ttValign();
          const ttTextAlign0 = ttTextAlign();
          const ttNoticeV0 = ttNoticeSel();
          const ttPressMid = ttPress("↕");
          await wait(360);
          const ttLitAfterV = ttAlignLit();
          const ttHomeAfterV = ttHomeLit();
          const ttValignAfterV = ttValign();
          const ttTextAlignAfterV = ttTextAlign();
          const ttNoticeAfterV = ttNoticeSel();
          const ttDisabledAfterV = ttAlignBtn("↕")?.disabled === true ? 1 : 0;
          await wait(4200);
          const ttRefuseV = ttPress("↕");
          await wait(360);
          const ttNoticeRefuseV = ttNoticeSel();
          const ttLitRefuseV = ttAlignLit();
          const ttValignRefuseV = ttValign();
          ttUndoKey();
          await wait(360);
          const ttLitUndoV = ttAlignLit();
          const ttValignUndoV = ttValign();
          await wait(4200);
          const ttPressCenter = ttPress("中");
          await wait(360);
          const ttLitAfterH = ttAlignLit();
          const ttHomeAfterH = ttHomeLit();
          const ttTextAlignAfterH = ttTextAlign();
          const ttValignAfterH = ttValign();
          const ttNoticeAfterH = ttNoticeSel();
          await wait(4200);
          const ttRefuseH = ttPress("中");
          await wait(360);
          const ttNoticeRefuseH = ttNoticeSel();
          const ttLitRefuseH = ttAlignLit();
          const ttTextAlignRefuseH = ttTextAlign();
          ttUndoKey();
          await wait(360);
          const ttLitUndoH = ttAlignLit();
          const ttTextAlignUndoH = ttTextAlign();
          await wait(4200);
          // 最后这一击不落撤销：交本节末尾那句 solo:save 写回盘上，好让第二趟的起手与 disk() 各读到自己那一句。
          ttPress("↕");
          await wait(360);
          ttPress("中");
          await wait(360);
          const ttLitEnd = ttAlignLit();
          const ttHomeEnd = ttHomeLit();
          const ttValignEnd = ttValign();
          const ttTextAlignEnd = ttTextAlign();
          const ttNoticeEnd = ttNoticeSel();
          // 上一趟留下的那一格（这一趟新插的空格左边那一个；第一趟它左边是丙，件里两句话都没说，读不到 ↕ 与 中）：
          // 第二趟该读回上一趟壳自己写回盘上的那两句，回显是 ↕ 与 中 两枚同亮。
          const ttPrevCol = ttAlignCol - 1;
          let ttPrevLit = "";
          if (ttPrevCol >= 0) {
            ttPut(ttRow0Cell(ttPrevCol));
            await wait(240);
            ttPrevLit = ttAlignLit();
          }
          // 「查看网格线」那一枚（CT_TableTools.kuip:8 的 ShowGridlines，RB_Table 组里第一枚摆得出的）在装机这一趟量五条：
          // 起手这一档开着且导引真画得出、关掉之后墨那一本账一条不动、开与关之间纸一个像素不挪、按两下回到原档而件不脏，
          // 再加上「同一条边不许既有墨又有导引」（#196 之前页面替每一格发明一条边，量的正是这一枚）。
          // 夹具（.scratch/make-table-tools-smoke.ts）刚拼出来时一句边也没说，但**同一趟 smoke 早几步**的 #146 那一腿
          // 会在这张表上落框线（它按表形状认件，不按标题），所以「墨」这一趟不是一张白纸：断的是「墨只随文件说的那句动、
          // 不随这一档动」与「有墨的边不留导引」，不是「六格 24 条边全没墨」。
          // 按钮每次现问：这一页的 ribbon 会随 syncDocRibbon 重排，开头抓住的那枚节点当场离开文档。
          const glBtn = () => ttRowEls().find((b) => (b.textContent ?? "").includes("查看网格线"));
          const glViewOn = () => (document.getElementById("view")?.classList.contains("doc-grid") ? 1 : 0);
          const glLit = () => (glBtn()?.classList.contains("on") ? 1 : 0);
          const glTds = () => [...(ttFind()?.querySelectorAll("td") ?? [])].slice(0, 6);
          const glShadow = () => glTds().map((td) => (getComputedStyle(td).boxShadow === "none" ? 0 : 1)).join("");
          const glInk = () => glTds().map((td) => {
            const s = getComputedStyle(td);
            return [s.borderTopStyle, s.borderRightStyle, s.borderBottomStyle, s.borderLeftStyle].join("/");
          }).join("|");
          // 「页面替文件说话」那一枚缺陷的签名：同一条物理线上**既有墨又有导引**（导引只该留给没人画线的那些线）。
          // 反过来「既没墨又没导引」是这一档开着时屏幕上看不见地界。问「这条线上有没有墨」得问**那条线**而不是手上那一格：
          // 外圈那四条线是表自己说的，墨写在 table 元素身上（render.ts 的 edgeOf 同一口径），内部那一条缝由贴着它的两格
          // 任一格的墨代表。四条边的签名互不为子串，所以能从一整串 --gl 里逐条认出来（串本身画得出这条量在
          // .scratch/bs-td-probe.html）。
          const glSig = { top: "inset 0 1px 0", bottom: "inset 0 -1px 0", left: "inset 1px 0 0", right: "inset -1px 0 0" };
          const glAudit = () => {
            const t = ttFind();
            if (!t) return { edges: 0, both: "no-table", neither: "no-table" };
            const CAP = { top: "Top", bottom: "Bottom", left: "Left", right: "Right" };
            const OPP = { top: "bottom", bottom: "top", left: "right", right: "left" };
            const SIDES = ["top", "bottom", "left", "right"];
            const both = [];
            const neither = [];
            let edges = 0;
            const inkOf = (el, side) => getComputedStyle(el)["border" + CAP[side] + "Style"] !== "none";
            const near = (a, b) => Math.abs(a - b) < 1.5;
            const tr = t.getBoundingClientRect();
            const all = [...t.querySelectorAll("td")].map((td) => ({ el: td, q: td.getBoundingClientRect() }));
            const overX = (a, b) => a.q.left < b.q.right - 1 && a.q.right > b.q.left + 1;
            const overY = (a, b) => a.q.top < b.q.bottom - 1 && a.q.bottom > b.q.top + 1;
            const lineInk = (a, side) => {
              if (inkOf(a.el, side)) return true;
              const ring = side === "top" ? near(a.q.top, tr.top) : side === "bottom" ? near(a.q.bottom, tr.bottom)
                : side === "left" ? near(a.q.left, tr.left) : near(a.q.right, tr.right);
              if (ring && inkOf(t, side)) return true;
              const other = OPP[side];
              for (let j = 0; j < all.length; j++) {
                const b = all[j];
                if (b.el === a.el) continue;
                const face = side === "top" ? near(b.q.bottom, a.q.top) && overX(a, b)
                  : side === "bottom" ? near(b.q.top, a.q.bottom) && overX(a, b)
                  : side === "left" ? near(b.q.right, a.q.left) && overY(a, b) : near(b.q.left, a.q.right) && overY(a, b);
                if (face && inkOf(b.el, other)) return true;
              }
              return false;
            };
            const sample = glTds();
            for (let i = 0; i < sample.length; i++) {
              const a = { el: sample[i], q: sample[i].getBoundingClientRect() };
              const list = getComputedStyle(a.el).getPropertyValue("--gl");
              for (let k = 0; k < 4; k++) {
                const side = SIDES[k];
                edges++;
                const ink = lineInk(a, side);
                const guide = list.indexOf(glSig[side]) >= 0;
                if (ink && guide) both.push(i + "." + side);
                if (!ink && !guide) neither.push(i + "." + side);
              }
            }
            return { edges, both: both.join(","), neither: neither.join(",") };
          };
          const glRect = () => {
            const t = ttFind();
            const r = t ? t.getBoundingClientRect() : null;
            const box = r ? Math.round(r.width * 10) / 10 + "x" + Math.round(r.height * 10) / 10 : "no-table";
            return box + "|" + glTds().map((td) => {
              const q = td.getBoundingClientRect();
              return Math.round(q.width * 100) + "x" + Math.round(q.height * 100);
            }).join(",");
          };
          const glGroups = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] .grp")];
          const glGroupIdx = glGroups.findIndex((g) => [...g.querySelectorAll("button")].some((b) => (b.textContent ?? "").includes("查看网格线")));
          const glStartView = glViewOn();
          const glStartLit = glLit();
          const glStartShadow = glShadow();
          const glStartInk = glInk();
          const glStartAudit = glAudit();
          const glStartRect = glRect();
          // 视图那一档不该把件弄脏，也不该给谁回绝一句（这一枚归视图，不归选区）。
          // 起手那一档标题上已经挂着 ●（上面几腿是真编辑），所以这里问的是「按两下前后那枚 ● 有没有换边」，不是它亮不亮。
          const glDirty0 = (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0;
          glBtn()?.click();
          await wait(320);
          const glOffView = glViewOn();
          const glOffLit = glLit();
          const glOffShadow = glShadow();
          const glOffInk = glInk();
          // 关掉这一档之后量的这一条只报 both，不报 neither——量到的不是产品改了导引：那串 --gl 里嵌着
          // var(--gl-ink)，而 --gl-ink 只挂在容器那枚 doc-grid 上（index.html 的 463-464 行），档一关这个变量按
          // CSS 变量的规矩整个失效，computed --gl 就序列化成空串。所以那一档问得出「没有哪条边墨与导引叠在一起」，
          // 问不出「导引的字面还是不是原来那串」——后者由起手那一条与画得出/退场那两条一起管。
          const glOffAudit = glAudit();
          const glOffRect = glRect();
          const glOffNotice = ttNoticeSel();
          glBtn()?.click();
          await wait(320);
          const glBackView = glViewOn();
          const glBackLit = glLit();
          const glBackShadow = glShadow();
          const glBackRect = glRect();
          const glDirty1 = (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0;
          // 「表格属性」那一枚（CT_TableTools.kuip:7 的 TableProperties，RB_Table 组里排在 ShowGridlines **之前**）
          // 在装机这一趟量六条：按钮与组序摆得出、对话框起手摆的是件里那一句、改完纸跟着动、撤销回得来、
          // 再开一次「取消」不动文件、「原样确定」（这里按「取消」那一支等价：一个字段没碰）不留脏字。
          // 起点是件里**说过**的那几句：夹具说了一句 w:tblLayout fixed、没说过对齐/缩进/默认边距（make-table-tools-smoke.ts），
          // 所以起手 固定列宽=勾，其余格子空着。落笔挑的是**左对齐 + 1.27 厘米缩进**：居中/右对齐在纸上只改
          // margin-left:auto，而这张表已宽过正文栏，auto 解析出来的像素与「没说」分不开——左对齐加缩进才是纸上量得到的那一格。
          // 默认边距落 0.19 厘米（≈108 缇）在纸上就是 td 的 padding-left。最后整条撤销，盘上那件交本节末尾那句
          // solo:save 写回（这一腿不许牵动 #190/#195 那两笔字节账）。
          const tpGrp = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] .grp")]
            .find((g) => [...g.querySelectorAll("button")].some((b) => b.textContent === "属性"));
          const tpGrpLabels = tpGrp ? [...tpGrp.querySelectorAll("button")].map((b) => b.textContent).join("|") : "";
          const tpBtn = () => ttRowEls().find((b) => b.textContent === "属性");
          const tpDlg = () => document.querySelector(".solo-tp-dialog");
          const tpField = (n) => tpDlg()?.querySelector("[name='" + n + "']") ?? null;
          const tpShot = () => {
            const t = ttFind();
            const td = ttAlignTd();
            if (!t) return "no-table";
            const s = getComputedStyle(t);
            return [s.tableLayout, s.marginLeft, td ? getComputedStyle(td).paddingLeft : "no-cell"].join("|");
          };
          const tpSet = (el, v) => {
            if (!el) return 0;
            el.value = v;
            el.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          const tpBefore = tpShot();
          const tpBtnPresent = tpBtn() ? 1 : 0;
          const tpBtnLabel = tpBtn()?.textContent ?? "";
          tpBtn()?.click();
          await wait(320);
          const tpOpen = tpDlg()?.open === true ? 1 : 0;
          const tpTitle = tpDlg()?.querySelector("h2")?.textContent ?? "";
          const tpRows = tpDlg()?.querySelectorAll(".solo-tp-row,.solo-tp-check").length ?? 0;
          const tpEchoAlign = tpField("tp-align")?.value ?? "?";
          const tpEchoIndent = tpField("tp-indent")?.value ?? "?";
          const tpEchoFixed = tpField("tp-fixed")?.checked === true ? 1 : 0;
          const tpEchoMargin = ["top", "left", "bottom", "right"].map((s) => tpField("tp-m-" + s)?.value ?? "?").join(",");
          tpSet(tpField("tp-align"), "left");
          tpSet(tpField("tp-indent"), "1.27");
          if (tpField("tp-fixed")) {
            tpField("tp-fixed").checked = false;
            tpField("tp-fixed").dispatchEvent(new Event("change", { bubbles: true }));
          }
          tpSet(tpField("tp-m-left"), "0.19");
          await wait(140);
          const tpError = tpDlg()?.querySelector(".solo-tp-error")?.textContent ?? "";
          tpDlg()?.querySelector(".solo-tp-apply")?.click();
          await wait(400);
          const tpGone = tpDlg() ? 0 : 1;
          const tpAfter = tpShot();
          const tpWritten = tpAfter !== tpBefore ? 1 : 0;
          ttUndoKey();
          await wait(400);
          const tpUndo = tpShot();
          const tpUndoBack = tpUndo === tpBefore ? 1 : 0;
          tpBtn()?.click();
          await wait(320);
          const tpReopen = tpDlg()?.open === true ? 1 : 0;
          tpDlg()?.querySelector(".solo-tp-cancel")?.click();
          await wait(220);
          const tpCancelGone = tpDlg() ? 0 : 1;
          const tpCancelClean = tpShot() === tpBefore ? 1 : 0;
          const tpDialogsLeft = document.querySelectorAll(".solo-tp-dialog").length;
          // 「标题行重复」（RB_Data 组那一枚，WPS 的 Repeat Header Rows）在装机这一趟量五条：按钮摆得出、起手灭着且
          // 一行都不带 data-tbl-header；光标在第一行按下去之后从表头到光标那一行都标上；⌘Z 一起退；光标在最后一行按
          // 时命令自己回绝并说出理由；开着修订按同一条只添一份「行属性」记录、撤销之后那一行跟着退。整条撤销收尾，
          // 不动盘上那件的字节账。这一件的表只有两行、一页放得下，所以续页复件数在这一趟恒为 0——复件与「给表体让位」
          // 由 .scratch/tbl-header-dom.ts 那一页量（真 Chrome 里 3 页 62 行）。
          const rhBtn = () => ttRowEls().find((b) => b.textContent === "标题行重复");
          const rhRows = () => [...(ttFind()?.rows ?? [])].filter((r) => !r.classList.contains("solo-repeat-header-row"));
          const rhStamped = () => rhRows().map((r) => (r.getAttribute("data-tbl-header") === "1" ? 1 : 0)).join("");
          const rhOn = () => (rhBtn()?.classList.contains("on") ? 1 : 0);
          const rhClones = () => ttFind()?.querySelectorAll(".solo-repeat-header-row").length ?? -1;
          const rhItems = () => [...document.querySelectorAll(".solo-rev-item")]
            .filter((it) => (it.querySelector(".solo-rev-kind")?.textContent ?? "") === "行属性").length;
          ttPut(ttRow0Cell(0));
          await wait(240);
          const rhBtnPresent = rhBtn() ? 1 : 0;
          const rhStartDisabled = rhBtn()?.disabled === true ? 1 : 0;
          const rhStartOn = rhOn();
          const rhStartStamped = rhStamped();
          const rhStartClones = rhClones();
          rhBtn()?.click();
          await wait(360);
          const rhAfterOn = rhOn();
          const rhAfterStamped = rhStamped();
          const rhAfterClones = rhClones();
          const rhNotice = ttNoticeSel();
          ttUndoKey();
          await wait(360);
          const rhUndoStamped = rhStamped();
          const rhUndoOn = rhOn();
          // 小条是四秒后退场的浮层：上一击那句「最后一行不能是标题行」挂着就按下一击，「这一击没被回绝」就是假绿。
          await wait(4200);
          ttPut(ttBottomRow()?.cells[0]);
          await wait(240);
          rhBtn()?.click();
          await wait(360);
          const rhLastStamped = rhStamped();
          const rhLastNotice = ttNoticeSel();
          await wait(4200);
          revBtn()?.click();
          await wait(200);
          const rhTrackOn = revBtn()?.classList.contains("on") ? 1 : 0;
          if (!ttPaneOpen()) rb("审阅窗格")?.click();
          await wait(260);
          const rhItemsBefore = rhItems();
          ttPut(ttRow0Cell(0));
          await wait(240);
          rhBtn()?.click();
          await wait(400);
          const rhItemsSigned = rhItems();
          const rhStampedSigned = rhStamped();
          ttUndoKey();
          await wait(400);
          const rhItemsUndone = rhItems();
          const rhStampedUndone = rhStamped();
          revBtn()?.click();
          await wait(160);
          const rhTrackEnd = revBtn()?.classList.contains("on") ? 1 : 0;
          // 光标交回表外：下面那一段量的还是「表一退场、页签跟着退、档回到开始」，与这一腿各说各的。
          const ttCursorBackInCell = ttAlignTd() ? 1 : 0;
          // 光标退出表挑的是**表外那一段**，不是「第一张纸最后一个孩子」：表涨宽后会挪到别的纸上，
          // 第一张纸的末 child 可能就是这只表——那样量到的是分页，不是那一页的退场。
          const ttOut = [...document.querySelectorAll("#view .para")].find((el) => !el.closest("table"));
          const ttBackRange = document.createRange();
          ttBackRange.selectNodeContents(ttOut);
          ttBackRange.collapse(true);
          const ttBackSel = document.getSelection();
          ttBackSel?.removeAllRanges();
          if (ttBackSel) ttBackSel.addRange(ttBackRange);
          ttOut?.focus?.();
          document.dispatchEvent(new Event("selectionchange"));
          await wait(260);
          const ttTabBackHidden = ttTab("表格工具")?.hidden === true ? 1 : 0;
          const ttHomeOn = ttTab("开始")?.classList.contains("on") ? 1 : 0;
          const ttInTableAgain = ttOut ? 0 : 1;
          // 光标退到表外之后那六枚该一起灰着、一枚都不亮（表外没有格子可摆），开始页那三枚同样灰着。
          const ttAlignOutLit = ttAlignLit();
          const ttAlignOutDisabled = ttAlignKeys.filter((k) => ttAlignBtn(k)?.disabled === true).length;
          const ttHomeOutDisabled = ["↑", "↕", "↓"].filter((k) => ttHomeBtn(k)?.disabled === true).length;
          // 「转换成文本」（CT_TableTools.kuip:89 的 ConvertTableToText，RB_Data 那组的第二格）在装机这只盒子里量七条：
          // ①光标在表外时那一格灰着、四档都在（首档是「转换成文本」自己，转换没有回显可言）；②这一腿在正文里自己插
          // 一张干净的 2×2——insertTable 把光标放进新表第一格，那一格跟着不灰、字面仍是首档，而按表序 index 问回来的
          // 那一只与按光标问回来的是同一只；③开着修订按「逗号分隔」不许把表变没——OOXML 没有「一张表变成文本」这种
          // 修订，命令自己回绝、小条说一句、表一张不少、导引线一枚不多；④关回修订后落「制表符分隔」这一档，那张表
          // 从纸上没了、位子上那些段里多出导引线（一档一枚 span.tab，只数表外那些，算的是行数×(列数−1)），带「甲」
          // 那一只一张没少；⑤页签跟着退场、那一格重新灰着（命令把光标放回表原来那个位子上的第一段）；⑥第一下 ⌘Z
          // 撤转换：表回到同一个 index、行列与导引线回起手、页签重新露；⑦第二下 ⌘Z 撤插表：表数回起手、甲那一只还
          // 在、导引线也回起手。整条以撤销收尾，不许牵动盘上那件的字节账（两趟开的同一件）。
          // 「落得下去」这一条为什么不在「甲」那一只上量：前面那几腿在它身上留了一枚待判的属性改动，命令对它自己回绝
          // （「这张表上有一条待判的属性改动，转成文本会把它丢掉」）；那条回绝分支连同件里那句 w:tab 的字节账归
          // .scratch/tbl-text-dom.ts（它把同一份 writeDocx 的 XML 数出来）与 tests/doc-table-to-text.test.ts，
          // 这里断的是「装机的壳按这一格落得下去、落下去纸上是什么、退得回来」。
          const cvChooser = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] select")]
            .find((x) => (x.title || "").indexOf("转换成文本") === 0);
          const cvTabSpanN = () => [...document.querySelectorAll("#view .page span.tab")].filter((x) => !x.closest("table")).length;
          const cvTables = () => [...document.querySelectorAll("#view .page table.doc-table")];
          // 新插那一只表只留它在整篇表序里的位置：每按一次命令纸重铺一回，抓在手里的元素当场离开文档。
          const cvTableAt = (i) => cvTables()[i] ?? null;
          const cvCaretTable = () => {
            const n = document.getSelection()?.anchorNode ?? null;
            const el = n ? (n.nodeType === Node.TEXT_NODE ? n.parentElement : n) : null;
            return el?.closest("#view .page table.doc-table") ?? null;
          };
          const cvPress = (value) => { const el = cvChooser(); if (!el) return 0; el.value = value; el.dispatchEvent(new Event("change")); return 1; };
          const cvChooserHit = cvChooser() ? 1 : 0;
          const cvOutDisabled = cvChooser()?.disabled === true ? 1 : 0;
          const cvOptions = cvChooser()?.options.length ?? 0;
          const cvHomeValueOut = cvChooser()?.value ?? "";
          const cvN0 = cvTables().length;
          const cvTabs0 = cvTabSpanN();
          const cvTrackStart = revBtn()?.classList.contains("on") ? 1 : 0;
          const cvInsertBtn = rb("表格");
          const cvInsertHit = cvInsertBtn ? 1 : 0;
          cvInsertBtn?.click();
          await wait(220);
          const cvPickEl = document.querySelector(".gridpick button[data-c='2'][data-r='2']");
          const cvPickHit = cvPickEl ? 1 : 0;
          cvPickEl?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
          await wait(360);
          const cvTablesIns = cvTables().length;
          const cvNewIdx = cvTables().indexOf(cvCaretTable());
          ttPut(cvTableAt(cvNewIdx)?.rows[0]?.cells[0]);
          await wait(260);
          const cvCaretSame = cvCaretTable() === cvTableAt(cvNewIdx) ? 1 : 0;
          const cvNewRows = cvTableAt(cvNewIdx)?.rows.length ?? 0;
          const cvNewCols = cvTableAt(cvNewIdx)?.rows[0]?.cells.length ?? 0;
          const cvNewNoJia = cvTableAt(cvNewIdx) ? ((cvTableAt(cvNewIdx).textContent ?? "").indexOf("甲") < 0 ? 1 : 0) : 0;
          const cvInDisabled = cvChooser()?.disabled === true ? 1 : 0;
          const cvHomeValue = cvChooser()?.value ?? "";
          // 上面那几腿（标题行重复末行回绝）留下的浮层四秒后退场：先等它退场，「这一条有没有说话」才问得准。
          await wait(4300);
          const cvQuiet = ttNoticeSel();
          const cvPressHit = cvPress("tab");
          await wait(460);
          const cvNotice = ttNoticeSel();
          const cvTablesAfter = cvTables().length;
          const cvJiaStill = ttFind() ? 1 : 0;
          const cvTabsAfter = cvTabSpanN();
          const cvTabHidden = ttTab("表格工具")?.hidden === true ? 1 : 0;
          const cvDisabledAfter = cvChooser()?.disabled === true ? 1 : 0;
          ttUndoKey();
          await wait(460);
          const cvUndoTables = cvTables().length;
          const cvUndoRows = cvTableAt(cvNewIdx)?.rows.length ?? 0;
          const cvUndoCols = cvTableAt(cvNewIdx)?.rows[0]?.cells.length ?? 0;
          const cvUndoTabs = cvTabSpanN();
          const cvTabReShown = ttTab("表格工具")?.hidden === true ? 0 : 1;
          // 撤回来那一只表站着时才问修订那一条：这一条不改纸，表数与导引线得停在撤销后的读数上。
          revBtn()?.click();
          await wait(220);
          const cvTrackOn = revBtn()?.classList.contains("on") ? 1 : 0;
          ttPut(cvTableAt(cvNewIdx)?.rows[0]?.cells[0]);
          await wait(260);
          const cvRefusePress = cvPress("comma");
          await wait(420);
          const cvRefuseNotice = ttNoticeSel();
          const cvRefuseTables = cvTables().length;
          const cvRefuseTabs = cvTabSpanN();
          revBtn()?.click();
          await wait(200);
          const cvTrackEnd = revBtn()?.classList.contains("on") ? 1 : 0;
          ttUndoKey();
          await wait(460);
          const cvUndo2Tables = cvTables().length;
          const cvUndo2Jia = ttFind() ? 1 : 0;
          const cvUndo2Tabs = cvTabSpanN();
          await wait(4300);
          const cvUndoNotice = ttNoticeSel();
          // 「分布列」（CT_TableTools.kuip:54 的 DistributeColumns，idMso="TableColumnsDistribute"，RB_Adjust 那组
          // 第三枚、挂在 :58 那枚 TableAutoAdjust 的清单里）在装机这只盒子里量六条：①光标在「甲」那只表里时那一枚
          // 摆得出、按得动，光标退到表外时灰着（与上面那两格同一个开关）；②先拿自家的「单元格宽度」那一格把光标
          // 那一列摆宽一厘米——纸上那几列因此**不等**，这一击才有不等可摆；③按下去每一列摆到同一个宽（相邻两列
          // 差不到 0.5px），而整只盒一寸不挪：均分把那几列的和留在原地，那句 w:tblW 没理由重述；④立刻再按一次——
          // 这会儿每列已经一样宽，命令该回绝、把理由说在小条上、纸上一寸不挪；⑤一次 ⌘Z 退掉的是**第一次**那一击
          // （纸上重新不等、盒回到摆宽后那个数），这一条同时是「回绝那一次没进撤销栈」的证据：它要是留了一条空步，
          // ⌘Z 只会退掉那条空步、均分还挂在纸上；⑥再一次 ⌘Z 退掉宽度那一格，列宽与盒宽回到起手——整条撤销收尾，
          // 这一腿不许牵动盘上那件的字节账（下面 d1/d2 那两本的列数与网格和照旧，与本腿无关）。
          // 为什么量在「甲」那一只：新插的那只 2×2 两列本来就等宽，只量得到回绝那一条分支；而「甲」这一只在两趟里
          // 分别是 4 列与 5 列（上一腿前面那几腿各留一张、各插一列），所以这里断的全是**每趟自己那本算式**
          // （不等 → 齐 → 回绝 → 退两步回起手），两趟之间比的只有形状与那些算式，不是绝对列宽。
          // WPS 真机在悬光标上落哪几列、均分时动不动那句 w:tblW 均 未取证；这里断的是自家那条口径（落一整列、
          // 碰到的列之和留在原地）在装机的壳里落得下去、退得回来。件里那串字节的形状归 tests/doc-distribute.test.ts
          // 与 .scratch/col-distribute-dom.ts（它把同一份 writeDocx 的 XML 数出来）。
          const dcBtn = () => rb("分布列");
          const dcColsOf = () => [...(ttFind()?.rows[0]?.cells ?? [])]
            .map((c) => Math.round(c.getBoundingClientRect().width * 100) / 100);
          const dcSpread = (a) => (a.length ? Math.round((Math.max(...a) - Math.min(...a)) * 100) / 100 : -1);
          const dcJoin = (a) => a.join(",");
          ttPut(ttRow0Cell(0));
          await wait(260);
          const dcTrackStart = revBtn()?.classList.contains("on") ? 1 : 0;
          const dcBtnPresent = dcBtn() ? 1 : 0;
          const dcBtnIn = dcBtn()?.disabled === true ? 1 : 0;
          const dcCols0 = dcColsOf();
          const dcBox0 = ttBox();
          const dcEcho0 = ttSpin("宽度")?.value ?? "";
          const dcWidenTo = dcEcho0 !== "" && Number.isFinite(Number(dcEcho0))
            ? (Math.round((Number(dcEcho0) + 1) * 1000) / 1000).toFixed(3)
            : "5.292";
          const dcWidenHit = ttType(ttSpin("宽度"), dcWidenTo);
          await wait(420);
          const dcColsW = dcColsOf();
          const dcBoxW = ttBox();
          // 浮层四秒退场：上一击那句要是还挂着，「这一击没被回绝」就成了假绿。
          await wait(4300);
          const dcQuiet = ttNoticeSel();
          const dcPress = dcBtn() && dcBtn().disabled !== true ? (dcBtn().click(), 1) : 0;
          await wait(460);
          const dcColsAfter = dcColsOf();
          const dcBoxAfter = ttBox();
          const dcAfterNotice = ttNoticeSel();
          const dcPress2 = dcBtn() && dcBtn().disabled !== true ? (dcBtn().click(), 1) : 0;
          await wait(460);
          const dcRefuseNotice = ttNoticeSel();
          const dcColsRefuse = dcColsOf();
          const dcBoxRefuse = ttBox();
          ttUndoKey();
          await wait(460);
          const dcColsUndo1 = dcColsOf();
          const dcBoxUndo1 = ttBox();
          ttUndoKey();
          await wait(460);
          const dcColsUndo2 = dcColsOf();
          const dcBoxUndo2 = ttBox();
          const dcSpreadUndo2 = dcSpread(dcColsUndo2);
          const dcTablesEnd = cvTables().length;
          const dcOutPara = [...document.querySelectorAll("#view .para")].find((el) => !el.closest("table"));
          ttPut(dcOutPara);
          await wait(260);
          const dcOutDisabled = dcBtn()?.disabled === true ? 1 : 0;
          // 「分布行」（CT_TableTools.kuip:53 的 DistributeRows，idMso="TableRowsDistribute"，与 :54 那枚
          // DistributeColumns 同挂 RB_Adjust（:55）、照声明的序排在它**前面**）在装机这只盒子里量六条：
          // ①光标在带「甲」那只表里时那一枚摆得出、按得动，退到表外时灰着（与同组那三件同一个开关）；
          // ②状态栏那一格自报「厘米」——下面那一格的字面按 1 厘米 = 567 缇算，单位不是这一档时这条账自己作废；
          // ③先拿自家的「单元格高度」那一格把光标那一行按**纸上量到的**那条高再撑 600 缇，纸上那几行因此不等，
          //    这一击才有不等可摆；④按下去每一行摆到同一条高（相邻两行差不到 0.5px），整只盒的高一寸不挪：
          //    均分把那几行的和留在原地；⑤立刻再按一次——这会儿每行已经一样高，命令该回绝、把理由说在小条上、
          //    纸上一寸不挪；⑥一次 ⌘Z 退掉的是**第一次**那一击（纸上重新不等、盒回到撑高后那个数），这一条同时是
          //    「回绝那一次没进撤销栈」的证据：它要是留了一条空步，⌘Z 只会退掉那条空步、均分还挂在纸上；
          //    再一次 ⌘Z 退掉高度那一格，行高与盒高回到起手——整条撤销收尾，这一腿不许牵动盘上那件的字节账。
          // 为什么行高要从纸上量而不是从件里那句抄：授权 46 份 58 张表里 20 张一句 w:trHeight 都没说、按行数是
          // 580 说 / 118 没说（.scratch/row-distribute-audience.py），而说了的那 580 枚全是最小说法（atLeast 580 /
          // exact 0）——件里那句只是地板，「这一行现在多高」只有铺排答得出。这一件的几行起手一句没说，所以这一腿
          // 的算式用的全是本趟量到的像素；两趟之间比的只有形状与算式（行数是前面几腿留的账），不是绝对行高。
          // WPS 真机在悬光标上落哪几行、atLeast 追不上内容时它怎么收尾，均 未取证。
          const drBtn = () => rb("分布行");
          const drRowsOf = () => [...(ttFind()?.rows ?? [])]
            .filter((r) => !r.classList.contains("solo-sheet-gap-row") && !r.classList.contains("solo-repeat-header-row"))
            .map((r) => Math.round(r.getBoundingClientRect().height * 100) / 100);
          const drSpread = (a) => (a.length ? Math.round((Math.max(...a) - Math.min(...a)) * 100) / 100 : -1);
          const drJoin = (a) => a.join(",");
          const drSum = (a) => Math.round(a.reduce((s, h) => s + h, 0) * 100) / 100;
          const drBox = () => Math.round((ttFind()?.getBoundingClientRect().height ?? 0) * 100) / 100;
          ttPut(ttRow0Cell(0));
          await wait(260);
          const drTrackStart = revBtn()?.classList.contains("on") ? 1 : 0;
          const drBtnPresent = drBtn() ? 1 : 0;
          const drBtnIn = drBtn()?.disabled === true ? 1 : 0;
          const drUnit = document.querySelector(".solo-doc-status-unit")?.textContent ?? "";
          const drRows0 = drRowsOf();
          const drSum0 = drSum(drRows0);
          const drBox0 = drBox();
          const drWidenTo = (((drRows0[0] ?? 0) * 15 + 600) / 567).toFixed(3);
          const drWidenHit = ttType(ttSpin("高度"), drWidenTo);
          await wait(420);
          // 浮层四秒退场：上一击那句要是还挂着，「这一击没被回绝」就成了假绿。
          await wait(4300);
          const drQuiet = ttNoticeSel();
          const drRowsW = drRowsOf();
          const drSumW = drSum(drRowsW);
          const drBoxW = drBox();
          const drPress = drBtn() && drBtn().disabled !== true ? (drBtn().click(), 1) : 0;
          await wait(460);
          const drRowsAfter = drRowsOf();
          const drSumAfter = drSum(drRowsAfter);
          const drBoxAfter = drBox();
          const drAfterNotice = ttNoticeSel();
          const drPress2 = drBtn() && drBtn().disabled !== true ? (drBtn().click(), 1) : 0;
          await wait(460);
          const drRefuseNotice = ttNoticeSel();
          const drRowsRefuse = drRowsOf();
          const drBoxRefuse = drBox();
          ttUndoKey();
          await wait(460);
          const drRowsUndo1 = drRowsOf();
          const drSumUndo1 = drSum(drRowsUndo1);
          const drBoxUndo1 = drBox();
          ttUndoKey();
          await wait(460);
          const drRowsUndo2 = drRowsOf();
          const drBoxUndo2 = drBox();
          const drSpreadUndo2 = drSpread(drRowsUndo2);
          const drOutPara = [...document.querySelectorAll("#view .para")].find((el) => !el.closest("table"));
          ttPut(drOutPara);
          await wait(260);
          const drOutDisabled = drBtn()?.disabled === true ? 1 : 0;
          // 「字体」那一组（CT_TableTools.kuip:129 的 RB_Table_Font，组 tip TableToolFontGroup，夹在 RB_Adjust(:55)
          // 与 RB_Alignment(:156) 之间）在装机这只盒子里量四条：
          // ①这一组打进包之后仍摆得出、仍按 WPS 那一页的组序坐在第 5 组（左边那一组带着「分布行」、右边那一组是
          //   对齐那六枚），组里七枚控件一字排开（:133 的字体框 + :134 的字号 + :138/:141/:144 的粗/斜/线 +
          //   :147 的字体颜色 + :150 的突出显示）；两页那两枚字体框各自按 WPS 声明的数摆（本组 :133 的 112、
          //   开始页 CT_Home.kuip:110 的 130）——装机窗口就是产品开的那只，量的是那份渲染，不是再拼一遍探针页。
          // ②带子被这一组撑宽（自然宽 > 可视宽）之后，最后一枚「突出显示」靠行右那枚自家的箭头滚进可视框——
          //   断的是**可达**，不是「够宽」；不许为了让它露出来删掉组里任何一枚。
          // ③两页同源：表里这一枚「粗」按下去，开始页同名那枚跟着亮，再按回来一起灭，纸上那一枚 run 的字重跟着走。
          //   WPS 也是同一个 ksoCmd 挂两处（:138/:141/:144 的 idQ 就是 Bold/Italic/Underline，:147/:150 同指 FontColorPicker/
          //   HighlightColorPicker），所以「两页各说一句」在这枚上是要害，不是风格。
          // ④单位账与回显在同一腿里：字号挑 24 磅，纸上那枚 run 就是 32px（1 磅 = 4/3 像素，96dpi）；字体框打字进
          //   黑体之后开始页那格回显同一个名字（两页读同一份 markState）。挑的都是**选区盖住那几枚 run**——悬光标
          //   上字体框走的是待用格式（#89/#90），不落笔在这一格已有的字上。
          // 这一腿按完全部撤销，不牵动盘上那件的字节账（件里那六个槽的形状归 .scratch/table-tools-dom.ts 的 K 腿
          // 与 tests/doc-run-props.test.ts）；WPS 真机这一组按下去落不落 w:rFonts/w:sz、颜色档怎么排，均 未取证。
          await wait(4200);
          ttPut(ttRow0Cell(0));
          await wait(260);
          const tfTabShow = ttTab("表格工具") ? 1 : 0;
          ttTab("开始")?.click();
          await wait(240);
          const tfHomeFaceNow = () => document.querySelector("#ribbon .rbar[data-tab='home'] input.facebox");
          const tfFaceNow = () => document.querySelector("#ribbon .rbar[data-tab='table'] input.facebox");
          const tfHomeFace = tfHomeFaceNow();
          const tfHomeFaceW = tfHomeFace ? Math.round(tfHomeFace.getBoundingClientRect().width) : -1;
          const tfHomeBold0 = ttLitIn(ttHomeBtn("粗"));
          ttTab("表格工具")?.click();
          await wait(240);
          const tfBar = document.querySelector("#ribbon .rbar[data-tab='table']");
          const tfBarShown = tfBar?.hidden === true ? 0 : 1;
          const tfGroups = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] .grp")];
          const tfGroupIdx = tfGroups.findIndex((g) => [...g.querySelectorAll("button")].map((b) => b.textContent).join("|") === "粗|斜|线");
          const tfGroup = tfGroupIdx < 0 ? null : tfGroups[tfGroupIdx];
          const tfCtlN = tfGroup ? tfGroup.querySelectorAll(".facebox,.sizebox,.pick,.rb,.swatch").length : -1;
          const tfFaceW = tfFaceNow() ? Math.round(tfFaceNow().getBoundingClientRect().width) : -1;
          const tfPrevDist = tfGroupIdx > 0 && [...tfGroups[tfGroupIdx - 1].querySelectorAll("button")].some((b) => b.textContent === "分布行") ? 1 : 0;
          const tfNextAlign = tfGroupIdx >= 0 && tfGroupIdx + 1 < tfGroups.length &&
            [...tfGroups[tfGroupIdx + 1].querySelectorAll("button")].map((b) => b.textContent).join("|") === "↑|↕|↓|左|中|右" ? 1 : 0;
          // 字号那一格 #221 起是 ComboBox（input.sizebox 带 datalist），不再是 select；其余几格照旧。
          const tfSel = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] select, #ribbon .rbar[data-tab='table'] input.sizebox")].find((s) => s.title === label);
          const tfNatural = tfBar?.scrollWidth ?? -1;
          const tfVisible = tfBar?.clientWidth ?? -1;
          const tfRoverPresent = document.querySelector("#ribbon .rover") ? 1 : 0;
          const tfIn = (el) => {
            if (!el || !tfBar) return 0;
            const b = tfBar.getBoundingClientRect();
            const r = el.getBoundingClientRect();
            return r.left >= b.left - 0.5 && r.right <= b.right + 0.5 ? 1 : 0;
          };
          const tfHlStart = tfIn(tfSel("突出显示"));
          let tfRoverClicks = 0;
          const tfRoverEl = document.querySelector("#ribbon .rover");
          while (tfIn(tfSel("突出显示")) === 0 && tfRoverClicks < 8) {
            if (tfRoverEl?.hidden === true) tfBar.scrollLeft = tfBar.scrollWidth;
            else tfRoverEl?.click();
            tfRoverClicks++;
            await wait(180);
          }
          const tfHlReached = tfIn(tfSel("突出显示"));
          const tfRun = () => {
            const td = ttRow0Cell(0);
            return td?.querySelector(".para span") ?? td?.querySelector(".para") ?? null;
          };
          const tfSelectRun = () => {
            const el = tfRun();
            if (!el || !el.textContent) return 0;
            const r = document.createRange();
            r.selectNodeContents(el);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            return 1;
          };
          const tfBtn = (label) => ttRowEls().find((x) => x.textContent === label);
          // computed 的 font-family 串里带 CSS 那对双引号（"PingFang SC"），而烟测读数是一条 JSON——引号不抹掉，
          // 外面那条正则会把这一格截断成半串，「黑体在不在里面」就量不到了。
          const tfFam = (el) => cssOf(el, "fontFamily").replace(/"/g, "");
          const tfPress = (label) => {
            const b = tfBtn(label);
            if (!b || b.disabled === true) return 0;
            b.click();
            return 1;
          };
          const tfPick = (sel, v) => {
            if (!sel) return 0;
            sel.value = v;
            sel.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          const tfHomeSel = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] select, #ribbon .rbar[data-tab='home'] input.sizebox")].find((s) => s.title === label);
          // 两页那枚「字体颜色」色板（CT_Home.kuip:108 与 CT_TableTools.kuip:147 同一条 ksoCmd）。
          const tfColorBox = () => document.querySelector("#ribbon .rbar[data-tab='table'] input.swatch[title='字体颜色']");
          const tfHomeColorBox = () => document.querySelector("#ribbon .rbar[data-tab='home'] input.swatch[title='字体颜色']");
          const tfColorPick = (hex) => {
            const i = tfColorBox();
            if (!i || i.disabled === true) return 0;
            i.value = hex;
            i.dispatchEvent(new Event("input", { bubbles: true }));
            return 1;
          };
          // 色板那一格交回来的是 #rrggbb（小写），纸上 computed 的是 rgb(...)：两边折成同一个形状才敢比。
          // 这里不用正则：这段脚本活在一条模板串里，正则里的反斜杠会被外层字符串吃掉一半。
          const tfHex = (s) => {
            const t = String(s || "");
            if (t.indexOf("rgb") !== 0) return t;
            const open = t.indexOf("(");
            const parts = open < 0 ? [] : t.slice(open + 1).split(",");
            if (parts.length < 3) return t;
            const hx = (v) => { const n = parseInt(v, 10); return Number.isFinite(n) ? (n < 16 ? "0" : "") + n.toString(16) : "??"; };
            return ("#" + hx(parts[0]) + hx(parts[1]) + hx(parts[2])).toLowerCase();
          };
          const tfText0 = tfRun()?.textContent ?? "";
          await wait(200);
          const tfSelectHit = tfSelectRun();
          // 起手那几枚回显要在**摊开选区之后**取，与 tfClean() 同一口径：折叠光标报的是那一点的待用格式（#89/#90），
          // 摊开报的是那几枚 run 已有的字——起手按折叠取、结束时按摊开比，两句就不是同一个量，永远对不齐。
          const tfBold0Start = ttLitIn(tfBtn("粗"));
          const tfFaceHome0 = tfHomeFaceNow()?.value ?? "?";
          const tfFaceTable0 = tfFaceNow()?.value ?? "?";
          const tfSizeTable0 = tfSel("字号")?.value ?? "?";
          // 起手那两枚色板的读数：这一腿按下去之前，格子报的是这段字**穿着**的墨（件里没说过颜色时，那就是纸上画的那层）。
          const tfColorTable0 = tfColorBox()?.value ?? "?";
          const tfColorHome0 = tfHomeColorBox()?.value ?? "?";
          const tfColorCss0 = tfHex(cssOf(tfRun(), "color"));
          const tfWeight0 = cssOf(tfRun(), "fontWeight");
          const tfSize0 = cssOf(tfRun(), "fontSize");
          const tfFamily0 = tfFam(tfRun());
          // 四击之间隔 1000ms 不是「等渲染」，是**跟着产品自己那条合并规矩排步**：mark() 里 now-burst<900 就不落新快照
          // （doc-edit.ts:1822「Typing coalesces into one entry per pause」）。按 360ms 排时这四击在栈上只有 1~2 条，
          // 撤两步读到的是「按完不粗那一记」之后的中间态——纸上 400、回显全灭看着像起手，件里却留着那句 w:b val="0"。
          // 隔过 900ms 每一击各占一条，这一腿的撤销步数才等于它自己按下的击数，上限也才真的是「只撤自己」。
          const tfPressBold = tfPress("粗");
          await wait(1000);
          tfSelectRun();
          const tfBoldLitTable = ttLitIn(tfBtn("粗"));
          const tfBoldLitHome = ttLitIn(ttHomeBtn("粗"));
          const tfWeightBold = cssOf(tfRun(), "fontWeight");
          const tfPressBold2 = tfPress("粗");
          await wait(1000);
          tfSelectRun();
          const tfBoldOffTable = ttLitIn(tfBtn("粗"));
          const tfBoldOffHome = ttLitIn(ttHomeBtn("粗"));
          const tfWeightOff = cssOf(tfRun(), "fontWeight");
          const tfPressSize = tfPick(tfSel("字号"), "24");
          await wait(1000);
          const tfSizePx = cssOf(tfRun(), "fontSize");
          const tfSizeEchoTable = tfSel("字号")?.value ?? "?";
          const tfSizeEchoHome = tfHomeSel("字号")?.value ?? "?";
          const tfPressFace = ttType(tfFaceNow(), "黑体");
          // 这一眼也要隔过 900ms：下一击（色板）落在它之后，间隔不够就在栈上被合并成一条，撤销步数就对不上击数。
          await wait(1000);
          const tfFamilyPx = tfFam(tfRun());
          const tfFaceEchoHome = tfHomeFaceNow()?.value ?? "?";
          const tfFaceEchoTable = tfFaceNow()?.value ?? "?";
          const tfBoldLitAfterFace = ttLitIn(tfBtn("粗"));
          // 撤销的步数按这一腿自己按下的击数走（tfPresses），**不按条件早退**——纸与回显看不出「说过」与「没说过」的差别：
          // 撤两步落回的是「按完不粗那一记」之后，纸上 400、两页那枚灭、字号格显空，与起手一字不差，可件里那句
          // w:b w:val="0" 还在（run 上 false 算说过，ownRunStyle 的口径，tests/doc-run-props.test.ts:92 同形钉着）。
          // 上一版据此早退就漏了这一枚；每一击之间各隔过 1000ms（> mark() 那条 900ms 合并窗，doc-edit.ts:1822），
          // 栈上确实是击数那么多条，撤满才回到「这一腿第一击之前」。上限也就是击数：多撤一步就要吃掉「对齐」那一腿
          // 不落撤销、交壳写回盘上的那两句话，所以再拿 tfNeighbor0/tfNeighborEnd 当绊线——少撤则 tfCleanEnd 红，
          // 多撤则邻格那两句话红，两头都当场看得见，不靠调宽条件换绿。
          const tfBlur = () => { const a = document.activeElement; if (a && a.blur) a.blur(); };
          // 问回显前必须自己**把这一趟点亮**：回显挂在 document 的 selectionchange 上，而摊选区（tfSelectRun）
          // 只改 Range，不发这个事件——不发就没有人来问回显，两页那几格谁的也不写。
          // 「正在打字的那一格不写」那条礼节在这台烟测里依旧站着（show: false 的隐藏窗口，产品把焦点交回纸上的
          // page.focus() 在此不成），但这一腿碰不到它：第五击按的是色板的 input，从头到尾没把焦点交给任何一格，
          // 撤完那一站量出来是 tfFocusAfterUndo=BODY.（上一版四击、字体框还是 activeElement 时量出来是「表框」；
          // 两回都由 tfFocusAfterUndo 当场报出，读的是签名，不写进判据）。
          // 无头 Chrome 里同一句 page.focus() 是成的，所以这条签名只在装机腿出现——它问的是回显同不同源，
          // 不是「带着焦点那一枚该不该写」；那一格该不该写由 DOM 腿与产品自己的焦点走向定，不由这里改产品。
          const tfEcho = () => { tfBlur(); tfSelectRun(); document.dispatchEvent(new Event("selectionchange")); };
          // 第五击：在那枚色板里挑一枚蓝（走 input，与用户取完色同一趟），落的是 rPr 里那句 w:color。
          const tfPressColor = tfColorPick("#0000ff");
          await wait(1000);
          const tfColorPx = tfHex(cssOf(tfRun(), "color"));
          tfEcho();
          const tfColorEchoTable = tfColorBox()?.value ?? "?";
          const tfColorEchoHome = tfHomeColorBox()?.value ?? "?";
          const tfClean = () => {
            tfEcho();
            return cssOf(tfRun(), "fontWeight") === tfWeight0 && cssOf(tfRun(), "fontSize") === tfSize0 &&
              tfFam(tfRun()) === tfFamily0 &&
              tfHex(cssOf(tfRun(), "color")) === tfColorCss0 &&
              ttLitIn(tfBtn("粗")) === tfBold0Start && ttLitIn(ttHomeBtn("粗")) === tfHomeBold0 &&
              (tfHomeFaceNow()?.value ?? "") === tfFaceHome0 && (tfFaceNow()?.value ?? "") === tfFaceTable0 &&
              (tfSel("字号")?.value ?? "?") === tfSizeTable0 &&
              (tfColorBox()?.value ?? "?") === tfColorTable0 && (tfHomeColorBox()?.value ?? "?") === tfColorHome0;
          };
          const tfNeighbor = () => ttValign() + "|" + ttTextAlign();
          const tfNeighbor0 = tfNeighbor();
          const tfPresses = tfPressBold + tfPressBold2 + tfPressSize + tfPressFace + tfPressColor;
          let tfUndoSteps = 0;
          await wait(200);
          while (tfUndoSteps < tfPresses) {
            tfSelectRun();
            ttUndoKey();
            tfUndoSteps++;
            await wait(420);
          }
          const tfFocusAfterUndo = (() => {
            const a = document.activeElement;
            if (!a) return "无";
            if (a === tfFaceNow()) return "表框";
            return (a.tagName || "?") + "." + (a.className || "");
          })();
          const tfCleanEnd = tfClean() ? 1 : 0;
          // 红的时候要能一眼看出是哪一枚没回起手，所以把 tfClean() 那五枚逐条也报出来（同一次读法，不另起口径）。
          tfEcho();
          const tfEndBoldTable = ttLitIn(tfBtn("粗"));
          const tfEndBoldHome = ttLitIn(ttHomeBtn("粗"));
          const tfEndFaceHome = tfHomeFaceNow()?.value ?? "?";
          const tfEndFaceTable = tfFaceNow()?.value ?? "?";
          const tfEndSizeEcho = tfSel("字号")?.value ?? "?";
          const tfEndColorTable = tfColorBox()?.value ?? "?";
          const tfEndColorHome = tfHomeColorBox()?.value ?? "?";
          const tfNeighborEnd = tfNeighbor();
          tfSelectRun();
          const tfWeightEnd = cssOf(tfRun(), "fontWeight");
          const tfFamilyEnd = tfFam(tfRun());
          const tfSizeEnd = cssOf(tfRun(), "fontSize");
          const tfColorEnd = tfHex(cssOf(tfRun(), "color"));
          ttCmd = {
            ttTabPresent, ttTabHidden, ttBarHidden, ttBarOpen, ttBtns, ttHit: ttHit ? 1 : 0, ttDisabled, ttNotice,
            ttPxBefore, ttPxAfter, ttWrittenBefore, ttWrittenAfter, ttColsBefore, ttColsAfter, ttTabBackHidden, ttHomeOn, ttInTableAgain,
            ttDelHit: ttDelBtn ? 1 : 0, ttDelDisabled, ttJiaBefore, ttJiaGone, ttJiaBack, ttJiaLinesBefore, ttOutBefore, ttOutAfter, ttOutBack,
            ttDelNotice, ttPages, ttPageKids, ttTablesInfo,
            ttTrackBefore, ttTrackOff, ttTrackOn, ttDel2Hit: ttDel2 ? 1 : 0, ttDel2Disabled, ttTrackTable, ttStruck, ttTrackNotice,
            ttTrackUndoTable, ttStruckBack, ttTrackEnd,
            ttChooserHit, ttChooserDisabled, ttNoticeBeforeUp, ttBottomBefore, ttUpPick, ttUpNotice, ttBottomAfterUp,
            ttLeftPick, ttTargetText, ttRow0Before, ttTextBefore, ttRow0AfterLeft, ttTextAfterLeft, ttLeftNotice, ttRow0Back, ttTextBack,
            ttCellTrackOn, ttPaneRead, ttCellDelBefore, ttSignPick, ttRow0Signed, ttTextSigned, ttCellDelSigned, ttSignedNotice,
            ttCellDelAfterUndo, ttRow0AfterUndo, ttCellTrackEnd,
            ttSizeHit, ttHeightEcho0, ttWidthEcho0, ttRowHBefore, ttHeightPress, ttRowHAfter, ttHeightEcho1,
            ttRowHUndone, ttHeightEcho2, ttBoxBeforeW, ttCol0BeforeW, ttCol1BeforeW, ttWidthPress, ttBoxAfterW,
            ttCol0AfterW, ttCol1AfterW, ttWidthEcho1, ttBoxUndoneW, ttHeightEchoOut, ttWidthEchoOut,
            ttOutPress, ttOutNotice, ttHeightEchoRefused, ttRowHRefused, ttBoxRefused,
            ttAlignTrack, ttAlignTabShow, ttAlignBarShown, ttAlignGroupIdx, ttAlignGroupN, ttAlignSpinN,
            ttAlignCol, ttRow0N: ttRow0N(), ttAlignCellText, ttStartLit, ttStartHome, ttValign0, ttTextAlign0, ttNoticeV0,
            ttPressMid, ttDisabledAfterV, ttLitAfterV, ttHomeAfterV, ttValignAfterV, ttTextAlignAfterV, ttNoticeAfterV,
            ttRefuseV, ttNoticeRefuseV, ttLitRefuseV, ttValignRefuseV, ttLitUndoV, ttValignUndoV,
            ttPressCenter, ttLitAfterH, ttHomeAfterH, ttTextAlignAfterH, ttValignAfterH, ttNoticeAfterH,
            ttRefuseH, ttNoticeRefuseH, ttLitRefuseH, ttTextAlignRefuseH, ttLitUndoH, ttTextAlignUndoH,
            ttLitEnd, ttHomeEnd, ttValignEnd, ttTextAlignEnd, ttNoticeEnd, ttPrevLit, ttCursorBackInCell,
            ttAlignOutLit, ttAlignOutDisabled, ttHomeOutDisabled,
            glBtnPresent: glBtn() ? 1 : 0, glBtnLabel: glBtn()?.textContent ?? "", glGroupIdx: glGroupIdx,
            glStartView, glStartLit, glStartShadow, glStartInk, glStartEdges: glStartAudit.edges, glStartBoth: glStartAudit.both, glStartNeither: glStartAudit.neither, glStartRect,
            glOffView, glOffLit, glOffShadow, glOffInk, glOffEdges: glOffAudit.edges, glOffBoth: glOffAudit.both, glOffNeither: glOffAudit.neither, glOffRect, glOffNotice,
            glBackView, glBackLit, glBackShadow, glBackRect, glDirty0, glDirty1,
            tpGrpLabels, tpBtnPresent, tpBtnLabel, tpOpen, tpTitle, tpRows,
            tpEchoAlign, tpEchoIndent, tpEchoFixed, tpEchoMargin, tpError,
            tpBefore, tpAfter, tpWritten, tpUndoBack, tpReopen, tpCancelGone, tpCancelClean, tpDialogsLeft,
            rhBtnPresent, rhStartDisabled, rhStartOn, rhStartStamped, rhStartClones,
            rhAfterOn, rhAfterStamped, rhAfterClones, rhNotice, rhUndoStamped, rhUndoOn,
            rhLastStamped, rhLastNotice, rhTrackOn, rhItemsBefore, rhItemsSigned, rhStampedSigned,
            rhItemsUndone, rhStampedUndone, rhTrackEnd,
            cvChooserHit, cvOptions, cvOutDisabled, cvHomeValueOut, cvN0, cvTabs0, cvTrackStart,
            cvInsertHit, cvPickHit, cvTablesIns, cvNewIdx, cvCaretSame, cvNewRows, cvNewCols, cvNewNoJia,
            cvInDisabled, cvHomeValue, cvQuiet, cvPressHit, cvNotice, cvTablesAfter, cvJiaStill, cvTabsAfter,
            cvTabHidden, cvDisabledAfter, cvUndoTables, cvUndoRows, cvUndoCols, cvUndoTabs, cvTabReShown,
            cvTrackOn, cvRefusePress, cvRefuseNotice, cvRefuseTables, cvRefuseTabs, cvTrackEnd,
            cvUndo2Tables, cvUndo2Jia, cvUndo2Tabs, cvUndoNotice,
            dcBtnPresent, dcBtnIn, dcOutDisabled, dcTrackStart, dcEcho0, dcWidenTo, dcWidenHit,
            dcCols0: dcJoin(dcCols0), dcSpread0: dcSpread(dcCols0), dcBox0,
            dcColsW: dcJoin(dcColsW), dcSpreadW: dcSpread(dcColsW), dcBoxW,
            dcQuiet, dcPress, dcColsAfter: dcJoin(dcColsAfter), dcSpreadAfter: dcSpread(dcColsAfter), dcBoxAfter, dcAfterNotice,
            dcPress2, dcRefuseNotice, dcColsRefuse: dcJoin(dcColsRefuse), dcBoxRefuse,
            dcColsUndo1: dcJoin(dcColsUndo1), dcSpreadUndo1: dcSpread(dcColsUndo1), dcBoxUndo1,
            dcColsUndo2: dcJoin(dcColsUndo2), dcSpreadUndo2, dcBoxUndo2, dcTablesEnd,
            drBtnPresent, drBtnIn, drOutDisabled, drTrackStart, drUnit, drWidenHit, drWidenTo,
            drRows0: drJoin(drRows0), drSpread0: drSpread(drRows0), drSum0, drBox0,
            drRowsW: drJoin(drRowsW), drSpreadW: drSpread(drRowsW), drSumW, drBoxW,
            drQuiet, drPress, drRowsAfter: drJoin(drRowsAfter), drSpreadAfter: drSpread(drRowsAfter), drSumAfter, drBoxAfter, drAfterNotice,
            drPress2, drRefuseNotice, drRowsRefuse: drJoin(drRowsRefuse), drBoxRefuse,
            drRowsUndo1: drJoin(drRowsUndo1), drSpreadUndo1: drSpread(drRowsUndo1), drSumUndo1, drBoxUndo1,
            drRowsUndo2: drJoin(drRowsUndo2), drSpreadUndo2, drBoxUndo2,
            tfTabShow, tfBarShown, tfGroupIdx, tfCtlN, tfFaceW, tfHomeFaceW, tfPrevDist, tfNextAlign,
            tfNatural, tfVisible, tfRoverPresent, tfHlStart, tfRoverClicks, tfHlReached,
            tfFaceHome0, tfFaceTable0, tfSizeTable0, tfText0, tfSelectHit, tfWeight0, tfSize0, tfFamily0,
            tfBold0Start, tfHomeBold0, tfPressBold, tfBoldLitTable, tfBoldLitHome, tfWeightBold,
            tfPressBold2, tfBoldOffTable, tfBoldOffHome, tfWeightOff,
            tfPressSize, tfSizePx, tfSizeEchoTable, tfSizeEchoHome,
            tfPressFace, tfFamilyPx, tfFaceEchoHome, tfFaceEchoTable, tfBoldLitAfterFace,
            tfPresses, tfUndoSteps, tfFocusAfterUndo, tfCleanEnd, tfEndBoldTable, tfEndBoldHome, tfEndFaceHome, tfEndFaceTable, tfEndSizeEcho,
            tfPressColor, tfColorTable0, tfColorHome0, tfColorCss0, tfColorPx, tfColorEchoTable, tfColorEchoHome,
            tfEndColorTable, tfEndColorHome, tfColorEnd,
            tfWeightEnd, tfFamilyEnd, tfSizeEnd, tfNeighbor0, tfNeighborEnd,
            ttDirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0,
          };
        }
        // #214：「表格属性」这一族在**装机这只壳**里第一次落笔并写回盘上。#197 那一趟（上面 tpBefore/tpAfter 那几条）
        // 落的是左对齐 + 缩进、量完整条撤销——它证的是「按得动、纸跟着动、撤得回」，不是「落得到盘上」；而 #213 把
        // 居中那一支换成 calc 之后，「居中 + 缇」这条**从没被点过**：授权 58 张表里 14 枚说居中、8 枚两句同说，
        // 落笔面只走过 #213 之前那条分得开的档。这一腿按件名「写盘」认领自己那三张，别家的字节账一律不牵动。
        // 三张各占一种起始态，问三种形状（每行问两半：表那一句在 w:tblPr 里，#216 新接的行那一句 w:trPr/w:jc 在行自己身上）：
        // 摆1 表与行两句都没说 → 表点居中 + 0.5 厘米、行点居中（盘上凭空多出三枚，纸上走 calc 那一支、只有光标那一行挪）；
        // 摆2 表只说居中 → 原样确定（一个字节也不许多写，脏字也不许多亮，表那句也不许抄给行）；
        // 摆3 表说居中 + 288、末行说 trHeight 600 + 居中 → 光标放末行，左缩进填成 0（盘上必须是**明写的 0**，不是删掉——
        // 「说了 0」与「没说」画到同一处，文件里却是两句话：授权 46 份里 w:tblInd 说 0 的有 30 枚，见 §154），
        // 行那句选「左对齐」（= 摘掉 w:jc，不许补一句 val=left；授权 190 枚行对齐里 left 一枚也没明写过，见 §157），
        // 同座的兄弟孩子 w:trHeight 得留在原位——摘一句不许顺手把邻座也搬走。
        // 落完按壳自己的 solo:save 写回；第二趟直接开上一趟写回的那份——起手摆出来的就该是上一趟落下的那两枚，
        // 且再点一次同样的值不再动字节。这一段整篇在模板字符串里：不写类型、不写反引号、正则不带转义。
        const tpwDoc = (document.getElementById("title")?.textContent ?? "").includes("写盘");
        let tablePropsWriteLeg = null;
        if (tpwDoc) {
          // 起手先把「修订」按回关着：上面 #149/#178 那一族把跟踪点开着交过来，开着落这一击会顺带写一枚
          // w:tblPrChange（实测第 1 趟盘上那五枚孩子就是 tblW/jc/tblInd/tblLayout/tblPrChange）。那是**修订**那一族
          // 的形状，不是摆位这一族要问的；两条账混在一张表上，反读就说不清哪一句是谁写的。关掉由产品自己那枚开关来
          // （与 #190 那一腿同一条路），量完按回原样交给后面的腿。开着落笔那一支由 #149 的单测与 #197 那一腿管。
          const tpwTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const tpwTrackOn = () => (tpwTrackBtn()?.classList.contains("on") ? 1 : 0);
          const track0 = tpwTrackOn();
          if (track0) {
            tpwTrackBtn()?.click();
            await wait(320);
          }
          const trackOff = tpwTrackOn();
          // 内容盒按**这张表自己那张纸**算，不按第一张：前面那些腿会往正文里插表插字，这三张可能被推到第二张纸上，
          // 拿第一张的盒沿去量第二张上的矩形，读回来的是两张纸的差，不是摆位。
          const tpwBox = (el) => {
            const p = el ? el.closest(".page") : null;
            if (!p) return null;
            const z = zPaper(p);
            const cs = getComputedStyle(p);
            const r = p.getBoundingClientRect();
            return { z: z, left: r.left / z + parseFloat(cs.paddingLeft || "0"), right: r.right / z - parseFloat(cs.paddingRight || "0") };
          };
          const tpwTable = (needle) => [...document.querySelectorAll("#view .page table.doc-table")].find((t) => (t.textContent ?? "").includes(needle));
          const tpwRead = (needle) => {
            const t = tpwTable(needle);
            const b = tpwBox(t);
            if (!t || !b) return null;
            const r = t.getBoundingClientRect();
            return { css: t.getAttribute("style") ?? "", gL: round1(r.left / b.z - b.left), gR: round1(b.right - r.right / b.z), w: round1(r.width / b.z), boxW: round1(b.right - b.left) };
          };
          // #216：行那一句（w:trPr/w:jc）落在**那一行自己**那枚 tr 上，不在表上——所以逐行读，表那一句的读数管不到它。
          // 行内样式同时带着兄弟孩子那一句（摆3 末行的「height:40px」= 600 缇 ÷ 15）：摘掉居中时它得留在原处。
          const tpwRowRead = (needle) => {
            const t = tpwTable(needle);
            if (!t) return null;
            return [...t.querySelectorAll("tr")].map((tr) => ({
              css: tr.getAttribute("style") ?? "",
              align: getComputedStyle(tr).textAlign,
            }));
          };
          // 光标放回那张表第一格（或指定那一行的第一格）：命令按选区派发，光标不在表里时「属性」那枚是灰的——那不该由这一腿绕过。
          // 照 #190 那一腿的做法（selectNodeContents + collapse），不抓格子里那个 span 的第一个孩子：
          // 格子里那一串是壳自己画的，画成什么形状不该由这一腿来假设。
          const tpwPut = (needle, rowIdx) => {
            const rows = tpwTable(needle)?.querySelectorAll("tr");
            const cell = rows?.[rowIdx ?? 0]?.querySelector("td");
            if (!cell) return 0;
            const rg = document.createRange();
            rg.selectNodeContents(cell);
            rg.collapse(true);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(rg);
            cell.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const tpwBtn = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] button")].find((b) => b.textContent === "属性");
          const tpwDlg = () => document.querySelector(".solo-tp-dialog");
          const tpwField = (n) => tpwDlg()?.querySelector("[name='" + n + "']") ?? null;
          // 单位从对话框自己那行标签里读（「左缩进（厘米）」）：壳的测量单位不落在文件里，写死的期望会随档位漂。
          const tpwUnit = () => {
            const lab = tpwField("tp-indent")?.parentElement;
            const t = lab ? String(lab.textContent) : "";
            const a = t.indexOf("（");
            const e = t.indexOf("）");
            return a >= 0 && e > a ? t.slice(a + 1, e) : "";
          };
          const tpwSet = (el, v, kind) => {
            if (!el) return 0;
            el.value = v;
            el.dispatchEvent(new Event(kind, { bubbles: true }));
            return 1;
          };
          const tpwOpen = async (needle, rowIdx) => {
            const put = tpwPut(needle, rowIdx);
            await wait(220);
            const bar = document.querySelector("#ribbon .rbar[data-tab='table']");
            const shown = bar?.hidden === false ? 1 : 0;
            if (bar?.hidden === true) [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === "表格工具")?.click();
            await wait(200);
            const btn = tpwBtn();
            const disabled = btn?.disabled === true ? 1 : 0;
            const hit = btn ? (btn.click(), 1) : 0;
            await wait(300);
            return { put: put, shown: shown, btnPresent: btn ? 1 : 0, disabled: disabled, hit: hit, open: tpwDlg()?.open === true ? 1 : 0, title: tpwDlg()?.querySelector("h2")?.textContent ?? "" };
          };
          const tpwEcho = () => ({
            align: tpwField("tp-align")?.value ?? "?",
            indent: tpwField("tp-indent")?.value ?? "?",
            fixed: tpwField("tp-fixed")?.checked === true ? 1 : 0,
            margin: ["top", "left", "bottom", "right"].map((s) => tpwField("tp-m-" + s)?.value ?? "?").join(","),
            unit: tpwUnit(),
            // #216：行那一句是**另一格**（tp-row-align），读的是光标所在那一行自己那句，不被表那句带着走。
            rowAlign: tpwField("tp-row-align")?.value ?? "?",
            rowOpts: tpwField("tp-row-align")?.options?.length ?? 0,
            // #222：「允许跨页断行」勾读的是那一行有没有 w:cantSplit（没写 = 允许 = 勾着）。
            rowBreak: tpwField("tp-row-allow-break")?.checked === true ? 1 : 0,
          });
          // 勾这一类不认 .value，要落的是 checked 再发一次 change（与固定列宽那枚同一条路）。
          const tpwCheck = (el, on) => {
            if (!el) return 0;
            el.checked = on;
            el.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          const tpwApply = () => {
            const b = tpwDlg()?.querySelector(".solo-tp-apply");
            return b ? (b.click(), 1) : 0;
          };
          const dirtyNow = () => ((document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0);
          const tpwSteps = [];
          // 摆1：表与行两句都没说 → 表那半点居中 + 0.5 厘米，行那半点居中：一次「确定」落两处，
          // 盘上凭空多出三枚（tblPr 的 jc/tblInd 加行自己那枚 trPr/jc），纸上只有光标那一行挪。
          {
            const o = await tpwOpen("摆1");
            const before = tpwRead("摆1");
            const rowsBefore = tpwRowRead("摆1");
            const echo = tpwEcho();
            const dirtyBefore = dirtyNow();
            const setAlign = tpwSet(tpwField("tp-align"), "center", "change");
            const setIndent = tpwSet(tpwField("tp-indent"), "0.5", "input");
            const setRow = tpwSet(tpwField("tp-row-align"), "center", "change");
            await wait(140);
            const err = tpwDlg()?.querySelector(".solo-tp-error")?.textContent ?? "";
            const applied = tpwApply();
            await wait(420);
            tpwSteps.push({ name: "摆1", ...o, echo: echo, dirtyBefore: dirtyBefore, setAlign: setAlign, setIndent: setIndent, setRow: setRow, err: err, applied: applied, gone: tpwDlg() ? 0 : 1, before: before, rowsBefore: rowsBefore, rows: tpwRowRead("摆1"), after: tpwRead("摆1"), dirtyAfter: dirtyNow() });
          }
          // 摆2：表只说居中、行没说 → 一格不碰，直接按「确定」：表那句不许抄给行。
          {
            const o = await tpwOpen("摆2");
            const before = tpwRead("摆2");
            const rowsBefore = tpwRowRead("摆2");
            const echo = tpwEcho();
            const dirtyBefore = dirtyNow();
            const applied = tpwApply();
            await wait(420);
            tpwSteps.push({ name: "摆2", ...o, echo: echo, dirtyBefore: dirtyBefore, err: "", applied: applied, gone: tpwDlg() ? 0 : 1, before: before, rowsBefore: rowsBefore, rows: tpwRowRead("摆2"), after: tpwRead("摆2"), dirtyAfter: dirtyNow() });
          }
          // 摆3：表说居中 + 288、末行说 trHeight 600 + 居中 → 光标放**末行**，只把左缩进改成 0、行那句选「左对齐」（= 摘掉）。
          {
            const o = await tpwOpen("摆3", 1);
            const before = tpwRead("摆3");
            const rowsBefore = tpwRowRead("摆3");
            const echo = tpwEcho();
            const dirtyBefore = dirtyNow();
            const setIndent = tpwSet(tpwField("tp-indent"), "0", "input");
            const setRow = tpwSet(tpwField("tp-row-align"), "left", "change");
            // #222：末行本来没说 w:cantSplit（勾着 = 允许断行）；勾掉这一下要在盘上落一枚空孩子。
            const setBreak = tpwCheck(tpwField("tp-row-allow-break"), false);
            await wait(140);
            const err = tpwDlg()?.querySelector(".solo-tp-error")?.textContent ?? "";
            const applied = tpwApply();
            await wait(420);
            tpwSteps.push({ name: "摆3", ...o, echo: echo, dirtyBefore: dirtyBefore, setIndent: setIndent, setRow: setRow, setBreak: setBreak, err: err, applied: applied, gone: tpwDlg() ? 0 : 1, before: before, rowsBefore: rowsBefore, rows: tpwRowRead("摆3"), after: tpwRead("摆3"), dirtyAfter: dirtyNow() });
          }
          // #222 续：再开一次只读起手——上一击写进模型的那枚 w:cantSplit 得让「允许跨页断行」灭掉。
          // 「取消」不动模型也不写盘：盘上那件就是摆3落下的样子，正好交给本节末尾的 solo:save 与外部反读。
          {
            const o = await tpwOpen("摆3", 1);
            const echo = tpwEcho();
            tpwDlg()?.querySelector(".solo-tp-cancel")?.click();
            await wait(220);
            tpwSteps.push({ name: "摆3-回显", ...o, echo: echo, gone: tpwDlg() ? 0 : 1 });
          }
          fire("solo:save");
          await wait(1000);
          // 交还原样：跟踪是上面那一族点开着交过来的，量完按回去，后面的腿才仍拿到它自己那一档。
          if (track0) {
            tpwTrackBtn()?.click();
            await wait(320);
          }
          const trackBack = tpwTrackOn();
          tablePropsWriteLeg = {
            unit: tpwSteps[0]?.echo?.unit ?? "",
            trackBefore: track0,
            trackOff: trackOff,
            trackBack: trackBack,
            steps: tpwSteps,
            savedDirty: dirtyNow(),
            dialogsLeft: document.querySelectorAll(".solo-tp-dialog").length,
            colW: round1((() => { const b = tpwBox(tpwTable("摆1")); return b ? b.right - b.left : 0; })()),
            // 写回之后再看一眼那三张：后面那些腿（插字/插表）动的是别处，这一族的行内样式得到这一刻仍该是刚才落的那句。
            // 第二趟的起手与这一串同一处，才是「壳把自己写回的那份读回来、照它摆」——比步骤里那次紧跟的读数更晚、更硬。
            final: ["摆1", "摆2", "摆3"].map((n) => {
              const r = tpwRead(n);
              return { name: n, css: r?.css ?? "", gL: r?.gL ?? null, gR: r?.gR ?? null, w: r?.w ?? null, boxW: r?.boxW ?? null, rows: tpwRowRead(n) };
            }),
            pages: document.querySelectorAll("#view .page").length,
          };
        }
        // #219：测量单位三档（毫米/英寸/磅）——「表格属性 → 左缩进」那格在每一档下摆什么字面、填进去的数按哪一档的
        // 缇数落笔。单位走产品自己那条路：状态栏「测量单位」那格 → 「选项」对话框 → 确定（单位只活在这个窗口里，
        // 量完按回厘米交出去）。三张表各起始一句 w:tblInd（288 / −129 / 907 缇），各填一种新缇数：纸上挪的距离与
        // 盘上那句都得对上；第四张一句都没说——格子必须空着（「没说」不是 0），按「取消」一个字节也不许动。
        // 这一段整篇在模板字符串里：不写反引号、不写美元花括号。
        const tuDoc = (document.getElementById("title")?.textContent ?? "").includes("表格单位");
        let tableUnitsLeg = null;
        if (tuDoc) {
          // 修订先按回关着（与 #214 同一条路）：别的腿把跟踪开着交过来，开着落这一击会在 tblPr 里多一枚
          // w:tblPrChange，盘上就出现两句 w:tblInd（新值与记录里的旧值），这一族的字节账说不清是谁写的。
          // 量完按回原样交给后面的腿。
          const tuTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const tuTrackOn = () => (tuTrackBtn()?.classList.contains("on") ? 1 : 0);
          const tuTrack0 = tuTrackOn();
          if (tuTrack0) {
            tuTrackBtn()?.click();
            await wait(320);
          }
          const tuTrackOff = tuTrackOn();
          const tuTable = (needle) => [...document.querySelectorAll("#view .page table.doc-table")].find((t) => (t.textContent ?? "").includes(needle));
          const tuGap = (needle) => {
            const t = tuTable(needle);
            const p = t ? t.closest(".page") : null;
            if (!t || !p) return null;
            const z = zPaper(p);
            const cs = getComputedStyle(p);
            const r = p.getBoundingClientRect();
            const left = r.left / z + parseFloat(cs.paddingLeft || "0");
            const box = t.getBoundingClientRect();
            return { l: round1(box.left / z - left), w: round1(box.width / z) };
          };
          const tuCell = () => document.querySelector(".solo-doc-status-unit")?.textContent ?? "";
          const tuPickUnit = async (u) => {
            document.querySelector(".solo-doc-status-unit")?.click();
            await wait(240);
            const d = document.querySelector(".solo-op-dialog");
            const open = d && d.open ? 1 : 0;
            const radio = d ? d.querySelector("input[name='op-unit'][value='" + u + "']") : null;
            const pre = radio ? (radio.checked ? 1 : 0) : 0;
            radio?.click();
            await wait(80);
            d?.querySelector(".solo-op-apply")?.click();
            await wait(300);
            return { open: open, radio: radio ? 1 : 0, pre: pre, closed: document.querySelector(".solo-op-dialog") ? 0 : 1, cell: tuCell() };
          };
          const tuPut = (needle) => {
            const cell = tuTable(needle)?.querySelector("td");
            if (!cell) return 0;
            const rg = document.createRange();
            rg.selectNodeContents(cell);
            rg.collapse(true);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(rg);
            cell.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const tuDlg = () => document.querySelector(".solo-tp-dialog");
          const tuField = (n) => tuDlg()?.querySelector("[name='" + n + "']") ?? null;
          const tuLabelUnit = () => {
            const lab = tuField("tp-indent")?.parentElement;
            const t = lab ? String(lab.textContent) : "";
            const a = t.indexOf("（");
            const e = t.indexOf("）");
            return a >= 0 && e > a ? t.slice(a + 1, e) : "";
          };
          const tuDirty = () => ((document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0);
          const tuOpen = async (needle) => {
            const put = tuPut(needle);
            await wait(240);
            const bar = document.querySelector("#ribbon .rbar[data-tab='table']");
            if (bar?.hidden === true) [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === "表格工具")?.click();
            await wait(200);
            const btn = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] button")].find((b) => b.textContent === "属性");
            const disabled = btn?.disabled === true ? 1 : 0;
            if (btn) btn.click();
            await wait(300);
            return { put: put, btnPresent: btn ? 1 : 0, disabled: disabled, open: tuDlg()?.open === true ? 1 : 0 };
          };
          const tuSet = (el, v) => {
            if (!el) return 0;
            el.value = v;
            el.dispatchEvent(new Event("input", { bubbles: true }));
            return 1;
          };
          const tuHit = (sel) => {
            const b = tuDlg()?.querySelector(sel);
            return b ? (b.click(), 1) : 0;
          };
          const tuSteps = [];
          const tuPlan = [["毫米表", "mm", "毫米", "20"], ["英寸表", "in", "英寸", "0.75"], ["磅表", "pt", "磅", "36"]];
          for (const [name, u, cn, literal] of tuPlan) {
            const pick = await tuPickUnit(u);
            const g0 = tuGap(name);
            const o = await tuOpen(name);
            const echo = { unit: tuLabelUnit(), indent: tuField("tp-indent")?.value ?? "?" };
            const dirty0 = tuDirty();
            const set = tuSet(tuField("tp-indent"), literal);
            await wait(140);
            const err = tuDlg()?.querySelector(".solo-tp-error")?.textContent ?? "";
            const applied = tuHit(".solo-tp-apply");
            await wait(460);
            tuSteps.push({ name: name, unit: u, cn: cn, literal: literal, pick: pick, open: o, echo: echo, set: set, err: err, applied: applied, gone: tuDlg() ? 0 : 1, g0: g0, g1: tuGap(name), dirty0: dirty0, dirty1: tuDirty() });
          }
          {
            const pick = await tuPickUnit("cm");
            const g0 = tuGap("无缩进表");
            const o = await tuOpen("无缩进表");
            const echo = { unit: tuLabelUnit(), indent: tuField("tp-indent")?.value ?? "?" };
            const dirty0 = tuDirty();
            const cancelled = tuHit(".solo-tp-cancel");
            await wait(320);
            tuSteps.push({ name: "无缩进表", unit: "cm", cn: "厘米", literal: "", pick: pick, open: o, echo: echo, set: 0, err: "", applied: 0, cancelled: cancelled, gone: tuDlg() ? 0 : 1, g0: g0, g1: tuGap("无缩进表"), dirty0: dirty0, dirty1: tuDirty() });
          }
          fire("solo:save");
          await wait(1000);
          if (tuTrack0) {
            tuTrackBtn()?.click();
            await wait(320);
          }
          tableUnitsLeg = { unitCellEnd: tuCell(), trackBefore: tuTrack0, trackOff: tuTrackOff, trackBack: tuTrackOn(), steps: tuSteps, dialogsLeft: document.querySelectorAll(".solo-tp-dialog, .solo-op-dialog").length, savedDirty: tuDirty() };
        }
        // #215：「允许西文在单词中间换行」那一勾在**装机这只壳**里走一趟并落回盘上。DOM 探针（.scratch/wrap-dom.ts）
        // 证的是同一个 openParaSpacing/setParagraphSpacing 在浏览器页里的形状；只有装机壳能给的那一条是：
        // 壳把自己写回盘上的那份再读回来、照它摆，且同一句话第二次说时不再多写一个字节。三摆问三种起始态：
        // 摆1 没说这一枚 → 勾一下、确定：纸上从整词换行走成拦腰断（首行铺到内宽），盘上裸写一枚 <w:wordWrap/>；
        // 摆2 自己说了允许 → 一格不碰直接确定：那一枚留在原位，别的字节也不许多写；
        // 摆3 从样式链借来「允许」→ 摆出来是勾着的（对话框说的是页面上正在画的那一份），但段落不认领，
        //      原样确定之后那一句 w:pPr 里仍不许多出这一枚——把借来的话抄成自己说的，就是替文件签了一句它没说过的。
        // 勾选只在「此刻没勾」时落一次，所以第二趟开第一趟写回的那份时，三摆都走「原样确定」那条路。
        // 跟踪改动先按回关着（与 #214 那一腿同一条路）：开着落这一击会顺带写一枚 w:pPrChange，那是修订那一族的
        // 形状，摆位这一族的字节账问不清是谁写的；量完按回原样交给后面的腿。
        const wrapDoc = (document.getElementById("title")?.textContent ?? "").includes("换行");
        let wrapLeg = null;
        if (wrapDoc) {
          const wrapTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const wrapTrackOn = () => (wrapTrackBtn()?.classList.contains("on") ? 1 : 0);
          const wrapTrack0 = wrapTrackOn();
          if (wrapTrack0) {
            wrapTrackBtn()?.click();
            await wait(320);
          }
          const wrapTrackOff = wrapTrackOn();
          const wrapPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          // 断法量在带那个长西文词的那一枚 span 上：逐行矩形由 Range.getClientRects() 报，行数=词被挪走后多出来的那一行，
          // 首行宽=那一行铺到哪儿（整词换行时停在词前，拦腰断时铺到内宽）。页内宽按这张纸自己的盒沿算（同 #214 那条）。
          const wrapLineRead = (needle) => {
            const el = wrapPara(needle);
            if (!el) return null;
            const node = [...el.querySelectorAll("span")].map((x) => x.firstChild).find((n) => n && (n.textContent ?? "").includes("International"));
            let lines = 0;
            let first = 0;
            if (node) {
              const rg = document.createRange();
              rg.selectNodeContents(node);
              const rects = [...rg.getClientRects()];
              lines = rects.length;
              first = rects.length ? round1(rects[0].width) : 0;
            }
            const page = el.closest(".page");
            const z = page ? zPaper(page) : 1;
            const cs = page ? getComputedStyle(page) : null;
            const pr = page ? page.getBoundingClientRect() : null;
            const inner = pr && cs ? round1((pr.width - parseFloat(cs.paddingLeft || "0") - parseFloat(cs.paddingRight || "0")) / z) : 0;
            return { wb: getComputedStyle(el).wordBreak, lines: lines, first: first, inner: inner, box: round1(el.getBoundingClientRect().width / z) };
          };
          const wrapLong = (needle) => [...(wrapPara(needle)?.querySelectorAll("span") ?? [])].map((x) => x.firstChild).find((n) => n && (n.textContent ?? "").includes("International"));
          const wrapPut = (needle) => {
            const node = wrapLong(needle);
            if (!node) return 0;
            const rg = document.createRange();
            rg.setStart(node, 0);
            rg.collapse(true);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(rg);
            wrapPara(needle)?.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const wrapSelect = () => [...document.querySelectorAll("#ribbon select")].find((x) => x.title === "行距");
          const wrapDlg = () => document.querySelector(".solo-pg-dialog");
          const wrapField = (n) => wrapDlg()?.querySelector('[name="' + n + '"]') ?? null;
          const wrapOpen = async (needle) => {
            const put = wrapPut(needle);
            await wait(280);
            const sel = wrapSelect();
            if (sel) {
              sel.value = "more";
              sel.dispatchEvent(new Event("change", { bubbles: true }));
            }
            await wait(340);
            const d = wrapDlg();
            const box = wrapField("pg-wrap");
            const spans = d ? [...d.querySelectorAll(".solo-pg-snap > span")] : [];
            return {
              put: put,
              sel: sel ? 1 : 0,
              open: d?.open === true ? 1 : 0,
              // 第二枚勾的措辞（第一枚是 #134 的「如果定义了文档网格，则对齐到网格」）：WPS 那一页的中文在 .qm 那堵墙
              // 后面读不出，这句是本产品的通用说法、未取证——这一腿只钉它摆的是哪一格、不钉措辞来自谁。
              label: spans[1] ? String(spans[1].textContent).trim() : "",
              checked: box ? (box.checked ? 1 : 0) : -1,
              hint: d?.querySelector('[data-hint="pg-wrap"]')?.textContent ?? "",
              before: wrapField("pg-before")?.value ?? "?",
              after: wrapField("pg-after")?.value ?? "?",
              rule: wrapField("pg-rule")?.value ?? "?",
              value: wrapField("pg-value")?.value ?? "?",
              snap: wrapField("pg-snap") ? (wrapField("pg-snap").checked ? 1 : 0) : -1,
            };
          };
          const wrapApply = async () => {
            const b = wrapDlg()?.querySelector(".solo-pg-apply");
            const hit = b ? (b.click(), 1) : 0;
            await wait(420);
            return { applied: hit, gone: wrapDlg() ? 0 : 1 };
          };
          const wrapDirty = () => ((document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0);
          const wrapSteps = [];
          for (const name of ["摆1", "摆2", "摆3"]) {
            const o = await wrapOpen(name);
            const before = wrapLineRead(name);
            const dirtyBefore = wrapDirty();
            // 只在「此刻没勾」时落一次勾选：说了允许的那两段（摆2 自己说的、摆3 借来的）第二趟起手就是勾着的，
            // 于是这一腿在第二趟一次也不按——那才是「同一句话再说一遍不改字节」的落笔形状。
            const box = wrapField("pg-wrap");
            const clicked = o.checked === 0 && box ? (box.click(), 1) : 0;
            if (clicked) await wait(240);
            const hintAfter = wrapDlg()?.querySelector('[data-hint="pg-wrap"]')?.textContent ?? "";
            const ap = await wrapApply();
            // 纸上那两读数另起名：o 里 before/after 已经占了对话框那两格（段前/段后），
            // 沿用同名会把「件里那句话读回来还是那句」这条证据原地盖掉。
            wrapSteps.push({ name: name, ...o, clicked: clicked, hintAfter: hintAfter, ...ap, dirtyBefore: dirtyBefore, paperBefore: before, paperAfter: wrapLineRead(name), dirtyAfter: wrapDirty() });
          }
          fire("solo:save");
          await wait(1200);
          if (wrapTrack0) {
            wrapTrackBtn()?.click();
            await wait(320);
          }
          wrapLeg = {
            trackBefore: wrapTrack0,
            trackOff: wrapTrackOff,
            trackBack: wrapTrackOn(),
            steps: wrapSteps,
            savedDirty: wrapDirty(),
            dialogsLeft: document.querySelectorAll(".solo-pg-dialog").length,
            // 写回之后再看一眼那三段：后面那些腿动的是别处，而第二趟的起手就该摆成这一串——两处同一处才算
            // 「壳把自己写回盘上的那一句读回来并照它摆」，比步骤里那次紧跟的读数更晚、更硬。
            final: ["摆1", "摆2", "摆3"].map((n) => ({ name: n, ...wrapLineRead(n) })),
            pages: document.querySelectorAll("#view .page").length,
          };
        }
        // #220：「段落 → 换行和分页」那四颗勾在**装机这只壳**里走一趟并落回盘上。四摆问四种起始态：
        // 摆1 段前分页此刻没说 → 勾上之后这张纸要重切（这一段挪到新一张的纸沿、纸数多一张）；
        // 摆2 自己说了与下段同页（裸写一枚）→ 摆出来勾着，原样确定一个字节不多写；
        // 摆3 自己说了关孤行控制 → 摆出来没勾，勾上之后段落裸写一枚（WPS 拼「开」就是裸写，授权口径 46 件）；
        // 摆4 段中不分页是样式 Lent 说的 → 摆出来勾着但段落不认领，原样确定之后仍不多写这一枚。
        // 跟踪改动先按回关着（与 #214/#215 同一条路），量完按回原样交给后面的腿。
        // 那四枚记号在纸上由 render.ts 盖成 data-keep-next / data-keep-lines / data-widow 与 page-break-before 那一句，
        // 所以这一腿既读对话框那一格，也读纸上那一段到底摆着哪几枚。
        const keepDoc = (document.getElementById("title")?.textContent ?? "").includes("分页四勾");
        let keepLeg = null;
        if (keepDoc) {
          const keepTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const keepTrackOn = () => (keepTrackBtn()?.classList.contains("on") ? 1 : 0);
          const keepTrack0 = keepTrackOn();
          if (keepTrack0) {
            keepTrackBtn()?.click();
            await wait(320);
          }
          const keepTrackOff = keepTrackOn();
          const keepPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          const keepPut = (needle) => {
            const el = keepPara(needle);
            if (!el) return 0;
            const node = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode();
            if (!node) return 0;
            const rg = document.createRange();
            rg.setStart(node, 0);
            rg.collapse(true);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(rg);
            el.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          // 纸数是**画出来的**：一张连续的 .page 由条带切成几张，条带一条不多一张。页归属按条带上沿比。
          const keepStrips = () => [...document.querySelectorAll("#view .solo-sheet-strip")].map((x) => parseFloat(x.style.top) || 0).sort((a, b) => a - b);
          const keepMarks = (needle) => {
            const el = keepPara(needle);
            if (!el) return null;
            const cs = getComputedStyle(el);
            const page = el.closest(".page");
            const z = page ? zPaper(page) : 1;
            const pr = page ? page.getBoundingClientRect() : null;
            const er = el.getBoundingClientRect();
            const strips = keepStrips();
            const top = pr ? round1((er.top - pr.top) / z) : -1;
            const sheet = strips.filter((t) => t < top - 1).length;
            const sheetTop = sheet ? strips[sheet - 1] : 0;
            return {
              next: el.hasAttribute("data-keep-next") ? 1 : 0,
              lines: el.hasAttribute("data-keep-lines") ? 1 : 0,
              widow: el.getAttribute("data-widow") === "0" ? 0 : 1,
              before: cs.breakBefore === "always" || cs.pageBreakBefore === "always" ? 1 : 0,
              top: top,
              sheet: sheet,
              pages: strips.length + 1,
              // 这一段离它那一张的上沿多远：段前分页落笔后它就贴着新一张的沿（只剩上边距与页缝那道让位）。
              fromSheetTop: round1(top - sheetTop),
            };
          };
          const keepDlg = () => document.querySelector(".solo-pg-dialog");
          const keepField = (n) => keepDlg()?.querySelector('[name="' + n + '"]') ?? null;
          const keepSelect = () => [...document.querySelectorAll("#ribbon select")].find((x) => x.title === "行距");
          const keepOpen = async (needle) => {
            const put = keepPut(needle);
            await wait(280);
            const sel = keepSelect();
            if (sel) {
              sel.value = "more";
              sel.dispatchEvent(new Event("change", { bubbles: true }));
            }
            await wait(340);
            const d = keepDlg();
            const on = (n) => (keepField(n)?.checked ? 1 : 0);
            const rows = d ? [...d.querySelectorAll(".solo-pg-keep label span")] : [];
            return {
              put: put,
              sel: sel ? 1 : 0,
              open: d?.open === true ? 1 : 0,
              labels: rows.map((r) => String(r.textContent).trim()).join("/"),
              widow: on("pg-widow"),
              next: on("pg-keep-next"),
              lines: on("pg-keep-lines"),
              before: on("pg-break-before"),
              hint: d?.querySelector('[data-hint="pg-keep"]')?.textContent ?? "",
              before120: keepField("pg-before")?.value ?? "?",
            };
          };
          const keepApply = async () => {
            const b = keepDlg()?.querySelector(".solo-pg-apply");
            const hit = b ? (b.click(), 1) : 0;
            await wait(460);
            return { applied: hit, gone: keepDlg() ? 0 : 1 };
          };
          const keepDirty = () => ((document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0);
          const keepSteps = [];
          for (const name of ["摆1", "摆2", "摆3", "摆4"]) {
            const o = await keepOpen(name);
            const marksBefore = keepMarks(name);
            const pagesBefore = keepStrips().length + 1;
            const dirtyBefore = keepDirty();
            // 只在「此刻没勾」时落一次：第二趟起手那四格已经是第一趟写回的那一句，于是这一腿一次也不按。
            const clicked = [];
            if (name === "摆1" && o.before === 0 && keepField("pg-break-before")) {
              keepField("pg-break-before").click();
              clicked.push("pg-break-before");
            }
            if (name === "摆3" && o.widow === 0 && keepField("pg-widow")) {
              keepField("pg-widow").click();
              clicked.push("pg-widow");
            }
            if (clicked.length) await wait(260);
            const hintAfter = keepDlg()?.querySelector('[data-hint="pg-keep"]')?.textContent ?? "";
            const ap = await keepApply();
            keepSteps.push({
              name: name, ...o, hintAfter: hintAfter, clicked: clicked.join(","), ...ap,
              dirtyBefore: dirtyBefore, dirtyAfter: keepDirty(),
              marksBefore: marksBefore, marksAfter: keepMarks(name),
              pagesBefore: pagesBefore, pagesAfter: keepStrips().length + 1,
            });
          }
          fire("solo:save");
          await wait(1400);
          if (keepTrack0) {
            keepTrackBtn()?.click();
            await wait(320);
          }
          keepLeg = {
            trackBefore: keepTrack0,
            trackOff: keepTrackOff,
            trackBack: keepTrackOn(),
            steps: keepSteps,
            savedDirty: keepDirty(),
            dialogsLeft: document.querySelectorAll(".solo-pg-dialog").length,
            final: ["摆1", "摆2", "摆3", "摆4"].map((n) => ({ name: n, ...keepMarks(n) })),
            pages: document.querySelectorAll("#view .page").length,
          };
        }
        // #221：字号框在装机这只壳里走一趟。那一格在 WPS 里是 ComboBox（KxWpsFontSizeComboBoxCommand，
        // fontSizeList 21 档、fixedWidth="60"）：下拉只列那 21 档，而框里还能打字。这一腿问四件事——
        // 下拉是不是那 21 档（24/42 不在）、框与同一对里的字体框同顶同高且宽 60、三段各摆一种起始态
        // （号1 自己说 14.5 不在档上 / 号2 什么都没说框空着 / 号3 自己说 12），打进来的数按半点落笔。
        // 每一步只在「此刻还不是那个数」时敲一次，所以第二趟起手盘上已经是这一句，一敲也不打。
        const sizeDoc = (document.getElementById("title")?.textContent ?? "").includes("字号框");
        let sizeLeg = null;
        if (sizeDoc) {
          const szPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          const szBox = () => document.querySelector("#ribbon .rbar[data-tab='home'] input.sizebox");
          const szTwin = () => document.querySelector("#ribbon .rbar[data-tab='table'] input.sizebox");
          const szFace = () => document.querySelector("#ribbon .rbar[data-tab='home'] input.facebox");
          const szPut = (needle) => {
            const el = szPara(needle);
            if (!el) return 0;
            // 按**文字节点**两端铺选区（与真人拖选落下的形状一样）：markState 从选区起点那一枚节点往上问字号，
            // 起在 .para 上就问不到 run 自己那一句（它在那枚 span 上，不在段上）。
            // 认的是**装着这几个字的那一枚节点**，不是段里的第一枚：前几腿会给末段点一颗列表符号，段头先多出
            // 一枚 .bullet，从第一枚文字起选，问到的就是符号那一格，不是这段字自己穿的字号。
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let node = null;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) {
                node = n;
                break;
              }
            }
            if (!node) return 0;
            const r = document.createRange();
            r.setStart(node, 0);
            r.setEnd(node, (node.textContent ?? "").length);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            el.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          // 回显由 selectionchange 那一趟刷，而框自己还捏着焦点时它不写（正在打字那格不动）——先 blur 再问。
          const szEcho = (needle) => {
            szBox()?.blur();
            szPut(needle);
            return szBox()?.value ?? "?";
          };
          // 纸上量的是那枚文字所在元素的 computed 字号（pt 折成 px 是 4/3，13.5pt → 18px）：与 markState 问的是
          // 同一处——装着这几个字的那枚文字节点往上的第一枚元素（不是段头那颗列表符号）。
          const szPx = (needle) => {
            const el = szPara(needle);
            if (!el) return -1;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) return round1(parseFloat(getComputedStyle(n.parentElement ?? el).fontSize));
            }
            return round1(parseFloat(getComputedStyle(el).fontSize));
          };
          const szType = (needle, text) => {
            szPut(needle);
            const b = szBox();
            if (!b) return 0;
            b.value = text;
            b.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          const szBox0 = szBox();
          const szFace0 = szFace();
          // 先回到「开始」那一档：字号框只在那一页上，别的档看不见也点不着。
          const szHomeTab = [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === "开始");
          if (szHomeTab && !szHomeTab.classList.contains("on")) {
            szHomeTab.click();
            await wait(220);
          }
          szBox0?.focus();
          await wait(160);
          const szFocus = document.activeElement === szBox0 ? 1 : 0;
          const szListEl = document.getElementById("doc-sizes") ? 1 : 0;
          const szList = [...document.querySelectorAll("#doc-sizes option")].map((o) => o.value);
          const szRect = szBox0?.getBoundingClientRect();
          const szFaceRect = szFace0?.getBoundingClientRect();
          const szTwin0 = szTwin();
          const szTwinList = szTwin0?.getAttribute("list") ?? "";
          const szTwinType = szTwin0?.type ?? "";
          szBox0?.blur();
          const szEcho0 = { a: szEcho("号1"), b: szEcho("号2"), c: szEcho("号3") };
          const szSel = (() => {
            const s = document.getSelection();
            return s && s.rangeCount ? (s.isCollapsed ? "collapsed" : "range") : "none";
          })();
          const szPx0 = { a: szPx("号1"), b: szPx("号2"), c: szPx("号3") };
          const szClicks = [];
          // 号2 起手什么都没说：打 13.7 → 按半点落到 13.5，框里也只留 13.5（纸上 18px）。
          if (szEcho("号2") === "") {
            szType("号2", "13.7");
            await wait(240);
            szClicks.push("号2:13.7");
          }
          const szRound = { echo: szEcho("号2"), px: szPx("号2") };
          // 打进来的不是磅数就不落笔：框退回件里此刻那一句，纸上不动。
          if (szEcho("号2") === "13.5") {
            szType("号2", "abc");
            await wait(240);
            szClicks.push("号2:abc");
          }
          const szJunk = { echo: szEcho("号2"), px: szPx("号2") };
          if (szEcho("号2") !== "72") {
            szType("号2", "72");
            await wait(240);
            szClicks.push("号2:72");
          }
          // 24 不在 WPS 那 21 档里（fontSizeList 无 24）：能打字的这条路正是为了这种数。
          if (szEcho("号3") !== "24") {
            szType("号3", "24");
            await wait(240);
            szClicks.push("号3:24");
          }
          const szAfter = {};
          for (const n of ["号1", "号2", "号3"]) szAfter[n] = { echo: szEcho(n), px: szPx(n) };
          fire("solo:save");
          await wait(1400);
          sizeLeg = {
            tag: szBox0?.tagName ?? "",
            type: szBox0?.type ?? "",
            listId: szBox0?.getAttribute("list") ?? "",
            twinTag: szTwin0?.tagName ?? "",
            twinType: szTwinType,
            twinList: szTwinList,
            listCount: szList.length,
            listHas24: szList.includes("24") ? 1 : 0,
            listHas42: szList.includes("42") ? 1 : 0,
            listHas5: szList.includes("5") ? 1 : 0,
            listHas72: szList.includes("72") ? 1 : 0,
            w: szRect ? round1(szRect.width) : -1,
            dh: szRect && szFaceRect ? round1(szRect.height - szFaceRect.height) : -99,
            dtop: szRect && szFaceRect ? round1(szRect.top - szFaceRect.top) : -99,
            focus: szFocus,
            listEl: szListEl,
            sel: szSel,
            echo0: szEcho0,
            px0: szPx0,
            round: szRound,
            junk: szJunk,
            after: szAfter,
            clicks: szClicks.join(","),
            savedDirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0,
          };
        }
        // #223：清除格式在装机这只壳里走一趟。WPS 把这一枚摆在 RB_Font 里字体名与字号之后、加粗之前
        // （CT_Home.kuip 的 ClearFormatsRainbow，idMso="ClearFormatting"、icon="eraser_kd"、
        // hotKey="E"），中文界面字面「清除格式」（mui/zh_CN/wps.qm）。这一腿问四件事——
        // 组内次序（字体框/字号框/清除格式/加粗自左向右）、甲组整组直接格式被撤到与什么都没说的丙组同脸、
        // 乙组那枚只有 w:rStyle/w:lang 的 run 一个字节不动、连按第二次不再变。
        const cfDoc = (document.getElementById("title")?.textContent ?? "").includes("清除格式");
        let clearLeg = null;
        if (cfDoc) {
          const cfHomeTab = [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === "开始");
          if (cfHomeTab && !cfHomeTab.classList.contains("on")) {
            cfHomeTab.click();
            await wait(220);
          }
          // 前面几腿把「修订」开着交过来（量的是别的族，量完按回原样）。开着落这一击，撤干净的 rPr 旁边会
          // 多一枚 w:rPrChange 记着原来那组话——那是修订那一族要问的形状，与本族「整组直接格式撤掉了没有」
          // 混在一张纸上，反读就说不清是谁写的。先按回关着（走产品自己那枚开关），量完按回原样。
          // 开着清一次会留哪一枚改由 #149 那一族的单测管。
          const cfTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const cfTrackOn = () => (cfTrackBtn()?.classList.contains("on") ? 1 : 0);
          const cfTrack0 = cfTrackOn();
          if (cfTrack0) {
            cfTrackBtn()?.click();
            await wait(320);
          }
          const cfTrackOff = cfTrackOn();
          const cfPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          const cfPut = (needle) => {
            const el = cfPara(needle);
            if (!el) return 0;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let node = null;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) {
                node = n;
                break;
              }
            }
            if (!node) return 0;
            const r = document.createRange();
            r.setStart(node, 0);
            r.setEnd(node, (node.textContent ?? "").length);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            el.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const cfLook = (needle) => {
            const el = cfPara(needle);
            if (!el) return { w: "-1" };
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let host = el;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) {
                host = n.parentElement ?? el;
                break;
              }
            }
            // 上下标的缩一圈套在 run 自己那一枚里面：纸上的线、底色、字号都写在**外面**那一枚上，
            // 从里面问只问得到继承来的那几项。往上走到 run 自己那一枚为止——再往上那一圈是修订的壳
            // （.page .rev-fmt 自己画一条点线表示「这一枚的格式被改过」），它的线不是 run 说的线，
            // 问到它头上会把「清干净了」读成还剩一条下划线。
            const cfRev = (n) => !!n && (n.classList.contains("rev-fmt") || n.classList.contains("rev"));
            while (host.parentElement && host.parentElement !== el && !cfRev(host.parentElement)) host = host.parentElement;
            let wrap = host;
            while (wrap.parentElement && wrap.parentElement !== el) wrap = wrap.parentElement;
            const cs = getComputedStyle(host);
            return {
              w: cs.fontWeight,
              style: cs.fontStyle,
              deco: cs.textDecorationLine,
              color: cs.color,
              size: round1(parseFloat(cs.fontSize)),
              bg: cs.backgroundColor,
              va: cs.verticalAlign,
              fmtRev: wrap.classList.contains("rev-fmt") ? 1 : 0,
              u: host.getAttribute("data-u") ?? "",
            };
          };
          const cfSizeBox = () => document.querySelector("#ribbon .rbar[data-tab='home'] input.sizebox");
          const cfFaceBox = () => document.querySelector("#ribbon .rbar[data-tab='home'] input.facebox");
          const cfBtn = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button")].find((b) => (b.textContent ?? "").trim() === "清除格式");
          const cfBold = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button")].find((b) => (b.getAttribute("title") ?? "").startsWith("加粗"));
          // 回显跟着 selectionchange 走，而框还捏着焦点时不刷——先 blur 再问（与字号那一腿同一条）。
          const cfEcho = (needle) => {
            cfSizeBox()?.blur();
            cfFaceBox()?.blur();
            cfPut(needle);
            return { size: cfSizeBox()?.value ?? "?", face: cfFaceBox()?.value ?? "?", bold: cfBold()?.classList.contains("on") ? 1 : 0 };
          };
          const cfOrder = (() => {
            const face = cfFaceBox()?.getBoundingClientRect();
            const size = cfSizeBox()?.getBoundingClientRect();
            const btn = cfBtn()?.getBoundingClientRect();
            const bold = cfBold()?.getBoundingClientRect();
            if (!face || !size || !btn || !bold) return null;
            return { faceLtSize: face.left < size.left ? 1 : 0, sizeLtClear: size.left < btn.left ? 1 : 0, clearLtBold: btn.left < bold.left ? 1 : 0 };
          })();
          const cfBefore = { a: cfLook("整组都说"), b: cfLook("只有槽"), c: cfLook("什么都没说"), echo: cfEcho("整组都说"), htmlB: (cfPara("只有槽")?.innerHTML ?? "").length, btn: cfBtn() ? 1 : 0 };
          const cfClicks = [];
          // 甲组起手是整组都说（粗/斜/线/删/红/22pt/黄底/上标/加宽）：只在这一句还看得出与丙组不同时才按一次。
          if (cfBefore.a.w !== cfBefore.c.w || cfBefore.a.color !== cfBefore.c.color) {
            cfPut("整组都说");
            cfBtn()?.click();
            await wait(240);
            cfClicks.push("甲");
          }
          const cfAfter = { a: cfLook("整组都说"), b: cfLook("只有槽"), c: cfLook("什么都没说"), echo: cfEcho("整组都说"), htmlB: (cfPara("只有槽")?.innerHTML ?? "").length };
          // 第二趟起手盘上已经是清完的样子；再按一次应当一个字也不再变。
          if (cfAfter.a.w === cfAfter.c.w) {
            cfPut("整组都说");
            cfBtn()?.click();
            await wait(240);
            cfClicks.push("甲2");
          }
          const cfAgain = { a: cfLook("整组都说"), echo: cfEcho("整组都说") };
          // 量完把手按回原样交给后面的腿（与 #214/#219 同一条）。
          if (cfTrack0) {
            cfTrackBtn()?.click();
            await wait(320);
          }
          fire("solo:save");
          await wait(1400);
          clearLeg = {
            order: cfOrder,
            btn: cfBefore.btn,
            before: cfBefore,
            after: cfAfter,
            again: cfAgain,
            clicks: cfClicks.join(","),
            trackBefore: cfTrack0,
            trackOff: cfTrackOff,
            trackBack: cfTrackOn(),
            savedDirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0,
          };
        }
        // #226：开始页「样式」画廊在装机这只壳里走一趟。WPS 的 RB_Style 在 RB_Paragraph 之后、RB_Editing
        // 之前（CT_Home.kuip 的 idQ 序），下拉里列的是**这份文件自己声明的**段落样式，第一格是默认那枚
        // （选它就是撤掉 w:pStyle）。这一腿问五件事：选项清单与次序、甲组换成「标题 1」后纸上换成 16 磅
        // Arial 深蓝粗体且段距跟着样式走、乙组自己写的缩进与粗体留着、丙组换回正文后 pStyle 撤掉且回显
        // 跟着光标回到空值、丁组换成「引用」后字符样式那张脸与那点墨留着而斜体从段落链借来。
        const galDoc = (document.getElementById("title")?.textContent ?? "").includes("样式画廊");
        let styleLeg = null;
        if (galDoc) {
          const galHome = [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === "开始");
          if (galHome && !galHome.classList.contains("on")) {
            galHome.click();
            await wait(220);
          }
          // 前面几腿量别的族时可能把「修订」开着交过来；开着换样式会多一枚 w:pPrChange，与本族问的
          // 「落笔只有一句 w:pStyle」混在一起。先按回关着（走产品自己那枚开关），量完按回原样。
          const galTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const galTrackOn = () => (galTrackBtn()?.classList.contains("on") ? 1 : 0);
          const galTrack0 = galTrackOn();
          if (galTrack0) {
            galTrackBtn()?.click();
            await wait(320);
          }
          const galSel = [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] select")].find((x) => x.title === "样式");
          const galPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          const galPut = (needle) => {
            const el = galPara(needle);
            if (!el) return 0;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let node = null;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) {
                node = n;
                break;
              }
            }
            if (!node) return 0;
            const r = document.createRange();
            r.setStart(node, 0);
            r.setEnd(node, (node.textContent ?? "").length);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            el.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const galPick = (value) => {
            if (!galSel) return 0;
            galSel.value = value;
            galSel.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          const galLook = (needle) => {
            const el = galPara(needle);
            if (!el) return null;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let host = el;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) {
                host = n.parentElement ?? el;
                break;
              }
            }
            while (host.parentElement && host.parentElement !== el) host = host.parentElement;
            const cs = getComputedStyle(host);
            const ps = getComputedStyle(el);
            return {
              tag: el.tagName,
              outline: el.getAttribute("data-outline") ?? "",
              w: cs.fontWeight,
              style: cs.fontStyle,
              color: cs.color,
              size: round1(parseFloat(cs.fontSize)),
              family: cs.fontFamily,
              top: round1(parseFloat(ps.marginTop)),
              left: round1(parseFloat(ps.marginLeft)),
            };
          };
          const galOpts = galSel ? [...galSel.options].map((o) => o.value + "=" + o.textContent) : [];
          const galListsBtn = [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button")].find((b) => (b.getAttribute("title") ?? "").startsWith("变回普通段落"));
          const galFindBtn = [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button")].find((b) => (b.getAttribute("title") ?? "").startsWith("查找与替换"));
          const galOrder = (() => {
            const sel = galSel?.getBoundingClientRect();
            const lists = galListsBtn?.getBoundingClientRect();
            const find = galFindBtn?.getBoundingClientRect();
            if (!sel || !lists || !find) return null;
            return { afterLists: sel.left > lists.left ? 1 : 0, beforeFind: sel.left < find.left ? 1 : 0 };
          })();
          const galEcho = () => (galSel ? galSel.value : "?");
          galPut("甲组");
          await wait(180);
          const galEcho0 = galEcho();
          galPick("2");
          await wait(320);
          const galJia = galLook("甲组");
          galPut("戊组");
          await wait(180);
          const galEchoQuote = galEcho();
          galPut("甲组");
          await wait(180);
          const galEchoJia = galEcho();
          galPut("乙组");
          await wait(160);
          galPick("2");
          await wait(320);
          const galYi = galLook("乙组");
          galPut("丙组");
          await wait(180);
          galPick("");
          await wait(320);
          const galBing = galLook("丙组");
          galPut("丙组");
          await wait(180);
          const galEchoBing = galEcho();
          galPut("丁组");
          await wait(180);
          const galEchoDing0 = galEcho();
          galPick("3");
          await wait(320);
          const galDing = galLook("丁组");
          galPut("丁组");
          await wait(180);
          const galEchoDing = galEcho();
          if (galTrack0) {
            galTrackBtn()?.click();
            await wait(320);
          }
          fire("solo:save");
          await wait(1400);
          styleLeg = {
            opts: galOpts,
            order: galOrder,
            echo0: galEcho0,
            echoQuote: galEchoQuote,
            echoJia: galEchoJia,
            echoBing: galEchoBing,
            echoDing0: galEchoDing0,
            echoDing: galEchoDing,
            jia: galJia,
            yi: galYi,
            bing: galBing,
            ding: galDing,
            trackBefore: galTrack0,
            trackBack: galTrackOn(),
            savedDirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0,
          };
        }
        // #227：开始页那两枚「上标/下标」在装机这只壳里走一趟。WPS 的 CT_Home.kuip:50-51 声明这两条命令，
        // :136-138 又把它们折成一枚 Font4（singlerowfold="1"），摆在删除线与字体效果之间。这一腿问四件事——
        // 组内次序（删除线/上标/下标自左向右）、件里三种说法各自回显（乙上标亮、丙下标亮、丁明说 baseline 两枚都不亮）、
        // 甲组按上标再按下标是「换一档」而不是叠两句、乙组与丁组按回自己那一档是撤槽（纸上回到不抬不落）、
        // 戊组那句 22 磅在字号框里仍读得回来（缩一圈只是纸的排版层）。盘上落回哪几句由装机器反读那一步问。
        const vaDoc = (document.getElementById("title")?.textContent ?? "").includes("上下标");
        let vertLeg = null;
        if (vaDoc) {
          const vaHomeTab = [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === "开始");
          if (vaHomeTab && !vaHomeTab.classList.contains("on")) {
            vaHomeTab.click();
            await wait(220);
          }
          // 同 #223/#226：修订开着落这一击会多一枚 w:rPrChange，与本族问的「同槽只有一句」混在一起。
          const vaTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const vaTrackOn = () => (vaTrackBtn()?.classList.contains("on") ? 1 : 0);
          const vaTrack0 = vaTrackOn();
          if (vaTrack0) {
            vaTrackBtn()?.click();
            await wait(320);
          }
          const vaPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          const vaPut = (needle) => {
            const el = vaPara(needle);
            if (!el) return 0;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let node = null;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) {
                node = n;
                break;
              }
            }
            if (!node) return 0;
            const r = document.createRange();
            r.setStart(node, 0);
            r.setEnd(node, (node.textContent ?? "").length);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            el.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          // 缩一圈套在 run 自己那一枚里面：里头那枚说 vertical-align，外头那枚说文件写的磅数。两处各问各的，
          // 问到同一枚头上就会把「22 磅」读成 16.5（.75em 算出来的纸面值），那种读数正是这一族要防的降级。
          const vaLook = (needle) => {
            const el = vaPara(needle);
            if (!el) return null;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let node = null;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) {
                node = n;
                break;
              }
            }
            if (!node) return null;
            const inner = node.parentElement ?? el;
            let outer = inner;
            while (outer.parentElement && outer.parentElement !== el) outer = outer.parentElement;
            const ci = getComputedStyle(inner);
            const co = getComputedStyle(outer);
            return { va: ci.verticalAlign, inner: round1(parseFloat(ci.fontSize)), outer: round1(parseFloat(co.fontSize)) };
          };
          const vaBtn = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button")].find((b) => (b.textContent ?? "").trim() === label);
          const vaEcho = (needle) => {
            vaPut(needle);
            return { sup: vaBtn("上标")?.classList.contains("on") ? 1 : 0, sub: vaBtn("下标")?.classList.contains("on") ? 1 : 0 };
          };
          const vaSizeBox = () => document.querySelector("#ribbon .rbar[data-tab='home'] input.sizebox");
          const vaStrike = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button")].find((b) => (b.getAttribute("title") ?? "").startsWith("删除线"));
          const vaOrder = (() => {
            const strike = vaStrike()?.getBoundingClientRect();
            const sup = vaBtn("上标")?.getBoundingClientRect();
            const sub = vaBtn("下标")?.getBoundingClientRect();
            if (!strike || !sup || !sub) return null;
            return { strikeLtSup: strike.left < sup.left ? 1 : 0, supLtSub: sup.left < sub.left ? 1 : 0 };
          })();
          const vaEchoYi = vaEcho("乙组");
          const vaEchoBing = vaEcho("丙组");
          const vaEchoDing = vaEcho("丁组");
          const vaEchoJia = vaEcho("甲组");
          // 甲组：一件什么都没说的 run，按上标再按下标。第二下是「换一档」，两句不许并排站着。
          vaPut("甲组");
          vaBtn("上标")?.click();
          await wait(240);
          const vaJiaSup = { echo: vaEcho("甲组"), look: vaLook("甲组") };
          vaBtn("下标")?.click();
          await wait(240);
          const vaJiaSub = { echo: vaEcho("甲组"), look: vaLook("甲组") };
          // 丙组：本来下标，按上标换过来。
          vaPut("丙组");
          vaBtn("上标")?.click();
          await wait(240);
          const vaBingSup = { echo: vaEcho("丙组"), look: vaLook("丙组") };
          // 乙组：本来上标，再按一次上标就是撤回——纸上回到不抬不落，回显两枚都灭。
          vaPut("乙组");
          vaBtn("上标")?.click();
          await wait(240);
          const vaYiOff = { echo: vaEcho("乙组"), look: vaLook("乙组") };
          // 丁组：明说 baseline 的 run 按上去再按回来，原来那句让位、回来时槽撤掉（不补写一句 baseline）。
          vaPut("丁组");
          vaBtn("上标")?.click();
          await wait(240);
          const vaDingOn = { echo: vaEcho("丁组"), look: vaLook("丁组") };
          vaBtn("上标")?.click();
          await wait(240);
          const vaDingOff = { echo: vaEcho("丁组"), look: vaLook("丁组") };
          // 戊组：22 磅的上标——字号框读文件那一句，纸里头那枚才是缩过的。
          vaPut("戊组");
          await wait(180);
          const vaSizeEcho = { value: vaSizeBox()?.value ?? "?", look: vaLook("戊组") };
          if (vaTrack0) {
            vaTrackBtn()?.click();
            await wait(320);
          }
          fire("solo:save");
          await wait(1400);
          vertLeg = {
            order: vaOrder,
            echoYi: vaEchoYi,
            echoBing: vaEchoBing,
            echoDing: vaEchoDing,
            echoJia: vaEchoJia,
            jiaSup: vaJiaSup,
            jiaSub: vaJiaSub,
            bingSup: vaBingSup,
            yiOff: vaYiOff,
            dingOn: vaDingOn,
            dingOff: vaDingOff,
            sizeEcho: vaSizeEcho,
            trackBefore: vaTrack0,
            trackBack: vaTrackOn(),
            savedDirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0,
          };
        }
        // #231：开始页「更改大小写」那一格在装机这只壳里走一趟。WPS 的 CT_Home.kuip:121 把这几条命令折成一枚
        // 折叠钮 Font1（referenceCommand="ChangeCaseRainbow"、idMso="ChangeCaseGallery"、ksoCmd="ChangeCase"、
        // rbHotKey="7"），组内排在字号放大/缩小之后、ClearFormatsRainbow 之前（:120-123）。档位与档位名按
        // Word 那五档摆：mui/zh_CN/wps.qm 只证明「更改大小写」「全部大写」这些字面在包里有（那是一张扁平哈希池），
        // 不证明下拉就摆这五档、也不证明次序——未取证。这一腿问四件事：① 格在开始页有面、与字号框同组、
        // 字号到更改大小写到清除格式自左向右；② 落「全部大写」把拉丁抬起来、汉字一个字不动、格子落回占位、件报脏；
        // ③ 改的是字面不是格式——丁组自己那句粗与 16 磅在纸上仍读得回来，撤销之后字面回原样而格式还在；
        // ④ ß 不摊成 SS（长度不变是这一族的根，WPS 真机在那一点上怎么落未取证）。
        // 乙组那句 w:tab 与盘上落回的 w:t 字面由装机器反读那一步问。
        const csDoc = (document.getElementById("title")?.textContent ?? "").includes("大小写");
        let caseLeg = null;
        if (csDoc) {
          const csHomeTab = [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === "开始");
          if (csHomeTab && !csHomeTab.classList.contains("on")) {
            csHomeTab.click();
            await wait(220);
          }
          // 同 #223/#226/#227：修订开着时这一笔会划出 w:del/w:ins，与本族问的「字面改了、格式没改」混在一起。
          const csTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const csTrackOn = () => (csTrackBtn()?.classList.contains("on") ? 1 : 0);
          const csTrack0 = csTrackOn();
          if (csTrack0) {
            csTrackBtn()?.click();
            await wait(320);
          }
          const ZW = String.fromCharCode(8203);
          const csPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          // 每问一次都重新查：落笔那一趟 render() 换的是 host.innerHTML，上一趟拿到的节点已经不在树上。
          const csText = (needle) => {
            const el = csPara(needle);
            return el ? (el.textContent ?? "").split(ZW).join("") : null;
          };
          const csWalk = (needle) => {
            const el = csPara(needle);
            if (!el) return null;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) return n;
            }
            return null;
          };
          // 整段圈进来：焦点得先落回页里，ribbon 那一格才把这一截当用户圈的选区（与 applyStyle 同一条回放）。
          const csPut = (needle) => {
            const el = csPara(needle);
            if (!el) return 0;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            const nodes = [];
            for (let n = walk.nextNode(); n; n = walk.nextNode()) nodes.push(n);
            if (!nodes.length) return 0;
            const lastNode = nodes[nodes.length - 1];
            const r = document.createRange();
            r.setStart(nodes[0], 0);
            r.setEnd(lastNode, (lastNode.textContent ?? "").length);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            el.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const csBox = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] select.pick")].find((x) => (x.getAttribute("title") ?? "").includes("更改大小写"));
          const csPick = (label) => {
            const box = csBox();
            if (!box) return 0;
            const opt = [...box.options].find((o) => o.textContent === label);
            if (!opt) return 0;
            box.value = opt.value;
            box.dispatchEvent(new Event("change"));
            return 1;
          };
          const csSizeBox = () => document.querySelector("#ribbon .rbar[data-tab='home'] input.sizebox");
          const csClearBtn = () => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button")].find((b) => (b.textContent ?? "").trim() === "清除格式");
          const csGroup = csBox() && csSizeBox() ? (csBox().parentElement === csSizeBox().parentElement ? 1 : 0) : 0;
          const csOptions = [...(csBox() ? csBox().options : [])].map((o) => o.textContent).join("|");
          const csOrder = (() => {
            const a = csSizeBox()?.getBoundingClientRect();
            const b = csBox()?.getBoundingClientRect();
            const c = csClearBtn()?.getBoundingClientRect();
            if (!a || !b || !c) return null;
            return { sizeLtCase: a.left < b.left ? 1 : 0, caseLtClear: b.left < c.left ? 1 : 0 };
          })();
          const csLook = (needle) => {
            const node = csWalk(needle);
            if (!node || !node.parentElement) return null;
            const cs = getComputedStyle(node.parentElement);
            return { weight: cs.fontWeight, size: round1(parseFloat(cs.fontSize)) };
          };
          const csDirty = () => ((document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0);
          // 甲组：一句 run 里汉字与拉丁并排。抬起来之后汉字一个字不能少，格子要回到占位（这一格是选方向的，不是开关）。
          csPut("甲组");
          const csJiaBefore = csText("甲组");
          csPick("全部大写");
          await wait(260);
          const csJiaUpper = { text: csText("甲组"), echo: csBox()?.value ?? "?", dirty: csDirty() };
          csPut("甲组");
          csPick("小写");
          await wait(260);
          const csJiaLower = csText("甲组");
          // 乙组：三枚 run，中间那枚只有 w:tab 没有字面。两头的字都抬起来，那一枚不许被吃掉（盘上还剩一枚由反读问）。
          csPut("乙组");
          csPick("全部大写");
          await wait(260);
          const csYi = { text: csText("乙组"), len: (csText("乙组") ?? "").length };
          // 丙组：ß 摊开是两枚字符，家里按「长度不变」那一档留着原字面。整段改前改后一样长才算这一条。
          csPut("丙组");
          const csBingBefore = csText("丙组");
          csPick("全部大写");
          await wait(260);
          const csBing = { text: csText("丙组"), sameLen: (csText("丙组") ?? "").length === (csBingBefore ?? "").length ? 1 : 0 };
          // 丁组：自己写了粗与 16 磅。切换大小写只改字面，那两句格式得仍站在纸上；撤销回原样，格式还得在。
          csPut("丁组");
          csPick("切换大小写");
          await wait(260);
          const csDingToggled = { text: csText("丁组"), look: csLook("丁组") };
          [...document.querySelectorAll("#ribbon button.rb")].find((b) => (b.textContent ?? "").trim() === "撤销")?.click();
          await wait(260);
          const csDingUndone = { text: csText("丁组"), look: csLook("丁组") };
          if (csTrack0) {
            csTrackBtn()?.click();
            await wait(320);
          }
          fire("solo:save");
          await wait(1400);
          caseLeg = {
            hasBox: csBox() ? 1 : 0,
            group: csGroup,
            order: csOrder,
            options: csOptions,
            jiaBefore: csJiaBefore,
            jiaUpper: csJiaUpper,
            jiaLower: csJiaLower,
            yi: csYi,
            bingBefore: csBingBefore,
            bing: csBing,
            dingToggled: csDingToggled,
            dingUndone: csDingUndone,
            trackBefore: csTrack0,
            trackBack: csTrackOn(),
            savedDirty: csDirty(),
          };
        }
        // #232：制表位前导符 w:leader 在装机这只壳里走一趟。家里那 22 例（.scratch/tab-leader-dom.ts）已经把
        // 「token 只跟着这一枚真正停进去的那一格」与「有墨/无墨两胞胎的几何一字不变」量完了；这一腿问的是
        // **这只壳**：打包进 renderer.js 的那 27 行 [data-leader] 规则、真铺排、真 w:defaultTabStop 之下，
        // 五段各枚制表符挂的 token、::after 上没上墨、宽度与落点各是什么形状；再在甲线尾上打一次字，问那一截
        // 点线跟不跟着走（盘上的 w:leader 由 .scratch/tab-leader-install-check.sh 独立读回）。
        const ldDoc = (document.getElementById("title")?.textContent ?? "").includes("前导符");
        let leaderLeg = null;
        if (ldDoc) {
          // 每问一次都重新查：落笔那一趟 render() 换的是 host.innerHTML，上一趟拿到的节点已经不在树上。
          const ldPara = (needle) =>
            [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          const ldTab = (t, baseLeft) => {
            const cs = getComputedStyle(t, "::after");
            const bg = cs.backgroundImage;
            const bgc = cs.backgroundColor;
            const painted = (bg && bg !== "none") || (bgc && bgc !== "rgba(0, 0, 0, 0)" && bgc !== "transparent");
            const r = t.getBoundingClientRect();
            return {
              token: t.dataset.leader ?? "",
              ink: painted ? 1 : 0,
              after: cs.height + "@" + cs.bottom + " " + cs.position,
              w: round1(r.width),
              left: round1(r.left - baseLeft),
              h: round1(r.height),
              text: t.textContent,
            };
          };
          const ldRead = (needle) => {
            const el = ldPara(needle);
            if (!el) return null;
            const page = el.closest(".page");
            const base = page ? page.getBoundingClientRect().left : 0;
            return {
              attr: el.getAttribute("data-tab-stops") ?? "",
              tabs: [...el.querySelectorAll("span.tab")].map((t) => ldTab(t, base)),
            };
          };
          const ldDirty = () => ((document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0);
          const ldTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const ldTrack0 = ldTrackBtn()?.classList.contains("on") ? 1 : 0;
          if (ldTrack0) {
            ldTrackBtn().click();
            await wait(320);
          }
          const ldJia = ldRead("甲线");
          const ldYi = ldRead("乙线");
          const ldBing = ldRead("丙线");
          const ldDing = ldRead("丁线");
          const ldWu = ldRead("戊线");
          const ldTail = ldRead("尾段");
          const ldCount = () => ({
            spans: document.querySelectorAll("#view .page span.tab").length,
            tokens: document.querySelectorAll("#view .page span.tab[data-leader]").length,
            zero: [...document.querySelectorAll("#view .page span.tab")].filter((t) => round1(t.getBoundingClientRect().height) !== 0).length,
          });
          const ldBefore = ldCount();
          // 甲线尾上打一次字：右对齐那枚位得把制表符垫出来的宽让给新长的字面，而那一截点线得跟着这枚制表符
          // 一起短——短到还在（不是撤掉），token 也一个字节不改。
          const ldEl = ldPara("甲线");
          if (ldEl) {
            const walk = document.createTreeWalker(ldEl, NodeFilter.SHOW_TEXT);
            let last = null;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) last = n;
            if (last) {
              const r = document.createRange();
              r.setStart(last, (last.textContent ?? "").length);
              r.collapse(true);
              const s = document.getSelection();
              s.removeAllRanges();
              s.addRange(r);
              ldEl.closest(".page").focus();
              document.execCommand("insertText", false, "打字");
              await wait(320);
            }
          }
          const ldJiaAfter = ldRead("甲线");
          const ldAfter = ldCount();
          if (ldTrack0) {
            const back = [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
            back?.click();
            await wait(320);
          }
          fire("solo:save");
          await wait(1400);
          leaderLeg = {
            jia: ldJia,
            yi: ldYi,
            bing: ldBing,
            ding: ldDing,
            wu: ldWu,
            tail: ldTail,
            before: ldBefore,
            jiaAfter: ldJiaAfter,
            after: ldAfter,
            text: (ldPara("甲线")?.textContent ?? "").split(String.fromCharCode(8203)).join(""),
            trackBefore: ldTrack0,
            trackBack: ldTrackBtn()?.classList.contains("on") ? 1 : 0,
            savedDirty: ldDirty(),
          };
        }
        // #234：「制表位」对话框在**装机这只盒子**里走两趟——两趟跑的是同一段腿代码，问的是同一批事，
        // 由 .scratch/tab-stops-install-check.sh 拿两份日志与两份盘上的件各问各的。家里那 35 例
        // （.scratch/tab-stops-dom.ts）量的是真 Chrome 里的 DOM；这一腿只问打包件才答得了的四件：
        // ① ribbon 上那枚「制表位」在不在、排在那一组的第几枚（打进 renderer.js 才算）；
        // ② 对话框那一段样式跟着打包件上身了没有（.solo-ts-dialog 的宽度、那枚「确定」的底色）；
        // ③ 一次「确定」之后**真铺排**下纸上那一枚的落点，与写回盘上的那两句（w:tabs、settings.xml 的
        //   w:defaultTabStop——纸与盘两条账各管各的，谁也不替谁背书）；
        // ④ 撤销与重做在打包件的栈上是**一整件事**：「撤销」一次纸上退回去，「重做」一次又落回来。
        // ⑤ 「改了才报脏」的反面：取消、填错、原样确定三条都不许动那一枚 ●，真落笔的那一条必须动。
        //   进这一条时纸已经脏着——在它前面的腿（docPart 打字、画者拖动）各自验的正是「改了字节要报脏」，
        //   所以起手先存一次把 ● 清零（dirty.baseline），后面每一格读数才是绝对口径，不是两笔脏的差。
        // ⑥ 「前导符」那一格在盒子里真按一下，两个方向都按：无→点线要纸上铺出墨、点线→无要把墨撤干净，
        //   而盘上那一句跟着重拼。按一下的对象挑的是**那一枚 Tab 真停进去的格**（丁段自己的 left 720）：
        //   甲段 8306 那枚虽然带着 dot，甲这一枚 Tab 却停在 720/2880，从没落进 8306——墨只跟停进去的那一格，
        //   拿甲的 8306 按「纸上有没有墨」永远问不出答案（清单与字节仍按原样带走，那是另一条账）。
        // 字面全按对话框自己那格的单位换算（壳的测量单位不落进文件，写死的期望会随档位漂）。
        const tsDoc = (document.getElementById("title")?.textContent ?? "").includes("制表位这一击");
        let tabStopsLeg = null;
        if (tsDoc) {
          const tsPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          const tsDirty = () => ((document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0);
          const tsBtn = () => [...document.querySelectorAll("#ribbon button")].find((b) => String(b.title ?? "").indexOf("制表位") === 0);
          const tsDlg = () => document.querySelector(".solo-ts-dialog");
          const tsList = () => tsDlg()?.querySelector("[name='ts-list']") ?? null;
          const tsRows = () => [...(tsList()?.options ?? [])];
          const tsField = (n) => tsDlg()?.querySelector("[name='" + n + "']") ?? null;
          const tsUnit = () => String(tsField("ts-pos")?.parentElement?.querySelector("span")?.textContent ?? "");
          const tsPer = () => (tsUnit() === "毫米" ? 1440 / 25.4 : tsUnit() === "英寸" ? 1440 : tsUnit() === "磅" ? 20 : 1440 / 2.54);
          const tsDec = () => (tsUnit() === "毫米" ? 2 : tsUnit() === "英寸" ? 4 : tsUnit() === "磅" ? 2 : 3);
          // 缇 ↔ 那一格的字面：与 doc-units.ts 同一套（每单位缇数与该档能精确往返的最小小数位）。
          const tsText = (tw) => String(Number((tw / tsPer()).toFixed(tsDec())));
          const tsTwips = (txt) => { const v = parseFloat(txt); return isFinite(v) ? Math.round(v * tsPer()) : NaN; };
          const tsType = (n, v) => { const el = tsField(n); if (el) { el.value = v; el.dispatchEvent(new Event("input", { bubbles: true })); } };
          const tsMark = (v) => { const r = tsDlg()?.querySelector("input[name='ts-val'][value='" + v + "']"); if (r) { r.checked = true; r.dispatchEvent(new Event("change", { bubbles: true })); } };
          const tsPickRow = (i) => { const l = tsList(); if (l) { l.selectedIndex = i; l.dispatchEvent(new Event("change", { bubbles: true })); } };
          const tsClick = (n) => { const el = tsDlg()?.querySelector("[name='" + n + "']"); el?.click(); };
          const tsApply = () => tsDlg()?.querySelector(".solo-ts-apply")?.click();
          const tsCancel = () => tsDlg()?.querySelector(".solo-ts-cancel")?.click();
          const tsRb = (label) => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === label);
          const tsHint = (n) => String(tsDlg()?.querySelector("[data-hint='" + n + "']")?.textContent ?? "");
          const tsErr = () => String(tsDlg()?.querySelector(".solo-ts-error")?.textContent ?? "");
          // 那一枚制表符身上的墨（口径同 leaderLeg 的 ldTab）：data-leader 的 token 加 ::after 到底画没画。
          const tsTabEl = (needle, i) => {
            const el = tsPara(needle);
            if (!el) return null;
            const ts = [...el.querySelectorAll("span.tab")];
            return ts[i ?? 0] ?? null;
          };
          const tsInk = (needle, i) => {
            const t = tsTabEl(needle, i);
            if (!t) return null;
            const cs = getComputedStyle(t, "::after");
            const painted = (cs.backgroundImage && cs.backgroundImage !== "none") ||
              (cs.backgroundColor && cs.backgroundColor !== "rgba(0, 0, 0, 0)" && cs.backgroundColor !== "transparent");
            return { attr: String(t.dataset.leader ?? ""), ink: painted ? 1 : 0, span: round1(t.getBoundingClientRect().width) };
          };
          const tsPick = (n, v) => { const el = tsField(n); if (el) { el.value = v; el.dispatchEvent(new Event("change", { bubbles: true })); } };
          const tsPut = (needle) => {
            const el = tsPara(needle);
            if (!el) return 0;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let node = null;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) { node = n; break; }
            }
            if (!node) return 0;
            const r = document.createRange();
            r.setStart(node, 1);
            r.setEnd(node, 1);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            el.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          // 那一枚制表符之后那截字的落点，按**这张纸自己**的 zoom 与左内边距折回纸像素（口径同 .scratch/tab-stops-dom.ts）。
          const tsSeg = (needle, idx) => {
            const el = tsPara(needle);
            const page = el?.closest(".page");
            if (!el || !page) return null;
            const ts = [...el.querySelectorAll("span.tab")];
            const t = ts[idx ?? 0];
            if (!t) return null;
            const z = parseFloat(page.style.getPropertyValue("zoom")) || 1;
            const base = page.getBoundingClientRect().left;
            const rg = document.createRange();
            const nxt = ts[(idx ?? 0) + 1] ?? null;
            if (nxt) { rg.setStartAfter(t); rg.setEndBefore(nxt); } else { rg.selectNodeContents(el); rg.setStartAfter(t); }
            const rects = [...rg.getClientRects()].filter((b) => b.width > 0);
            if (!rects.length) return null;
            const top = rects[0].top;
            const line = rects.filter((b) => Math.abs(b.top - top) <= 1);
            const L = Math.min(...line.map((b) => b.left));
            const R = Math.max(...line.map((b) => b.right));
            const to = (v) => round1((v - base) / z - (parseFloat(getComputedStyle(page).paddingLeft) || 0));
            return { left: to(L), right: to(R), mid: to((L + R) / 2) };
          };
          // 修订得按回关着：开着落这一击会顺带写一枚 w:pPrChange，那是修订那一族的形状，这一族不问它（量完按回原样）。
          const tsTrackBtn = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const tsTrackOn = () => (tsTrackBtn()?.classList.contains("on") ? 1 : 0);
          const tsTrack0 = tsTrackOn();
          const dOnEntry = tsDirty();
          if (tsTrack0) { tsTrackBtn().click(); await wait(320); }
          const dAfterTrack = tsDirty();
          // 进这一条时纸已经脏了：这一趟在它前面的腿（docPart 打字、画者拖动）各自验的正是「改了字节要报脏」，
          // 所以「取消/原样确定不许报脏」在这一格里没法按绝对口径问。先存一次把脏清零，后面每一格读数才都是绝对口径
          // （存下来的字节本来就是那些腿改过的，这一族最后仍按同一份盘子比对，不影响盘上那两条账）。
          let dBaseline = dAfterTrack;
          if (dAfterTrack) { fire("solo:save"); await wait(1400); dBaseline = tsDirty(); }
          const btn = tsBtn();
          const groupKids = btn ? [...(btn.parentElement?.children ?? [])] : [];
          const opened = [];
          // —— 1) 起手摆的是件里那一句：清单、三栏、默认格、那行小字，外加打包件的那一段样式 ——
          opened.push(tsPut("甲起"));
          btn?.click();
          await wait(260);
          const dStyle = tsDlg() ? getComputedStyle(tsDlg()) : null;
          const aStyle = tsField("ts-pos") ? getComputedStyle(tsDlg()?.querySelector(".solo-ts-apply") ?? document.body) : null;
          const read0 = {
            put: opened[0],
            rows: tsRows().map((o) => o.value).join(","),
            texts: tsRows().map((o) => String(o.textContent)).join("|"),
            pos: String(tsField("ts-pos")?.value ?? ""),
            val: String(tsDlg()?.querySelector("input[name='ts-val']:checked")?.value ?? ""),
            leader: String(tsField("ts-leader")?.value ?? ""),
            def: String(tsField("ts-default")?.value ?? ""),
            defTwips: tsTwips(String(tsField("ts-default")?.value ?? "")),
            unit: tsUnit(),
            hint: tsHint("ts-list"),
            boxW: dStyle ? round1(parseFloat(dStyle.width)) : 0,
            boxBorder: dStyle ? String(dStyle.borderTopWidth) + "/" + String(dStyle.borderTopColor) : "",
            applyBg: aStyle ? String(aStyle.backgroundColor) : "",
            btnTitle: String(btn?.title ?? "").slice(0, 8),
            btnIndex: groupKids.indexOf(btn ?? undefined) + 1,
            btnCount: groupKids.length,
          };
          tsCancel();
          await wait(200);
          const segJiaCancel = tsSeg("甲起", 0);
          const dAfterCancel = tsDirty();
          // —— 2) 填错的那一格：那行红字说出来、「确定」按下不落笔 ——
          btn?.click();
          await wait(240);
          const dAfterReopen = tsDirty();
          tsType("ts-pos", "不是一格数");
          await wait(200);
          const dAfterBadType = tsDirty();
          const badErr = tsErr();
          const badDisabled = tsDlg()?.querySelector(".solo-ts-apply")?.disabled === true;
          const segBeforeBad = tsSeg("甲起", 0);
          tsApply();
          await wait(240);
          const segAfterBad = tsSeg("甲起", 0);
          const badStillOpen = !!tsDlg();
          const dAfterBadApply = tsDirty();
          if (tsDlg()) tsCancel();
          await wait(200);
          const dAfterBadCancel = tsDirty();
          // —— 3) 原样确定：一字不动、不进撤销栈、不报脏 ——
          const dirtyPre = tsDirty();
          btn?.click();
          await wait(240);
          const segPreApply = tsSeg("甲起", 0);
          tsApply();
          await wait(260);
          const segPostApply = tsSeg("甲起", 0);
          const dirtyAfterNoop = tsDirty();
          // —— 4) 第一趟才落笔：加一枚 2880 居中、清掉 720 那一枚 ——
          const firstPass = read0.rows === "720,8306";
          let rowsAfterEdit = "";
          let segEdited = null;
          let segUndone = null;
          let segRedone = null;
          let cDef = null;
          let cDefAfter = null;
          let defEcho = "";
          let dAfterEdit = -1;
          let dAfterUndo = -1;
          let dAfterDef = -1;
          let rowsDing = "";
          let rowsDing2 = "";
          let leadEcho0 = "";
          let leadEcho1 = "";
          let leadEchoBack = "";
          let leadRow0 = "";
          let leadRowSet = "";
          let inkDingA = null;
          let inkDingB = null;
          let inkDingC = null;
          let dAfterLeader = -1;
          let dAfterLeaderBack = -1;
          if (firstPass) {
            btn?.click();
            await wait(240);
            tsType("ts-pos", tsText(2880));
            tsMark("center");
            tsClick("ts-set");
            await wait(200);
            tsPickRow(0);
            tsClick("ts-del");
            await wait(200);
            rowsAfterEdit = tsRows().map((o) => o.value).join(",");
            tsApply();
            await wait(300);
            segEdited = tsSeg("甲起", 0);
            dAfterEdit = tsDirty();
            tsRb("撤销")?.click();
            await wait(300);
            segUndone = tsSeg("甲起", 0);
            dAfterUndo = tsDirty();
            tsRb("重做")?.click();
            await wait(300);
            segRedone = tsSeg("甲起", 0);
            // 默认制表位：件级那一句改成 720（丙段整段没写位，纸上那一枚走的正是这一格）。
            btn?.click();
            await wait(240);
            cDef = tsSeg("丙起", 0);
            defEcho = String(tsField("ts-default")?.value ?? "");
            tsType("ts-default", tsText(720));
            tsApply();
            await wait(300);
            cDefAfter = tsSeg("丙起", 0);
            dAfterDef = tsDirty();
            // —— 5) 前导符那一档在装机壳里真按一下：丁段自己那枚 left 720 说的是「无」，而丁那一枚 Tab 正好停进这一格
            // （甲的 8306 带 dot，但甲那枚 Tab 停在 720/2880，从没落进 8306 那一格——墨只跟真正停进去的那一格，
            // 所以这一击的纸上证据必须落在丁身上，不能拿甲凑）。点上「点线」要铺墨，点回「无」要把墨撤干净。
            tsPut("丁起");
            btn?.click();
            await wait(240);
            rowsDing = tsRows().map((o) => o.value).join(",");
            leadEcho0 = String(tsField("ts-leader")?.value ?? "");
            leadRow0 = String(tsRows()[0]?.textContent ?? "");
            inkDingA = tsInk("丁起", 0);
            tsPick("ts-leader", "dot");
            await wait(140);
            tsClick("ts-set");
            await wait(220);
            leadEcho1 = String(tsField("ts-leader")?.value ?? "");
            leadRowSet = String(tsRows()[0]?.textContent ?? "");
            tsApply();
            await wait(320);
            inkDingB = tsInk("丁起", 0);
            dAfterLeader = tsDirty();
            // 反方向：重开一张，清单那一行回显的得是刚落盘的那一枚（dot），点回「无」后墨要归零。
            // 选区每次重画后都得重新按回丁那一格：落笔会换掉纸上的节点，指错了段就成了「清掉丁的制表位」。
            tsPut("丁起");
            btn?.click();
            await wait(240);
            rowsDing2 = tsRows().map((o) => o.value).join(",");
            leadEchoBack = String(tsField("ts-leader")?.value ?? "");
            tsPick("ts-leader", "");
            await wait(140);
            tsClick("ts-set");
            await wait(220);
            tsApply();
            await wait(320);
            inkDingC = tsInk("丁起", 0);
            dAfterLeaderBack = tsDirty();
            // 选区挪回甲：下面 readEnd 问的仍是甲那一份清单，别把这一族的读数换成丁的。
            tsPut("甲起");
            await wait(120);
          }
          const paper = { jia: tsSeg("甲起", 0), yi: tsSeg("乙起", 0), bing: tsSeg("丙起", 0), ding: tsSeg("丁起", 0), wu: tsSeg("戊起", 0) };
          btn?.click();
          await wait(240);
          const readEnd = { rows: tsRows().map((o) => o.value).join(","), def: String(tsField("ts-default")?.value ?? ""), defTwips: tsTwips(String(tsField("ts-default")?.value ?? "")), hint: tsHint("ts-list") };
          if (tsDlg()) tsCancel();
          await wait(200);
          if (tsTrack0) { tsTrackBtn()?.click(); await wait(300); }
          fire("solo:save");
          await wait(1400);
          tabStopsLeg = {
            firstPass: firstPass ? 1 : 0,
            read0,
            cancel: { gone: tsDlg() ? 1 : 0, segLeft: segJiaCancel ? segJiaCancel.left : null, dirty: dAfterCancel },
            bad: { errChars: badErr.length, disabled: badDisabled ? 1 : 0, moved: segBeforeBad && segAfterBad ? round1(segAfterBad.left - segBeforeBad.left) : null, open: badStillOpen ? 1 : 0 },
            dirty: { onEntry: dOnEntry, afterTrack: dAfterTrack, baseline: dBaseline, afterCancel: dAfterCancel, afterReopen: dAfterReopen, afterBadType: dAfterBadType, afterBadApply: dAfterBadApply, afterBadCancel: dAfterBadCancel, afterEdit: dAfterEdit, afterUndo: dAfterUndo, afterDef: dAfterDef, afterLeader: dAfterLeader, afterLeaderBack: dAfterLeaderBack },
            noop: { dirtyPre, dirtyAfterNoop, moved: segPreApply && segPostApply ? round1(segPostApply.left - segPreApply.left) : null, dialogsLeft: document.querySelectorAll(".solo-ts-dialog").length },
            edit: { rowsAfterEdit, segEdited, segUndone, segRedone, defEcho, cDef, cDefAfter },
            leader: { rowsDing, rowsDing2, echo0: leadEcho0, echo1: leadEcho1, echoBack: leadEchoBack, row0: leadRow0, rowSet: leadRowSet, inkA: inkDingA, inkB: inkDingB, inkC: inkDingC },
            paper,
            readEnd,
            ribbonButtons: document.querySelectorAll("#ribbon button").length,
            trackBefore: tsTrack0,
            trackBack: tsTrackOn(),
            savedDirty: tsDirty(),
          };
        }

        // #224：「显示原文」那一档纸面穿的是每条属性改动记录里的 prev（render.ts 走 doc.originalView()），
        // 所以读数的框也得说同一句。这一腿问三处同源：段落对话框里段前那一格、行距那一档、开始页对齐那四枚
        // 亮的是哪一枚；切回「显示修订」再问一遍，两档各自与纸面 text-align 对齐。
        const orDoc = (document.getElementById("title")?.textContent ?? "").includes("原文读数");
        let origLeg = null;
        if (orDoc) {
          const orPara = (needle) => [...document.querySelectorAll("#view .page .para")].find((x) => (x.textContent ?? "").includes(needle));
          const orPut = (needle) => {
            const el = orPara(needle);
            if (!el) return 0;
            const walk = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let node = null;
            for (let n = walk.nextNode(); n; n = walk.nextNode()) {
              if ((n.textContent ?? "").includes(needle)) {
                node = n;
                break;
              }
            }
            if (!node) return 0;
            const r = document.createRange();
            r.setStart(node, 0);
            r.setEnd(node, (node.textContent ?? "").length);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            el.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const orSelect = (tip) => [...document.querySelectorAll("#ribbon select")].find((x) => x.title === tip);
          const orView = (v) => {
            const s = orSelect("显示对文档的修改");
            if (!s) return 0;
            s.value = v;
            s.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          // 开始页对齐那四枚此刻亮着哪一枚：markState() 交 align，syncDocRibbon 按它点亮一枚。
          const orLit = () =>
            [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button.rb")]
              .filter((b) => ["左对齐", "居中", "右对齐", "两端对齐"].includes(b.getAttribute("title") ?? "") && b.classList.contains("on"))
              .map((b) => (b.textContent ?? "").trim())
              .join(",");
          const orDlg = () => document.querySelector(".solo-pg-dialog");
          const orField = (n) => orDlg()?.querySelector('[name="' + n + '"]') ?? null;
          const orPaper = (needle) => {
            const el = orPara(needle);
            return el ? getComputedStyle(el).textAlign : "?";
          };
          const orOpen = async (needle) => {
            const put = orPut(needle);
            await wait(320);
            const lit = orLit();
            const paper = orPaper(needle);
            const sel = orSelect("行距");
            if (sel) {
              sel.value = "more";
              sel.dispatchEvent(new Event("change", { bubbles: true }));
            }
            await wait(360);
            const d = orDlg();
            const box = orField("pg-value");
            return {
              put: put,
              lit: lit,
              paper: paper,
              open: d?.open === true ? 1 : 0,
              before: orField("pg-before")?.value ?? "?",
              after: orField("pg-after")?.value ?? "?",
              rule: orField("pg-rule")?.value ?? "?",
              value: box?.value ?? "?",
              valueOff: box ? (box.disabled ? 1 : 0) : -1,
            };
          };
          const orClose = async () => {
            orDlg()?.querySelector(".solo-pg-cancel")?.click();
            await wait(280);
            return orDlg() ? 0 : 1;
          };
          const orSteps = [];
          for (const v of ["marks", "original", "marks"]) {
            const set = orView(v);
            await wait(460);
            const o = await orOpen("被改过段距的一段");
            const c = await orClose();
            orSteps.push({ view: v, set: set, res: o, closed: c });
          }
          // 没记录那一段：两档同脸，读数与纸面都不该跟着档变——证明这不是把每一段都翻个面。
          const orCtrl = [];
          for (const v of ["marks", "original"]) {
            const set = orView(v);
            await wait(460);
            const c = await orOpen("没被改过的一段");
            orCtrl.push({ view: v, set: set, res: c, closed: await orClose() });
          }
          // 量完把档放回「显示修订」交给后面的腿。
          orView("marks");
          await wait(360);
          origLeg = { steps: orSteps, ctrl: orCtrl, dialogsLeft: document.querySelectorAll(".solo-pg-dialog").length, savedDirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0 };
        }
        // #225：同一条 originalView 也穿到表的读数上（doc-edit.ts 的 shownTable）。件由 .scratch/
        // make-original-table-read-smoke.ts 现拼：一张表把表/行/格三层记录都挂上，第二行一句没说。这一腿在三档之间拨
        // （marks → original → marks），每档问同一批地方：纸面（表那一圈的行内摆位/地/顶缝、卯格的地/顶缝/竖排/宽）、
        // 表格工具那两枚自旋格、「表格属性」对话框（对齐/左缩进/固定列宽/左边距/行对齐）、「边框与底纹」对话框在
        // **格**那一层（光标停在卯）与**表**那一层（拖选卯到午）的回显，外加「标题行重复」与竖向三枚亮的是哪一枚。
        // 读数框说的必须是纸面此刻正在画的那一句：原文档说记录里的 prev，显示修订档说现值。
        // 这一段整篇在模板字符串里：不写反引号、不写美元花括号。
        const otDoc = (document.getElementById("title")?.textContent ?? "").includes("原文表层");
        let origTableLeg = null;
        if (otDoc) {
          const otTable = () => [...document.querySelectorAll("#view .page table.doc-table")].find((t) => (t.textContent ?? "").includes("卯"));
          const otCell = (n) => [...(otTable()?.querySelectorAll("td") ?? [])].find((c) => (c.textContent ?? "").includes(n));
          const otSpan = (n) => otCell(n)?.querySelector("span")?.firstChild ?? null;
          const otSelect = (tip) => [...document.querySelectorAll("#ribbon select")].find((x) => x.title === tip);
          const otView = (v) => {
            const s = otSelect("显示对文档的修改");
            if (!s) return 0;
            s.value = v;
            s.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          const otBoxPick = () => [...document.querySelectorAll("#ribbon select")].find((s) => s.dataset && s.dataset.scope !== undefined);
          const otLeg = (value) => {
            const s = otBoxPick();
            if (!s) return 0;
            s.value = value;
            s.dispatchEvent(new Event("change", { bubbles: true }));
            return 1;
          };
          const otTab = (label) => [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === label);
          const otPaper = () => {
            const t = otTable();
            const c = otCell("卯");
            if (!t || !c) return null;
            const ts = getComputedStyle(t);
            const cs = getComputedStyle(c);
            const r = c.getBoundingClientRect();
            return {
              ml: t.style.marginLeft,
              mr: t.style.marginRight,
              tblBg: ts.backgroundColor,
              tblTop: ts.borderTopStyle + "/" + ts.borderTopWidth + "/" + ts.borderTopColor,
              cellBg: cs.backgroundColor,
              cellTop: cs.borderTopStyle + "/" + cs.borderTopWidth + "/" + cs.borderTopColor,
              va: cs.verticalAlign,
              w: Math.round(r.width * 10) / 10,
              h: Math.round(r.height * 10) / 10,
            };
          };
          const otPut = (n) => {
            const node = otSpan(n);
            if (!node) return 0;
            const rg = document.createRange();
            rg.setStart(node, 0);
            rg.collapse(true);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(rg);
            otCell(n)?.closest(".page")?.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const otAll = () => {
            const a = otSpan("卯");
            const z = otSpan("午");
            if (!a || !z) return 0;
            const rg = document.createRange();
            rg.setStart(a, 0);
            rg.setEnd(z, (z.textContent ?? "").length);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(rg);
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const otSpin = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] .rspin")]
            .find((x) => (x.textContent || "").trim().indexOf(label) === 0)?.querySelector("input");
          const otLit = () => {
            const vTitles = ["顶端对齐", "垂直居中", "底端对齐"];
            const v = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] button")]
              .filter((b) => vTitles.includes(b.getAttribute("title") ?? "") && b.classList.contains("on"))
              .map((b) => (b.getAttribute("title") ?? "").slice(0, 2))
              .join(",");
            const head = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] button")].find((b) => b.textContent === "标题行重复");
            return { v: v, head: head?.classList.contains("on") ? 1 : 0 };
          };
          const otProps = async () => {
            otTab("表格工具")?.click();
            await wait(200);
            const btn = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] button")].find((b) => b.textContent === "属性");
            if (btn) btn.click();
            await wait(300);
            const d = document.querySelector(".solo-tp-dialog");
            const f = (n) => d?.querySelector('[name="' + n + '"]') ?? null;
            return {
              open: d?.open === true ? 1 : 0,
              align: f("tp-align")?.value ?? "?",
              indent: f("tp-indent")?.value ?? "?",
              fixed: f("tp-fixed")?.checked ? 1 : 0,
              mleft: f("tp-m-left")?.value ?? "?",
              row: f("tp-row-align")?.value ?? "?",
            };
          };
          const otPropsClose = async () => {
            document.querySelector(".solo-tp-dialog .solo-tp-cancel")?.click();
            await wait(240);
            return document.querySelector(".solo-tp-dialog") ? 0 : 1;
          };
          const otBorder = async () => {
            otTab("开始")?.click();
            await wait(160);
            const opened = otLeg("more");
            await wait(280);
            const d = document.querySelector(".solo-pb-dialog");
            const f = (n) => d?.querySelector('[name="' + n + '"]') ?? null;
            const out = {
              opened: opened,
              open: d ? 1 : 0,
              scope: d?.querySelector(".solo-pb-scope b")?.textContent ?? "",
              top: f("pb-top")?.value ?? "?",
              sz: f("pb-top-sz")?.value ?? "?",
              co: f("pb-top-co")?.value ?? "?",
              ground: f("pb-shading")?.value ?? "?",
            };
            d?.querySelector(".solo-pb-cancel")?.click();
            await wait(240);
            out.gone = document.querySelector(".solo-pb-dialog") ? 0 : 1;
            return out;
          };
          const otSteps = [];
          for (const v of ["marks", "original", "marks"]) {
            const set = otView(v);
            await wait(460);
            const put = otPut("卯");
            await wait(280);
            const paper = otPaper();
            otTab("表格工具")?.click();
            await wait(200);
            const spins = { h: otSpin("高度")?.value ?? "?", w: otSpin("宽度")?.value ?? "?" };
            const lit = otLit();
            const props = await otProps();
            const propsClosed = await otPropsClose();
            // 格那一层：光标停在卯里（刚才那一格还在）。
            const cellBorder = await otBorder();
            // 表那一层：拖选卯到午，整张表铺满才算「表」。
            const all = otAll();
            await wait(280);
            const scope = otBoxPick()?.dataset?.scope ?? "";
            const tableBorder = await otBorder();
            otSteps.push({ view: v, set: set, put: put, paper: paper, spins: spins, lit: lit, props: props, propsClosed: propsClosed, cellBorder: cellBorder, all: all, scope: scope, tableBorder: tableBorder });
          }
          // 量完把档放回「显示修订」交给后面的腿（与 #224 同一条）。
          otView("marks");
          await wait(320);
          origTableLeg = {
            steps: otSteps,
            dialogsLeft: document.querySelectorAll(".solo-tp-dialog, .solo-pb-dialog").length,
            savedDirty: (document.getElementById("title")?.textContent ?? "").includes("●") ? 1 : 0,
          };
        }
        // #236 表格「排序」在装机壳里走一趟。件「排序-装机.docx」是一张带 w:tblHeader 的两列表（3/10/2），
        // 三次开合问三件事：取消不动行序也不动光标；降序把整行挪走（第二格的名字跟着键走）；再升序排回。
        // 行序与光标都从页面上读（光标是壳自己按模型落的，不是探针替它铺的），最终 save 由文件外的 python
        // 读 document.xml 核对行序与标题行那句 w:tblHeader。别的腿把这件改脏过，所以起手先存一次把 ● 归零。
        const sortDoc = (document.getElementById("title")?.textContent ?? "").includes("排序");
        let sortLeg = null;
        if (sortDoc) {
          fire("solo:save");
          await wait(900);
          const sTable = () => [...document.querySelectorAll("#view .page table.doc-table")].find((t) => (t.textContent ?? "").includes("键"));
          const sKeys = () => [...(sTable()?.querySelectorAll("tr") ?? [])].filter((tr) => !tr.classList.contains("solo-sheet-gap-row")).map((tr) => (tr.children[0]?.textContent ?? "").trim()).join(",");
          const sCaret = () => {
            const s = document.getSelection();
            const r = s && s.rangeCount ? s.getRangeAt(0) : null;
            if (!r) return { cell: "", collapsed: 0 };
            const n = r.startContainer;
            const el = n && n.nodeType === 1 ? n : n.parentElement;
            const td = el && typeof el.closest === "function" ? el.closest("td") : null;
            return { cell: td ? (td.textContent ?? "").trim() : "", collapsed: r.collapsed ? 1 : 0 };
          };
          const sPut = () => {
            const cell = [...(sTable()?.querySelectorAll("td") ?? [])].find((c) => (c.textContent ?? "").trim() === "3");
            const span = cell?.querySelector("span");
            if (!cell || !span || !span.firstChild) return 0;
            const r = document.createRange();
            r.setStart(span.firstChild, 1);
            r.collapse(true);
            const s = document.getSelection();
            s?.removeAllRanges();
            if (s) s.addRange(r);
            cell.focus?.();
            document.dispatchEvent(new Event("selectionchange"));
            return 1;
          };
          const sDlg = () => document.querySelector(".solo-sort-dialog");
          const sOpen = async () => {
            const bar = document.querySelector("#ribbon .rbar[data-tab='table']");
            if (bar?.hidden === true) [...document.querySelectorAll(".rtabs button")].find((b) => b.textContent === "表格工具")?.click();
            await wait(200);
            const btn = [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] button")].find((b) => b.textContent === "排序");
            const present = btn ? 1 : 0;
            const disabled = btn?.disabled === true ? 1 : 0;
            if (btn) btn.click();
            await wait(300);
            return { present: present, disabled: disabled, open: sDlg()?.open === true ? 1 : 0 };
          };
          const sEcho = () => {
            const d = sDlg();
            if (!d) return { open: 0 };
            const c = d.querySelector("[name='sort-column']");
            const k = d.querySelector("[name='sort-kind']");
            const o = d.querySelector("[name='sort-order']");
            const h = d.querySelector("[name='sort-header']");
            return { open: 1, cols: c ? c.options.length : 0, col: c ? c.value : "?", kind: k ? k.value : "?", order: o ? o.value : "?", header: h && h.checked ? 1 : 0 };
          };
          const sCancel = async () => {
            const b = sDlg()?.querySelector(".solo-sort-cancel");
            if (!b) return 0;
            b.click();
            await wait(260);
            return sDlg() ? 0 : 1;
          };
          const sApply = async (order) => {
            const o = sDlg()?.querySelector("[name='sort-order']");
            if (!o) return 0;
            o.value = order;
            o.dispatchEvent(new Event("change", { bubbles: true }));
            sDlg()?.querySelector(".solo-sort-apply")?.click();
            await wait(560);
            return sDlg() ? 0 : 1;
          };
          const sTrack = () => [...document.querySelectorAll("#ribbon button.rb")].find((b) => b.textContent === "修订");
          const sTrack0 = sTrack()?.classList.contains("on") ? 1 : 0;
          if (sTrack0) {
            sTrack()?.click();
            await wait(320);
          }
          const put = sPut();
          await wait(280);
          const orderBefore = sKeys();
          const opened = await sOpen();
          const echo = sEcho();
          const cancelled = await sCancel();
          const orderAfterCancel = sKeys();
          const caretAfterCancel = sCaret();
          const dirtyAfterCancel = dirtyNow();
          const reopened = await sOpen();
          const appliedDesc = await sApply("desc");
          const orderDesc = sKeys();
          const caretDesc = sCaret();
          const dirtyDesc = dirtyNow();
          const reopenedAsc = await sOpen();
          const appliedAsc = await sApply("asc");
          const orderAsc = sKeys();
          const caretAsc = sCaret();
          const dirtyAsc = dirtyNow();
          fire("solo:save");
          await wait(900);
          sortLeg = {
            table: sTable() ? 1 : 0, put: put, orderBefore: orderBefore, opened: opened, echo: echo,
            cancelled: cancelled, orderAfterCancel: orderAfterCancel, caretAfterCancel: caretAfterCancel, dirtyAfterCancel: dirtyAfterCancel,
            reopened: reopened, appliedDesc: appliedDesc, orderDesc: orderDesc, caretDesc: caretDesc, dirtyDesc: dirtyDesc,
            reopenedAsc: reopenedAsc, appliedAsc: appliedAsc, orderAsc: orderAsc, caretAsc: caretAsc, dirtyAsc: dirtyAsc,
            dialogsLeft: document.querySelectorAll(".solo-sort-dialog").length,
          };
          if (sTrack0) {
            sTrack()?.click();
            await wait(320);
          }
        }
        // 上面那些落点把光标留在了格里；后面那三条断命令与图片要按这一趟原本的位置落在正文里，所以先放回纸的末尾。
        const pageNow = document.querySelector("#view .page");
        const lastBlock = pageNow?.children[pageNow.children.length - 1];
        if (lastBlock) {
          const br = document.createRange();
          br.selectNodeContents(lastBlock);
          br.collapse(false);
          const bs = document.getSelection();
          bs?.removeAllRanges();
          if (bs) bs.addRange(br);
          lastBlock.focus?.();
        }
        await wait(160);
        fire("solo:save");
        await wait(900);
        const tblBoxCmd = {
          tblSelected, tblScope, tblLegs, tblLegNames, tblTitle, tblClicked,
          tblBeforeSolid, tblBefore, tblAfterSolid, tblAfter, tblScopeKept, tblDirty,
          tblReclick, tblReclickSame, tblUndo, tblAfterUndo, tblScopeAfterUndo, tblSpanAfterUndo,
          tblRedo, tblAfterRedo, cellScope, cellLegs, cellClicked, cellTopNow,
          dlgOpened, dlgScope, dlgRows, dlgGround, dlgEcho, dlgPick, dlgApply, dlgGone, cellLeftNow, dlgScopeKept,
          g1Opened, cellGroundRow, cellGroundEcho, cellGroundPick, g1Submit, cellGroundNow, g1Gone,
          tblGroundScope, g2Opened, tblGroundRow, tblGroundEcho, tblGroundPick, g2Submit,
          tblGroundNow, cellGroundStill, siblingGround, g2Gone,
          revPaneOpened, revInRows, revInKinds, revTracking, revTableScope, revTopBefore, revTableStroke, revTopAfter,
          revRowsTable, revCellScope, revCellBottomBefore, revCellStroke, revCellBottomAfter,
          revRowsAll, revKindsAll, revAuthor, revCellAuthor, revCountText,
          revDecide, revAcceptHit, revRowsAfterAccept, revTopAfterAccept,
          revRejectHit, revRowsAfterReject, revCellBottomAfterReject, revCellTopStill, revCellGroundStill,
          revPaneClosed, revDirty,
        };

        // WPS 把三条断命令挂在同一个下拉里（wpsentries.kuip 的 CT_Insert.RB_Pages.InsertBreaks.*），所以这里三颗
        // 都点。标记是 br[data-brk]：从前这条读的是 [style*="page-break"]，那是 #104 之前的画法，改完之后它恒为 0
        // —— 冒烟件自己在报「分页符没插进去」。（这段跑在模板字符串里，注释不能带反引号，反引号会终止字符串。）
        const brShapes = () => [...document.querySelectorAll("#view .page br[data-brk]")].map((b) => b.getAttribute("data-brk"));
        const newShape = (before, after) => {
          const rest = before.slice();
          for (const s of after) {
            const i = rest.indexOf(s);
            if (i < 0) return s;
            rest.splice(i, 1);
          }
          return "";
        };
        const s0 = brShapes();
        rb("分页符")?.click();
        await wait(150);
        const s1 = brShapes();
        const shapePage = newShape(s0, s1);
        rb("换行符")?.click();
        await wait(150);
        const s2 = brShapes();
        const shapeLine = newShape(s1, s2);
        rb("分栏符")?.click();
        await wait(150);
        const s3 = brShapes();
        const shapeColumn = newShape(s2, s3);
        const breaks = s3.length - s0.length;
        const breakShapes = { page: shapePage, line: shapeLine, column: shapeColumn };
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
        // 收笔不等 blur：这只壳里窗口不在前台时 blur 事件根本不来，刚打进去的字就只留在 DOM 里。
        // 走生产路径的存盘（solo:save → saveCommand → editor.commit()）把它交回模型——真用户按 ⌘S 也是这条路。
        fire("solo:save");
        await wait(600);
        // 带子跟着纸走：DOM 门禁跑的是 src/ui/index.html 生成的一份副本，这一腿要证明的是 shipping 的这只盒子。
        // 只报结构与「纸数对不上」这类数，不报任何一个可能是文档内容的字。纸盒要现问：这一段前面已经重排过，
        // 探针开头抓住的那一个 page 早就离开文档了。
        const bandPaper = document.querySelector("#view .page");
        const bandClones = [...document.querySelectorAll("#view .solo-hf-clone")];
        const bandChips = (el) => [...el.querySelectorAll("[data-field]")].map((c) => c.textContent);
        const bandReal = (which) => [...document.querySelectorAll("#view .hf-" + which)].find((el) => !el.classList.contains("solo-hf-clone"));
        const bandHead = bandReal("header");
        const bandFoot = bandReal("footer");
        const bandStripEls = [...document.querySelectorAll("#view .solo-sheet-strip")];
        // 条带与复件的摆位写的是「纸的像素」，zoom 把纸连同缝一起放大。这一腿报的数一律换算回纸的像素：
        // 漏乘或多乘 zoom，stripSpace 就会是 false、缝距就会差出一大截，纸盒里量的与 CSS 写的对不上。
        const bandZoom = parseFloat(bandPaper?.style.getPropertyValue("zoom")) || 1;
        const bandPaperTop = bandPaper?.getBoundingClientRect().top ?? 0;
        const bandLocal = (el) => Math.round(((el.getBoundingClientRect().top - bandPaperTop) / bandZoom) * 10) / 10;
        // 复件按各自那一张的缝摆：报的是「复件的字到缝沿的实际间距」，节说的那个距离不进这一段（那是模型里的数）。
        const bandGap = (a, b, which) => {
          const r = a.getBoundingClientRect();
          const s = b.getBoundingClientRect();
          const pad = (parseFloat(getComputedStyle(a).paddingTop) || 0) * bandZoom;
          const padB = (parseFloat(getComputedStyle(a).paddingBottom) || 0) * bandZoom;
          return Math.round(((which === "header" ? r.top + pad - s.bottom : s.top - (r.bottom - padB)) / bandZoom) * 10) / 10;
        };
        const bandHeads = bandClones.filter((el) => el.classList.contains("hf-header"));
        const bandFeet = bandClones.filter((el) => el.classList.contains("hf-footer"));
        // #157：复件按**这一张所属的节**取部件。这一读数报的是「这一份复件照的是哪一份」，报的是部件在件里的
        // 名字（orig = 住在纸里那一份可点的原件，其余是 word/headerN.xml 这一类母本名），不是任何一个文档内容的字。
        // 认法只有一句：比第一段——复件不带命令条（hf-bar 那一段摘掉），拿整段比会被这一处差异顶成「条条都不同」。
        // 只有一节的件里数组全是 orig；按节取通了才会出现第二种名字，而且它只该出现在跨进第二节之后的那几张上。
        const bandFirstP = (el) => (el && el.querySelector(".hf-p") ? el.querySelector(".hf-p").textContent || "" : "");
        const bandSrcEls = [...document.querySelectorAll("#view .solo-hf-src")];
        const bandLabel = (el, visible, kind) => {
          const t = bandFirstP(el);
          if (visible && t === bandFirstP(visible)) return "orig";
          const src = bandSrcEls.find((x) => x.dataset.hfKind === kind && bandFirstP(x) === t);
          return src ? src.dataset.hfPart || "?" : "?";
        };
        const bandHeadSrc = bandHeads.map((el) => bandLabel(el, bandHead, "header"));
        const bandFootSrc = bandFeet.map((el) => bandLabel(el, bandFoot, "footer"));
        // #157：带子比页边距高时，正文得往下让位。让位的数由摆位那一趟写在纸盒上——两枚 CSS 变量管首张与末张
        // （纸的 padding 就是那两张的边距带），两枚 data 属性按节列出每一张上沿要推多远（铺排的容量、缝距的推距、
        // 状态栏的纵向位置都读这两枚）。这一腿问的是 shipping 的这只盒子里「两处是不是同一个数、 padding 是不是
        // 真的被顶开了」，而不是纸盒外那一份副本。换算一律：视口像素减纸盒上沿、除 zoom 折回纸的像素、乘 15 折回
        // 缇（CDP 标定 15.0001 缇每像素，误差落不进 1 缇）。
        const bandPaperRect = bandPaper?.getBoundingClientRect();
        const bandPaperBottom = bandPaperRect ? bandPaperRect.bottom : 0;
        const bandIsDecor = (el) =>
          el.classList.contains("solo-sheet-strip") ||
          el.classList.contains("solo-hf-clone") ||
          el.classList.contains("solo-hf-master") ||
          el.classList.contains("doc-page-border");
        // 「每一条都……」这种判定在空清单上恒真：那一条腿没量过任何东西，却报出一个 true。空的时候就报
        // 「无X(这一格不算数)」——读日志的人要的是这一格到底有没有主语，不是一个好看的布尔。
        const bandAll = (label, els, test) => (els.length === 0 ? "无" + label + "(这一格不算数)" : els.every(test));
        const bandBody = [...(bandPaper ? bandPaper.children : [])].filter((el) => !bandIsDecor(el));
        const bandTw = (px) => Math.round(px * 15);
        const bandLocalPx = (viewportPx) => (viewportPx - bandPaperTop) / bandZoom;
        const bandVarPx = (name) => {
          const v = bandPaper ? getComputedStyle(bandPaper).getPropertyValue(name) : "";
          const n = parseFloat(v);
          return Number.isFinite(n) ? n : null;
        };
        const bandListPx = (attr) => {
          const v = bandPaper ? bandPaper.getAttribute(attr) || "" : "";
          return v.split(",").map((s) => parseFloat(s)).filter((n) => Number.isFinite(n));
        };
        // 每一张的首块正文离这一张的上沿多远：让位进没进铺排，读的就是这一串。第 0 项是首张。
        const bandSheetTops = [bandPaperTop].concat(bandStripEls.map((el) => el.getBoundingClientRect().bottom));
        const bandSheetBodyStart = bandSheetTops.map((t) => {
          const el = bandBody.find((b) => b.getBoundingClientRect().top >= t - 1);
          return el ? bandTw((el.getBoundingClientRect().top - t) / bandZoom) : -1;
        });
        // 原件守的是这一叠纸的首尾：报「带子的第一段离纸沿多少缇」与「它是否落在该落的那一张里」。
        const bandTextEdge = (el, which) => {
          const p = el ? el.querySelector(".hf-p") : null;
          if (!p || !bandPaperRect) return null;
          const r = p.getBoundingClientRect();
          return which === "header"
            ? { top: bandTw(bandLocalPx(r.top)), bottom: bandTw(bandLocalPx(r.bottom)), pastPaperBottom: r.bottom > bandPaperBottom }
            : { top: bandTw((bandPaperBottom - r.top) / bandZoom), bottom: bandTw((bandPaperBottom - r.bottom) / bandZoom), pastPaperTop: r.top < bandPaperTop };
        };
        const bandHeadText = bandTextEdge(bandHead, "header");
        const bandFootText = bandTextEdge(bandFoot, "footer");
        const docBands = {
          strips: bandStripEls.length,
          clones: bandClones.length,
          // 复件是纸内的装饰，两份原件也住在纸里（守这一叠纸的首尾），复件不顶替它们。
          clonesInsidePaper: bandClones.filter((el) => bandPaper?.contains(el)).length,
          originalsInsidePaper: [bandHead, bandFoot].filter((el) => el && bandPaper?.contains(el)).length,
          // 原件 resting 态不可编辑，这一句只来自渲染写的 contenteditable=false：带子住在**可编辑的纸盒里**，
          // 少这一句它就继承成可编辑，editHF 开头那句 isContentEditable 判断直接 return——带子点不开、打进去的
          // 字落进正文、sync 又把它当装饰跳过，两边一起丢。
          originalsRestLocked: [bandHead, bandFoot].every((el) => !!el && el.isContentEditable === false),
          // 两个口径分开报：attribute 是渲染写的那一句（冒烟自己把带子点开过就会是 true，那是这一腿的副作用），
          // isContentEditable 才是继承过之后的真值。只报后者会把「烟测没点回原处」报成「渲染忘了锁」。
          mastersEditableAttr: [bandHead, bandFoot].map((el) => (el ? String(el.getAttribute("contenteditable")) : "none")),
          masters: document.querySelectorAll("#view .solo-hf-master").length,
          // 复件不进编辑、不带「插入页码」那枚按钮、也不被算进正文。空清单上的 every 恒真，那一格根本没测
          // 过东西——报「无X(这一格不算数)」而不是报一个舒服的 true。
          clonesLocked: bandAll("复件", bandClones, (el) => el.isContentEditable === false && el.querySelector(".hf-page") === null && !el.hasAttribute("data-hf")),
          cloneCopies: bandHeads.length + bandFeet.length,
          // 每一条缝：CSS 写的上沿 与 纸盒里量回来的上沿（都按纸的像素）。同一趟画的条带与复件守同一个空间。
          stripSpace: bandStripEls.map((el) => [parseFloat(el.style.top), bandLocal(el)]),
          stripSpaceOk: bandAll("页缝条带", bandStripEls, (el) => Math.abs(parseFloat(el.style.top) - bandLocal(el)) <= 2),
          headGaps: bandHeads.map((el, i) => (bandStripEls[i] ? bandGap(el, bandStripEls[i], "header") : null)),
          footGaps: bandFeet.map((el, i) => (bandStripEls[i] ? bandGap(el, bandStripEls[i], "footer") : null)),
          chips: bandClones.flatMap(bandChips).concat(bandHead ? bandChips(bandHead) : [], bandFoot ? bandChips(bandFoot) : []),
          chipsNumeric: bandAll(
            "页码字面",
            bandClones.concat(bandHead ? [bandHead] : [], bandFoot ? [bandFoot] : []).flatMap(bandChips),
            (t) => /^[0-9]+$/.test(t || ""),
          ),
          headChips: bandHead ? bandChips(bandHead) : [],
          footChips: bandFoot ? bandChips(bandFoot) : [],
          editableHead: bandHead?.isContentEditable,
          editableFoot: bandFoot?.isContentEditable,
          bandHeadSrc,
          bandFootSrc,
          bandHeadSects: new Set(bandHeadSrc).size,
          bandFootSects: new Set(bandFootSrc).size,
          bandSrcParts: bandSrcEls.map((el) => (el.dataset.hfPart || "") + ":" + (el.dataset.hfKind || "")),
          bandSources: bandSrcEls.length,
          // 母本躲在纸盒外：不可编辑、不带 hf 这个名字（带子的字面样式与编辑接单都只认 .hf 那一条规则）。
          bandSourcesLocked: bandAll("母本", bandSrcEls, (el) => el.style.display === "none" && el.isContentEditable === false && !el.classList.contains("hf") && !el.hasAttribute("data-hf")),
          // 打字进原件、复件只是照它的字画：纸盒的直接孩子里，正文块与装饰（条带、复件、两份原件、页面边框）
          // 各占几枚，且装饰必须全排在正文之后——indexBlocks 按 children[i] 与 blocks[i] 对位认领，装饰插进
          // 正文中间就等于让正文全体错位。
          paperKids: bandPaper ? bandPaper.children.length : 0,
          paperBodyKids: bandBody.length,
          paperDecor: bandPaper ? bandPaper.children.length - bandBody.length : 0,
          decorCounts: [
            bandStripEls.length,
            bandClones.length,
            document.querySelectorAll("#view .solo-hf-master").length,
            document.querySelectorAll("#view .page > .doc-page-border").length,
          ],
          decorTrailing: (() => {
            const kids = [...(bandPaper ? bandPaper.children : [])];
            const n = kids.length - bandBody.length;
            return n === 0 || kids.slice(-n).every(bandIsDecor);
          })(),
          // #157 让位：CSS 变量（首张的上下内边距由它顶开）与按节的属性清单问的是同一趟量出来的数；
          // 而纸的上/下内边距必须等于「节说的边距 + 这一端让出的那一截」——两处不是同一个数，铺排的容量与
          // 画出来的纸就开始各说各话。
          yieldTopPx: bandVarPx("--band-yield-t"),
          yieldBottomPx: bandVarPx("--band-yield-b"),
          yieldTopList: bandListPx("data-band-yield-t"),
          yieldBottomList: bandListPx("data-band-yield-b"),
          yieldVarMatchesList: (() => {
            const t = bandListPx("data-band-yield-t");
            const b = bandListPx("data-band-yield-b");
            const vt = bandVarPx("--band-yield-t");
            const vb = bandVarPx("--band-yield-b");
            return vt !== null && vb !== null && t[0] === vt && b[0] === vb;
          })(),
          paddingCarriesYield: (() => {
            if (!bandPaper) return null;
            const csp = getComputedStyle(bandPaper);
            const declaredTop = parseFloat(bandPaper.style.getPropertyValue("--pad-t")) || 0;
            const declaredBottom = parseFloat(bandPaper.style.getPropertyValue("--pad-b")) || 0;
            const yT = bandVarPx("--band-yield-t") || 0;
            const yB = bandVarPx("--band-yield-b") || 0;
            return (
              Math.abs((parseFloat(csp.paddingTop) || 0) - (declaredTop + yT)) <= 1 &&
              Math.abs((parseFloat(csp.paddingBottom) || 0) - (declaredBottom + yB)) <= 1
            );
          })(),
          // 每一张的首块正文离这一张上沿多远（缇）：第 0 项是首张。带子不超过边距线时这一串就是边距。
          bodyStartPerSheetTw: bandSheetBodyStart,
          headTextTw: bandHeadText,
          footTextTw: bandFootText,
          // 带子的字与正文不相压：页眉那一段收在正文第一块之前，页脚那一段起在正文最后一块之后。
          bandClearsBody: (() => {
            const first = bandBody[0];
            const last = bandBody[bandBody.length - 1];
            const hp = bandHead ? bandHead.querySelector(".hf-p") : null;
            const fp = bandFoot ? bandFoot.querySelector(".hf-p") : null;
            if (!first || !last || !hp || !fp) return null;
            return (
              hp.getBoundingClientRect().bottom <= first.getBoundingClientRect().top + 1 &&
              fp.getBoundingClientRect().top >= last.getBoundingClientRect().bottom - 1
            );
          })(),
          zoom: bandPaper?.style.getPropertyValue("zoom") || "",
        };
        // #156：换缩放档问的还是同一张纸。Ribbon 的 ⊕ 每步 +10%，五步到 150%；缝的道数、复件份数、每一条缝在
        // 纸上的位置（都换算回纸的像素）都得与 100% 同一份。铺排那趟曾拿视口像素的行盒去比纸的像素的容量，150%
        // 下纸数虚增。DOM 门禁（.scratch/hf-dom.ts）跑的是 src/ui/index.html 的生成副本，这一腿证的是这只 shipping 的盒子。
        const bandAt = () => {
          const paper = document.querySelector("#view .page");
          const els = [...document.querySelectorAll("#view .solo-sheet-strip")];
          const z = parseFloat(paper?.style.getPropertyValue("zoom")) || 1;
          const top0 = paper ? paper.getBoundingClientRect().top : 0;
          // 现问：换档重画过一整趟，条带与复件都是这一趟新造的节点（原件与母本还在），拿 docBands 之前抓住的
          // 那一份去认「照的是哪一份」会把复件与新的条带错开——一律现问。
          const vis = (which) => [...document.querySelectorAll("#view .hf-" + which)].find((el) => !el.classList.contains("solo-hf-clone"));
          const srcOf = (which) => [...document.querySelectorAll("#view .solo-hf-clone.hf-" + which)].map((el) => bandLabel(el, vis(which), which));
          return {
            label: document.querySelector("#ribbon .grp button[title='恢复 100%']")?.textContent,
            zoom: Math.round(z * 100) / 100,
            strips: els.length,
            clones: document.querySelectorAll("#view .solo-hf-clone").length,
            tops: els.map((el) => Math.round(((el.getBoundingClientRect().top - top0) / z) * 10) / 10),
            // #157：换档重画之后，每一张照的还是自己那一节的部件吗（母本的枚数与名字也得一字不改）。
            headSrc: srcOf("header"),
            footSrc: srcOf("footer"),
            srcParts: [...document.querySelectorAll("#view .solo-hf-src")].map((el) => el.dataset.hfPart || "?"),
          };
        };
        const sameTops = (a, b) => a.length === b.length && a.every((v, i) => Math.abs(v - b[i]) <= 2);
        const sameList = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);
        const zoom100 = bandAt();
        for (let i = 0; i < 5; i++) {
          rb("⊕")?.click();
        }
        await wait(800);
        const zoom150 = bandAt();
        rb(zoom150.label)?.click();
        await wait(800);
        const zoomBack = bandAt();
        const zoomBands = {
          at100: zoom100,
          at150: zoom150,
          back100: zoomBack,
          bandsKept:
            sameList(zoom100.headSrc, zoom150.headSrc) &&
            sameList(zoom100.footSrc, zoom150.footSrc) &&
            sameList(zoom100.headSrc, zoomBack.headSrc) &&
            sameList(zoom100.footSrc, zoomBack.footSrc) &&
            sameList(zoom100.srcParts, zoomBack.srcParts),
          same:
            zoom100.zoom === 1 &&
            zoom150.zoom === 1.5 &&
            zoomBack.zoom === 1 &&
            zoom150.strips === zoom100.strips &&
            zoom150.clones === zoom100.clones &&
            sameTops(zoom150.tops, zoom100.tops) &&
            zoomBack.strips === zoom100.strips &&
            zoomBack.clones === zoom100.clones &&
            sameTops(zoomBack.tops, zoom100.tops),
        };
        // #158：每一道缝都得有东西把它切开。条带的道数 == 垫片**落点**的道数（段落/格内空 span 与垫片行
        // 按位置去重；#234 的行内那一刀按格各垫一枚、同一道缝不再数成几道），且状态栏「共 N 页」报的 N 与
        // 条带数 + 1 同一份。旧写法在缝落进两枚表行之间时既不垫也不画（垫片进表格会改行高），于是纸数照旧往上涨、
        // 画出来的纸却少一道缝——两张纸糊在一起。DOM 门禁钉的是垫片行的形状，这一腿钉的是 shipping 盒子里
        // 「纸数与画出来的缝」不能再各说各话。
        const seamSplit = (() => {
          const spans = document.querySelectorAll("#view .page .solo-sheet-gap").length;
          const rows = document.querySelectorAll("#view .page tr.solo-sheet-gap-row").length;
          const strips = document.querySelectorAll("#view .page .solo-sheet-strip").length;
          // 一道缝可以由几枚垫片切：行内那一刀按格垫（#234，一格一枚、都落在同一条刀口上），按落点去重才是
          // 缝的道数。旧式相等（strips === spans + rows）在那个形状上会把一道缝数成两道。
          const cuts = [...document.querySelectorAll("#view .page .solo-sheet-gap, #view .page tr.solo-sheet-gap-row")].map((el) => Math.round(el.getBoundingClientRect().top));
          const seamCuts = [...new Set(cuts)].length;
          const cell = [...document.querySelectorAll(".solo-doc-status-loc")].find((b) => (b.title || "").indexOf("页数") === 0);
          const m = /([0-9]+)/.exec(cell ? cell.textContent || "" : "");
          const n = m ? Number(m[1]) : -1;
          return { spans: spans, rows: rows, strips: strips, seams: seamCuts, statusPages: n, cutEverySeam: strips === seamCuts, pagesMatch: n > 0 && strips + 1 === n };
        })();
        const docInsert = {
          tableRows,
          tableCols,
          breaks,
          breakShapes,
          docPics,
          hfText,
          docBands,
          zoomBands,
          seamSplit,
          // #174：文件里那句「字距 82 缇」在装机件里也得把墨推开。逐枚 span 量计算样式——「normal」是「这句没说」，
          // 其余每一枚都得是那一句换算出来的像素（96dpi 下 1pt = 4/3px，所以 82 缇 = 5.46667px、-6 缇 = -0.4px；
          // 说了「w:val=0」的那一枚 Blink 交回 normal，.scratch/char-spacing-dom.ts 的 zero_ruler 自己量过）。
          // 夹具 /tmp/char-spacing-smoke.docx（.scratch/make-char-spacing-smoke.ts）里该画上的是五枚：自己说过
          // 的三枚 + 样式借来的两枚。DOM 闸门量的是同一份夹具的墨宽差，这一腿只问装出去的盒子里画没画。
          charSpacing: (() => {
            const painted = [...document.querySelectorAll("#view .page span")]
              .map((s) => getComputedStyle(s).letterSpacing)
              .filter((v) => v !== "normal" && v !== "");
            return { painted: painted.length, values: [...new Set(painted)].join(",") };
          })(),
          // #217：w:fitText（调整宽度）那一组在装机件里也摊到文件说的宽度上。按「段 + data-fit」拼组，报的是
          // 每组几枚、文件说的缇数、量出来的宽度（一个十进制）与 run 自己那句 --fit-char-ls（没有就是空）。
          // 摊出来的 --fit-ls 不报：它是量出来的，不进模型；同组各枚本应同值，组上的宽度已经把它量进了。
          fitText: (() => {
            const paras = [...document.querySelectorAll("#view .page .para")];
            const groups = new Map();
            let lone = 0;
            for (const el of document.querySelectorAll("#view .page [data-fit]")) {
              const at = paras.indexOf(el.closest(".para"));
              const name = el.getAttribute("data-fit");
              // 没有号的 run 自成一摊（与 applyFitText 同一条判据）：同段里两枚空号不许挤成一组。
              const key = at + "/" + (name === "" ? "#" + lone++ : name);
              const g = groups.get(key) || { n: 0, tw: el.getAttribute("data-fit-tw"), w: 0, own: "" };
              g.n += 1;
              g.w += el.getBoundingClientRect().width;
              const own = el.style.getPropertyValue("--fit-char-ls").trim();
              if (own) g.own = own;
              groups.set(key, g);
            }
            return [...groups.values()].map((g) => g.n + "x" + g.tw + "=" + g.w.toFixed(1) + (g.own ? "/" + g.own : ""));
          })(),
          // #218：w:tblPr/w:tblpPr 浮动表在装机件里钉在哪——报**纸的像素**里的 left/top（视口读数减纸盒上沿、
          // 除回 zoom）、宽高、data-tblp-page 与文件说的那两个缇数，外加每张表的 computed position。
          // 比对留给 .scratch/tblppr-install-check.sh：纸高 16838 缇 = 1122.53px、页缝 24px、1117 缇 = 74.47px、
          // 10912 缇 = 727.47px 都是它自己算的。不报文档正文，只报几何。
          tblpPr: (() => {
            const p = document.querySelector("#view .page");
            if (!p) return [];
            const zz = Number.parseFloat(p.style.getPropertyValue("zoom")) || 1;
            const box = p.getBoundingClientRect();
            const r2 = (n) => Math.round(n * 100) / 100;
            return [...p.querySelectorAll("table")].map((el) => {
              const r = el.getBoundingClientRect();
              // 首格首句：装机件的判据按件里给表起的名字认表，不按张数——烟测自己后面还会往正文尾插表。
              const cell = el.querySelector("td, th") || el;
              return {
                n: (cell.textContent || "").replace(/[\\s\\u00a0]+/g, " ").trim().slice(0, 8),
                px: el.getAttribute("data-tblp-x") || "",
                py: el.getAttribute("data-tblp-y") || "",
                page: el.getAttribute("data-tblp-page") || "",
                pos: getComputedStyle(el).position,
                l: r2((r.left - box.left) / zz),
                t: r2((r.top - box.top) / zz),
                w: r2(r.width / zz),
                h: r2(r.height / zz),
              };
            });
          })(),
        };

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
        const pptPart = { slides: document.querySelectorAll("#view .thumb").length, stage: !!document.querySelector("#view .stage"), title: document.getElementById("title")?.textContent, tabs: [...document.querySelectorAll("#ribbon .rtabs button")].map((t) => t.textContent).join("/"), homeCaps: [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] .grp")].map((g) => g.dataset.cap || "-").join("/"), homeIcons: document.querySelectorAll("#ribbon .rbar[data-tab='home'] button.rb.ico .ic svg path").length, homeIconText: [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button.rb.ico")].every((b) => b.textContent === b.querySelector(".lb")?.textContent), undoInRibbon: [...document.querySelectorAll("#ribbon .rbar button.rb")].filter((b) => b.textContent === "撤销" || b.textContent === "重做").length, viewCaps: [...document.querySelectorAll("#ribbon .rbar[data-tab='view'] .grp")].map((g) => g.dataset.cap || "-").join("/") };
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
        // Views: the sorter paints every page, a double click opens that page in the
        // normal view, the notes page shows the same slide plus its notes, and each
        // button reports which view is live.
        const viewBtn = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='view'] button.rb")].find((b) => b.textContent === label);
        viewBtn("幻灯片浏览")?.click();
        await wait(220);
        const sorterView = { painted: !!document.querySelector("#view .deck.sorter"), cells: document.querySelectorAll("#view .sorter-cell").length, pressed: viewBtn("幻灯片浏览")?.getAttribute("aria-pressed") };
        document.querySelectorAll("#view .sorter-cell")[1]?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
        await wait(220);
        const sorterBack = { normal: !!document.querySelector("#view .stage"), picked: [...document.querySelectorAll("#view .thumb")].findIndex((t) => t.classList.contains("on")) };
        viewBtn("备注页")?.click();
        await wait(220);
        const notesBody = document.querySelector("#view .notes-sheet-text");
        const notesView = { painted: !!document.querySelector("#view .notes-sheet"), editable: notesBody?.getAttribute("contenteditable") };
        if (notesBody) {
          notesBody.focus();
          notesBody.textContent = "口头备注落在盘上";
          notesBody.dispatchEvent(new Event("input", { bubbles: true }));
          notesBody.blur();
          await wait(120);
          notesView.text = notesBody.textContent;
          notesView.empty = notesBody.dataset.empty;
        }
        // Reading view: one slide fills the stage, the bar pages it, and Esc returns.
        viewBtn("阅读视图")?.click();
        await wait(220);
        const readingSlide = document.querySelector("#view .reading-slide");
        const readingStage = document.querySelector("#view .reading-stage");
        const hs = readingSlide?.getBoundingClientRect();
        const st = readingStage?.getBoundingClientRect();
        const readingView = {
          painted: !!document.querySelector("#view .deck.reading"),
          pressed: viewBtn("阅读视图")?.getAttribute("aria-pressed"),
          count: document.querySelector("#view .reading-count")?.textContent,
          fitted: !!hs && !!st && hs.width <= st.width && hs.width > st.width * 0.3,
        };
        window.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight" }));
        await wait(150);
        readingView.next = document.querySelector("#view .reading-count")?.textContent;
        window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
        await wait(220);
        readingView.left = !!document.querySelector("#view .stage");
        const pptViews = { sorter: sorterView, back: sorterBack, notes: notesView, reading: readingView, normal: !!document.querySelector("#view .stage"), pressed: viewBtn("普通视图")?.getAttribute("aria-pressed") };
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
        rb("撤销")?.click();
        await wait(100);
        const undoTitle = !document.querySelector(".sheet-chart svg")?.textContent.includes("图表编辑探针");
        rb("重做")?.click();
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
        rb("撤销")?.click(); await wait(100);
        const undoSize = Math.abs(parseFloat(document.querySelector(".sheet-chart")?.style.width || "0") - originalWidth) < .01;
        rb("重做")?.click(); await wait(100);
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
        // Cell borders (#237): the ribbon's three presets land on the selection's four sides.
        // Read computed styles, with the sheet's own grid line counting as "no cell border";
        // the run then saves with the box applied so an independent reader can see it on disk.
        tap("D1")?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, button: 0 }));
        await wait(60);
        tap("D3")?.dispatchEvent(new PointerEvent("pointermove", { bubbles: true }));
        await wait(60);
        document.querySelector(".sheet-wrap")?.dispatchEvent(new PointerEvent("pointerup", { bubbles: true }));
        await wait(80);
        const sheetEdge = (ref) => {
          const el = tap(ref);
          if (!el) return "";
          const cs = getComputedStyle(el);
          return ["top", "right", "bottom", "left"].map((s) => {
            const w = cs.getPropertyValue("border-" + s + "-width");
            const c = cs.getPropertyValue("border-" + s + "-color");
            return w === "0px" || c === "rgb(220, 223, 228)" ? "0" : w + " " + cs.getPropertyValue("border-" + s + "-style") + " " + c;
          }).join(" | ");
        };
        rb("所有框")?.click();
        await wait(150);
        const borderOn = sheetEdge("D2");
        rb("无框")?.click();
        await wait(150);
        const borderOff = sheetEdge("D2");
        rb("所有框")?.click();
        await wait(150);
        const sheetBorder = { on: borderOn, off: borderOff, first: sheetEdge("D1"), last: sheetEdge("D3") };
        // 页签条（#238）：表格页按 WPS etongmani.kui 的 KRbTabWidgetCommand 分页，四档各自的命令
        // 要真落在自己那一档里；换档只留一档可见；撤销/重做是字面不是箭头字形（箭头在本机字栈里是豆腐块）。
        const sheetTabs = [...document.querySelectorAll("#ribbon .rtabs button")];
        const barHas = (key, labels) => {
          const bar = document.querySelector("#ribbon .rbar[data-tab='" + key + "']");
          const own = bar ? [...bar.querySelectorAll("button")].map((b) => b.textContent) : [];
          return labels.every((t) => own.includes(t));
        };
        const dataTabBtn = sheetTabs.find((b) => b.textContent === "数据");
        dataTabBtn?.click();
        await wait(60);
        const sheetRibbon = {
          tabs: sheetTabs.map((b) => b.textContent).join("/"),
          home: barHas("home", ["撤销", "重做", "所有框", "色阶", "查找"]),
          insert: barHas("insert", ["图片", "柱状图", "编辑图表"]),
          data: barHas("data", ["筛选", "清除筛选", "升序", "降序"]),
          view: barHas("view", ["冻结", "解冻"]),
          oneBar: [...document.querySelectorAll("#ribbon .rbar")].filter((el) => el.hidden !== true).length === 1,
          ariaOnData: dataTabBtn?.getAttribute("aria-selected") === "true",
          dataBarShown: document.querySelector("#ribbon .rbar[data-tab='data']")?.hidden !== true,
          tofu: [...document.querySelectorAll("#ribbon button")].some((b) => /[↶↷]/.test(b.textContent || "")),
          overflow: !!document.querySelector("#ribbon .rover"),
        };
        sheetTabs[0]?.click();
        await wait(60);
        // 组名不是只写在属性上：量一下 ::after 的 content 与带子高度，证明那行字真被排了出来。
        const homeGrps = [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] .grp")];
        sheetRibbon.caps = homeGrps.map((g) => g.dataset.cap || "-").join("/");
        sheetRibbon.capInk = homeGrps.map((g) => getComputedStyle(g, "::after").content.replace(/"/g, "")).join("/");
        sheetRibbon.capBar = Math.round(document.querySelector("#ribbon .rbar[data-tab='home']").getBoundingClientRect().height);
        // 图标不是只在类名上：数到带子里真排出来的 SVG 路径，且按钮的 textContent 仍是标签那行字（图标里没有文字节点）。
        sheetRibbon.icons = document.querySelectorAll("#ribbon .rbar[data-tab='home'] button.rb.ico .ic svg path").length;
        sheetRibbon.iconText = [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] button.rb.ico")].every((b) => b.textContent === b.querySelector(".lb")?.textContent);
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
        return { docPart, bandLeg, bandTabsLeg, editMarksLeg, sortLeg, rowSplitLeg, rulerLeg, painterLeg, leaderLeg, tabStopsLeg, tablePlaceLeg, tablePropsWriteLeg, tableUnitsLeg, wrapLeg, keepLeg, sizeLeg, clearLeg, styleLeg, vertLeg, caseLeg, origLeg, origTableLeg, underLeg, paraBox, paraBoxCmd, tblBoxCmd, tableToolsCmd: ttCmd, docInsert, docOutline, pptPart, pptInsert, pptShow, pptViews, pptSaved, sheetPart, sheetFill, sheetPic, sheetCF, sheetGT, sheetNote, sheetChart, sheetChartEdit, sheetChartGeometry, sheetZoom, sheetBorder, sheetRibbon, docNewPart, errorState, recovered };
      })()`;
      void last?.webContents
        .executeJavaScript(probe)
        .then((r) => console.log("SOLO_SMOKE", JSON.stringify({ ...r, inPlaceWrites: smokeWrites, saveAsWrites: smokeSaves })))
        .catch((e: unknown) => {
          console.log("SOLO_SMOKE_ERROR", e instanceof Error ? (e.stack ?? e.message) : String(e));
          console.log("SOLO_SMOKE_CONSOLE", JSON.stringify(smokeConsole));
        })
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
