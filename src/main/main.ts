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
        const sizeBox = document.querySelector("#ribbon select.pick");
        const faceRect = faceBox?.getBoundingClientRect();
        const sizeRect = sizeBox?.getBoundingClientRect();
        const ribbonPair = faceRect && sizeRect && Math.abs(faceRect.top - sizeRect.top) < 0.5 && Math.abs(faceRect.height - sizeRect.height) < 0.5 && Math.round(faceRect.width) === 130 ? 1 : 0;
        marker?.click();
        await wait(150);
        const after = document.querySelector("#view .page");
        const docPart = { title: document.getElementById("title")?.textContent, editable: after?.getAttribute("contenteditable"), blocksBefore: blocks, blocksAfter: after?.children.length, typed: after?.textContent?.includes("SMOKE"), bulletsBefore: bullets, bulletsWearingOwnFont: bulletWorn, bulletsOwnSize: bulletSized, bulletsOwnColor: bulletInked, bulletsAfter: after?.querySelectorAll(".bullet").length, ribbonGroups: document.querySelectorAll("#ribbon .grp").length, fontBoxSpans, fontBoxEcho, ribbonPair };

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
        const ringOf = (needle) => [...document.querySelectorAll("#view .page table.doc-table")].find((t) => (t.textContent ?? "").includes(needle));
        const cellOf = (needle) => [...(ringOf(needle)?.querySelectorAll("td") ?? [])].find((c) => (c.textContent ?? "").includes(needle));
        const spanOf = (needle) => cellOf(needle)?.querySelector("span")?.firstChild ?? null;
        const ringBottom = (el) => (el ? cssOf(el, "borderBottomStyle") + "/" + cssOf(el, "borderBottomWidth") : "");
        const boxPick = () => [...document.querySelectorAll("#ribbon select")].find((s) => s.dataset && s.dataset.scope !== undefined);
        const legCount = () => {
          const s = boxPick();
          return s ? [...s.options].length : -1;
        };
        const clickLeg = (value) => {
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
        const tblUndo = ribbonBtn("↶") ? (ribbonBtn("↶").click(), 1) : 0;
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
        const tblRedo = ribbonBtn("↷") ? (ribbonBtn("↷").click(), 1) : 0;
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
          const tfCtlN = tfGroup ? tfGroup.querySelectorAll(".facebox,.pick,.rb,.swatch").length : -1;
          const tfFaceW = tfFaceNow() ? Math.round(tfFaceNow().getBoundingClientRect().width) : -1;
          const tfPrevDist = tfGroupIdx > 0 && [...tfGroups[tfGroupIdx - 1].querySelectorAll("button")].some((b) => b.textContent === "分布行") ? 1 : 0;
          const tfNextAlign = tfGroupIdx >= 0 && tfGroupIdx + 1 < tfGroups.length &&
            [...tfGroups[tfGroupIdx + 1].querySelectorAll("button")].map((b) => b.textContent).join("|") === "↑|↕|↓|左|中|右" ? 1 : 0;
          const tfSel = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='table'] select")].find((s) => s.title === label);
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
          const tfHomeSel = (label) => [...document.querySelectorAll("#ribbon .rbar[data-tab='home'] select")].find((s) => s.title === label);
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
        // #158：每一道缝都得有东西把它切开。条带的道数 == 段落垫片 + 垫片行，且状态栏「共 N 页」报的 N 与
        // 条带数 + 1 同一份。旧写法在缝落进两枚表行之间时既不垫也不画（垫片进表格会改行高），于是纸数照旧往上涨、
        // 画出来的纸却少一道缝——两张纸糊在一起。DOM 门禁钉的是垫片行的形状，这一腿钉的是 shipping 盒子里
        // 「纸数与画出来的缝」不能再各说各话。
        const seamSplit = (() => {
          const spans = document.querySelectorAll("#view .page .solo-sheet-gap").length;
          const rows = document.querySelectorAll("#view .page tr.solo-sheet-gap-row").length;
          const strips = document.querySelectorAll("#view .page .solo-sheet-strip").length;
          const cell = [...document.querySelectorAll(".solo-doc-status-loc")].find((b) => (b.title || "").indexOf("页数") === 0);
          const m = /([0-9]+)/.exec(cell ? cell.textContent || "" : "");
          const n = m ? Number(m[1]) : -1;
          return { spans: spans, rows: rows, strips: strips, statusPages: n, cutEverySeam: strips === spans + rows, pagesMatch: n > 0 && strips + 1 === n };
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
        return { docPart, bandLeg, underLeg, paraBox, paraBoxCmd, tblBoxCmd, tableToolsCmd: ttCmd, docInsert, docOutline, pptPart, pptInsert, pptShow, pptSaved, sheetPart, sheetFill, sheetPic, sheetCF, sheetGT, sheetNote, sheetChart, sheetChartEdit, sheetChartGeometry, sheetZoom, docNewPart, errorState, recovered };
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
