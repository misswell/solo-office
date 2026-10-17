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
        // 以及光标退回一格时列子是不是回到九档。件里那句 w:tblBorders 由 .scratch/tbl-cmd-install-check.sh 在外面问。
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
        // 光标退回一格：那一层说的是格（九档），落点从 table 那一圈换到那一格自己那一圈的顶缝。
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
        return { docPart, paraBox, paraBoxCmd, tblBoxCmd, docInsert, docOutline, pptPart, pptInsert, pptShow, pptSaved, sheetPart, sheetFill, sheetPic, sheetCF, sheetGT, sheetNote, sheetChart, sheetChartEdit, sheetChartGeometry, sheetZoom, docNewPart, errorState, recovered };
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
