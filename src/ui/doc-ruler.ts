/**
 * 水平标尺：#228（WPS 视图·显示组的 `ViewRulerWord`）。
 *
 * WPS 侧实据：`office6/res/commands/wps/ongmani/CT_View.kuip` 的 `RB_Show`（`idMso="GroupViewShowHide"`、
 * `groupTip="ShowGroup"`、`text="@Show"`）里，`ViewRulerWord`（`ksoCmd="Ruler"`、`customTip="Ruler"`、
 * `qtspyName="Ruler"`、`hotKey="l"`、`rbHotKey="R"`、`text="@Ruler"`）排在 `sp_ShowDocumentMap` 与
 * `CommentOrTrackingChangeMarkup` 之间，同组的另有 `TableGridlines` / `TaskPaneInViewTab` / `ViewGridlinesWord`；
 * `wpsentries.kuip` 里有 `CT_View.RB_Show.ViewRulerWord`→`ViewRuler`、`ViewMenu.Ruler`，打印预览另有一枚
 * `CT_PrintPreview.RB_PrintShow.PrintShowViewRuler`；右键菜单声明了 `Ruler Popup Menu` 的
 * `FirstLineIndent` / `HangingIndent` / 三种制表位对齐。所以「视图页有一枚标尺开关」是声明实据；**真机上标尺的
 * 长什么样、数字从哪一格起、主刻度与细分的步长、拖标记的吸附步长都未取证**（WPS 真机卡在首次启动那页，只能
 * owner 点）。这一趟只画「能站得住的那些数」，不碰拖动：标尺上每一枚刻度与标记的位置都来自纸面自己。
 *
 * 三个来源，没有一个新发明的数：
 *
 * - **纸与正文栏**：`renderWord` 写在 `.page` 行内样式上的 `--page-w`（纸宽）、`--pad-l`（左边距 + 装订线）、
 *   `--pad-r`（右边距），单位是纸的像素（`twipsToPx` 的 15 缇/px）。标尺与纸同宽、同一套 `zoom`，
 *   正文栏的左沿就落在 `--pad-l` 上。
 * - **刻度**：测量单位是状态栏那一格（`doc-units.ts` 的 `currentUnit`，#122）。刻度值由单位换像素，
 *   数字档位只按「相邻数字至少排得下」挑 1/2/5×10^k——这是版面规则，不是 WPS 的读数。
 * - **缩进与制表位标记**：缩进不问模型、也不重跑一遍 `indResolve`——直接量纸面那个 `.para` 的盒子
 *   （`getBoundingClientRect` 除回 `zoom`，再减纸盒左沿）。这样「纸怎么摆、标尺就怎么标」，不会像
 *   `w:jc` 那样出现第二份解释。制表位读的正是 `doc-layout.ts` 画制表符用的那枚 `data-tab-stops`
 *   （`parseTabStops` 同一份解析），位置从**正文栏左沿**起算（§17.3.2.9，与 `applyTabStops` 同一条）。
 *
 * 表格里的段不画标记：格子里的缩进基准是格沿，不是正文栏左沿（同一个理由写在 `applyTabStops` 里挡 `w:tc`
 * 的那一段）。带子里（`.hf-p`）同理——它不在 `.page` 的正文流里。
 */

import { pxToTwips, twipsToPx } from "../core/docx.js";
import { TAB_STOPS_ATTR, parseTabStops } from "./doc-layout.js";
import { currentUnit, pxPerUnit, unitDef, type Unit } from "./doc-units.js";

/** 一枚刻度：`x` 是离**正文栏左沿**的纸面像素，`label` 只有主刻度才有。 */
export interface RulerTick {
  x: number;
  major: boolean;
  label?: string;
}

/** 数字之间至少要让出这么宽（纸面像素）。9px 的字，「10」连间距大约 20px，留到 30 就绝不会叠。 */
const LABEL_GAP_PX = 30;

/** 主刻度步长按 1 → 2 → 5 → 10 → … 往上走，直到相邻两个数字排得下。 */
const nextStep = (s: number): number => {
  const decade = Math.pow(10, Math.floor(Math.log10(s) + 1e-9));
  const mantissa = Math.round(s / decade);
  return (mantissa === 1 ? 2 : mantissa === 2 ? 5 : 10) * decade;
};

/**
 * 正文栏那一横条上该画哪些刻度。`0` 在左沿（Word/WPS 的惯例：正文栏左沿是 0，页边距摆在阴影区里；
 * 真机未取证），向纸的右沿数；一步单位折不成整数像素时按纸面像素截断，不让最后一枚越过右沿。
 */
export function rulerTicks(textWidthPx: number, u: Unit = currentUnit()): RulerTick[] {
  const unitPx = pxPerUnit(u);
  if (!(textWidthPx > 0) || !(unitPx > 0)) return [];
  let step = 1;
  while (step * unitPx < LABEL_GAP_PX) step = nextStep(step);
  // 细分：四等分（Word 的厘米尺是 0.25 cm 一枚）；四分之一挤不下就二等分；还挤不下就不画细分。
  const quarter = (step * unitPx) / 4;
  const minor = quarter >= 4 ? step / 4 : step * unitPx / 2 >= 4 ? step / 2 : step;
  const out: RulerTick[] = [];
  const maxUnits = textWidthPx / unitPx;
  // 步长是 1/2/5 的十进数，最多两位小数（0.25），用整数步进避开浮点累加。
  const stepUnits = Math.round(minor * 100);
  for (let k = 0; ; k += 1) {
    const value = (k * stepUnits) / 100;
    if (value > maxUnits + 1e-9) break;
    const ratio = value / step;
    const major = Math.abs(ratio - Math.round(ratio)) < 1e-9;
    const x = value * unitPx;
    if (x > textWidthPx + 0.001) break;
    out.push(major ? { x, major, label: String(Math.round(value * 100) / 100) } : { x, major });
  }
  return out;
}

/** 标尺上那三枚缩进标记与制表位标记，`x` 都是离纸盒左沿的纸面像素。 */
export interface RulerMarks {
  left?: number;
  first?: number;
  right?: number;
  /** 段自己画在盒子里的左右内边距（列表缩进与段落边框的留白都在这一条里），纸面像素。 */
  inlineLeft?: number;
  inlineRight?: number;
  tabs: { x: number; val: "left" | "center" | "right" }[];
}

const pxOf = (v: string): number => {
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : 0;
};

/**
 * 光标所在段在纸上摆出来的三条边（首行、左、右）与它自己说的制表位。
 * 量的是**盒子**：`getBoundingClientRect()` 给视口像素，除回 `zoom` 就是纸的像素，再减纸盒左沿——
 * 与 `doc-layout.ts` 量页沿同一条口径。段在表格里（`td`/`th`）或带子里（不在 `.page` 里）时不给标记：
 * 那两处缩进的基准不是正文栏左沿。
 */
export function rulerMarksOf(page: HTMLElement, para: HTMLElement | null, zoom: number): RulerMarks {
  const marks: RulerMarks = { tabs: [] };
  if (!para || !page.contains(para) || para.closest("td, th")) return marks;
  const scale = zoom > 0 ? zoom : 1;
  const rect = page.getBoundingClientRect();
  const box = para.getBoundingClientRect();
  const cs = getComputedStyle(para);
  const padLeft = pxOf(cs.paddingLeft);
  const padRight = pxOf(cs.paddingRight);
  const left = (box.left - rect.left) / scale + padLeft;
  marks.left = left;
  marks.first = left + pxOf(cs.textIndent);
  marks.right = (box.right - rect.left) / scale - padRight;
  marks.inlineLeft = padLeft;
  marks.inlineRight = padRight;
  // 制表位的位置与 `applyTabStops` 一样从正文栏左沿起算；这里只画，不问几何，所以直接用同一枚 `twipsToPx`。
  const column = pxOf(page.style.getPropertyValue("--pad-l"));
  for (const stop of parseTabStops(para.getAttribute(TAB_STOPS_ATTR) ?? "")) {
    marks.tabs.push({ x: column + twipsToPx(stop.posTwips), val: stop.val });
  }
  return marks;
}

/**
 * 标尺上能拖的三枚标记。制表位（`.ruler-tab`）不在这条路上：WPS 那边它拖到哪一档、右键菜单怎么落，
 * 都还没有实据。
 */
export type RulerDragKind = "first" | "left" | "right";

/** 拖标记要的几何，全是纸面像素（`zoom` 之前那一层）：纸宽、左右页边距、段自己的左右内边距，三条边各在哪。 */
export interface RulerDragGeometry {
  pageW: number;
  padL: number;
  padR: number;
  inlineLeft: number;
  inlineRight: number;
  leftX: number;
  firstX: number;
  rightX: number;
}

/**
 * 纸面上的一个 x 折回件里那句话说的是谁，单位缇。三条都与 `rulerMarksOf` 量的盒子同一条口径：
 * 左缩进减掉页边距与段自己的 padding（列表缩进画在 padding 里），右缩进从正文栏右沿往回数，
 * 首行给的是**相对左标记的偏移**——负数就是悬挂。吸附步长不在这里发明（WPS 真机未取证），
 * 拖到哪个像素就落哪个 15 缇。
 */
export function indentTwipsOf(kind: RulerDragKind, x: number, geom: RulerDragGeometry): number {
  if (kind === "left") return Math.round(pxToTwips(x - geom.padL - geom.inlineLeft));
  if (kind === "first") return Math.round(pxToTwips(x - geom.leftX));
  return Math.round(pxToTwips(geom.pageW - geom.padR - x - geom.inlineRight));
}

/** 拖动要问编辑器的三件事：能不能拖、怎么落笔、落不了时把标记摆回去。 */
export interface RulerDragHooks {
  /** 修订视图「原始状态」下纸是只读的，尺上的标记也不该动。 */
  enabled(): boolean;
  /** 落进件里；没改动（或改不了）返回 false，由 `repaint()` 把标记摆回原处。 */
  commit(kind: RulerDragKind, twips: number): boolean;
  repaint(): void;
}

const dragHooks = new WeakMap<HTMLElement, RulerDragHooks>();

const geomOf = (ruler: HTMLElement): RulerDragGeometry | null => {
  const raw = ruler.dataset.dragGeom;
  if (!raw) return null;
  try {
    const g = JSON.parse(raw) as RulerDragGeometry;
    return Number.isFinite(g.pageW) && g.pageW > 0 && Number.isFinite(g.leftX) ? g : null;
  } catch {
    return null;
  }
};

const zoomOf = (ruler: HTMLElement): number => {
  const z = parseFloat(ruler.style.zoom);
  return Number.isFinite(z) && z > 0 ? z : 1;
};

const place = (ruler: HTMLElement, kind: RulerDragKind, x: number): void => {
  const el = ruler.querySelector<HTMLElement>(`[data-ruler-mark="${kind}"]`);
  if (!el) return;
  el.style.left = `${x.toFixed(2)}px`;
  el.dataset.x = String(Math.round(x * 10) / 10);
};

/** 左标记不许压到右标记上，右标记也不许压到左标记上；三枚都留在纸里。 */
const clampX = (kind: RulerDragKind, x: number, geom: RulerDragGeometry): number => {
  let lo = 0;
  let hi = geom.pageW;
  if (kind === "left") hi = Math.min(hi, geom.rightX);
  if (kind === "right") lo = Math.max(lo, geom.leftX);
  return Math.min(hi, Math.max(lo, x));
};

/**
 * 给一枚标尺挂上拖动：按下 `.ruler-ind` 就跟指针走，松手才落进件里。监听器挂在标尺自己身上
 * （指针捕获把 move/up 都送回它），纸重铺换掉标尺时一起消失，不会往 `window` 上攒。
 *
 * 拖左标记时首行标记跟着走（件里首行存的是相对左标记的偏移：`w:firstLine` 或 `w:hanging`），
 * 拖首行越过左标记就换档：往右是 `w:firstLine`，往左是 `w:hanging`。
 */
export function bindRulerDrag(ruler: HTMLElement, hooks: RulerDragHooks): void {
  dragHooks.set(ruler, hooks);
  if (ruler.dataset.dragBound === "1") return;
  ruler.dataset.dragBound = "1";
  let state: {
    kind: RulerDragKind;
    geom: RulerDragGeometry;
    zoom: number;
    pointerId: number;
    startClientX: number;
    baseX: number;
    x: number;
    moved: boolean;
  } | null = null;

  const finish = (commit: boolean): void => {
    if (!state) return;
    const st = state;
    state = null;
    const live = dragHooks.get(ruler);
    if (!live) return;
    if (commit && st.moved && live.enabled()) {
      if (live.commit(st.kind, indentTwipsOf(st.kind, st.x, st.geom))) return;
    }
    live.repaint();
  };

  const onDown = (e: PointerEvent): void => {
    const live = dragHooks.get(ruler);
    if (!live?.enabled() || state) return;
    const target = e.target instanceof Element ? e.target.closest<HTMLElement>("[data-ruler-mark]") : null;
    if (!target || !ruler.contains(target)) return;
    const kind = target.dataset.rulerMark as RulerDragKind | undefined;
    const geom = kind ? geomOf(ruler) : null;
    if (!kind || !geom) return;
    e.preventDefault();
    const baseX = kind === "left" ? geom.leftX : kind === "first" ? geom.firstX : geom.rightX;
    state = { kind, geom, zoom: zoomOf(ruler), pointerId: e.pointerId, startClientX: e.clientX, baseX, x: baseX, moved: false };
    try {
      ruler.setPointerCapture(e.pointerId);
    } catch {
      // 合成事件（探针）没有可捕获的指针：不捕获也照拖，监听器本来就挂在尺子上。
    }
  };

  const onMove = (e: PointerEvent): void => {
    if (!state || e.pointerId !== state.pointerId) return;
    const dx = e.clientX - state.startClientX;
    if (!state.moved && Math.abs(dx) < 0.5) return;
    state.moved = true;
    const x = clampX(state.kind, state.baseX + dx / state.zoom, state.geom);
    state.x = x;
    place(ruler, state.kind, x);
    if (state.kind === "left") place(ruler, "first", state.geom.firstX + (x - state.baseX));
  };

  const onUp = (e: PointerEvent): void => {
    if (!state || e.pointerId !== state.pointerId) return;
    try {
      ruler.releasePointerCapture(e.pointerId);
    } catch {
      // 捕获已经没了：照常收尾。
    }
    finish(e.type === "pointerup");
  };

  ruler.addEventListener("pointerdown", onDown);
  ruler.addEventListener("pointermove", onMove);
  ruler.addEventListener("pointerup", onUp);
  ruler.addEventListener("pointercancel", onUp);
}

/** 标尺这一趟的入参：单位、缩放，以及光标那一段（没有就给空段）。 */
export interface RulerPaint {
  unit: Unit;
  zoom: number;
  caret: HTMLElement | null;
}

/**
 * 把标尺画进它自己的元素（`ensureRuler` 造的那一枚）。每次整体重铺——刻度最多几十枚、标记三枚，
 * 比「记住上一趟、挑着改」省事，也不会留着上一段那枚错位的标记。
 */
export function drawRuler(ruler: HTMLElement, page: HTMLElement, paint: RulerPaint): void {
  const width = pxOf(page.style.getPropertyValue("--page-w"));
  const padL = pxOf(page.style.getPropertyValue("--pad-l"));
  const padR = pxOf(page.style.getPropertyValue("--pad-r"));
  const textWidth = Math.max(0, width - padL - padR);
  ruler.style.width = `${width}px`;
  ruler.style.zoom = String(paint.zoom > 0 ? paint.zoom : 1);
  ruler.dataset.rulerUnit = paint.unit;
  ruler.dataset.rulerPage = String(Math.round(width * 10) / 10);
  ruler.dataset.rulerText = String(Math.round(textWidth * 10) / 10);
  const nodes: HTMLElement[] = [];
  // 正文栏那一块亮一些，两侧的页边距留在底色的阴影里——标尺上「字从哪一列起」一眼就看得出来。
  const area = document.createElement("div");
  area.className = "ruler-area";
  area.style.left = `${padL}px`;
  area.style.width = `${textWidth}px`;
  nodes.push(area);
  for (const tick of rulerTicks(textWidth, paint.unit)) {
    const el = document.createElement("i");
    el.className = `ruler-tick${tick.major ? " major" : ""}`;
    el.style.left = `${(padL + tick.x).toFixed(2)}px`;
    el.dataset.tickX = String(Math.round(tick.x * 10) / 10);
    if (tick.label !== undefined) {
      const b = document.createElement("b");
      b.textContent = tick.label;
      el.append(b);
    }
    nodes.push(el);
  }
  const marks = rulerMarksOf(page, paint.caret, paint.zoom);
  if (marks.left !== undefined && marks.first !== undefined && marks.right !== undefined) {
    const geom: RulerDragGeometry = {
      pageW: width,
      padL,
      padR,
      inlineLeft: marks.inlineLeft ?? 0,
      inlineRight: marks.inlineRight ?? 0,
      leftX: marks.left,
      firstX: marks.first,
      rightX: marks.right,
    };
    ruler.dataset.dragGeom = JSON.stringify(geom);
  } else {
    delete ruler.dataset.dragGeom;
  }
  for (const which of ["first", "left", "right"] as const) {
    const x = marks[which];
    if (x === undefined) continue;
    const el = document.createElement("i");
    el.className = `ruler-ind ${which}`;
    el.style.left = `${x.toFixed(2)}px`;
    el.dataset.rulerMark = which;
    el.dataset.x = String(Math.round(x * 10) / 10);
    nodes.push(el);
  }
  for (const tab of marks.tabs) {
    const el = document.createElement("i");
    el.className = `ruler-tab val-${tab.val === "left" ? "l" : tab.val === "right" ? "r" : "c"}`;
    el.style.left = `${tab.x.toFixed(2)}px`;
    el.dataset.rulerTab = tab.val;
    el.dataset.x = String(Math.round(tab.x * 10) / 10);
    nodes.push(el);
  }
  const unit = document.createElement("span");
  unit.className = "ruler-unit";
  unit.textContent = unitDef(paint.unit).name;
  nodes.push(unit);
  ruler.replaceChildren(...nodes);
}

/** 标尺元素挂在纸盒前面（同一个滚动容器里），`render()` 整块重铺 innerHTML 之后由 `ensureRuler` 再挂回去。 */
export function ensureRuler(host: HTMLElement, page: HTMLElement): HTMLElement {
  const found = host.querySelector<HTMLElement>(":scope > .doc-ruler");
  if (found) return found;
  const el = document.createElement("div");
  el.className = "doc-ruler";
  el.contentEditable = "false";
  el.setAttribute("aria-hidden", "true");
  page.before(el);
  return el;
}
