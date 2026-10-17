import type { TableBorder } from "../core/docx.js";

/**
 * 「边框与底纹」对话框：段落那一圈 `w:pBdr` 与那一层 `w:shd` 从此有了一条**能说话**的路。
 * #144 做到「读得出、画得出、保存不丢、铺排算它」，但命令面是空的——一句「这一段要个盒子」只能靠手改文件。
 *
 * 摆位与形状都由 WPS 自己声明：`res/wpscommon.kuip` 里没有这一族，命令在懒加载的
 * `res/commands/wps/ongmani/CT_Home.kuip`：
 * - `:271` `<KxPropComposeCommand id="BordersAndShading" idMso="BordersShadingDialogWord" ksoCmd="BorderShading"
 *   dialogHint="1" hotKey="o" text="@Borders and Shading"/>`——`dialogHint="1"` 说的就是「这一条开一个对话框」；
 * - `:272-287` 那枚下拉 `RbTableOutsideBorder`（idMso="TableBordersMenu"，`referenceCommand="OutsideBorders"`
 *   即按钮本体那一击）按 `<separator/>` 分四段：`Bottom/Top/Left/Right`｜`NoBorder/AllBorders/OutsideBorders/
 *   CrossBorder`｜`HMiddleBorder/VMiddleBorder`｜`BordersAndShading`；
 * - `res/wpsentries.kuip:669` 把它挂在 `CT_Home.RB_Paragraph.*.RbTableOutsideBorder.BordersAndShading`，
 *   也就是**主页 ribbon 的段落组**里，与「行距」那一格同组。
 * `CrossBorder`/`VMiddleBorder` 那两条说的是 `w:insideH`/`w:insideV`——`CT_PBdr`（§17.3.2.36）没有这两个槽
 * （它另有 `w:between`/`w:bar`，授权口径 46 份里 0 枚，`.scratch/corpus-recheck.mjs` 数的就是 `w:pBdr` 那一对标签里的孩子），所以 `HMiddleBorder` 落在段落上是 `w:between`，而那两条在段落那张
 * 列子里没有对应的一句话，不摆。#146 起，同一个口子按选区派发：整表选中落 `w:tblBorders`，格选区落 `w:tcBorders`，
 * 内部横竖那三条永远落表（见文末 `tableLegPatch`）。这十条腿在 WPS 是**一枚控件**按上下文派发，不是两张 ribbon。
 *
 * **一格一格「不改」**：与「段落」「页面设置」同一条纪律——留空/选着「不改」的那一格一个字节都不写，因为
 * 「没说」与「说了 0」是两件事，而 `w:pBdr` 是逐边解决的（§17.7.2），动哪一边就只有那一边归这一段。四格
 * （线型/磅/间距/颜色）都在**每一条边自己那一行**里：件里的 `w:sz`/`w:space`/`w:color` 本来就是逐边写的
 * （26 枚 `w:pBdr` 的 74 枚边孩子人人都带），没有「共享的那三格」这种形状。一条边原本没有线（自己与样式链
 * 都没说）而只填了磅/间距/颜色，这一击作废并说明——线型是 `CT_Border` 里说不出来的那一枚，不能替这一段编。
 *
 * 默认形状照件，不猜：`.scratch/pbdr-audience.ts`（46 份授权件）里 26 枚 `w:pBdr`（24 枚在 `word/styles.xml`、
 * 2 枚在页眉部件、正文段落自己写的 0 枚），侧数 bottom 26 / top、left、right 各 16；属性形状只有四种组合
 * ——`sz=0 space=1`（32 枚）、`sz=0 space=4`（32 枚）、`sz=6 space=1`（8 枚，全是要画线的那 8 枚
 * `w:val="single"`）、`sz=0 space=0`（2 枚），`w:color` 74/74 写 `auto`，`w:shadow` 0 枚。所以 ribbon 那几条腿
 * 落下「单线」时交的是 `val="single" sz=6 space=1 color="auto"`（件里**画过线**的边只证过 1 磅这一档间距，左右
 * 两条画线的边一枚样本都没有，所以不给它们编 4 磅），落下「无框」时四边一起交 `val="none" sz=0`
 * 加那一边自己的间距档（顶/底 1、左/右 4——那 16 枚四边全说的样式就是这个形状）；
 * 底纹 30/30 枚写 `w:val="clear" w:color="auto"` 加一个十六进制 `w:fill`，「无地」这一格交的是 `w:fill="auto"`
 * ——那是**一句**「这一段要页面自己的地」，与「没写底纹」不是一件事（段落层 0 例、`w:tcPr` 层 1,600 例）。
 * 各格中文措辞（「边框与底纹」「单线」「双线」「间距」…）是这一族软件的通用说法，本机 `.qm` 读不出，**未取证**。
 */
const STYLE_ID = "solo-para-border-styles";

/** 对话框里那条下拉摆得出的线型：`""` 是「不改」，`none` 是一句「这一边不画线」。 */
const VALS: [string, string][] = [
  ["", "不改"],
  ["none", "无"],
  ["single", "单线"],
  ["double", "双线"],
  ["dashed", "虚线"],
  ["dotted", "点线"],
];

/**
 * 对话框摆得出的那四条边。`CT_PBdr` 另有 `w:between`（这一段与下一段之间那条）与 `w:bar`（贯穿左边距的那一条），
 * 这一族的对话框没有它们的行——`w:bar` 没有落笔处，`w:between` 只有 ribbon 那条「内部横框线」说得出来。
 */
export type ParaBoxSide = "top" | "left" | "bottom" | "right";
/** 段落这一层**落笔**说得出的边：对话框那四条，加 ribbon 那两条腿要的 `w:between`。 */
export type ParaEdge = ParaBoxSide | "between";
const SIDES: [ParaBoxSide, string][] = [
  ["top", "上"],
  ["left", "左"],
  ["bottom", "下"],
  ["right", "右"],
];
/** 对话框摆得出的那四条边，顺序与对话框一致；回显那本账按它逐边取。 */
export const PARA_BOX_SIDES: readonly ParaBoxSide[] = SIDES.map(([side]) => side);

/** 一条边被说到的那几句；缺的键 = 这一句不改，沿用这一段这一边原来的说法。 */
export type EdgePatch = {
  /** `w:val`，ST_Border 的一枚字面。 */
  val?: string;
  /** `w:sz`，八分之一磅（对话框按磅收，换算只在这里做一次）。 */
  sizeEighths?: number;
  /** `w:space`，磅：正文到这条线的距离。 */
  spacePt?: number;
  /** `w:color`：`auto` 或六位十六进制，按件的说法（不带 `#`）。 */
  color?: string;
};

/**
 * 一次「确定」交回来的东西：`sides` 一条也没有且 `shading` 也没有 = 这一击什么都没问。
 * 边名按选区那一层给（段落四条、表六条、格四条），补丁的形状是同一份。
 */
export type BoxPatch<S extends string> = {
  sides?: Partial<Record<S, EdgePatch>>;
  /** 底纹 `w:fill`：`auto`（无地）或 `#RRGGBB`。键不在 = 不改。 */
  shading?: string;
};
export type ParaBoxPatch = BoxPatch<ParaEdge>;

/** 打开时摆出来的读数（段落那一层）：第一段那一套框（含样式链借来的那一份），加一颗「所选几段说法不一致」的提示。 */
export type ParaBoxState = BoxState<ParaBoxSide>;

const CSS = `
.solo-pb-dialog {
  width: 380px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-pb-dialog form { display: grid; gap: 12px; padding: 14px 16px 16px; }
.solo-pb-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-pb-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-pb-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-pb-sides { display: grid; gap: 8px; grid-template-columns: 3em repeat(4, minmax(0, 1fr)); align-items: center; }
.solo-pb-sides b { font-weight: 500; color: #4b5563; font-size: 12px; }
.solo-pb-sides label { display: grid; gap: 2px; font-size: 11px; color: #6b7280; }
/* 「应用于」那一行说的是这一击落在哪一层：同一枚对话框在段落、整张表、所选格子三处摆的边不一样。 */
.solo-pb-scope { margin: 0; font-size: 12px; color: #6b7280; }
.solo-pb-scope b { color: #2b2b2b; font-weight: 600; }
.solo-pb-dialog input, .solo-pb-dialog select {
  min-width: 0; padding: 4px 6px; font: 13px/1.4 inherit; color: inherit;
  border: 1px solid #c9d3e6; border-radius: 6px; background: #fff;
}
.solo-pb-ground label, .solo-pb-ground { display: grid; gap: 3px; font-size: 12px; color: #4b5563; }
.solo-pb-hint { color: #6b7280; font-size: 11px; min-height: 14px; }
.solo-pb-error { margin: 0; min-height: 18px; color: #b42318; font-size: 12px; }
.solo-pb-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-pb-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-pb-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
.solo-pb-apply[disabled] { opacity: .5; cursor: default; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

/** 颜色的位置：`auto`（或别的 token）按件的说法原样走，十六进制归成六位大写不带 `#`。 */
function colorToken(text: string): string | undefined {
  const t = text.trim();
  if (!t) return undefined;
  if (/^#?[0-9A-Fa-f]{6}$/.test(t)) return t.replace(/^#/, "").toUpperCase();
  return /^[A-Za-z][-\w]*$/.test(t) ? t : undefined;
}

/** 底纹那一格同一条口径，但模型里颜色的形状带 `#`（读端 `hex()` 就是这么给的）。 */
function fillValue(text: string): string | undefined {
  const t = text.trim();
  if (!t) return undefined;
  if (/^#?[0-9A-Fa-f]{6}$/.test(t)) return "#" + t.replace(/^#/, "").toUpperCase();
  return /^[A-Za-z][-\w]*$/.test(t) ? t : undefined;
}

/** 磅 → `w:sz` 的八分之一磅；留空给 `undefined`（不改），负数与非数由调用方判坏。 */
function readNumber(input: HTMLInputElement): { value?: number; bad: boolean } {
  const t = input.value.trim();
  if (!t) return { bad: false };
  const v = Number(t);
  return Number.isFinite(v) && v >= 0 ? { value: v, bad: false } : { bad: true };
}

/** 打开时摆出来的读数：第一处那一套框，加一颗「多处说法不一致」的提示。 */
export type BoxState<S extends string> = {
  sides: Partial<Record<S, TableBorder>>;
  shading?: string;
  /** 几处摆不出同一个答案：那一行小字说明，格子摆的是第一处的话，填了的格子写给全部。 */
  mixed?: boolean;
};

export type BoxDialogOpts<S extends string> = {
  /** 这一击落在哪一层。对话框把这句写在「应用于」那一格里：点的到底是段落、整张表还是所选格子，读者得看得见。 */
  appliesTo: string;
  /** 那一层说得出的边；顺序就是对话框里从上到下的顺序，回显那本账也按它逐边取。 */
  sides: readonly [S, string][];
  /**
   * 那一层有没有底纹的槽**写得回去**。三层现在都有：段落的 `w:pPr/w:shd`（#145）、表与格的
   * `w:tblPr/w:shd`、`w:tcPr/w:shd`（#152，同一个 `shadingSlotXml` 落进各自的槽序）。
   */
  shading?: boolean;
  state: BoxState<S>;
  apply: (patch: BoxPatch<S>) => void;
};

export function showBoxDialog<S extends string>(opts: BoxDialogOpts<S>): void {
  const { state, appliesTo, apply } = opts;
  document.querySelectorAll<HTMLDialogElement>(".solo-pb-dialog").forEach((old) => {
    old.close();
    old.remove();
  });
  injectStyles();
  const dialog = document.createElement("dialog");
  dialog.className = "solo-pb-dialog";
  dialog.setAttribute("aria-label", "边框与底纹");
  const sideRow = ([side, label]: [S, string]): string =>
    `<b>${label}</b><label>线型<select name="pb-${side}">${VALS.map(([v, t]) => `<option value="${v}">${t}</option>`).join("")}</select></label>
     <label>磅<input name="pb-${side}-sz" inputmode="decimal" autocomplete="off" value="" placeholder="不改"></label>
     <label>间距<input name="pb-${side}-sp" inputmode="decimal" autocomplete="off" value="" placeholder="不改"></label>
     <label>颜色<input name="pb-${side}-co" autocomplete="off" value="" placeholder="auto"></label>`;
  dialog.innerHTML =
    `<form><header><h2>边框与底纹</h2><button type="button" class="solo-pb-close" aria-label="关闭">×</button></header>
    <p class="solo-pb-scope">应用于：<b>${appliesTo}</b></p>
    <div class="solo-pb-sides">${opts.sides.map(sideRow).join("")}</div>
    ${opts.shading ? `<label class="solo-pb-ground">底纹<input name="pb-shading" autocomplete="off" value="${state.shading ?? ""}" placeholder="留空不改；auto 或十六进制"><small class="solo-pb-hint" data-hint="pb-shading"></small></label>` : `<small class="solo-pb-hint" data-hint="pb-shading"></small>`}
    <p class="solo-pb-error" role="status"></p>
    <footer><button type="button" class="solo-pb-cancel">取消</button><button type="submit" class="solo-pb-apply">确定</button></footer></form>`;

  const box = (name: string): HTMLInputElement => dialog.querySelector<HTMLInputElement>(`[name="${name}"]`)!;
  const select = (name: string): HTMLSelectElement => dialog.querySelector<HTMLSelectElement>(`[name="${name}"]`)!;
  const error = dialog.querySelector<HTMLElement>(".solo-pb-error")!;
  const applyButton = dialog.querySelector<HTMLButtonElement>(".solo-pb-apply")!;
  const hint = dialog.querySelector<HTMLElement>('[data-hint="pb-shading"]')!;
  const shading = opts.shading ? box("pb-shading") : null;

  // 每一边把件（或样式链）正在画的那几句摆进格子：线型认得出就选中，认不出来（`thick`、`hairline`…）就把它
  // 插成额外的一档，而不是让格子空着——留空在这里是「不改」，把件的说法当成用户新点的一句话不算话。
  // 摆出来只为「原样确定」不漂移：内容没变的边由逐边比掉，一个字节也不会多写。
  for (const [side] of opts.sides) {
    const edge = state.sides[side];
    if (!edge) continue;
    const pick = select(`pb-${side}`);
    if (!VALS.some(([v]) => v === edge.val)) pick.add(new Option(edge.val, edge.val));
    pick.value = edge.val;
    if (edge.size !== undefined) box(`pb-${side}-sz`).value = String(edge.size / 8);
    if (edge.space !== undefined) box(`pb-${side}-sp`).value = String(edge.space);
    if (edge.color) box(`pb-${side}-co`).value = edge.color;
  }
  hint.textContent = state.mixed ? "所选几处说法不一致：不一致的那几行留空（留空=不改），原样确定不改内容；填了的格子写给全部。" : "";

  const build = (): BoxPatch<S> | undefined => {
    const patch: BoxPatch<S> = {};
    const sides: Partial<Record<S, EdgePatch>> = {};
    for (const [side, label] of opts.sides) {
      const val = select(`pb-${side}`).value;
      const sz = readNumber(box(`pb-${side}-sz`));
      const sp = readNumber(box(`pb-${side}-sp`));
      const color = colorToken(box(`pb-${side}-co`).value);
      if (sz.bad || sp.bad) {
        error.textContent = "磅与间距要填 0 或更大的数字，留空表示不改。";
        applyButton.disabled = true;
        return undefined;
      }
      if (box(`pb-${side}-co`).value.trim() && color === undefined) {
        error.textContent = "颜色要填 auto 或六位十六进制（如 FFD966），留空表示不改。";
        applyButton.disabled = true;
        return undefined;
      }
      const edge: EdgePatch = {};
      if (val) edge.val = val;
      if (sz.value !== undefined) edge.sizeEighths = Math.round(sz.value * 8);
      if (sp.value !== undefined) edge.spacePt = sp.value;
      if (color !== undefined) edge.color = color;
      if (!Object.keys(edge).length) continue;
      // 「这一边原本没有线」时磅/间距/颜色挂不住：没有边可改，写出去就是替这一段新编一条线（线型是 `CT_Border`
      // 里唯一说不出来的那一枚，§17.3.2.36 的每一边都由 `w:val` 带头）。原本有线（自己的或样式链借的）就只改数。
      if (!val && !state.sides[side]) {
        error.textContent = `「${label}」边原本没有线，要先点明线型才改得了磅、间距与颜色。`;
        applyButton.disabled = true;
        return undefined;
      }
      sides[side] = edge;
    }
    const fill = shading ? fillValue(shading.value) : undefined;
    if (shading?.value.trim() && fill === undefined) {
      error.textContent = "底纹要填 auto（无地）或六位十六进制（如 FFF2CC），留空表示不改。";
      applyButton.disabled = true;
      return undefined;
    }
    if (Object.keys(sides).length) patch.sides = sides;
    if (fill !== undefined) patch.shading = fill;
    error.textContent = "";
    applyButton.disabled = false;
    return patch;
  };
  dialog.addEventListener("input", () => build());
  dialog.addEventListener("change", () => build());
  dialog.addEventListener("keydown", (e) => e.stopPropagation());

  const sel = document.getSelection();
  const saved = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
  const restore = (): void => {
    if (!saved || !sel) return;
    sel.removeAllRanges();
    sel.addRange(saved);
  };
  const shut = (): void => {
    dialog.close();
    dialog.remove();
    restore();
  };
  dialog.addEventListener("close", () => {
    if (!dialog.isConnected) return;
    dialog.remove();
    restore();
  });
  dialog.querySelector<HTMLButtonElement>(".solo-pb-close")!.onclick = shut;
  dialog.querySelector<HTMLButtonElement>(".solo-pb-cancel")!.onclick = shut;
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    const patch = build();
    if (!patch) return;
    shut();
    apply(patch);
  };
  document.body.append(dialog);
  build();
  dialog.showModal();
  dialog.querySelector<HTMLSelectElement>(`[name="pb-${opts.sides[0]?.[0]}"]`)?.focus();
}

/** 段落那一层的入口：四条边加底纹，说的就是这一段（几段一起选时按第一段的说法摆，不一致由那行小字说）。 */
export function showParaBorderDialog(state: ParaBoxState, apply: (patch: ParaBoxPatch) => void): void {
  showBoxDialog({ appliesTo: "段落", sides: SIDES, shading: true, state, apply });
}

/**
 * 表那一层：六条边（比段落多出内部横竖两条，`CT_TblBorders` 才有的那两枚孩子）加这张表自己的地。
 * 底纹这一行自 #152 起接得上：`w:tblPr/w:shd` 读得出、`syncTableBox` 写得回同一枚槽。
 */
export function showTableBoxDialog(state: BoxState<TableBoxSide>, apply: (patch: BoxPatch<TableBoxSide>) => void): void {
  showBoxDialog({ appliesTo: "整张表", sides: TABLE_EDGES, shading: true, state, apply });
}

/** 格那一层：每一格自己那四条边与它自己的地；确定之后逐格落，本来就这么说的格一个字节不动。 */
export function showCellBoxDialog(state: BoxState<CellBoxSide>, apply: (patch: BoxPatch<CellBoxSide>) => void): void {
  showBoxDialog({ appliesTo: "所选单元格", sides: CELL_EDGES, shading: true, state, apply });
}

/** 这一边「说了不要线」时的形状：`w:sz="0"` 加一件里最常见的那档间距（32+32 枚 `space=1`/`space=4`）。 */
export const NONE_EDGE = (space = 1): TableBorder => ({ val: "none", size: 0, space, color: "auto" });
/** 件自己画一条线时的形状：8/8 枚 `w:val="single"` 全是 `sz=6 space=1 color=auto`（`.scratch/pbdr-audience.ts`）。 */
export const SINGLE_EDGE: TableBorder = { val: "single", size: 6, space: 1, color: "auto" };
/**
 * 表与格那一层「新说一条边」的形状，与段落那一族**不是**同一套数——那两个数是量出来的
 * （`.scratch/tblbdr-census.ts`，授权口径 46 份件、58 张表、4,136 格）：表级 336 枚边孩子**全带 `w:space="0"`**
 * 且 `w:color="auto"`，画线那 332 枚按 `w:sz` 分 4（260 枚）与 12（72 枚），**一枚 `sz="6"` 也没有**；
 * 格级 1,924 枚同样 `space=0`（1,912 枚说了，12 枚是 `w:val="nil"` 没写），线型 single 里 `sz=4` 是多数。
 * 段落那一族的 `sz=6 space=1` 在这两层没有观众，搬过来就是替文件编一个它从来没写过的形状。
 */
export const TABLE_SINGLE_EDGE: TableBorder = { val: "single", size: 4, space: 0, color: "auto" };
export const TABLE_NONE_EDGE: TableBorder = { val: "none", size: 0, space: 0, color: "auto" };
/** 一条边从没被这一层说过时，按哪一层的形状补——段落与表/格各一档。 */
export type EdgeShape = { single: TableBorder; none: TableBorder };
export const PARA_EDGE_SHAPE: EdgeShape = { single: SINGLE_EDGE, none: NONE_EDGE() };
export const TABLE_EDGE_SHAPE: EdgeShape = { single: TABLE_SINGLE_EDGE, none: TABLE_NONE_EDGE };

/**
 * ribbon 那几条腿：`none` = 这一层说得出的边一起说不要，其余按件的形状给那一条边画线。
 * 「无框」那一句的形状也照件，不是随手抹掉：那 66 枚 `w:val="none"` 的边孩子全带 `w:sz="0"`，间距按边分档
 * ——顶/底 32 枚写 `w:space="1"`、左/右 32 枚写 `w:space="4"`（16 枚四边全说的样式正是这个形状）。
 * `w:between` 一枚样本都没有（46 份件、26 枚 `w:pBdr` 里 0 枚），它说的是**横**着的那一条，间距跟顶/底那一档 1，
 * 不替它另编一个数。
 *
 * 这四条聚合腿按 WPS 声明的那四条在段落这一层的说法派发：`AllBorders` = 四条边加中间那条、
 * `OutsideBorders` = 只有四条边（选区只有一段时这两条确实是同一句话，跨段才分得开）、
 * `HMiddleBorder` = 只有中间那条、`NoBorder` = 一起说不要。`VMiddleBorder`（内部竖框线）与
 * `CrossBorder`（横竖框线）在这里说不出——`CT_PBdr` 没有竖着的那一枚 inside 孩子（`w:bar` 说的是贯穿
 * 左边距的那一条，不是「内部」），所以段落那张列子不摆这两档，摆出来就是给用户一条会落错的腿。
 */
export type BorderLeg = "none" | "all" | "outside" | "insideH" | ParaBoxSide;

/** 段落这一层「四条边」与「中间那条」——`w:between` 是这一段与下一段之间的那一条。 */
const PARA_OUTER: readonly ParaEdge[] = ["top", "left", "bottom", "right"];
const PARA_MIDDLE: readonly ParaEdge[] = ["between"];

export const legPatch = (leg: BorderLeg): ParaBoxPatch => {
  const sides: Partial<Record<ParaEdge, EdgePatch>> = {};
  const at = (side: ParaEdge): void => {
    if (leg === "none") {
      sides[side] = { val: "none", sizeEighths: 0, spacePt: side === "left" || side === "right" ? 4 : 1, color: "auto" };
      return;
    }
    sides[side] = { val: SINGLE_EDGE.val, sizeEighths: SINGLE_EDGE.size, spacePt: SINGLE_EDGE.space, color: SINGLE_EDGE.color };
  };
  if (leg === "none" || leg === "all") [...PARA_OUTER, ...PARA_MIDDLE].forEach(at);
  else if (leg === "outside") PARA_OUTER.forEach(at);
  else if (leg === "insideH") PARA_MIDDLE.forEach(at);
  else at(leg);
  return { sides };
};

/**
 * 把补丁落到一段的框上：**换掉整个物件**而不是往共享的那一个里改——样式链借来的那一份是按引用带在字段上的
 * （`mergeParaBorders`），往里改就是替样式改它自己，别处借过同一套框的段落会跟着变。
 * 逐边合：没被点名的边继续留原件/样式链的那一份，被点名的边只把填了的属性换掉，其余沿用这一段这一边原来的
 * 说法。一条边说不出线型（原本没有线，这一击又没点线型）时写不进模型——那是替文件新编一条线。
 */
export function applyBoxPatch<S extends string>(
  borders: Partial<Record<S, TableBorder>> | undefined,
  patch: { sides?: Partial<Record<string, EdgePatch>> },
  shape: EdgeShape = PARA_EDGE_SHAPE,
): Partial<Record<S, TableBorder>> | undefined {
  if (!patch.sides) return borders;
  const next: Partial<Record<S, TableBorder>> = { ...borders };
  for (const [side, edge] of Object.entries(patch.sides) as [S, EdgePatch][]) {
    const was = next[side];
    const val = edge.val ?? was?.val;
    if (!val) continue;
    // 「说了不要线」时磅就是 0：件里 66 枚 `w:val="none"` 的边孩子没有一枚写别的数。留着一条 12 的磅在一句
    // 「这一边不画线」上，下一个读的人按它算盒子厚度，而这一段说的根本没有线。间距不跟着归零——那一句这一段
    // （或样式）自己说过，件里 `none` 的边也带着数（1 与 4 两档）。
    const off = val === "none";
    const merged: TableBorder = {
      val,
      size: edge.sizeEighths ?? (off ? shape.none.size : was?.size ?? shape.single.size),
      space: edge.spacePt ?? was?.space ?? (off ? shape.none.space : shape.single.space),
      color: edge.color ?? was?.color ?? "auto",
    };
    next[side] = merged;
  }
  return next;
}

/**
 * 表与格那一侧的边名：`CT_TblBorders`（§17.4.76）说得内部横线与两条斜线，`CT_TcBorders`（§17.4.77）
 * 只有这一格自己的边——格子说不出「内部」，那是表才有的那一枚孩子。斜线 `w:tl2br`/`w:tr2bl` 与
 * `w:between`、`w:start`/`w:end` 这一族还没接命令：页面画不出斜线，读写按原件带着走。
 */
export type TableBoxSide = "top" | "bottom" | "left" | "right" | "insideH" | "insideV";
export type CellBoxSide = "top" | "bottom" | "left" | "right";
/** 表那一侧的十条腿（`Bottom/Top/Left/Right`｜`NoBorder/AllBorders/OutsideBorders/CrossBorder`｜`HMiddle/VMiddle`）。 */
export type TableLeg = CellBoxSide | "none" | "all" | "outside" | "cross" | "insideH" | "insideV";
/** 格那一侧的腿：四条边 + 无框 + 每格都带框。内部横竖那三条在格选区里派发去向表。 */
export type CellLeg = CellBoxSide | "none" | "all";

/** 一句「这一边要画线」：线型与磅跟着这一击走，`w:space`/`w:color` 继续用件里那一边自己说过的。
 * 磅按表与格那一层的形状（`TABLE_SINGLE_EDGE`），不是段落那一族的 6：授权口径那两层里画线的边共 1,924 枚，
 * `w:sz` 落在 4 的 1,512 枚、12 的 334 枚、2 的 66 枚、没说的 12 枚（`.scratch/corpus-recheck.mjs` #2）——
 * 4 是多数档，不是唯一档。 */
const ON: EdgePatch = { val: "single", sizeEighths: TABLE_SINGLE_EDGE.size };
/** 一句「这一边不画线」：磅归 0 由 `applyBoxPatch` 落，间距不跟着归零（件里 `none` 的边也带着数）。 */
const OFF: EdgePatch = { val: "none" };
const TABLE_EDGE_LABELS: readonly [TableBoxSide, string][] = [
  ["top", "上"],
  ["left", "左"],
  ["bottom", "下"],
  ["right", "右"],
  ["insideH", "内部横缝"],
  ["insideV", "内部竖缝"],
];
const CELL_EDGE_LABELS: readonly [CellBoxSide, string][] = [
  ["top", "上"],
  ["left", "左"],
  ["bottom", "下"],
  ["right", "右"],
];
/** 对话框按层摆出来的边：表六条（多内部横竖）、格四条、段落四条。 */
export const TABLE_EDGES: readonly [TableBoxSide, string][] = TABLE_EDGE_LABELS;
export const CELL_EDGES: readonly [CellBoxSide, string][] = CELL_EDGE_LABELS;
const TABLE_SIDES: readonly TableBoxSide[] = TABLE_EDGE_LABELS.map(([side]) => side);
const CELL_SIDES: readonly CellBoxSide[] = CELL_EDGE_LABELS.map(([side]) => side);
/** 回显那本账按层逐边取，顺序与对话框一致。 */
export const TABLE_BOX_SIDES: readonly TableBoxSide[] = TABLE_SIDES;
export const CELL_BOX_SIDES: readonly CellBoxSide[] = CELL_SIDES;
const OUTER: readonly TableBoxSide[] = ["top", "left", "bottom", "right"];

/**
 * 表级那一腿落在 `w:tblBorders` 上的边集合。逐边写回时只有这里点名的边会被重拼，其余（含件写了而模型
 * 没有字段的 `w:between`/`w:themeColor`）按件的字节走。
 */
export function tableLegPatch(leg: TableLeg): BoxPatch<TableBoxSide> {
  const sides: Partial<Record<TableBoxSide, EdgePatch>> = {};
  const put = (edge: EdgePatch, list: readonly TableBoxSide[]): void => {
    for (const side of list) sides[side] = edge;
  };
  if (leg === "none") put(OFF, TABLE_SIDES);
  else if (leg === "all") put(ON, TABLE_SIDES);
  else if (leg === "outside") put(ON, OUTER);
  else if (leg === "cross") put(ON, ["insideH", "insideV"]);
  else if (leg === "insideH" || leg === "insideV") put(ON, [leg]);
  else put(ON, [leg]);
  return { sides };
}

/** 格级那一腿落在每一格自己的 `w:tcBorders` 上；`insideH`/`insideV` 在这里没有槽，见 `tableLegPatch`。 */
export function cellLegPatch(leg: CellLeg): BoxPatch<CellBoxSide> {
  const sides: Partial<Record<CellBoxSide, EdgePatch>> = {};
  const edge = leg === "none" ? OFF : ON;
  for (const side of leg === "none" || leg === "all" ? CELL_SIDES : [leg]) sides[side] = edge;
  return { sides };
}
