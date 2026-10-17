import { LINE_ONE } from "../core/docx.js";
import type { LineSpacingRule } from "../core/docx.js";
import { currentUnit, twipText, twipsFromUnit, unitDef } from "./doc-units.js";
import type { Unit } from "./doc-units.js";

/**
 * 「段落」对话框：竖向节奏那三个数——段前、段后、行距，外加一颗「如果定义了文档网格，则对齐到网格」的勾。
 * #132 把样式表代劳的那 8px 判死之后，这一族**只剩文件自己
 * 说过的那一条路**，而这条路此前没有任何一条 UI 命令能走。授权口径（46 份 WPS 写出的件，`.scratch/corpus-recheck2.mjs`；
 * 部件范围由 `.scratch/spacing-part-scope.py` 定死——「3,948」= 挂在真 `w:p` 的 `w:pPr` 里的 `w:spacing`，跨
 * `word/document.xml` 3,924 + `word/footnotes.xml` 12 + `word/endnotes.xml` 12 三件；页眉页脚那 74 块 `w:pPr` 一枚不写，
 * `word/styles.xml` 的 246 枚是样式自己的段落属性、不在这一档）：3,948 枚里说段前的只有 78 枚、说段后的
 * 244 枚——绝大多数两边都
 * 没说；行距那一家人人都连着规则一起说（`w:line` 与 `w:lineRule` 3,948/3,948 成对），规则档位 `auto` 2,890 /
 * `exact` 1,004 / `atLeast` 54。换一只口径只看 `word/document.xml`、且不过滤父亲，给的是说行距 3,924 / 说段前 68 /
 * 说段后 232，另有 14 枚同名的 `w:spacing` 坐在 `w:rPr` 里说的是字符间距而非段距，正文 7,430 段里 7,357 段写了
 * `w:pPr`（`.scratch/ind-spacing-band-shapes.py`，与 `.scratch/ind-spacing-band-shapes.ts` 两种解析各数一遍相符）。
 * 能画不能改，
 * 用户想说的是「0.5 行」而页面只会照着文件的旧数画，所以本片把这条命令接上。
 *
 * 摆位按 WPS 的 Writer 自己声明的那一族命令来：`wpsentries.kuip` 里
 * `CT_Home.RB_Paragraph.*.m_LineSpacing` 挂着 `LineSpacing1_0 / 1_5 / 2_0 / 2_5 / 3_0 / LineSpacingMore`
 * 六条（`wpscommon.kuip` 把前三条的字面写成 `1.0`/`1.5`/`2.0`），另有 `CT_Paragraph.RB_ParagraphOnParagraph.*` 的
 * `ChangeLineSpacing / IncreaseLineSpacing / ReduceLineSpacing / TightLineSpacing / LooseLineSpacing / CleanLineSpacing`。
 * 那一格下拉的五个倍数由 `renderer.ts` 直接落，**「更多」这一条走这里**：段前/段后与 最小值/固定值 都没有下拉档，
 * 只能在这个对话框里说。对话框上各条中文措辞（「段前」「行距」「固定值」…）是这一族软件的通用说法，
 * 本机这份 bundle 里 `.qm` 读不出（`doc-units.ts` 头记的是同一堵墙），**未取证**。
 *
 * **只写用户动过的那一句**。`w:line` 与 `w:lineRule` 在件里从来没单独出现过（授权口径那三件部件的 3,948 枚成对、
 * 0 枚只写一边，部件范围见本节开头），
 * 所以行距那一格落下去必然连着规则一起交；段前段后两个格子留空 = 这一笔不改（与「页面设置」里那三格距离同一个
 * 口径），因为「没说」与「说了 0」是两件事——#132 的测试钉过这一点。用户没碰过的格子一个字节都不写：段落从样式链
 * 借来的那个数（`paraLent`）本来就是样式的话，抄到段落身上就是替文件编一句话。
 *
 * 段前段后的格子按状态栏那格的测量单位读数（`doc-units.ts`），缇只有一份、在这里换算一次。格子按该单位的
 * twip 精确位数预填（「页面设置」那条同一个口径），所以没人改过的对话框按一次「确定」不改任何字节。
 * 件把同一个间距用两笔写（`w:before` 缇 + `w:beforeLines` 1/100 行，授权口径段前 16 枚、段后 64 枚这么写，80 枚里有
 * 68 枚的商落在 312 缇那一档，`.scratch/line-tiers.mjs` 同批）时，
 * 下面那行小字把用户填的数按**件自己的那一档**换算成「几行」摆出来——那一档的来路在 `linePitchTwips`。
 *
 * 第三颗勾是「如果定义了文档网格，则对齐到网格」（`w:snapToGrid`，§17.3.1.87）。它的归属也是 WPS 自己声明的：
 * `wpsentries.kuip:1108-1110` 在段落那一页挂了三条——`CT_Paragraph.RB_GridlinesOnParagraph`（name
 * ="ParagraphGroupGridlines"）、`…GridlinesInParaLayoutTab`（ParagraphGridlines）与
 * `…SnapToGridInParaLayoutTab`（**ParagraphSnapToGrid**）。13 份 `.kuip` 里没有任何 `DocumentGrid`/`LinePitch`
 * 命令，即节那一层的网格档位**没有可照的 WPS 命令**，本片不做；做的这一颗正是段落级的 `ParagraphSnapToGrid`。
 * 受众（`.scratch/corpus-recheck2.mjs` 与 `.scratch/line-tiers.mjs`，同一套授权口径 46 份）：`document.xml` 里 406 枚
 * `w:snapToGrid`，全部挂在 `w:pPr` 块里（不限层共 410 枚，`w:rPr` 里 0 枚，§17.3.1.87 不放它在那儿，读取也不看那儿）；
 * 说了这一格的正文段里**说开 180（裸写）、说关 226（`w:val="0"`）**，而说关的 226 枚有 206 枚落在声明了行网格的节下
 * （节级口径：按段落后面那一枚 `w:sectPr` 判定；剩下 20 枚在从来没要过网格的节里，关的是空气）；43/46 份首节声明了
 * 行网格，首节 `w:linePitch` 取 312 缇 41 份 / 360 三份 / 285 两份。拼法只有两种：说开**裸写** 180 枚、说关
 * `w:val="0"` 226 枚，没有一枚写 `"1"` 或 `"false"`；`w:pPr` 里它的前兄弟固定是 `adjustRightInd`（334/410 枚），
 * 后兄弟最多是 `spacing`（284 枚）与 `textAlignment`（74 枚），与 `PPR_ORDER` 同一条。
 */
const STYLE_ID = "solo-para-spacing-styles";

const CSS = `
.solo-pg-dialog {
  width: 340px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-pg-dialog form { display: grid; gap: 12px; padding: 14px 16px 16px; }
.solo-pg-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-pg-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-pg-dialog .solo-pg-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-pg-fields { display: grid; gap: 10px 12px; grid-template-columns: repeat(2, minmax(0, 1fr)); }
.solo-pg-row { display: grid; gap: 8px; grid-template-columns: minmax(0, 1.4fr) minmax(0, 1fr); align-items: end; }
.solo-pg-fields label, .solo-pg-row label { display: grid; grid-template-columns: minmax(0, 1fr); gap: 3px; font-size: 12px; color: #4b5563; }
.solo-pg-num { display: flex; align-items: center; gap: 6px; min-width: 0; }
.solo-pg-dialog input, .solo-pg-dialog select {
  flex: 1; min-width: 0; padding: 4px 6px; font: inherit; color: inherit;
  border: 1px solid #c9d3e6; border-radius: 6px; background: #fff;
}
.solo-pg-dialog select { width: 100%; }
.solo-pg-num span { color: #6b7280; flex: none; }
.solo-pg-hint { color: #6b7280; font-size: 11px; min-height: 14px; }
.solo-pg-snap { display: grid; gap: 3px; font-size: 12px; color: #4b5563; }
.solo-pg-snap > span { display: flex; align-items: center; gap: 7px; }
.solo-pg-snap input { width: 14px; height: 14px; flex: none; margin: 0; padding: 0; }
.solo-pg-error { margin: 0; min-height: 18px; color: #b42318; font-size: 12px; }
.solo-pg-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-pg-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-pg-dialog .solo-pg-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
.solo-pg-dialog .solo-pg-apply[disabled] { opacity: .5; cursor: default; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

/** 对话框交回来的那份数：一个键不存在就是「这一笔不改」，值为 0 是「这一段明写着不要这一笔」。 */
export type ParaSpacingPatch = {
  /** `w:spacing@w:before`，缇。 */
  beforeTwips?: number;
  /** `w:spacing@w:after`，缇。 */
  afterTwips?: number;
  /** `w:spacing@w:line` 的原值：`auto` 是 240 分之一行，`exact`/`atLeast` 是缇。 */
  line?: number;
  /** 与 `line` 同进同出的那一条规则。 */
  lineRule?: LineSpacingRule;
  /** `w:snapToGrid`：勾着 = 吸本节行网格，没勾 = 这一段明写着不吸。键不存在 = 这一笔不改。 */
  snapToGrid?: boolean;
};

/** 打开时对话框上摆的那份数，来自段落**解析后**的状态（样式链借来的也算，那正是页面上画的）。 */
export type ParaSpacingState = ParaSpacingPatch & {
  /** 件在这一段里认的一行多高（缇）：`w:before`／(`w:beforeLines`/100)。说不出来就没有那行小字。 */
  linePitchTwips?: number;
  /**
   * 段落所在那一节的行网格一格多高（`w:docGrid@w:linePitch`，缇）；`undefined` = 那一节没定义行网格
   * （§17.3.2.9 的 `w:type` 不是 `lines`，或整句没写）。勾着但没有网格时排版一点不动，所以这一格要说出来。
   */
  docGridTwips?: number;
  /**
   * 所选跨段且那几段对「吸不吸网格」说法不一致。这一格摆不出三态（缺省就是「吸」，所以没说的也摆成勾着），
   * 于是由那行小字说明；用户一旦改过勾选，值就按同一份写给所有所选段。
   */
  snapMixed?: boolean;
};

/** 一个 `w:line` 折算成「几倍行距」，只在 `auto` 那一族里有意义（`LINE_ONE` = 单倍，§17.3.1.27）。 */
const multipleOf = (line: number): number => line / LINE_ONE;

const ROWS: [string, string][] = [
  ["auto:1", "单倍行距"],
  ["auto:1.5", "1.5 倍行距"],
  ["auto:2", "2 倍行距"],
  ["auto", "多倍行距"],
  ["atLeast", "最小值"],
  ["exact", "固定值"],
];
/** 只有这三档要填「设置值」：单倍/1.5/2 自己就把数说完了。 */
const NEEDS_VALUE = new Set(["auto", "atLeast", "exact"]);

/** What one box holds: nothing typed means "not asked about", and anything unparseable is refused. */
type Reading = { value?: number; bad: boolean };

const readGap = (input: HTMLInputElement, u: Unit): Reading => {
  const text = input.value.trim();
  if (!text) return { bad: false };
  const v = Number(text);
  // 与「页面设置」同一条：不设上界，件里 `w:before` 收任何缇数；非数与负数拒绝。
  if (!Number.isFinite(v) || v < 0) return { bad: true };
  return { value: twipsFromUnit(v, u), bad: false };
};

/**
 * Opens 段落 for the paragraph the caret (or selection) sits on. `state` is what the page is drawing now;
 * a property no file and no style stated arrives as no value, and its row is shown as 单倍行距 rather than
 * being filled with an invented number——按下「确定」时没动过的那一格一个字节都不写。
 */
export function showParaSpacingDialog(state: ParaSpacingState, apply: (patch: ParaSpacingPatch) => void): void {
  document.querySelectorAll<HTMLDialogElement>(".solo-pg-dialog").forEach((old) => {
    old.close();
    old.remove();
  });
  injectStyles();
  const dialog = document.createElement("dialog");
  dialog.className = "solo-pg-dialog";
  dialog.setAttribute("aria-label", "段落");
  const u = currentUnit();
  const unitName = unitDef(u).name;
  const gap = (name: string, label: string, value: number | undefined): string =>
    `<label>${label}<span class="solo-pg-num"><input name="${name}" inputmode="decimal" autocomplete="off" value="${value === undefined ? "" : twipText(value)}"><span>${unitName}</span></span><small class="solo-pg-hint" data-hint="${name}"></small></label>`;
  dialog.innerHTML =
    `<form><header><h2>段落</h2><button type="button" class="solo-pg-close" aria-label="关闭">×</button></header>
    <div class="solo-pg-fields">
    ${gap("pg-before", "段前", state.beforeTwips)}
    ${gap("pg-after", "段后", state.afterTwips)}
    </div>
    <div class="solo-pg-row">
    <label>行距<select name="pg-rule">${ROWS.map(([v, t]) => `<option value="${v}">${t}</option>`).join("")}</select></label>
    <label><span data-value-label>设置值</span><span class="solo-pg-num"><input name="pg-value" inputmode="decimal" autocomplete="off" value=""><span data-value-unit></span></span></label>
    </div>
    <label class="solo-pg-snap"><span><input type="checkbox" name="pg-snap">如果定义了文档网格，则对齐到网格</span><small class="solo-pg-hint" data-hint="pg-snap"></small></label>
    <p class="solo-pg-error" role="status"></p>
    <footer><button type="button" class="solo-pg-cancel">取消</button><button type="submit" class="solo-pg-apply">确定</button></footer></form>`;

  const field = (name: string): HTMLInputElement => dialog.querySelector<HTMLInputElement>(`[name="${name}"]`)!;
  const before = field("pg-before");
  const after = field("pg-after");
  const value = field("pg-value");
  const snap = field("pg-snap");
  const rule = dialog.querySelector<HTMLSelectElement>('select[name="pg-rule"]')!;
  const valueUnit = dialog.querySelector<HTMLElement>("[data-value-unit]")!;
  const valueLabel = dialog.querySelector<HTMLElement>("[data-value-label]")!;
  const error = dialog.querySelector<HTMLElement>(".solo-pg-error")!;
  const applyButton = dialog.querySelector<HTMLButtonElement>(".solo-pg-apply")!;

  // 行距那一格预填成页面正在画的那一档；用户没换过它就不交这一笔（`touched` 只管这一格）。
  let touched = false;
  // 「对齐到网格」同一条纪律，另一只 `snapTouched`：这一枚开关的缺省是**吸**（§17.3.1.87 说 `w:snapToGrid`
  // 没写就算 on），所以没说的段落也摆成勾着——但勾着不等于用户说了，不交这一笔，件里那句借来的话照旧归样式。
  let snapTouched = false;
  const pick = (): string => {
    const { line, lineRule } = state;
    if (line === undefined) return "auto:1";
    if (lineRule === "exact") return "exact";
    if (lineRule === "atLeast") return "atLeast";
    const m = multipleOf(line);
    return m === 1 ? "auto:1" : m === 1.5 ? "auto:1.5" : m === 2 ? "auto:2" : "auto";
  };
  rule.value = pick();
  snap.checked = state.snapToGrid !== false;
  /** 「设置值」那一格装的是倍数还是长度，跟着行距那一档走；单倍/1.5/2 自己说完了数，格子禁用。 */
  const needsValue = (): boolean => NEEDS_VALUE.has(rule.value);
  const isMultiple = (): boolean => rule.value === "auto";
  /**
   * 上一回这一格装的是哪一种数。同一个单位之间来回切档（例如 固定值 ↔ 最小值，两笔都按缇）不动用户敲的字；
   * 跨单位时**不拿旧单位的数冒充新单位的数**——把 `w:line="500"`（固定值 500 缇）当成「500 倍」填进多倍行距
   * 那一格，就是替件编一个谁也没说过的数。这时候要么按段落**当前**那一档预填，要么空着。
   */
  let shownKind: "none" | "mult" | "len" = "none";
  const syncValueBox = (): void => {
    const need = needsValue();
    const kind = !need ? "none" : isMultiple() ? "mult" : "len";
    value.disabled = !need;
    valueLabel.textContent = kind === "mult" ? "倍数" : kind === "len" ? "行高" : "设置值";
    valueUnit.textContent = kind === "mult" ? "倍" : kind === "len" ? unitName : "";
    if (kind === shownKind) return;
    shownKind = kind;
    if (kind === "none") {
      value.value = "";
      return;
    }
    // 只有件那一档说的单位与这一格要的单位是同一个，才把那个数摆进来。
    const lends = kind === "mult" ? state.lineRule === "auto" : state.lineRule === rule.value;
    value.value = lends && state.line !== undefined ? (kind === "mult" ? String(multipleOf(state.line)) : twipText(state.line)) : "";
  };
  /** 段前段后与「对齐到网格」那三行小字：件自己两笔一起说时，把用户填的数换算成「几行」。说不出一档就没有这一行。 */
  const syncHints = (): void => {
    const pitch = state.linePitchTwips;
    for (const [name, input] of [["pg-before", before], ["pg-after", after]] as const) {
      const hint = dialog.querySelector<HTMLElement>(`[data-hint="${name}"]`)!;
      const v = Number(input.value.trim());
      hint.textContent = pitch && input.value.trim() && Number.isFinite(v) ? `≈ ${(twipsFromUnit(v, u) / pitch).toFixed(2)} 行（件里一行=${twipText(pitch)}${unitName}）` : "";
    }
    // 网格那一行的话只能说文件自己说的：本节有没有行网格、一格多高（`w:docGrid`）。没有网格时这一勾
    // 落下去排版一点不动（`render.ts` 的 `lineHeight` 拿不到 `pitchPx`），所以那句话必须摆出来，不让用户勾了才发现。
    const grid = dialog.querySelector<HTMLElement>('[data-hint="pg-snap"]')!;
    const pitchTxt = state.docGridTwips === undefined ? "" : `本节一格 ${twipText(state.docGridTwips)}${unitName}`;
    grid.textContent = state.docGridTwips === undefined
      ? "本节没有定义行网格，勾与不勾都不改排版。"
      : state.snapMixed && !snapTouched
        ? `所选几段说法不一致：有的吸，有的不吸（${pitchTxt}）。改一次勾选会把同一个答案写给全部。`
        : snap.checked
          ? `行高按本节的网格取整格：${pitchTxt}。`
          : `这一段不吸网格：行高只按它自己说的算（${pitchTxt} 不用）。`;
  };
  const build = (): ParaSpacingPatch | undefined => {
    const b = readGap(before, u);
    const a = readGap(after, u);
    if (b.bad || a.bad) {
      error.textContent = "段前与段后要填 0 或更大的数字，留空表示不改。";
      applyButton.disabled = true;
      return undefined;
    }
    const patch: ParaSpacingPatch = {};
    if (b.value !== undefined) patch.beforeTwips = b.value;
    if (a.value !== undefined) patch.afterTwips = a.value;
    if (snapTouched) patch.snapToGrid = snap.checked;
    if (touched) {
      if (needsValue()) {
        const v = Number(value.value.trim());
        if (!Number.isFinite(v) || v <= 0) {
          error.textContent = isMultiple() ? "多倍行距要填大于 0 的倍数。" : `行高要填大于 0 ${unitName}的数。`;
          applyButton.disabled = true;
          return undefined;
        }
        // `auto` 的 `w:line` 是 240 分之一行，`exact`/`atLeast` 的是缇：单位在这一格上分岔，换算只在这里做一次。
        patch.line = isMultiple() ? Math.round(v * LINE_ONE) : twipsFromUnit(v, u);
        patch.lineRule = isMultiple() ? "auto" : rule.value === "exact" ? "exact" : "atLeast";
      } else {
        const m = Number(rule.value.split(":")[1]);
        patch.line = Math.round(m * LINE_ONE);
        patch.lineRule = "auto";
      }
    }
    error.textContent = "";
    applyButton.disabled = false;
    return patch;
  };
  const refresh = (): void => {
    syncValueBox();
    syncHints();
    build();
  };
  // 与「页面设置」同一条：`input` 里不回头改写用户正在敲的格子，只在换档/开关时重排。
  dialog.addEventListener("input", () => {
    syncHints();
    build();
  });
  dialog.addEventListener("change", (e) => {
    const name = (e.target as HTMLElement).getAttribute("name");
    if (name === "pg-rule") touched = true;
    if (name === "pg-snap") snapTouched = true;
    syncValueBox();
    syncHints();
    build();
  });
  // The editor owns the page's shortcuts; a dialog typing over them would move the caret behind it.
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
  dialog.querySelector<HTMLButtonElement>(".solo-pg-close")!.onclick = shut;
  dialog.querySelector<HTMLButtonElement>(".solo-pg-cancel")!.onclick = shut;
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    const patch = build();
    if (!patch) return;
    shut();
    apply(patch);
  };
  document.body.append(dialog);
  syncValueBox();
  syncHints();
  build();
  dialog.showModal();
  before.focus();
  before.select();
}
