import { TAB_LEADER_TOKENS } from "../core/docx.js";
import type { TabStop } from "../core/docx.js";
import { currentUnit, twipText, twipsFromUnit, unitDef } from "./doc-units.js";
import type { Unit } from "./doc-units.js";

/**
 * 「制表位」对话框：一段那一列 `w:tabs` 的**逐位置**编辑，加上件级的那一枚 `word/settings.xml/w:defaultTabStop`。
 * 摆位照 WPS 自己的声明：`office6/res/commands/wps/ongmani/CT_Home.kuip:309` 在主页段落组里挂了
 * `<KxLegacyTriggerCommand aid="780" customTip="Tabs" dialogHint="1" idMso="Tabs" ksoCmd="Tabs"
 * icon="ruler_format_tabs_kd" text="@Tabs"/>`——`dialogHint="1"` 说的是这一条开的是一张对话框，而 :337 把它排在
 * `RB_Paragraph`（`tb_Paragraph`）里 `Sort` 之后，:315 又把它列进「段落」下拉的最后一条。所以入口在开始页段落那一簇的
 * 尾巴上，命令本身只负责开这张对话框。WPS 那张对话框里分几栏、各栏什么字面**未取证**（`.qm` 只证词在不在：
 * `mui/zh_CN/wps.qm` 制表位 14 处、默认制表位 1 处、前导符 4 处、对齐方式 16 处、左对齐 29、居中对齐 15、右对齐 31、
 * 小数点对齐 3、全部清除 1；`kso.qm` 制表位 3、默认制表位 1、小数点对齐 1、全部清除 1；读的是 UTF-16**BE**，
 * 且要按两种字节对齐各扫一遍——只按偏移 0 解码会把有的词读成 0），这里的措辞是这一族软件的通用说法，
 * 摆的是家已有落笔处的那几栏。**WPS 的字面一条也不抄进产品。**
 *
 * 对齐只摆 左/居中/右 三档，不是四档或六档：授权 46 份 320 块 `w:tabs`、462 枚孩子里 `w:val` 只拼了四种——
 * `left` 266、`right` 144、`center` 50、`clear` 2，**`decimal`/`bar` 各 0 枚**（`.scratch/tabs-slot-shape.py`）。
 * 而读取那一头（`tabStopsOf`）把模型收不下的档一律收成 `left`，所以在这里摆「小数点对齐」会写出一句
 * `w:val="left"`——一枚说谎的按钮。那一档要摆，得先让模型收得下、页面画得出来（记在待办里）。
 * `clear` 也不是本对话框的一档：交回来的清单只说「这些位置归段落」，被从清单里去掉的那些位置由写手
 * （`paraTabStopEdits`）分两种说法——入件时（段自己的字节或样式链）确实有话的那一格写成 `w:val="clear"`，
 * 正是件里那 2 枚的拼法；同一趟里前一击添上、随后又被删掉的那一格则从块里摘走，段落不替一句没人说过的话立案。
 *
 * 前导符摆 无 + `TAB_LEADER_TOKENS` 六档：常量与页面那 27 行 `[data-leader=…]` 规则同一个主人（`docx.ts` 的
 * `TAB_LEADER_TOKENS` 头注记的就是这条纪律），所以这里直接从它生成，不另拼一套。件里的受众只有 `dot` 94 枚
 * （枚枚 `dot`），其余五档 0 枚——它们画得出来（CSS 各有一档），只是没有件真的这样写过。
 *
 * **只写用户动过的那两句**。清单与开窗时那一份逐位置相同就整键不交（`patch.stops` 缺席），所以「原样确定」在任何
 * 选区下都不改一个字节——`tests/doc-tab-stops.test.ts` 钉的就是这一条；跨段说法不一致时（`mixed`）不改就不替全部
 * 重写，改了才把同一份清单写给每一段。「默认制表位」留空 = 这一笔不改，而件没说这一枚时（`state.defaultTwips` 是
 * `undefined`）格子也空着——不替文件编一个 720：`w:defaultTabStop` 在授权 46 份里 46/46 都说了这一枚
 * （420 × 43、720 × 3，枚枚自闭合、只有 `w:val`，`.scratch/settings-tabstop-shape.py`），但**这一份**没说就是没说。
 *
 * 位置按状态栏那格的测量单位读数（`doc-units.ts`），缇只有一份、在这里换算一次；格子按该单位的 twip 精确位数
 * 预填，所以没人动过的对话框按一次「确定」不改任何字节。`w:pos` 是 `ST_SignedTwipsMeasure`，读取那头
 * （`tabStopsOf`）只认「是不是个数」，所以这里不设上下界，只拒非数——拿 15 缇/px 那口径去猜一个「合理的最大值」，
 * 是给件加一句它没说过的话。
 *
 * 位置升序、位置不重复：件就是这么写的（318/320 块升序、0 块降序、0 块重复位置），清单在`设置`时按位置并入并
 * 重新排序，与写手 `tabsSlotXml` 插入新孩子的口径是同一句，两边各排一次不会打架。
 */
const STYLE_ID = "solo-tab-stops-styles";

const CSS = `
.solo-ts-dialog {
  width: 372px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-ts-dialog form { display: grid; gap: 10px; padding: 14px 16px 16px; }
.solo-ts-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-ts-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-ts-dialog .solo-ts-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-ts-top { display: grid; gap: 10px; grid-template-columns: minmax(0, 1fr) auto; align-items: start; }
.solo-ts-top > label { display: grid; gap: 3px; font-size: 12px; color: #4b5563; min-width: 0; }
.solo-ts-dialog select, .solo-ts-dialog input {
  padding: 4px 6px; font: inherit; color: inherit;
  border: 1px solid #c9d3e6; border-radius: 6px; background: #fff;
}
.solo-ts-listbox { min-width: 0; width: 100%; font: 12px/1.45 ui-monospace, Menlo, monospace; }
.solo-ts-btns { display: grid; gap: 6px; }
.solo-ts-btns button { padding: 4px 12px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; white-space: nowrap; }
.solo-ts-fields { display: grid; gap: 8px 12px; grid-template-columns: minmax(0, 1.2fr) minmax(0, 1fr); }
.solo-ts-fields label, .solo-ts-align { display: grid; gap: 3px; font-size: 12px; color: #4b5563; min-width: 0; }
.solo-ts-num { display: flex; align-items: center; gap: 6px; min-width: 0; }
.solo-ts-num input { flex: 1; min-width: 0; }
.solo-ts-num span { color: #6b7280; flex: none; }
.solo-ts-align { border: 0; margin: 0; padding: 0; display: grid; gap: 2px; }
.solo-ts-align legend { padding: 0; font-size: 12px; color: #4b5563; }
.solo-ts-align span { display: flex; align-items: center; gap: 6px; font-size: 12px; }
.solo-ts-align input { width: 13px; height: 13px; margin: 0; padding: 0; }
.solo-ts-hint { color: #6b7280; font-size: 11px; min-height: 14px; }
.solo-ts-error { margin: 0; min-height: 18px; color: #b42318; font-size: 12px; }
.solo-ts-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-ts-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-ts-dialog .solo-ts-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
.solo-ts-dialog .solo-ts-apply[disabled] { opacity: .5; cursor: default; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

/** 「制表位」对话框交回来的两句：缺键就是那一句不改。 */
export type TabStopPatch = {
  /**
   * 段落此刻要认的**整份**清单（逐位置，`w:val="clear"` 不在这里——被去掉的那些位置由写手补出那一句）。
   * 与开窗时那一份一模一样时这一键缺席，所以「原样确定」不落字节。
   */
  stops?: TabStop[];
  /** `word/settings.xml` 的 `w:defaultTabStop@w:val`，缇。 */
  defaultTwips?: number;
};

/** 对话框打开时页面上正在画的那一份（段落的是解析过样式链的合并结果，样式借来的也算——那正是纸上画的）。 */
export type TabStopState = {
  stops?: TabStop[];
  /** 件级那一枚；`undefined` = 这一份件的 `word/settings.xml` 没说，格子留空。 */
  defaultTwips?: number;
  /** 所选跨段且那几段的制表位清单不一致：清单摆第一段的样子，由那行小字说出来。 */
  mixed?: boolean;
};

const ALIGN_LABELS: Record<string, string> = { left: "左对齐", center: "居中对齐", right: "右对齐" };
const LEADER_LABELS: Record<string, string> = {
  spaces: "空格",
  dot: "点线",
  dash: "短划线",
  line: "长划线",
  heavy: "粗线",
  middleDot: "中间点",
};
const alignLabel = (val: string): string => ALIGN_LABELS[val] ?? val;
const leaderLabel = (leader?: string): string => (leader ? LEADER_LABELS[leader] ?? leader : "无");
const byPos = (a: TabStop, b: TabStop): number => a.posTwips - b.posTwips;
/** 两份清单逐位置、逐档、逐前导符比一遍；只有真的一样才不交这一笔。 */
const sameStops = (a: TabStop[], b: TabStop[]): boolean => JSON.stringify(a) === JSON.stringify(b);

/**
 * 打开「制表位」。`state` 是页面此刻画的那一份；改完把新的整份清单与（动过的）默认制表位一次交回。
 */
export function showTabStopsDialog(state: TabStopState, apply: (patch: TabStopPatch) => void): void {
  document.querySelectorAll<HTMLDialogElement>(".solo-ts-dialog").forEach((old) => {
    old.close();
    old.remove();
  });
  injectStyles();
  const dialog = document.createElement("dialog");
  dialog.className = "solo-ts-dialog";
  dialog.setAttribute("aria-label", "制表位");
  const u: Unit = currentUnit();
  const unitName = unitDef(u).name;
  const shownStops = [...(state.stops ?? [])].sort(byPos);
  // 手上这份是工作副本：每一枚都另取物件，`清除`/`设置` 改的是它，绝不动 `state.stops`——那一份可能是样式链
  // 上的数组本体（`mergeTabStops` 段落一枚没说时交回链上那一份的**同一个引用**）。
  const working: TabStop[] = shownStops.map((s) => ({ ...s }));
  dialog.innerHTML =
    `<form><header><h2>制表位</h2><button type="button" class="solo-ts-close" aria-label="关闭">×</button></header>
    <div class="solo-ts-top">
    <label>制表位<select class="solo-ts-listbox" name="ts-list" size="6"></select></label>
    <div class="solo-ts-btns">
    <button type="button" name="ts-set">设置</button>
    <button type="button" name="ts-del">清除</button>
    <button type="button" name="ts-clear">全部清除</button>
    </div>
    </div>
    <div class="solo-ts-fields">
    <label>制表位位置<span class="solo-ts-num"><input name="ts-pos" inputmode="decimal" autocomplete="off"><span>${unitName}</span></span></label>
    <label>前导符<select name="ts-leader">${["", ...TAB_LEADER_TOKENS]
      .map((t) => `<option value="${t}">${t === "" ? "无" : LEADER_LABELS[t]}</option>`)
      .join("")}</select></label>
    <fieldset class="solo-ts-align"><legend>对齐方式</legend>${Object.entries(ALIGN_LABELS)
      .map(([v, t]) => `<span><input type="radio" name="ts-val" value="${v}" id="ts-val-${v}"><label for="ts-val-${v}">${t}</label></span>`)
      .join("")}</fieldset>
    </div>
    <small class="solo-ts-hint" data-hint="ts-list"></small>
    <label>默认制表位<span class="solo-ts-num"><input name="ts-default" inputmode="decimal" autocomplete="off" value="${state.defaultTwips === undefined ? "" : twipText(state.defaultTwips)}"><span>${unitName}</span></span></label>
    <small class="solo-ts-hint" data-hint="ts-default"></small>
    <p class="solo-ts-error" role="status"></p>
    <footer><button type="button" class="solo-ts-cancel">取消</button><button type="submit" class="solo-ts-apply">确定</button></footer></form>`;

  const list = dialog.querySelector<HTMLSelectElement>('[name="ts-list"]')!;
  const pos = dialog.querySelector<HTMLInputElement>('[name="ts-pos"]')!;
  const leader = dialog.querySelector<HTMLSelectElement>('[name="ts-leader"]')!;
  const def = dialog.querySelector<HTMLInputElement>('[name="ts-default"]')!;
  const radios = [...dialog.querySelectorAll<HTMLInputElement>('input[name="ts-val"]')];
  const error = dialog.querySelector<HTMLElement>(".solo-ts-error")!;
  const applyButton = dialog.querySelector<HTMLButtonElement>(".solo-ts-apply")!;
  const listHint = dialog.querySelector<HTMLElement>('[data-hint="ts-list"]')!;
  const defHint = dialog.querySelector<HTMLElement>('[data-hint="ts-default"]')!;

  const chosenVal = (): string => (radios.find((r) => r.checked) ?? radios[0]!).value;
  const setVal = (val: string): void => {
    const hit = radios.find((r) => r.value === val);
    // 清单里若有模型收不下的档（读取那头不会交出来，这里是给坏件留的后路），认不出来就回落到第一档，
    // 而不是让「对齐方式」这一组一枚都不勾——没勾的那一组按「设置」会静默写成第一档，更容易骗人。
    (hit ?? radios[0]!).checked = true;
  };

  /** 列表的一行说什么：档、位置、前导符，三样都是页面上真画得出的那三样。 */
  const rowText = (s: TabStop): string => `${alignLabel(s.val)}  ${twipText(s.posTwips)} ${unitName}  ${leaderLabel(s.leader)}`;
  const renderList = (selectPos?: number): void => {
    const at = list.selectedIndex;
    list.innerHTML = working.map((s) => `<option value="${s.posTwips}">${rowText(s)}</option>`).join("");
    list.size = Math.max(4, Math.min(6, working.length || 4));
    const want = selectPos === undefined ? at : working.findIndex((s) => s.posTwips === selectPos);
    list.selectedIndex = want >= 0 ? want : working.length ? Math.min(want, working.length - 1) : -1;
    syncRow();
  };
  /** 选中行挪到哪儿，下面那三栏就摆成那一枚说的话——「设置」同一位置是同位并入，不会顺手改档。 */
  const syncRow = (): void => {
    const s = working[list.selectedIndex];
    if (!s) return;
    pos.value = twipText(s.posTwips);
    setVal(s.val);
    leader.value = s.leader ?? "";
  };

  /** 一个格子的读数：留空是「没说」，非数是「说坏了」。 */
  const readTwips = (input: HTMLInputElement): { value?: number; bad: boolean } => {
    const text = input.value.trim();
    if (!text) return { bad: false };
    const v = Number(text);
    // `w:pos` 与 `w:defaultTabStop` 都不设上下界（前者是有符号缇数），所以只拒非数；件里 `w:defaultTabStop`
    // 只说过 420 与 720 两个数，那是件的用词，不是这一格的合法性。
    if (!Number.isFinite(v)) return { bad: true };
    return { value: twipsFromUnit(v, u), bad: false };
  };

  dialog.querySelector<HTMLButtonElement>('[name="ts-set"]')!.onclick = () => {
    const r = readTwips(pos);
    if (r.bad || r.value === undefined) return;
    const stop: TabStop = { posTwips: r.value, val: chosenVal() as TabStop["val"], ...(leader.value ? { leader: leader.value } : {}) };
    const at = working.findIndex((s) => s.posTwips === stop.posTwips);
    // 同一个位置只容一枚：件里 0 块有重复位置。
    if (at >= 0) working[at] = stop;
    else working.push(stop);
    working.sort(byPos);
    renderList(stop.posTwips);
    build();
  };
  dialog.querySelector<HTMLButtonElement>('[name="ts-del"]')!.onclick = () => {
    if (list.selectedIndex < 0) return;
    working.splice(list.selectedIndex, 1);
    renderList();
    build();
  };
  dialog.querySelector<HTMLButtonElement>('[name="ts-clear"]')!.onclick = () => {
    working.length = 0;
    renderList();
    build();
  };
  list.onchange = () => syncRow();

  const build = (): TabStopPatch | undefined => {
    const p = readTwips(pos);
    const d = readTwips(def);
    working.sort(byPos);
    if (p.bad || d.bad) {
      error.textContent = "制表位位置与默认制表位要填数字，留空表示不改。";
      applyButton.disabled = true;
      return undefined;
    }
    error.textContent = "";
    applyButton.disabled = false;
    const patch: TabStopPatch = {};
    if (!sameStops(working, shownStops)) patch.stops = working.map((s) => ({ ...s }));
    // 与开窗时那一个数不同才算用户说了这一句：件说的是 420、格子里也是 420，按「确定」不该改成 420。
    if (d.value !== undefined && d.value !== state.defaultTwips) patch.defaultTwips = d.value;
    return patch;
  };
  const syncHints = (): void => {
    const pitch = state.defaultTwips ?? 420;
    listHint.textContent = state.mixed
      ? `所选几段的制表位不一致，这里摆的是第一段那份。改过的清单会按同一份写给全部所选段。`
      : working.length
        ? `位置从正文左缘起算；一段没有制表位时 Tab 迈 ${twipText(pitch)}${unitDef(u).name}（默认制表位）。`
        : `这一段没有制表位：Tab 只按默认制表位走。填位置、选对齐与前导符，再按「设置」。`;
    defHint.textContent = state.defaultTwips === undefined
      ? "这一份件的 word/settings.xml 没说默认制表位；留空就是不说，填了才写那一句。"
      : `一段没有制表位时 Tab 迈多远（${twipText(state.defaultTwips)}${unitName}）；这一段自己那几格不算在内。`;
  };
  const refresh = (): void => {
    syncHints();
    build();
  };
  dialog.addEventListener("input", refresh);
  dialog.addEventListener("change", refresh);
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
  dialog.querySelector<HTMLButtonElement>(".solo-ts-close")!.onclick = shut;
  dialog.querySelector<HTMLButtonElement>(".solo-ts-cancel")!.onclick = shut;
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    const patch = build();
    if (!patch) return;
    shut();
    apply(patch);
  };
  document.body.append(dialog);
  renderList(working.length ? working[0]!.posTwips : undefined);
  // 开窗时选中的那一枚同时填进上面那三栏：第一档「左对齐」本来就是缺省，所以没人按「设置」时这一格一个字也不写。
  setVal(working[0]?.val ?? "left");
  pos.value = working.length ? twipText(working[0]!.posTwips) : "";
  leader.value = working[0]?.leader ?? "";
  refresh();
  dialog.showModal();
  def.focus();
}
