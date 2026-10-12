import type { Spreadsheet } from "../core/xlsx.js";
import { FILTER_OPERATORS, filterMatches, filterRange, filterText, type FilterOperator, type FilterRule } from "../core/filter.js";
import { refOf } from "../core/sheet-ops.js";

const labels: Record<FilterOperator, string> = { equal: "等于", notEqual: "不等于", lessThan: "小于", lessThanOrEqual: "小于或等于", greaterThan: "大于", greaterThanOrEqual: "大于或等于" };

export function openFilterEditor(host: HTMLElement, wb: Spreadsheet, sheetName: string, id: number, commit: (rule?: FilterRule) => void, sort: (descending: boolean) => void): void {
  host.querySelector(".filter-editor")?.remove();
  const sheet = wb.sheets.find((s) => s.name === sheetName), filter = sheet?.autoFilter;
  if (!sheet || !filter) return;
  wb.refreshFilters();
  const box = filterRange(filter.ref), old = filter.columns.find((c) => c.id === id)?.rule;
  const choices = new Map<string, { value: string; count: number }>();
  for (let row = box.r1 + 1; row <= box.r2; row++) {
    if (sheet.hiddenRows?.has(row) || !filter.columns.filter((c) => c.id !== id).every((c) => filterMatches(wb.value(sheetName, refOf(box.c1 + c.id, row)), c.rule))) continue;
    const value = filterText(wb.value(sheetName, refOf(box.c1 + id, row))), key = value.toLowerCase();
    const found = choices.get(key);
    choices.set(key, { value: found?.value ?? value, count: (found?.count ?? 0) + 1 });
  }
  if (old?.kind === "values") for (const value of [...old.values, ...(old.blank ? [""] : [])]) if (!choices.has(value.toLowerCase())) choices.set(value.toLowerCase(), { value, count: 0 });
  const dialog = document.createElement("dialog");
  dialog.className = "filter-editor";
  dialog.innerHTML = `<form><header><strong></strong><button type="button" class="filter-close">×</button></header>
    <div class="filter-sort"><button type="button" data-sort="asc">升序</button><button type="button" data-sort="desc">降序</button></div>
    <label>筛选方式<select name="filter-mode"><option value="values">按值选择</option><option value="custom">条件筛选</option></select></label>
    <section class="filter-values"><input name="filter-search" type="search" placeholder="搜索筛选值" aria-label="搜索筛选值"><div class="filter-check-tools"><button type="button" data-check="all">全选</button><button type="button" data-check="none">全不选</button></div><div class="filter-options"></div></section>
    <section class="filter-custom" hidden><label>条件一<select name="filter-op1"></select><input name="filter-val1" aria-label="条件一的值"></label><label><input type="checkbox" name="filter-second">使用第二个条件</label><select name="filter-combine"><option value="and">同时满足</option><option value="or">满足任一个</option></select><label>条件二<select name="filter-op2"></select><input name="filter-val2" aria-label="条件二的值"></label><small>文本支持 * 和 ? 通配符；~ 可转义通配符。</small></section>
    <p class="filter-error" role="alert"></p><footer><button type="button" class="filter-clear">清除此列筛选</button><button type="button" class="filter-cancel">取消</button><button type="submit" class="filter-apply">应用</button></footer></form>`;
  dialog.querySelector("strong")!.textContent = `筛选：${filterText(wb.value(sheetName, refOf(box.c1 + id, box.r1))) || refOf(box.c1 + id, box.r1)}`;
  const field = <T extends HTMLInputElement | HTMLSelectElement>(name: string): T => dialog.querySelector(`[name="${name}"]`)!;
  const mode = field<HTMLSelectElement>("filter-mode"), search = field<HTMLInputElement>("filter-search"), second = field<HTMLInputElement>("filter-second");
  const options = dialog.querySelector<HTMLElement>(".filter-options")!;
  const checks: HTMLInputElement[] = [];
  for (const { value, count } of [...choices.values()].sort((a, b) => a.value.localeCompare(b.value, "zh-CN", { numeric: true }))) {
    const label = document.createElement("label"), check = document.createElement("input"), text = document.createElement("span");
    check.type = "checkbox"; check.dataset.value = value;
    check.checked = old?.kind !== "values" || (value === "" ? old.blank : old.values.some((v) => v.toLowerCase() === value.toLowerCase()));
    text.textContent = `${value || "（空白）"} (${count})`;
    label.append(check, text); options.append(label); checks.push(check);
  }
  for (const n of [1, 2]) {
    const select = field<HTMLSelectElement>(`filter-op${n}`);
    for (const op of FILTER_OPERATORS) select.add(new Option(labels[op], op));
    if (old?.kind === "custom" && old.tests[n - 1]) { select.value = old.tests[n - 1].operator; field<HTMLInputElement>(`filter-val${n}`).value = old.tests[n - 1].value; }
  }
  mode.value = old?.kind === "custom" ? "custom" : "values";
  second.checked = old?.kind === "custom" && old.tests.length === 2;
  field<HTMLSelectElement>("filter-combine").value = old?.kind === "custom" && !old.and ? "or" : "and";
  const refresh = (): void => {
    dialog.querySelector<HTMLElement>(".filter-values")!.hidden = mode.value !== "values";
    dialog.querySelector<HTMLElement>(".filter-custom")!.hidden = mode.value !== "custom";
    field<HTMLSelectElement>("filter-op2").disabled = !second.checked;
    field<HTMLInputElement>("filter-val2").disabled = !second.checked;
    for (const check of checks) check.parentElement!.hidden = !(check.dataset.value || "（空白）").toLowerCase().includes(search.value.toLowerCase());
  };
  dialog.addEventListener("change", refresh); search.oninput = refresh;
  dialog.querySelectorAll<HTMLButtonElement>("[data-check]").forEach((b) => b.onclick = () => { for (const check of checks) check.checked = b.dataset.check === "all"; });
  dialog.addEventListener("keydown", (e) => e.stopPropagation());
  const error = dialog.querySelector<HTMLElement>(".filter-error")!;
  const apply = (rule?: FilterRule): void => { try { commit(rule); dialog.close(); } catch (e) { error.textContent = e instanceof Error ? e.message : String(e); } };
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    if (mode.value === "values") {
      if (checks.every((c) => c.checked)) { apply(); return; }
      apply({ kind: "values", values: checks.filter((c) => c.checked && c.dataset.value !== "").map((c) => c.dataset.value!), blank: checks.some((c) => c.checked && c.dataset.value === "") });
    } else {
      const tests = [1, ...(second.checked ? [2] : [])].map((n) => ({ operator: field<HTMLSelectElement>(`filter-op${n}`).value as FilterOperator, value: field<HTMLInputElement>(`filter-val${n}`).value }));
      apply({ kind: "custom", and: field<HTMLSelectElement>("filter-combine").value === "and", tests });
    }
  };
  dialog.querySelectorAll<HTMLButtonElement>("[data-sort]").forEach((b) => b.onclick = () => { try { sort(b.dataset.sort === "desc"); dialog.close(); } catch (e) { error.textContent = String(e); } });
  dialog.querySelector<HTMLButtonElement>(".filter-clear")!.onclick = () => apply();
  for (const cls of ["filter-close", "filter-cancel"]) dialog.querySelector<HTMLButtonElement>(`.${cls}`)!.onclick = () => dialog.close();
  if (filter.unsupported) { error.textContent = filter.unsupported; dialog.querySelectorAll<HTMLButtonElement>(".filter-apply,.filter-clear,[data-sort]").forEach((b) => b.disabled = true); }
  dialog.addEventListener("close", () => { dialog.remove(); host.focus(); }, { once: true });
  host.append(dialog); refresh(); dialog.showModal();
}
