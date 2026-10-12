import type { ChartModel, Spreadsheet } from "../core/xlsx.js";
import { CHART_PALETTES, renderChartSvg } from "../core/chart.js";

/** Modal drafts use the production model, so invalid references never commit partial edits. */
export function openChartEditor(host: HTMLElement, wb: Spreadsheet, sheet: string, index: number, commit: (next: ChartModel) => void, remove: () => void): void {
  host.querySelector<HTMLDialogElement>(".chart-editor")?.close();
  const chart = wb.sheet(sheet)?.charts[index];
  if (!chart) return;
  const dialog = document.createElement("dialog");
  dialog.className = "chart-editor";
  dialog.setAttribute("aria-label", "编辑图表");
  dialog.innerHTML = `<form><header><h2>编辑图表</h2><button type="button" class="chart-close" aria-label="关闭">×</button></header>
    <div class="chart-fields"><label>标题<input name="chart-title" autocomplete="off"></label>
    <label>图表类型<select name="chart-type"><option value="bar">柱状图</option><option value="horizontal">条形图</option><option value="line">折线图</option><option value="area">面积图</option><option value="pie">饼图</option><option value="doughnut">环形图</option><option value="scatter">散点图</option></select></label>
    <label>配色<select name="chart-palette"><option value="keep">保留现有配色</option><option value="office">经典</option><option value="ocean">海洋</option><option value="warm">暖色</option></select></label>
    <label class="chart-source">数据区域<input name="chart-source" autocomplete="off" placeholder="例如 '销售 数据'!A1:C8"><small>第一行为系列名称；多列数据的第一列为分类或 X 坐标。留空保留各系列现有引用。</small></label>
    <label class="chart-legend"><input type="checkbox" name="chart-legend">显示图例</label></div>
    <div class="chart-preview" aria-label="图表预览"></div><p class="chart-error" role="status"></p>
    <footer><button type="button" class="chart-delete">删除图表</button><span></span><button type="button" class="chart-cancel">取消</button><button type="submit" class="chart-apply">应用</button></footer></form>`;
  const field = <T extends HTMLInputElement | HTMLSelectElement>(name: string): T => dialog.querySelector(`[name="${name}"]`)!;
  const title = field<HTMLInputElement>("chart-title"), type = field<HTMLSelectElement>("chart-type"), palette = field<HTMLSelectElement>("chart-palette"), source = field<HTMLInputElement>("chart-source"), legend = field<HTMLInputElement>("chart-legend");
  title.value = chart.title;
  type.value = chart.kind === "bar" && chart.barDir === "bar" ? "horizontal" : chart.kind;
  source.value = chart.srcRef ?? "";
  legend.checked = chart.showLegend !== false;
  const preview = dialog.querySelector<HTMLElement>(".chart-preview")!, error = dialog.querySelector<HTMLElement>(".chart-error")!, apply = dialog.querySelector<HTMLButtonElement>(".chart-apply")!;
  let draft: ChartModel | undefined;
  const refresh = (): void => {
    try {
      draft = wb.prepareChartEdit(sheet, index, { kind: type.value === "horizontal" ? "bar" : type.value as ChartModel["kind"], barDir: type.value === "horizontal" ? "bar" : "col", title: title.value, source: source.value, colors: palette.value === "keep" ? undefined : CHART_PALETTES[palette.value], showLegend: legend.checked });
      preview.innerHTML = renderChartSvg(draft, 500, 280);
      error.textContent = "";
      apply.disabled = false;
    } catch (e) {
      draft = undefined;
      error.textContent = e instanceof Error ? e.message : String(e);
      apply.disabled = true;
    }
  };
  preview.innerHTML = renderChartSvg(wb.chartForDisplay(sheet, chart), 500, 280);
  dialog.addEventListener("input", refresh);
  dialog.addEventListener("change", refresh);
  dialog.addEventListener("keydown", (e) => e.stopPropagation());
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    refresh();
    if (!draft) return;
    dialog.close();
    commit(draft);
  };
  const close = (): void => dialog.close();
  dialog.querySelector<HTMLButtonElement>(".chart-close")!.onclick = close;
  dialog.querySelector<HTMLButtonElement>(".chart-cancel")!.onclick = close;
  dialog.querySelector<HTMLButtonElement>(".chart-delete")!.onclick = () => { dialog.close(); remove(); };
  dialog.addEventListener("close", () => { dialog.remove(); host.focus(); }, { once: true });
  host.append(dialog);
  refresh();
  dialog.showModal();
  title.focus();
}
