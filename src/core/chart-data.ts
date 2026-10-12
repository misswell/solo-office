import { parseRange, rangeSpec, refOf, type Rect } from "./sheet-ops.js";
import type { ChartSeries } from "./xlsx.js";

/** A1 chart references retain sheet identity, including apostrophes and ! in names. */
export function chartRange(spec: string, defaultSheet: string): { sheet: string; box: Rect } {
  const m = /^(?:(?:'((?:[^']|'')*)'|([^'!]+))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d*))?$/.exec(spec.trim());
  if (!m) throw new Error("请输入有效数据区域，例如 '销售 数据'!A1:C8");
  const box = parseRange(m[3] + (m[4] ? `:${m[4]}` : ""));
  if (box.c2 >= 16384 || box.r2 >= 1048576) throw new Error("数据区域超出 XLSX 行列范围");
  // Bound synchronous preview/cache allocation, independently of workbook size.
  if ((box.c2 - box.c1 + 1) * (box.r2 - box.r1 + 1) > 100000) throw new Error("图表数据区域最多包含 100,000 个单元格");
  return { sheet: m[1]?.replace(/''/g, "'") ?? m[2] ?? defaultSheet, box };
}

export const chartRef = (sheet: string, spec: string): string => `'${sheet.replace(/'/g, "''")}'!${spec.replace(/([A-Za-z]+)(\d+)/g, "$$$1$$$2")}`;

export function chartSeriesForRange(sheet: string, box: Rect, value: (ref: string) => unknown): ChartSeries[] {
  if (box.r2 <= box.r1) throw new Error("数据区域需要包含标题行和至少一行数据");
  const withCats = box.c2 > box.c1;
  const cats = Array.from({ length: box.r2 - box.r1 }, (_, i) => String(value(refOf(box.c1, box.r1 + i + 1)) ?? ""));
  const series: ChartSeries[] = [];
  for (let c = box.c1 + (withCats ? 1 : 0); c <= box.c2; c++) {
    series.push({
      name: String(value(refOf(c, box.r1)) ?? "") || `系列${series.length + 1}`,
      nameRef: chartRef(sheet, refOf(c, box.r1)),
      catRef: withCats ? chartRef(sheet, `${refOf(box.c1, box.r1 + 1)}:${refOf(box.c1, box.r2)}`) : undefined,
      valRef: chartRef(sheet, `${refOf(c, box.r1 + 1)}:${refOf(c, box.r2)}`),
      cats: withCats ? [...cats] : [],
      vals: Array.from({ length: box.r2 - box.r1 }, (_, i) => {
        const v = value(refOf(c, box.r1 + i + 1));
        return typeof v === "number" && Number.isFinite(v) ? v : null;
      }),
    });
  }
  return series;
}

export function chartDataFingerprint(series: ChartSeries[]): string {
  return JSON.stringify(series.map((s) => [s.name, s.cats, s.vals]));
}

/** Structural grid edits move the referenced cells even when A1 addresses are absolute. */
export function shiftChartRef(spec: string, defaultSheet: string, editedSheet: string, axis: "row" | "col", at: number, count: number): string {
  try {
    const { sheet, box } = chartRange(spec, defaultSheet);
    if (sheet.toLowerCase() !== editedSheet.toLowerCase()) return spec;
    const a = axis === "row" ? box.r1 : box.c1, b = axis === "row" ? box.r2 : box.c2;
    let from: number, to: number;
    if (count < 0) { from = a >= at ? a - count : a; to = b >= at ? b - count : b; }
    else {
      if (a >= at && b < at + count) return "#REF!";
      from = a < at ? a : a >= at + count ? a - count : at;
      to = b < at ? b : b >= at + count ? b - count : at - 1;
    }
    const next = axis === "row" ? { ...box, r1: from, r2: to } : { ...box, c1: from, c2: to };
    return chartRef(sheet, rangeSpec(next));
  } catch { return spec; }
}

export function chartSourceRange(series: ChartSeries[], defaultSheet: string): string | undefined {
  try {
    const first = series[0];
    if (!first?.nameRef || !first.valRef) return;
    const n = chartRange(first.nameRef, defaultSheet);
    const v = chartRange(first.valRef, defaultSheet);
    if (v.sheet !== n.sheet || v.box.c1 !== n.box.c1 || v.box.c2 !== n.box.c1 || v.box.r1 !== n.box.r1 + 1) return;
    let c1 = n.box.c1;
    if (first.catRef) {
      const cats = chartRange(first.catRef, defaultSheet);
      if (cats.sheet !== n.sheet || cats.box.c1 !== c1 - 1 || cats.box.c2 !== c1 - 1 || cats.box.r1 !== v.box.r1 || cats.box.r2 !== v.box.r2) return;
      c1--;
    }
    for (let i = 0; i < series.length; i++) {
      const s = series[i];
      const name = chartRange(s.nameRef ?? "", defaultSheet);
      const val = chartRange(s.valRef ?? "", defaultSheet);
      if (name.sheet !== n.sheet || val.sheet !== n.sheet || name.box.c1 !== n.box.c1 + i || name.box.c2 !== name.box.c1 || name.box.r1 !== n.box.r1 || name.box.r2 !== n.box.r1 || val.box.c1 !== name.box.c1 || val.box.c2 !== name.box.c1 || val.box.r1 !== v.box.r1 || val.box.r2 !== v.box.r2 || s.catRef !== first.catRef) return;
    }
    return chartRef(n.sheet, rangeSpec({ c1, r1: n.box.r1, c2: n.box.c1 + series.length - 1, r2: v.box.r2 }));
  } catch { return; }
}
