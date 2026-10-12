import { buildXml, parseOrderedXml, rawAttrs, rawChildren, rawFind, rawFindAll, tagLocalOf, tagOf } from "./opc.js";
import type { ChartModel } from "./xlsx.js";

/**
 * Chart rendering and generation. `renderChartSvg` draws the model as SVG for
 * the sheet canvas; `chartPartXml` emits the DrawingML chart part a new chart
 * saves as. Both are pure so the round trip is testable without a DOM.
 * REA/Ghidra observation (2026-10-11): excelrw's FUN_00183858 calls the
 * imported KCTSeries::setDataSourceContext at 0x183dd8; createCoreChart and
 * setStyleID are separate boundaries. This supports owning series references
 * separately from type/style. XML mapping and rendering follow OOXML and
 * independent-consumer checks, not an unverified claim about WPS internals.
 * Raw Evidence: .scratch/chart-evidence-20261011/ (never packaged).
 */

export const CHART_PALETTES: Record<string, string[]> = {
  ocean: ["087F8C", "5AAA95", "86A873", "BBCF9B", "386641", "A7C957"],
  warm: ["C8553D", "E08E45", "F8B259", "D4A373", "9C6644", "B56576"],
  office: ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "264478", "9E480E"],
};
const paletteOf = (chart: ChartModel): string[] => {
  const colors = chart.colors?.filter((c) => /^[0-9a-f]{6}$/i.test(c));
  return colors?.length ? colors : CHART_PALETTES.office;
};

const esc = (s: string): string => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** A "nice" axis top: rounds up to 1/2/5×10^k so ticks read like a spreadsheet's. */
export function niceMax(v: number): number {
  if (v <= 0) return 1;
  const exp = Math.floor(Math.log10(v));
  const base = Math.pow(10, exp);
  for (const m of [1, 2, 2.5, 5, 10]) if (v <= m * base) return m * base;
  return 10 * base;
}

interface Box {
  w: number;
  h: number;
}

function frame(chart: ChartModel, b: Box): string {
  const body = b.w >= 240 && b.h >= 120;
  const title = chart.title && body ? `<text x="${b.w / 2}" y="20" text-anchor="middle" font-size="13" font-weight="600" fill="#333">${esc(chart.title)}</text>` : "";
  return (
    `<rect x="0" y="0" width="${b.w}" height="${b.h}" rx="6" fill="#fff" stroke="#d9dde2"/>` +
    (body ? `<rect x="6" y="6" width="${b.w - 12}" height="${b.h - 12}" rx="4" fill="none" stroke="#eceff3"/>` : "") +
    title
  );
}

function legend(items: string[], b: Box, chart: ChartModel): { y: number; svg: string } {
  if (!items.length || chart.showLegend === false) return { y: b.h, svg: "" };
  const y = b.h - 18;
  let x = 14;
  let svg = "";
  items.forEach((name, i) => {
    svg += `<rect x="${x}" y="${y - 8}" width="9" height="9" fill="#${paletteOf(chart)[i % paletteOf(chart).length]}"/>`;
    svg += `<text x="${x + 13}" y="${y}" font-size="10.5" fill="#444">${esc(name.slice(0, 14))}</text>`;
    x += 20 + name.length * 6.2;
  });
  return { y, svg };
}

export function renderChartSvg(chart: ChartModel, widthPx: number, heightPx: number): string {
  const b = { w: widthPx, h: heightPx };
  const flat = chart.series.filter((s) => s.vals.some((v) => v !== null));
  if (!flat.length) return `<svg xmlns="http://www.w3.org/2000/svg" width="${b.w}" height="${b.h}" viewBox="0 0 ${b.w} ${b.h}">${frame(chart, b)}</svg>`;
  if (chart.kind === "pie" || chart.kind === "doughnut") return `<svg xmlns="http://www.w3.org/2000/svg" width="${b.w}" height="${b.h}" viewBox="0 0 ${b.w} ${b.h}">${pieSvg(chart, b)}</svg>`;
  if (chart.kind === "scatter") return `<svg xmlns="http://www.w3.org/2000/svg" width="${b.w}" height="${b.h}" viewBox="0 0 ${b.w} ${b.h}">${scatterSvg(chart, b)}</svg>`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${b.w}" height="${b.h}" viewBox="0 0 ${b.w} ${b.h}">${axisChartSvg(chart, b)}</svg>`;
}

function axisChartSvg(chart: ChartModel, b: Box): string {
  const n = Math.max(1, ...chart.series.map((s) => s.vals.length));
  const cats = chart.series[0]?.cats.length ? chart.series[0].cats : Array.from({ length: n }, (_, i) => String(i + 1));
  const leg = legend(chart.series.map((s) => s.name), b, chart);
  const x0 = 50, y0 = leg.y - 22, w = Math.max(20, b.w - x0 - 18), h = Math.max(20, y0 - 34);
  const values = chart.series.flatMap((s) => s.vals.filter((v): v is number => v !== null));
  const minValue = Math.min(0, ...values), maxValue = Math.max(0, ...values);
  const min = minValue < 0 ? -niceMax(-minValue) : 0, max = maxValue > 0 ? niceMax(maxValue) : (min === 0 ? 1 : 0);
  const fraction = (v: number): number => (v - min) / (max - min);
  const horizontal = chart.kind === "bar" && chart.barDir === "bar";
  const band = (horizontal ? h : w) / n;
  const color = (i: number): string => `#${paletteOf(chart)[i % paletteOf(chart).length]}`;
  let svg = frame(chart, b);
  for (let i = 0; i <= 4; i++) {
    const t = i / 4, v = min + (max - min) * t;
    const x = x0 + t * w, y = y0 - t * h;
    const label = Number(v.toPrecision(5));
    svg += horizontal
      ? `<line x1="${x}" y1="34" x2="${x}" y2="${y0}" stroke="#eef1f4"/><text x="${x}" y="${y0 + 14}" text-anchor="middle" font-size="10" fill="#666">${label}</text>`
      : `<line x1="${x0}" y1="${y}" x2="${x0 + w}" y2="${y}" stroke="#eef1f4"/><text x="${x0 - 5}" y="${y + 4}" text-anchor="end" font-size="10" fill="#666">${label}</text>`;
  }
  const zero = fraction(0);
  svg += horizontal ? `<line x1="${x0 + zero * w}" y1="34" x2="${x0 + zero * w}" y2="${y0}" stroke="#999"/>` : `<line x1="${x0}" y1="${y0 - zero * h}" x2="${x0 + w}" y2="${y0 - zero * h}" stroke="#999"/>`;
  cats.forEach((cat, i) => {
    svg += horizontal
      ? `<text x="${x0 - 5}" y="${34 + band * (i + .5) + 4}" text-anchor="end" font-size="10" fill="#666">${esc(cat.slice(0, 8))}</text>`
      : `<text x="${x0 + band * (i + .5)}" y="${y0 + 14}" text-anchor="middle" font-size="10" fill="#666">${esc(cat.slice(0, 8))}</text>`;
  });
  chart.series.forEach((s, si) => {
    if (chart.kind === "bar") {
      const sw = band * .72 / chart.series.length;
      s.vals.forEach((v, i) => {
        if (v === null) return;
        const start = Math.min(zero, fraction(v)), size = Math.abs(fraction(v) - zero);
        svg += horizontal
          ? `<rect x="${x0 + start * w}" y="${34 + band * (i + .14) + sw * si}" width="${size * w}" height="${Math.max(.5, sw - 1)}" fill="${color(si)}"/>`
          : `<rect x="${x0 + band * (i + .14) + sw * si}" y="${y0 - Math.max(zero, fraction(v)) * h}" width="${Math.max(.5, sw - 1)}" height="${size * h}" fill="${color(si)}"/>`;
      });
    } else {
      let pts: string[] = [];
      const flush = (): void => {
        if (!pts.length) return;
        if (chart.kind === "area") {
          const firstX = pts[0].split(",")[0], lastX = pts[pts.length - 1].split(",")[0];
          svg += `<polygon points="${firstX},${y0 - zero * h} ${pts.join(" ")} ${lastX},${y0 - zero * h}" fill="${color(si)}" opacity=".4"/>`;
        } else svg += `<polyline points="${pts.join(" ")}" fill="none" stroke="${color(si)}" stroke-width="2"/>`;
        pts = [];
      };
      s.vals.forEach((v, i) => {
        if (v === null) { flush(); return; }
        const x = x0 + band * (i + .5), y = y0 - fraction(v) * h;
        pts.push(`${x},${y}`);
        if (chart.kind !== "area") svg += `<circle cx="${x}" cy="${y}" r="2.6" fill="${color(si)}"/>`;
      });
      flush();
    }
  });
  return svg + leg.svg;
}

function pieSvg(chart: ChartModel, b: Box): string {
  const s = chart.series[0];
  const items = s.vals.map((v, i) => ({ v: v ?? 0, cat: s.cats[i] ?? "", i })).filter((x) => x.v > 0);
  const leg = legend(items.map((x) => x.cat), b, chart);
  const cx = b.w / 2 - 30;
  const cy = (leg.y + (chart.title ? 26 : 8)) / 2;
  const r = Math.max(16, Math.min(b.w / 2 - 60, (leg.y - (chart.title ? 30 : 12)) / 2));
  const total = items.reduce((a, x) => a + x.v, 0) || 1;
  let angle = -Math.PI / 2;
  let svg = frame(chart, b);
  const isDoughnut = chart.kind === "doughnut";
  items.forEach((x, idx) => {
    const sweep = (x.v / total) * Math.PI * 2;
    const x1 = cx + r * Math.cos(angle);
    const y1 = cy + r * Math.sin(angle);
    angle += sweep;
    const x2 = cx + r * Math.cos(angle);
    const y2 = cy + r * Math.sin(angle);
    const large = sweep > Math.PI ? 1 : 0;
    const fill = `#${paletteOf(chart)[idx % paletteOf(chart).length]}`;
    if (items.length === 1) {
      svg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${fill}"/>`;
      if (isDoughnut) svg += `<circle cx="${cx}" cy="${cy}" r="${r * .55}" fill="#fff"/>`;
    } else if (isDoughnut) {
      const r2 = r * 0.55;
      const ix2 = cx + r2 * Math.cos(angle);
      const iy2 = cy + r2 * Math.sin(angle);
      const ix1 = cx + r2 * Math.cos(angle - sweep);
      const iy1 = cy + r2 * Math.sin(angle - sweep);
      svg += `<path d="M${x1},${y1} A${r},${r} 0 ${large} 1 ${x2},${y2} L${ix2},${iy2} A${r2},${r2} 0 ${large} 0 ${ix1},${iy1} Z" fill="${fill}"/>`;
    } else {
      svg += `<path d="M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${large} 1 ${x2},${y2} Z" fill="${fill}"/>`;
    }
    const mid = angle - sweep / 2;
    svg += `<text x="${cx + r * 0.68 * Math.cos(mid)}" y="${cy + r * 0.68 * Math.sin(mid) + 3.5}" text-anchor="middle" font-size="10" fill="#fff">${Math.round((x.v / total) * 100)}%</text>`;
  });
  svg += leg.svg;
  return svg;
}

function scatterSvg(chart: ChartModel, b: Box): string {
  const leg = legend(chart.series.map((s) => s.name), b, chart);
  const points = chart.series.flatMap((s, si) => s.vals.flatMap((y, i) => {
    const raw = s.cats[i];
    const x = raw === undefined ? i + 1 : raw.trim() ? Number(raw) : NaN;
    return y !== null && Number.isFinite(x) ? [{ x, y, si }] : [];
  }));
  const minX = Math.min(0, ...points.map((p) => p.x)), maxX = Math.max(1, ...points.map((p) => p.x));
  const minY = Math.min(0, ...points.map((p) => p.y)), maxY = Math.max(1, ...points.map((p) => p.y));
  const x0 = 50, y0 = leg.y - 22, w = b.w - x0 - 18, h = y0 - 34;
  let svg = frame(chart, b);
  for (let i = 0; i <= 4; i++) {
    const t = i / 4;
    svg += `<line x1="${x0}" y1="${y0 - h * t}" x2="${x0 + w}" y2="${y0 - h * t}" stroke="#eef1f4"/><text x="${x0 - 5}" y="${y0 - h * t + 4}" text-anchor="end" font-size="10">${Number((minY + (maxY - minY) * t).toPrecision(4))}</text><text x="${x0 + w * t}" y="${y0 + 14}" text-anchor="middle" font-size="10">${Number((minX + (maxX - minX) * t).toPrecision(4))}</text>`;
  }
  for (const p of points) svg += `<circle cx="${x0 + (p.x - minX) / (maxX - minX) * w}" cy="${y0 - (p.y - minY) / (maxY - minY) * h}" r="3.4" fill="#${paletteOf(chart)[p.si % paletteOf(chart).length]}"/>`;
  return svg + leg.svg;
}

// ------------------------------------------------------------- chart part XML

const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
// Shared child order for the supported series/axes; absent optional fields keep their schema position.
const SERIES_ORDER = ["idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "marker", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "smooth", "extLst"];
const AXIS_ORDER = ["axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "crossBetween", "majorUnit", "minorUnit", "extLst"];


const numRef = (f: string, vals: (number | null)[], startIdx: number): string =>
  `<c:numRef><c:f>${esc(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals
    .map((v, i) => (v === null ? "" : `<c:pt idx="${startIdx + i}"><c:v>${v}</c:v></c:pt>`))
    .join("")}</c:numCache></c:numRef>`;

const strRef = (f: string, vals: string[], startIdx: number): string =>
  `<c:strRef><c:f>${esc(f)}</c:f><c:strCache><c:ptCount val="${vals.length}"/>${vals
    .map((v, i) => `<c:pt idx="${startIdx + i}"><c:v>${esc(v)}</c:v></c:pt>`)
    .join("")}</c:strCache></c:strRef>`;

/** The chart part a new chart saves as — the same shape openpyxl and WPS read. */
export function chartPartXml(chart: ChartModel, _sheetName: string): string {
  const scatter = chart.kind === "scatter";
  const pie = chart.kind === "pie" || chart.kind === "doughnut";
  const fill = (color: string): string => `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
  const series = chart.series.map((s, i) => {
    const color = paletteOf(chart)[i % paletteOf(chart).length];
    const cat = s.catRef
      ? (scatter || s.catNumeric ? numRef(s.catRef, s.cats.map((x) => x.trim() && Number.isFinite(Number(x)) ? Number(x) : null), 0) : strRef(s.catRef, s.cats, 0))
      : (scatter ? `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${s.vals.length}"/>${s.vals.map((_, j) => `<c:pt idx="${j}"><c:v>${j + 1}</c:v></c:pt>`).join("")}</c:numLit>` : `<c:strLit><c:ptCount val="${s.cats.length}"/>${s.cats.map((x, j) => `<c:pt idx="${j}"><c:v>${esc(x)}</c:v></c:pt>`).join("")}</c:strLit>`);
    const vals = s.valRef ? numRef(s.valRef, s.vals, 0) : `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${s.vals.length}"/>${s.vals.map((x, j) => x === null ? "" : `<c:pt idx="${j}"><c:v>${x}</c:v></c:pt>`).join("")}</c:numLit>`;
    const tx = s.nameRef ? strRef(s.nameRef, [s.name], 0) : `<c:v>${esc(s.name)}</c:v>`;
    const points = pie ? s.vals.map((_, j) => `<c:dPt><c:idx val="${j}"/><c:spPr>${fill(paletteOf(chart)[j % paletteOf(chart).length])}</c:spPr></c:dPt>`).join("") : "";
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${tx}</c:tx><c:spPr>${fill(color)}<a:ln>${fill(color)}</a:ln></c:spPr>${points}` +
      (s.cats.length || scatter ? `<c:${scatter ? "xVal" : "cat"}>${cat}</c:${scatter ? "xVal" : "cat"}>` : "") +
      `<c:${scatter ? "yVal" : "val"}>${vals}</c:${scatter ? "yVal" : "val"}></c:ser>`;
  }).join("");
  const tag = { bar: "barChart", line: "lineChart", area: "areaChart", pie: "pieChart", doughnut: "doughnutChart", scatter: "scatterChart", unknown: "barChart" }[chart.kind];
  const extra = chart.kind === "bar" ? `<c:barDir val="${chart.barDir ?? "col"}"/><c:grouping val="clustered"/>`
    : chart.kind === "line" || chart.kind === "area" ? `<c:grouping val="standard"/>`
    : scatter ? `<c:scatterStyle val="marker"/>` : "";
  const group = `<c:${tag}>${extra}<c:varyColors val="${pie ? 1 : 0}"/>${series}${chart.kind === "doughnut" ? '<c:holeSize val="55"/>' : ""}${pie ? "" : '<c:axId val="111111111"/><c:axId val="222222222"/>'}</c:${tag}>`;
  const horizontal = chart.kind === "bar" && chart.barDir === "bar";
  const axis = (id: number, cross: number, pos: string, kind: string): string => `<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/></c:${kind}>`;
  const axes = pie ? "" : axis(111111111, 222222222, horizontal ? "l" : "b", scatter ? "valAx" : "catAx") + axis(222222222, 111111111, horizontal ? "b" : "l", "valAx");
  const title = chart.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(chart.title)}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = chart.showLegend === false ? "" : '<c:legend><c:legendPos val="b"/><c:layout/></c:legend>';
  const generated = `${DECL}<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>${title}<c:plotArea><c:layout/>${group}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
  return chart.sourceXml ? patchChart(chart, generated) : generated;
}

/** Patch only owned properties; unrelated chart/space XML and drawing anchors survive. */
function patchChart(model: ChartModel, generated: string): string {
  const tree = parseOrderedXml(model.sourceXml!);
  const fresh = parseOrderedXml(generated);
  const root = rawFind(tree, "chartSpace");
  const chart = rawFind(rawChildren(root), "chart");
  const plot = rawFind(rawChildren(chart), "plotArea");
  const newChart = rawFind(fresh, "chart");
  const newPlot = rawFind(fresh, "plotArea");
  const group = rawChildren(plot).find((x) => /Chart$/.test(tagLocalOf(x) ?? ""));
  const newGroup = rawChildren(newPlot).find((x) => /Chart$/.test(tagLocalOf(x) ?? ""));
  if (!root || !chart || !plot || !group || !newGroup) throw new Error("无法安全编辑此图表");
  Object.assign(rawAttrs(root), { "@_xmlns:c": C_NS, "@_xmlns:a": A_NS });
  const replace = (parent: any, locals: string[], nodes: any[], before?: string[]): void => {
    const children = rawChildren(parent);
    const first = children.findIndex((x) => locals.includes(tagLocalOf(x) ?? ""));
    const rest = children.filter((x) => !locals.includes(tagLocalOf(x) ?? ""));
    const boundary = before ? rest.findIndex((x) => before.includes(tagLocalOf(x) ?? "")) : -1;
    rest.splice(first >= 0 ? Math.min(first, rest.length) : boundary >= 0 ? boundary : rest.length, 0, ...nodes);
    children.splice(0, children.length, ...rest);
  };
  const fields = model.editFields ?? [];
  if (fields.includes("title")) replace(chart, ["title", "autoTitleDeleted"], rawChildren(newChart).filter((x) => ["title", "autoTitleDeleted"].includes(tagLocalOf(x) ?? "")), ["plotArea"]);
  if (fields.includes("legend")) replace(chart, ["legend"], rawChildren(newChart).filter((x) => tagLocalOf(x) === "legend"), ["plotVisOnly", "dispBlanksAs", "extLst"]);
  if (tagLocalOf(group) !== tagLocalOf(newGroup) || fields.includes("type")) {
    const oldSeries = rawChildren(group).filter((x) => tagLocalOf(x) === "ser");
    const newSeries = rawChildren(newGroup).filter((x) => tagLocalOf(x) === "ser");
    newSeries.forEach((ser, i) => {
      if (!oldSeries[i]) return;
      const compatible = ["spPr", "dPt", "dLbls", "extLst", ...(["bar", "line", "area", "scatter"].includes(model.kind) ? ["trendline", "errBars"] : []), ...(["line", "scatter"].includes(model.kind) ? ["marker", "smooth"] : [])];
      for (const local of compatible) {
        if (local !== "extLst" && fields.includes("colors")) continue;
        const kept = rawChildren(oldSeries[i]).filter((x) => tagLocalOf(x) === local);
        if (kept.length) replace(ser, [local], kept, SERIES_ORDER.slice(SERIES_ORDER.indexOf(local) + 1));
      }
    });
    // Preserve compatible axes, including bounds, number formats and typography.
    const oldAxes = rawChildren(plot).filter((x) => ["catAx", "valAx", "dateAx"].includes(tagLocalOf(x) ?? ""));
    const oldIds = rawChildren(group).filter((x) => tagLocalOf(x) === "axId").map((x) => rawAttrs(x)["@_val"]);
    const newAxes = rawChildren(newPlot).filter((x) => ["catAx", "valAx"].includes(tagLocalOf(x) ?? ""));
    const chosen = newAxes.map((axis, i) => {
      const old = oldAxes.find((x) => rawAttrs(rawFind(rawChildren(x), "axId"))["@_val"] === oldIds[i]);
      if (old && (tagLocalOf(old) === tagLocalOf(axis) || (tagLocalOf(old) === "dateAx" && tagLocalOf(axis) === "catAx"))) return structuredClone(old);
      if (old && tagLocalOf(old) === "dateAx") {
        for (const local of ["scaling", "numFmt", "title", "spPr", "txPr", "majorGridlines", "minorGridlines"]) {
          const kept = rawChildren(old).filter((x) => tagLocalOf(x) === local);
          if (kept.length) replace(axis, [local], kept, AXIS_ORDER.slice(AXIS_ORDER.indexOf(local) + 1));
        }
      }
      return axis;
    });
    const ids = chosen.map((x) => rawAttrs(rawFind(rawChildren(x), "axId"))["@_val"]);
    const wasHorizontal = rawAttrs(rawChildren(group).find((x) => tagLocalOf(x) === "barDir") ?? {})["@_val"] === "bar";
    const horizontal = model.kind === "bar" && model.barDir === "bar";
    chosen.forEach((axis, i) => {
      rawAttrs(rawFind(rawChildren(axis), "crossAx"))["@_val"] = ids[1 - i];
      const pos = rawFind(rawChildren(axis), "axPos");
      // A preserved axis rotates only when changing between horizontal/vertical bars.
      if (axis !== newAxes[i] && wasHorizontal !== horizontal) {
        const oldPos = rawAttrs(pos)["@_val"];
        rawAttrs(pos)["@_val"] = ({ b: "l", l: "b", t: "r", r: "t" } as Record<string, string>)[oldPos] ?? oldPos;
      }
    });
    rawChildren(newGroup).filter((x) => tagLocalOf(x) === "axId").forEach((x, i) => { rawAttrs(x)["@_val"] = ids[i]; });
    replace(plot, [tagLocalOf(group)!, "catAx", "valAx", "dateAx", "serAx"], [newGroup, ...chosen], ["dTable", "spPr", "extLst"]);
  } else {
    const oldSeries = rawChildren(group).filter((x) => tagLocalOf(x) === "ser");
    const newSeries = rawChildren(newGroup).filter((x) => tagLocalOf(x) === "ser");
    const merged = newSeries.map((freshSer, i) => {
      const ser = oldSeries[i];
      if (!ser) return freshSer;
      // Refresh references/caches while retaining markers, labels, trendlines and extensions.
      for (const local of ["tx", "cat", "val", "xVal", "yVal"]) {
        const nodes = rawChildren(freshSer).filter((x) => tagLocalOf(x) === local);
        if (nodes.length || ((local === "cat" || local === "xVal") && fields.includes("source"))) replace(ser, [local], nodes, SERIES_ORDER.slice(SERIES_ORDER.indexOf(local) + 1));
      }
      if (fields.includes("colors")) {
        const generatedProps = rawFind(rawChildren(freshSer), "spPr");
        const oldProps = rawChildren(ser).find((x) => tagLocalOf(x) === "spPr");
        const fills = ["solidFill", "noFill", "gradFill", "pattFill", "blipFill", "grpFill"];
        if (oldProps) {
          replace(oldProps, fills, rawChildren(generatedProps).filter((x) => tagLocalOf(x) === "solidFill"), ["ln", "effectLst", "extLst"]);
          const ln = rawChildren(oldProps).find((x) => tagLocalOf(x) === "ln");
          const newLn = rawChildren(generatedProps).find((x) => tagLocalOf(x) === "ln");
          if (ln) replace(ln, fills, rawChildren(newLn));
          else rawChildren(oldProps).push(newLn);
        } else replace(ser, ["spPr"], [generatedProps], ["marker", "dPt", "dLbls", "cat", "val", "xVal", "yVal", "extLst"]);
        // Pie palettes own each point's fill, never its explosion/label/extension data.
        if (model.kind === "pie" || model.kind === "doughnut") {
          const oldPoints = rawChildren(ser).filter((x) => tagLocalOf(x) === "dPt");
          const newPoints = rawChildren(freshSer).filter((x) => tagLocalOf(x) === "dPt");
          const points = [...oldPoints];
          for (const point of newPoints) {
            const idx = rawAttrs(rawFind(rawChildren(point), "idx"))["@_val"];
            const old = oldPoints.find((x) => rawAttrs(rawFind(rawChildren(x), "idx"))["@_val"] === idx);
            if (!old) { points.push(point); continue; }
            const spPr = rawChildren(old).find((x) => tagLocalOf(x) === "spPr");
            const freshProps = rawFind(rawChildren(point), "spPr");
            if (spPr) replace(spPr, fills, rawChildren(freshProps), ["ln", "effectLst", "extLst"]);
            else replace(old, ["spPr"], [freshProps], ["pictureOptions", "extLst"]);
          }
          replace(ser, ["dPt"], points, ["dLbls", "cat", "val", "xVal", "yVal", "extLst"]);
        }
      }
      return ser;
    });
    replace(group, ["ser"], merged, ["dLbls", "gapWidth", "overlap", "axId", "extLst"]);
  }

  return DECL + buildXml(tree.filter((x) => tagOf(x) !== "?xml"));
}

/** Complex chart structures remain opaque, but their local A1 references still follow cells. */
export function rewriteChartSourceReferences(xml: string, rewrite: (ref: string) => string): string | undefined {
  const tree = parseOrderedXml(xml);
  const root = rawFind(tree, "chartSpace");
  if (!root) return;
  const namespaces = rawAttrs(root);
  const prefixes = Object.entries(namespaces).filter(([key, value]) => key.startsWith("@_xmlns") && value === C_NS).map(([key]) => key === "@_xmlns" ? "" : key.slice("@_xmlns:".length) + ":");
  let changed = false;
  for (const local of ["numRef", "strRef", "multiLvlStrRef"]) for (const ref of rawFindAll(tree, local)) {
    if (!prefixes.some((p) => tagOf(ref) === p + local)) continue;
    const f = rawChildren(ref).find((x) => prefixes.some((p) => tagOf(x) === p + "f"));
    if (!f) continue;
    const old = rawChildren(f).map((x) => String(x["#text"] ?? "")).join("");
    const next = rewrite(old);
    if (next === old) continue;
    changed = true;
    rawChildren(f).splice(0, rawChildren(f).length, { "#text": next });
    // Optional caches are invalid after a reference shift; consumers recalculate from f.
    const children = rawChildren(ref);
    const kept = children.filter((x) => !["numCache", "strCache", "multiLvlStrCache"].includes(tagLocalOf(x) ?? ""));
    children.splice(0, children.length, ...kept);
  }
  return changed ? DECL + buildXml(tree.filter((x) => tagOf(x) !== "?xml")) : undefined;
}
