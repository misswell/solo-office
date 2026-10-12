import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { strToU8, zipSync } from "fflate";
import { Spreadsheet, drawingAnchorAtPx, imageBoxPx, type ChartModel } from "../src/core/xlsx.js";
import { writeXlsx } from "../src/core/write.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { rect } from "../src/core/sheet-ops.js";
import { History } from "../src/core/history.js";
import { chartRange } from "../src/core/chart-data.js";
import { CHART_PALETTES, renderChartSvg } from "../src/core/chart.js";

const fixture = (): Uint8Array => new Uint8Array(readFileSync(new URL("./fixtures/chart-sample.xlsx", import.meta.url)));
function sample(): Spreadsheet {
  const wb = Spreadsheet.blank();
  for (const [r, row] of [["X", "收入", "支出"], ["1", "12", "8"], ["3", "-4", "15"], ["10", "20", "5"]].entries()) row.forEach((v, c) => wb.setCell(0, `${String.fromCharCode(65 + c)}${r + 1}`, v));
  wb.addChart("bar", 0, rect(0, 0, 2, 3), { c: 4, r: 1 }, "原图表");
  return wb;
}
const edit = (wb: Spreadsheet, patch: Partial<Parameters<Spreadsheet["prepareChartEdit"]>[2]> = {}): void => {
  const old = wb.sheets[0].charts[0];
  wb.updateChart(0, 0, wb.prepareChartEdit(0, 0, { kind: old.kind, title: old.title, barDir: old.barDir, showLegend: old.showLegend !== false, ...patch }));
};
function customize(wb: Spreadsheet, transform: (xml: string) => string): Spreadsheet {
  const pkg = OpcPackage.from(writeXlsx(wb));
  const path = "xl/charts/chart1.xml";
  pkg.parts.set(path, strToU8(transform(text(pkg.bytes(path)!))));
  return Spreadsheet.openBytes(zipSync(Object.fromEntries(pkg.parts)));
}

describe("chart editing through workbook save/open", () => {
  it.each(["bar", "line", "area", "pie", "doughnut", "scatter"] as ChartModel["kind"][])("converts an imported chart to %s without moving its drawing", (kind) => {
    const bytes = writeXlsx(sample()), wb = Spreadsheet.openBytes(bytes);
    edit(wb, { kind, title: "新标题 & <测试>", colors: CHART_PALETTES.ocean, showLegend: false });
    const pkg = OpcPackage.from(writeXlsx(wb)), old = OpcPackage.from(bytes);
    const back = Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].charts[0];
    expect(back.kind).toBe(kind);
    expect(back.title).toBe("新标题 & <测试>");
    expect(back.series[0].vals).toEqual([12, -4, 20]);
    expect(back.showLegend).toBe(false);
    expect(back.series[0].valRef).toBe("'Sheet1'!$B$2:$B$4");
    expect(pkg.bytes("xl/drawings/drawing1.xml")).toEqual(old.bytes("xl/drawings/drawing1.xml"));
    const xml = pkg.xml("xl/charts/chart1.xml");
    if (kind === "pie" || kind === "doughnut") expect(xml.chartSpace.chart.plotArea.valAx).toBeUndefined();
    if (kind === "scatter") expect(xml.chartSpace.chart.plotArea.scatterChart.ser[0].xVal.numRef.f["#text"]).toBe("'Sheet1'!$A$2:$A$4");
  });

  it("updates title/colors in an original openpyxl part, leaving other charts byte-identical", () => {
    const bytes = fixture(), wb = Spreadsheet.openBytes(bytes);
    edit(wb, { title: "修改标题", colors: CHART_PALETTES.warm });
    const pkg = OpcPackage.from(writeXlsx(wb)), old = OpcPackage.from(bytes);
    expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].charts[0].title).toBe("修改标题");
    expect(pkg.bytes("xl/charts/chart2.xml")).toEqual(old.bytes("xl/charts/chart2.xml"));
    expect(pkg.bytes("xl/charts/chart3.xml")).toEqual(old.bytes("xl/charts/chart3.xml"));
    expect(text(pkg.bytes("xl/charts/chart1.xml")!)).toContain("C8553D");
  });

  it("preserves unowned extensions and series attributes on title-only edits", () => {
    const bytes = writeXlsx(sample()), pkg = OpcPackage.from(bytes);
    const path = "xl/charts/chart1.xml";
    const xml = text(pkg.bytes(path)!).replace("</c:ser>", '<c:extLst><c:ext uri="test"><x:payload xmlns:x="urn:solo-test" flag="keep"/></c:ext></c:extLst></c:ser>').replace("</c:chartSpace>", '<c:printSettings><c:pageSetup orientation="landscape"/></c:printSettings></c:chartSpace>');
    pkg.parts.set(path, strToU8(xml));
    const wb = Spreadsheet.openBytes(zipSync(Object.fromEntries(pkg.parts)));
    edit(wb, { title: "Only title", source: "A1:C3" });
    const saved = text(OpcPackage.from(writeXlsx(wb)).bytes(path)!);
    expect(saved).toContain('flag="keep"');
    expect(saved).toContain('orientation="landscape"');
  });

  it("changes cross-sheet sources with quoted apostrophes and updates formula-derived caches", () => {
    const wb = sample();
    const data = wb.addSheet("O'Brien ! 数据");
    for (const [ref, value] of [["A1", "类别"], ["B1", "新系列"], ["A2", "甲"], ["A3", "乙"], ["B2", "=2*6"], ["B3", "30"]]) wb.setCell(data.name, ref, value);
    edit(wb, { source: "'O''Brien ! 数据'!A1:B3" });
    wb.setCell(data.name, "B2", "=7*6");
    const display = wb.chartForDisplay("Sheet1", wb.sheets[0].charts[0]);
    expect(display.series[0].vals).toEqual([42, 30]);
    const back = Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].charts[0];
    expect(back.series[0].name).toBe("新系列");
    expect(back.series[0].cats).toEqual(["甲", "乙"]);
    expect(back.series[0].vals).toEqual([42, 30]);
    expect(back.series[0].valRef).toBe("'O''Brien ! 数据'!$B$2:$B$3");
  });

  it("refreshes imported chart caches on ordinary cell edits", () => {
    const wb = Spreadsheet.openBytes(writeXlsx(sample()));
    wb.setCell(0, "B2", "88");
    const pkg = OpcPackage.from(writeXlsx(wb));
    const chart = pkg.xml("xl/charts/chart1.xml").chartSpace.chart.plotArea.barChart;
    expect(chart.ser[0].val.numRef.numCache.pt[0].v["#text"]).toBe("88");
    expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].charts[0].series[0].vals[0]).toBe(88);
  });

  it("validates malformed, missing and oversized references atomically", () => {
    const wb = sample(), before = wb.snapshot(), rev = wb.revision;
    for (const source of ["A1:B3:C4", "missing!A1:B3", "XFE1:XFE4", "A1:B100001", "A1:B1"]) expect(() => edit(wb, { source })).toThrow();
    expect(wb.snapshot()).toEqual(before);
    expect(wb.revision).toBe(rev);
    expect(chartRange("'O''Brien ! 数据'!$A$1:$B$3", "Sheet1").sheet).toBe("O'Brien ! 数据");
  });

  it("follows source cells through row insertion, deletion and quoted sheet renames", () => {
    const wb = Spreadsheet.openBytes(writeXlsx(sample()));
    wb.insertRows(0, 1);
    let chart = wb.sheets[0].charts[0];
    expect(chart.series[0].valRef).toBe("'Sheet1'!$B$3:$B$5");
    expect(wb.chartForDisplay("Sheet1", chart).series[0].vals).toEqual([12, -4, 20]);
    wb.deleteRows(0, 1);
    wb.renameSheet(0, "O'Brien ! 数据");
    chart = wb.sheets[0].charts[0];
    expect(chart.series[0].valRef).toBe("'O''Brien ! 数据'!$B$2:$B$4");
    const back = Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].charts[0];
    expect(back.series[0].vals).toEqual([12, -4, 20]);
    expect(back.series[0].valRef).toBe(chart.series[0].valRef);
  });

  it("undoes and redoes chart edits and isolates nested snapshot data", () => {
    const wb = sample(), history = new History();
    history.record(wb);
    edit(wb, { kind: "doughnut", title: "改后", colors: CHART_PALETTES.ocean });
    expect(history.undoStep(wb)).toBe(true);
    expect(wb.sheets[0].charts[0].title).toBe("原图表");
    expect(history.redoStep(wb)).toBe(true);
    expect(wb.sheets[0].charts[0].kind).toBe("doughnut");
    const snap = wb.snapshot();
    wb.sheets[0].charts[0].series[0].vals[0] = 999;
    wb.sheets[0].charts[0].colors![0] = "FFFFFF";
    expect(snap.sheets[0].charts[0].series[0].vals[0]).toBe(12);
    expect(snap.sheets[0].charts[0].colors![0]).toBe("087F8C");
  });

  it("renders horizontal, negative and numeric X values without invalid SVG geometry", () => {
    const wb = sample();
    edit(wb, { barDir: "bar" });
    const svg = renderChartSvg(wb.sheets[0].charts[0], 500, 300);
    expect(svg).not.toMatch(/(?:width|height)="-/);
    edit(wb, { kind: "scatter" });
    expect(renderChartSvg(wb.sheets[0].charts[0], 500, 300)).not.toMatch(/NaN|Infinity/);
  });

  it("retains stacked chart bytes and refuses lossy editing", () => {
    const pkg = OpcPackage.from(writeXlsx(sample()));
    const path = "xl/charts/chart1.xml";
    pkg.parts.set(path, strToU8(text(pkg.bytes(path)!).replace('grouping val="clustered"', 'grouping val="stacked"')));
    const bytes = zipSync(Object.fromEntries(pkg.parts)), wb = Spreadsheet.openBytes(bytes);
    expect(() => edit(wb, { title: "不可写" })).toThrow(/堆积/);
    wb.setCell(0, "Z1", "其他编辑");
    expect(OpcPackage.from(writeXlsx(wb)).bytes(path)).toEqual(pkg.bytes(path));
  });

  it("preserves compatible axis bounds and number formats through type/direction/title edits", () => {
    const wb = customize(sample(), (xml) => xml.replace('<c:crossAx val="111111111"/>', '<c:numFmt formatCode="0.00%" sourceLinked="0"/><c:crossAx val="111111111"/>').replaceAll('</c:scaling>', '<c:max val="100"/><c:min val="-10"/></c:scaling>'));
    edit(wb, { barDir: "bar" });
    edit(wb, { title: "旋转后仍保留轴设置" });
    let plot = OpcPackage.from(writeXlsx(wb)).xml("xl/charts/chart1.xml").chartSpace.chart.plotArea;
    expect(plot.valAx.scaling.max["@_val"]).toBe("100");
    expect(plot.valAx.numFmt["@_formatCode"]).toBe("0.00%");
    expect(plot.valAx.axPos["@_val"]).toBe("b");
    edit(wb, { kind: "line" });
    plot = OpcPackage.from(writeXlsx(wb)).xml("xl/charts/chart1.xml").chartSpace.chart.plotArea;
    expect(plot.valAx.scaling.min["@_val"]).toBe("-10");
    expect(plot.valAx.numFmt["@_formatCode"]).toBe("0.00%");
    expect(plot.valAx.axPos["@_val"]).toBe("l");
  });

  it("keeps point explosion/extension data when changing pie palettes", () => {
    const original = sample();
    edit(original, { kind: "pie" });
    const wb = customize(original, (xml) => xml.replace('<c:idx val="0"/><c:spPr>', '<c:idx val="0"/><c:explosion val="25"/><c:spPr>').replace('</c:dPt>', '<c:extLst><c:ext uri="point-extra"/></c:extLst></c:dPt>'));
    edit(wb, { colors: CHART_PALETTES.warm });
    const point = OpcPackage.from(writeXlsx(wb)).xml("xl/charts/chart1.xml").chartSpace.chart.plotArea.pieChart.ser[0].dPt[0];
    expect(point.explosion["@_val"]).toBe("25");
    expect(point.extLst.ext["@_uri"]).toBe("point-extra");
    expect(point.spPr.solidFill.srgbClr["@_val"]).toBe("C8553D");
  });

  it("keeps existing series extensions when the source adds another series", () => {
    const wb = customize(sample(), (xml) => xml.replace('</c:ser>', '<c:extLst><c:ext uri="first-series-extra"/></c:extLst></c:ser>'));
    wb.setCell(0, "D1", "新增"); wb.setCell(0, "D2", "4"); wb.setCell(0, "D3", "6"); wb.setCell(0, "D4", "8");
    edit(wb, { source: "A1:D4" });
    const series = OpcPackage.from(writeXlsx(wb)).xml("xl/charts/chart1.xml").chartSpace.chart.plotArea.barChart.ser;
    expect(series).toHaveLength(3);
    expect(series[0].extLst.ext["@_uri"]).toBe("first-series-extra");
    expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].charts[0].series[2].vals).toEqual([4, 6, 8]);
  });

  it("preserves data labels while changing source and chart type together", () => {
    const wb = customize(sample(), (xml) => xml.replace('<c:cat>', '<c:dLbls><c:showVal val="1"/></c:dLbls><c:cat>'));
    edit(wb, { kind: "line", source: "A1:B3" });
    const series = OpcPackage.from(writeXlsx(wb)).xml("xl/charts/chart1.xml").chartSpace.chart.plotArea.lineChart.ser;
    expect(series.dLbls.showVal["@_val"]).toBe("1");
    expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].charts[0].series[0].vals).toEqual([12, -4]);
  });

  it("keeps a date axis when converting between category chart types", () => {
    const wb = customize(sample(), (xml) => xml.replaceAll("c:catAx", "c:dateAx"));
    edit(wb, { kind: "line" });
    const plot = OpcPackage.from(writeXlsx(wb)).xml("xl/charts/chart1.xml").chartSpace.chart.plotArea;
    expect(plot.dateAx.axId["@_val"]).toBe("111111111");
    expect(plot.catAx).toBeUndefined();
  });

  it("rewrites references in opaque stacked charts after a sheet rename without flattening them", () => {
    const wb = customize(sample(), (xml) => xml.replace('grouping val="clustered"', 'grouping val="stacked"').replace('</c:chartSpace>', '<c:extLst><c:ext uri="opaque-extra"/></c:extLst></c:chartSpace>'));
    wb.renameSheet(0, "新数据");
    const xml = text(OpcPackage.from(writeXlsx(wb)).bytes("xl/charts/chart1.xml")!);
    expect(xml).toContain('grouping val="stacked"');
    expect(xml).toContain('uri="opaque-extra"');
    const parsed = OpcPackage.from(writeXlsx(wb)).xml("xl/charts/chart1.xml").chartSpace.chart.plotArea.barChart;
    expect(parsed.ser[0].val.numRef.f["#text"]).toBe("'新数据'!$B$2:$B$4");
    expect(xml).not.toContain("Sheet1");
    expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].charts[0].series[0].vals).toEqual([12, -4, 20]);
  });

  it("moves/resizes an imported chart without changing chart content or frame identity", () => {
    const bytes = writeXlsx(sample()), wb = Spreadsheet.openBytes(bytes), original = OpcPackage.from(bytes);
    wb.setChartBox(0, 0, { x: 123, y: 47, w: 640, h: 380 });
    const pkg = OpcPackage.from(writeXlsx(wb)), back = Spreadsheet.openBytes(writeXlsx(wb));
    expect(imageBoxPx(back.sheets[0], back.sheets[0].charts[0])).toEqual({ x: 123, y: 47, w: 640, h: 380 });
    expect(pkg.bytes("xl/charts/chart1.xml")).toEqual(original.bytes("xl/charts/chart1.xml"));
    expect(pkg.xml("xl/drawings/drawing1.xml").wsDr.oneCellAnchor.graphicFrame.nvGraphicFramePr.cNvPr["@_name"]).toBe("原图表");
  });

  it("resizes a two-cell anchor while retaining anchoring policy and graphic-frame metadata", () => {
    const pkg = OpcPackage.from(writeXlsx(sample()));
    const path = "xl/drawings/drawing1.xml";
    const xml = text(pkg.bytes(path)!).replaceAll("xdr:oneCellAnchor", "xdr:twoCellAnchor").replace('<xdr:twoCellAnchor>', '<xdr:twoCellAnchor editAs="absolute">').replace(/<xdr:ext[^>]*\/>/, '<xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>').replace('<xdr:cNvGraphicFramePr/>', '<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>');
    pkg.parts.set(path, strToU8(xml));
    const wb = Spreadsheet.openBytes(zipSync(Object.fromEntries(pkg.parts)));
    wb.setChartBox(0, 0, { x: 171, y: 61, w: 420, h: 260 });
    const saved = OpcPackage.from(writeXlsx(wb));
    expect(text(saved.bytes(path)!)).toContain('noGrp="1"');
    expect(saved.xml(path).wsDr.oneCellAnchor).toBeUndefined();
    expect(saved.xml(path).wsDr.twoCellAnchor["@_editAs"]).toBe("absolute");
    const back = Spreadsheet.openBytes(writeXlsx(wb));
    expect(imageBoxPx(back.sheets[0], back.sheets[0].charts[0])).toEqual({ x: 171, y: 61, w: 420, h: 260 });
  });

  it("preserves opaque sibling anchors and default namespaces while moving and deleting modeled charts", () => {
    const source = sample();
    source.addChart("line", 0, rect(0, 0, 2, 3), { c: 10, r: 10 }, "第二图表");
    const pkg = OpcPackage.from(writeXlsx(source)), path = "xl/drawings/drawing1.xml";
    let xml = text(pkg.bytes(path)!);
    const frame = xml.match(/<xdr:graphicFrame>[\s\S]*?<\/xdr:graphicFrame>/)![0];
    const absolute = `<xdr:absoluteAnchor><xdr:pos x="100" y="200"/><xdr:ext cx="952500" cy="952500"/>${frame}<xdr:clientData/></xdr:absoluteAnchor>`;
    const shape = '<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="99" name="Opaque shape"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:blipFill><a:blip r:embed="rIdOpaque"/></a:blipFill></xdr:spPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>';
    xml = xml.replace('</xdr:wsDr>', `${absolute}${shape}</xdr:wsDr>`).replace('<xdr:wsDr ', '<xdr:wsDr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replaceAll('xdr:', '').replace('xmlns:xdr=', 'xmlns=');
    pkg.parts.set(path, strToU8(xml));
    const relPath = "xl/drawings/_rels/drawing1.xml.rels";
    pkg.parts.set(relPath, strToU8(text(pkg.bytes(relPath)!).replace('</Relationships>', '<Relationship Id="rIdOpaque" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/opaque.png"/></Relationships>')));
    pkg.parts.set("xl/media/opaque.png", new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jJ5kAAAAASUVORK5CYII=", "base64")));
    const original = pkg.xml(path).wsDr, wb = Spreadsheet.openBytes(zipSync(Object.fromEntries(pkg.parts)));
    expect(wb.sheets[0].charts).toHaveLength(2);
    wb.setChartBox(0, 0, { x: 123, y: 47, w: 640, h: 380 });
    const saved = OpcPackage.from(writeXlsx(wb)), tree = saved.xml(path).wsDr;
    expect(tree.absoluteAnchor).toEqual(original.absoluteAnchor);
    expect(tree.oneCellAnchor[2]).toEqual(original.oneCellAnchor[2]);
    expect(tree.oneCellAnchor[1]).toEqual(original.oneCellAnchor[1]);
    expect(text(saved.bytes(path)!)).not.toContain('<xdr:');
    expect(saved.bytes("xl/media/opaque.png")).toEqual(pkg.bytes("xl/media/opaque.png"));
    wb.removeChart(0, 1);
    const deleted = OpcPackage.from(writeXlsx(wb)).xml(path).wsDr;
    expect(deleted.oneCellAnchor).toHaveLength(2);
    expect(deleted.oneCellAnchor[1]).toEqual(original.oneCellAnchor[2]);
    expect(deleted.absoluteAnchor).toEqual(original.absoluteAnchor);
  });

  it("moves and deletes the selected anchor when charts share a part", () => {
    const source = sample();
    source.addChart("line", 0, rect(0, 0, 2, 3), { c: 10, r: 10 }, "第二图表");
    const pkg = OpcPackage.from(writeXlsx(source)), path = "xl/drawings/drawing1.xml";
    pkg.parts.set(path, strToU8(text(pkg.bytes(path)!).replace('r:id="rId2"', 'r:id="rId1"')));
    const wb = Spreadsheet.openBytes(zipSync(Object.fromEntries(pkg.parts)));
    expect(wb.sheets[0].charts[0].part).toBe(wb.sheets[0].charts[1].part);
    const first = imageBoxPx(wb.sheets[0], wb.sheets[0].charts[0]);
    wb.setChartBox(0, 1, { x: 171, y: 61, w: 420, h: 260 });
    let back = Spreadsheet.openBytes(writeXlsx(wb));
    expect(imageBoxPx(back.sheets[0], back.sheets[0].charts[0])).toEqual(first);
    expect(imageBoxPx(back.sheets[0], back.sheets[0].charts[1])).toEqual({ x: 171, y: 61, w: 420, h: 260 });
    wb.removeChart(0, 0);
    back = Spreadsheet.openBytes(writeXlsx(wb));
    expect(back.sheets[0].charts).toHaveLength(1);
    expect(imageBoxPx(back.sheets[0], back.sheets[0].charts[0])).toEqual({ x: 171, y: 61, w: 420, h: 260 });
  });

  it("deletes one picture anchor while retaining another that shares its media part", () => {
    const source = Spreadsheet.blank(), png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jJ5kAAAAASUVORK5CYII=";
    source.addImage(0, { c: 1, r: 1 }, png, 40, 40, "第一图");
    source.addImage(0, { c: 5, r: 5 }, png, 60, 60, "第二图");
    const pkg = OpcPackage.from(writeXlsx(source)), path = "xl/drawings/_rels/drawing1.xml.rels";
    pkg.parts.set(path, strToU8(text(pkg.bytes(path)!).replace('../media/image2.png', '../media/image1.png')));
    const wb = Spreadsheet.openBytes(zipSync(Object.fromEntries(pkg.parts)));
    expect(wb.sheets[0].images[0].path).toBe(wb.sheets[0].images[1].path);
    const history = new History(); history.record(wb);
    wb.removeImage(0, 0);
    expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].images.map((i) => i.name)).toEqual(["第二图"]);
    history.undoStep(wb);
    expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].images.map((i) => i.name)).toEqual(["第一图", "第二图"]);
  });

  it("inverts nonuniform and hidden grid dimensions and rejects invalid geometry atomically", () => {
    const wb = sample(), sheet = wb.sheets[0];
    sheet.colWidths.set(0, 20); sheet.hiddenCols.add(1); sheet.rowHeights.set(0, 30);
    const anchor = drawingAnchorAtPx(sheet, 151, 47);
    expect(anchor).toEqual({ col: 2, row: 1, colOff: 7 * 9525, rowOff: 7 * 9525 });
    const rev = wb.revision;
    expect(() => wb.setChartBox(0, 0, { x: NaN, y: 0, w: 100, h: 100 })).toThrow();
    expect(wb.revision).toBe(rev);
    const history = new History(); history.record(wb);
    wb.setChartBox(0, 0, { x: 151, y: 47, w: 600, h: 400 });
    history.undoStep(wb);
    expect(wb.sheets[0].charts[0].cx).toBe(500 * 9525);
  });
});
