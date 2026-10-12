import { describe, expect, it } from "vitest";
import { Spreadsheet } from "../src/core/xlsx.js";
import { Evaluator, parse } from "../src/core/formula.js";
import { writeXlsx } from "../src/core/write.js";
import { History } from "../src/core/history.js";
import { OpcPackage } from "../src/core/opc.js";
import { BUILTIN } from "../src/core/numfmt.js";
import { parseGrid, parseRange, rangeSpec, rect, refOf, shiftRefs } from "../src/core/sheet-ops.js";
import { makeXlsx } from "./helpers/make.js";

const fresh = (): Spreadsheet => Spreadsheet.openBytes(makeXlsx());
const box = (spec: string) => parseRange(spec);

describe("structural editing", () => {
  it("moves cells and rewrites relative and absolute references on insert", () => {
    const wb = fresh();
    wb.insertRows("汇总", 1, 2);
    expect(wb.sheets[0].cells.get("B6")?.formula).toBe("SUM(B4:B5)");
    expect(wb.sheets[0].cells.get("C5")?.formula).toBe("B5/$B$7");
    expect(wb.value("汇总", "B6")).toBe(30800.5);
    expect(wb.value("汇总", "A8")).toBe(42);
  });

  it("shrinks a range that straddles a deleted band and errors one inside it", () => {
    expect(shiftRefs("SUM(B2:B3)", "row", 1, 1)).toBe("SUM(B2:B2)");
    expect(shiftRefs("SUM(B2:B5)", "row", 1, 2)).toBe("SUM(B2:B3)");
    expect(shiftRefs("SUM(A1:B2)", "row", 1, 1)).toBe("SUM(A1:B1)");
    expect(shiftRefs("SUM(B2:B3)", "row", 1, 2)).toBe("SUM(#REF!)");
    expect(shiftRefs("SUM(A1:B9)", "row", 1, 1)).toBe("SUM(A1:B8)");
    expect(shiftRefs("B2/$B$5", "row", 1, 1)).toBe("#REF!/$B$4");
    expect(shiftRefs("'数据2'!B2", "col", 0, -1)).toBe("'数据2'!C2");
    expect(shiftRefs('CONCAT("B2",B3)', "row", 1, 1)).toBe('CONCAT("B2",B2)');
  });

  it("drops the deleted cells, shifts the rest and keeps values computing", () => {
    const wb = fresh();
    wb.deleteRows("汇总", 1, 1);
    expect(wb.value("汇总", "B2")).toBe(18300.5);
    expect(wb.sheets[0].cells.get("B3")?.formula).toBe("SUM(B2:B2)");
    expect(wb.value("汇总", "B3")).toBe(18300.5);
    expect(wb.sheets[0].cells.has("A2")).toBe(true);
    expect(wb.value("汇总", "A2")).toBe("Q2");
    expect(wb.value("汇总", "A5")).toBe(42);
    // a row deleted over here leaves the other sheet's geometry and its refs alone
    expect(wb.sheets[0].cells.get("A5")?.formula).toBe("'数据2'!B2");
    expect(wb.sheets[1].cells.get("B2")?.constant).toBe(42);
    expect(wb.value("数据2", "B2")).toBe(42);
  });

  it("rewrites a reference to the edited sheet that lives on another sheet", () => {
    const wb = fresh();
    wb.setCell("数据2", "C3", "=汇总!B4");
    wb.deleteRows("汇总", 1, 1);
    expect(wb.sheets[1].cells.get("C3")?.formula).toBe("汇总!B3");
    expect(wb.value("数据2", "C3")).toBe(18300.5);
  });

  it("shifts column widths, row heights and merges with the lines they belong to", () => {
    const wb = fresh();
    wb.insertCols("汇总", 0, 1);
    expect(wb.sheets[0].colWidths.get(1)).toBe(18);
    expect(wb.sheets[0].merges).toContain("B1:D1");
    wb.insertRows("汇总", 0, 1);
    expect(wb.sheets[0].merges).toContain("B2:D2");
    expect(wb.sheets[0].freeze).toEqual({ rows: 1, cols: 0 });
    wb.deleteRows("汇总", 0, 2);
    expect(wb.sheets[0].merges).toEqual([]);
  });

  it("keeps a defined name pointing at its cell across a structural edit", () => {
    const wb = fresh();
    wb.insertRows("汇总", 0, 2);
    const ev = new Evaluator(wb, "汇总");
    expect(ev.eval(parse("利润率"), "汇总")).toBe(0.1875);
    expect(wb.value("汇总", "C4")).toBe(0.1875);
  });
});

describe("styles, values and clipboard", () => {
  it("applies a style patch across a range and reports it back", () => {
    const wb = fresh();
    wb.setStyle("汇总", box("A2:B3"), { bold: true, bgColor: "FFC00000", numFmt: "#,##0" });
    expect(wb.styleAt("汇总", 0, 1).bold).toBe(true);
    expect(wb.styleAt("汇总", 1, 2).bgColor).toBe("FFC00000");
    expect(wb.display("汇总", "B2").text).toBe("12,500");
    expect(wb.styleAt("汇总", 2, 2).bold).toBe(false);
  });

  it("formats numbers the way the toolbar would", () => {
    const wb = fresh();
    wb.setStyle("汇总", box("B2:B3"), { numFmt: "0.00%" });
    expect(wb.display("汇总", "B2").text).toBe("1250000.00%");
    wb.setStyle("汇总", box("B2:B3"), { numFmt: "#,##0" });
    expect(wb.display("汇总", "B2").text).toBe("12,500");
    wb.setStyle("汇总", box("B2:B3"), { numFmt: "0" });
    expect(wb.display("汇总", "B2").text).toBe("12500");
  });

  it("clears contents, formats or both", () => {
    const wb = fresh();
    wb.clear("汇总", box("B2:B2"), "content");
    expect(wb.value("汇总", "B2")).toBe(null);
    expect(wb.styleAt("汇总", 1, 1).bold).toBe(false);
    wb.setStyle("汇总", box("B2:B2"), { bold: true });
    wb.clear("汇总", box("B2:B2"), "format");
    expect(wb.styleAt("汇总", 1, 1).bold).toBe(false);
    wb.setCell("汇总", "B2", "7");
    wb.clear("汇总", box("B2:B2"), "all");
    expect(wb.sheets[0].cells.has("B2")).toBe(false);
  });

  it("pastes tab-separated text as typed cells", () => {
    const wb = fresh();
    const area = wb.pasteText("汇总", 0, 9, "甲\t1\t2026-03-01\n乙\t3\tTRUE");
    expect(area).toEqual(rect(0, 9, 2, 10));
    expect(wb.value("汇总", "A10")).toBe("甲");
    expect(wb.value("汇总", "B11")).toBe(3);
    expect(wb.value("汇总", "C11")).toBe(true);
    expect(wb.value("汇总", "C10")).toBe(46082);
    expect(parseGrid('x\t"a\tb"\n')).toEqual([["x", "a\tb"]]);
  });

  it("collects status-bar statistics over a selection", () => {
    const wb = fresh();
    const s = wb.stats("汇总", box("B2:B3"));
    expect(s).toMatchObject({ count: 2, numeric: 2, sum: 30800.5, min: 12500, max: 18300.5 });
    expect(wb.stats("汇总", box("B2:B5"))).toMatchObject({ count: 4, numeric: 4, sum: 77001.25, avg: 19250.3125 });
  });

  it("merges and unmerges blocks without losing the anchor", () => {
    const wb = fresh();
    wb.merge("汇总", box("D1:E2"));
    expect(wb.sheets[0].merges).toContain("D1:E2");
    expect(wb.mergeSpan("汇总", 4, 1)).toEqual(rect(3, 0, 4, 1));
    expect(wb.sheets[0].cells.has("E2")).toBe(false);
    wb.unmerge("汇总", box("D1:E2"));
    expect(wb.sheets[0].merges).toEqual(["A1:C1"]);
  });

  it("sorts rows by one column, blanks last", () => {
    const wb = fresh();
    ["甲", "乙", "丙"].forEach((t, i) => wb.setCell("汇总", `E${i + 2}`, t));
    wb.setCell("汇总", "F2", "30");
    wb.setCell("汇总", "F3", "10");
    wb.setCell("汇总", "F4", "20");
    wb.sort("汇总", box("E2:F5"), 5, false);
    expect(["E2", "E3", "E4", "E5"].map((r) => wb.value("汇总", r))).toEqual(["乙", "丙", "甲", null]);
    expect(["F2", "F3", "F4", "F5"].map((r) => wb.value("汇总", r))).toEqual([10, 20, 30, null]);
    wb.sort("汇总", box("E2:F5"), 5, true);
    expect(wb.value("汇总", "E2")).toBe("甲");
  });

  it("finds and replaces over displayed text", () => {
    const wb = fresh();
    expect(wb.find("汇总", "q1")).toEqual(["A2"]);
    expect(wb.replaceAll("汇总", "Q1", "第一季度")).toBe(1);
    expect(wb.value("汇总", "A2")).toBe("第一季度");
  });
});

describe("sheets and history", () => {
  it("adds, renames and removes sheets, updating qualified references", () => {
    const wb = fresh();
    wb.renameSheet(1, "明细");
    expect(wb.sheets[1].name).toBe("明细");
    expect(wb.sheets[0].cells.get("A6")?.formula).toBe("'明细'!B2");
    expect(wb.value("汇总", "A6")).toBe(42);
    const added = wb.addSheet();
    expect(added.name).toBe("Sheet3");
    expect(wb.removeSheet(2)).toBe(true);
    expect(wb.removeSheet(0)).toBe(true);
    expect(wb.sheets.map((s) => s.name)).toEqual(["明细"]);
    expect(wb.removeSheet(0)).toBe(false);
  });

  it("steps back and forward over edits", () => {
    const wb = fresh();
    const history = new History();
    history.record(wb);
    wb.setCell("汇总", "B2", "99");
    expect(wb.value("汇总", "B4")).toBe(99 + 18300.5);
    expect(history.canUndo(wb)).toBe(true);
    history.undoStep(wb);
    expect(wb.value("汇总", "B2")).toBe(12500);
    expect(history.canRedo(wb)).toBe(true);
    history.redoStep(wb);
    expect(wb.value("汇总", "B2")).toBe(99);
    history.record(wb);
    wb.deleteRows("汇总", 1, 1);
    history.undoStep(wb);
    expect(wb.value("汇总", "B2")).toBe(99);
    expect(wb.sheets[0].merges).toEqual(["A1:C1"]);
  });
});

describe("xlsx writer round trip", () => {
  const wb = fresh();
  const bytes = writeXlsx(wb);
  const again = Spreadsheet.openBytes(bytes);

  it("keeps the package readable and complete", () => {
    expect(OpcPackage.from(bytes).kind).toBe("xlsx");
    expect(again.sheets.map((s) => s.name)).toEqual(["汇总", "数据2"]);
    expect(again.activeSheet).toBe(0);
  });

  it("keeps the style chain: number formats, fonts, fills, borders, alignment", () => {
    expect(again.styleAt("汇总", 1, 0)).toMatchObject({
      bold: true,
      fontColor: "1F3864",
      bgColor: "FFF2CC",
      align: "center",
      border: { top: true, right: true, bottom: true, left: true },
    });
    expect(again.styleAt("汇总", 2, 1).numFmt).toBe("0.00%");
    expect(again.styleAt("汇总", 2, 4).fontColor).toBe("C00000");
    expect(again.display("汇总", "C2").text).toBe("18.75%");
    expect(again.display("汇总", "C5").text).toBe("1848.03");
  });

  it("keeps shared strings with their runs, geometry and cross-sheet formulas", () => {
    expect(again.sheets[0].cells.get("A4")?.runs?.map((r) => r.text)).toEqual(["合计", "（含税）"]);
    expect(again.sheets[0].cells.get("A4")?.runs?.[1]?.bold).toBe(true);
    expect(again.sheets[0].colWidths.get(0)).toBe(18);
    expect(again.sheets[0].colWidths.get(3)).toBe(13);
    expect(again.sheets[0].freeze).toEqual({ rows: 1, cols: 0 });
    expect(again.sheets[0].merges).toEqual(["A1:C1"]);
    expect(again.sheets[0].cells.get("C4")?.formula).toBe("B4/$B$5");
    expect(again.value("汇总", "A6")).toBe(42);
    expect(again.value("汇总", "B6")).toBe("#DIV/0!");
    expect(new Evaluator(again, "汇总").eval(parse("利润率"), "汇总")).toBe(0.1875);
  });

  it("carries edits through a save as well as the original file", () => {
    const edited = fresh();
    edited.setCell("汇总", "E1", "=SUM(B2:B3)*2");
    edited.setStyle("汇总", box("E1:E1"), { bold: true, numFmt: "#,##0.00", wrap: true });
    edited.setColWidth("汇总", 4, 22);
    edited.setRowHeight("汇总", 0, 30);
    edited.setFreeze("汇总", 1, 4);
    const back = Spreadsheet.openBytes(writeXlsx(edited));
    expect(back.value("汇总", "E1")).toBe(61601);
    expect(back.display("汇总", "E1").text).toBe("61,601.00");
    expect(back.styleAt("汇总", 4, 0)).toMatchObject({ bold: true, wrap: true });
    expect(back.sheets[0].colWidths.get(4)).toBe(22);
    expect(back.sheets[0].rowHeights.get(0)).toBe(30);
    expect(back.sheets[0].freeze).toEqual({ rows: 1, cols: 4 });
    expect(refOf(4, 0)).toBe("E1");
    expect(rangeSpec(rect(0, 0, 1, 1))).toBe("A1:B2");
  });

  it("keeps built-in accounting formats by id", () => {
    expect(again.styleAt("汇总", 1, 3).numFmt).toBe(BUILTIN[43]);
    expect(again.display("汇总", "B4").text).toBe("30,801");
    const back = Spreadsheet.openBytes(writeXlsx(again));
    expect(back.display("汇总", "B4").text).toBe("30,801");
    expect(back.styleTable().length).toBe(again.styleTable().length);
  });

  it("keeps the style table stable across saves", () => {
    const back = Spreadsheet.openBytes(writeXlsx(Spreadsheet.openBytes(writeXlsx(wb))));
    expect(back.styleTable().length).toBe(again.styleTable().length);
  });
});

describe("a workbook created from nothing", () => {
  it("round-trips through its own writer and reader", () => {
    const wb = Spreadsheet.blank();
    expect(wb.sheets.map((s) => s.name)).toEqual(["Sheet1"]);
    expect(wb.source()).toBeUndefined();

    wb.setCell("Sheet1", "A1", "新建");
    wb.setCell("Sheet1", "B1", "42");
    const back = Spreadsheet.openBytes(writeXlsx(wb));
    expect(back.sheets.map((s) => s.name)).toEqual(["Sheet1"]);
    expect(back.value("Sheet1", "A1")).toBe("新建");
    expect(back.value("Sheet1", "B1")).toBe(42);
  });
});
