import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { Spreadsheet, rowHeightPx } from "../src/core/xlsx.js";
import { writeXlsx } from "../src/core/write.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { filterDataRegion, filterMatches, filterRange, type FilterOperator } from "../src/core/filter.js";
import { History } from "../src/core/history.js";
import { rect } from "../src/core/sheet-ops.js";

function sample(): Spreadsheet {
  const wb = Spreadsheet.blank();
  const rows = [["项目", "地区", "金额"], ["Alpha", "东", "10"], ["Beta", "西", "20"], ["Alpha", "西", "30"], ["Gamma", "东", "40"], ["", "东", "50"]];
  rows.forEach((row, r) => row.forEach((v, c) => wb.setCell(0, `${String.fromCharCode(65 + c)}${r + 1}`, v)));
  wb.enableAutoFilter(0, "A1:C6");
  return wb;
}
const hidden = (wb: Spreadsheet): number[] => { wb.refreshFilters(); return [...wb.sheets[0].filteredRows!]; };
const patch = (wb: Spreadsheet, change: (xml: string) => string): Spreadsheet => {
  const pkg = OpcPackage.from(writeXlsx(wb)), path = "xl/worksheets/sheet1.xml";
  pkg.parts.set(path, strToU8(change(text(pkg.bytes(path)!))));
  return Spreadsheet.openBytes(zipSync(Object.fromEntries(pkg.parts)));
};

describe("AutoFilter workbook editing and OOXML roundtrip", () => {
  it("combines per-column value lists while keeping the header and blank criterion", () => {
    const wb = sample();
    wb.setFilterColumn(0, 0, { kind: "values", values: ["alpha"], blank: false });
    expect(hidden(wb)).toEqual([2, 4, 5]);
    wb.setFilterColumn(0, 1, { kind: "values", values: ["西"], blank: false });
    expect(hidden(wb)).toEqual([1, 2, 4, 5]);
    wb.setFilterColumn(0, 0, { kind: "values", values: [], blank: true });
    wb.setFilterColumn(0, 1);
    expect(hidden(wb)).toEqual([1, 2, 3, 4]);
    expect(rowHeightPx(wb.sheets[0], 0)).toBe(20);
    expect(rowHeightPx(wb.sheets[0], 1)).toBe(0);
  });

  it.each<[FilterOperator, number[]]>([["equal", [1, 3, 4, 5]], ["notEqual", [2]], ["lessThan", [2, 3, 4, 5]], ["lessThanOrEqual", [3, 4, 5]], ["greaterThan", [1, 2]], ["greaterThanOrEqual", [1]]])("evaluates %s numeric criteria", (operator, rows) => {
    const wb = sample(); wb.setFilterColumn(0, 2, { kind: "custom", and: false, tests: [{ operator, value: "20" }] });
    expect(hidden(wb)).toEqual(rows);
  });

  it("supports dual AND/OR predicates and wildcard escaping", () => {
    const wb = sample();
    wb.setFilterColumn(0, 2, { kind: "custom", and: true, tests: [{ operator: "greaterThan", value: "10" }, { operator: "lessThan", value: "40" }] });
    expect(hidden(wb)).toEqual([1, 4, 5]);
    wb.setFilterColumn(0, 2, { kind: "custom", and: false, tests: [{ operator: "equal", value: "10" }, { operator: "equal", value: "40" }] });
    expect(hidden(wb)).toEqual([2, 3, 5]);
    const match = (value: string, pattern: string): boolean => filterMatches(value, { kind: "custom", and: false, tests: [{ operator: "equal", value: pattern }] });
    expect(match("Alpha", "a?ph*")).toBe(true); expect(match("a*b", "a~*b")).toBe(true);
    expect(match("axb", "a~*b")).toBe(false); expect(match("a?b", "a~?b")).toBe(true);
    expect(match("a\nb", "a*b")).toBe(true);
  });

  it("recalculates criteria from formula values after a cross-sheet input changes", () => {
    const wb = sample(); wb.addSheet("参数"); wb.setCell("参数", "A1", "15"); wb.setCell(0, "C2", "='参数'!A1");
    wb.setFilterColumn(0, 2, { kind: "custom", and: false, tests: [{ operator: "greaterThan", value: "20" }] });
    expect(hidden(wb)).toContain(1);
    wb.setCell("参数", "A1", "35"); expect(hidden(wb)).not.toContain(1);
  });

  it("writes criteria and row visibility, clears filtering without revealing manually hidden rows", () => {
    const wb = sample(); wb.hideRow(0, 4, true);
    wb.setFilterColumn(0, 0, { kind: "values", values: ["Alpha"], blank: false });
    const back = Spreadsheet.openBytes(writeXlsx(wb)), pkg = OpcPackage.from(writeXlsx(wb));
    expect(hidden(back)).toEqual([2, 4, 5]);
    expect(pkg.xml("xl/worksheets/sheet1.xml").worksheet.sheetPr["@_filterMode"]).toBe("1");
    // Preserve the unambiguous manual row through clear/save; manual+filter overlap
    // in imported files shares one hidden flag and cannot be distinguished.
    wb.clearFilterCriteria(0);
    const clear = Spreadsheet.openBytes(writeXlsx(wb));
    expect(hidden(clear)).toEqual([]); expect(clear.sheets[0].hiddenRows?.has(4)).toBe(true);
    expect(clear.sheets[0].autoFilter?.ref).toBe("A1:C6");
    const tree = OpcPackage.from(writeXlsx(wb)).xml("xl/worksheets/sheet1.xml").worksheet;
    expect(tree.sheetPr?.["@_filterMode"]).toBeUndefined();
    wb.removeAutoFilter(0); expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].autoFilter).toBeUndefined();
  });

  it("preserves manual visibility through supported filter open/clear when the row matches", () => {
    const wb = sample(); wb.hideRow(0, 1, true);
    wb.setFilterColumn(0, 0, { kind: "values", values: ["Alpha"], blank: false });
    const back = Spreadsheet.openBytes(writeXlsx(wb)); back.clearFilterCriteria(0);
    const reopened = Spreadsheet.openBytes(writeXlsx(back));
    expect(reopened.sheets[0].hiddenRows?.has(1)).toBe(true);
    expect(reopened.sheets[0].hiddenRows?.has(2)).toBe(false);
  });

  it("roundtrips custom conditions and blank-only filters", () => {
    const wb = sample(); wb.setFilterColumn(0, 2, { kind: "custom", and: true, tests: [{ operator: "greaterThanOrEqual", value: "20" }, { operator: "lessThanOrEqual", value: "40" }] });
    expect(hidden(Spreadsheet.openBytes(writeXlsx(wb)))).toEqual([1, 5]);
    wb.clearFilterCriteria(0); wb.setFilterColumn(0, 0, { kind: "values", values: [], blank: true });
    expect(hidden(Spreadsheet.openBytes(writeXlsx(wb)))).toEqual([1, 2, 3, 4]);
    wb.setFilterColumn(0, 0, { kind: "values", values: [], blank: false });
    expect(hidden(Spreadsheet.openBytes(writeXlsx(wb)))).toEqual([1, 2, 3, 4, 5]);
  });

  it("preserves unknown criteria and cached row visibility on unrelated edits", () => {
    const wb = patch(sample(), (xml) => xml.replace('<autoFilter ref="A1:C6"/>', '<autoFilter ref="A1:C6"><filterColumn colId="0"><colorFilter dxfId="0" cellColor="1"/></filterColumn></autoFilter>').replace('<row r="2"', '<row hidden="1" r="2"'));
    const rev = wb.revision;
    expect(wb.sheets[0].autoFilter?.unsupported).toBeTruthy();
    expect(() => wb.setFilterColumn(0, 0, { kind: "values", values: ["Alpha"], blank: false })).toThrow();
    expect(wb.revision).toBe(rev); wb.setCell(0, "D2", "无关编辑");
    const xml = text(OpcPackage.from(writeXlsx(wb)).bytes("xl/worksheets/sheet1.xml")!);
    expect(xml).toContain('<colorFilter dxfId="0" cellColor="1"/>');
    expect(Spreadsheet.openBytes(writeXlsx(wb)).sheets[0].hiddenRows?.has(1)).toBe(true);
  });

  it("keeps unowned root/column extensions when editing another criterion", () => {
    const initial = sample(); initial.setFilterColumn(0, 0, { kind: "values", values: ["Alpha"], blank: false });
    const pkg = OpcPackage.from(writeXlsx(initial));
    const valid = text(pkg.bytes("xl/worksheets/sheet1.xml")!).replace('<autoFilter ref="A1:C6">', '<autoFilter ref="A1:C6" custom="keep">').replace('<filterColumn colId="0">', '<filterColumn colId="0" hiddenButton="0">').replace('</filterColumn>', '<extLst><ext uri="column-extra"/></extLst></filterColumn>').replace('</autoFilter>', '<sortState ref="A2:C6"/><extLst><ext uri="root-extra"/></extLst></autoFilter>');
    pkg.parts.set("xl/worksheets/sheet1.xml", strToU8(valid));
    const source = Spreadsheet.openBytes(zipSync(Object.fromEntries(pkg.parts)));
    source.setFilterColumn(0, 2, { kind: "custom", and: false, tests: [{ operator: "greaterThan", value: "15" }] });
    const saved = OpcPackage.from(writeXlsx(source)).xml("xl/worksheets/sheet1.xml").worksheet.autoFilter;
    expect(saved["@_custom"]).toBe("keep"); expect(saved.filterColumn[0]["@_hiddenButton"]).toBe("0");
    expect(saved.filterColumn[0].extLst.ext["@_uri"]).toBe("column-extra"); expect(saved.extLst.ext["@_uri"]).toBe("root-extra"); expect(saved.sortState["@_ref"]).toBe("A2:C6");
  });

  it("tracks inserted/deleted rows and columns and removes a deleted header", () => {
    const wb = sample(); wb.setFilterColumn(0, 2, { kind: "values", values: ["20"], blank: false });
    wb.insertRows(0, 0); expect(wb.sheets[0].autoFilter?.ref).toBe("A2:C7");
    wb.insertCols(0, 1); expect(wb.sheets[0].autoFilter?.ref).toBe("A2:D7"); expect(wb.sheets[0].autoFilter?.columns[0].id).toBe(3);
    const back = Spreadsheet.openBytes(writeXlsx(wb)); expect(back.sheets[0].autoFilter?.columns[0].id).toBe(3);
    wb.deleteCols(0, 0); expect(wb.sheets[0].autoFilter?.ref).toBe("A2:C7"); expect(wb.sheets[0].autoFilter?.columns[0].id).toBe(2);
    wb.deleteRows(0, 1); expect(wb.sheets[0].autoFilter).toBeUndefined(); expect(hidden(wb)).toEqual([]);
  });

  it("undoes filtering and preserves manual row visibility in history", () => {
    const wb = sample(), history = new History(); wb.hideRow(0, 4, true); history.record(wb);
    wb.setFilterColumn(0, 0, { kind: "values", values: ["Alpha"], blank: false }); expect(hidden(wb)).toEqual([2, 4, 5]);
    history.undoStep(wb); expect(hidden(wb)).toEqual([]); expect(wb.sheets[0].hiddenRows?.has(4)).toBe(true);
    history.redoStep(wb); expect(hidden(wb)).toEqual([2, 4, 5]);
  });

  it("rejects invalid ranges/operators atomically and finds a bounded current region", () => {
    const wb = sample(), rev = wb.revision;
    for (const ref of ["A0:B2", "XFE1:XFE2", "A1:A1048577", "Sheet1!A1:B3", "A1:B3 trailing"]) expect(() => wb.enableAutoFilter(0, ref)).toThrow();
    expect(() => wb.setFilterColumn(0, 3, { kind: "values", values: [], blank: false })).toThrow();
    expect(() => wb.setFilterColumn(0, 0, { kind: "custom", and: true, tests: [] })).toThrow();
    expect(wb.revision).toBe(rev);
    const occupied = new Set(["1,1", "2,1", "1,2", "2,2", "1,4", "2,4"]);
    expect(filterDataRegion({ c: 2, r: 2 }, rect(0, 0, 10, 10), (c, r) => occupied.has(`${c},${r}`))).toEqual(rect(1, 1, 2, 2));
    expect(filterRange("$C$6:$A$1")).toEqual(rect(0, 0, 2, 5));
  });
});
