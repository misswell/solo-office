import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { Spreadsheet, schemeColor, type CellStyle } from "../src/core/xlsx.js";
import { writeXlsx } from "../src/core/write.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { borderLineCss } from "../src/ui/sheet-view.js";

/**
 * Cell borders carry a line token and an ink, and both are visible: over the authorized 73-file xlsx cohort,
 * 41 files draw `double` (243 instances) and `medium` (81) sides, and 693 sides name a colour (plus 176 theme
 * slots with 46 tints). The old model kept four booleans and drew every side as one thin dark line.
 *
 * The fixture below is the shape WPS writes: a border record per side, an RGB or theme colour child, and
 * sides that state no line as a bare `<left/>`. A bare side must stay bare — the old reader treated an absent
 * `@_style` as drawn, so a rule under one row also boxed the other three sides.
 */
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>`;

const STYLES = `<styleSheet xmlns="${NS}">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="3">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FF7F7F7F"/></left><right/><top style="double"><color rgb="FF3F3F3F"/></top><bottom/><diagonal/></border>
<border><left/><right/><top/><bottom style="medium"><color theme="5" tint="0.3999755851924192"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/></cellXfs>
</styleSheet>`;

function bordered(): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      DECL +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      DECL +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    "xl/workbook.xml": strToU8(
      DECL + `<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="表" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      DECL +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`,
    ),
    "xl/styles.xml": strToU8(DECL + STYLES),
    "xl/theme/theme1.xml": strToU8(DECL + THEME),
    "xl/worksheets/sheet1.xml": strToU8(
      DECL + `<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c></row></sheetData></worksheet>`,
    ),
  };
  return zipSync(files, { level: 6 });
}

const open = (bytes: Uint8Array): Spreadsheet => Spreadsheet.openBytes(bytes);
const stylesPart = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("xl/styles.xml")!);

describe("xlsx cell borders keep their line and ink", () => {
  it("reads each side's own token and colour, and leaves a side that states no line bare", () => {
    const wb = open(bordered());
    expect(wb.styleAt("表", 0, 0).border).toEqual({
      left: { style: "thin", color: "7F7F7F" },
      top: { style: "double", color: "3F3F3F" },
    });
  });

  it("resolves a theme colour with its tint to the ink Excel names", () => {
    const wb = open(bordered());
    // accent2 (ED7D31) at tint 0.4 is Excel's "Lighter 40%" F4B183; 0.6 is "Lighter 60%" F8CBAD.
    expect(wb.styleAt("表", 1, 0).border).toEqual({ bottom: { style: "medium", color: "F4B183" } });
    expect(schemeColor([undefined, undefined, undefined, undefined, undefined, "ED7D31"], 5, 0.3999755851924192)).toBe("F4B183");
    expect(schemeColor([undefined, undefined, undefined, undefined, undefined, "ED7D31"], 5, 0.6)).toBe("F8CBAD");
    expect(schemeColor([undefined, undefined, undefined, undefined, undefined, "ED7D31"], 5, -0.25)).toBe("C55A11");
    expect(schemeColor([undefined, undefined, undefined, undefined, undefined, "ED7D31"], 5, 0)).toBe("ED7D31");
    expect(schemeColor(undefined, 5, 0)).toBeUndefined();
  });

  it("draws the token it read: weight, line style and ink", () => {
    expect(borderLineCss({ style: "double", color: "3F3F3F" })).toBe("3px double #3F3F3F");
    expect(borderLineCss({ style: "medium", color: "F8CBAD" })).toBe("2px solid #F8CBAD");
    expect(borderLineCss({ style: "thin", color: "7F7F7F" })).toBe("1px solid #7F7F7F");
    expect(borderLineCss({ style: "dashed" })).toBe("1px dashed #40464f");
    expect(borderLineCss({ style: "thick" })).toBe("3px solid #40464f");
    expect(borderLineCss({ style: "somethingNew" })).toBe("1px solid #40464f");
  });

  it("writes an editor's line and colour back, and clears every side when asked to", () => {
    const wb = Spreadsheet.blank();
    const line: CellStyle["border"] = { top: { style: "double", color: "FF0000" }, left: { style: "thick" } };
    wb.setStyle(0, { c1: 0, r1: 0, c2: 0, r2: 0 }, { border: line });
    const bytes = writeXlsx(wb);
    const xml = stylesPart(bytes);
    expect(xml).toContain('<top style="double"><color rgb="FFFF0000"/></top>');
    expect(xml).toContain('<left style="thick"></left>');
    expect(xml).toContain("<right/>");
    expect(open(bytes).styleAt(0, 0, 0).border).toEqual(line);

    const cleared = open(bytes);
    cleared.setStyle(0, { c1: 0, r1: 0, c2: 0, r2: 0 }, { border: undefined });
    expect(open(writeXlsx(cleared)).styleAt(0, 0, 0).border).toBeUndefined();
  });

  it("keeps a sourced workbook's own border records when another cell is edited", () => {
    const wb = open(bordered());
    const before = stylesPart(bordered());
    wb.setCell("表", "A1", "42");
    const bytes = writeXlsx(wb);
    const after = stylesPart(bytes);
    // The source records are patched, not rebuilt: the two ruled borders survive byte for byte.
    expect(before).toContain('<border><left style="thin"><color rgb="FF7F7F7F"/></left><right/><top style="double"><color rgb="FF3F3F3F"/></top><bottom/><diagonal/></border>');
    expect(after).toContain('<border><left style="thin"><color rgb="FF7F7F7F"/></left><right/><top style="double"><color rgb="FF3F3F3F"/></top><bottom/><diagonal/></border>');
    expect(after).toContain('<bottom style="medium"><color theme="5" tint="0.3999755851924192"/></bottom>');
    expect(open(bytes).styleAt("表", 0, 0).border).toEqual({ left: { style: "thin", color: "7F7F7F" }, top: { style: "double", color: "3F3F3F" } });
  });

  it("counts the line when deciding two styles differ", () => {
    const wb = Spreadsheet.blank();
    wb.setStyle(0, { c1: 0, r1: 0, c2: 0, r2: 0 }, { border: { top: { style: "thin" } } });
    wb.setStyle(0, { c1: 1, r1: 0, c2: 1, r2: 0 }, { border: { top: { style: "double" } } });
    const reopened = open(writeXlsx(wb));
    expect(reopened.styleAt(0, 0, 0).border).toEqual({ top: { style: "thin" } });
    expect(reopened.styleAt(0, 1, 0).border).toEqual({ top: { style: "double" } });
  });
});

// Guard against a fixture that silently stopped carrying the parts it claims to.
it("fixture carries styles, theme and sheet", () => {
  const parts = Object.keys(unzipSync(bordered()));
  expect(parts).toEqual(expect.arrayContaining(["xl/styles.xml", "xl/theme/theme1.xml", "xl/worksheets/sheet1.xml"]));
});
