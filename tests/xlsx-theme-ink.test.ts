import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { Spreadsheet, schemeColor } from "../src/core/xlsx.js";

/**
 * 字体与填充的墨色也走主题槽：授权 73 份 xlsx 的 styles.xml 里 383 枚字体色写 `theme=`（73/73 份，
 * 2,546 个单元格 / 65 份实际用上）、968 只纯色填充写 `theme=`（41/73 份，84 个单元格 / 1 份实际用上），
 * 都没有 rgb。此前只读 `rgb`，这些字和格子全部退回默认黑/无色。这里钉住“rgb 优先、theme 其次、
 * tint 按 clrScheme 折算”，并让语料里的主题件确实解出颜色。
 */
const CORPUS = ".scratch/wps-corpus";
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>`;

const STYLES = `<styleSheet xmlns="${NS}">
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="14"/><color theme="4"/><name val="Calibri"/></font>
<font><sz val="12"/><color theme="1"/><name val="Calibri"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor theme="4"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor theme="5" tint="0.5"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
</cellXfs>
</styleSheet>`;

function themed(): Spreadsheet {
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
      DECL +
        `<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>主题字</t></is></c><c r="B1" s="2" t="inlineStr"><is><t>主题底</t></is></c><c r="C1" s="3" t="inlineStr"><is><t>调淡底</t></is></c></row></sheetData></worksheet>`,
    ),
  };
  return Spreadsheet.openBytes(zipSync(files, { level: 6 }));
}

describe("xlsx 主题墨色（a:clrScheme）落到字与格", () => {
  it("主题字体色与主题填充按 clrScheme 解出，rgb 仍优先，tint 折算", () => {
    const wb = themed();
    // styleAt 的坐标是 (列, 行)；三格同在第 1 行。
    expect(wb.styleAt("表", 0, 0).fontColor).toBe("4472C4");
    expect(wb.styleAt("表", 1, 0).fontColor).toBe("000000");
    expect(wb.styleAt("表", 1, 0).bgColor).toBe("4472C4");
    expect(wb.styleAt("表", 2, 0).bgColor).toBe(schemeColor(["000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31"], 5, 0.5));
  });

  it.skipIf(!existsSync(CORPUS))("语料里的主题字体与主题填充都解出了六位色值", () => {
    const files = readdirSync(CORPUS).sort().filter((f) => /\.xlsx$/i.test(f)).map((f) => join(CORPUS, f));
    let fontThemes = 0;
    let fillThemes = 0;
    let resolvedFonts = 0;
    let resolvedFills = 0;
    const fontDecks = new Set<string>();
    const fillDecks = new Set<string>();
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const styles = new TextDecoder().decode(unzipSync(bytes)["xl/styles.xml"] ?? new Uint8Array());
      const fontBodies = (styles.match(/<fonts\b[\s\S]*?<\/fonts>/)?.[0] ?? "").match(/<font>[\s\S]*?<\/font>/g) ?? [];
      const fillBodies = (styles.match(/<fills\b[\s\S]*?<\/fills>/)?.[0] ?? "").match(/<fill>[\s\S]*?<\/fill>/g) ?? [];
      const themedFonts = fontBodies.filter((f) => /<color\b[^>]*theme=/.test(f) && !/<color\b[^>]*rgb=/.test(f));
      const themedFills = fillBodies.filter(
        (f) => /patternType="solid"/.test(f) && /<fgColor\b[^>]*theme=/.test(f) && !/<fgColor\b[^>]*rgb=/.test(f),
      );
      const wb = Spreadsheet.openBytes(bytes);
      const stylesUsed = (wb as unknown as { styles: { fontColor?: string; bgColor?: string }[] }).styles;
      if (themedFonts.length) {
        fontDecks.add(file);
        fontThemes += themedFonts.length;
        resolvedFonts += Math.min(themedFonts.length, stylesUsed.filter((s) => s.fontColor !== undefined).length);
      }
      if (themedFills.length) {
        fillDecks.add(file);
        fillThemes += themedFills.length;
        resolvedFills += Math.min(themedFills.length, stylesUsed.filter((s) => s.bgColor !== undefined).length);
      }
      for (const s of stylesUsed) {
        if (s.fontColor !== undefined) expect(s.fontColor, file.split("__").pop()).toMatch(/^[0-9A-F]{6}$/);
        if (s.bgColor !== undefined) expect(s.bgColor, file.split("__").pop()).toMatch(/^[0-9A-F]{6}$/);
      }
    }
    expect(fontThemes, "主题字体色受众").toBeGreaterThanOrEqual(383);
    expect(fillThemes, "主题填充受众").toBeGreaterThanOrEqual(968);
    expect(resolvedFonts).toBeGreaterThan(0);
    expect(resolvedFills).toBeGreaterThan(0);
    expect(fontDecks.size).toBeGreaterThanOrEqual(73);
    expect(fillDecks.size).toBeGreaterThanOrEqual(41);
  });
});
