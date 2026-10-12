import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { OpcPackage, text } from "../src/core/opc.js";
import { Spreadsheet } from "../src/core/xlsx.js";
import { writeXlsx } from "../src/core/write.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { Presentation } from "../src/core/pptx.js";
import { writePptx } from "../src/core/write-pptx.js";
import { makeDocx } from "./helpers/make.js";

/**
 * Files WPS itself writes: package-absolute relationship targets, inline strings
 * instead of a shared-string table, and the parts a hand-rolled writer tends to
 * drop — themes, table and drawing parts, conditional formats, custom properties.
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

const partsOf = (bytes: Uint8Array): Map<string, Uint8Array> => new Map(Object.entries(unzipSync(bytes)) as [string, Uint8Array][]);
const bodyOf = (bytes: Uint8Array, path: string): string => text(partsOf(bytes).get(path)!);

function pkg(files: Record<string, string>): Uint8Array {
  const out: Record<string, Uint8Array> = {};
  for (const [path, body] of Object.entries(files)) out[path] = strToU8(XML_DECL + body);
  return zipSync(out, { level: 6 });
}

/** The unowned parts a save must carry through byte-identically. */
const CARRIED = ["docProps/custom.xml", "xl/theme/theme1.xml", "xl/tables/table1.xml", "xl/drawings/drawing1.xml", "xl/worksheets/_rels/sheet1.xml.rels"];

function makeWpsXlsx(): Uint8Array {
  const sheet1 = `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">
<dimension ref="A1:B4"/>
<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
<sheetData>
<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>产品</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>销量</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>甲</t></is></c><c r="B2"><v>120</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>乙</t></is></c><c r="B3"><v>80</v></c></row>
<row r="4"><c r="A4" t="inlineStr"><is><t>合计</t></is></c><c r="B4"><f>SUM(B2:B3)</f><v>200</v></c></row>
</sheetData>
<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" operator="lessThan" priority="1"><formula>100</formula></cfRule></conditionalFormatting>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
<drawing r:id="rId1"/>
<tableParts count="1"><tablePart r:id="rId2"/></tableParts>
</worksheet>`;

  const sheet2 = `<worksheet xmlns="${NS_MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>设置</t></is></c></row></sheetData></worksheet>`;

  const styles = `<styleSheet xmlns="${NS_MAIN}">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><color theme="4" tint="-0.25"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="center"/></xf></cellXfs>
<cellStyles count="2"><cellStyle name="常规" xfId="0" builtinId="0"/><cellStyle name="强调" xfId="0" builtinId="3"/></cellStyles>
<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>
</styleSheet>`;

  const overrides = [
    ["/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"],
    ["/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"],
    ["/xl/worksheets/sheet2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"],
    ["/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"],
    ["/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"],
    ["/xl/tables/table1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"],
    ["/xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml"],
    ["/docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml"],
  ]
    .map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`)
    .join("");

  return pkg({
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`,
    // WPS writes package-absolute targets; a reader that joins them under xl/ finds nothing.
    "_rels/.rels": `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="/xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheets><sheet name="数据" sheetId="1" r:id="rId1"/><sheet name="设置" sheetId="2" state="hidden" r:id="rId2"/></sheets><definedNames><definedName name="合计">数据!$B$4</definedName></definedNames><calcPr calcId="1"/></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${NS_R}/worksheet" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${NS_R}/worksheet" Target="/xl/worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${NS_R}/styles" Target="/xl/styles.xml"/><Relationship Id="rId4" Type="${NS_R}/theme" Target="/xl/theme/theme1.xml"/></Relationships>`,
    "xl/styles.xml": styles,
    "xl/worksheets/sheet1.xml": sheet1,
    "xl/worksheets/sheet2.xml": sheet2,
    "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${NS_R}/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId2" Type="${NS_R}/table" Target="../tables/table1.xml"/></Relationships>`,
    "xl/tables/table1.xml": `<table xmlns="${NS_MAIN}" id="1" name="表1" displayName="表1" ref="A1:B4"><tableColumns count="2"><tableColumn id="1" name="产品"/><tableColumn id="2" name="销量"/></tableColumns></table>`,
    "xl/drawings/drawing1.xml": `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>`,
    "xl/theme/theme1.xml": `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="WPS"/>`,
    "docProps/custom.xml": `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="来源"><vt:lpwstr xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">WPS</vt:lpwstr></property></Properties>`,
  });
}

describe("a workbook WPS wrote", () => {
  it("resolves package-absolute targets and reads every cell", () => {
    const bytes = makeWpsXlsx();
    const pkgSrc = OpcPackage.from(bytes);
    expect(pkgSrc.resolve("xl/workbook.xml", { id: "rId1", type: "", target: "/xl/worksheets/sheet1.xml", targetMode: "Internal" })).toBe("xl/worksheets/sheet1.xml");
    const wb = Spreadsheet.openBytes(bytes);
    expect(wb.sheets.map((s) => s.name)).toEqual(["数据", "设置"]);
    expect(wb.value("数据", "A2")).toBe("甲");
    expect(wb.value("数据", "B2")).toBe(120);
    expect(wb.value("设置", "A1")).toBe("设置");
  });

  it("an untouched save reproduces every part", () => {
    const bytes = makeWpsXlsx();
    const before = partsOf(bytes);
    const after = partsOf(writeXlsx(Spreadsheet.openBytes(bytes)));
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
    for (const [path, data] of before) expect(after.get(path)).toEqual(data);
  });

  it("an edit keeps the parts, sheet elements and style sections the model does not own", () => {
    const bytes = makeWpsXlsx();
    const before = partsOf(bytes);
    const wb = Spreadsheet.openBytes(bytes);
    wb.setCell("数据", "B2", "999");
    const out = writeXlsx(wb);
    const after = partsOf(out);

    for (const path of CARRIED) expect(after.get(path)).toEqual(before.get(path));
    expect(bodyOf(out, "_rels/.rels")).toContain('Target="/xl/workbook.xml"');

    const back = Spreadsheet.openBytes(out);
    expect(back.value("数据", "B2")).toBe(999);
    expect(back.value("数据", "A2")).toBe("甲");

    const sheet = bodyOf(out, "xl/worksheets/sheet1.xml");
    expect(sheet).toContain("<conditionalFormatting");
    expect(sheet).toContain("dxfId=\"0\"");
    expect(sheet).toContain("<drawing");
    expect(sheet).toContain("<tableParts");
    expect(sheet).toContain("<pageMargins");
    // A conditional format points into the style sheet's dxf table by index.
    const styles = bodyOf(out, "xl/styles.xml");
    expect(styles).toContain("<dxf>");
    expect(styles).toContain("强调");
    const workbook = bodyOf(out, "xl/workbook.xml");
    expect(workbook).toContain('state="hidden"');
    expect(workbook).toContain("合计");
  });

  it("keeps a hidden sheet hidden through a save", () => {
    const wb = Spreadsheet.openBytes(makeWpsXlsx());
    wb.setCell("数据", "B3", "81");
    const back = Spreadsheet.openBytes(writeXlsx(wb));
    expect(back.sheets[1].hidden).toBe("hidden");
    expect(back.value("设置", "A1")).toBe("设置");
  });

  it("an edit that creates no styles leaves the style sheet byte-identical", () => {
    const bytes = makeWpsXlsx();
    const before = partsOf(bytes);
    const wb = Spreadsheet.openBytes(bytes);
    wb.setCell("数据", "B2", "999");
    const out = writeXlsx(wb);
    // Theme colours and tints live only in these bytes; rebuilding the sheet loses them.
    expect(partsOf(out).get("xl/styles.xml")).toEqual(before.get("xl/styles.xml"));
    expect(bodyOf(out, "xl/worksheets/sheet1.xml")).toMatch(/<c r="A1" s="1" t="s">/);
    const back = Spreadsheet.openBytes(out);
    expect(back.styleAt("数据", 0, 0)).toEqual(wb.styleAt("数据", 0, 0));
    expect((bodyOf(out, "xl/worksheets/sheet1.xml").match(/<\?xml/g) ?? []).length).toBe(1);
  });

  it("a style the editor creates appends after the file's own records", () => {
    const bytes = makeWpsXlsx();
    const wb = Spreadsheet.openBytes(bytes);
    wb.setStyle("数据", { c1: 1, r1: 1, c2: 1, r2: 1 }, { bold: true });
    const out = writeXlsx(wb);
    const styles = bodyOf(out, "xl/styles.xml");
    expect(styles).toContain('color theme="4" tint="-0.25"');
    expect(styles).toContain('<fonts count="3">');
    expect(styles).toContain('<cellXfs count="3">');
    const back = Spreadsheet.openBytes(out);
    expect(back.styleAt("数据", 1, 1).bold).toBe(true);
    // The cell that already had a style keeps the record it came from.
    expect(back.styleAt("数据", 0, 0).align).toBe("center");
    expect(back.value("数据", "B4")).toBe(200);
  });
});

describe("a document WPS wrote", () => {
  it("an untouched save reproduces every part", () => {
    const bytes = makeDocx();
    const before = partsOf(bytes);
    const after = partsOf(writeDocx(WordDocument.openBytes(bytes)));
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
    for (const [path, data] of before) expect(after.get(path)).toEqual(data);
  });

  it("an edit keeps properties, styles and media byte-identical", () => {
    const bytes = makeDocx();
    const before = partsOf(bytes);
    const doc = WordDocument.openBytes(bytes);
    doc.blocks.push({ kind: "paragraph", inlines: [{ text: "补记" }] });
    const out = writeDocx(doc);
    const after = partsOf(out);
    for (const path of ["docProps/core.xml", "word/styles.xml", "word/numbering.xml", "word/media/trend.png"]) {
      expect(after.get(path)).toEqual(before.get(path));
    }
    expect(bodyOf(out, "word/document.xml")).toContain("补记");
  });
});

function makeWpsPptx(): Uint8Array {
  const slide = `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="${NS_R}">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="文本框"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="635000" y="480000"/><a:ext cx="3000000" cy="800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:p><a:r><a:t>标题</a:t></a:r></a:p></p:txBody></p:sp>
<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="连接符"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="100000" y="2000000"/><a:ext cx="2000000" cy="500000"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr>
<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp>
</p:spTree></p:cSld></p:sld>`;

  return pkg({
    "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
    "_rels/.rels": `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    "ppt/presentation.xml": `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="${NS_R}"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${NS_R}/slide" Target="slides/slide1.xml"/></Relationships>`,
    "ppt/slides/slide1.xml": slide,
  });
}

describe("a deck with a shape that has no text body", () => {
  it("keeps text typed into it", () => {
    const pres = Presentation.openBytes(makeWpsPptx());
    const slide = pres.slides[0];
    const connector = slide.shapes.find((s) => s.name === "连接符");
    expect(connector?.paragraphs).toEqual([]);
    connector!.paragraphs = [{ level: 0, runs: [{ text: "补注" }] }];
    connector!.textDirty = true;
    pres.markDirty(slide);

    const out = writePptx(pres);
    const written = bodyOf(out, "ppt/slides/slide1.xml");
    const connectorXml = /<p:cxnSp>[\s\S]*?<\/p:cxnSp>/.exec(written)?.[0] ?? "";
    expect(connectorXml).toContain("补注");
    // The created body sits after the style element, where CT_Connector orders it.
    expect(connectorXml.indexOf("<p:txBody>")).toBeGreaterThan(connectorXml.indexOf("<p:style>"));

    const back = Presentation.openBytes(out);
    expect(back.slides[0].shapes.find((s) => s.name === "连接符")?.paragraphs?.[0].runs[0].text).toBe("补注");
    expect(back.slides[0].shapes.find((s) => s.name === "文本框")?.paragraphs?.[0].runs[0].text).toBe("标题");
  });
});
