import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { strToU8, zipSync } from "fflate";
import { Spreadsheet } from "../src/core/xlsx.js";
import { writeXlsx } from "../src/core/write.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text as partText } from "../src/core/opc.js";
import { rect, refOf } from "../src/core/sheet-ops.js";
import { scaleColorAt } from "../src/core/xlsx.js";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";

describe("sheet pictures", () => {
  it("round-trips an inserted picture through the package", () => {
    const wb = Spreadsheet.blank();
    wb.addImage("Sheet1", { c: 2, r: 1 }, PNG, 120, 80, "logo");
    const bytes = writeXlsx(wb);
    const back = Spreadsheet.openBytes(bytes);
    const img = back.sheet("Sheet1")!.images[0];
    expect(img.col).toBe(2);
    expect(img.row).toBe(1);
    expect(img.cx).toBe(120 * 9525);
    expect(img.cy).toBe(80 * 9525);
    expect(img.path).toMatch(/^xl\/media\/image\d+\.png$/);
    expect(img.src.startsWith("data:image/png;base64,")).toBe(true);

    const pkg = OpcPackage.from(bytes);
    expect(pkg.parts.has("xl/drawings/drawing1.xml")).toBe(true);
    expect(pkg.parts.has(img.path)).toBe(true);
    const sheetRels = [...pkg.rels("xl/worksheets/sheet1.xml").values()];
    expect(sheetRels.some((r) => r.type.endsWith("/drawing"))).toBe(true);
    const drawRels = [...pkg.rels("xl/drawings/drawing1.xml").values()];
    expect(drawRels[0].type.endsWith("/image")).toBe(true);
    expect(pkg.contentType("xl/drawings/drawing1.xml")).toContain("drawing+xml");
    // The sheet body gained exactly one drawing reference.
    expect(pkg.text("xl/worksheets/sheet1.xml")!.match(/<drawing /g)).toHaveLength(1);
  });

  it("keeps a second picture on its own anchor and reuses freed media names", () => {
    const wb = Spreadsheet.blank();
    wb.addImage("Sheet1", { c: 0, r: 0 }, PNG, 40, 40);
    wb.addImage("Sheet1", { c: 4, r: 3 }, PNG, 60, 30);
    const back = Spreadsheet.openBytes(writeXlsx(wb));
    const images = back.sheet("Sheet1")!.images;
    expect(images).toHaveLength(2);
    expect(images[0].path).not.toBe(images[1].path);
    expect(images[1].col).toBe(4);
  });

  it("drops a removed picture on save", () => {
    const wb = Spreadsheet.blank();
    wb.addImage("Sheet1", { c: 0, r: 0 }, PNG, 40, 40);
    wb.removeImage("Sheet1", 0);
    const pkg = OpcPackage.from(writeXlsx(wb));
    expect(pkg.text("xl/worksheets/sheet1.xml")!).not.toContain("<drawing ");
    expect(pkg.parts.has("xl/drawings/drawing1.xml")).toBe(false);
  });
});

describe("fill handle", () => {
  const wb = (): Spreadsheet => {
    const w = Spreadsheet.blank();
    w.setCell("Sheet1", "A1", "1");
    w.setCell("Sheet1", "A2", "4");
    w.setCell("Sheet1", "B1", "=A1*2");
    w.setCell("Sheet1", "C1", "备注");
    return w;
  };

  it("extends a numeric series downward", () => {
    const w = wb();
    w.fill("Sheet1", rect(0, 0, 0, 1), rect(0, 0, 0, 4));
    expect(w.value("Sheet1", "A3")).toBe(7);
    expect(w.value("Sheet1", "A4")).toBe(10);
    expect(w.value("Sheet1", "A5")).toBe(13);
  });

  it("shifts relative references and copies text sideways", () => {
    const w = wb();
    w.fill("Sheet1", rect(0, 0, 2, 0), rect(0, 0, 2, 1));
    // A2 continues its own series (single source ⇒ +1), so B2 reads A2*2 = 4.
    expect(w.sheet("Sheet1")!.cells.get("B2")!.formula).toBe("A2*2");
    expect(w.value("Sheet1", "A2")).toBe(2);
    expect(w.value("Sheet1", "B2")).toBe(4);
    expect(w.value("Sheet1", "C2")).toBe("备注");
  });

  it("fills upward with a descending series", () => {
    const w = wb();
    w.setCell("Sheet1", "A2", "10");
    w.fill("Sheet1", rect(0, 1, 0, 1), rect(0, 0, 0, 1));
    expect(w.value("Sheet1", "A1")).toBe(9);
  });
});

describe("docx inserts", () => {
  it("writes and re-reads a table, a picture and a page break", async () => {
    const doc = WordDocument.blank();
    doc.blocks.push(
      {
        kind: "table",
        columns: 3,
        rows: [
          [0, 1, 2].map((c): import("../src/core/docx.js").TableCell => ({ paragraphs: [{ kind: "paragraph", inlines: [{ text: `格${c}` }] }], widthTwips: 3008 })),
        ],
      },
      { kind: "paragraph", inlines: [], pageBreakBefore: true },
      {
        kind: "paragraph",
        inlines: [{ text: "", image: { src: PNG, width: 96, height: 64, name: "logo", path: "word/media/solo-image1.png" } }],
      },
    );
    const pkg = OpcPackage.from(writeDocx(doc));
    const back = WordDocument.openBytes(writeDocx(doc));
    // blocks[0] is the blank document's opening paragraph; the inserts follow it.
    expect(back.blocks[1].kind).toBe("table");
    const table = back.blocks[1] as Extract<(typeof back.blocks)[number], { kind: "table" }>;
    expect(table.rows[0][2].paragraphs[0].kind === "paragraph" && table.rows[0][2].paragraphs[0].inlines[0].text).toBe("格2");
    expect(back.blocks[2].kind === "paragraph" && back.blocks[2].pageBreakBefore).toBe(true);
    const img = back.blocks[3].kind === "paragraph" ? back.blocks[3].inlines[0].image : undefined;
    expect(img?.path).toBe("word/media/solo-image1.png");
    expect(pkg.parts.has("word/media/solo-image1.png")).toBe(true);
    expect(pkg.text("word/document.xml")!).toContain("<w:pageBreakBefore/>");
  });
});

describe("conditional formatting", () => {
  const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const W = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const cfWorkbook = (): Uint8Array => {
    const sheet =
      `${DECL}<worksheet xmlns="${W}" xmlns:r="${R}"><sheetData>` +
      `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>9</v></c></row>` +
      `<row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>` +
      `<row r="3"><c r="A3"><v>5</v></c></row>` +
      `</sheetData>` +
      `<conditionalFormatting sqref="B1:B2">` +
      `<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>` +
      `<cfRule type="iconSet" priority="2"><iconSet iconSet="3TrafficLights1"/></cfRule>` +
      `</conditionalFormatting>` +
      `<conditionalFormatting sqref="C1"><cfRule type="cellIs" dxfId="0" priority="3" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>` +
      `</worksheet>`;
    const styles =
      `${DECL}<styleSheet xmlns="${W}"><fonts count="1"><font><sz val="11"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>` +
      `<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
      `</styleSheet>`;
    return zipSync({
      "[Content_Types].xml": strToU8(`${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`),
      "_rels/.rels": strToU8(`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      "xl/workbook.xml": strToU8(`${DECL}<workbook xmlns="${W}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
      "xl/_rels/workbook.xml.rels": strToU8(`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
      "xl/worksheets/sheet1.xml": strToU8(sheet),
      "xl/styles.xml": strToU8(styles),
    });
  };

  it("reads colorScale, cellIs with dxf, and keeps unknown kinds raw", () => {
    const wb = Spreadsheet.openBytes(cfWorkbook());
    const cond = wb.sheet("Sheet1")!.cond;
    expect(cond).toHaveLength(2);
    const scale = cond[0].rules.find((r) => r.type === "colorScale");
    expect(scale?.colors).toEqual(["F8696B", "63BE7B"]);
    expect(cond[0].rules.find((r) => r.type === "iconSet")).toBeDefined();
    const cellIs = cond[1].rules[0];
    expect(cellIs.op).toBe("greaterThan");
    expect(cellIs.formula).toBe("5");
    expect(cellIs.dxfId).toBe(0);
    expect(wb.dxfOf(0)?.bgColor).toBe("FFC7CE");
    expect(wb.dxfOf(0)?.fontColor).toBe("9C0006");
  });

  it("round-trips untouched rules byte-safely, including the ones it cannot evaluate", () => {
    const src = Spreadsheet.openBytes(cfWorkbook());
    const back = Spreadsheet.openBytes(writeXlsx(src));
    const rules = back.sheet("Sheet1")!.cond.flatMap((c) => c.rules);
    expect(rules.find((r) => r.type === "iconSet")).toBeDefined();
    expect(rules.find((r) => r.type === "colorScale")?.colors).toEqual(["F8696B", "63BE7B"]);
    // The saved sheet XML still carries the iconSet rule text.
    const sheetXml = partText(OpcPackage.from(writeXlsx(src)).bytes("xl/worksheets/sheet1.xml")!)!;
    expect(sheetXml).toContain("3TrafficLights1");
  });

  it("applies an editor rule and clears rules over a box", () => {
    const wb = Spreadsheet.openBytes(cfWorkbook());
    wb.setCond("Sheet1", rect(0, 0, 0, 2), { type: "dataBar", color: "638EC6" });
    let saved = writeXlsx(wb);
    let back = Spreadsheet.openBytes(saved);
    expect(back.sheet("Sheet1")!.cond.some((c) => c.rules.some((r) => r.type === "dataBar"))).toBe(true);
    expect(back.sheet("Sheet1")!.cond.some((c) => c.rules.some((r) => r.type === "iconSet"))).toBe(true);

    back.clearCond("Sheet1", rect(1, 0, 1, 1)); // the B1:B2 group
    saved = writeXlsx(back);
    const final = Spreadsheet.openBytes(saved).sheet("Sheet1")!.cond;
    expect(final.some((c) => c.ref === "B1:B2")).toBe(false);
    expect(final.some((c) => c.rules.some((r) => r.type === "iconSet"))).toBe(false);
    expect(final.some((c) => c.rules.some((r) => r.type === "dataBar"))).toBe(true);
  });

  it("applies an editor cellIs rule and writes its dxf", () => {
    const wb = Spreadsheet.blank();
    wb.setCell("Sheet1", "A1", "5");
    wb.setCell("Sheet1", "A2", "50");
    wb.setCondCellIs("Sheet1", rect(0, 0, 0, 1), "greaterThan", "10");
    const saved = writeXlsx(wb);
    const styles = partText(OpcPackage.from(saved).bytes("xl/styles.xml")!)!;
    expect(styles).toContain("<dxfs");
    expect(styles).toContain("FFC7CE");
    const back = Spreadsheet.openBytes(saved);
    const rule = back.sheet("Sheet1")!.cond[0].rules[0];
    expect(rule.type).toBe("cellIs");
    expect(back.dxfOf(rule.dxfId!)?.bgColor).toBe("FFC7CE");
  });

  it("keeps aboveAverage and containsText rules verbatim", () => {
    const src = Spreadsheet.openBytes(cfWorkbook());
    void src;
    const bytes = cfWorkbook();
    // Hand-edit the sheet part to carry two more rule kinds.
    const pkg = OpcPackage.from(bytes);
    const sheetXml = partText(pkg.bytes("xl/worksheets/sheet1.xml")!)!.replace(
      "</worksheet>",
      `<conditionalFormatting sqref="A1:A3"><cfRule type="aboveAverage" priority="9"/><cfRule type="containsText" priority="10" operator="containsText" text="3"><formula>NOT(ISERROR(SEARCH("3",A1)))</formula></cfRule></conditionalFormatting></worksheet>`,
    );
    const rebuilt = zipSync({
      "xl/worksheets/sheet1.xml": strToU8(sheetXml),
      "xl/workbook.xml": pkg.bytes("xl/workbook.xml")!,
      "xl/_rels/workbook.xml.rels": pkg.bytes("xl/_rels/workbook.xml.rels")!,
      "xl/styles.xml": pkg.bytes("xl/styles.xml")!,
      "[Content_Types].xml": pkg.bytes("[Content_Types].xml")!,
      "_rels/.rels": pkg.bytes("_rels/.rels")!,
    });
    const wb = Spreadsheet.openBytes(rebuilt);
    const kinds = wb.sheet("Sheet1")!.cond.flatMap((c) => c.rules.map((r) => r.type));
    expect(kinds).toContain("aboveAverage");
    expect(kinds).toContain("containsText");
    const sheetXml2 = partText(OpcPackage.from(writeXlsx(wb)).bytes("xl/worksheets/sheet1.xml")!)!;
    expect(sheetXml2).toContain("aboveAverage");
    expect(sheetXml2).toContain("NOT(ISERROR(SEARCH");
  });

  it("interpolates color scales", () => {
    expect(scaleColorAt(["FF0000", "00FF00"], 0)).toBe("FF0000");
    expect(scaleColorAt(["FF0000", "00FF00"], 1)).toBe("00FF00");
    expect(scaleColorAt(["000000", "FFFFFF"], 0.5)).toBe("808080");
    expect(scaleColorAt(["000000", "808080", "FFFFFF"], 0.25)).toBe("404040");
  });
});

describe("docx header/footer", () => {
  const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const hfDocx = (headerBody: string): Uint8Array =>
    zipSync({
      "[Content_Types].xml": strToU8(`${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`),
      "_rels/.rels": strToU8(`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/document.xml": strToU8(`${DECL}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>正文</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId5"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`),
      "word/_rels/document.xml.rels": strToU8(`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${R}/header" Target="header1.xml"/></Relationships>`),
      "word/header1.xml": strToU8(headerBody),
    });

  it("reads header texts, edits them, and writes them back", () => {
    const doc = WordDocument.openBytes(
      hfDocx(`${DECL}<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>公司机密</w:t></w:r></w:p></w:hdr>`),
    );
    expect(doc.header?.editable).toBe(true);
    expect(doc.header?.paragraphs).toEqual(["公司机密"]);

    doc.setHeaderTexts("header", ["已修订"]);
    const back = WordDocument.openBytes(writeDocx(doc));
    expect(back.header?.paragraphs).toEqual(["已修订"]);
    const pkg = OpcPackage.from(writeDocx(doc));
    expect(pkg.text("word/header1.xml")).toContain("已修订");
    expect(pkg.text("word/document.xml")).toContain("headerReference");
  });

  it("refuses to edit a rich header and keeps its bytes", () => {
    const rich = `${DECL}<w:hdr xmlns:w="${W}"><w:tbl><w:tr><w:tc><w:p><w:r><w:t>logo</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>文本</w:t></w:r></w:p></w:hdr>`;
    const bytes = hfDocx(rich);
    const doc = WordDocument.openBytes(bytes);
    expect(doc.header?.editable).toBe(false);
    doc.setHeaderTexts("header", ["篡改"]);
    const out = writeDocx(doc);
    expect(partText(OpcPackage.from(out).bytes("word/header1.xml")!)).toContain("<w:tbl>");
  });

  it("creates a header on a blank document, wiring rels and sectPr", () => {
    const doc = WordDocument.blank();
    doc.ensureHeaderFooter("header");
    doc.setHeaderTexts("header", ["新页眉"]);
    const pkg = OpcPackage.from(writeDocx(doc));
    const back = WordDocument.openBytes(writeDocx(doc));
    expect(back.header?.paragraphs).toEqual(["新页眉"]);
    expect(pkg.text("word/header1.xml")).toContain("新页眉");
    expect(pkg.text("word/document.xml")).toContain(`headerReference`);
    const rels = [...pkg.rels("word/document.xml").values()];
    expect(rels.some((r) => r.type.endsWith("/header"))).toBe(true);
    expect(pkg.contentType("word/header1.xml")).toContain("header+xml");
  });
});

describe("charts", () => {
  it("reads openpyxl-written charts with their series", () => {
    const bytes = new Uint8Array(readFileSync(new URL("./fixtures/chart-sample.xlsx", import.meta.url).pathname));
    const wb = Spreadsheet.openBytes(bytes);
    const charts = wb.sheets[0].charts;
    expect(charts.map((c) => c.kind)).toEqual(["bar", "line", "pie"]);
    const bar = charts[0];
    expect(bar.title).toBe("销售对比");
    expect(bar.series.map((s) => s.name)).toEqual(["一月", "二月"]);
    expect(bar.series[0].vals).toEqual([12, 8, 20]);
    expect(bar.series[0].cats).toEqual(["A", "B", "C"]);
    expect(bar.cx).toBeGreaterThan(0);
    // Saving untouched keeps the chart parts byte-identical.
    const pkg = OpcPackage.from(writeXlsx(wb));
    const src = OpcPackage.from(bytes);
    expect(pkg.bytes("xl/charts/chart1.xml")!).toEqual(src.bytes("xl/charts/chart1.xml")!);
    expect(pkg.bytes("xl/charts/chart2.xml")!).toEqual(src.bytes("xl/charts/chart2.xml")!);
    expect(pkg.bytes("xl/charts/chart3.xml")!).toEqual(src.bytes("xl/charts/chart3.xml")!);
  });

  it("writes a chart the editor built and reads it back", () => {
    const wb = Spreadsheet.blank();
    for (const [a, b, c] of [["类别", "Q1", "Q2"], ["东", "10", "14"], ["南", "22", "8"], ["西", "15", "19"]] as [string, string, string][]) {
      const r = wb.bounds("Sheet1").maxRow + 2;
      wb.setCell("Sheet1", `A${r}`, a);
      wb.setCell("Sheet1", `B${r}`, b);
      wb.setCell("Sheet1", `C${r}`, c);
    }
    wb.addChart("bar", "Sheet1", rect(0, 1, 2, wb.bounds("Sheet1").maxRow), { c: 4, r: 0 }, "季度销售");
    const saved = writeXlsx(wb);
    const pkg = OpcPackage.from(saved);
    expect(pkg.parts.has("xl/charts/chart1.xml")).toBe(true);
    expect(pkg.contentType("xl/charts/chart1.xml")).toContain("chart+xml");
    expect([...pkg.rels("xl/drawings/drawing1.xml").values()].some((r) => r.type.endsWith("/chart"))).toBe(true);
    const back = Spreadsheet.openBytes(saved);
    const chart = back.sheet("Sheet1")!.charts[0];
    expect(chart.kind).toBe("bar");
    expect(chart.title).toBe("季度销售");
    expect(chart.series[1].vals).toEqual([14, 8, 19]);
    expect(chart.col).toBe(4);
  });
});

describe("pptx alignment", () => {
  it("maps the editor's alignment words to algn and back", () => {
    // Import lazily to keep this test file's xlsx/docx focus obvious.
    return (async () => {
      const { Presentation } = await import("../src/core/pptx.js");
      const { writePptx } = await import("../src/core/write-pptx.js");
      const pres = Presentation.blank();
      const slide = pres.slides[0];
      slide.shapes.push({
        kind: "shape",
        name: "文本",
        geometry: "rect",
        x: 100,
        y: 100,
        width: 200,
        height: 80,
        paragraphs: [{ level: 0, align: "center", runs: [{ text: "中" }] }],
      });
      pres.markDirty(slide);
      const back = Presentation.openBytes(writePptx(pres));
      const paragraph = back.slides[0].shapes[0].paragraphs?.[0];
      expect(paragraph?.align).toBe("center");
    })();
  });
});
