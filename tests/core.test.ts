import { describe, expect, it } from "vitest";
import { Spreadsheet } from "../src/core/xlsx.js";
import { WordDocument } from "../src/core/docx.js";
import { Presentation } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { Evaluator, colToIndex, indexToCol, parse } from "../src/core/formula.js";
import { BUILTIN, format } from "../src/core/numfmt.js";
import { writeXlsx } from "../src/core/write.js";
import { renderWord, renderSlideHtml } from "../src/core/render.js";
import { makeDocx, makePptx, makeXlsx } from "./helpers/make.js";

class Stub implements EvaluatorLike {
  constant(): any {
    return null;
  }
  formula(): string | undefined {
    return undefined;
  }
  name(): undefined {
    return undefined;
  }
}
type EvaluatorLike = ConstructorParameters<typeof Evaluator>[0] extends infer T ? T : never;

const evalFormula = (f: string): unknown => {
  const ev = new Evaluator(new Stub() as any, "Sheet1");
  return ev.eval(parse(f), "Sheet1");
};

describe("formula language", () => {
  it("respects Excel operator precedence", () => {
    expect(evalFormula("=1+2*3")).toBe(7);
    expect(evalFormula("=(1+2)*3")).toBe(9);
    expect(evalFormula("=2^3^2")).toBe(512); // ^ is right-associative
    expect(evalFormula('="a"&"b"&1')).toBe("ab1");
    expect(evalFormula("=50%")).toBe(0.5);
    expect(evalFormula("=-2^2")).toBe(4); // unary binds tighter than ^
  });

  it("compares and short-circuits errors", () => {
    expect(evalFormula('=1<"2"')).toBe(true);
    expect(evalFormula("=3>5")).toBe(false);
    expect(evalFormula("=1/0")).toBe("#DIV/0!");
    expect(evalFormula("=NOPE(1)")).toBe("#NAME?");
    expect(evalFormula("=(1/0)+5")).toBe("#DIV/0!");
  });

  it("implements the shipped function library", () => {
    expect(evalFormula("=SUM(1,2,3)")).toBe(6);
    expect(evalFormula("=ROUND(2.3456,2)")).toBe(2.35);
    expect(evalFormula("=IF(AND(TRUE,NOT(FALSE)),1,2)")).toBe(1);
    expect(evalFormula('=MID("hello",2,3)')).toBe("ell");
    expect(evalFormula('=TEXTJOIN("- ",TRUE,"a","","b")')).toBe("a- b");
    expect(evalFormula("=SUMPRODUCT(1,2)+MEDIAN(1,2,100)")).toBe(4); // 1*2 + median(1,2,100)
    expect(evalFormula("=DATE(2026,1,1)")).toBe(46023);
    expect(evalFormula("=YEAR(46023)")).toBe(2026);
    expect(evalFormula("=IFERROR(1/0,-1)")).toBe(-1);
    expect(evalFormula("=STDEV(2,4,4,4,5,5,7,9)")).toBeCloseTo(2.138, 3);
  });

  it("round-trips column names", () => {
    for (const col of ["A", "Z", "AA", "AZ", "BA", "ZZ", "AAA"]) expect(indexToCol(colToIndex(col))).toBe(col);
  });
});

describe("number formats", () => {
  it("renders the built-in codes", () => {
    expect(format(0.1875, "0.00%")).toBe("18.75%");
    expect(format(1234567.891, "#,##0.00")).toBe("1,234,567.89");
    expect(format(-1234, "#,##0;(#,##0)")).toBe("(1,234)");
    expect(format(46023, "yyyy-mm-dd")).toBe("2026-01-01");
    expect(format(0.75, "h:mm")).toBe("18:00");
    expect(format(42, "General")).toBe("42");
  });

  it("shows General as the number itself, not a rounded integer", () => {
    expect(format(30800.5, "General")).toBe("30800.5");
    expect(format(15400.25, "General")).toBe("15400.25");
    expect(format(1 / 3, "General")).toBe("0.333333333");
    expect(format(-0.1875, "General")).toBe("-0.1875");
    expect(format(1e12, "General")).toBe("1E+12");
    expect(format(1.2345678901e-12, "General")).toBe("1.23456789E-12");
  });

  it("keeps the accounting built-ins", () => {
    expect(BUILTIN[43]).toBe('_ * #,##0_ ;_ * -#,##0_ ;_ * "-"??_ ;_ @_ ');
    expect(format(30800.5, BUILTIN[43]!)).toBe("30,801");
    expect(format(30800.5, BUILTIN[42]!)).toBe("$30,800.50");
    expect(format(0.5, BUILTIN[49]!)).toBe("0.5");
  });
});

describe("xlsx reader", () => {
  const wb = Spreadsheet.openBytes(makeXlsx());

  it("reads workbook structure", () => {
    expect(wb.sheets.map((s) => s.name)).toEqual(["汇总", "数据2"]);
    expect(wb.activeSheet).toBe(0);
    const sheet = wb.sheets[0];
    expect([...sheet.colWidths.keys()]).toContain(0);
    expect(sheet.freeze).toEqual({ rows: 1, cols: 0 });
    expect(sheet.merges).toContain("A1:C1");
  });

  it("resolves shared strings, rich runs and inline strings", () => {
    expect(wb.value("汇总", "A2")).toBe("Q1");
    expect(wb.sheets[0].cells.get("A4")?.runs?.map((r) => r.text)).toEqual(["合计", "（含税）"]);
    expect(wb.value("数据2", "A2")).toBe("跨表引用");
  });

  it("recomputes stored formulas instead of trusting cached values", () => {
    expect(wb.value("汇总", "B4")).toBe(30800.5);
    expect(wb.value("汇总", "B5")).toBeCloseTo(15400.25, 6);
    expect(wb.value("汇总", "C5")).toBeCloseTo(1848.03, 2);
    expect(wb.value("汇总", "A6")).toBe(42); // cross-sheet reference
    expect(wb.value("汇总", "B6")).toBe("#DIV/0!");
  });

  it("translates shared formulas to each slave cell", () => {
    expect(wb.sheets[0].cells.get("C3")?.formula).toBe("B3/$B$5");
    expect(wb.sheets[0].cells.get("C4")?.formula).toBe("B4/$B$5");
    expect(wb.value("汇总", "C4")).toBeCloseTo(30800.5 / 15400.25, 6);
  });

  it("applies number formats from the style chain", () => {
    expect(wb.display("汇总", "C2").text).toBe("18.75%");
    expect(wb.styleOf(wb.sheets[0].cells.get("B1")!).bold).toBe(true);
    expect(wb.styleOf(wb.sheets[0].cells.get("B1")!).bgColor).toBe("FFF2CC");
  });

  it("honours defined names", () => {
    expect(wb.value("汇总", "C2")).toBe(0.1875);
    const ev = new Evaluator(wb, "汇总");
    expect(ev.eval(parse("利润率"), "汇总")).toBe(0.1875);
  });

  it("edits a cell and propagates through the dependency graph", () => {
    const editable = Spreadsheet.openBytes(makeXlsx());
    editable.setCell("汇总", "B2", "30000");
    expect(editable.value("汇总", "B4")).toBe(30000 + 18300.5);
    expect(editable.value("汇总", "C5")).toBeCloseTo((30000 + 18300.5) * 0.06, 6);
    editable.setCell("汇总", "B3", "=B2*2");
    expect(editable.value("汇总", "B4")).toBe(30000 + 60000);
    editable.setCell("汇总", "B2", "=B3");
    expect(editable.value("汇总", "B2")).toBe("#CYCLE!");
  });

  it("writes a package the reader can load back", () => {
    const bytes = writeXlsx(wb);
    const again = Spreadsheet.openBytes(bytes);
    expect(again.sheets.map((s) => s.name)).toEqual(["汇总", "数据2"]);
    expect(again.value("汇总", "B4")).toBe(30800.5);
    expect(again.value("汇总", "C2")).toBe(0.1875);
    expect(again.sheets[0].merges).toContain("A1:C1");
    expect(OpcPackage.from(bytes).kind).toBe("xlsx");
  });
});

describe("docx reader", () => {
  const doc = WordDocument.openBytes(makeDocx());

  it("maps body order into blocks", () => {
    expect(doc.blocks[0]).toMatchObject({ kind: "paragraph", styleId: "Title" });
    expect(doc.blocks[1]).toMatchObject({ styleId: "Heading1" });
    expect(doc.blocks.some((b) => b.kind === "table")).toBe(true);
    expect(doc.sections[0].pageWidthTwips).toBe(11906);
    expect(doc.title).toBe("季度经营分析");
  });

  it("carries run formatting and hyperlinks", () => {
    const p = doc.blocks[2] as any;
    const texts = p.inlines.map((i: any) => i.text);
    expect(texts.join("")).toContain("收入同比增长 18.75%");
    expect(p.inlines.find((i: any) => i.bold)?.color).toBe("#C00000");
    expect(p.inlines.find((i: any) => i.italic)?.sizePt).toBe(10);
    expect(p.inlines.find((i: any) => i.hyperlink)).toMatchObject({ text: "明细表", hyperlink: "https://example.com/detail" });
  });

  it("numbers lists per numbering.xml", () => {
    const bullets = doc.blocks.filter((b) => b.kind === "paragraph" && (b as any).numPr) as any[];
    expect(bullets).toHaveLength(4);
    expect(doc.levelOf(bullets[0].numPr.numId, 0)?.numFmt).toBe("bullet");
    expect(doc.levelOf(bullets[2].numPr.numId, 0)?.numFmt).toBe("decimal");
  });

  it("embeds media as data URIs and renders to HTML", () => {
    const html = renderWord(doc);
    expect(html).toContain("data:image/png;base64,");
    expect(html).toContain("<h1");
    expect(html).toContain('class="doc-table"');
    expect(html).toContain("page-break-before:always");
    // 宽高与底纹仍旧按文件那句上身（这一条只查「按序落在同一枚 style 里」，
    // 后面接着的是导引清单 `--gl`——那一串的形状由 tests/doc-table-gridlines.test.ts 判）。
    expect(html).toMatch(/<td style="width:160px;background:#DEEAF6;/);
  });
});

describe("pptx reader", () => {
  const pres = Presentation.openBytes(makePptx());

  it("reads slide size and order", () => {
    expect(pres.slides).toHaveLength(1);
    expect(pres.slideWidthEmu).toBe(9144000);
    expect(pres.slideHeightPx).toBe(720);
  });

  it("positions shapes and reads text runs", () => {
    const [title] = pres.slides[0].shapes.filter((s) => s.name === "标题");
    expect(title.x).toBeCloseTo(635000 / 9525, 3);
    expect(title.paragraphs?.[0].runs[0]).toMatchObject({ text: "季度经营复盘", bold: true, sizePt: 36 });
    const ellipse = pres.slides[0].shapes.find((s) => s.name === "椭圆");
    expect(ellipse?.geometry).toBe("ellipse");
    expect(ellipse?.fill).toBe("#FFD966");
    expect(ellipse?.lineColor).toBe("#BF8F00");
  });

  it("reads pictures, tables and notes", () => {
    const pic = pres.slides[0].shapes.find((s) => s.kind === "picture");
    expect(pic?.src).toContain("data:image/");
    const table = pres.slides[0].shapes.find((s) => s.kind === "table");
    expect(table?.rows?.flat().map((c) => c.text)).toEqual(["指标", "达成", "收入", "108%"]);
    expect(pres.slides[0].notes).toContain("口头补充");
  });

  it("reads bullets from a nested buChar element", () => {
    const body = pres.slides[0].shapes.find((s) => s.name === "正文");
    expect(body?.paragraphs?.map((p) => p.bullet)).toEqual(["•", "◦", undefined]);
    // `algn="r"` is normalized to the editor's own alignment words on read.
    expect(body?.paragraphs?.[2].align).toBe("right");
  });

  it("renders absolutely positioned HTML", () => {
    const html = renderSlideHtml(pres, 0);
    expect(html).toContain("position:absolute");
    expect(html).toContain("clip-path:ellipse");
    expect(html).toContain("季度经营复盘");
  });
});
