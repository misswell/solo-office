import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync, zipSync, strToU8 } from "fflate";
import { Spreadsheet, colWidthPx, rowHeightPx } from "../src/core/xlsx.js";
import { makeXlsx } from "./helpers/make.js";

/**
 * 表的默认行高与默认列宽在 `sheetFormatPr` 里：授权 73 份 xlsx / 80 张表**每张都写了**
 * `defaultRowHeight`，其中 13.5 磅 61 张、14.4 磅 15 张、15 磅 3 张、16.8 磅 1 张；`defaultColWidth`
 * 写 9 字符的 42 张。此前渲染固定按 15 磅 / 8.43 字符画，13.5 磅的表每行高 2px、9 字符的列每列窄 4px。
 * 这里钉住“表自己的默认值优先、没写才落 OOXML 默认”，并用语料逐表对表。
 */
const CORPUS = ".scratch/wps-corpus";
const dec = new TextDecoder();

function withFormat(defaults: string): Spreadsheet {
  const zip = unzipSync(makeXlsx());
  const path = "xl/worksheets/sheet1.xml";
  const xml = dec.decode(zip[path]!).replace("</sheetViews>", `</sheetViews><sheetFormatPr ${defaults}/>`);
  zip[path] = strToU8(xml);
  return Spreadsheet.openBytes(zipSync(zip, { level: 6 }));
}

describe("xlsx 表的默认行高/列宽（sheetFormatPr）", () => {
  it("默认值按表读；显式行高与列宽仍优先，隐藏列仍为 0", () => {
    const wb = withFormat('defaultRowHeight="13.5" defaultColWidth="9"');
    const sheet = wb.sheets[0]!;
    expect(sheet.defaultRowHeight).toBe(13.5);
    expect(sheet.defaultColWidth).toBe(9);
    expect(rowHeightPx(sheet, 0)).toBe(18);
    // sheet1 第 1 列宽 18 字符、第 2..4 列宽 13 字符，都压过默认 9。
    expect(colWidthPx(sheet, 0)).toBe(130);
    expect(colWidthPx(sheet, 3)).toBe(94);
    // 第二张表没写 sheetFormatPr：落 OOXML 默认 15 磅 / 8.43 字符。
    const plain = wb.sheets[1]!;
    expect(plain.defaultRowHeight).toBeUndefined();
    expect(rowHeightPx(plain, 0)).toBe(20);
    expect(colWidthPx(plain, 0)).toBe(61);
  });

  it.skipIf(!existsSync(CORPUS))("语料 80 张表的默认值逐表与原始 sheetFormatPr 一致", () => {
    const files = readdirSync(CORPUS).sort().filter((f) => /\.xlsx$/i.test(f)).map((f) => join(CORPUS, f));
    const rawRows = new Map<string, number>();
    let sheets = 0;
    let at135 = 0;
    let at9 = 0;
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(file));
      const zip = unzipSync(bytes);
      const wanted: number[] = [];
      const wantedCols: number[] = [];
      for (const path of Object.keys(zip).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))) {
        const xml = dec.decode(zip[path]!);
        const fmt = /<sheetFormatPr\b[^>]*>/.exec(xml)?.[0];
        expect(fmt, `${file.split("__").pop()} ${path} 缺 sheetFormatPr`).toBeDefined();
        const row = Number(/defaultRowHeight="([\d.]+)"/.exec(fmt!)?.[1]);
        expect(Number.isFinite(row), `${path} defaultRowHeight`).toBe(true);
        wanted.push(row);
        rawRows.set(String(row), (rawRows.get(String(row)) ?? 0) + 1);
        if (row === 13.5) at135++;
        const col = /defaultColWidth="([\d.]+)"/.exec(fmt!)?.[1];
        if (col) {
          wantedCols.push(Number(col));
          if (Number(col) === 9) at9++;
        }
        sheets++;
      }
      const wb = Spreadsheet.openBytes(bytes);
      expect(
        wb.sheets.map((s) => s.defaultRowHeight).sort((a, b) => a! - b!),
        `${file.split("__").pop()} 的默认行高`,
      ).toEqual(wanted.sort((a, b) => a - b));
      expect(
        wb.sheets.map((s) => s.defaultColWidth).filter((v) => v !== undefined).sort((a, b) => a! - b!),
        `${file.split("__").pop()} 的默认列宽`,
      ).toEqual(wantedCols.sort((a, b) => a - b));
      // 没有显式 ht 的那一行按表自己的默认值落像素。
      for (const sheet of wb.sheets) {
        const first = sheet.rowHeights.has(0);
        if (!first) expect(rowHeightPx(sheet, 0)).toBe(Math.round(((sheet.defaultRowHeight ?? 15) * 4) / 3));
      }
    }
    expect(sheets, "有 sheetFormatPr 的表").toBeGreaterThanOrEqual(80);
    expect(at135, "13.5 磅默认行高的表").toBeGreaterThanOrEqual(61);
    expect(at9, "9 字符默认列宽的表").toBeGreaterThanOrEqual(42);
    expect([...rawRows.keys()].sort()).toEqual(["13.5", "14.4", "15", "16.8"]);
  });
});
