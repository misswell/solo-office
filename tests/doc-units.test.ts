import { describe, expect, it } from "vitest";
import { UNITS, twipText, twipsFromUnit } from "../src/ui/doc-units.js";

/**
 * The edit axis of `doc-units.ts`: every whole twip count 0..60000 must survive a trip through a box in any
 * of the four units — the dialog prefills with these decimals and parses the same text back, so an identity
 * here is exactly "a dialog nobody retouched cannot move a margin or an edge". The inequality the decimals
 * come from is spelled in the module's header; this file is the traversal it points at.
 */
describe("docx 测量单位：编辑轴的缇往返", () => {
  it("0..60000 每个整数缇在四种单位下写进格子再读回来都是它自己", () => {
    const wrong: string[] = [];
    for (const u of UNITS.map((x) => x.id)) {
      for (let tw = 0; tw <= 60000; tw++) {
        const back = twipsFromUnit(Number(twipText(tw, u)), u);
        if (back !== tw) wrong.push(`${u} ${tw} → ${twipText(tw, u)} → ${back}`);
      }
    }
    expect(wrong.slice(0, 5)).toEqual([]);
  });

  it("一句话口径的样本：同一批缇数在四家单位下的字面", () => {
    // 1440 缇 = 2.54 厘米 = 25.4 毫米 = 1 英寸 = 72 磅，是四家换算互相咬合的那一点。
    expect(twipText(1440, "cm")).toBe("2.54");
    expect(twipText(1440, "mm")).toBe("25.4");
    expect(twipText(1440, "in")).toBe("1");
    expect(twipText(1440, "pt")).toBe("72");
    // 851 缇就是页面设置头注里那个数：三位小数能把它原样带回，两位不能。
    expect(twipText(851, "cm")).toBe("1.501");
    expect(twipsFromUnit(1.5, "cm")).toBe(850);
    expect(twipsFromUnit(Number(twipText(851, "cm")), "cm")).toBe(851);
    // 英寸下同一个数要四位：0.59 会回成 850。
    expect(twipText(851, "in")).toBe("0.591");
    expect(twipsFromUnit(0.59, "in")).toBe(850);
    // A4 的两条边在四家单位下的精确字面（11906 = 21.001 厘米，没有更短的写法回到它）。
    expect(twipText(11906, "cm")).toBe("21.001");
    expect(twipText(11906, "mm")).toBe("210.01");
    expect(twipText(11906, "in")).toBe("8.2681");
    expect(twipText(11906, "pt")).toBe("595.3");
    expect(twipText(16838, "in")).toBe("11.6931");
    // 信纸在英寸下正好是它名字上的那两个数。
    expect(twipText(12240, "in")).toBe("8.5");
    expect(twipText(15840, "in")).toBe("11");
  });

  it("尾零不留：整数的字面就是整数，两次往返不会长胖", () => {
    for (const u of UNITS.map((x) => x.id)) {
      for (const tw of [0, 20, 283, 1440, 23811, 60000]) {
        const once = twipText(tw, u);
        expect(once).toBe(String(Number(once)));
        expect(twipText(twipsFromUnit(Number(once), u), u)).toBe(once);
      }
    }
  });

  it("默认单位是厘米，与页面设置原来的口径同轴", () => {
    expect(twipText(851)).toBe("1.501");
    expect(twipsFromUnit(1.501)).toBe(851);
  });
});
