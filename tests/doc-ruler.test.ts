import { describe, expect, it } from "vitest";
import { pxPerUnit } from "../src/ui/doc-units.js";
import { rulerTicks } from "../src/ui/doc-ruler.js";

/**
 * 标尺刻度是纯算式：`0` 在正文栏左沿，数字档按 1/2/5×10^k 挑到排得下为止，细分四等分。
 * 这里钉的是「一整条正文栏上量出来的东西」——数字不叠、刻度不越过右沿、换单位只换口径不换纸。
 * 纸面与标记（DOM、`zoom`、`getBoundingClientRect`）在 `.scratch/ruler-dom.ts` 与装机烟测里量。
 */
const A4_TEXT = 793.7 - 96 - 96; // A4 宽减 WPS 默认左右边距（1 英寸），纸的像素

describe("doc-ruler 刻度", () => {
  it("厘米：数字每 1 厘米一枚，细分到 0.25 厘米", () => {
    const ticks = rulerTicks(A4_TEXT, "cm");
    const majors = ticks.filter((t) => t.major);
    expect(majors[0]!.x).toBe(0);
    expect(majors.map((t) => t.label)).toEqual(["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15"]);
    // 相邻主刻度就是一个单位（折成纸面像素）。
    expect(majors[1]!.x).toBeCloseTo(pxPerUnit("cm"), 6);
    // 主刻度之间三枚细分。
    expect(ticks.filter((t) => !t.major && t.x < pxPerUnit("cm")).length).toBe(3);
    // 一枚都不越过正文栏右沿。
    expect(Math.max(...ticks.map((t) => t.x))).toBeLessThanOrEqual(A4_TEXT + 0.001);
  });

  it("毫米：1 毫米排不下数字，档位升到 10 毫米（数字仍按毫米报）", () => {
    const ticks = rulerTicks(A4_TEXT, "mm");
    const majors = ticks.filter((t) => t.major);
    expect(majors.slice(0, 3).map((t) => t.label)).toEqual(["0", "10", "20"]);
    expect(majors[1]!.x).toBeCloseTo(pxPerUnit("mm") * 10, 6);
    expect(pxPerUnit("mm") * 2.5).toBeGreaterThanOrEqual(4);
  });

  it("英寸：数字每一英寸一枚，细分到四分之一英寸", () => {
    const ticks = rulerTicks(A4_TEXT, "in");
    const majors = ticks.filter((t) => t.major);
    expect(majors.map((t) => t.label)).toEqual(["0", "1", "2", "3", "4", "5", "6"]);
    expect(majors[1]!.x).toBe(96);
    expect(ticks.filter((t) => !t.major && t.x < 96).length).toBe(3);
  });

  it("磅：1 磅太窄，档位一路升到 50 磅", () => {
    const ticks = rulerTicks(A4_TEXT, "pt");
    const majors = ticks.filter((t) => t.major);
    expect(majors[0]!.label).toBe("0");
    expect(majors[1]!.label).toBe("50");
    expect(majors[1]!.x).toBeCloseTo(pxPerUnit("pt") * 50, 6);
  });

  it("同一把尺换单位只换数字口径，正文栏右沿不变", () => {
    for (const unit of ["cm", "mm", "in", "pt"] as const) {
      const ticks = rulerTicks(A4_TEXT, unit);
      expect(ticks.length).toBeGreaterThan(2);
      expect(ticks[0]!.x).toBe(0);
      expect(Math.max(...ticks.map((t) => t.x))).toBeLessThanOrEqual(A4_TEXT + 0.001);
      // 主刻度逐年递增，数字不重复。
      const labels = ticks.filter((t) => t.major).map((t) => Number(t.label));
      for (let i = 1; i < labels.length; i += 1) expect(labels[i]!).toBeGreaterThan(labels[i - 1]!);
    }
  });

  it("量不出正文栏时一枚都不画", () => {
    expect(rulerTicks(0, "cm")).toEqual([]);
    expect(rulerTicks(-5, "cm")).toEqual([]);
    expect(rulerTicks(Number.NaN, "cm")).toEqual([]);
  });
});
