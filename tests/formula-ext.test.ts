import { beforeEach, describe, expect, it } from "vitest";
import { Evaluator, parse, type Value } from "../src/core/formula.js";

/**
 * Extension-function tests. Every expected value comes from an INDEPENDENT
 * reference, never from the implementation under test:
 *  - Python reference implementation computed from the textbook definitions
 *    (cross-checked against the published Microsoft Docs worked examples noted
 *    inline, e.g. PMT(0.08/12,10,10000) = -1037.03);
 *  - small hand-checked datasets for the *_IFS family.
 * See .scratch/rea-notes.md §8 for the WPS-side KF_* evidence.
 */

const grid = new Map<string, Value>();
const setCell = (ref: string, v: Value): void => {
  grid.set(`Sheet1!${ref}`, v);
};
const clearGrid = (): void => grid.clear();

const stub = {
  constant: (sheet: string, col: number, row: number): Value => grid.get(`${sheet}!${String.fromCharCode(65 + col)}${row + 1}`) ?? null,
  formula: (): string | undefined => undefined,
  name: (): undefined => undefined,
};

const evalFormula = (src: string): Value => {
  const ev = new Evaluator(stub, "Sheet1");
  return ev.sub(parse(src), "Sheet1");
};

const close = (a: Value, b: number, digits = 6): void => {
  expect(typeof a === "number" && Number.isFinite(a)).toBe(true);
  expect(a as number).toBeCloseTo(b, digits);
};

// Fixed dataset used across the *_IFS cases; the expected numbers are
// verifiable by inspection.
function seedIfs(): void {
  clearGrid();
  //        A(区域)   B(季度)  C(金额)
  setCell("A1", "东");
  setCell("B1", 1);
  setCell("C1", 100);
  setCell("A2", "南");
  setCell("B2", 1);
  setCell("C2", 200);
  setCell("A3", "东");
  setCell("B3", 2);
  setCell("C3", 300);
  setCell("A4", "西");
  setCell("B4", 2);
  setCell("C4", 400);
  setCell("A5", "东");
  setCell("B5", 3);
  setCell("C5", "文本"); // 文本金额：SUMIFS/AVERAGEIFS 必须跳过
}

describe("statistics (population spread, shape, regression)", () => {
  it("population spread matches the Python reference", () => {
    close(evalFormula("STDEV.P(1,2,3,4,5)"), 1.4142135623730951);
    close(evalFormula("STDEVP(1,2,3,4,5)"), 1.4142135623730951);
    close(evalFormula("VAR.P(1,2,3,4,5)"), 2);
    close(evalFormula("VARP(1,2,3,4,5)"), 2);
    // Range reference over cells, blanks and text skipped like Excel.
    setCell("H1", 1);
    setCell("H2", 2);
    setCell("H3", 3);
    setCell("H4", 4);
    setCell("H5", 5);
    setCell("H6", "文本");
    setCell("H7", "");
    close(evalFormula("STDEV.P(H1:H7)"), 1.4142135623730951);
  });

  it("shape statistics match Microsoft Docs examples", () => {
    close(evalFormula("GEOMEAN(4,5,8,7,11,4,3)"), 5.476986969656962);
    close(evalFormula("HARMEAN(4,5,8,7,11,4,3)"), 5.028375962061728);
    close(evalFormula("AVEDEV(4,5,6,7,5,4,3)"), 1.0204081632653061);
    close(evalFormula("DEVSQ(4,5,8,7,11,4,3)"), 48);
    close(evalFormula("MODE(5,3,2,3,2,1,4)"), 2);
    // Ties resolve to the smallest value.
    close(evalFormula("MODE(3,3,2,2,1)"), 2);
    // No repeat → #N/A.
    expect(evalFormula("MODE(1,2,3)")).toBe("#N/A");
    // Non-positive input → #NUM!.
    expect(evalFormula("GEOMEAN(4,0,3)")).toBe("#NUM!");
    expect(evalFormula("HARMEAN(-1,2,3)")).toBe("#NUM!");
  });

  it("percentile and quartile interpolate like Excel (INC)", () => {
    clearGrid();
    [1, 2, 3, 4, 5].forEach((v, i) => setCell(`A${i + 1}`, v));
    close(evalFormula("PERCENTILE(A1:A5,0.3)"), 2.2);
    close(evalFormula("PERCENTILE.INC(A1:A5,0.3)"), 2.2);
    close(evalFormula("QUARTILE(A1:A5,1)"), 2);
    close(evalFormula("QUARTILE.INC(A1:A5,3)"), 4);
    close(evalFormula("PERCENTILE(A1:A5,0)"), 1);
    close(evalFormula("PERCENTILE(A1:A5,1)"), 5);
    expect(evalFormula("PERCENTILE(A1:A5,1.5)")).toBe("#NUM!");
    expect(evalFormula("QUARTILE(A1:A5,5)")).toBe("#NUM!");
  });

  it("regression measures match the Docs datasets", () => {
    const set = (): void => {
      clearGrid();
      // y = A2:A8, x = B2:B8 — the Microsoft SLOPE example (0.305556).
      const ys = [2, 3, 9, 1, 8, 7, 5];
      const xs = [6, 5, 11, 7, 5, 4, 4];
      ys.forEach((v, i) => setCell(`A${i + 2}`, v));
      xs.forEach((v, i) => setCell(`B${i + 2}`, v));
    };
    set();
    close(evalFormula("SLOPE(A2:A8,B2:B8)"), 0.3055555555555556);
    close(evalFormula("INTERCEPT(A2:A8,B2:B8)"), 3.1666666666666665);
    close(evalFormula("RSQ(A2:A8,B2:B8)"), 0.05795019157088123);
    close(evalFormula("CORREL(A2:A8,B2:B8)"), 0.24072846861827315);
    close(evalFormula("PEARSON(A2:A8,B2:B8)"), 0.24072846861827315);
    close(evalFormula("COVARIANCE.P(A2:A8,B2:B8)"), 1.5714285714285714);
    close(evalFormula("COVARIANCE.S(A2:A8,B2:B8)"), 1.8333333333333333);
    close(evalFormula("COVAR(A2:A8,B2:B8)"), 1.5714285714285714);
    // The CORREL Docs example (0.997054) on its own dataset.
    clearGrid();
    ["3", "2", "4", "5", "6"].forEach((v, i) => setCell(`A${i + 2}`, Number(v)));
    ["9", "7", "12", "15", "17"].forEach((v, i) => setCell(`B${i + 2}`, Number(v)));
    close(evalFormula("CORREL(A2:A6,B2:B6)"), 0.9970544855015815);
    close(evalFormula("COVAR(A2:A6,B2:B6)"), 5.2);
    // Zero variance in the x range → #DIV/0!; constant y just gives slope 0.
    clearGrid();
    setCell("A1", 2);
    setCell("A2", 2);
    setCell("B1", 1);
    setCell("B2", 5);
    close(evalFormula("SLOPE(A1:A2,B1:B2)"), 0);
    expect(evalFormula("SLOPE(B1:B2,A1:A2)")).toBe("#DIV/0!");
    expect(evalFormula("RSQ(A1:A2,B1:B2)")).toBe("#DIV/0!");
  });

  it("AVERAGEA counts text as zero and booleans as one", () => {
    clearGrid();
    setCell("A1", "文字");
    setCell("A2", 1);
    setCell("A3", true);
    setCell("A4", 3);
    close(evalFormula("AVERAGEA(A1:A4)"), 1.25);
  });

  it("RANK.EQ and RANK.AVG handle order and ties", () => {
    clearGrid();
    [10, 20, 20, 40].forEach((v, i) => setCell(`A${i + 1}`, v));
    expect(evalFormula("RANK.EQ(20,A1:A4)")).toBe(2); // descending: 40 then the 20s
    close(evalFormula("RANK.AVG(20,A1:A4)"), 2.5);
    expect(evalFormula("RANK.EQ(20,A1:A4,1)")).toBe(2); // ascending: 10 then the 20s
    expect(evalFormula("RANK.EQ(99,A1:A4)")).toBe("#N/A");
  });
});

describe("paired-criteria plurals", () => {
  beforeEach(seedIfs);

  it("COUNTIFS across two criteria ranges", () => {
    expect(evalFormula('COUNTIFS(A1:A5,"东",B1:B5,2)')).toBe(1);
    expect(evalFormula('COUNTIFS(A1:A5,"东")')).toBe(3);
    expect(evalFormula('COUNTIFS(B1:B5,">1")')).toBe(3);
  });

  it("SUMIFS skips text amounts and matches conjunctions", () => {
    close(evalFormula('SUMIFS(C1:C5,A1:A5,"东")'), 400);
    close(evalFormula('SUMIFS(C1:C5,A1:A5,"东",B1:B5,2)'), 300);
    expect(evalFormula('SUMIFS(C1:C5,A1:A5,"不存在")')).toBe(0);
  });

  it("AVERAGEIFS/MAXIFS/MINIFS with empty results", () => {
    close(evalFormula('AVERAGEIFS(C1:C5,A1:A5,"东")'), 200);
    expect(evalFormula('AVERAGEIFS(C1:C5,A1:A5,"不存在")')).toBe("#DIV/0!");
    close(evalFormula('MAXIFS(C1:C5,A1:A5,"东")'), 300);
    close(evalFormula('MINIFS(C1:C5,B1:B5,1)'), 100);
    expect(evalFormula('MAXIFS(C1:C5,A1:A5,"不存在")')).toBe(0);
  });

  it("mismatched range sizes are #VALUE!", () => {
    expect(evalFormula('COUNTIFS(A1:A5,"东",A1:A2,"东")')).toBe("#VALUE!");
    expect(evalFormula('SUMIFS(C1:C5,A1:A2,"东")')).toBe("#VALUE!");
  });
});

describe("math and information extensions", () => {
  it("exponential and logarithm families with domain guards", () => {
    close(evalFormula("EXP(1)"), Math.E);
    close(evalFormula("LN(EXP(2))"), 2);
    close(evalFormula("LOG(8,2)"), 3);
    close(evalFormula("LOG(100)"), 2);
    close(evalFormula("LOG10(1000)"), 3);
    expect(evalFormula("LN(0)")).toBe("#NUM!");
    expect(evalFormula("LOG(8,1)")).toBe("#DIV/0!");
    expect(evalFormula("LOG10(-1)")).toBe("#NUM!");
  });

  it("multiple rounding follows the sign rules", () => {
    close(evalFormula("CEILING(2.5,1)"), 3);
    close(evalFormula("CEILING(-2.5,2)"), -2);
    close(evalFormula("CEILING(-2.5,-2)"), -4);
    expect(evalFormula("CEILING(2.5,-2)")).toBe("#NUM!");
    close(evalFormula("FLOOR(2.5,1)"), 2);
    close(evalFormula("FLOOR(-2.5,1)"), -3);
    expect(evalFormula("FLOOR(2.5,-1)")).toBe("#NUM!");
    close(evalFormula("MROUND(10,3)"), 9);
    close(evalFormula("MROUND(-10,-3)"), -9);
    expect(evalFormula("MROUND(10,-3)")).toBe("#NUM!");
  });

  it("parity and the #N/A family", () => {
    expect(evalFormula("ISEVEN(4)")).toBe(true);
    expect(evalFormula("ISODD(4)")).toBe(false);
    expect(evalFormula("ISODD(3.9)")).toBe(true);
    expect(evalFormula("IFNA(NA(),\"替代\")")).toBe("替代");
    expect(evalFormula("IFNA(7,\"替代\")")).toBe(7);
    expect(evalFormula("ISNA(NA())")).toBe(true);
    expect(evalFormula("ISNA(1)")).toBe(false);
  });
});

describe("financial time-value functions (Microsoft Docs references)", () => {
  it("PMT/FV/PV/NPER/RATE against published examples", () => {
    close(evalFormula("PMT(0.08/12,10,10000)"), -1037.0320893591636);
    close(evalFormula("PMT(0.08/12,120,10000)"), -121.32759435535777);
    close(evalFormula("PMT(0.08/12,10,10000,0,1)"), -1030.16, 2);
    close(evalFormula("FV(0.06/12,10,-200,-500,1)"), 2581.4033740601362, 4);
    close(evalFormula("PV(0.08/12,240,500)"), -59777.14585118777);
    close(evalFormula("NPER(0.12/12,-100,-1000,10000,1)"), 59.67386567429457);
    close(evalFormula("RATE(48,-200,8000)*12"), 0.09241766985842485, 6);
  });

  it("NPV/IRR/MIRR against published examples", () => {
    close(evalFormula("NPV(0.1,-10000,3000,4200,6800)"), 1188.4434123352212);
    // IRR/MIRR take the cashflow as one range, like a real spreadsheet row.
    setCell("J1", -70000);
    [12000, 15000, 18000, 21000, 26000].forEach((v, i) => setCell(`J${i + 2}`, v));
    close(evalFormula("IRR(J1:J6)"), 0.0866309480365316, 6);
    setCell("J1", -70000);
    setCell("J2", 12000);
    setCell("J3", 15000);
    close(evalFormula("IRR(J1:J3)"), -0.44350694133474045, 6);
    setCell("J1", -120000);
    [39000, 30000, 21000, 37000, 46000].forEach((v, i) => setCell(`J${i + 2}`, v));
    close(evalFormula("MIRR(J1:J6,0.1,0.12)"), 0.1260941303659051, 6);
  });

  it("IPMT/PPMT split the payment and preserve the identity", () => {
    close(evalFormula("IPMT(0.1,3,3,8000)"), -292.4471299093658, 4);
    close(evalFormula("IPMT(0.1,1,3,8000)"), -800, 6);
    close(evalFormula("PPMT(0.1,3,3,8000)"), -2924.4712990936528, 4);
    close(evalFormula("PPMT(0.1,1,3,8000)"), -2416.9184290030184, 4);
    close(evalFormula("IPMT(0.1,3,3,8000)+PPMT(0.1,3,3,8000)"), evalFormula("PMT(0.1,3,8000)") as number, 8);
    // period outside [1, nper] → #NUM!
    expect(evalFormula("IPMT(0.1,4,3,8000)")).toBe("#NUM!");
    expect(evalFormula("PPMT(0.1,0,3,8000)")).toBe("#NUM!");
  });

  it("CUMIPMT/CUMPRINC against the Docs loan example", () => {
    close(evalFormula("CUMIPMT(0.09/12,360,125000,13,24,0)"), -11135.232130750845, 3);
    close(evalFormula("CUMPRINC(0.09/12,360,125000,13,24,0)"), -934.1071234208764, 3);
    close(evalFormula("CUMIPMT(0.09/12,360,125000,1,1,0)"), -937.5, 6);
    expect(evalFormula("CUMIPMT(0.09/12,360,125000,0,24,0)")).toBe("#NUM!");
    expect(evalFormula("CUMIPMT(0.09/12,360,125000,24,13,0)")).toBe("#NUM!");
    expect(evalFormula("CUMIPMT(-0.09/12,360,125000,13,24,0)")).toBe("#NUM!");
  });

  it("depreciation schedules", () => {
    close(evalFormula("SLN(30000,7500,10)"), 2250);
    expect(evalFormula("SLN(30000,7500,0)")).toBe("#DIV/0!");
    close(evalFormula("SYD(30000,7500,10,1)"), 4090.909090909091);
    close(evalFormula("SYD(30000,7500,10,10)"), 409.09090909090907);
    expect(evalFormula("SYD(30000,7500,10,11)")).toBe("#NUM!");
    close(evalFormula("DDB(2400,300,3650,1)"), 1.3150684931506849, 4);
    close(evalFormula("DDB(2400,300,120,1)"), 40);
    close(evalFormula("DDB(2400,300,10,1)"), 480);
  });

  it("convergence failures surface as #NUM!", () => {
    // IRR: no sign change in the cashflow — no root to find.
    clearGrid();
    [5, 6, 7].forEach((v, i) => setCell(`A${i + 1}`, v));
    expect(evalFormula("IRR(A1:A3)")).toBe("#NUM!");
    // RATE: pv+fv=0 with zero payments — the annuity equation has no root.
    expect(evalFormula("RATE(10,0,-1000,0)")).toBe("#NUM!");
  });
});

describe("argument semantics shared by the new library", () => {
  it("blank and text cells are ignored by statistical ranges", () => {
    clearGrid();
    setCell("A1", 1);
    setCell("A2", "");
    setCell("A3", "文本");
    setCell("A4", 4);
    close(evalFormula("STDEV.P(A1:A4)"), 1.5);
    close(evalFormula("VAR.P(A1:A4)"), 2.25);
    close(evalFormula("GEOMEAN(A1:A4)"), 2);
    close(evalFormula("AVEDEV(A1:A4)"), 1.5);
    close(evalFormula("DEVSQ(A1:A4)"), 4.5);
    close(evalFormula("PERCENTILE(A1:A4,0.5)"), 2.5);
  });

  it("errors inside ranges propagate", () => {
    clearGrid();
    setCell("A1", 1);
    setCell("A2", "#N/A");
    setCell("A3", 3);
    expect(evalFormula("STDEV.P(A1:A3)")).toBe("#N/A");
    expect(evalFormula("CORREL(A1:A3,A1:A3)")).toBe("#N/A");
    expect(evalFormula("IRR(A1:A3)")).toBe("#N/A");
    // SUMIFS: a matched cell in the target range propagates; the criteria
    // range treats error cells as no-match.
    clearGrid();
    setCell("A1", 1);
    setCell("A2", "#N/A");
    setCell("A3", 3);
    setCell("B1", 1);
    setCell("B2", 1);
    setCell("B3", 1);
    expect(evalFormula('SUMIFS(A1:A3,B1:B3,">0")')).toBe("#N/A");
    setCell("B2", "#N/A");
    close(evalFormula('SUMIFS(A1:A3,B1:B3,">0")'), 4);
  });

  it("range orientation does not matter and single cells still coerce", () => {
    clearGrid();
    setCell("A1", 3);
    setCell("B1", 4);
    close(evalFormula("STDEV.P(A1:B1)"), 0.5);
    close(evalFormula("EXP(A1)"), Math.exp(3));
    close(evalFormula("MROUND(B1,3)"), 3);
  });

  it("existing core functions still evaluate after the registry merge", () => {
    clearGrid();
    setCell("A1", 1);
    setCell("A2", 2);
    setCell("A3", 3);
    close(evalFormula("SUM(A1:A3)"), 6);
    expect(evalFormula("IF(SUM(A1:A3)>5,\"大\",\"小\")")).toBe("大");
    close(evalFormula("ROUND(2.6,0)"), 3);
    expect(evalFormula("COUNTIF(A1:A3,\">1\")")).toBe(2);
    close(evalFormula("MEDIAN(A1:A3)"), 2);
    expect(evalFormula("UNKNOWNFN(1)")).toBe("#NAME?");
  });
});
