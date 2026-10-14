import { describe, expect, it } from "vitest";
import { packPages, placementsAcrossSections, placementsOf, rowInPage, sheetOf, type LayoutUnit } from "../src/ui/doc-layout.js";
import { DEFAULT_SECTION, type Section } from "../src/core/docx.js";

/**
 * 分页的那一半纯的：`packPages` 不认识 DOM，只把一列单位往固定高度的纸上铺。行盒怎么量出来的（以及它的口径）
 * 由 `doc-layout.ts` 文件头与 `.scratch/pagination-dom.ts` 那道闸门钉，这里钉的是**铺**的规则本身——
 * 溢出、正好铺满、断点必另起一页、比一页还高的一枚只占一页。
 *
 * 单位高度一律用 px 的小数：A4（`w:pgSz` 16838 缇）在浏览器里是 1122.53px 高，减掉 `w:pgMar` 上下各 1440 缇
 * （96px）得到正文能用的 930.53px；页面 CSS 写的是 12pt／`line-height:1.62`，12pt = 16px，一行就是 25.92px。
 * 这两个数都在下面用着，不是随手取的整数（`.scratch/pagination-dom.ts` 量的正是它们）。
 */

const u = (height: number, forcePage = false): LayoutUnit => ({ height, forcePage });
const column = (n: number, height: number): LayoutUnit[] => Array.from({ length: n }, () => u(height));

const A4_BODY = 930.5333333333333;
const LINE = 25.92;

describe("packPages", () => {
  it("空的一列仍是一张纸", () => {
    expect(packPages([], A4_BODY)).toBe(1);
    expect(packPages([], 0)).toBe(1);
  });

  it("铺得下就一张，一张都放不下也还是一张", () => {
    expect(packPages(column(10, LINE), A4_BODY)).toBe(1);
    expect(packPages([u(2_000)], A4_BODY)).toBe(1);
  });

  it("整页能装 35 行 12pt／1.62，第 36 行挪到下一页", () => {
    // 35 × 25.92 = 907.2 ≤ 930.53，36 × 25.92 = 933.12 溢出。
    expect(packPages(column(35, LINE), A4_BODY)).toBe(1);
    expect(packPages(column(36, LINE), A4_BODY)).toBe(2);
    expect(packPages(column(70, LINE), A4_BODY)).toBe(2);
    expect(packPages(column(71, LINE), A4_BODY)).toBe(3);
  });

  it("小数加出来的那点误差不会多滚出一页", () => {
    // 0.1 + 0.1 + 0.1 在浮点里是 0.30000000000000004：纸高 0.3 的三行没有溢出。
    expect(packPages(column(3, 0.1), 0.3)).toBe(1);
    expect(packPages(column(4, 0.1), 0.3)).toBe(2);
  });

  it("断点必另起一页，哪怕这一页还空着", () => {
    // 第一段带 page-break-before：WPS 那侧也是先给一张空页再开始排。
    expect(packPages([u(LINE, true), u(LINE)], A4_BODY)).toBe(2);
    // 连按两次分页符：中间那一页是空的，但它是一页。
    expect(packPages([u(20), u(0, true), u(0, true), u(20)], A4_BODY)).toBe(3);
  });

  it("断点落在段尾也数得出那一页（后面没有内容的那一枚不占行）", () => {
    // 一枚 0 高的记号 = 纸上没有新的行盒，但另起了一页。
    const m = packPages([u(20), u(0, true)], A4_BODY);
    expect(m).toBe(2);
    expect(m).not.toBe(1);
  });

  it("同一处既断页又放不下时，只挪一次", () => {
    // 前两枚铺到 120 > 100 已经要翻页，第二枚又带着断点：翻页这件事在一处只发生一次。
    expect(packPages([u(60), u(60, true)], 100)).toBe(2);
    expect(packPages([u(60), u(60)], 100)).toBe(2);
  });

  it("比一页还高的那一枚只占一页，溢出去的不另算一页", () => {
    expect(packPages([u(1_000), u(10)], 100)).toBe(2);
    expect(packPages([u(1_000)], 100)).toBe(1);
  });

  it("负数与 0 高的普通单位不吃页", () => {
    expect(packPages([u(-500), u(0), u(A4_BODY)], A4_BODY)).toBe(1);
  });
});

/**
 * 逐枚落页那一份：`packPages` 报的是总数，状态栏的「第 X 页」要的是**某一枚**落在第几页，
 * 两者必须出自同一趟走法（`packPages` 就是取的它的最后一枚），所以边界条件一起钉。
 */
describe("placementsOf", () => {
  it("与 units 一一对应，最后一枚就是总页数", () => {
    const units = [u(50), u(50), u(50)];
    expect(placementsOf(units, 100)).toEqual([1, 1, 2]);
    expect(placementsOf(units, 100).at(-1)).toBe(packPages(units, 100));
  });

  it("空的一列没有落页，但页数仍是 1", () => {
    expect(placementsOf([], A4_BODY)).toEqual([]);
    expect(packPages([], A4_BODY)).toBe(1);
  });

  it("断点把后面那一枚抬到下一页，它后面接着排", () => {
    expect(placementsOf([u(50), u(0, true), u(50)], 100)).toEqual([1, 2, 2]);
  });

  it("比一页还高的一枚占一页，它后面那一枚从下一页开始", () => {
    expect(placementsOf([u(1_000), u(10)], 100)).toEqual([1, 2]);
  });

  it("量不出纸高时全落在第 1 页", () => {
    expect(placementsOf([u(50), u(50)], 0)).toEqual([1, 1]);
    expect(placementsOf([u(50), u(0, true)], -1)).toEqual([1, 1]);
  });

  it("35 行的边界上，第 35 枚在第 1 页、第 36 枚落到第 2 页", () => {
    const places = placementsOf(column(36, LINE), A4_BODY);
    expect(places[34]).toBe(1);
    expect(places[35]).toBe(2);
  });
});

/**
 * 一节自己的纸那一份：`sheetOf` 从节的声明算容量 = 页高 − 上下边距（缇 ÷ 15），与「第一节按纸上
 * 画出来的 padding 量」是两副来路——第一节的纸上还叠着编辑器的内边距，其余各节的纸**没画出来**，
 * 只能按声明算（`doc-layout.ts` 文件头；`.scratch/pagination-dom.ts` 第 13 节钉换纸的那条路）。
 */
describe("sheetOf", () => {
  it("A4 默认（16838 − 1440 − 1440 缇）就是 930.53px，与页面上量的正文高同一个数", () => {
    expect(sheetOf(DEFAULT_SECTION)?.contentHeightPx).toBeCloseTo(A4_BODY, 9);
  });

  it("节说了自己的纸与边距就按它算", () => {
    const short: Section = { ...DEFAULT_SECTION, pageHeightTwips: 6000 };
    expect(sheetOf(short)?.contentHeightPx).toBeCloseTo((6000 - 2880) / 15, 9);
    const roomy: Section = { ...DEFAULT_SECTION, margins: { ...DEFAULT_SECTION.margins, top: 720, bottom: 720 } };
    expect(sheetOf(roomy)?.contentHeightPx).toBeCloseTo((16838 - 1440) / 15, 9);
  });

  it("页高连上下边距都不够留的节不给纸，不加一个编出来的数", () => {
    expect(sheetOf({ ...DEFAULT_SECTION, pageHeightTwips: 2880 })).toBeUndefined();
    expect(sheetOf({ ...DEFAULT_SECTION, pageHeightTwips: 1000 })).toBeUndefined();
  });
});

/**
 * 多节铺排那一份：每一节内的铺法与单节逐枚相同（单节件就是一张纸的特例，下面有对拍），
 * 节界必另起一页（哪怕上一节那页还空着——断点是声明），从新节第一枚起换新节的容量。
 */
describe("placementsAcrossSections", () => {
  it("单节件与 placementsOf 逐枚相同（同一套铺法，只差几张纸）", () => {
    const units = [u(50), u(50), u(50), u(0, true), u(50)];
    const secs = units.map(() => 0);
    expect(placementsAcrossSections(units, secs, [{ contentHeightPx: 100 }])).toEqual(placementsOf(units, 100));
    expect(placementsAcrossSections(column(36, LINE), column(36, LINE).map(() => 0), [{ contentHeightPx: A4_BODY }])).toEqual(
      placementsOf(column(36, LINE), A4_BODY),
    );
  });

  it("节界必另起一页，哪怕上一节那页还空着", () => {
    const units = [u(10), u(10)];
    expect(placementsAcrossSections(units, [0, 1], [{ contentHeightPx: 100 }, { contentHeightPx: 100 }])).toEqual([1, 2]);
    // 上一节只留了一枚 0 高的记号：那一页还是被抬走了，断点是声明。
    expect(placementsAcrossSections([u(0), u(10)], [0, 1], [{ contentHeightPx: 100 }, { contentHeightPx: 100 }])).toEqual([1, 2]);
    // 流从第 1 页开始：第一节之前没有内容时「另起一页」没有可抬的上一页，不前空出一页。
    expect(placementsAcrossSections([u(10)], [1], [{ contentHeightPx: 100 }, { contentHeightPx: 100 }])).toEqual([1]);
  });

  it("新节的第一枚另起一页之后，与它同节的接着排", () => {
    const units = [u(10), u(10), u(10), u(10)];
    expect(placementsAcrossSections(units, [0, 0, 1, 1], [{ contentHeightPx: 100 }, { contentHeightPx: 100 }])).toEqual([1, 1, 2, 2]);
  });

  it("容量换着算：短纸上的两枚翻页，长纸上的同样两枚挤在一页", () => {
    const units = [u(40), u(40)];
    expect(placementsAcrossSections(units, [1, 1], [{ contentHeightPx: 100 }, { contentHeightPx: 50 }])).toEqual([1, 2]);
    expect(placementsAcrossSections(units, [0, 0], [{ contentHeightPx: 100 }, { contentHeightPx: 50 }])).toEqual([1, 1]);
  });

  it("比新节一张纸还高的一枚只占一页（按新节的容量钳）", () => {
    expect(placementsAcrossSections([u(1_000), u(10)], [1, 1], [{ contentHeightPx: 100 }, { contentHeightPx: 50 }])).toEqual([1, 2]);
  });

  it("节界与断点落在同一枚上时两次都发生（连按分页符一样的叠法）", () => {
    expect(placementsAcrossSections([u(10), u(10, true)], [0, 1], [{ contentHeightPx: 100 }, { contentHeightPx: 100 }])).toEqual([1, 3]);
  });

  it("空的一列没有落页", () => {
    expect(placementsAcrossSections([], [], [{ contentHeightPx: 100 }])).toEqual([]);
  });
});

/**
 * 行号那一份纯的：`rowInPage` 拿 `placementsOf` 的答案在同一页里数到第几枚。口径是**页内重数**——
 * 一页的第一枚（有内容的）就是那一页的第 1 行，0 高的断点记号不占行。真机是否逐页重数**未取证**
 * （状态栏的闸门 `.scratch/pagination-dom.ts` 第 11 节按这条口径钉）。
 */
describe("rowInPage", () => {
  it("第 1 页里顺着数", () => {
    const units = [u(10), u(10), u(10), u(10)];
    const places = placementsOf(units, 25);
    expect(places).toEqual([1, 1, 2, 2]);
    expect(rowInPage(places, units, 0)).toBe(1);
    expect(rowInPage(places, units, 1)).toBe(2);
  });

  it("换页归 1：第 2 页第一枚是那一页的第 1 行，不跟着上一页连数", () => {
    const units = [u(10), u(10), u(10), u(10)];
    const places = placementsOf(units, 25);
    expect(rowInPage(places, units, 2)).toBe(1);
    expect(rowInPage(places, units, 3)).toBe(2);
  });

  it("0 高的断点记号不占行", () => {
    const units = [u(10), u(0, true), u(10)];
    const places = placementsOf(units, 100);
    expect(places).toEqual([1, 2, 2]);
    expect(rowInPage(places, units, 2)).toBe(1);
  });

  it("取点走到记号上（防呆路径）算第 1 行，不报 0", () => {
    const units = [u(10), u(0, true)];
    const places = placementsOf(units, 100);
    expect(rowInPage(places, units, 1)).toBe(1);
  });

  it("越界的序号不猜", () => {
    expect(rowInPage([1], [u(10)], -1)).toBeUndefined();
    expect(rowInPage([1], [u(10)], 1)).toBeUndefined();
    expect(rowInPage([], [], 0)).toBeUndefined();
  });
});
