import { describe, expect, it } from "vitest";
import {
  bandYieldPx,
  lineCharges,
  packPages,
  placementsAcrossSections,
  placementsOf,
  rowInPage,
  sheetOf,
  splitRowBands,
  type LayoutUnit,
} from "../src/ui/doc-layout.js";
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

describe("keepNext / keepLines：文件声明的两条保页规则", () => {
  it("keepNext：段尾那一枚和下一块的第一枚同页——差一枚装不下时整枚挪下去", () => {
    const units: LayoutUnit[] = [
      { height: 40, forcePage: false },
      { height: 40, forcePage: false },
      { height: 20, forcePage: false, keepNext: true },
      { height: 50, forcePage: false },
    ];
    expect(placementsOf(units, 100)).toEqual([1, 1, 2, 2]);
    const loose: LayoutUnit[] = units.map((u) => ({ height: u.height, forcePage: u.forcePage }));
    expect(placementsOf(loose, 100)).toEqual([1, 1, 1, 2]);
  });
  it("keepLines：整段一坨——一页装不下整段就整段挪到下一页起头", () => {
    const units: LayoutUnit[] = [
      { height: 60, forcePage: false },
      { height: 20, forcePage: false, keepLines: 60 },
      { height: 40, forcePage: false },
    ];
    expect(placementsOf(units, 100)).toEqual([1, 2, 2]);
    const loose: LayoutUnit[] = units.map((u) => ({ height: u.height, forcePage: u.forcePage }));
    expect(placementsOf(loose, 100)).toEqual([1, 1, 2]);
  });
  it("keepLines 比一页还高时这条不算数：与没写同铺法（整段放哪都在拆）", () => {
    const tall: LayoutUnit[] = [
      { height: 60, forcePage: false },
      { height: 80, forcePage: false, keepLines: 180 },
      { height: 50, forcePage: false },
    ];
    const loose: LayoutUnit[] = tall.map((u) => ({ height: u.height, forcePage: u.forcePage }));
    expect(placementsOf(tall, 100)).toEqual(placementsOf(loose, 100));
    expect(placementsOf(tall, 100)).toEqual([1, 2, 3]);
  });
  it("两条叠一起：标题 keepNext + 下一段 keepLines——标题跟着整段走，不孤零零留在页底", () => {
    const units: LayoutUnit[] = [
      { height: 30, forcePage: false },
      { height: 20, forcePage: false, keepNext: true },
      { height: 60, forcePage: false, keepLines: 60 },
      { height: 40, forcePage: false },
    ];
    expect(placementsOf(units, 100)).toEqual([1, 2, 2, 3]);
    const loose: LayoutUnit[] = units.map((u) => ({ height: u.height, forcePage: u.forcePage }));
    expect(placementsOf(loose, 100)).toEqual([1, 1, 2, 2]);
  });
  it("keepNext 遇上声明断点：下一枚自己另起一页时链不成立，各走各的", () => {
    const units: LayoutUnit[] = [
      { height: 90, forcePage: false },
      { height: 10, forcePage: false, keepNext: true },
      { height: 10, forcePage: true },
    ];
    expect(placementsOf(units, 100)).toEqual([1, 1, 2]);
  });
});

describe("widowControl：缺省即开的孤行控制", () => {
  /** 一段 `n` 枚行盒，每枚带上它在段里的枚序（`unitBlocksOf` 在段落说「没关」时盖的就是这一族）。 */
  const run = (height: number, n: number): LayoutUnit[] =>
    Array.from({ length: n }, (_, index) => ({ height, forcePage: false, widowLine: { index, total: n } }));
  const loose = (units: readonly LayoutUnit[]): LayoutUnit[] => units.map((u) => ({ height: u.height, forcePage: u.forcePage }));

  it("页尾只剩段首一行：那一行也挪下去（整段起页），比没写时多带一枚", () => {
    const units: LayoutUnit[] = [{ height: 60, forcePage: false }, ...run(30, 3)];
    expect(placementsOf(units, 100)).toEqual([1, 2, 2, 2]);
    expect(placementsOf(loose(units), 100)).toEqual([1, 1, 2, 2]);
  });
  it("页头只剩段尾一行：把它前面那一行一起带下去（新页至少两行）", () => {
    const units: LayoutUnit[] = [{ height: 30, forcePage: false }, ...run(30, 3)];
    expect(placementsOf(units, 100)).toEqual([1, 1, 2, 2]);
    expect(placementsOf(loose(units), 100)).toEqual([1, 1, 1, 2]);
  });
  it("两行的段：两条修正同时命中，等价于整段不拆（keepLines 的两行特例）", () => {
    const units: LayoutUnit[] = [{ height: 60, forcePage: false }, ...run(30, 2)];
    expect(placementsOf(units, 100)).toEqual([1, 2, 2]);
    expect(placementsOf(loose(units), 100)).toEqual([1, 1, 2]);
  });
  it("页界本来就落在段首（枚序 0）时不动：不为此多挪一枚、也不多算一页", () => {
    const units: LayoutUnit[] = [{ height: 80, forcePage: false }, ...run(30, 3)];
    expect(placementsOf(units, 100)).toEqual(placementsOf(loose(units), 100));
    expect(placementsOf(units, 100)).toEqual([1, 2, 2, 2]);
  });
  it("往前挪不越过声明的断点：段首那枚自己 `forcePage` 时页界由文件说死，让给它", () => {
    const units: LayoutUnit[] = [
      { height: 50, forcePage: false },
      { height: 30, forcePage: true, widowLine: { index: 0, total: 2 } },
      { height: 30, forcePage: false, widowLine: { index: 1, total: 2 } },
    ];
    expect(placementsOf(units, 55)).toEqual([1, 2, 3]);
  });
});

/**
 * 一枚行盒「占多高」这条口径的持久件。DOM 那侧由 `.scratch/pagination-dom.ts` 与 `.scratch/grid-dom.ts`
 * 量真纸（闸门不进包，只护得住本地）；这里护的是**规则本身**，node 就能跑。
 *
 * 数是从闸门抄来的一组模型读数，不是随手整数：
 * - 312 缇的网格（`w:docGrid w:linePitch="312"` = 20.8px）上一格 20.8，段的块上沿 206、第一行墨顶 205
 *   （字自己洒进上一格 1px），三行的段块下沿 268.4——`para-gap-dom` 的 `原始账` 与 `grid-dom` 的格宽读数。
 * - 11 磅中文行不吸网格时：声明 `line-height` 23.76px，块自己的推进 24.77px（`lh-drift-dom`）。
 * 钉的是「按**流里的推进** charging」这一条：起点取 `from`、末枚收到 `to − from − 前面各枚`。
 * 按碎片（墨水）边沿 charging 时每一枚多 1px，一页 44 格只数得出 42——第 132 轮的负控把这条亲手翻出来过：
 * 那一版 `tests/` 883 条全绿，只有 DOM 闸门变红，所以规则必须在这里也钉一枚。
 */
describe("lineCharges：一枚行盒占多高，问它在纸上推进了多少", () => {
  const PITCH = 20.8; // 312 缇
  const grid = [
    { top: 205, bottom: 226.8 },
    { top: 225.8, bottom: 247.6 },
    { top: 246.6, bottom: 268.4 },
  ];

  it("网格上三行：每枚正好一格，枚高之和 == 块画出来的推进", () => {
    const c = lineCharges(grid, PITCH, 206, 268.4);
    expect(c).toEqual([PITCH, PITCH, PITCH]);
    expect(c.reduce((a, b) => a + b, 0)).toBeCloseTo(268.4 - 206, 10);
  });

  it("墨顶比格顶高那 1px 不花流里的地：拿碎片边沿 charging 就凭空多出一格的一截", () => {
    // `from`/`to` 不可信（下一个 `br` 零高、量不到自己那行）才退回边沿——那是这一族唯一的退路。
    const ink = lineCharges(grid, PITCH, Number.NaN, Number.NaN);
    expect(ink[0]).toBeCloseTo(20.8, 10);
    expect(ink[ink.length - 1]).toBeCloseTo(21.8, 10);
    const flow = lineCharges(grid, PITCH, 206, 268.4).reduce((a, b) => a + b, 0);
    expect(ink.reduce((a, b) => a + b, 0) - flow).toBeCloseTo(1, 10);
  });

  it("一页装几格由这个口径决定：按推进 charging 得 44 格，按墨水边沿 charging 只数得出 42", () => {
    // 930.53px 的一页：44 × 20.8 = 915.2 装得下，45 × 20.8 = 936 溢出。
    expect(packPages(column(44, PITCH), A4_BODY)).toBe(1);
    expect(packPages(column(45, PITCH), A4_BODY)).toBe(2);
    // 同样三行、每枚 21.8（墨顶到墨底）：42 枚 915.6 还在纸上，43 枚 937.4 就滚出第二页。
    expect(packPages(column(42, 21.8), A4_BODY)).toBe(1);
    expect(packPages(column(43, 21.8), A4_BODY)).toBe(2);
  });

  it("声明够矮时跟着字体内容盒长：11 磅中文行的枚高是 24.77，不是那句 23.76", () => {
    const runs = [
      { top: 101, bottom: 125 },
      { top: 125.77, bottom: 150 },
      { top: 150.54, bottom: 174 },
    ];
    const c = lineCharges(runs, 23.76, 100, 174.31);
    expect(c[0]).toBeCloseTo(25.77, 10); // 块上沿 100 到下一枚墨顶：段首那截空白也是这段占的
    expect(c[1]).toBeCloseTo(24.77, 10);
    expect(c[2]).toBeCloseTo(23.77, 10); // 末枚收在总推进减去前面各枚，不超纸
    expect(c.reduce((a, b) => a + b, 0)).toBeCloseTo(74.31, 10);
  });

  it("一行的段没有「前面各枚」可减：整段就是它画出来的那一截推进", () => {
    expect(lineCharges([{ top: 205, bottom: 225 }], PITCH, 206, 226.8)[0]).toBeCloseTo(20.8, 10);
    // 行高比推进大时（`exact` 那一族反过来）地板是行高本身。
    expect(lineCharges([{ top: 205, bottom: 210 }], 30, 206, 211)).toEqual([30]);
  });
});

/**
 * 带子从正文里占走的那一截（#157）。真机取到的口径是「带子贴在页边距之内、正文从它下沿起排」：夹具
 * `/tmp/tall-header.docx`（A4、四边 1440 缇、`w:header` 720 缇、页眉 10 行）里 WPS 的正文上沿 ≈3445 缇，
 * 自家改此前恒为 1440 缇——差的 2005 缇 = 133.67px 正是这一式子该算出来的数。
 *
 * 这一式子只认三个数：节说的距离、带子量出来的高、它坐在哪一条边距里。距离没说 → 贴纸沿（0）；
 * 带子整个在边距线以内 → 让 0，那一截空白本来就是纸的边距，不是从正文里借的。
 */
describe("bandYieldPx：带子越过边距线的那一截", () => {
  const TWIPS = (px: number): number => px * 15;

  it("带子坐在边距线以内时一分都不让", () => {
    // 距离 720 缇（48px）加带子 40px = 88px，还在 1440 缇（96px）的边距里。
    expect(bandYieldPx(720, 40, 1440)).toBe(0);
  });

  it("越过边距线的那一截才从正文里占走", () => {
    expect(bandYieldPx(720, 60, 1440)).toBeCloseTo(12, 10);
  });

  it("节没说距离就带子贴纸沿", () => {
    // 发明了 Word 的 docDefaults（720 缇）会把一个没人写过的数算进让位里。
    expect(bandYieldPx(undefined, 108, 1440)).toBeCloseTo(12, 10);
    expect(bandYieldPx(undefined, 108, 1440)).toBeCloseTo(bandYieldPx(0, 108, 1440), 10);
  });

  it("真机那一次差的就是这一截：181.67px 的页眉在 1440 缇的边距上让出 2005 缇", () => {
    const over = bandYieldPx(720, 181.66666666666666, 1440);
    expect(TWIPS(over)).toBeCloseTo(2005, 6);
    // 正文上沿 = 边距 + 让位 = 3445 缇，与真机量到的那个数同一格。
    expect(TWIPS(96 + over)).toBeCloseTo(3445, 6);
  });

  it("距离本身写在边距线之外时，带子空着也让出那一截", () => {
    // `w:header="2160"` 落在 1440 缇的边距之外：摆位在纸上是 144px，正文得从它下面起排。
    expect(bandYieldPx(2160, 0, 1440)).toBeCloseTo(48, 10);
  });
});

/**
 * 表行按行盒切带（#234）：`splitRowBands` 只认几何，不认识 DOM——`tableBandsOf` 把每一格的行盒
 * （`{top,bottom}`）收来，切法在这里钉住。三条不变量：
 * - 各带的高之和恒等于行高（装箱读的是这个和，页数不因切法变）；
 * - 刀口只落在**行与行之间**——任何一格的一枚行盒都不许被一刀穿过；
 * - 行首那一带 `cut: null`（仍是整行一枚的口径），其余各带的 `cut` 就是它自己的上沿。
 */
describe("splitRowBands：可断的行按行盒切成几枚单位", () => {
  const LINE = 25.92;
  const line = (n: number, top: number): { top: number; bottom: number } => ({ top, bottom: top + n * LINE });
  const total = (bands: { height: number }[]): number => bands.reduce((n, b) => n + b.height, 0);

  it("格子里没有量到行盒时整行一枚，cut 为 null", () => {
    expect(splitRowBands(100, 400, [])).toEqual([{ top: 100, height: 300, cut: null }]);
    expect(splitRowBands(100, 400, [[]])).toEqual([{ top: 100, height: 300, cut: null }]);
  });

  it("0 高/负高的行不产刀口", () => {
    expect(splitRowBands(100, 100, [])).toEqual([{ top: 100, height: 0, cut: null }]);
    expect(splitRowBands(100, 99, [])).toEqual([{ top: 100, height: 0, cut: null }]);
  });

  it("一格五行切五带，首带 cut=null、其余 cut 是自己的上沿", () => {
    const lines = Array.from({ length: 5 }, (_, i) => line(1, 100 + i * LINE));
    const bands = splitRowBands(100, 100 + 5 * LINE, [lines]);
    expect(bands.length).toBe(5);
    for (const b of bands) expect(b.height).toBeCloseTo(LINE, 10);
    expect(bands.map((b) => b.cut)).toEqual([null, 100 + LINE, 100 + 2 * LINE, 100 + 3 * LINE, 100 + 4 * LINE]);
    expect(total(bands)).toBeCloseTo(5 * LINE, 10);
  });

  it("刀口取各格行盒上沿的并集", () => {
    // 左格三行、右格两行：两格行高一致，刀口就是两条线，右格最后一行比左格短。
    const left = [line(1, 0), line(1, LINE), line(1, 2 * LINE)];
    const right = [line(1, 0), line(1, LINE)];
    const bands = splitRowBands(0, 3 * LINE, [left, right]);
    expect(bands.map((b) => b.height)).toEqual([LINE, LINE, LINE]);
  });

  it("穿行盒的候选丢掉：一格两行错开时不在它中间下刀", () => {
    // 左格第一行 0..2 行高（字号大），右格两行各 1 行高；右格第二行的上沿落在左格第一行中间。
    const left = [{ top: 0, bottom: 2 * LINE }];
    const right = [line(1, 0), line(1, LINE)];
    const bands = splitRowBands(0, 2 * LINE, [left, right]);
    expect(bands).toEqual([{ top: 0, height: 2 * LINE, cut: null }]);
  });

  it("紧挨着的两个候选（<0.5px）只算一刀", () => {
    // 两格的行首各差 0.2px：两处都不穿任何行盒，但只落一刀（靠下那一枚丢掉）。
    const a = [{ top: 0, bottom: LINE }, { top: LINE, bottom: 2 * LINE }];
    const b = [{ top: 0, bottom: LINE - 0.3 }, { top: LINE + 0.2, bottom: 2 * LINE + 0.2 }];
    const bands = splitRowBands(0, 2 * LINE, [a, b]);
    expect(bands).toEqual([
      { top: 0, height: LINE, cut: null },
      { top: LINE, height: LINE, cut: LINE },
    ]);
  });

  it("两格行盒错开半格设备像素：容差按缩放放开后仍能下刀", () => {
    // 装机件在 1.59 档量到两格行首差 0.795 视口 px（布局 0.5 = dpr 2 上的半格）：缺省 0.5 的容差把这一刀判成
    // 穿过右格第一行，整行退成一枚整带；容差随缩放放到 1.59 之后，这半格仍算落在行沿上。
    const a = [line(1, 0), line(1, LINE)];
    const b = [line(1, 0.795), line(1, LINE + 0.795)];
    expect(splitRowBands(0, 2 * LINE, [a, b])).toEqual([{ top: 0, height: 2 * LINE, cut: null }]);
    const bands = splitRowBands(0, 2 * LINE, [a, b], 1.59);
    expect(bands.map((b) => b.height)).toEqual([LINE, LINE]);
    expect(bands[0]!.cut).toBeNull();
  });

  it("容差放宽也不许切穿错开超过容差的行盒", () => {
    // 右格第二行整行压在左格第一行中间（错开一行高）：放宽到 1.59 也不下刀。
    const a = [line(2, 0)];
    const b = [line(1, 0), line(1, LINE)];
    expect(splitRowBands(0, 2 * LINE, [a, b], 1.59)).toEqual([{ top: 0, height: 2 * LINE, cut: null }]);
  });

  it("行的上下沿本身不是刀口", () => {
    const lines = [
      { top: -LINE, bottom: 0 },
      { top: 0, bottom: LINE },
      { top: LINE, bottom: 2 * LINE },
      { top: 2 * LINE, bottom: 3 * LINE },
      { top: 3 * LINE, bottom: 4 * LINE },
    ];
    const bands = splitRowBands(0, 3 * LINE, [lines]);
    expect(bands.length).toBe(3);
    expect(bands[0]!.cut).toBeNull();
    expect(total(bands)).toBeCloseTo(3 * LINE, 10);
  });

  it("每格的首行不是刀口：单行的行整枚一枚，不切出空的第一带", () => {
    // 真机件上量到过这个假刀口：首行上沿进候选，一行的行被切成两带，垫片 1629px 插在格首、整行沉下去。
    expect(splitRowBands(100, 100 + LINE, [[{ top: 100, bottom: 100 + LINE }]])).toEqual([
      { top: 100, height: LINE, cut: null },
    ]);
    // 两格首行错开 6px：首行不进候选，这一带也就没有刀口（错开的那条线还会穿另一格的行盒）。
    const a = [{ top: 106, bottom: 106 + LINE }];
    const b = [{ top: 100, bottom: 100 + LINE }];
    const bands = splitRowBands(100, 106 + LINE, [a, b]);
    expect(bands).toHaveLength(1);
    expect(bands[0]!.cut).toBeNull();
    expect(bands[0]!.height).toBeCloseTo(6 + LINE, 10);
  });

  it("各行带的高加回去等于整行高（装箱不因切法换页）", () => {
    const lines = Array.from({ length: 40 }, (_, i) => line(1, i * LINE));
    const whole = splitRowBands(0, 40 * LINE, []);
    const cut = splitRowBands(0, 40 * LINE, [lines]);
    expect(total(whole)).toBeCloseTo(total(cut), 10);
    expect(cut.length).toBe(40);
  });
});
