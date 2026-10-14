/**
 * 度量单位：WPS 的状态栏在最右边一排里挂着一格 `KxWpsStatusUnitCommand`（`office6/res/wpscommon.kuip`，
 * 459 行，夹在 字数 与 编码 之间），detail 说的是「@Measurement unit. Click to open "Options" dialog.」
 * ——点开的是「选项」对话框。真机上那一格显示什么字面、选项里摆哪几种单位、默认是哪一种、选完存在
 * 哪里（配置？注册表？）**都未取证**（`.qm` 那堵墙与 `doc-word-count.ts` 头注记的是同一件）。
 *
 * 这一条是我们先接上的一个**有界**切片：选项对话框（`doc-options.ts`）只摆「测量单位」一组，
 * 四种单位 厘米/毫米/英寸/磅 与这个次序是我们自己定的集合，默认厘米与状态栏那格原来的口径一致。
 * 选择**只活在这个窗口里**：应用还没有设置存储，重启回到厘米（真机存去哪里未取证）。
 *
 * 受众有两条腿，两条腿的小数位**不是一套**，各有各的口径：
 *
 * - **读数**（`doc-status-bar.ts` 的「纵向位置」那一格，`formatLength`）：以 CSS 参考像素为底
 *   （1 英寸 = 96px = 2.54 厘米 = 25.4 毫米 = 72 磅，其余三个方向由此定义），每位报几位小数是我们
 *   自己定的（厘米/英寸 两位、毫米/磅 一位，让四种的绝对精度大致相当），真机报几位、报什么字面**未取证**。
 * - **编辑**（`doc-page-setup.ts` 的九个格子，`twipText`/`twipsFromUnit`）：以缇（twip）为底，位数不是
 *   为好看定的，是**能精确往返的最小口径**——0..60000 每个整数缇数写进格子再读回来必须是它自己，否则
 *   一个没人动过行距/边距的对话框按一次「确定」就把文件改了。证明：格子回读的误差最多半格，要求它
 *   不超过半个缇；相邻两个缇的间距又必须大于半个格距（否则两个不同的缇写成同一个字面）。两条不等式
 *   在四单位下的解是 厘米 3 位（0.0005 厘米 = 0.283 缇；相邻 0.00176 厘米）、毫米 2 位（0.005 毫米 =
 *   0.283 缇；相邻 0.0176 毫米）、英寸 4 位（0.00005 英寸 = 0.072 缇；相邻 0.00069 英寸）、磅 2 位
 *   （0.005 磅 = 0.1 缇；相邻 0.05 磅）。全遍历的恒等断言在 `tests/doc-units.test.ts`。
 */

export type Unit = "cm" | "mm" | "in" | "pt";

export interface UnitDef {
  id: Unit;
  /** 中文名：选项对话框上的字面，也是状态栏那格显示的字，读数后缀同一个词。 */
  name: string;
  /** **读数**的位数口径（`formatLength`）。编辑格子的位数是另一套，见 `twipText`。 */
  decimals: number;
}

/** 选项里摆的四种，次序就是对话框上的次序（真机摆哪几种、什么序未取证）。 */
export const UNITS: readonly UnitDef[] = [
  { id: "cm", name: "厘米", decimals: 2 },
  { id: "mm", name: "毫米", decimals: 1 },
  { id: "in", name: "英寸", decimals: 2 },
  { id: "pt", name: "磅", decimals: 1 },
];

export const unitDef = (u: Unit): UnitDef => UNITS.find((x) => x.id === u)!;

let unit: Unit = "cm";
const listeners = new Set<() => void>();

export const currentUnit = (): Unit => unit;

/** 换单位：变了才通知（同值重选不惊动订阅者）。订阅者在通知里现读，不缓存。 */
export function setUnit(next: Unit): void {
  if (next === unit) return;
  unit = next;
  for (const fn of [...listeners]) fn();
}

export function onUnitChange(fn: () => void): () => void {
  listeners.add(fn);
  return () => {
    listeners.delete(fn);
  };
}

const PER_PX: Record<Unit, number> = { cm: 2.54 / 96, mm: 25.4 / 96, in: 1 / 96, pt: 72 / 96 };

/** 一个长度读数的字面：「2.54 厘米」。小数位照该单位的口径。 */
export function formatLength(px: number, u: Unit = unit): string {
  const def = unitDef(u);
  return `${(px * PER_PX[u]).toFixed(def.decimals)} ${def.name}`;
}

/** 编辑轴的每单位缇数：1 英寸 = 1440、1 磅 = 20 是 OOXML 的定义，厘米与毫米是它的导数。 */
const TWIPS_PER_UNIT: Record<Unit, number> = { cm: 1440 / 2.54, mm: 1440 / 25.4, in: 1440, pt: 20 };

/** 编辑轴的小数位：头注里那两条不等式在四单位下的解。改这里先改 `tests/doc-units.test.ts` 的遍历。 */
const TWIP_DECIMALS: Record<Unit, number> = { cm: 3, mm: 2, in: 4, pt: 2 };

/** 一个缇数写进格子的字面（尾零不留，「42」不写「42.000」）。 */
export function twipText(tw: number, u: Unit = unit): string {
  return String(Number((tw / TWIPS_PER_UNIT[u]).toFixed(TWIP_DECIMALS[u])));
}

/** 格子里的数换回缇数，四舍五入到整缇。 */
export function twipsFromUnit(v: number, u: Unit = unit): number {
  return Math.round(v * TWIPS_PER_UNIT[u]);
}
