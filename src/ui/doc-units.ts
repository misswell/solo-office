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
 * 这条腿接上的受众是状态栏的「纵向位置」那一格（`doc-status-bar.ts`）：它本来写死报厘米两位小数，
 * 现在按当前单位报。页面设置对话框**仍报厘米**：它三位小数的口径是与厘米绑着量的（851 缇 = 1.501
 * 厘米，三位小数能把每个整数缇数原样带回，见 `doc-page-setup.ts` 头注），换成英寸/磅要先把「多少位
 * 小数能原样带回一个缇」重新量一遍，这一轮不做。
 *
 * 换算的底是 CSS 参考像素（1 英寸 = 96px = 2.54 厘米 = 25.4 毫米 = 72 磅，其余三个方向由此定义）。
 * 每种单位报几位小数是我们自己定的（厘米/英寸 两位、毫米/磅 一位，让四种的绝对精度大致相当），
 * 真机报几位、报什么字面**未取证**。
 */

export type Unit = "cm" | "mm" | "in" | "pt";

export interface UnitDef {
  id: Unit;
  /** 中文名：选项对话框上的字面，也是状态栏那格显示的字，读数后缀同一个词。 */
  name: string;
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
