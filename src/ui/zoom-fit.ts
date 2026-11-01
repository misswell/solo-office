/**
 * 视图页 `RB_Zoom` 里 WPS 声明的两枚腿（`ongmani/CT_View.kuip` 的 `ZoomPageWidth`——`text="@Page Width"`、
 * 下拉里 `index="-1"`——与 `ZoomOnePage`——`text="@One Page"`、`index="-3"`）的式子，外加落笔时的取整与钳位。
 * 件里给的是这两枚的名字与位置；把纸量进窗口那两枚式子是这一族自己推的（WPS 落笔的精确数与「是否记住上次缩放」
 * 未机器取证）。这里不认识 DOM：窗口和纸的数由 `renderer.ts` 交进来。
 *
 * 「双页」（`ZoomTwoPages`）不在这里：我们的纸是一列连续铺开的，两页并排要的是另一套铺排，不是另一个缩放数。
 */

/** `#view` 为竖滚动条留的宽。打开件时那初始档与「页宽」用的是同一个式子，所以只留这一处来路。 */
export const VIEW_SCROLLBAR_PX = 15;

/** ribbon 那几档共用的上下界。 */
export const ZOOM_MIN = 0.4;
export const ZOOM_MAX = 2.5;

/**
 * 按纸定档的原始比：页宽 = 可用宽 ÷ 纸宽；单页 = 页宽那一头与「可视高 ÷ 纸高」里小的那一个——一整张纸要
 * 放进窗口，两个方向都得放下。纸的像素不成数（`w:pgSz` 没说、或说成 0）就没有这一档。
 */
export function paperFit(
  fit: "width" | "page",
  viewW: number,
  viewH: number,
  paperW: number,
  paperH: number,
): number | undefined {
  if (!(paperW > 0) || !(paperH > 0)) return undefined;
  const byWidth = (viewW - VIEW_SCROLLBAR_PX) / paperW;
  return fit === "width" ? byWidth : Math.min(byWidth, viewH / paperH);
}

/**
 * 落笔的档。「按纸定档」向下取到两位小数：向上取会把纸放得比窗口还大，而那两档说的正是「放得下」——
 * 放不下就是没做到。⊖/⊕ 的 ±0.1 仍走四舍五入。两条都钳到同一界。
 */
export function snapZoom(z: number, fit: boolean): number {
  const two = fit ? Math.floor(z * 100) / 100 : Math.round(z * 100) / 100;
  return Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, two));
}
