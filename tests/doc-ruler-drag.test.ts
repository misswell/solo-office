import { describe, expect, it } from "vitest";
import { twipsToPx } from "../src/core/docx.js";
import { indentTwipsOf, type RulerDragGeometry } from "../src/ui/doc-ruler.js";

/**
 * 标尺拖动的换算：纸面像素 → 件里那一句话（缇）。所有输入都是纸的像素（`zoom` 之前那一层），
 * 所以这里没有缩放变量可测——尺与纸同一个 `zoom`，手势那边先除回过。
 *
 * A4 宽 11906 缇、左右页边距 1800：纸 793.7px、正文栏左沿 120px、右沿 673.7px。
 */
const PW = twipsToPx(11906);
const PAD = twipsToPx(1800);
const geom = (over: Partial<RulerDragGeometry> = {}): RulerDragGeometry => ({
  pageW: PW,
  padL: PAD,
  padR: PAD,
  inlineLeft: 0,
  inlineRight: 0,
  leftX: PAD,
  firstX: PAD,
  rightX: PAD + (PW - PAD * 2),
  ...over,
});

describe("标尺标记的 px→缇换算", () => {
  it("左标记：正文栏左沿是 0 缇，往里 48px 是 720 缇", () => {
    expect(indentTwipsOf("left", PAD, geom())).toBe(0);
    expect(indentTwipsOf("left", PAD + twipsToPx(720), geom())).toBe(720);
  });

  it("左标记减掉段自己的 padding（列表缩进画在 padding 里，不在 margin 里）", () => {
    const list = geom({ inlineLeft: 32, leftX: PAD + 32 + twipsToPx(720) });
    expect(indentTwipsOf("left", list.leftX, list)).toBe(720);
  });

  it("首行：相对左标记的偏移，正数往里、负数悬挂", () => {
    const g = geom({ leftX: PAD + twipsToPx(720), firstX: PAD + twipsToPx(720) + twipsToPx(420) });
    expect(indentTwipsOf("first", g.firstX, g)).toBe(420);
    expect(indentTwipsOf("first", g.leftX, g)).toBe(0);
    expect(indentTwipsOf("first", g.leftX - twipsToPx(240), g)).toBe(-240);
  });

  it("右标记：从正文栏右沿往回数，末行右沿是 0 缇", () => {
    const g = geom();
    expect(indentTwipsOf("right", PW - PAD, g)).toBe(0);
    expect(indentTwipsOf("right", PW - PAD - twipsToPx(360), g)).toBe(360);
  });

  it("右标记减掉段自己的 padding", () => {
    const g = geom({ inlineRight: 16 });
    expect(indentTwipsOf("right", PW - PAD - 16 - twipsToPx(360), g)).toBe(360);
  });

  it("拖回原处读出来的就是件里原来那个数（720/420/360 往返）", () => {
    const g = geom({
      leftX: PAD + twipsToPx(720),
      firstX: PAD + twipsToPx(720) + twipsToPx(420),
      rightX: PW - PAD - twipsToPx(360),
    });
    expect(indentTwipsOf("left", g.leftX, g)).toBe(720);
    expect(indentTwipsOf("first", g.firstX, g)).toBe(420);
    expect(indentTwipsOf("right", g.rightX, g)).toBe(360);
  });

  it("半像素也照 15 缇/px 落，不替 WPS 发明吸附步长", () => {
    expect(indentTwipsOf("left", PAD + 0.5, geom())).toBe(8); // 7.5 → 8
    expect(indentTwipsOf("left", PAD + 2.5, geom())).toBe(38); // 37.5 → 38
  });
});
