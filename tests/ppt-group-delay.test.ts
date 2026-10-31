import { describe, expect, it } from "vitest";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { strToU8, unzipSync, zipSync } from "fflate";
import { Presentation, type Shape } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { makePptx } from "./helpers/make.js";

/**
 * 授权语料 35 份里只有 1 枚 `p:grpSp`（wideLayout/presetLayout_14），它的 `chOff`/`chExt`
 * 与组的 `off`/`ext` 相同；组员坐标写在子空间里，按幻灯片坐标读会把整组再平移一次。
 * 同一份里还有唯一一枚不支持的预设 `flowChartDelay`（D 形），此前画成直角矩形。
 */
const CORPUS = ".scratch/wps-corpus";
const dec = new TextDecoder();

function withShapes(replacement: string): Presentation {
  const zip = unzipSync(makePptx());
  const before = dec.decode(zip["ppt/slides/slide1.xml"]);
  const slide = before.replace(/<p:sp><p:nvSpPr><p:cNvPr id="4" name="椭圆"\/>[\s\S]*?<\/p:sp>/, replacement);
  expect(slide).not.toBe(before);
  zip["ppt/slides/slide1.xml"] = strToU8(slide);
  return Presentation.openBytes(zipSync(zip, { level: 6 }));
}

/** 组框 200×100、子空间 400×200：组员 (300,50)+200×100 应落到组内 (50,25)+100×50。 */
const GROUP = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="组"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="952500"/><a:chOff x="1905000" y="0"/><a:chExt cx="3810000" cy="1905000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="5" name="组内矩形"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="2857500" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="00B050"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp></p:grpSp>`;

const DELAY = `<p:sp><p:nvSpPr><p:cNvPr id="4" name="延迟"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="1905000"/></a:xfrm><a:prstGeom prst="flowChartDelay"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`;

function flatten(shapes: Shape[]): Shape[] {
  const out: Shape[] = [];
  for (const s of shapes) {
    out.push(s);
    if (s.children) out.push(...flatten(s.children));
  }
  return out;
}

describe("pptx 组变换与 flowChartDelay", () => {
  it("组员坐标按 chOff/chExt 映射进组框，渲染时只落一次组偏移", () => {
    const pres = withShapes(GROUP);
    const group = pres.slides[0].shapes.find((s) => s.kind === "group");
    expect(group).toMatchObject({ x: 100, y: 100, width: 200, height: 100 });
    const child = group?.children?.[0];
    expect(child).toMatchObject({ x: 50, y: 25, width: 100, height: 50 });

    const html = renderSlideHtml(pres, 0);
    expect(html).toContain("left:100px;top:100px;width:200px;height:100px");
    expect(html).toContain("left:50px;top:25px;width:100px;height:50px");
    expect(html).not.toContain("left:300px");
  });

  it("flowChartDelay 按右半椭圆裁形，不再画成直角矩形", () => {
    const pres = withShapes(DELAY);
    const html = renderSlideHtml(pres, 0);
    expect(html).toContain("border-radius:0 50% 50% 0 / 0 50% 50% 0");
  });

  it.skipIf(!existsSync(CORPUS))("语料里 1 枚组、1 枚 flowChartDelay 都按新口径落地", () => {
    const dir = CORPUS;
    const files = readdirSync(dir)
      .sort()
      .filter((f) => /\.pptx$/i.test(f))
      .map((f) => join(dir, f));
    let groups = 0;
    let delays = 0;
    for (const file of files) {
      const pres = Presentation.openBytes(new Uint8Array(readFileSync(file)));
      for (const slide of pres.slides) {
        for (const s of flatten(slide.shapes)) {
          if (s.kind === "group") {
            groups++;
            // 组员必须落在组框内（留 1px 舍入）：子空间坐标若按幻灯片坐标读会跑出框外。
            for (const c of s.children ?? []) {
              expect(c.x).toBeGreaterThanOrEqual(-1);
              expect(c.y).toBeGreaterThanOrEqual(-1);
              expect(c.x + c.width).toBeLessThanOrEqual(s.width + 1);
              expect(c.y + c.height).toBeLessThanOrEqual(s.height + 1);
            }
          }
          if (s.geometry === "flowChartDelay") {
            delays++;
            expect(renderSlideHtml(pres, slide.index)).toContain("border-radius:0 50% 50% 0 / 0 50% 50% 0");
          }
        }
      }
    }
    expect(files.length).toBeGreaterThanOrEqual(35);
    expect(groups).toBe(1);
    expect(delays).toBe(1);
  });
});
