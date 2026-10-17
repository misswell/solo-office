import { describe, expect, it } from "vitest";
import type { ParaBorderSet, Paragraph } from "../src/core/docx.js";
import { WordDocument, paraBorderChanges, syncParaSlots } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { applyBoxPatch, legPatch, type ParaBoxPatch } from "../src/ui/doc-para-border.js";
import { makeChainDocx } from "./helpers/make-paragraph.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
const byText = (doc: WordDocument, needle: string): Paragraph =>
  doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph").find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
const boxOf = (xml: string, needle: string): string =>
  /<w:pBdr\b[^>]*\/>|<w:pBdr\b[^>]*>[\s\S]*?<\/w:pBdr>/.exec(paraXml(xml, needle))?.[0] ?? "";
const shdOf = (xml: string, needle: string): string => /<w:shd[^>]*\/>/.exec(paraXml(xml, needle))?.[0] ?? "";
const drawnStyle = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<(?:div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>(?:(?!<\/(?:div|h[1-6])>)[\s\S])*<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return / style="([^"]*)"/.exec(m[0])?.[1] ?? "";
  }
  return "";
};

/**
 * 段落「边框与底纹」那条命令**落下去之后**的账。#144 让框与地读得出、画得出、保存不丢，但命令面是空的；
 * #145 接上对话框与 ribbon 那几条腿，于是第一次有了一段可以按用户的话去问的文件：
 * `setParagraphBox` 交的是 `applyBoxPatch(段落手里的框, 补丁)` 再 `syncParaSlots`，与对话框、与那几条腿同一个口子。
 *
 * 每一击的落笔形状都对着件量出来的那两种（`.scratch/pbdr-audience.ts`：46 份授权件、26 枚 `w:pBdr` / 74 枚边孩子）：
 * - 要画线的 8 枚全是 `w:val="single" w:sz="6" w:space="1" w:color="auto"`；
 * - 说不要的 66 枚全是 `w:sz="0"`，间距只有 1 与 4 两档，颜色 74/74 写 `auto`；
 * - 底纹 30/30 枚写 `w:val="clear" w:color="auto"` 加一枚十六进制 `w:fill`。
 * 三条纪律在这一族里最要紧，所以每条都有个反向对照：**逐边**（§17.7.2，只写动过的那一边）、**对着读快照比**
 * （原样确定不该多写一个字节）、**没动的边按件的字节带着走**（含模型没有字段的 `w:themeColor`/`w:frame`，也含
 * 属性原本的写法与先后）。
 */
describe("docx 段落「边框与底纹」的落笔", () => {
  /** 样式借一整圈，段落自己说的各有不同：四段各自代表一条命令会碰到的四种情形。 */
  const DOC = makeChainDocx(
    [
      { id: "Plain" },
      {
        id: "Boxed",
        pPr: '<w:pBdr><w:top w:val="double" w:sz="12" w:space="2" w:color="auto"/><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>',
      },
    ],
    [
      { text: "整段没说过框" },
      { pStyle: "Boxed", text: "上下是样式说的" },
      { pStyle: "Boxed", pPr: '<w:pBdr><w:bottom w:color="auto" w:val="none" w:sz="0" w:space="1"/></w:pBdr>', text: "自己退回了底线" },
      {
        pPr:
          '<w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto" w:themeColor="accent1" w:frame="1"/></w:pBdr>' +
          '<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>',
        text: "自己说了顶线与一层地",
      },
    ],
  );

  /** 一次点击：把补丁落到那一段的框上，再走 `syncParaSlots`——`setParagraphBox` 逐段做的正是这两步。 */
  const click = (doc: WordDocument, needle: string, patch: ParaBoxPatch): Paragraph => {
    const p = byText(doc, needle);
    if (patch.sides) p.borders = applyBoxPatch(p.borders, patch);
    if (patch.shading !== undefined) p.shading = patch.shading;
    syncParaSlots(p);
    return p;
  };

  it("外侧框线那条腿：一段从没说过框的段落，第一次说话按件的形状与 `CT_PBdr` 的边序写四边", () => {
    const doc = open(DOC);
    click(doc, "整段没说过框", legPatch("outside"));
    const side = (name: string): string => `<w:${name} w:val="single" w:sz="6" w:space="1" w:color="auto"/>`;
    expect(boxOf(save(doc), "整段没说过框")).toBe(`<w:pBdr>${["top", "left", "bottom", "right"].map(side).join("")}</w:pBdr>`);
    // 别的段一段没被碰到：件里借来的那一圈不落进段落。
    expect(boxOf(save(doc), "上下是样式说的")).toBe("");
    // 页上画得出来（#144 那一条从命令这一侧重走一遍）。
    expect(drawnStyle(doc, "整段没说过框")).toContain("border-top:1px solid currentColor");
  });

  it("左边框那条腿只说那一边：从样式借来的上下不被抄进段落，`w:pBdr` 里也只有它", () => {
    const doc = open(DOC);
    click(doc, "上下是样式说的", legPatch("left"));
    expect(boxOf(save(doc), "上下是样式说的")).toBe('<w:pBdr><w:left w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>');
  });

  /**
   * 件里 26 枚 `w:pBdr` 有 24 枚在 `word/styles.xml`、正文段落自己写的 0 枚：把借来的那一圈整套装进段落身上，
   * 每一枚都是替文件新编的一句样式的话。这一条把那个风险摆在最刁的地方——用户点的那一边**恰好**与样式说的一模一样。
   */
  it("原样确定不是用户的话：样式说过那一边、段落再说一遍同样的数，一个字节都不写", () => {
    const doc = open(DOC);
    const p = click(doc, "上下是样式说的", legPatch("bottom"));
    // 字段里留着解析过样式链的那一份（对话框该摆它），但 `paraBorderChanges` 对着读快照比，比不出动过的边。
    expect(p.borders).toEqual({
      top: { val: "double", size: 12, space: 2, color: "auto" },
      bottom: { val: "single", size: 6, space: 1, color: "auto" },
    });
    expect(paraBorderChanges(p)).toBeUndefined();
    expect(boxOf(save(doc), "上下是样式说的")).toBe("");
    expect(p.pPrSlots?.map((s) => s.tag)).toEqual(["pStyle"]);
  });

  /**
   * 「无框」是一句**说出来的话**，不是抹掉：`w:val="none"` 连着 `w:sz="0"` 与那一边自己的间距档，这一层说得出
   * 的五条边一起说（四条加中间那一条——漏了中间那条，「内部横框线」落下去的线就撤不掉），才退得掉样式借来的
   * 那一圈（§17.7.2 逐边解决）。语料里这么写的 66 枚边孩子全带 `w:sz="0"`，顶/底写 `space="1"`
   * （32 枚）、左/右写 `space="4"`（32 枚）——那 16 枚四边全说的样式就是这个形状；中间那一条没有样本，
   * 按它**横**边那一档走 1。
   */
  it("无框那条腿写下四边与中间那一条 `w:val=\"none\"`，把样式借来的那圈线退回样式那里", () => {
    const doc = open(DOC);
    click(doc, "上下是样式说的", legPatch("none"));
    const side = (name: string, space: number): string => `<w:${name} w:val="none" w:sz="0" w:space="${space}" w:color="auto"/>`;
    expect(boxOf(save(doc), "上下是样式说的")).toBe(
      `<w:pBdr>${[side("top", 1), side("left", 4), side("bottom", 1), side("between", 1), side("right", 4)].join("")}</w:pBdr>`,
    );
    expect(drawnStyle(doc, "上下是样式说的")).not.toMatch(/border-/);
    // 没被点到的段落照旧：件里那一句「这条线不要」仍然画不出来，样式那条顶线仍然画得出来。
    expect(drawnStyle(doc, "自己退回了底线")).not.toContain("border-bottom");
    expect(drawnStyle(doc, "自己说了顶线与一层地")).toContain("border-top");
  });

  /**
   * 没动的边按件的字节带着走，动的边只换用户填了的那一格。这里动的顶边在件里带两枚模型没有字段的东西
   * （`w:themeColor`、`w:frame`），底纹那一句的 `w:val`/`w:color` 同样是件的说法——它们都得留在原位。
   */
  it("只改顶边的磅：线型、间距、颜色与件写的主题色、边线框一起留着，属性按 `CT_Border` 的序排", () => {
    const doc = open(DOC);
    click(doc, "自己说了顶线与一层地", { sides: { top: { sizeEighths: 24 } } });
    expect(boxOf(save(doc), "自己说了顶线与一层地")).toBe(
      '<w:pBdr><w:top w:val="single" w:sz="24" w:space="1" w:color="auto" w:themeColor="accent1" w:frame="1"/></w:pBdr>',
    );
    // 底纹这一击没填，所以那一枚一个字节不动——件自己的写法（包括十六进制怎么写）不归这一族管。
    expect(shdOf(save(doc), "自己说了顶线与一层地")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>');
    // 顶边的颜色没填：`auto` 沿用这一段这一边原来的说法，而不是替它挑一种颜色。磅数按八分之一磅走：
    // 24/8 = 3 磅 = 4 CSS px（`borderCss` 的 `size/6`），间距还是 1 磅 = 1.333 px。
    expect(drawnStyle(doc, "自己说了顶线与一层地")).toContain("border-top:4px solid currentColor;padding-top:1.333px");
  });

  it("段落自己写过、这一击没点的边，连件里另立的属性先后一起照旧", () => {
    const doc = open(DOC);
    click(doc, "自己退回了底线", legPatch("top"));
    // 底线在件里写的是 `w:color` 在 `w:val` 前，没动它就还是那个写法；新写的顶边按件的顺序。
    expect(boxOf(save(doc), "自己退回了底线")).toBe(
      '<w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/><w:bottom w:color="auto" w:val="none" w:sz="0" w:space="1"/></w:pBdr>',
    );
    // 样式借来的顶边被这一段自己那条顶线按逐边解决掉，页面从这一句起画它。
    expect(drawnStyle(doc, "自己退回了底线")).toContain("border-top:1px solid currentColor");
    expect(drawnStyle(doc, "自己退回了底线")).not.toContain("border-bottom");
  });

  it("底纹填一种地按件的说法写十六进制，`auto` 那句 token 原样回填", () => {
    const doc = open(DOC);
    // 交进来的形状是读端给一枚颜色画的那个形状（`#RRGGBB` 大写），对话框的 `fillValue` 归的就是它。
    click(doc, "整段没说过框", { shading: "#FFC000" });
    expect(shdOf(save(doc), "整段没说过框")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="FFC000"/>');
    expect(drawnStyle(doc, "整段没说过框")).toContain("background-color:#FFC000");

    const back = open(DOC);
    click(back, "自己说了顶线与一层地", { shading: "auto" });
    // 「这一段要页面自己的地」是一句话，不是没说：不写 `w:shd` 就等于让样式那层地铺进来。
    expect(shdOf(save(back), "自己说了顶线与一层地")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
    expect(drawnStyle(back, "自己说了顶线与一层地")).not.toMatch(/background/);
  });

  /**
   * 样式借来的那个盒子在字段里是**按引用**带的（`mergeParaBorders`），所以往它里面改就是替样式改它自己——
   * 别处借过同一套框的段落会跟着变。命令面必须换掉整个物件，这一条是那条纪律唯一能抓住的地方。
   */
  it("落笔换掉整个物件，样式手里那一套一个属性没被动过", () => {
    const doc = open(DOC);
    const p = byText(doc, "上下是样式说的");
    const lent = doc.paragraphStyleOf("Boxed")!.borders!;
    expect(p.borders).toBe(lent);
    p.borders = applyBoxPatch(p.borders, legPatch("outside"));
    expect(p.borders).not.toBe(lent);
    expect(lent).toEqual({
      top: { val: "double", size: 12, space: 2, color: "auto" },
      bottom: { val: "single", size: 6, space: 1, color: "auto" },
    });
    // 其余三条边是这一段新说的；底线与样式说的一模一样，于是由 `paraBorderChanges` 比掉。
    expect(Object.keys(paraBorderChanges(p)!)).toEqual(["top", "left", "right"]);
  });

  /** 说不出线型的一边写不进模型：那是替这一段编一条它没点过的线（对话框同一条判定，在 `build()` 里拦）。 */
  it("只填磅而不点线型时，原本没有线的那一边不进模型", () => {
    const patch: ParaBoxPatch = { sides: { left: { spacePt: 3 }, top: { sizeEighths: 12 } } };
    expect(applyBoxPatch(undefined, patch)).toEqual({});
    const already: ParaBorderSet = { top: { val: "thick", size: 24, space: 4, color: "FF0000" } };
    expect(applyBoxPatch(already, patch)).toEqual({ top: { val: "thick", size: 12, space: 4, color: "FF0000" } });
    // 沿用这一段这一边原来的说法，包括件里那枚模型摆不出来的线型，不顺手归成 `single`。
    expect(applyBoxPatch(already, { sides: { top: { val: "double" } } })).toEqual({
      top: { val: "double", size: 24, space: 4, color: "FF0000" },
    });
  });

  it("点过之后再存一次，那一句不再改字节", () => {
    const doc = open(DOC);
    // 两种颜色在模型里是两个形状，各按各自那一族读进来的样子走：底纹是读端归过的一枚 `#RRGGBB`
    // （`shadingOf` 过 `hex()`），边框的 `w:color` 是件自己写的字面，不带 `#`（`borderOf` 原样带）。
    // 写回时两边都落回件的说法（`shadingFillToken` 去掉 `#`，边框直接抄），所以第二次存不再有差异。
    click(doc, "上下是样式说的", { sides: { right: { val: "dashed", sizeEighths: 8, spacePt: 2, color: "C00000" } }, shading: "#E2EFDA" });
    const once = writeDocx(doc);
    const again = save(open(once));
    expect(boxOf(again, "上下是样式说的")).toBe('<w:pBdr><w:right w:val="dashed" w:sz="8" w:space="2" w:color="C00000"/></w:pBdr>');
    expect(shdOf(again, "上下是样式说的")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="E2EFDA"/>');
    // 读回来的段对着新的快照比，第二次保存看不出这一段改过什么。
    const read = open(once);
    expect(paraBorderChanges(byText(read, "上下是样式说的"))).toBeUndefined();
  });

  /**
   * #153：WPS 那枚下拉声明的 `AllBorders`/`OutsideBorders`/`HMiddleBorder` 三条落在段落上说的是
   * `w:pBdr/w:between`（`CT_PBdr` 没有 insideH/insideV）。三条各写各的，中间那一条排在其自己的边序里
   * （`PBDR_ORDER`：top、left、bottom、**between**、right、bar），不是接在四条边后面。
   */
  it("所有框线说五条：四条边加中间那一条，`w:between` 按 `CT_PBdr` 的边序排在 bottom 与 right 之间", () => {
    const doc = open(DOC);
    click(doc, "整段没说过框", legPatch("all"));
    const side = (name: string): string => `<w:${name} w:val="single" w:sz="6" w:space="1" w:color="auto"/>`;
    expect(boxOf(save(doc), "整段没说过框")).toBe(
      `<w:pBdr>${["top", "left", "bottom", "between", "right"].map(side).join("")}</w:pBdr>`,
    );
  });

  it("外侧框线只说四条边——它与所有框线差的就是中间那一条", () => {
    const doc = open(DOC);
    click(doc, "整段没说过框", legPatch("outside"));
    expect(boxOf(save(doc), "整段没说过框")).not.toContain("w:between");
    expect(Object.keys(byText(doc, "整段没说过框").borders!)).toEqual(["top", "left", "bottom", "right"]);
  });

  it("内部横框线只说中间那一条：样式借来的上下一个字节不抄过来", () => {
    const doc = open(DOC);
    click(doc, "上下是样式说的", legPatch("insideH"));
    expect(boxOf(save(doc), "上下是样式说的")).toBe('<w:pBdr><w:between w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>');
  });

  it("无框线把中间那一条一起说不要：不然那条线撤不掉", () => {
    const doc = open(DOC);
    click(doc, "上下是样式说的", legPatch("insideH"));
    click(doc, "上下是样式说的", legPatch("none"));
    expect(boxOf(save(doc), "上下是样式说的")).toContain(
      '<w:between w:val="none" w:sz="0" w:space="1" w:color="auto"/>',
    );
    // 撤掉之后纸上不再画那一条（画与撤是同一句话的两头）。
    expect(drawnStyle(doc, "上下是样式说的")).not.toContain("background-image");
  });

  /**
   * 落笔要有反馈：中间那一条画在**这一段自己的下沿**，且只有下一段接着它、两段又是同一个样式时才画
   * （`betweenOf`）。两条控制放在一起——同样式的下一段画，不同样式的下一段不画，所以「下一段说不说」不该由这一击代答。
   */
  it("中间那一条只画在同样式的两段之间：同样式的下一段画，样式不同的下一段不画", () => {
    const doc = open(DOC);
    click(doc, "上下是样式说的", legPatch("insideH")); // 下一段「自己退回了底线」同为 Boxed
    click(doc, "整段没说过框", legPatch("insideH")); // 下一段是 Boxed，样式不同
    expect(drawnStyle(doc, "上下是样式说的")).toContain("background-image:linear-gradient(currentColor,currentColor)");
    expect(drawnStyle(doc, "整段没说过框")).not.toContain("background-image");
  });

  it("底框线与中间那一条各画一条：一条 CSS 边画不出两条，故用背景贴着 padding 盒的底", () => {
    const doc = open(DOC);
    click(doc, "上下是样式说的", legPatch("all"));
    const css = drawnStyle(doc, "上下是样式说的");
    expect(css).toContain("border-bottom:1px solid currentColor"); // 这一段自己说了 single(6)，样式借来的 double(12) 被它换掉
    expect(css).toContain("background-image:linear-gradient(currentColor,currentColor)");
    expect(css).toContain("background-size:100% 1px");
    // padding-bottom 只有一格，取两者里大的那一个，不各推一条把前一条挤掉。
    expect(css.match(/padding-bottom/g)).toHaveLength(1);
    // 那一格垫的是「间距 1 磅 + 那一条自己 1px」：只算间距的话，线会压在下一段的字上（变异 MU3 抓出来的空档）。
    expect(css).toContain("padding-bottom:2.333px");
  });

  /** 件里说了中间那一条而这一击没碰它：字节按原件带回，纸上照旧画得出来。 */
  it("件里那句 `w:between` 不动它时原样带回，读得出也画得出", () => {
    const bytes = makeChainDocx(
      [{ id: "Plain" }],
      [
        { pStyle: "Plain", pPr: '<w:pBdr><w:between w:val="single" w:sz="12" w:space="2" w:color="0070C0"/></w:pBdr>', text: "件里说了中间那条" },
        { pStyle: "Plain", text: "接着的那一段" },
      ],
    );
    const doc = open(bytes);
    expect(byText(doc, "件里说了中间那条").borders?.between).toEqual({ val: "single", size: 12, space: 2, color: "0070C0" });
    expect(boxOf(save(doc), "件里说了中间那条")).toBe('<w:pBdr><w:between w:val="single" w:sz="12" w:space="2" w:color="0070C0"/></w:pBdr>');
    expect(drawnStyle(doc, "件里说了中间那条")).toContain("background-image:linear-gradient(#0070C0,#0070C0)");
  });
});
