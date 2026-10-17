import { describe, expect, it } from "vitest";
import type { Paragraph } from "../src/core/docx.js";
import { WordDocument, carriedShape, syncParaSlots } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeChainDocx, makeParaPropsDocx } from "./helpers/make-paragraph.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const byText = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
/** The whole element, children and all — a border set *is* its children, so the opening tag proves nothing. */
const boxOf = (xml: string, needle: string): string =>
  /<w:pBdr\b[^>]*\/>|<w:pBdr\b[^>]*>[\s\S]*?<\/w:pBdr>/.exec(paraXml(xml, needle))?.[0] ?? "";
const shdOf = (xml: string, needle: string): string => /<w:shd[^>]*\/>/.exec(paraXml(xml, needle))?.[0] ?? "";
/**
 * The `style` the page puts on the paragraph carrying `needle`, or "" when it puts none. The paragraph's own
 * element is the only place its box and ground can show up, so this is what "the page draws it" means here.
 */
const drawnStyle = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<(?:div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>(?:(?!<\/(?:div|h[1-6])>)[\s\S])*<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return / style="([^"]*)"/.exec(m[0])?.[1] ?? "";
  }
  return "";
};

/**
 * 段落那一圈 `w:pBdr` 与那一层 `w:shd`（§17.3.2.36 / §17.3.2.33）：读进模型、画上页面、保存带回，以及**不该由
 * 段落说的时候别说**。
 *
 * 形状与受众都量自授权语料（`.scratch/pbdr-audience.ts`：46 份 docx 家族件、7,430 个正文段）：
 * - 26 枚 `w:pBdr` 里 16 枚把 top/left/bottom/right 四边全说成 `w:val="none" w:color="auto" w:sz="0"`，
 *   `w:space` 顶边底边写 1、左右写 4；24 枚在 `word/styles.xml`，正文段落自己写的 0 枚；
 * - 8 枚样式写一条 `w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"`；
 * - 段落底纹 30 枚（4 份件，全在 `word/document.xml`），全是 `w:val="clear" w:color="auto" w:fill="十六进制"`。
 *   `w:fill="auto"` 在段落层 0 例，但不是编出来的形状：同一批件的 `w:tcPr/w:shd` 有 1,600/1,646 枚那么写，
 *   那是文件自己说「要页面自己的地」的惯用写法，故这里按 token 带着走，不当「没说」（见 `ParaStyle.shading`）；
 * - `w:between`、`w:bar`、`w:shadow` 各 0 枚：读得出、不画（也还没有实据可照）。
 */
describe("docx 段落边框与底纹（w:pBdr / w:shd）", () => {
  /** A style that draws a box and a ground, one that refuses every side, and the four paragraphs they lend to. */
  const BOXED = makeChainDocx(
    [
      { id: "Plain" },
      {
        id: "Boxed",
        next: "Plain",
        pPr: '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>',
      },
      {
        id: "Refused",
        pPr:
          '<w:pBdr><w:top w:val="none" w:color="auto" w:sz="0" w:space="1"/><w:left w:val="none" w:color="auto" w:sz="0" w:space="4"/><w:bottom w:val="none" w:color="auto" w:sz="0" w:space="1"/><w:right w:val="none" w:color="auto" w:sz="0" w:space="4"/></w:pBdr>',
      },
    ],
    [
      { pStyle: "Boxed", text: "整圈是样式说的" },
      { pStyle: "Boxed", pPr: '<w:pBdr><w:bottom w:val="none" w:color="auto" w:sz="0" w:space="1"/></w:pBdr>', text: "只把那条线退回去" },
      { pStyle: "Refused", text: "四边都说不画" },
      { pStyle: "Boxed", pPr: '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>', text: "底纹说要页面的地" },
      { text: "什么都没说的" },
    ],
  );

  it("reads each side with the stroke, width, gap and colour the file gave it", () => {
    const own = open(makeParaPropsDocx());
    // `w:sz` is eighths of a point and `w:space` whole points; both keep their own scale on the way to a pixel.
    expect(byText(own, "整段自述").borders).toEqual({ bottom: { val: "single", size: 6, space: 1, color: "auto" } });
    expect(byText(own, "整段自述").shading).toBe("#FFF2CC");
  });

  it("takes the box and the ground from the style chain when the paragraph says nothing of its own", () => {
    const doc = open(BOXED);
    const p = byText(doc, "整圈是样式说的");
    expect(p.borders).toEqual({ bottom: { val: "single", size: 6, space: 1, color: "auto" } });
    expect(p.shading).toBe("#FFF2CC");
    // 借来的那一套与链上的是**同一个物件**：写侧与 `nextSplit` 都按引用认「这一段自己说过没有」。
    expect(p.borders).toBe(doc.paragraphStyleOf("Boxed")!.borders);
  });

  it("keeps a side the paragraph refused, instead of reading `none` as a silence", () => {
    const doc = open(BOXED);
    const p = byText(doc, "只把那条线退回去");
    // §17.7.2 一条属性一条属性地解：这一段退回去的是底线，样式那句 `w:shd` 照旧上身。
    expect(p.borders).toEqual({ bottom: { val: "none", size: 0, space: 1, color: "auto" } });
    expect(p.borders).not.toBe(doc.paragraphStyleOf("Boxed")!.borders);
    expect(p.shading).toBe("#FFF2CC");
  });

  it("draws the line it says and no line for the sides that say none", () => {
    const doc = open(BOXED);
    // 6 个 1/8 磅 = 1 CSS px，1 磅的间距 = 1.333 px，`w:color="auto"` 是当前的墨色。
    expect(drawnStyle(doc, "整圈是样式说的")).toContain("border-bottom:1px solid currentColor;padding-bottom:1.333px;background-color:#FFF2CC");
    expect(drawnStyle(doc, "只把那条线退回去")).not.toContain("border-bottom");
    expect(drawnStyle(doc, "只把那条线退回去")).toContain("background-color:#FFF2CC");
    // 判据是「这一边没上样式」，不是「style 属性是空串」：阶梯总给这一段带着字高，那一格跟框、底纹无关。
    for (const needle of ["四边都说不画", "什么都没说的"]) {
      expect(drawnStyle(doc, needle)).not.toMatch(/border-/);
      expect(drawnStyle(doc, needle)).not.toMatch(/background/);
    }
  });

  /**
   * 一句 `w:fill="auto"` 说的是「这一段要页面自己的地」，与「这一段没说底纹」是两句话（§85 那条教训的第 N 次
   * 上身）：折成没说不但页上会把样式那层地铺进一段明令不要地的段落，写侧的 `moved()` 也再看不出这一段改过什么。
   * 段落层 0 例、`w:tcPr` 层 1,600 例——受众是这一击的，不是件里已有的。
   */
  it("keeps a stated `w:fill=\"auto\"` a refusal of the style's ground, not a silence", () => {
    const p = byText(open(BOXED), "底纹说要页面的地");
    expect(p.shading).toBe("auto");
    const drawn = drawnStyle(open(BOXED), "底纹说要页面的地");
    expect(drawn).not.toMatch(/background/);
    // 框照旧：§17.7.2 一条属性一条属性地解，这一段只退了底纹。
    expect(drawn).toContain("border-bottom:1px solid currentColor");
  });

  it("carries the element's own bytes through a save that touched nothing", () => {
    const read = part(makeParaPropsDocx());
    const out = save(open(makeParaPropsDocx()));
    expect(boxOf(out, "整段自述")).toBe(boxOf(read, "整段自述"));
    expect(shdOf(out, "整段自述")).toBe(shdOf(read, "整段自述"));
  });

  /**
   * 一件里 24 枚 `w:pBdr` 在 `word/styles.xml`、正文段落自己写的 0 枚（`.scratch/pbdr-audience.ts`）：把链上借来
   * 的那个盒子抄到段落身上，每一枚都是替文件新编的一句样式的话。动了别的属性也不该顺手把它认领过来。
   */
  it("does not make a paragraph claim the box its style drew", () => {
    const doc = open(BOXED);
    const p = byText(doc, "整圈是样式说的");
    p.align = "center";
    syncParaSlots(p);
    const out = save(doc);
    expect(boxOf(out, "整圈是样式说的")).toBe("");
    expect(shdOf(out, "整圈是样式说的")).toBe("");
    expect(paraXml(out, "整圈是样式说的")).toContain('<w:jc w:val="center"/>');
    // 保存回来的件再读一次，仍然是「样式说的、段落没说」，且画出来还是那条线。
    const again = open(writeDocx(doc));
    expect(byText(again, "整圈是样式说的").borders).toBe(again.paragraphStyleOf("Boxed")!.borders);
    expect(drawnStyle(again, "整圈是样式说的")).toContain("border-bottom:1px solid currentColor");
  });

  /**
   * 拒绝一条格式改动时，框与底纹与那几项槽位族的开关是同一件事：记录没说的就是当时没说，得把手里的清掉、
   * 向样式链要回来。不清，被否决的那一句还挂在字段上——页上画错，`moved()` 也再看不出这一段改过什么。
   * 记录那句 `w:pPr` 是**整套**改动前的属性，所以它自己的字节（`prevSlots`）整枚替掉段落手里那一套：被否决的
   * 那两枚跟着走，不是留在原地不动。
   */
  it("拒绝那条改动：记录没说的那一圈与那一层回到样式链那份", () => {
    const doc = open(
      makeChainDocx(
        [
          {
            id: "Boxed",
            pPr: '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>',
          },
        ],
        [
          {
            pStyle: "Boxed",
            text: "框与底纹是刚改的",
            pPr:
              '<w:pBdr><w:bottom w:val="none" w:color="auto" w:sz="0" w:space="1"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="E2EFDA"/>' +
              '<w:pPrChange w:id="31" w:author="林晓" w:date="2026-03-01T08:00:00Z"><w:pPr><w:pStyle w:val="Boxed"/><w:jc w:val="both"/></w:pPr></w:pPrChange>',
          },
        ],
      ),
    );
    const p = byText(doc, "框与底纹是刚改的");
    expect(p.borders).toEqual({ bottom: { val: "none", size: 0, space: 1, color: "auto" } });
    expect(p.shading).toBe("#E2EFDA");
    expect(doc.resolveRevisions("reject", 31)).toBe(1);
    // 记录只写了样式名与对齐：链上的那条底线与那层地才是这一段原本的的说法，且按引用还回来（这一段没说）。
    expect(p.borders).toBe(doc.paragraphStyleOf("Boxed")!.borders);
    expect(p.shading).toBe("#FFF2CC");
    expect(drawnStyle(doc, "框与底纹是刚改的")).toContain("border-bottom:1px solid currentColor");
    expect(drawnStyle(doc, "框与底纹是刚改的")).toContain("background-color:#FFF2CC");
    // 被否决的那两枚不在保存出来的件里，也不会有哪一枚被字段替段落新签一句。
    const out = save(doc);
    expect(boxOf(out, "框与底纹是刚改的")).toBe("");
    expect(shdOf(out, "框与底纹是刚改的")).toBe("");
    expect(paraXml(out, "框与底纹是刚改的")).toContain('<w:pStyle w:val="Boxed"/>');
    // 存下去再读回来仍然是这一套：字段、快照与件说得一致，下一次保存不会再替段落补一句。
    const again = open(writeDocx(doc));
    expect(byText(again, "框与底纹是刚改的").shading).toBe("#FFF2CC");
    expect(boxOf(save(again), "框与底纹是刚改的")).toBe("");
  });

  it("spells a box a paragraph said for itself when it has no `w:pPr` slots to fall back on", () => {
    const doc = open(BOXED);
    const split = doc.nextSplit(byText(doc, "只把那条线退回去"))!;
    expect(split.styleId).toBe("Plain");
    // 自己那句「这条线不要」跟着走；样式 Boxed 的底纹是 Boxed 说的话，Plain 没说，所以留在样式那里。
    expect(split.borders).toEqual({ bottom: { val: "none", size: 0, space: 1, color: "auto" } });
    expect(split.shading).toBeUndefined();
    expect(split.pPrSlots).toBeUndefined();
    // 新段没有文字，所以这里对着整件那一段认——件里只剩这一段，一枚 `w:pBdr` 也不可能有第二个来源。
    doc.blocks = [split];
    const out = save(doc);
    // 拼法照件：一边一枚 `<w:side>`、按 `CT_PBdr` 的边序；`none` 照样写出来，因为「说了不要」是这一段的句子。
    expect(/<w:pBdr\b[^>]*>[\s\S]*?<\/w:pBdr>/.exec(out)?.[0] ?? "").toBe(
      '<w:pBdr><w:bottom w:val="none" w:sz="0" w:space="1" w:color="auto"/></w:pBdr>',
    );
    expect(/<w:shd[^>]*\/>/.exec(out)?.[0] ?? "").toBe("");
    // 一句 `w:fill="auto"` 由字段签出来时，token 原样回填，不会被大写成一个 schema 里没有的词。
    const token = { ...split, shading: "auto" } as Paragraph;
    const withToken = open(BOXED);
    withToken.blocks = [token];
    expect(/<w:shd[^>]*\/>/.exec(save(withToken))?.[0] ?? "").toBe('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
  });

  /**
   * 回车拆段带形状那一条：清单由 `carriedShape` 从模型的字段自己拼。手数的那一版漏过槽位族四项（三条保持规则、
   * 不吸网格）与 `w:leftChars`，现在连着框与底纹一次带走——这一段没有 `w:pPr` 槽位，字段是它唯一的说法。
   */
  it("carries every paragraph field, the slot-only family included, to the half Enter makes", () => {
    const p = {
      kind: "paragraph",
      styleId: "Boxed",
      align: "center",
      indentLeftChars: 200,
      keepNext: true,
      snapToGrid: false,
      borders: { bottom: { val: "single", size: 6, space: 1, color: "auto" } },
      shading: "#FFF2CC",
      inlines: [],
    } as Paragraph;
    expect(carriedShape(p)).toEqual({
      styleId: "Boxed",
      align: "center",
      indentLeftChars: 200,
      keepNext: true,
      snapToGrid: false,
      borders: { bottom: { val: "single", size: 6, space: 1, color: "auto" } },
      shading: "#FFF2CC",
    });
    // 「这一段之前必须断一页」不是版式而是断点，`fresh` 与 `nextSplit` 都把它摘掉。
    expect(carriedShape({ ...p, pageBreakBefore: true }).pageBreakBefore).toBe(true);
  });
});
