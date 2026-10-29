import { describe, expect, it } from "vitest";
import type { Paragraph } from "../src/core/docx.js";
import { LINE_ONE, carriedShape, lineGridPitchTwips, linePitchTwips, paraSpacingPitch, paraStyleOf, syncParaSlots, WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { childSpans, elementContent, OpcPackage, text } from "../src/core/opc.js";
import { makeChainDocx, makeGridRevisionDocx, makeSpacingDocx, makeWrapRevisionDocx } from "./helpers/make-paragraph.js";
import { renderWord } from "../src/core/render.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const findPara = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
/** 那一段自己写下的 `w:pPr` 整句：形状对不对，看字节而不是看字段。 */
const pPrOf = (xml: string, needle: string): string => {
  const para = [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
  const slot = childSpans(elementContent(para)).find((k) => k.tag === "pPr");
  return slot ? elementContent(slot.xml) : "";
};
const spacingOf = (xml: string, needle: string): string => /<w:spacing[^>]*\/>/.exec(pPrOf(xml, needle))?.[0] ?? "";
/** 段落标记 `w:rPr` 里那一枚同名的 `w:spacing`（字符间距），它跟段距不是一个东西。 */
const markSpacingOf = (xml: string, needle: string): string =>
  /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(pPrOf(xml, needle))?.[0] ?? "";

/**
 * 段前／段后／行距的**写出**侧。读与画早在 #72 与 #130~#132 就归了文件，这一族此前没有任何一条命令能改它。
 *
 * 形状不是编的：`.scratch/spacing-shape.ts` 扫过 201 份 WPS 写出的件的 10,248 枚 `w:spacing`——
 * - 属性先后只有一簇，`before, beforeLines, beforeAutospacing, after, afterLines, afterAutospacing, line, lineRule`；
 * - `w:line` 与 `w:lineRule` 从来没单独出现过（7,810 枚成对，0 枚只写一边）；
 * - 同一个间距用两笔写（缇 + 1/100 行）的有段前 2,007 枚、段后 1,850 枚，商几乎全落 312 缇一档，
 *   1,865 枚前后同档、0 枚不一致；另有 33+91 枚只说 1/100 行而不给缇，那种档量不出来。
 * 这些数与「谁能动、谁不该动」的口径写在 `syncParaSlots` 的头注里。
 *
 * 单位换算按件自己的那一档：改缇时把 1/100 行 那一笔按同一档重算（`w:before="312" w:beforeLines="100"` 改 624
 * 缇 → `w:beforeLines="200"`）；件没说档就不写这一笔，而不是照 App 的行高造一个。
 */
describe("docx 段距与行距的写出", () => {
  const DOC = makeSpacingDocx(
    [
      { spacing: 'w:before="120" w:after="60" w:line="360" w:lineRule="auto"', markRPr: '<w:rPr><w:spacing w:val="20"/></w:rPr>', text: "两笔都说的" },
      { spacing: 'w:before="312" w:beforeLines="100" w:after="312" w:afterLines="100"', text: "按行写档的" },
      { spacing: 'w:before="312"', text: "只说缇的" },
      { spacing: 'w:before="0" w:beforeAutospacing="1"', text: "让软件自定段距的" },
      { spacing: 'w:line="578" w:lineRule="exact"', text: "固定值的" },
      { spacing: 'w:line="363" w:lineRule="atLeast"', text: "最小值的" },
      { pStyle: "Lent", spacing: 'w:before="0"', text: "行距靠样式的" },
      { text: "什么都没说的" },
    ],
    '<w:spacing w:line="480" w:lineRule="auto"/>',
  );

  it("一件没动过地存一遍，那一句一个字节都不变", () => {
    const round = save(open(DOC));
    for (const needle of ["两笔都说的", "按行写档的", "只说缇的", "让软件自定段距的", "固定值的", "最小值的", "行距靠样式的", "什么都没说的"]) {
      expect(spacingOf(round, needle)).toBe(spacingOf(part(DOC), needle));
    }
    // 整段 `w:pPr` 也一样：没写字节的段落不该因为存盘而被重新拼一遍。
    expect(pPrOf(round, "两笔都说的")).toBe(pPrOf(part(DOC), "两笔都说的"));
  });

  it("没动过那三个数时，一个字节都不写", () => {
    const doc = open(DOC);
    const p = findPara(doc, "按行写档的");
    // 把字段按**读到的原值**交回去（对话框里没人改过格子就是这个形状），`moved()` 认得出这不是用户的话。
    p.spacingBefore = 312;
    p.spacingAfter = 312;
    syncParaSlots(p);
    expect(spacingOf(save(doc), "按行写档的")).toBe('<w:spacing w:before="312" w:beforeLines="100" w:after="312" w:afterLines="100"/>');
  });

  it("改段前时按件自己那一档把 1/100 行 也带走，段后原样留在后面", () => {
    const doc = open(DOC);
    const p = findPara(doc, "按行写档的");
    p.spacingBefore = 624;
    syncParaSlots(p);
    expect(spacingOf(save(doc), "按行写档的")).toBe('<w:spacing w:before="624" w:beforeLines="200" w:after="312" w:afterLines="100"/>');
  });

  it("件没给那一档时不写 1/100 行，也不凭空造一个", () => {
    const doc = open(DOC);
    const p = findPara(doc, "只说缇的");
    p.spacingBefore = 624;
    syncParaSlots(p);
    expect(spacingOf(save(doc), "只说缇的")).toBe('<w:spacing w:before="624"/>');
  });

  it("用户说了一个数，就把「让软件自定」那一笔按件里的灭法写掉", () => {
    const doc = open(DOC);
    const p = findPara(doc, "让软件自定段距的");
    expect(p.spacingBefore).toBe(0);
    p.spacingBefore = 300;
    syncParaSlots(p);
    expect(spacingOf(save(doc), "让软件自定段距的")).toBe('<w:spacing w:before="300" w:beforeAutospacing="0"/>');
  });

  it("`w:line` 与 `w:lineRule` 同进同出，且排在缇那几笔之后", () => {
    const doc = open(DOC);
    const p = findPara(doc, "两笔都说的");
    p.lineSpacing = 480;
    p.lineSpacingRule = "auto";
    syncParaSlots(p);
    expect(spacingOf(save(doc), "两笔都说的")).toBe('<w:spacing w:before="120" w:after="60" w:line="480" w:lineRule="auto"/>');
    // 段落标记里那枚 `w:spacing w:val` 说的是字符间距：整枚留在 `w:rPr` 的字节里，段距这条分支碰不到它。
    expect(markSpacingOf(save(doc), "两笔都说的")).toBe('<w:rPr><w:spacing w:val="20"/></w:rPr>');
  });

  it("固定值换成倍数时，单位跟着规则一起换", () => {
    const doc = open(DOC);
    const p = findPara(doc, "固定值的");
    p.lineSpacing = LINE_ONE;
    p.lineSpacingRule = "auto";
    syncParaSlots(p);
    expect(spacingOf(save(doc), "固定值的")).toBe('<w:spacing w:line="240" w:lineRule="auto"/>');

    const other = open(DOC);
    const q = findPara(other, "最小值的");
    q.lineSpacing = 578;
    q.lineSpacingRule = "exact";
    syncParaSlots(q);
    expect(spacingOf(save(other), "最小值的")).toBe('<w:spacing w:line="578" w:lineRule="exact"/>');
  });

  it("只动行距时，段落自己写过的段前留在原位；从样式借来的那个数不被抄进段落", () => {
    const doc = open(DOC);
    const p = findPara(doc, "行距靠样式的");
    expect(p.lineSpacing).toBe(480);
    expect(p.asRead?.lineSpacing).toBe(480);
    p.lineSpacing = 240;
    p.lineSpacingRule = "auto";
    syncParaSlots(p);
    // `w:before="0"` 是这一段自己说的（与「没说」是两件事，#132 钉过），改行距不能把它顶掉。
    expect(spacingOf(save(doc), "行距靠样式的")).toBe('<w:spacing w:before="0" w:line="240" w:lineRule="auto"/>');
  });

  it("什么都没写过的段落，改段前才第一次写下这一句", () => {
    const doc = open(DOC);
    const p = findPara(doc, "什么都没说的");
    expect(p.pPrSlots).toBeUndefined();
    p.spacingAfter = 120;
    p.lineSpacing = 480;
    p.lineSpacingRule = "auto";
    syncParaSlots(p);
    // 没有槽的段落由 `ownParaStyle` 按字段拼（`w:line` 配着默认规则一起写）。
    expect(pPrOf(save(doc), "什么都没说的")).toBe('<w:spacing w:after="120" w:line="480" w:lineRule="auto"/>');
  });

  it("改过的那一句存下去再读回来，第二次存不再改字节", () => {
    const doc = open(DOC);
    const p = findPara(doc, "按行写档的");
    p.spacingBefore = 624;
    syncParaSlots(p);
    const once = writeDocx(doc);
    const again = part(writeDocx(WordDocument.openBytes(once)));
    expect(spacingOf(again, "按行写档的")).toBe('<w:spacing w:before="624" w:beforeLines="200" w:after="312" w:afterLines="100"/>');
  });

  it("那一档量不出来时 `paraSpacingPitch` 不猜", () => {
    const doc = open(DOC);
    expect(paraSpacingPitch(findPara(doc, "按行写档的"))).toBe(312);
    expect(paraSpacingPitch(findPara(doc, "只说缇的"))).toBeUndefined();
    expect(paraSpacingPitch(findPara(doc, "什么都没说的"))).toBeUndefined();
    expect(linePitchTwips("0", "100")).toBeUndefined();
    expect(linePitchTwips("312", "0")).toBeUndefined();
  });
});

/** 那一段自己写下的 `w:snapToGrid`——按 `w:pPr` 的直接孩子找，段落标记 `w:rPr` 里那枚同名的不算（201 份里 19 枚挂在那儿）。 */
const snapOf = (xml: string, needle: string): string => childSpans(pPrOf(xml, needle)).find((k) => k.tag === "snapToGrid")?.xml ?? "";

/**
 * 「如果定义了文档网格，则对齐到网格」的**写出**侧（§17.3.1.87 `w:snapToGrid`）。读与画早在 #130 就归了文件，
 * 这一枚开关此前没有任何一条命令能改它。
 *
 * 拼法与位置都不是编的（`.scratch/snap-grid-audience.ts`，同一批 201 份 WPS 写出的件）：
 * - `document.xml` 3,466 枚里 `p>w:pPr` 直接孩子 3,394 枚；另有 72 枚挂在 `w:rPr`（`pPr>rPr` 19、`r>rPr` 53），
 *   §17.3.1.87 不放它在那儿，读取不看、写出也不碰。
 * - 说法只有两种：裸写 1,416 枚（= 说开，缺省也是开）、`w:val="0"` 2,381 枚（= 说关）。**没有一枚**写 `"1"`、
 *   `"false"` 或 `"none"`，所以关掉必须写 `"0"`，打开必须裸写。
 * - 正文段（13,426 段）里说开 1,353、说关 672，而这 672 段**全在定义了行网格的节下**（201 份里 200 份首节写
 *   `w:docGrid w:type="lines"`）。
 * - `w:pPr` 里它的前后兄弟固定是 `adjustRightInd → snapToGrid → spacing|jc|ind`（2,185 + 597 + 270 枚），
 *   与 `PPR_ORDER` 同一条，所以新写的那一枚必须落在 `w:spacing` 之前。
 */
describe("docx 「与文档网格对齐」的写出", () => {
  const GRID = makeSpacingDocx(
    [
      { snapToGrid: 'w:val="0"', spacing: 'w:line="240" w:lineRule="auto"', text: "说不吸的" },
      { snapToGrid: "", text: "裸写说吸的" },
      { spacing: 'w:before="120"', text: "没提网格的" },
      { markRPr: '<w:rPr><w:snapToGrid w:val="0"/></w:rPr>', text: "标记里有一枚的" },
      { pStyle: "Lent", text: "网格靠样式的" },
      { text: "整段没 pPr 的" },
    ],
    '<w:snapToGrid w:val="0"/>',
    '<w:docGrid w:type="lines" w:linePitch="312"/>',
  );

  it("件里那两种拼法一件没动时原样带回", () => {
    const round = save(open(GRID));
    expect(snapOf(round, "说不吸的")).toBe('<w:snapToGrid w:val="0"/>');
    expect(snapOf(round, "裸写说吸的")).toBe("<w:snapToGrid/>");
    expect(snapOf(round, "没提网格的")).toBe("");
    // 段落标记 `w:rPr` 里那枚同名的开关不是段落属性：整枚留在原地，一个字没动。
    expect(snapOf(round, "标记里有一枚的")).toBe("");
    expect(/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(pPrOf(round, "标记里有一枚的"))?.[0]).toBe('<w:rPr><w:snapToGrid w:val="0"/></w:rPr>');
  });

  it("把读到的值原样交回去不是用户的话：一个字节都不写", () => {
    const doc = open(GRID);
    const p = findPara(doc, "说不吸的");
    expect(p.snapToGrid).toBe(false);
    p.snapToGrid = false;
    syncParaSlots(p);
    expect(pPrOf(save(doc), "说不吸的")).toBe(pPrOf(part(GRID), "说不吸的"));
  });

  it("段落没说过而用户关掉时，新写的一枚按件的拼法落在 `w:spacing` 之前", () => {
    const doc = open(GRID);
    const p = findPara(doc, "没提网格的");
    expect(p.snapToGrid).toBeUndefined();
    p.snapToGrid = false;
    syncParaSlots(p);
    expect(pPrOf(save(doc), "没提网格的")).toBe('<w:snapToGrid w:val="0"/><w:spacing w:before="120"/>');
    // 同一份件里另一段没被这一击碰到：它的 `w:spacing` 与原字节一模一样。
    expect(pPrOf(save(doc), "裸写说吸的")).toBe(pPrOf(part(GRID), "裸写说吸的"));
  });

  it("打开时裸写，不写 `w:val=\"1\"`；同一句里别的属性留在原位", () => {
    const doc = open(GRID);
    const p = findPara(doc, "说不吸的");
    p.snapToGrid = true;
    syncParaSlots(p);
    expect(pPrOf(save(doc), "说不吸的")).toBe('<w:snapToGrid/><w:spacing w:line="240" w:lineRule="auto"/>');
    // 段落标记里那枚同名的开关整枚照旧——它说的是断行符自己吸不吸，不是这一段的事。
    expect(/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(pPrOf(save(doc), "标记里有一枚的"))?.[0]).toBe('<w:rPr><w:snapToGrid w:val="0"/></w:rPr>');
  });

  it("样式链借来的那一枚不抄进段落；用户改口说开时段落自己认领", () => {
    const doc = open(GRID);
    const p = findPara(doc, "网格靠样式的");
    // 字段读的是解析过样式链的那一份，`asRead` 存的也是它——写出侧对着这一份判「是不是这一段自己改的口」。
    expect(p.snapToGrid).toBe(false);
    expect(p.asRead?.snapToGrid).toBe(false);
    expect(p.pPrSlots?.map((s) => s.tag)).toEqual(["pStyle"]);
    // 没动过 → 「不吸」还是样式的话，段落那一句里只有它自己点的名。
    expect(pPrOf(save(doc), "网格靠样式的")).toBe('<w:pStyle w:val="Lent"/>');
    const other = open(GRID);
    const q = findPara(other, "网格靠样式的");
    q.snapToGrid = true;
    syncParaSlots(q);
    expect(pPrOf(save(other), "网格靠样式的")).toBe('<w:pStyle w:val="Lent"/><w:snapToGrid/>');
  });

  it("整段没有 `w:pPr` 时由字段拼这一句（那条路也得接上），存下去再读回来不再改字节", () => {
    const doc = open(GRID);
    const p = findPara(doc, "整段没 pPr 的");
    expect(p.pPrSlots).toBeUndefined();
    p.snapToGrid = false;
    const once = writeDocx(doc);
    expect(pPrOf(part(once), "整段没 pPr 的")).toBe('<w:snapToGrid w:val="0"/>');
    const again = open(once);
    expect(findPara(again, "整段没 pPr 的").snapToGrid).toBe(false);
    expect(pPrOf(part(writeDocx(again)), "整段没 pPr 的")).toBe('<w:snapToGrid w:val="0"/>');
  });

  it("`lineGridPitchTwips` 只认件说得出间距的那种网格", () => {
    const withGrid = open(GRID);
    expect(lineGridPitchTwips(withGrid.sections[0]!)).toBe(312);
    const noGrid = open(makeSpacingDocx([{ text: "没有网格的" }]));
    expect(lineGridPitchTwips(noGrid.sections[0]!)).toBeUndefined();
    const otherType = open(makeSpacingDocx([{ text: "别档网格的" }], "", '<w:docGrid w:type="default" w:linePitch="312"/>'));
    expect(lineGridPitchTwips(otherType.sections[0]!)).toBeUndefined();
    const noPitch = open(makeSpacingDocx([{ text: "网格没间距的" }], "", '<w:docGrid w:type="lines"/>'));
    expect(lineGridPitchTwips(noPitch.sections[0]!)).toBeUndefined();
  });

  /**
   * 拒绝一条改动时，这一枚开关与别项是同一件事：记录没说的就是当时没说。`restoreParaStyle` 以前只清
   * `PARA_STYLE_KEYS` 那本账，而开关不在里头——拒绝完字段还顶着被否决的那一句，画错，下一次存盘还会替段落把它
   * 写回件里。三条保持规则原来有同一个口子，这一并补上（同一条判据、同一个循环）。
   */
  it("拒绝那条改动：记录没说的那一枚被清回去，说过的那一枚按记录的字节还回来", () => {
    const doc = open(makeGridRevisionDocx());
    const added = findPara(doc, "网格是刚加上的");
    expect(added.snapToGrid).toBe(true);
    expect(doc.resolveRevisions("reject", 21)).toBe(1);
    // 记录那一句没有 `w:snapToGrid`，样式链也没借，所以这一段回到「自己没说」（缺省仍是吸，与「说了不吸」两件事）。
    expect(added.snapToGrid).toBeUndefined();
    expect(snapOf(save(doc), "网格是刚加上的")).toBe("");
    // 存下去再读回来也一样：字段与快照都说「没说」，`moved()` 认不出这一击，不会替文件补一句。
    expect(snapOf(save(open(writeDocx(doc))), "网格是刚加上的")).toBe("");

    const removed = findPara(doc, "网格是刚被删掉的");
    expect(removed.snapToGrid).toBeUndefined();
    expect(doc.resolveRevisions("reject", 22)).toBe(1);
    expect(removed.snapToGrid).toBe(false);
    expect(pPrOf(save(doc), "网格是刚被删掉的")).toBe('<w:snapToGrid w:val="0"/><w:spacing w:line="300" w:lineRule="auto"/>');
  });

  it("回车拆段：上一段自己那句「不吸」跟着走，从样式借来的那句留在样式那里", () => {
    const own = open(makeChainDocx([{ id: "Head", next: "Body" }, { id: "Body" }], [{ pStyle: "Head", pPr: '<w:snapToGrid w:val="0"/>', text: "自己说的" }]));
    const split = own.nextSplit(findPara(own, "自己说的")!)!;
    expect(split.styleId).toBe("Body");
    expect(split.snapToGrid).toBe(false);

    const lent = open(
      makeChainDocx([{ id: "Head", next: "Body", pPr: '<w:snapToGrid w:val="0"/>' }, { id: "Body" }], [{ pStyle: "Head", text: "借来的" }]),
    );
    const borrowed = lent.nextSplit(findPara(lent, "借来的")!)!;
    // 「不吸」是 Head 说的话，Body 没说，所以新段不该认领它——写出来就是替文件在一段没说过的话上签了名。
    expect(findPara(lent, "借来的").snapToGrid).toBe(false);
    expect(borrowed.snapToGrid).toBeUndefined();
  });
});

/** 那一段自己写下的 `w:wordWrap`——按 `w:pPr` 的直接孩子找，段落标记 `w:rPr` 里那枚同名的不算。 */
const wrapOf = (xml: string, needle: string): string => childSpans(pPrOf(xml, needle)).find((k) => k.tag === "wordWrap")?.xml ?? "";

/**
 * 「允许西文在单词中间换行」的**读、画、写**（§17.4.77 `w:wordWrap`，WPS 段落对话框「中文版式」那一页的第一枚）。
 *
 * 受众是量的，不是照 `w:snapToGrid` 类推的（`.scratch/cjk-typo-audience.py`，授权口径 46 件）：正文 7,357 枚
 * `w:pPr` 里 200 枚说了这一格、**枚枚裸写**（9 份模板名），**说关 0 枚**，`word/styles.xml` 与 `w:docDefaults`
 * 各 0 枚。件里从没写过 `w:val="0"`，那一笔是本机对话框取消勾选时按 §17.4.77 的 on/off 写法补的，测试把两种拼法
 * 都钉住：读侧「说了允许」与「没说」在页面上画同一张纸（缺省是不许），但盘上是两句话。
 * 位置按 `PPR_ORDER`：`w:wordWrap` 在 `w:kinsoku` 之后、`w:snapToGrid` 之前，所以新写的一枚落在网格那一枚与
 * `w:spacing` 之前。
 */
describe("docx 「允许西文在单词中间换行」", () => {
  const WRAP = makeSpacingDocx(
    [
      { wordWrap: "", spacing: 'w:line="240" w:lineRule="auto"', text: "裸写说允许的" },
      { wordWrap: 'w:val="0"', text: "明着说不许的" },
      { wordWrap: "", snapToGrid: 'w:val="0"', text: "两枚都说过的" },
      { spacing: 'w:before="120"', text: "没提换行的" },
      { markRPr: '<w:rPr><w:wordWrap/></w:rPr>', text: "标记里有一枚的" },
      { pStyle: "Lent", text: "允许靠样式的" },
      { text: "整段没 pPr 的" },
    ],
    "<w:wordWrap/>",
  );

  it("裸写算「说了允许」，`w:val=\"0\"` 算说了不许，没提才是没说", () => {
    const doc = open(WRAP);
    expect(findPara(doc, "裸写说允许的").wordWrap).toBe(true);
    expect(findPara(doc, "明着说不许的").wordWrap).toBe(false);
    expect(findPara(doc, "没提换行的").wordWrap).toBeUndefined();
    // 段落标记 `w:rPr` 里那枚同名的开关不是段落属性：那一句说的是断行符自己，不替这一段答话。
    expect(findPara(doc, "标记里有一枚的").wordWrap).toBeUndefined();
  });

  it("只有说了允许的那一段改断词；`w:val=\"0\"` 与没提画同一张纸", () => {
    const doc = open(makeSpacingDocx([{ wordWrap: "", text: "说了的" }, { wordWrap: 'w:val="0"', text: "说了不许的" }, { text: "没说的" }]));
    const html = renderWord(doc);
    expect([...html.matchAll(/word-break:break-all/g)]).toHaveLength(1);
  });

  it("件里那两种拼法一件没动时原样带回，标记里那枚一个字节不动", () => {
    const round = save(open(WRAP));
    expect(wrapOf(round, "裸写说允许的")).toBe("<w:wordWrap/>");
    expect(wrapOf(round, "明着说不许的")).toBe('<w:wordWrap w:val="0"/>');
    expect(pPrOf(round, "两枚都说过的")).toBe('<w:wordWrap/><w:snapToGrid w:val="0"/>');
    expect(wrapOf(round, "没提换行的")).toBe("");
    expect(wrapOf(round, "标记里有一枚的")).toBe("");
    expect(/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(pPrOf(round, "标记里有一枚的"))?.[0]).toBe("<w:rPr><w:wordWrap/></w:rPr>");
  });

  it("把读到的值原样交回去不是用户的话：一个字节都不写", () => {
    const doc = open(WRAP);
    const p = findPara(doc, "裸写说允许的");
    expect(p.wordWrap).toBe(true);
    p.wordWrap = true;
    syncParaSlots(p);
    expect(pPrOf(save(doc), "裸写说允许的")).toBe(pPrOf(part(WRAP), "裸写说允许的"));
  });

  it("段落没说过而用户勾上时，新写的一枚裸写并排在 `w:spacing` 之前", () => {
    const doc = open(WRAP);
    const p = findPara(doc, "没提换行的");
    expect(p.wordWrap).toBeUndefined();
    p.wordWrap = true;
    syncParaSlots(p);
    expect(pPrOf(save(doc), "没提换行的")).toBe("<w:wordWrap/><w:spacing w:before=\"120\"/>");
  });

  it("勾掉时写 `w:val=\"0\"`，不写 `w:val=\"1\"`；同一句里别的属性留在原位", () => {
    const doc = open(WRAP);
    const p = findPara(doc, "裸写说允许的");
    p.wordWrap = false;
    syncParaSlots(p);
    expect(pPrOf(save(doc), "裸写说允许的")).toBe('<w:wordWrap w:val="0"/><w:spacing w:line="240" w:lineRule="auto"/>');
    // 新那一枚排在网格那一枚之前，两枚各说各的，互不认领。
    const q = findPara(doc, "两枚都说过的");
    q.wordWrap = false;
    syncParaSlots(q);
    expect(pPrOf(save(doc), "两枚都说过的")).toBe('<w:wordWrap w:val="0"/><w:snapToGrid w:val="0"/>');
    expect(/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(pPrOf(save(doc), "标记里有一枚的"))?.[0]).toBe("<w:rPr><w:wordWrap/></w:rPr>");
  });

  it("样式链借来的那一枚不抄进段落；用户改口说不许时段落自己认领", () => {
    const doc = open(WRAP);
    const p = findPara(doc, "允许靠样式的");
    expect(p.wordWrap).toBe(true);
    expect(p.asRead?.wordWrap).toBe(true);
    expect(p.pPrSlots?.map((s) => s.tag)).toEqual(["pStyle"]);
    expect(pPrOf(save(doc), "允许靠样式的")).toBe('<w:pStyle w:val="Lent"/>');
    const other = open(WRAP);
    const q = findPara(other, "允许靠样式的");
    q.wordWrap = false;
    syncParaSlots(q);
    expect(pPrOf(save(other), "允许靠样式的")).toBe('<w:pStyle w:val="Lent"/><w:wordWrap w:val="0"/>');
  });

  it("整段没有 `w:pPr` 时由字段拼这一句，存下去再读回来不再改字节", () => {
    const doc = open(WRAP);
    const p = findPara(doc, "整段没 pPr 的");
    expect(p.pPrSlots).toBeUndefined();
    p.wordWrap = true;
    const once = writeDocx(doc);
    expect(pPrOf(part(once), "整段没 pPr 的")).toBe("<w:wordWrap/>");
    const again = open(once);
    expect(findPara(again, "整段没 pPr 的").wordWrap).toBe(true);
    expect(pPrOf(part(writeDocx(again)), "整段没 pPr 的")).toBe("<w:wordWrap/>");
  });

  it("拒绝那条改动：记录没说的那一枚被清回去，说过的那一枚按记录的字节还回来", () => {
    const doc = open(makeWrapRevisionDocx());
    const added = findPara(doc, "换行是刚加上的");
    expect(added.wordWrap).toBe(true);
    expect(doc.resolveRevisions("reject", 31)).toBe(1);
    // 记录那一句没有 `w:wordWrap`，样式链也没借，所以这一段回到「自己没说」（缺省是不许，与「说了不许」两件事）。
    expect(added.wordWrap).toBeUndefined();
    expect(wrapOf(save(doc), "换行是刚加上的")).toBe("");
    expect(wrapOf(save(open(writeDocx(doc))), "换行是刚加上的")).toBe("");

    const removed = findPara(doc, "换行是刚被删掉的");
    expect(removed.wordWrap).toBeUndefined();
    expect(doc.resolveRevisions("reject", 32)).toBe(1);
    expect(removed.wordWrap).toBe(false);
    expect(pPrOf(save(doc), "换行是刚被删掉的")).toBe('<w:wordWrap w:val="0"/><w:spacing w:line="300" w:lineRule="auto"/>');
  });

  /**
   * 格式改动记录带的那本账必须是**全账**，也就是 `carriedShape` 而不是 `paraStyleOf`。这一条钉记账的物件本身，
   * 不钉某一条命令：那本账少列一项，签记录时「改之前那一句」就缺那一项，拒绝之后字段与被还回来的字节不同步——
   * 纸上还画着被否决的那一句。#134 的 `w:snapToGrid` 与 #215 的 `w:wordWrap` 走的正是这条路，
   * `.scratch/wrap-dom.ts` 量到的失败形状是：拒绝后 `wordWrap` 留 `undefined`、纸回 `normal`，而盘上写 `<w:wordWrap/>`。
   *
   * `paraStyleOf` 那半本不是遗留，是写出侧判「哪一句归段落自己写」的依据（`moved()`/`ownParaStyle`）：把槽位族
   * 八项补进去，`syncParaSlots` 就会把样式链借来的话替段落重写一遍。两本账各有各的主人，所以两个都钉住。
   */
  it("记录带的是全账：槽位族那八项在 `carriedShape` 里，写出侧那半本按设计不含", () => {
    const doc = open(
      makeChainDocx([{ id: "Body" }], [
        {
          pPr:
            '<w:keepNext/><w:keepLines/><w:widowControl w:val="0"/><w:wordWrap/><w:snapToGrid w:val="0"/>' +
            '<w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs>' +
            '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>' +
            '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>',
          text: "八项都说过的",
        },
        { text: "什么都没说的" },
      ]),
    );
    const eight = carriedShape(findPara(doc, "八项都说过的")) as Record<string, unknown>;
    for (const key of ["keepNext", "keepLines", "widowControl", "snapToGrid", "wordWrap", "borders", "shading", "tabStops"])
      expect(eight[key] === undefined ? `少了 ${key}` : key).toBe(key);
    const own = paraStyleOf(findPara(doc, "八项都说过的")) as Record<string, unknown>;
    for (const key of ["keepNext", "keepLines", "widowControl", "snapToGrid", "wordWrap", "borders", "shading", "tabStops"])
      expect(own[key], `那半本不该有 ${key}`).toBeUndefined();
    // 账里摆的是这一段说过的话：没说过的不进账，否则签记录会把模型上的空位当成「改之前的话」说一遍。
    const silent = carriedShape(findPara(doc, "什么都没说的")) as Record<string, unknown>;
    for (const key of ["keepNext", "keepLines", "widowControl", "snapToGrid", "wordWrap", "borders", "shading", "tabStops"])
      expect(silent[key], `空位 ${key} 不该进账`).toBeUndefined();
  });

  it("回车拆段：上一段自己那句「允许」跟着走，从样式借来的那句留在样式那里", () => {
    const own = open(makeChainDocx([{ id: "Head", next: "Body" }, { id: "Body" }], [{ pStyle: "Head", pPr: "<w:wordWrap/>", text: "自己说的" }]));
    const split = own.nextSplit(findPara(own, "自己说的")!)!;
    expect(split.styleId).toBe("Body");
    expect(split.wordWrap).toBe(true);

    const lent = open(makeChainDocx([{ id: "Head", next: "Body", pPr: "<w:wordWrap/>" }, { id: "Body" }], [{ pStyle: "Head", text: "借来的" }]));
    const borrowed = lent.nextSplit(findPara(lent, "借来的")!)!;
    // 「允许」是 Head 说的话，Body 没说，所以新段不该认领它。
    expect(findPara(lent, "借来的").wordWrap).toBe(true);
    expect(borrowed.wordWrap).toBeUndefined();
  });
});
