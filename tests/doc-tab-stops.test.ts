import { describe, expect, it } from "vitest";
import type { Paragraph, TabStop } from "../src/core/docx.js";
import { WordDocument, carriedShape, syncParaSlots } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { text, OpcPackage } from "../src/core/opc.js";
import { makeChainDocx, makeIndentDocx } from "./helpers/make-paragraph.js";
import { makeTabLeaderDocx } from "./helpers/make-tab-leader.js";
import { renderWord } from "../src/core/render.js";
import { parseTabStops } from "../src/ui/doc-layout.js";

/**
 * `w:pPr/w:tabs` (§17.4.19). The shapes these tests transcribe are the ones the authorized 46 份 write
 * (`.scratch/tabs-audience.py`, `.scratch/tabs-inside-table.py`): 370 blocks, every one of them under a `w:pPr`
 * (230 in `word/document.xml`, 90 in `word/styles.xml`, 12 in the bands), 106 of them naming two stops or more;
 * 520 stops, **all with `w:pos`**, spelled in four ways only — `left` 304, `right` 152, `center` 58, `clear` 6 —
 * with `w:leader` on 94 of them and every leader saying `dot`; `decimal` and `bar` 0, so the model has no field
 * for them. 0 of the 370 sit in a table cell. `w:defaultTabStop` is written by all 46 (43 × 420, 3 × 720).
 *
 * 前导符那五段的形状（`tests/helpers/make-tab-leader.ts`）是**这一本账与装机壳共用的一件**：`.scratch` 把它写成
 * `/tmp/制表位前导符-装机.docx`，`src/main/main.ts` 的 `leaderLeg` 在真 ribbon 与真纸面上读同一件，纸上的读数
 * 与盘上写回的字节一起对。所以这里期望的枚数——` w:leader="` 在正文 3 枚、`word/styles.xml` 1 枚——是**量出来的**，
 * 不是照着断言凑的。
 */

const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const stopsOf = (bytes: Uint8Array, at = 0) => paragraphs(open(bytes))[at]!.tabStops;

describe("docx 制表位读取", () => {
  it("段落自己写的每一枚都读出来，按缇排序，w:leader 留作文件自己的token", () => {
    const docx = makeChainDocx([], [
      { pPr: '<w:tabs><w:tab w:val="right" w:pos="8306"/><w:tab w:val="left" w:pos="720" w:leader="dot"/></w:tabs>', text: "两枚" },
    ]);
    expect(stopsOf(docx)).toEqual([
      { posTwips: 720, val: "left", leader: "dot" },
      { posTwips: 8306, val: "right", leader: undefined },
    ]);
  });

  it("空的 w:tabs 是「什么都没说」，不是「一处不落」", () => {
    expect(stopsOf(makeChainDocx([], [{ pPr: "<w:tabs/>", text: "空块" }]))).toBeUndefined();
  });

  it("w:val 漏写按 ST_TabVal 的默认值 left；w:pos 漏写的那一枚不收", () => {
    const docx = makeChainDocx([], [{ pPr: '<w:tabs><w:tab w:pos="4153"/><w:tab w:val="right"/></w:tabs>', text: "两种漏写" }]);
    expect(stopsOf(docx)).toEqual([{ posTwips: 4153, val: "left", leader: undefined }]);
  });

  it("样式链与段落逐**位置**合：段落那一枚只顶掉同位置的，样式其余的原样留着", () => {
    // nearest-wins over the whole element would lose 4153 here — and 106 of the 370 blocks write two or more at once.
    const docx = makeChainDocx(
      [{ id: "T", pPr: '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="center" w:pos="4153"/></w:tabs>' }],
      [{ pStyle: "T", pPr: '<w:tabs><w:tab w:val="right" w:pos="4153"/></w:tabs>', text: "顶掉一枚" }],
    );
    expect(stopsOf(docx)).toEqual([
      { posTwips: 720, val: "left", leader: undefined },
      { posTwips: 4153, val: "right", leader: undefined },
    ]);
  });

  it("两级样式各写各的：往上的那一跳不把下面那跳的清单整批抹掉", () => {
    const docx = makeChainDocx(
      [
        { id: "Base", pPr: '<w:tabs><w:tab w:val="left" w:pos="2160"/></w:tabs>' },
        { id: "Child", basedOn: "Base", pPr: '<w:tabs><w:tab w:val="right" w:pos="9000"/></w:tabs>' },
      ],
      [{ pStyle: "Child", text: "链上两级" }],
    );
    expect(stopsOf(docx)).toEqual([
      { posTwips: 2160, val: "left", leader: undefined },
      { posTwips: 9000, val: "right", leader: undefined },
    ]);
  });

  it("段落写 w:val=\"clear\" 就只办事不占位：抹掉链上同位置那一枚，别的位置不动", () => {
    // 授权 46 份 6 枚 `clear`，枚枚带 `w:pos`（`.scratch/tabs-audience.py`）。
    const docx = makeChainDocx(
      [{ id: "T", pPr: '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="8306"/></w:tabs>' }],
      [{ pStyle: "T", pPr: '<w:tabs><w:tab w:val="clear" w:pos="8306"/></w:tabs>', text: "取消那枚" }],
    );
    expect(stopsOf(docx)).toEqual([{ posTwips: 720, val: "left", leader: undefined }]);
  });

  it("拆段的新那一段带走上一段算好的制表位（`MODEL_ONLY_PARA_KEYS` 那本账）", () => {
    const p = paragraphs(open(makeChainDocx([], [{ pPr: '<w:tabs><w:tab w:val="right" w:pos="8306"/></w:tabs>', text: "有位的段" }])))[0]!;
    expect(carriedShape(p).tabStops).toEqual([{ posTwips: 8306, val: "right", leader: undefined }]);
  });
});

describe("docx 制表位的写端与画端", () => {
  /** One keystroke, then save: the path a document is actually put through. */
  const edited = (bytes: Uint8Array): string => {
    const doc = open(bytes);
    const p = paragraphs(doc)[0]!;
    p.inlines[0] = { ...p.inlines[0]!, text: "打过字的" };
    return part(writeDocx(doc));
  };

  it("改一次正文，w:tabs 按原件那一块带回来，不多写一枚", () => {
    const bytes = makeChainDocx([], [{ pPr: '<w:tabs><w:tab w:val="right" w:pos="8306"/></w:tabs>', text: "整段自述" }]);
    const saved = edited(bytes);
    expect(saved.match(/<w:tabs>/g)?.length).toBe(1);
    expect(saved).toContain('<w:tab w:val="right" w:pos="8306"/>');
    expect(saved.match(/<w:tab /g)?.length).toBe(1);
  });

  it("`data-tab-stops` 把算好的位置交给量的人；没有位的段落不带这一句", () => {
    const html = renderWord(open(makeChainDocx([], [
      { pPr: '<w:tabs><w:tab w:val="right" w:pos="8306"/><w:tab w:val="center" w:pos="4153"/></w:tabs>', text: "有位" },
      { text: "没有位", rPr: "" },
    ])));
    expect(html).toContain('data-tab-stops="4153:c,8306:r"');
    expect(html.match(/data-tab-stops/g)?.length).toBe(1);
    // 画端只认这一句字面：读回来就是同一张位置表。
    expect(parseTabStops("4153:c,8306:r")).toEqual([
      { posTwips: 4153, val: "center" },
      { posTwips: 8306, val: "right" },
    ]);
    expect(parseTabStops("8306,,:x")).toEqual([]);
  });

  it("w:leader 的原词跟着位置一起进 data-tab-stops；不在 ST_LeaderType 里的那一枚不写", () => {
    // 授权 46 份带 `w:leader` 的 94 枚枚枚说 `dot`（`.scratch/tabs-leader-roundtrip.ts`）；`blink` 是本机造的
    // 坏形状，断的是「页面上不替它画墨」，不是「件里说了不算」——模型照原词带着它，写回去也还是那一句。
    const bytes = makeChainDocx([], [
      { pPr: '<w:tabs><w:tab w:val="right" w:pos="8306" w:leader="dot"/><w:tab w:val="left" w:pos="720" w:leader="blink"/></w:tabs>', text: "两档" },
    ]);
    expect(paragraphs(open(bytes))[0]!.tabStops).toEqual([
      { posTwips: 720, val: "left", leader: "blink" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ]);
    expect(renderWord(open(bytes))).toContain('data-tab-stops="720:l,8306:r:dot"');
    expect(parseTabStops("720:l,8306:r:dot")).toEqual([
      { posTwips: 720, val: "left" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ]);
  });

  it("改一次正文，w:leader 随原件那一块带回来", () => {
    // 写端不重拼 `w:tabs`：这一枚走的是段落自己那批槽位（`MODEL_ONLY_PARA_KEYS` 那本账），
    // 所以认不认得这个档都不影响它回到盘上——`tabs-leader-roundtrip.ts` 量的 92→92 同一条路。
    const saved = edited(
      makeChainDocx([], [{ pPr: '<w:tabs><w:tab w:val="right" w:pos="8306" w:leader="dot"/></w:tabs>', text: "整段自述" }]),
    );
    expect(saved).toContain('<w:tab w:val="right" w:pos="8306" w:leader="dot"/>');
    expect(saved.match(/ w:leader="/g)?.length).toBe(1);
  });
});

/**
 * 装机壳那一腿（`src/main/main.ts` 的 `leaderLeg`）与 `.scratch/tab-leader-install-check.sh` 共用这一件，
 * 所以这一族问的是「件本身在源层面就分得开」：模型读到的位置表、画端写出的那一句字面、重拼带回的原词。
 */
describe("docx 前导符夹具的五段形状", () => {
  const doc = () => WordDocument.openBytes(makeTabLeaderDocx());
  const byNeedle = (needle: string) =>
    paragraphs(doc()).find((p) => p.inlines.map((i) => ("text" in i ? i.text : "")).join("").includes(needle))!;
  const xml = (bytes: Uint8Array, partName = "word/document.xml") => text(OpcPackage.from(bytes).bytes(partName)!);

  it("甲一枚 dot、乙两枚只有一枚说 dot、丙从样式借、丁说 spaces、戊与尾段什么都没有", () => {
    expect(byNeedle("甲线").tabStops).toEqual([{ posTwips: 8306, val: "right", leader: "dot" }]);
    expect(byNeedle("乙线").tabStops).toEqual([
      { posTwips: 2880, val: "left", leader: "dot" },
      { posTwips: 8306, val: "right", leader: undefined },
    ]);
    expect(byNeedle("丙线").tabStops).toEqual([{ posTwips: 1440, val: "left", leader: "dot" }]);
    expect(byNeedle("丁线").tabStops).toEqual([{ posTwips: 4153, val: "left", leader: "spaces" }]);
    expect(byNeedle("戊线").tabStops).toBeUndefined();
    expect(byNeedle("尾段").tabStops).toBeUndefined();
  });

  it("画端四段各带一句，戊与尾段不带；六枚制表符一枚不多", () => {
    const html = renderWord(doc());
    expect([...html.matchAll(/data-tab-stops="([^"]*)"/g)].map((m) => m[1])).toEqual([
      "8306:r:dot",
      "2880:l:dot,8306:r",
      "1440:l:dot",
      "4153:l:spaces",
    ]);
    expect(html.match(/class="tab"/g)?.length).toBe(6);
  });

  it("整件重拼：三枚 w:leader 与样式那一枚原样回去，没写位的段不落 w:tabs", () => {
    const saved = xml(writeDocx(doc()));
    expect(saved.match(/ w:leader="/g)?.length).toBe(3);
    expect(xml(writeDocx(doc()), "word/styles.xml").match(/ w:leader="/g)?.length).toBe(1);
    expect(saved).toContain('<w:tab w:val="left" w:pos="4153" w:leader="spaces"/>');
    // 戊线那一段自己会说些什么，但绝不会被安上一枚位——按段取，别按字符窗口取（隔壁丁线正写着位）。
    const paraOf = (xml: string, needle: string) =>
      [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>).)*<\/w:p>/gs)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
    expect(paraOf(saved, "戊线")).not.toContain("<w:tabs>");
    expect(paraOf(saved, "尾段")).not.toContain("<w:tabs>");
  });
});

/**
 * 写回那一族（#234）。断的是「对话框按了确定之后，盘上落的正是用户动过的那几格」，参照的写法全部来自授权
 * 46 份自己：`.scratch/tabs-slot-shape.py` 量的 320 枚 `w:tabs`、462 枚孩子——属性序只有 `w:val w:pos`（368 枚）
 * 与 `w:val w:leader w:pos`（94 枚）两种，枚枚自闭合（成对写法 0 枚），318 枚块的孩子按 `w:pos` 升序（另 2 枚
 * 既非升序也非降序），`w:val="clear"` 2 枚、重复位置 0 处。认领的粒度因此是**位置**而不是整块：整枚抄上去会把
 * 样式链的话算到段落头上，而 `p.tabStops` 正是合并过链的那一份。
 */
describe("docx 制表位写回", () => {
  const paraOf = (xml: string, needle: string): string =>
    [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>).)*<\/w:p>/gs)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
  const tabsOf = (out: Uint8Array, needle: string): string => /<w:tabs>[\s\S]*?<\/w:tabs>/.exec(paraOf(part(out), needle))?.[0] ?? "";
  const stylesOf = (out: Uint8Array): string => text(OpcPackage.from(out).bytes("word/styles.xml")!);

  /** 「确定」那一下：字段换成新的整份清单，再让段落把自己**这一击**动过的位置认领下来。 */
  const setStops = (bytes: Uint8Array, needle: string, next: TabStop[]): Uint8Array => {
    const doc = open(bytes);
    const p = paragraphs(doc).find((x) => x.inlines.map((i) => ("text" in i ? i.text : "")).join("").includes(needle))!;
    p.tabStops = next;
    syncParaSlots(p);
    return writeDocx(doc);
  };

  /** 同一趟里按两次「确定」：`syncParaSlots` 每次都把段自己那几枚孩子的字节换下来，第二击比的该是**刚落下的
   * 那一串**，不是入件那一份——装机腿步骤 5（无→点线→无）走的正是这一条，只在盘上看得到。 */
  const setStopsTwice = (bytes: Uint8Array, needle: string, a: TabStop[], b: TabStop[] | undefined): Uint8Array => {
    const doc = open(bytes);
    const p = paragraphs(doc).find((x) => x.inlines.map((i) => ("text" in i ? i.text : "")).join("").includes(needle))!;
    p.tabStops = a;
    syncParaSlots(p);
    p.tabStops = b;
    syncParaSlots(p);
    return writeDocx(doc);
  };

  // 两枚孩子用的是量出来的那两种写法之一（`w:val w:pos`）与另一种（`w:val w:leader w:pos`）：462 枚里没有第三种序。
  const TWO = '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="8306" w:leader="dot"/></w:tabs>';
  const twoStopDocx = (label: string) => makeChainDocx([], [{ pPr: TWO, text: label }]);

  it("原样确定不漂移：内容没变时一枚都不重写，件那一块按原字节回去", () => {
    const out = setStops(twoStopDocx("两格没改"), "两格没改", [
      { posTwips: 720, val: "left" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ]);
    expect(tabsOf(out, "两格没改")).toBe(TWO);
    expect(part(out)).not.toContain('w:val="clear"');
  });

  it("新增的一枚按 w:pos 升序插进件那一块，没被点名的孩子一个字节不动", () => {
    const out = setStops(twoStopDocx("添一格"), "添一格", [
      { posTwips: 720, val: "left" },
      { posTwips: 4153, val: "center" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ]);
    // 新增那一枚身上没有 `w:leader` 可写，就一个字都不写（`spaces` 那一档也是这么退掉的）。
    expect(tabsOf(out, "添一格")).toBe(
      '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306" w:leader="dot"/></w:tabs>',
    );
  });

  it("只改一格时对没动的那一格不说；改动那一格按 CT_TabStop 的属性序落笔", () => {
    // 新写的 `w:leader` 插在 `w:val` 与 `w:pos` 中间，不是贴在末尾——94 枚带前导符的孩子枚枚是这个序。
    const out = setStops(twoStopDocx("改左那一格"), "改左那一格", [
      { posTwips: 720, val: "center", leader: "dash" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ]);
    expect(tabsOf(out, "改左那一格")).toBe(
      '<w:tabs><w:tab w:val="center" w:leader="dash" w:pos="720"/><w:tab w:val="right" w:pos="8306" w:leader="dot"/></w:tabs>',
    );
  });

  it("前导符退回「无」是把那句 w:leader 从盘上退掉，不是留着个空值", () => {
    const out = setStops(twoStopDocx("撤了点线"), "撤了点线", [
      { posTwips: 720, val: "left" },
      { posTwips: 8306, val: "right" },
    ]);
    expect(tabsOf(out, "撤了点线")).toBe('<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="8306"/></w:tabs>');
  });

  it("删掉样式借来的那一格写成 w:val=\"clear\"，段落只带那一枚，样式自己那两枚不动", () => {
    // §17.3.2.155 让一段作废样式给的那一格用的正是这一枚孩子；授权口径的块里本来就有（clear 2 枚 / 462 枚）。
    const bytes = makeChainDocx(
      [{ id: "T", pPr: '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="8306"/></w:tabs>' }],
      [{ pStyle: "T", text: "借两格" }],
    );
    const out = setStops(bytes, "借两格", [{ posTwips: 720, val: "left" }]);
    expect(tabsOf(out, "借两格")).toBe('<w:tabs><w:tab w:val="clear" w:pos="8306"/></w:tabs>');
    // 没被动过的那一格不抄：段落身上只有一句「8306 那格作废」，720 仍然只由样式说。
    expect(stylesOf(out)).toContain('<w:tab w:val="left" w:pos="720"/>');
    expect(paraOf(part(out), "借两格")).not.toContain('w:pos="720"');
  });

  it("样式借来的格原样确定时，段落身上不落 w:tabs", () => {
    // 这一枚是「不无中生有」的判据：`p.tabStops` 里那两格是链给的，把它换成内容相同的一份新数组不能算动过。
    const bytes = makeChainDocx(
      [{ id: "T", pPr: '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="8306"/></w:tabs>' }],
      [{ pStyle: "T", text: "借两格没改" }],
    );
    const out = setStops(bytes, "借两格没改", [
      { posTwips: 720, val: "left" },
      { posTwips: 8306, val: "right" },
    ]);
    expect(paraOf(part(out), "借两格没改")).not.toContain("<w:tabs>");
  });

  it("同一趟里加了前导符又撤掉：字节回到件原本那一串，纸上没墨盘上也不说有", () => {
    // 装机腿步骤 5 的那一击（无→点线→无）。第二击时字段与快照相等，判据若按「与入件那一份比」就成了净变化为零，
    // 第一击落下的 `w:leader="dot"` 会一直停在盘上——纸早就不画了。
    const out = setStopsTwice(twoStopDocx("加了点线又撤"), "加了点线又撤", [
      { posTwips: 720, val: "left", leader: "dot" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ], [
      { posTwips: 720, val: "left" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ]);
    expect(tabsOf(out, "加了点线又撤")).toBe(TWO);
  });

  it("同一趟里添的一格又被删掉：段落收回那句，不作废谁的话", () => {
    // 快照里没有 4153 这一格（件与链都没说过），所以它不该换来一枚 `w:val="clear"`——那是「一段作废样式给的
    // 那一格」的写法。前一击由段落自己写下、这一击由段落自己摘掉，块回到件原本那两枚孩子。
    const out = setStopsTwice(twoStopDocx("添一格又删"), "添一格又删", [
      { posTwips: 720, val: "left" },
      { posTwips: 4153, val: "center" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ], [
      { posTwips: 720, val: "left" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ]);
    expect(tabsOf(out, "添一格又删")).toBe(TWO);
    expect(paraOf(part(out), "添一格又删")).not.toContain('w:val="clear"');
  });

  it("件里本来那一格被删仍写 clear，同一趟里另一格的说法改动一起落", () => {
    // `drop` 与 `clear` 的分界是「入件时这一格有没有话」，不是「这一趟里有没有被碰过」：8306 是件自己写的，
    // 删掉它就得明说作废；720 那一枚是这一趟改的说法，由字段重拼。
    const out = setStopsTwice(twoStopDocx("改一枚删一枚"), "改一枚删一枚", [
      { posTwips: 720, val: "left", leader: "dot" },
      { posTwips: 4153, val: "center" },
      { posTwips: 8306, val: "right", leader: "dot" },
    ], [
      { posTwips: 720, val: "left", leader: "dot" },
    ]);
    expect(tabsOf(out, "改一枚删一枚")).toBe('<w:tabs><w:tab w:val="left" w:leader="dot" w:pos="720"/><w:tab w:val="clear" w:pos="8306"/></w:tabs>');
  });

  it("段本来一句没说：添上去又全部清除之后，块整个从段身上摘掉", () => {
    // 「全部清除」那一击落回 `undefined`。既不能留下 `clear`，也不能留下空的 `<w:tabs></w:tabs>`——
    // 段落对这一族又回到一句没说，与入件时同一个形状。
    const out = setStopsTwice(makeChainDocx([], [{ text: "添了再全清" }]), "添了再全清", [{ posTwips: 720, val: "left" }], undefined);
    expect(paraOf(part(out), "添了再全清")).not.toContain("<w:tabs>");
    expect(paraOf(part(out), "添了再全清")).not.toContain('w:val="clear"');
  });

  it("模型不收的 w:val=\"decimal\" 原样留在块里：位置对得上就不替件改说法", () => {
    // 读端把不在 ST_TabVal 四档里的那一枚按默认值 `left` 画（授权 46 份 decimal/bar 各 0 枚，模型没有它们的字段）；
    // 写端不能因此把件那句改掉，也不能把它丢掉。
    const out = setStops(makeChainDocx([], [{ pPr: '<w:tabs><w:tab w:val="decimal" w:pos="4000"/></w:tabs>', text: "借位添一格" }]), "借位添一格", [
      { posTwips: 720, val: "left" },
      { posTwips: 4000, val: "left" },
    ]);
    expect(tabsOf(out, "借位添一格")).toBe('<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="decimal" w:pos="4000"/></w:tabs>');
  });

  it("没有 w:pPr 的段（编辑器拆出来的那一段）从字段那条路写出同一句", () => {
    // 两条路共用 `paraTabStopEdits`，所以「槽位带着原件字节」与「只剩字段」不会说成两件事。
    const doc = open(makeChainDocx([], [{ text: "没有 pPr 的段" }]));
    const p = paragraphs(doc)[0]!;
    p.tabStops = [{ posTwips: 8306, val: "right", leader: "dot" }];
    syncParaSlots(p);
    // 段上没有槽位可参照，这一枚就是新写的：属性按 CT_TabStop 的序 val→leader→pos。
    expect(tabsOf(writeDocx(doc), "没有 pPr 的段")).toBe('<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs>');
  });
});

/**
 * 「默认制表位」那一枚（§17.7.2 的 `w:defaultTabStop`）不在 `w:pPr` 里，说的是整份件，落在
 * `word/settings.xml`。参照的写法同样来自授权 46 份：`.scratch/settings-tabstop-shape.py` 量到 46 份**一份不缺**
 * 地写了它、枚枚自闭合且只带 `w:val` 这一枚属性（420 × 43、720 × 3）；件里缺它的是 `文件 → 新建` 那形状，
 * 插入位置按 `.scratch/settings-child-order.py` 在同批件里量出来的「之后兄弟」列取最靠前的那一枚。
 */
describe("docx 默认制表位写回", () => {
  const settingsOf = (out: Uint8Array): string => text(OpcPackage.from(out).bytes("word/settings.xml") ?? new Uint8Array());

  it("件里已经有那一枚时只换 w:val 那一个数，其余的字节按原件带走", () => {
    const bytes = makeIndentDocx([{ text: "一段" }], "", '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:hyphenationZone w:val="420"/>');
    const doc = open(bytes);
    expect(doc.defaultTabStop()).toBe(720);
    expect(doc.edited).toBe(false);
    doc.setDefaultTabStop(420);
    // 落在部件里、不在 blocks 里的一枚也得报「改了」——漏进指纹就是对话框关了、盘上还是旧数（#180 同一族）。
    expect(doc.edited).toBe(true);
    const out = writeDocx(doc);
    expect(settingsOf(out)).toContain('<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/><w:hyphenationZone w:val="420"/>');
    expect(settingsOf(out).match(/defaultTabStop/g)?.length).toBe(1);
    expect(open(out).defaultTabStop()).toBe(420);
  });

  it("没人说过它：改一次正文，那件部件一个字节都不动", () => {
    const bytes = makeIndentDocx([{ text: "一段" }], "", '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>');
    const carried = text(OpcPackage.from(bytes).bytes("word/settings.xml")!);
    const doc = open(bytes);
    const p = paragraphs(doc)[0]!;
    p.inlines[0] = { ...p.inlines[0]!, text: "打过字的" };
    expect(doc.defaultTabStopEdit()).toBeUndefined();
    expect(settingsOf(writeDocx(doc))).toBe(carried);
  });

  it("件里缺那一枚时插在量出来的那些之后兄弟之前，件自己的顺序不乱", () => {
    const doc = open(makeIndentDocx([{ text: "一段" }], "", '<w:zoom w:percent="100"/><w:hyphenationZone w:val="420"/>'));
    expect(doc.defaultTabStop()).toBeUndefined();
    doc.setDefaultTabStop(420);
    expect(settingsOf(writeDocx(doc))).toContain('<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/><w:hyphenationZone w:val="420"/>');
  });

  it("件根本没有 word/settings.xml 时把部件、关系与内容类型三样一起补齐", () => {
    // 授权 46 份份份带这一件，但 `文件 → 新建` 的件可能没有；一处没有关系的部件是 Word 会忽略的部件，
    // 用户刚填的那个距离就不会在下次打开时还在。
    const doc = open(makeIndentDocx([{ text: "一段" }]));
    doc.setDefaultTabStop(420);
    const out = writeDocx(doc);
    expect(settingsOf(out)).toContain('<w:defaultTabStop w:val="420"/>');
    expect(text(OpcPackage.from(out).bytes("word/_rels/document.xml.rels")!)).toContain('relationships/settings" Target="settings.xml"');
    expect(text(OpcPackage.from(out).bytes("[Content_Types].xml")!)).toContain(
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
    );
    expect(open(out).defaultTabStop()).toBe(420);
  });

  it("件里已经挂好指向 settings 的关系时，换默认制表位不再发第二张身份证", () => {
    // 授权 46 份的 `word/settings.xml` 份份挂在**正文那一页**的关系表上（`.scratch/settings-rel-census.py`：
    // 件级 `.rels` 挂它的 0 份）。写出器按同一张表问「这件部件有关系了吗」，答「有」就不再补——
    // 补了就是一个部件两张身份证，读的一侧按哪一张走没有定数。
    const bytes = makeIndentDocx([{ text: "一段" }], "", '<w:defaultTabStop w:val="720"/>');
    expect(text(OpcPackage.from(bytes).bytes("word/_rels/document.xml.rels")!)).toContain('relationships/settings" Target="settings.xml"');
    const doc = open(bytes);
    doc.setDefaultTabStop(420);
    const out = OpcPackage.from(writeDocx(doc));
    expect(text(out.bytes("word/_rels/document.xml.rels")!).match(/relationships\/settings/g)?.length).toBe(1);
    expect(text(out.bytes("_rels/.rels")!)).not.toContain("settings");
    expect(open(writeDocx(doc)).defaultTabStop()).toBe(420);
  });

  it("它跟着节的撤销载体走：一次恢复回到件里那个数", () => {
    const doc = open(makeIndentDocx([{ text: "一段" }], "", '<w:defaultTabStop w:val="720"/>'));
    const step = doc.pageGeometryState();
    doc.setDefaultTabStop(420);
    expect(doc.defaultTabStop()).toBe(420);
    doc.restorePageGeometry(step);
    expect(doc.defaultTabStopEdit()).toBeUndefined();
    expect(doc.defaultTabStop()).toBe(720);
  });
});
