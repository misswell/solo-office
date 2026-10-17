import { describe, expect, it } from "vitest";
import type { Paragraph } from "../src/core/docx.js";
import { WordDocument, syncParaSlots } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeIndentDocx, makeIndentRevisionDocx } from "./helpers/make-paragraph.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));

const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const byText = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
/** The paragraph's own `w:ind`, as written. Every fixture here spells it self-closing, so the tag is the element. */
const indXml = (xml: string, needle: string): string => /<w:ind[^>]*\/>/.exec(paraXml(xml, needle))?.[0] ?? "";
/** The one `w:ind` a one-paragraph document can hold. */
const firstInd = (xml: string): string => /<w:ind[^>]*\/>/.exec(xml)?.[0] ?? "";
/**
 * The `style` the page puts on the paragraph carrying `needle`, or "" when it puts none. A paragraph's own
 * element is the only place its indent can show up, so this is what "the page draws it" means here.
 */
const drawnStyle = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<(?:div|h[1-6])\s[^>]*class="para[^"]*"[^>]*>(?:(?!<\/(?:div|h[1-6])>)[\s\S])*<\/(?:div|h[1-6])>/g)) {
    if (m[0].includes(needle)) return / style="([^"]*)"/.exec(m[0])?.[1] ?? "";
  }
  return "";
};

/**
 * `w:ind` as WPS writes it, with the numbers it writes. Every shape and value below was counted over the 86
 * documents this repository's docx shapes come from — 29,207 `w:ind` elements in their `word/document.xml`
 * (`.scratch/ind-census.ts`, `ind-census3.ts`, `ind-order.ts`, `ind-values.ts`):
 *
 * - `w:firstLine` is stated positive 6,332 times and `w:firstLineChars` 1,387 (over every part that holds one),
 *   and 1,147 elements state both, always at a ratio that works out to a whole font size per character
 *   (240 twips for 2 characters = 12pt).
 * - 83 elements state a positive `w:firstLine` next to `w:firstLineChars="0"` (`w:firstLine="642"` 70 of them)
 *   and 4 the other way round, which is the measurement that decides the precedence the drawing follows.
 * - 232 state only the character form, every one of them `w:firstLineChars="300"`.
 * - 185 state `w:hanging`, each of them also stating a `w:left` (183 positive, 2 negative) and none of them
 *   `w:firstLine` (`w:left="440" w:hanging="440"` 72 times, `w:left="425" w:hanging="425"` 52), so 悬挂缩进 is a
 *   left indent whose first line hangs back out of it.
 * - `w:right` is stated positive 29 times (`w:right="220"` 28 of them) and negative 104 times, every one of
 *   those `-107`, so the right indent is not a non-negative number here.
 * - "no first-line indent" is written out loud as `w:firstLine="0" w:firstLineChars="0"` 17,856 times.
 *
 * Before this, the model read `w:left` alone, so of what the element could state it honoured one attribute and
 * drew nothing for the other four.
 */
describe("docx 段落缩进（w:ind）：首行、悬挂、右缩进读得出也画得出", () => {
  it("reads each spelling the file can give the first line", () => {
    const doc = open(
      makeIndentDocx([
        { ind: 'w:firstLine="480"', text: "只有绝对值" },
        { ind: 'w:firstLine="480" w:firstLineChars="200"', text: "两笔都写" },
        { ind: 'w:firstLineChars="300"', text: "只有字符数" },
        { ind: 'w:firstLine="642" w:firstLineChars="0"', text: "字符数写零" },
        { ind: 'w:firstLine="0" w:firstLineChars="0"', text: "说不缩进" },
      ]),
    );
    expect(byText(doc, "只有绝对值")).toMatchObject({ indentFirstLine: 480, indentFirstLineChars: undefined });
    expect(byText(doc, "两笔都写")).toMatchObject({ indentFirstLine: 480, indentFirstLineChars: 200 });
    expect(byText(doc, "只有字符数")).toMatchObject({ indentFirstLine: undefined, indentFirstLineChars: 300 });
    // A stated zero is a statement, not an absence: it is what lets a paragraph say "no indent" to a style
    // that lends one, and it is the 83-element shape that keeps the twips value drawing.
    expect(byText(doc, "字符数写零")).toMatchObject({ indentFirstLine: 642, indentFirstLineChars: 0 });
    expect(byText(doc, "说不缩进")).toMatchObject({ indentFirstLine: 0, indentFirstLineChars: 0 });
  });

  it("reads the hanging and right indent a paragraph states about itself", () => {
    const doc = open(
      makeIndentDocx([
        { ind: 'w:left="440" w:hanging="440"', text: "悬挂缩进" },
        { ind: 'w:right="220"', text: "右缩进" },
        { ind: 'w:right="-107"', text: "右缩进为负" },
      ]),
    );
    expect(byText(doc, "悬挂缩进")).toMatchObject({ indentLeft: 440, indentHanging: 440, indentFirstLine: undefined });
    expect(byText(doc, "右缩进")).toMatchObject({ indentRight: 220 });
    // 107 of the corpus's elements state a negative `w:right`, so this is not a non-negative number here.
    expect(byText(doc, "右缩进为负")).toMatchObject({ indentRight: -107 });
  });

  it("draws the first line from the absolute value when the file states both", () => {
    const doc = open(makeIndentDocx([{ ind: 'w:firstLine="480" w:firstLineChars="200"', text: "两笔都写" }]));
    // 480 twips is 32 CSS px; the character form would have made it 2.00em, and the two agree because the
    // ratio in all 1,147 paired elements is a font size.
    expect(drawnStyle(doc, "两笔都写")).toContain("text-indent:32px");
  });

  it("keeps drawing the indent the 83 elements state when their character form says zero", () => {
    const doc = open(makeIndentDocx([{ ind: 'w:firstLine="642" w:firstLineChars="0"', text: "字符数写零" }]));
    // §17.3.1.14 has the character form override the absolute one. Following that here would erase 83 indents
    // WPS itself drew, to settle 4 elements, so the absolute value wins whenever it is stated non-zero.
    expect(drawnStyle(doc, "字符数写零")).toContain("text-indent:42.8px");
  });

  it("draws a characters-only first line in em and a stated zero as nothing", () => {
    const doc = open(
      makeIndentDocx([
        { ind: 'w:firstLineChars="300"', text: "只有字符数" },
        { ind: 'w:firstLine="0" w:firstLineChars="0"', text: "说不缩进" },
      ]),
    );
    // 3.00em is the only drawing available without a twips value: it follows the page's font size rather than
    // the paragraph mark's. 未取证 how wide WPS makes these 236.
    expect(drawnStyle(doc, "只有字符数")).toContain("text-indent:3.00em");
    expect(drawnStyle(doc, "说不缩进")).not.toContain("text-indent");
  });

  it("hangs the first line back out of the left indent it shares the element with", () => {
    const doc = open(makeIndentDocx([{ ind: 'w:left="440" w:hanging="440"', text: "悬挂缩进" }]));
    // margin-left 29.33px with text-indent -29.33px is what puts the first line at the margin and the rest of
    // the paragraph in from it — 悬挂缩进, in the shape all 185 of them have.
    expect(drawnStyle(doc, "悬挂缩进")).toContain("margin-left:29.333333333333332px;text-indent:-29.333333333333332px");
  });

  it("draws a right indent, including the 104 in a body that are negative", () => {
    const doc = open(
      makeIndentDocx([
        { ind: 'w:right="220"', text: "右缩进" },
        { ind: 'w:right="-107"', text: "右缩进为负" },
      ]),
    );
    expect(drawnStyle(doc, "右缩进")).toContain("margin-right:14.666666666666666px");
    expect(drawnStyle(doc, "右缩进为负")).toContain("margin-right:-7.133333333333334px");
  });

  it("lends the indent a paragraph does not state and honours the one that does", () => {
    const doc = open(
      makeIndentDocx(
        [
          { pStyle: "Lent", text: "整段靠样式" },
          { pStyle: "Lent", ind: 'w:firstLine="0" w:firstLineChars="0"', text: "自己说不缩进" },
          { pStyle: "Lent", ind: 'w:right="220"', text: "只写右缩进" },
        ],
        '<w:ind w:left="420" w:firstLine="480" w:firstLineChars="200"/>',
      ),
    );
    // Nothing stated, so all three come from the style — 480 twips of first line and the 420 of left.
    expect(byText(doc, "整段靠样式")).toMatchObject({ indentLeft: 420, indentFirstLine: 480, indentFirstLineChars: 200 });
    expect(drawnStyle(doc, "整段靠样式")).toContain("margin-left:28px;text-indent:32px");
    // A stated zero answers the style, and only for the attribute it states.
    expect(byText(doc, "自己说不缩进")).toMatchObject({ indentFirstLine: 0, indentLeft: 420 });
    expect(drawnStyle(doc, "自己说不缩进")).not.toContain("text-indent");
    // An attribute the paragraph leaves unstated still comes through, alongside the one it wrote.
    expect(byText(doc, "只写右缩进")).toMatchObject({ indentRight: 220, indentFirstLine: 480, indentLeft: 420 });
  });

  it("段落自己答过首行那一件事时，样式链上的 firstLine 不再盖住它", () => {
    const doc = open(
      makeIndentDocx(
        [
          { pStyle: "Lent", ind: 'w:left="425" w:leftChars="0" w:hanging="425" w:firstLineChars="0"', text: "文件那一枚" },
          { pStyle: "Lent", ind: 'w:left="480" w:hanging="480"', text: "只写悬挂" },
          { pStyle: "Lent", ind: 'w:firstLine="480"', text: "自己写缇" },
          { pStyle: "Lent", ind: 'w:leftChars="0"', text: "左边只说字符零" },
        ],
        '<w:ind w:left="420" w:firstLine="883"/>',
      ),
    );
    // 授权 46 份的 `word/document.xml` 里这样的 94 枚，其中 74 枚的样式链写了 `w:firstLine="883"`：段首先行
    // 由段落自己说明白「不缩进」，链上那一枚不该把它盖掉（§17.3.1.14 + §17.7.2，见 `indResolve`）。
    expect(byText(doc, "文件那一枚")).toMatchObject({ indentLeft: 425, indentFirstLine: undefined, indentFirstLineChars: 0, indentHanging: 425 });
    expect(drawnStyle(doc, "文件那一枚")).toContain("margin-left:28.333333333333332px;text-indent:-28.333333333333332px");
    // 62 枚只写自己的 `w:hanging`，其中 34 枚同样被链上的 883 盖过。
    expect(byText(doc, "只写悬挂")).toMatchObject({ indentFirstLine: undefined, indentHanging: 480 });
    expect(drawnStyle(doc, "只写悬挂")).toContain("text-indent:-32px");
    // 段落自己写了缇那一枚时照旧由它上身——这一支与上面那支的分别只在「段自己有没有答过」。
    expect(byText(doc, "自己写缇")).toMatchObject({ indentFirstLine: 480 });
    // 48 枚写 `w:leftChars="0"` 而不写 `w:left`：左边这一件事也由段落答了，链上 420 缇不许叠上来。
    expect(byText(doc, "左边只说字符零")).toMatchObject({ indentLeft: undefined, indentLeftChars: 0 });
    expect(drawnStyle(doc, "左边只说字符零")).not.toContain("margin-left");
    // 写回那一趟同样不许把链上那枚 883 塞进段落自己的 `w:ind`：认领的是段自己说过的属性。
    const own = byText(doc, "文件那一枚");
    syncParaSlots(own);
    const written = own.pPrSlots?.find((s) => s.tag === "ind")?.xml ?? "";
    expect(written).not.toContain('w:firstLine="883"');
    expect(written).toContain('w:hanging="425"');
  });

  it("saves what it read with every w:ind byte intact", () => {
    const bytes = makeIndentDocx([
      { ind: 'w:firstLine="480" w:firstLineChars="200"', text: "两笔都写" },
      { ind: 'w:left="440" w:hanging="440"', text: "悬挂缩进" },
      { ind: 'w:right="-107"', text: "右缩进为负" },
      { ind: 'w:firstLine="642" w:firstLineChars="0"', text: "字符数写零" },
    ]);
    const before = part(bytes);
    const after = save(open(bytes));
    // The file's own `w:ind` travels as text, so having four of its attributes in the model must cost a save
    // nothing: same attributes, same order, same `w:leftChars` no field has.
    for (const needle of ["两笔都写", "悬挂缩进", "右缩进为负", "字符数写零"]) {
      expect(indXml(after, needle)).toBe(indXml(before, needle));
    }
  });

  it("writes back every attribute a paragraph the editor made can state, in the order WPS uses", () => {
    const doc = WordDocument.blank();
    const p = paragraphs(doc)[0]!;
    expect(p.pPrSlots).toBeUndefined();
    p.indentLeft = 420;
    p.indentRight = 220;
    p.indentFirstLine = 480;
    p.indentFirstLineChars = 200;
    expect(firstInd(part(writeDocx(doc)))).toBe('<w:ind w:left="420" w:right="220" w:firstLine="480" w:firstLineChars="200"/>');

    const hung = WordDocument.blank();
    const h = paragraphs(hung)[0]!;
    h.indentLeft = 440;
    h.indentHanging = 440;
    // A paragraph spelled from fields used to write `w:left` alone, so the other three of these five died on
    // the way out of a split or a freshly typed paragraph no file had ever described.
    expect(firstInd(part(writeDocx(hung)))).toBe('<w:ind w:left="440" w:hanging="440"/>');
  });

  it("puts the recorded first line back when a format change is refused", () => {
    const doc = open(makeIndentRevisionDocx());
    expect(byText(doc, "缩进被改过")).toMatchObject({ indentFirstLine: 0, indentFirstLineChars: 0 });
    expect(doc.resolveRevisions("reject", 12)).toBe(1);
    const p = byText(doc, "缩进被改过");
    expect(p).toMatchObject({ indentFirstLine: 480, indentFirstLineChars: 200 });
    expect(drawnStyle(doc, "缩进被改过")).toContain("text-indent:32px");
    expect(indXml(save(doc), "缩进被改过")).toBe('<w:ind w:firstLine="480" w:firstLineChars="200"/>');
  });
});

/**
 * Where 增加缩进量／减少缩进量 puts the number, measured over the 39 distinct bodies WPS itself wrote inside
 * its own install bundle (`.scratch/indent-step.ts`, files deduped by content hash of `word/document.xml`):
 *
 * - `w:defaultTabStop` is stated 420 36 times and 720 3 times, and never left out — so the distance an indent
 *   spelled in twips moves is the file's own 默认制表位, and 420 缇 is 2 字符 at the 10.5pt these files write.
 * - a positive `w:leftChars` never arrives without its twips value, and the pairs that show up are 200↔420 (3),
 *   400↔840 (23), 303↔1207 (16), 100↔198 (14): dividing one by the other gives the paragraph's own em size
 *   (210/399/402 缇 per character), so a step moves both spellings and the twips value follows the ratio the
 *   file used rather than a width this model assumes.
 * - `w:left="0"` is stated 187 times, which is how these files say 回到边内. Whether the button keeps going past
 *   the left margin onto a negative value is 未取证 — the corpus holds `w:left="-1"`, `-41`, `-178`, but none of
 *   those files was made by clicking it, so this stops at 0.
 */
describe("docx 段落缩进：增加／减少缩进量走的那一步（两笔一起动）", () => {
  it("moves a character-counted indent by 2 字符 and keeps the file's own per-character width", () => {
    const doc = open(
      makeIndentDocx([
        { ind: 'w:left="840" w:leftChars="400"', text: "四字符" },
        { ind: 'w:left="1207" w:leftChars="303"', text: "十九点九磅的" },
      ]),
    );
    // 840/400 is 210 缇 per 1% character (10.5pt); 1260/600 states the same width, so the pair the file wrote
    // grows along its own line instead of being snapped to a number this model picked.
    expect(doc.nextIndent(byText(doc, "四字符"), 1)).toEqual({ indentLeft: 1260, indentLeftChars: 600 });
    expect(doc.nextIndent(byText(doc, "四字符"), -1)).toEqual({ indentLeft: 420, indentLeftChars: 200 });
    // 1207/303 is 399 缇 per character (19.9pt): rounding to 420 would move it by more than the step.
    expect(doc.nextIndent(byText(doc, "十九点九磅的"), 1)).toEqual({ indentLeft: 2004, indentLeftChars: 503 });
    expect(doc.nextIndent(byText(doc, "十九点九磅的"), -1)).toEqual({ indentLeft: 410, indentLeftChars: 103 });
  });

  it("steps a twips-only indent by the document's own 默认制表位 and invents no character count", () => {
    const doc = open(makeIndentDocx([{ ind: 'w:left="840"', text: "只写缇" }]));
    const p = byText(doc, "只写缇");
    expect(doc.defaultTabStop()).toBeUndefined();
    expect(doc.nextIndent(p, 1)).toEqual({ indentLeft: 1260 });
    expect(doc.nextIndent(p, 1)!.indentLeftChars).toBeUndefined();
    expect(doc.nextIndent(p, -1)).toEqual({ indentLeft: 420 });

    // The 3 files of the 39 that state 720 move by 720, not by the 420 the other 36 use.
    const wide = open(makeIndentDocx([{ ind: 'w:left="840"', text: "只写缇" }], "", '<w:defaultTabStop w:val="720"/>'));
    expect(wide.defaultTabStop()).toBe(720);
    expect(wide.nextIndent(byText(wide, "只写缇"), 1)).toEqual({ indentLeft: 1560 });
  });

  it("stops at the margin instead of writing a change that is nothing", () => {
    const doc = open(
      makeIndentDocx([
        { ind: 'w:left="0" w:leftChars="0"', text: "说了贴边" },
        { text: "没提过缩进" },
      ]),
    );
    // Both read as sitting on the edge, so 减少缩进量 has nowhere to go and the click must not put `w:ind` down.
    expect(doc.nextIndent(byText(doc, "说了贴边"), -1)).toBeUndefined();
    expect(doc.nextIndent(byText(doc, "没提过缩进"), -1)).toBeUndefined();
    // A stated zero is the 187-element shape: it moves out by 2 字符, and both spellings arrive together.
    expect(doc.nextIndent(byText(doc, "说了贴边"), 1)).toEqual({ indentLeft: 420, indentLeftChars: 200 });
    expect(doc.nextIndent(byText(doc, "没提过缩进"), 1)).toEqual({ indentLeft: 420 });
  });

  it("saves the moved pair over the file's own w:ind, leaving the attributes with no field in place", () => {
    const bytes = makeIndentDocx([
      { ind: 'w:left="840" w:leftChars="400" w:right="220" w:rightChars="100" w:firstLine="0" w:firstLineChars="0"', text: "点到的那段" },
      { ind: 'w:left="420" w:leftChars="200" w:hangingChars="0" w:hanging="0"', text: "别的那一段" },
    ]);
    const doc = open(bytes);
    const p = byText(doc, "点到的那段");
    const next = doc.nextIndent(p, 1)!;
    p.indentLeft = next.indentLeft;
    p.indentLeftChars = next.indentLeftChars;
    syncParaSlots(p);
    const after = save(doc);
    // `w:rightChars` has no field here, and `w:right`, both first-line spellings were not moved — they keep
    // their values and their places, so only the two numbers the step owns change.
    expect(indXml(after, "点到的那段")).toBe('<w:ind w:left="1260" w:leftChars="600" w:right="220" w:rightChars="100" w:firstLine="0" w:firstLineChars="0"/>');
    // The paragraph nobody clicked keeps its exact bytes even though its neighbour was rewritten.
    expect(indXml(after, "别的那一段")).toBe(indXml(part(bytes), "别的那一段"));
  });

  it("writes the indent a paragraph borrowed from its style down when the step moves it", () => {
    const doc = open(
      makeIndentDocx([{ pStyle: "Lent", text: "靠样式缩进" }], '<w:ind w:left="420" w:leftChars="200"/><w:jc w:val="center"/>'),
    );
    const p = byText(doc, "靠样式缩进");
    expect(p).toMatchObject({ indentLeft: 420, indentLeftChars: 200 });
    expect(p.pPrSlots?.map((s) => s.tag)).toEqual(["pStyle"]);
    const next = doc.nextIndent(p, -1)!;
    p.indentLeft = next.indentLeft;
    p.indentLeftChars = next.indentLeftChars;
    syncParaSlots(p);
    // 减少缩进量 on a borrowed indent answers the style out loud, in `CT_PPrBase`'s own slot for it: `w:pStyle`
    // keeps its place at the front and the new `w:ind` lands after it.
    expect(indXml(save(doc), "靠样式缩进")).toBe('<w:ind w:left="0" w:leftChars="0"/>');
    expect(paraXml(save(doc), "靠样式缩进").indexOf('<w:pStyle w:val="Lent"/>')).toBeLessThan(
      paraXml(save(doc), "靠样式缩进").indexOf("<w:ind"),
    );
  });

  it("draws the step on the page as well as writing it", () => {
    const doc = open(makeIndentDocx([{ ind: 'w:left="840" w:leftChars="400"', text: "点到的那段" }]));
    expect(drawnStyle(doc, "点到的那段")).toContain("margin-left:56px");
    const p = byText(doc, "点到的那段");
    const next = doc.nextIndent(p, 1)!;
    p.indentLeft = next.indentLeft;
    p.indentLeftChars = next.indentLeftChars;
    syncParaSlots(p);
    // 1260 缇 is 63pt is 84px — the character form would draw 6.00em, and the two agree because the pair the
    // step wrote is the pair the file wrote.
    expect(drawnStyle(doc, "点到的那段")).toContain("margin-left:84px");
  });
});
