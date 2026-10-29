import { describe, expect, it } from "vitest";
import type { Inline, Paragraph, PropSlot } from "../src/core/docx.js";
import { WordDocument, applyFacePick, applyRunPatch, hasLine } from "../src/core/docx.js";
import { SEEN_BY_VIEW, inlineLength, keepRunIdentity } from "../src/core/run-identity.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeChainDocx } from "./helpers/make-paragraph.js";
import { makeHanFaceDocx, makeHighlightDocx, makeOnOffDocx, makeTabRunDocx, makeUnderlineDocx } from "./helpers/make-run.js";

/**
 * #78 — the page is a rendering of the resolved run style, and the editor re-derives the model from it on every
 * keystroke. A rebuilt run therefore has to be told which run it came from: the file's own `w:rPr` children, the
 * read snapshot that says which properties on screen the run stated, the two faces of `w:rFonts` and a drawing's
 * bytes are not things an element can show.
 *
 * The DOM itself is what `.scratch/run-font-dom.ts` and `.scratch/run-kern-dom.ts` check. What these cases pin is
 * the rule underneath, on the lists that path produces: the view's own words (`SEEN_BY_VIEW`) plus the text.
 */

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
const rPrs = (xml: string): string[] =>
  [...xml.matchAll(/<w:rPr(?:\s[^>]*)?>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>|<w:rPr(?:\s[^>]*)?\/>/g)].map((m) => m[0]);
const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const paraAt = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text).join("").includes(needle))!;

/** What the page can tell us back: the fields `styleOf` reads off tags and inline CSS, the ones it reads off the
 *  stamped attributes, and the text as it now stands. Nothing else survives a rebuild — that is the gap. */
function viewOf(run: Inline, text = run.text): Inline {
  const out: Partial<Inline> = { text };
  for (const key of SEEN_BY_VIEW) {
    if (key === "text") continue;
    const value = run[key];
    if (value !== undefined) (out as Record<string, unknown>)[key] = value;
  }
  return out as Inline;
}

/** The view's rebuild of a whole paragraph, with `add` typed into the run holding `needle`. */
const retyped = (inlines: Inline[], needle: string, add: string): Inline[] => {
  const prev = inlines.map((r) => ({ ...r }));
  const next = prev.map((r) => (r.text.includes(needle) ? viewOf(r, r.text + add) : viewOf(r)));
  keepRunIdentity(next, prev);
  return next;
};

/**
 * The list the page hands back as it really is. `styleOf` only ever switches `italic` and `strike` to true — an
 * `<i>`, a `text-decoration` — and `vertAlign` to `superscript` or `subscript` through a `sup`/`sub` tag or
 * `vertical-align:super/sub`, so a property the file turned off has nothing on screen to show it and comes back as
 * nothing at all. `bold` is the one that can report its off side, through `font-weight`. `viewOf` above lets a run
 * keep a `false` it was handed, which is what the page cannot do.
 *
 * `highlight` and `underline` are deliberately not in that delete list: since #88 and #210 the renderer stamps the
 * name the document chose (`data-hl`, `data-u`) and `styleOf` reads the attribute, so a run that said
 * `w:val="none"` comes back saying `"none"`, and a `thick` line keeps being thick rather than a colour or a line
 * guessed from CSS — which is what those rounds measured as lost (`.scratch/hl-pending.ts`: `background:none`
 * hands Chromium's longhand back as `initial`, and only seven names survive a trip through CSS as something a
 * reverse lookup can name). `unstampedView` below is the case the two stamps do not cover.
 */
function pageView(run: Inline, text = run.text): Inline {
  const out = viewOf(run, text);
  for (const key of ["italic", "strike"] as const) if (run[key] === false) delete out[key];
  if (run.vertAlign === "baseline") delete out.vertAlign;
  return out;
}

/** The page when the markup carries none of our stamps: markup Blink made for a pending format, or pasted in.
 *  高亮与线名都只能按 CSS 答——有线就是 `single`，说了不要就一句都不说。 */
const unstampedView = (run: Inline, text = run.text): Inline => {
  const out = pageView(run, text);
  delete out.highlight;
  if (hasLine(out.underline)) out.underline = "single";
  else delete out.underline;
  return out;
};

/** The same rebuild, from the page as it actually answers. */
const retypedByPage = (inlines: Inline[], needle: string, add: string): Inline[] => {
  const prev = inlines.map((r) => ({ ...r }));
  const next = prev.map((r) => (r.text.includes(needle) ? pageView(r, r.text + add) : pageView(r)));
  keepRunIdentity(next, prev);
  return next;
};

/** Every field a run can hold, so a new one cannot join the list of things the page silently forgets. */
const FULL: Inline = {
  text: "正文",
  bold: true,
  italic: true,
  underline: "single",
  strike: true,
  sizePt: 12,
  color: "#C00000",
  highlight: "yellow",
  vertAlign: "subscript",
  font: "Calibri",
  fontEastAsia: "宋体",
  hyperlink: "#anchor",
  linkAttrs: 'w:history="1"',
  image: { src: "data:image/png;base64,AA", width: 10, height: 20, name: "p.png", path: "media/p.png" },
  drawing: { xml: "<w:drawing><wp:extent/></w:drawing>", width: 10, height: 20, name: "p.png" },
  break: true,
  brXml: '<w:br w:type="clear"/>',
  tab: true,
  rPrSlots: [{ tag: "kern", xml: '<w:kern w:val="2"/>' } as PropSlot],
  runText: "正文（同一枚 run 里还有 tab）",
  asRead: { bold: true },
};

describe("a run rebuilt from the page keeps what the page cannot show", () => {
  it("hands back every field the view has no word for", () => {
    const prev: Inline[] = [{ ...FULL }];
    const next: Inline[] = [viewOf(FULL, FULL.text + "打的字")];
    keepRunIdentity(next, prev);
    for (const key of Object.keys(FULL) as (keyof Inline)[]) {
      if (key === "text") continue; // the text is the one thing the page does tell us, and it grew
      expect(next[0]![key], String(key)).toEqual(FULL[key]);
    }
  });

  it("keeps a cancellation the page has nothing to show for", () => {
    const prev: Inline[] = [
      {
        text: "这句话",
        underline: "none",
        italic: false,
        asRead: { underline: "none", italic: false },
        rPrSlots: [
          { tag: "i", xml: '<w:i w:val="0"/>' } as PropSlot,
          { tag: "u", xml: '<w:u w:val="none"/>' } as PropSlot,
        ],
      },
    ];
    const next = [pageView(prev[0]!, "这句话打的字")];
    keepRunIdentity(next, prev);
    // A run that says nothing about the line is not the run taking its `none` back — only an off switch the page
    // can show is a contradiction. Left standing, the silence makes the writer delete what the file stated.
    // 线名这一格现在由 `data-u` 带回来（#210），页面有得说；把这枚钉拿走（Blink 自造的、或粘进来的标记）才是这里
    // 那条「页面什么都不说」的路，同一句 `none` 得由 `keepRunIdentity` 从上一枚 run 认领回来。
    expect(next[0]!.underline).toBe("none");
    expect(next[0]!.italic).toBe(false);

    const unstamped = [unstampedView(prev[0]!, "这句话打的字")];
    keepRunIdentity(unstamped, prev);
    expect(unstamped[0]!.underline).toBe("none");

    // The one property the page does switch off on its own stays the page's answer.
    const contradicted = [pageView({ ...prev[0]!, bold: false }, "这句话打的字")];
    keepRunIdentity(contradicted, [{ ...prev[0]!, bold: true }]);
    expect(contradicted[0]!.bold).toBe(false);
  });

  it("saves a typed sentence without erasing the no-underline it was typed into", () => {
    const doc = open(makeUnderlineDocx());
    const before = rPrs(save(doc));
    const p = paraAt(doc, "说明没有");
    p.inlines = retypedByPage(p.inlines, "说明没有", "打字");
    expect(p.inlines[0]!.underline).toBe("none");
    expect(rPrs(save(doc))).toEqual(before);
  });

  it("keeps a struck-off and a baseline run through the page's silence", () => {
    const prev: Inline[] = [
      { text: "没划", strike: false, asRead: { strike: false }, rPrSlots: [{ tag: "strike", xml: '<w:strike w:val="0"/>' } as PropSlot] },
      { text: "基线", vertAlign: "baseline", asRead: { vertAlign: "baseline" }, rPrSlots: [{ tag: "vertAlign", xml: '<w:vertAlign w:val="baseline"/>' } as PropSlot] },
    ];
    // The screen shows both of them plain, and `styleOf` has no answer for a run that sits on the line.
    const next = [pageView(prev[0]!, "没划打的字"), pageView(prev[1]!, "基线")];
    keepRunIdentity(next, prev);
    expect(next[0]!.strike).toBe(false);
    expect(next[1]!.vertAlign).toBe("baseline");

    // A sup the page does show is still the page's answer, the way `bold` was above.
    const contradicted = [pageView({ ...prev[1]!, vertAlign: "superscript" }, "基线")];
    keepRunIdentity(contradicted, [{ ...prev[1]! }]);
    expect(contradicted[0]!.vertAlign).toBe("superscript");
  });

  it("saves a typed sentence without erasing the no-strike or the baseline it was typed into", () => {
    const doc = open(makeOnOffDocx());
    const before = rPrs(save(doc));
    for (const needle of ["说了没划", "回到基线"]) {
      const p = paraAt(doc, needle);
      p.inlines = retypedByPage(p.inlines, needle, "打字");
      expect(p.inlines[0]!.text).toContain("打字");
    }
    expect(rPrs(save(doc))).toEqual(before);
  });

  it("saves a typed sentence without erasing the no-highlight it was typed into", () => {
    const doc = open(makeHighlightDocx());
    const before = rPrs(save(doc));
    // Three sentences the file wrote about colour: its own `none`, a §17.17.64 name the page could never guess
    // from CSS, and a `none` standing against a colour the style would lend. All three come back off `data-hl`.
    for (const needle of ["说了没高亮", "页面叫不出名字的高亮", "盖掉借来的高亮"]) {
      const p = paraAt(doc, needle);
      p.inlines = retypedByPage(p.inlines, needle, "打字");
      expect(p.inlines[0]!.highlight, needle).toBe(needle === "页面叫不出名字的高亮" ? "blue" : "none");
      expect(p.inlines[0]!.text).toContain("打字");
    }
    expect(rPrs(save(doc))).toEqual(before);
  });

  it("keeps a no-highlight through markup that carries no stamp", () => {
    const prev: Inline[] = [
      {
        text: "没高亮",
        highlight: "none",
        asRead: { highlight: "none" },
        rPrSlots: [{ tag: "highlight", xml: '<w:highlight w:val="none"/>' } as PropSlot],
      },
    ];
    // A span Blink made for a pending format, or pasted markup: nothing of ours on it, so it says nothing about
    // the colour the run gave. Page silence is not a contradiction of a cancellation.
    const next = [unstampedView(prev[0]!, "没高亮打的字")];
    keepRunIdentity(next, prev);
    expect(next[0]!.highlight).toBe("none");
    expect(next[0]!.asRead).toEqual({ highlight: "none" });

    // A colour the page does show is still the page's answer, the way `bold` was above.
    const contradicted = [pageView({ ...prev[0]!, highlight: "yellow" }, "没高亮")];
    keepRunIdentity(contradicted, [{ ...prev[0]! }]);
    expect(contradicted[0]!.highlight).toBe("yellow");
  });

  it("keeps the file's own \`w:rPr\` and the read snapshot through a keystroke", () => {
    const doc = open(
      makeChainDocx(
        [{ id: "Root", rPr: "<w:b/>" }],
        [{ pStyle: "Root", rPr: '<w:kern w:val="2"/><w:sz w:val="20"/>', text: "自己写了字距的" }, { text: "别的那一段" }],
      ),
    );
    const p = paraAt(doc, "自己写了字距的");
    const before = save(doc);
    p.inlines = retyped(p.inlines, "自己写了字距的", "打字");
    const after = save(doc);
    // `w:kern` came in with the run and the page never saw it; `w:b` is the style's, and the run only showed it.
    expect(rPrs(after)).toEqual(rPrs(before));
    expect(after).toContain('<w:kern w:val="2"/>');
    expect(paraXmlOf(after, "自己写了字距的")).not.toContain("<w:b/>");
    // The paragraph nobody typed in kept its nothing: the floor's size stays the floor's.
    expect(paraXmlOf(after, "别的那一段")).not.toContain("<w:sz");
  });

  it("lets a borrowed property become the run's own once the page shows it differently", () => {
    const doc = open(makeChainDocx([{ id: "Root", rPr: "<w:b/>" }], [{ pStyle: "Root", text: "跟着样式粗着" }]));
    const p = paraAt(doc, "跟着样式粗着");
    const prev = p.inlines.map((r) => ({ ...r }));
    // The toolbar un-bolding the run is the one thing the page does say: it comes back without the mark.
    const next = [viewOf({ ...prev[0]!, bold: false }, prev[0]!.text)];
    keepRunIdentity(next, prev);
    expect(next[0]!.bold).toBe(false);
    // The snapshot is what the run *showed*, resolved: the style's bold and the file's floor together, not a field list.
    expect(next[0]!.asRead).toEqual({ bold: true, sizePt: 11, font: "Calibri", fontEastAsia: "宋体" });
    p.inlines = next;
    const rPr = rPrs(save(doc))[0] ?? "";
    expect(rPr).toContain("<w:b"); // the run really did say so this time
    expect(rPr).not.toContain("<w:sz"); // and what it still only borrowed stays off the run
  });

  it("gives a picture's bytes to the one run it came from", () => {
    const prev: Inline[] = [{ text: "", image: FULL.image, drawing: FULL.drawing, sizePt: 10 }];
    // Two image elements aligned onto the one picture — the shape the list takes after a deletion next to it.
    const next: Inline[] = [viewOf(prev[0]!), viewOf(prev[0]!)];
    keepRunIdentity(next, prev);
    expect(next[0]!.drawing).toEqual(FULL.drawing);
    expect(next[1]!.drawing).toBeUndefined();
  });

  it("lines the untouched text up on both sides of the change", () => {
    const prev: Inline[] = [
      { text: "头", font: "A", fontEastAsia: "甲", rPrSlots: [{ tag: "kern", xml: '<w:kern w:val="2"/>' } as PropSlot] },
      { text: "中", font: "B", fontEastAsia: "乙" },
      { text: "尾", font: "C", fontEastAsia: "丙" },
    ];
    // The middle run's text is gone; the run after it did not move and must not take the deleted run's face.
    const next = [viewOf(prev[0]!), viewOf(prev[2]!, "尾")];
    keepRunIdentity(next, prev);
    expect(next[1]!.font).toBe("C");
    expect(next[1]!.fontEastAsia).toBe("丙");
    expect(next[0]!.rPrSlots).toHaveLength(1);
  });

  it("leaves a carried structure and an anchor to themselves", () => {
    const carry: Inline = { text: "", carry: { key: 3, tag: "bookmarkStart", xml: "<w:bookmarkStart w:id=\"3\"/>", show: "" } };
    const anchor: Inline = { text: "", anchor: { kind: "start", id: 7 } };
    const prev: Inline[] = [{ text: "文字", font: "A", rPrSlots: [{ tag: "kern", xml: "<w:kern/>" } as PropSlot] }, carry, anchor];
    const next: Inline[] = [viewOf(prev[0]!, "文字"), carry, anchor];
    keepRunIdentity(next, prev);
    expect(next[1]).toBe(carry);
    expect(next[2]).toBe(anchor);
    expect(next[1]!.rPrSlots).toBeUndefined();
    expect(next[2]!.font).toBeUndefined();
  });

  it("writes the same bytes for a document whose runs were only looked at", () => {
    const bytes = makeChainDocx(
      [{ id: "Root", rPr: '<w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/><w:sz w:val="24"/>' }],
      [
        { pStyle: "Root", text: "中文标题跟着样式" },
        { pStyle: "Root", rPr: '<w:lang w:val="zh-CN"/><w:rFonts w:eastAsia="楷体"/>', text: "自己写了语言与中文脸" },
      ],
    );
    const doc = open(bytes);
    const before = rPrs(save(doc));
    for (const p of paragraphs(doc)) p.inlines = retyped(p.inlines, p.inlines.map((r) => r.text).join(""), "打的字");
    expect(rPrs(save(doc))).toEqual(before);
  });

  it("页面重拼之后，同一枚 run 的两半改一遍脸仍并回一枚", () => {
    // #143：「这枚 run 说的是中文」不在页面上，靠 `keepRunIdentity` 从原点带回来。少了它，插在汉字段里的断行
    // 改字体时不补 `w:eastAsia`，两半各写一份性质，`w:cs`/`w:hint` 跟着多写一遍（§86 那份 9/9/9 就是这么来的）。
    const doc = open(makeTabRunDocx());
    const p = paragraphs(doc).find((x) => x.inlines.some((r) => r.break))!;
    const prev = p.inlines;
    const next = prev.map((r) => viewOf(r));
    keepRunIdentity(next, prev);
    for (const r of next) Object.assign(r, applyRunPatch(r, { font: "黑体" }));
    p.inlines = next;
    const para = paraXmlOf(save(doc), "分页符后的字");
    expect(para).toMatch(/<w:br w:type="page"\/><w:t>分页符后的字<\/w:t>/);
    expect(para).toContain('w:eastAsia="黑体"');
    expect((para.match(/<w:r\b/g) ?? []).length).toBe(1);
    expect((para.match(/w:cs=/g) ?? []).length).toBe(1);
  });

  it("原点是单孩子的 run，插在中间的分页符跟着整枚 run 回答「这说的是中文」", () => {
    // 上面那条走的是文件里本来就并列两个孩子的 run，`runText` 能从原点带回来。这一条是**我们自己造出的形状**：
    // 读进来时只有一枚 `w:t`（没有 `runText`），InsertPageBreak 之后页面上是三截。断行自己没有字，它问不到
    // 「这枚 run 说的是哪种文字」，只能问同一枚 run 的那一套 `w:rPr` 引用的兄弟们——引用就是写端并回一枚的依据。
    const doc = open(makeHanFaceDocx());
    const p = paraAt(doc, "这一句只写了西文脸");
    const prev = p.inlines.map((r) => ({ ...r }));
    // 页面上那枚断行是一个 `br[data-brk]` 元素，`inlinesOf` 把它推成 `{...ctx, break:true, brXml}`：段里解析出来
    // 的磅数它照旧带着，缺的是 `w:rPr` 与「这枚 run 说了哪种文字」。
    const next: Inline[] = [
      viewOf(prev[0]!, "这一"),
      { ...viewOf(prev[0]!, ""), break: true, brXml: '<w:br w:type="page"/>' },
      viewOf(prev[0]!, "句只写了西文脸"),
    ];
    keepRunIdentity(next, prev);
    expect(next[1]!.runText).toBe("这一句只写了西文脸");
    for (const r of next) Object.assign(r, applyRunPatch(r, { font: "黑体" }));
    p.inlines = next;
    const para = paraXmlOf(save(doc), "这一");
    expect(para).toMatch(/<w:br w:type="page"\/><w:t[^>]*>句只写了西文脸<\/w:t>/);
    expect((para.match(/<w:r\b/g) ?? []).length).toBe(1);
    expect((para.match(/w:cs=/g) ?? []).length).toBe(1);
    expect((para.match(/<w:sz\b/g) ?? []).length).toBe(1);
    expect(para).toContain('w:eastAsia="黑体"');
  });

  it("没带回槽引用的孩子不替别的 run 说话", () => {
    // 槽引用为空时写端本来也不会并（`joinableRun`），所以这里没有「同一枚 run」可问——汉字的判断照旧落在孩子
    // 自己那句话上。这一条钉住上面那条规矩的边界：重算只按引用并组，不按相邻。
    const prev: Inline[] = [{ text: "汉字" }, { text: "", break: true, brXml: "<w:br/>" }];
    const next = prev.map((r) => viewOf(r));
    keepRunIdentity(next, prev);
    expect(next[1]!.runText).toBeUndefined();
    expect(applyFacePick(next[1]!, "黑体")).toEqual({ font: "黑体" });
  });
});

function paraXmlOf(xml: string, needle: string): string {
  return [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)]
    .map((m) => m[0])
    .find((s) => s.includes(needle)) ?? "";
}

describe("the run list's own counting", () => {
  it("counts a mark as one character, which is how a caret offset is counted", () => {
    expect(inlineLength({ text: "" })).toBe(1);
    expect(inlineLength({ text: "三个字" })).toBe(3);
    expect(inlineLength({ text: "", break: true })).toBe(1);
  });
});
