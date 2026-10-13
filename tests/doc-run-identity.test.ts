import { describe, expect, it } from "vitest";
import type { Inline, Paragraph, PropSlot } from "../src/core/docx.js";
import { WordDocument } from "../src/core/docx.js";
import { SEEN_BY_VIEW, inlineLength, keepRunIdentity } from "../src/core/run-identity.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeChainDocx } from "./helpers/make-paragraph.js";
import { makeUnderlineDocx } from "./helpers/make-run.js";

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
 * The list the page hands back as it really is. `styleOf` only ever switches `italic`, `underline` and `strike`
 * to true — an `<i>`, a `<u>`, a `text-decoration` — so a property the file turned off has nothing on screen to
 * show it and comes back as nothing at all. `bold` is the one that can report its off side, through
 * `font-weight`. `viewOf` above lets a run keep a `false` it was handed, which is what the page cannot do.
 */
function pageView(run: Inline, text = run.text): Inline {
  const out = viewOf(run, text);
  for (const key of ["italic", "underline", "strike"] as const) if (run[key] === false) delete out[key];
  return out;
}

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
  underline: true,
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
        underline: false,
        italic: false,
        asRead: { underline: false, italic: false },
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
    expect(next[0]!.underline).toBe(false);
    expect(next[0]!.italic).toBe(false);

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
    expect(p.inlines[0]!.underline).toBe(false);
    expect(rPrs(save(doc))).toEqual(before);
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
