import { describe, expect, it } from "vitest";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { WordDocument, ownRunStyle, runStyleOf } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeChainDocx } from "./helpers/make-paragraph.js";

/**
 * §17.7.2's ladder seen from a run: its own `w:rPr`, then the character style its `w:rStyle` names, then the
 * paragraph style's chain and the default paragraph style beneath it, then `w:docDefaults/w:rPrDefault`.
 *
 * Measured over the 86 WPS-written documents (`.scratch/run-lent-impact.ts`, `.scratch/run-lent-check.ts`):
 * of the 61,702 runs the model reads in body paragraphs and first-level table cells, 60,578 already state one of
 * these properties in their own `w:rPr`, 768 write a `w:rPr` that states none of them, and 356 write no `w:rPr` at
 * all — WPS spells properties on nearly every run, but usually only some of them. 5,507 runs in 54 of the 86 files
 * now draw a property their own `w:rPr` never states (4,664 a size, 672 a colour, 660 a bold, 8 an underline),
 * which is both shapes: a field a partly-written `w:rPr` leaves open, and the 1,124 runs whose file stated nothing.
 * Among the 785 paragraphs the outline draws as a heading, 736 have at least one run doing so (654 bold, 348
 * colour, 215 size) — those are the headings that were showing this application's stylesheet rather than the
 * file's, which is what the previous round's `h1..h4` still owed.
 *
 * `w:rFonts` is held out of the lending on purpose (see `LENT_RUN_KEYS`): 350 of the 1,320 style-side faces name
 * a different `w:eastAsia` from their `w:ascii`, and the model's one `font` field is the ascii one, so handing it
 * over would put Chinese text in the Latin face. That holds back a face from 2,752 runs whose chain names one.
 *
 * The write side owes the files the mirror image: a save must not hand a `w:rPr` to a run whose file wrote none,
 * nor re-spell a borrowed value as one the run stated — `ownRunStyle` against the read snapshot, and
 * `.scratch/run-lent-bytes.ts` over all 86 confirms the 147,459 `<w:rPr>` in `word/document.xml` come back the
 * same after an edited save.
 */

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));

const paragraphs = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const byText = (doc: WordDocument, needle: string): Paragraph =>
  paragraphs(doc).find((p) => p.inlines.map((r) => r.text ?? "").join("").includes(needle))!;
const runOf = (doc: WordDocument, needle: string): Inline => byText(doc, needle).inlines.find((i) => i.text.includes(needle))!;
const paraXml = (xml: string, needle: string): string =>
  [...xml.matchAll(/<w:p(?:\s[^>]*)?>(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
const runXml = (xml: string, needle: string): string =>
  [...paraXml(xml, needle).matchAll(/<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*<\/w:r>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
/** The run's own `w:rPr`, as written — empty string when the save gave it none. */
const rPrXml = (xml: string, needle: string): string =>
  /<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>|<w:rPr\/>/.exec(runXml(xml, needle))?.[0] ?? "";
const drawnRun = (doc: WordDocument, needle: string): string => {
  for (const m of renderWord(doc).matchAll(/<span[^>]*\sstyle="([^"]*)"[^>]*>(?:(?!<\/span>)[\s\S])*?<\/span>/g)) {
    if (m[0].includes(needle)) return m[1] ?? "";
  }
  return "";
};

/** A heading whose bold and size live two styles up, exactly as the corpus's heading styles are built. */
const HEADING = [
  { id: "Root", rPr: '<w:b/><w:sz w:val="32"/>' },
  { id: "Title", basedOn: "Root", rPr: '<w:color w:val="C00000"/>' },
];

describe("docx run 格式从样式借：w:style 的 w:rPr 进阶梯", () => {
  it("reads the properties a paragraph's style states above the run", () => {
    const doc = open(makeChainDocx(HEADING, [{ pStyle: "Title", text: "两层给的粗体" }]));
    const r = runOf(doc, "两层给的粗体");
    expect(r.bold).toBe(true);
    expect(r.sizePt).toBe(16);
    expect(r.color).toMatch(/^#[0-9A-F]{6}$/);
    // The page draws them: this is what the h1–h4 headings were missing.
    expect(drawnRun(doc, "两层给的粗体")).toBe(`font-weight:600;font-size:16pt;color:${r.color};`);
  });

  it("takes each property from the first level that states it", () => {
    // The run's own size, the colour from its own style, the bold from that style's ancestor — three levels,
    // one field each, and none of them blanks the others.
    const doc = open(
      makeChainDocx(
        [
          { id: "Root", rPr: '<w:b/><w:sz w:val="32"/><w:color w:val="0000FF"/>' },
          { id: "Title", basedOn: "Root", rPr: '<w:sz w:val="28"/>' },
        ],
        [{ pStyle: "Title", rPr: '<w:sz w:val="20"/>', text: "各层各给一样" }],
      ),
    );
    const r = runOf(doc, "各层各给一样");
    expect(r.sizePt).toBe(10);
    expect(r.color).toBe("#0000FF");
    expect(r.bold).toBe(true);
  });

  it("lets the run cancel a bold the style hands down", () => {
    // `<w:b w:val="0"/>` is a stated value, not an absent one: §17.7.2's per-property rule stops at the run.
    const doc = open(makeChainDocx(HEADING, [{ pStyle: "Title", rPr: '<w:b w:val="0"/>', text: "标题里的不粗" }]));
    const r = runOf(doc, "标题里的不粗");
    expect(r.bold).toBe(false);
    expect(r.sizePt).toBe(16);
    expect(drawnRun(doc, "标题里的不粗")).toBe("font-size:16pt;color:#C00000;");
  });

  it("cancels at the style itself before the ancestor's bold comes through", () => {
    const doc = open(
      makeChainDocx(
        [
          { id: "Root", rPr: '<w:b/><w:sz w:val="32"/>' },
          { id: "Quiet", basedOn: "Root", rPr: '<w:b w:val="0"/>' },
        ],
        [{ pStyle: "Quiet", text: "样式自己关掉的粗体" }],
      ),
    );
    const r = runOf(doc, "样式自己关掉的粗体");
    expect(r.bold).toBe(false);
    expect(r.sizePt).toBe(16);
  });

  it("ranks the character style above the paragraph's and walks its own chain", () => {
    // `w:rStyle` is the run's second rung: its size wins over the heading's, while the properties the character
    // style leaves open still come from the paragraph below it — and the character chain's ancestor is hop two.
    const doc = open(
      makeChainDocx(
        [
          ...HEADING,
          { id: "Strong", type: "character", rPr: '<w:sz w:val="24"/>' },
          { id: "Loud", type: "character", basedOn: "Strong", rPr: "<w:i/>" },
        ],
        [
          { pStyle: "Title", rStyle: "Strong", text: "字符样式压过段落样式" },
          { pStyle: "Title", rStyle: "Loud", text: "字符样式的上一跳" },
        ],
      ),
    );
    const a = runOf(doc, "字符样式压过段落样式");
    expect(a.sizePt).toBe(12);
    expect(a.bold).toBe(true);
    const b = runOf(doc, "字符样式的上一跳");
    expect(b.sizePt).toBe(12);
    expect(b.italic).toBe(true);
    expect(b.color).toBe(a.color);
  });

  it("reaches a paragraph that names no style through the default style", () => {
    const doc = open(
      makeChainDocx(
        [{ id: "Normal", isDefault: true, rPr: "<w:i/>" }],
        [{ text: "没写样式的正文" }],
      ),
    );
    const r = runOf(doc, "没写样式的正文");
    expect(r.italic).toBe(true);
    // Nothing states a size below the paragraph, so §17.7.2's floor — `w:docDefaults/w:rPrDefault` — gives it.
    expect(r.sizePt).toBe(11);
  });

  it("keeps the floor under the paragraph chain when a character style hops through it", () => {
    // The same document seen from a run that names a character style: `Strong` states no size, so the size is
    // still the paragraph chain's 16pt. Folding the floor into each chain instead of putting it under the whole
    // ladder made the character hop carry the floor's 11pt up over it — the bottom rung winning a fight two
    // rungs above it. `.scratch/make-lent-demo.ts` draws this exact shape; the corpus never does, because the
    // 86 files' `w:rPrDefault` states only `w:rFonts`.
    const doc = open(
      makeChainDocx(
        [...HEADING, { id: "Strong", type: "character", rPr: "<w:i/>" }],
        [{ pStyle: "Title", rStyle: "Strong", text: "借来的字号没被地板顶掉" }],
      ),
    );
    const r = runOf(doc, "借来的字号没被地板顶掉");
    expect(r.sizePt).toBe(16);
    expect(r.italic).toBe(true);
    expect(r.bold).toBe(true);
    expect(drawnRun(doc, "借来的字号没被地板顶掉")).toBe(`font-weight:600;font-style:italic;font-size:16pt;color:${r.color};`);
  });

  it("does not lend a paragraph's properties from a character style marked default", () => {
    const doc = open(
      makeChainDocx([{ id: "DefaultChar", type: "character", isDefault: true, rPr: "<w:b/>" }], [{ text: "默认的是字符样式" }]),
    );
    expect(runOf(doc, "默认的是字符样式").bold).toBeUndefined();
  });

  it("holds the face back, because the model reads one of the two the style names", () => {
    // 350 of the 1,320 style-side `w:rFonts` give `w:eastAsia` a different face from `w:ascii`; `font` is the
    // ascii one, so lending it would set the Chinese in a Latin face the file chose for the Latin only.
    const doc = open(
      makeChainDocx(
        [{ id: "Root", rPr: '<w:rFonts w:ascii="Arial" w:eastAsia="SimSun"/><w:b/>' }],
        [{ pStyle: "Root", text: "中文不接西文字体" }, { pStyle: "Root", rPr: '<w:rFonts w:ascii="宋体"/>', text: "自己写了字体的" }],
      ),
    );
    expect(runOf(doc, "中文不接西文字体").font).toBeUndefined();
    expect(runOf(doc, "中文不接西文字体").bold).toBe(true);
    expect(runOf(doc, "自己写了字体的").font).toBe("宋体");
  });

  it("keeps an unedited run's `w:rPr` out of the saved file", () => {
    // The 1,124 runs a file wrote no property for (356 with no `w:rPr` at all, 768 whose `w:rPr` states none of
    // these): the page shows the heading's bold, and the file keeps not having a `w:rPr` on them. One edit
    // elsewhere is enough to make the save regenerate the part.
    const doc = open(makeChainDocx(HEADING, [{ pStyle: "Title", text: "两层给的粗体" }, { text: "被改了一下的正文" }]));
    byText(doc, "被改了一下的正文").align = "right";
    expect(drawnRun(doc, "两层给的粗体")).toContain("font-weight:600");
    expect(rPrXml(save(doc), "两层给的粗体")).toBe("");
    expect(paraXml(save(doc), "被改了一下的正文")).toContain('<w:jc w:val="right"/>');
  });

  it("writes only the property moved onto such a run", () => {
    const doc = open(makeChainDocx(HEADING, [{ pStyle: "Title", text: "被取消粗体的标题" }, { text: "别动" }]));
    const r = runOf(doc, "被取消粗体的标题");
    expect(r.bold).toBe(true);
    r.bold = false;
    expect(rPrXml(save(doc), "被取消粗体的标题")).toBe('<w:rPr><w:b w:val="0"/></w:rPr>');
    // The size and the colour it borrowed stay borrowed: writing them would move the heading's formatting onto
    // the run, where the document never had it.
    expect(rPrXml(save(doc), "别动")).toBe("");
  });

  it("leaves borrowed properties out of the set the run wrote for itself", () => {
    // 60,578 of the corpus's 61,702 modelled runs write their own `w:rPr`, most of them properties this model has
    // no field for; a chain-lent bold must not slip in beside the `w:kern` that only travels as the file's bytes.
    const doc = open(makeChainDocx(HEADING, [{ pStyle: "Title", rPr: '<w:kern w:val="2"/><w:sz w:val="20"/>', text: "自己写了字距的" }]));
    byText(doc, "自己写了字距的").align = "center";
    expect(runOf(doc, "自己写了字距的").bold).toBe(true);
    expect(rPrXml(save(doc), "自己写了字距的")).toBe('<w:rPr><w:kern w:val="2"/><w:sz w:val="20"/></w:rPr>');
  });

  it("round-trips a document whose runs borrow, without changing any run", () => {
    const bytes = makeChainDocx(
      [...HEADING, { id: "Strong", type: "character", rPr: '<w:sz w:val="24"/>' }],
      [
        { pStyle: "Title", text: "借来的" },
        { pStyle: "Title", rStyle: "Strong", text: "字符样式借来的" },
        { pStyle: "Title", rPr: '<w:b w:val="0"/>', text: "自己取消的" },
      ],
    );
    const doc = open(bytes);
    doc.blocks.push({ kind: "paragraph", inlines: [{ text: "编辑留下的痕迹" }] });
    const out = save(doc);
    expect(rPrXml(out, "借来的")).toBe("");
    // The run did write a `w:rPr` for itself — the one that names its character style — and that is what comes
    // back, with nothing added to it.
    expect(rPrXml(out, "字符样式借来的")).toBe('<w:rPr><w:rStyle w:val="Strong"/></w:rPr>');
    expect(rPrXml(out, "自己取消的")).toBe('<w:rPr><w:b w:val="0"/></w:rPr>');
    // Reading the ladder must not on its own count as an edit: with nothing touched the part comes back whole.
    expect(part(writeDocx(open(bytes)))).toBe(text(OpcPackage.from(bytes).bytes("word/document.xml")!));
  });

  it("spells a run the editor makes from its own fields", () => {
    // `ownRunStyle` is the mirror of `ownParaStyle`: what the snapshot already said is not the run's to write, and
    // a run this editor invented has no snapshot and states everything it shows.
    expect(ownRunStyle({ bold: true, sizePt: 16 }, { bold: true, sizePt: 10 })).toEqual({ sizePt: 10 });
    expect(ownRunStyle({ bold: true }, { bold: true })).toEqual({});
    expect(ownRunStyle(undefined, { bold: false })).toEqual({ bold: false });
    expect(runStyleOf({ text: "x", bold: true, sizePt: 16, rPrSlots: [] })).toEqual({ bold: true, sizePt: 16 });
  });
});
