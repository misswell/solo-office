import { describe, expect, it } from "vitest";
import type { Table } from "../src/core/docx.js";
import { WordDocument, rowShapeOf, signRowRevision, twipsToPx } from "../src/core/docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { applyRowHeights, distributeRows, type RowHeightMove } from "../src/core/doc-table-ops.js";
import { makeRowDistributeDocx } from "./helpers/make-tables.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const saved = (doc: WordDocument): string => text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
/** 四张：`unequal` 逐行说了又除不尽、`equal` 本来就等高、`unstated` 一句没说、`merged` 一带竖合并跨在两行上。 */
const quartet = (): { doc: WordDocument; unequal: Table; equal: Table; unstated: Table; merged: Table } => {
  const doc = open(makeRowDistributeDocx());
  const [unequal, equal, unstated, merged] = doc.blocks.filter((b) => b.kind === "table") as Table[];
  return { doc, unequal: unequal!, equal: equal!, unstated: unstated!, merged: merged! };
};
const heights = (t: Table): (number | undefined)[] => t.rows.map((row) => row[0]?.row?.heightTwips);
const every = (t: Table): number[] => t.rows.map((_, at) => at);
/** 编辑器那一击走的两步：core 先算（纸上量来的读数原样递下去），再按算出的答落笔。 */
const strike = (t: Table, rows: number[], onPaper: number[]): { ok: boolean; reason?: string; moved: RowHeightMove[] } => {
  const out = distributeRows(t, rows, onPaper);
  if (!out.ok) return { ok: false, reason: out.reason, moved: [] };
  return { ok: true, moved: applyRowHeights(t, out.rows) };
};

describe("docx「分布行」：等高而和不变（行自己那句 w:trHeight 换数、w:hRule 跟行原本的说法、表的宽与网格一个字不动）", () => {
  it("reads the four shapes the strike has to answer for", () => {
    const { unequal, equal, unstated, merged } = quartet();
    expect(heights(unequal)).toEqual([385, 402, 647, 647]);
    expect(heights(equal)).toEqual([600, 600, 600]);
    // 20/58 张表一行都没说（授权 46 份 ②：按行 580 说了 / 118 没说，同表混合 0 张）——件里说不出自己多高。
    expect(heights(unstated)).toEqual([undefined, undefined, undefined]);
    expect(heights(merged)).toEqual([200, 400, 600]);
    // 竖合并那一带仍是两行：`w:vMerge` 让每一行都留着自己的横边（226 枚、0 行不自成一行，⑥）。
    expect(merged.rows.map((row) => row.length)).toEqual([2, 2, 2]);
  });

  it("splits the remainder to the leading row and keeps the rows' sum exactly", () => {
    const { doc, unequal } = quartet();
    const out = strike(unequal, every(unequal), [385, 402, 647, 647]);
    expect(out.ok).toBe(true);
    expect(out.moved.length).toBe(4);
    // 2,081 ÷ 4 = 520 余 1：最上那一行多拿一缇，于是和还是 2,081——授权 46 份里 24/38 张走的就是这一条。
    expect(heights(unequal)).toEqual([521, 520, 520, 520]);
    const xml = saved(doc);
    expect(xml).toContain('<w:trPr><w:trHeight w:val="521" w:hRule="atLeast"/></w:trPr>');
    expect((xml.match(/<w:trHeight w:val="520" w:hRule="atLeast"\/>/g) ?? []).length).toBe(3);
    expect(xml).not.toContain('w:val="385"');
    // 行高是逐行说的，表的高本就是行叠出来的：这一击没有第二句字节要重述，表的宽与网格原样留着。
    expect(xml).toContain('<w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/>');
    expect(xml).toContain('<w:gridCol w:w="1500"/><w:gridCol w:w="1500"/>');
  });

  it("moves only the rows the selection touched, leaving the others' bytes alone", () => {
    const { doc, unequal } = quartet();
    const out = strike(unequal, [0, 1], [385, 402]);
    expect(out.ok && out.moved.length).toBe(2);
    // 787 ÷ 2 = 393 余 1 → 394 / 393；下面那两行没被碰到，那两句一个字节不动。
    expect(heights(unequal)).toEqual([394, 393, 647, 647]);
    const xml = saved(doc);
    expect(xml).toContain('<w:trHeight w:val="394" w:hRule="atLeast"/>');
    expect((xml.match(/<w:trHeight w:val="647" w:hRule="atLeast"\/>/g) ?? []).length).toBe(2);
  });

  it("takes the paper's reading when the file says nothing, and never invents w:hRule=\"exact\"", () => {
    const { doc, unstated } = quartet();
    // 602 ÷ 3 = 200 余 2：最上两行各多一缇。纸上的读数从铺排来（`doc-edit.ts` 逐行量 `getBoundingClientRect()`）。
    const out = strike(unstated, every(unstated), [100, 200, 302]);
    expect(out.ok).toBe(true);
    expect(heights(unstated)).toEqual([201, 201, 200]);
    const xml = saved(doc);
    expect((xml.match(/<w:trHeight w:val="201" w:hRule="atLeast"\/>/g) ?? []).length).toBe(2);
    expect(xml).toContain('<w:trHeight w:val="200" w:hRule="atLeast"/>');
    // 580 枚全说 atLeast、exact 0 枚（④）：编辑器插的行没有那句，补的是这一族唯一的写法，不是件里从没用过的那一种。
    expect(xml).not.toContain('w:hRule="exact"');
    // 行没有自己的对象：一行的每一格都得指到同一枚 `RowInfo`，否则写出与判定记录只认第一格。
    expect(unstated.rows[0]![1]!.row).toBe(unstated.rows[0]![0]!.row);
    expect(unstated.rows[0]![1]!.row?.heightTwips).toBe(201);
  });

  it("refuses rows that already stand equal, without spending a byte or an undo step", () => {
    const { doc, equal } = quartet();
    expect(strike(equal, every(equal), [600, 600, 600])).toMatchObject({ ok: false, reason: "这几行已经是等高" });
    expect(strike(equal, [9], [600])).toMatchObject({ ok: false, reason: "这一行不在表格里" });
    expect(heights(equal)).toEqual([600, 600, 600]);
    expect(saved(doc)).toBe(saved(open(makeRowDistributeDocx())));
  });

  it("judges「已经等高」by the paper, not by the numbers in the file", () => {
    const { doc, unequal } = quartet();
    // 纸上四行已经一样高（内容自己撑到 520），件里那四句 385 / 402 / 647 / 647 谁也没在纸上说话：
    // 这一击在纸上没有活。替没说话的那几行各补一句字节，是替件编一段它没有的墨。
    expect(strike(unequal, every(unequal), [520, 520, 520, 520])).toMatchObject({ ok: false, reason: "这几行已经是等高" });
    expect(heights(unequal)).toEqual([385, 402, 647, 647]);
    expect(saved(doc)).toBe(saved(open(makeRowDistributeDocx())));
  });

  it("refuses when the paper cannot measure, because one row has nothing to be divided against", () => {
    const { doc, unequal } = quartet();
    // 只剩一行量得到（其余 0 高：纸还没铺上、或那一行说了 `w:hidden`）——按 0 记进和是替它编一个假高度。
    expect(strike(unequal, [0, 1, 2, 3], [400, 0, 0, 0])).toMatchObject({ ok: false, reason: "量得出高度的只剩一行，等高要有两条边" });
    expect(strike(unequal, [0, 1], [0, 0])).toMatchObject({ ok: false, reason: "这几行在纸上量不出高度" });
    expect(heights(unequal)).toEqual([385, 402, 647, 647]);
    expect(saved(doc)).toBe(saved(open(makeRowDistributeDocx())));
  });

  it("puts the row the paper cannot measure out of the split and keeps the rest's sum", () => {
    const { doc, unequal } = quartet();
    const out = strike(unequal, [0, 1, 2, 3], [385, 402, 0, 647]);
    expect(out.ok && out.moved.length).toBe(3);
    // 中间那一行摘出去：1,434 ÷ 3 = 478，除得尽（38 张里 14 张走这一条）；它自己那句 647 原样留着。
    expect(heights(unequal)).toEqual([478, 478, 647, 478]);
    const xml = saved(doc);
    expect((xml.match(/<w:trHeight w:val="478" w:hRule="atLeast"\/>/g) ?? []).length).toBe(3);
    expect((xml.match(/<w:trHeight w:val="647" w:hRule="atLeast"\/>/g) ?? []).length).toBe(1);
  });

  it("distributes across a vertical merge without dropping a row, and never rewrites that band", () => {
    const { doc, merged } = quartet();
    const out = strike(merged, every(merged), [200, 400, 600]);
    // 中间那一行自己说的就是 400——均分算到它头上还是这个数，那一行不落笔（38 张里 14 张和除得尽，走的就是这一条）。
    expect(out.ok && out.moved.length).toBe(2);
    expect(heights(merged)).toEqual([400, 400, 400]);
    const xml = saved(doc);
    // 列那一击要摘掉「纸面上没有那条边」的列；行没有这一族——每一行都留着自己的横边，所以三行都在均分里。
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain('<w:vMerge/>');
    expect((xml.match(/<w:vMerge\b/g) ?? []).length).toBe(2);
    // 那三行如今各说 400：两行是这一笔换的，一行本来就说这个数。
    expect((xml.match(/<w:trHeight w:val="400" w:hRule="atLeast"\/>/g) ?? []).length).toBe(3);
  });

  it("draws the equalised rows on the page at the new heights", () => {
    const { doc, unequal } = quartet();
    const before = renderWord(doc);
    strike(unequal, every(unequal), [385, 402, 647, 647]);
    const after = renderWord(doc);
    expect(before).toContain(`height:${twipsToPx(647)}px`);
    expect(after).not.toContain(`height:${twipsToPx(647)}px`);
    expect((after.match(new RegExp(`height:${twipsToPx(520)}px`, "g")) ?? []).length).toBe(3);
    expect(after).toContain(`height:${twipsToPx(521)}px`);
    // 逐行画那一枚地板（`render.ts:rowStyle`）：没碰到的那三张表上的行，一句都不该跟着动。
    // 600 那四行——`equal` 的三行加 `merged` 的末行——都还站在自己那一句上。
    expect(after).toContain(`height:${twipsToPx(600)}px`);
    expect((after.match(new RegExp(`height:${twipsToPx(600)}px`, "g")) ?? []).length).toBe(4);
    expect((after.match(new RegExp(`height:${twipsToPx(200)}px`, "g")) ?? []).length).toBe(1);
  });

  it("signs the whole strike as one revision and lets the refusal put every row back", () => {
    const press = (id: number): { doc: WordDocument; t: Table } => {
      const doc = open(makeRowDistributeDocx());
      const t = doc.blocks.filter((b) => b.kind === "table")[0] as Table;
      const backs = t.rows.map((row) => ({ info: row[0]!.row!, was: rowShapeOf(row[0]!.row!), wasSlots: row[0]!.row?.slots ?? [] }));
      const out = strike(t, every(t), [385, 402, 647, 647]);
      expect(out.ok).toBe(true);
      for (const back of backs) signRowRevision(back.info, { id, author: "顾峰" }, back.was, back.wasSlots);
      return { doc, t };
    };
    const refused = press(601);
    expect(heights(refused.t)).toEqual([521, 520, 520, 520]);
    // 一次落笔 = 一枚 `w:id`：四行各挂一枚 `w:trPrChange`，写出的字节里那枚 id 只有一种。
    const signed = saved(refused.doc);
    expect((signed.match(/<w:trPrChange\b/g) ?? []).length).toBe(4);
    expect(new Set(Array.from(signed.matchAll(/<w:trPrChange w:id="(\d+)"/g), (m) => m[1]))).toEqual(new Set(["601"]));
    expect(refused.doc.resolveRevisions("reject", 601)).toBe(4);
    expect(heights(refused.t)).toEqual([385, 402, 647, 647]);
    const xml = saved(refused.doc);
    expect(xml).toContain('<w:trPr><w:trHeight w:val="385" w:hRule="atLeast"/></w:trPr>');
    expect((xml.match(/<w:trHeight w:val="647" w:hRule="atLeast"\/>/g) ?? []).length).toBe(2);
    expect(xml).not.toContain("trPrChange");

    const kept = press(602);
    expect(kept.doc.resolveRevisions("accept", 602)).toBe(4);
    expect(heights(kept.t)).toEqual([521, 520, 520, 520]);
    expect(saved(kept.doc)).toContain('<w:tblW w:w="3000" w:type="dxa"/>');
  });

  it("keeps the earlier single-row strike answering for one row only", () => {
    // 「分布行」与「单元格高度」用的是同一枚 `w:trHeight` 的落笔，不能把那条按行说一句话的口径挤掉：
    // 一行的落笔仍只写那一行，`w:hRule` 仍跟着行原本那句走。
    const { doc, unequal } = quartet();
    const moved = applyRowHeights(unequal, [{ at: 2, twips: 900 }]);
    expect(moved.length).toBe(1);
    expect(heights(unequal)).toEqual([385, 402, 900, 647]);
    const xml = saved(doc);
    expect(xml).toContain('<w:trPr><w:trHeight w:val="900" w:hRule="atLeast"/></w:trPr>');
    expect((xml.match(/<w:trHeight w:val="385" w:hRule="atLeast"\/>/g) ?? []).length).toBe(1);
  });
});
