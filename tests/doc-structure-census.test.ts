/**
 * 整包结构普查：一次保存不许让**任何一族**元素或属性从部件里消失，也不许在编辑之外的地方凭空造出族来。
 *
 * 前面每组用例都盯着一个已知形状（断、书签、图片内层、元素自身属性）。这条不带预设：把每个 xml 部件
 * 拆成两张计数表——元素名各出现多少次、`元素@属性名` 各挂多少次——然后整表对拍。所以"某族整族不见了"
 * 这种错，不管它是不是有人想到过，都会被点数对出来。
 *
 * 计数表只看名字与次数，不看正文与属性值：正文由 `doc-package-preservation.test.ts` 那组用例逐节点看着，
 * 两者合起来才是"没动过"。
 *
 * 表里允许三件事，每条都写明了上限或等量条件（见 `explainedByTheEdit`、`cancelSpaceMoves`、`capGrowth`）：
 * 加进去的那一个 run 自己造出 `r`/`t`；按规则重拼 `xml:space` 时在 `@space` 与 `@〔无属性〕` 之间等量搬家；
 * 批注引用标记补 Word 的 `CommentReference` 样式、以及以嵌套表收尾的单元格补那个必需的段落，各以件里真实的
 * 条数为上限。除此之外少一族、多一族都要报。
 */
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { OpcPackage } from "../src/core/opc.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { makeCommentedDocx } from "./helpers/make-comments.js";
import { makeContainerInHeaderDocx, makeContainersDocx, makeFieldCharDocx } from "./helpers/make-containers.js";
import { makeElementAttrsDocx, makeTableFidelityDocx } from "./helpers/make-tables.js";
import { makeParaPropsDocx } from "./helpers/make-paragraph.js";
import { makePreservationDocx } from "./helpers/make-preservation.js";
import { makeStraddleDocx } from "./helpers/make-revisions.js";
import { makeRunPropsDocx } from "./helpers/make-run.js";

type Bag = Map<string, number>;

const bump = (bag: Bag, key: string): void => void bag.set(key, (bag.get(key) ?? 0) + 1);
const localName = (name: string): string => name.split(":").pop() ?? name;

/**
 * `部件|元素` 记一次出现，`部件|元素@属性` 记一次挂载；没有属性的元素记 `@〔无属性〕`，
 * 这样"给本来不带属性的元素加一个属性"也会被算成换族——正文里 `w:t` 的 `xml:space` 正是这一类。
 */
function census(bytes: Uint8Array): Bag {
  const bag: Bag = new Map();
  for (const [path, data] of OpcPackage.from(bytes).parts) {
    if (!path.endsWith(".xml")) continue;
    const xml = Buffer.from(data).toString("utf8");
    for (const m of xml.matchAll(/<(?!\/|!|\?)([\w.:-]+)((?:\s+[\w.:-]+\s*=\s*"[^"]*")*)\s*\/?>/g)) {
      const el = localName(m[1]!);
      bump(bag, `${path}|${el}`);
      const attrs = [...m[2]!.matchAll(/([\w.:-]+)\s*=\s*"/g)].map((a) => localName(a[1]!));
      if (!attrs.length) bump(bag, `${path}|${el}@〔无属性〕`);
      for (const a of attrs) bump(bag, `${path}|${el}@${a}`);
    }
  }
  return bag;
}

/**
 * 一次最小改动：往头一个段落身上加一个 run。文档变脏，保存就必须从模型重建正文。
 * 改动只许这么大——它落在 body 末尾以外的一切结构都不该跟着动，所以多出来的族才有资格被问一句"凭什么"。
 */
function editAndSave(bytes: Uint8Array): Uint8Array {
  const doc = WordDocument.openBytes(bytes);
  const para = doc.blocks.find((b) => b.kind === "paragraph");
  if (!para || para.kind !== "paragraph") throw new Error("件里没有段落，这条普查没跑起来");
  para.inlines.push({ text: "回归" });
  return writeDocx(doc);
}

type Delta = { key: string; before: number; after: number };

/** 那一段的改动自己会造出 run 和文本；这两族（含它们的"无属性"计数）的变化是解释得通的。 */
const explainedByTheEdit = (key: string): boolean => /^word\/document\.xml\|(r|t)(@|$)/.test(key);

/**
 * 两处**定好了上限**的例外。别的都不许凭空长：
 *
 * - 批注引用标记的 `rStyle CommentReference` 是本编辑器按 Word 的写法补的（docx.ts:1009 与 write-docx.ts:372，
 *   doc-comments.test.ts 里有一条用例专门钉住这个形状）。原件给这个 run 写过的属性现在按属性带回，只有连
 *   `w:rPr` 都没有的标记才会被补上——上限 = 件里 `commentReference` 的条数。
 * - 单元格的最后一个块级孩子必须是段落。本机 86 份 WPS 真写出的件、30,227 个 `w:tc`，没有一个以 `w:tbl` 收尾，
 *   而 1,028 处嵌套表后面都跟着一个段落；件里本来有 m 个以表收尾的单元格，`p` 这两族就最多长出 m 条
 *   （write-docx.ts 的 cellBlocksXml 补的那个空段）。
 *
 * 上限的意义在于"多出来的数得出来"：超出的一律照报，普查不许被一句借口关掉。
 */
function capGrowth(list: Delta[], matches: RegExp, cap: number): Delta[] {
  if (!cap) return list;
  return list.flatMap((d) => {
    const grown = d.after - d.before;
    if (grown <= 0 || !matches.test(d.key)) return [d];
    const excess = grown - cap;
    return excess > 0 ? [{ key: d.key, before: d.after - excess, after: d.after }] : [];
  });
}

/** 这两条例外各自欠多少条——都按件里真实的形状数出来，不看用例标签。 */
function owedAllowances(bytes: Uint8Array): { anchors: number; cellsEndingWithTable: number } {
  const xml = Buffer.from(OpcPackage.from(bytes).parts.get("word/document.xml") ?? new Uint8Array()).toString("utf8");
  const count = (re: RegExp): number => (xml.match(re) ?? []).length;
  return {
    anchors: count(/<(?:[\w.:-]+:)?commentReference\b/g),
    cellsEndingWithTable: count(/<\/(?:[\w.:-]+:)?tbl>\s*<\/(?:[\w.:-]+:)?tc>/g),
  };
}

function saveDeltas(before: Bag, after: Bag, owed: { anchors: number; cellsEndingWithTable: number }): string[] {
  const list = cancelSpaceMoves(deltas(before, after, explainedByTheEdit));
  const anchored = capGrowth(list, /^word\/document\.xml\|(rPr|rStyle)(@|$)/, owed.anchors);
  return report(capGrowth(anchored, /^word\/document\.xml\|p(@|$)/, owed.cellsEndingWithTable));
}

function deltas(before: Bag, after: Bag, ignored: (key: string) => boolean): Delta[] {
  const out: Delta[] = [];
  for (const key of new Set([...before.keys(), ...after.keys()])) {
    if (ignored(key)) continue;
    const b = before.get(key) ?? 0;
    const a = after.get(key) ?? 0;
    if (a !== b) out.push({ key, before: b, after: a });
  }
  return out;
}

/**
 * `xml:space` 是按规则写的那一个属性（正文两端有空白才写，见 write-docx.ts 的 preserve 判定），
 * 所以它允许在 `X@space` 与 `X@〔无属性〕` 两族之间**等量**搬家——一次重新拼写完全可能裁掉或补上边空格。
 * 等量这一条本身就把它和"元素没了"分开：元素真丢了，`@〔无属性〕` 不会凭空多出同样的量。
 */
function cancelSpaceMoves(list: Delta[]): Delta[] {
  const by = new Map(list.map((d) => [d.key, d]));
  const paired = new Set<string>();
  for (const d of list) {
    if (!d.key.endsWith("@space") || paired.has(d.key)) continue;
    const twin = `${d.key.slice(0, -"space".length)}〔无属性〕`;
    const other = by.get(twin);
    if (!other || other.after - other.before !== d.before - d.after) continue;
    paired.add(d.key);
    paired.add(twin);
  }
  return list.filter((d) => !paired.has(d.key));
}

const report = (list: Delta[]): string[] => list.map((d) => `${d.key} ${d.before}→${d.after}`).sort();

const fixtures: [string, () => Uint8Array][] = [
  ["图形/断/书签/根声明", makePreservationDocx],
  ["域机制（fldChar/instrText）", makeFieldCharDocx],
  ["批注件", makeCommentedDocx],
  ["容器 sdt", makeContainersDocx],
  ["页眉里的容器", makeContainerInHeaderDocx],
  ["段落格式", makeParaPropsDocx],
  ["run 格式", makeRunPropsDocx],
  ["元素自身属性（paraId/rsid）", makeElementAttrsDocx],
  ["表格保真件", makeTableFidelityDocx],
  ["跨段修订", makeStraddleDocx],
];

/** 抹掉部件里真实存在的一个元素，用来证明这条普查真能看见"少一族"。 */
function sabotage(bytes: Uint8Array, localNameToLose: string): Uint8Array {
  const parts: Record<string, Uint8Array> = {};
  let hit = false;
  for (const [path, data] of OpcPackage.from(bytes).parts) {
    if (path === "word/document.xml" && !hit) {
      const xml = Buffer.from(data).toString("utf8");
      const m = new RegExp(`<(?:[\\w.-]+:)?${localNameToLose}\\b[^>]*(?:/>|>[\\s\\S]*?</(?:[\\w.-]+:)?${localNameToLose}>)`).exec(xml);
      if (m) {
        parts[path] = strToU8(xml.slice(0, m.index) + xml.slice(m.index + m[0].length));
        hit = true;
        continue;
      }
    }
    parts[path] = data;
  }
  if (!hit) throw new Error(`件里没有 ${localNameToLose}，这条对照没跑起来`);
  return zipSync(parts);
}

/** 只摘属性、元素留在原位：`jc@val` 少一条、`jc@〔无属性〕` 多一条，两张表都得动。 */
function sabotageAttr(bytes: Uint8Array, attrLocal: string): Uint8Array {
  const parts: Record<string, Uint8Array> = {};
  let hit = false;
  for (const [path, data] of OpcPackage.from(bytes).parts) {
    if (path === "word/document.xml" && !hit) {
      const xml = Buffer.from(data).toString("utf8");
      const m = new RegExp(`\\s(?:[\\w.-]+:)?${attrLocal}\\s*=\\s*"[^"]*"`).exec(xml);
      if (m) {
        parts[path] = strToU8(xml.slice(0, m.index) + xml.slice(m.index + m[0].length));
        hit = true;
        continue;
      }
    }
    parts[path] = data;
  }
  if (!hit) throw new Error(`件里没有 ${attrLocal} 属性，这条对照没跑起来`);
  return zipSync(parts);
}

describe("docx 保存：整包元素与属性族普查", () => {
  for (const [label, make] of fixtures) {
    it(`${label}：一族的元素或属性都不许少，除加进那一个 run 以外不许凭空多族`, () => {
      const bytes = make();
      const before = census(bytes);
      // 普查必须看得见东西：一件像样的件至少有几十族，否则"全通过"是空转。
      expect(before.size).toBeGreaterThan(30);
      expect(saveDeltas(before, census(editAndSave(bytes)), owedAllowances(bytes))).toEqual([]);
    });
  }

  it("对照：真抹掉一个元素，普查要报出少了一族", () => {
    const bytes = makePreservationDocx();
    const before = census(bytes);
    const after = census(sabotage(bytes, "br"));
    // 对照不走任何抵消与豁免：它要证明的是这张表数得出来。
    expect(report(deltas(before, after, () => false))).toContain("word/document.xml|br 2→1");
  });

  it("对照：属性从元素身上摘走，两张表都要动（少一条 @val，多一条 〔无属性〕）", () => {
    const bytes = makeElementAttrsDocx();
    const before = census(bytes);
    const after = census(sabotageAttr(bytes, "val"));
    const diffs = report(deltas(before, after, () => false));
    expect(diffs.filter((k) => k.includes("@val")).length).toBeGreaterThan(0);
    expect(diffs.filter((k) => k.includes("〔无属性〕") && !k.includes("@space")).length).toBeGreaterThan(0);
  });

  it("对照：抵消规则只认 space 与 〔无属性〕 之间那一次等量搬家", () => {
    const bag = (entries: [string, number][]): Bag => new Map(entries);
    // 等量一来一回：边空格被按规则裁掉，元素一条没少。
    const trimmed = cancelSpaceMoves(
      deltas(
        bag([["word/document.xml|t@space", 4], ["word/document.xml|t@〔无属性〕", 1]]),
        bag([["word/document.xml|t@space", 2], ["word/document.xml|t@〔无属性〕", 3]]),
        () => false,
      ),
    );
    expect(trimmed).toEqual([]);
    // 量对不上（少了 3 条 space，只回来 1 条）：留着报，抵消不许掩盖真正的少。
    expect(
      cancelSpaceMoves(
        deltas(
          bag([["word/document.xml|t@space", 4], ["word/document.xml|t@〔无属性〕", 1]]),
          bag([["word/document.xml|t@space", 1], ["word/document.xml|t@〔无属性〕", 2]]),
          () => false,
        ),
      ).length,
    ).toBe(2);
    // 元素整族少掉、`〔无属性〕` 一分不多：这条永远不该被抵消掉。
    expect(
      cancelSpaceMoves(
        deltas(
          bag([["word/document.xml|br", 2], ["word/document.xml|br@〔无属性〕", 2]]),
          bag([["word/document.xml|br", 1], ["word/document.xml|br@〔无属性〕", 1]]),
          () => false,
        ),
      ).length,
    ).toBe(2);
  });

  it("对照：两条例外只到上限为止，多出来的照样报", () => {
    const bag = (entries: [string, number][]): Bag => new Map(entries);
    const before = bag([["word/document.xml|rPr", 1], ["word/document.xml|p", 1]]);
    const after = bag([["word/document.xml|rPr", 4], ["word/document.xml|p", 4]]);
    // 上限 2，实际各长出 3 条：只有超出的那 1 条报得出来。
    expect(saveDeltas(before, after, { anchors: 2, cellsEndingWithTable: 2 })).toEqual([
      "word/document.xml|p 3→4",
      "word/document.xml|rPr 3→4",
    ]);
    // 件里既没有批注引用、也没有以表收尾的单元格（上限 0）：一条都不许放过。
    expect(saveDeltas(before, after, { anchors: 0, cellsEndingWithTable: 0 })).toEqual([
      "word/document.xml|p 1→4",
      "word/document.xml|rPr 1→4",
    ]);
    // 少掉的永远照报，上限只管"多出来"的那一侧。
    expect(saveDeltas(after, before, { anchors: 9, cellsEndingWithTable: 9 })).toEqual([
      "word/document.xml|p 4→1",
      "word/document.xml|rPr 4→1",
    ]);
  });

  it("新建的件按自己的形状数得出东西来", () => {
    const doc = WordDocument.blank();
    const first = doc.blocks[0];
    if (first.kind === "paragraph") first.inlines.push({ text: "新写的" });
    const bag = census(writeDocx(doc));
    expect(bag.get("word/document.xml|p")).toBe(1);
    expect(bag.get("word/document.xml|r")).toBe(1);
    expect(bag.get("word/document.xml|t@space")).toBeUndefined();
  });
});
