/**
 * 目录（TOC）里那一条条 `PAGEREF` 域的**缓存页码**（#248）。
 *
 * 目录条目的形状照 `event plan.dotm` 取证：一段一条，条目字是 HYPERLINK 域的结果，页码是 PAGEREF 域的
 * 结果——两个复杂域并列在同一段里，PAGEREF 的指令写着它指向的书签名（`PAGEREF _Toc27457 \h`），
 * 结果那截 `w:t` 里是**上次更新目录时算出的页码**。WPS「更新目录」做的就是把这些缓存按当前分页重算；
 * 这里做同一件事：给一个「书签 → 页码」的解析器，把每条 PAGEREF 的结果换成新数。
 *
 * 只动**结果**，不动指令：`w:instrText` 一个字节不改（书签名、`\h` 开关都是文件说的话），
 * `w:fldChar` 的三枚边界原样留着——Word/WPS 重新打开时这仍是一条活的域，只是缓存的数刷新了。
 */

export interface TocPageRef {
  /** `PAGEREF` 后面那个书签名（`_Toc27457`）。 */
  name: string;
  /** 结果那截现在的字面（`"1"`，或空串）。 */
  numberText: string;
  /** 结果那截 `w:t` 的内容在整段 xml 里的区间 `[start, end)`——替换只碰这一处。 */
  start: number;
  end: number;
}

/** 指令里认得出的 PAGEREF：`PAGEREF <书签名> [开关…]`。书签名按 JSON/XML 的习惯不含空白。 */
const PAGEREF = /\bPAGEREF\s+([^\s\\}]+)/;

/**
 * 一段（或一整块）xml 里所有 PAGEREF 的结果缓存，按出现次序。
 *
 * 走法是「指令 → 它那一层的 `separate` → 结果里的第一枚 `w:t`」。复杂域的嵌法允许结果里再嵌域
 * （目录条目没有，但形状上合法），所以 separate 只在**同一条** PAGEREF 之后找第一枚，且遇到 `end`
 * 之前没找到就算这条没有结果（读不出数，不猜）。
 */
export function tocPageRefs(xml: string): TocPageRef[] {
  const out: TocPageRef[] = [];
  const instr = /<(?:[\w.-]+:)?instrText\b[^>]*>([^<]*)<\/(?:[\w.-]+:)?instrText>/g;
  for (const m of xml.matchAll(instr)) {
    const hit = PAGEREF.exec(m[1] ?? "");
    if (!hit) continue;
    const from = (m.index ?? 0) + m[0].length;
    const sep = /<(?:[\w.-]+:)?fldChar\b[^>]*\b(?:fldCharType|w:fldCharType)="separate"[^>]*\/?>/.exec(xml.slice(from));
    if (!sep) continue;
    const afterSep = from + (sep.index ?? 0) + sep[0].length;
    const end = /<(?:[\w.-]+:)?fldChar\b[^>]*\b(?:fldCharType|w:fldCharType)="end"[^>]*\/?>/.exec(xml.slice(afterSep));
    const limit = end ? afterSep + (end.index ?? 0) : xml.length;
    const t = /<(?:[\w.-]+:)?t\b[^>]*>([^<]*)<\/(?:[\w.-]+:)?t>/.exec(xml.slice(afterSep, limit));
    if (!t) continue;
    const start = afterSep + (t.index ?? 0) + t[0].indexOf(">") + 1;
    out.push({ name: hit[1]!, numberText: t[1] ?? "", start, end: start + (t[1] ?? "").length });
  }
  return out;
}

/**
 * 把每条 PAGEREF 的缓存换成 `resolve(书签名)`。解析不出（书签不在件里、或那一页算不出）的条目**原样留着**：
 * 拿「1」去填一个算不出的页码，是把没量到的数写进文件。
 *
 * 返回新 xml 与改掉的条数；一条都没变时交回**同一个字符串**（调用方据此不重铺、不报脏）。
 */
export function rewriteTocCachedPages(
  xml: string,
  resolve: (name: string) => number | undefined,
): { xml: string; changed: number } {
  const refs = tocPageRefs(xml);
  let out = xml;
  let changed = 0;
  // 从后往前替换：前面的区间不会因为后面的长度变化而错位。
  for (let i = refs.length - 1; i >= 0; i--) {
    const ref = refs[i]!;
    const page = resolve(ref.name);
    if (page === undefined) continue;
    const next = String(page);
    if (ref.numberText === next) continue;
    out = out.slice(0, ref.start) + next + out.slice(ref.end);
    changed += 1;
  }
  return changed ? { xml: out, changed } : { xml, changed: 0 };
}
