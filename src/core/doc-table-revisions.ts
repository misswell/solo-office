import { childSpans, elementContent } from "./opc.js";
import type { PropSlot } from "./docx.js";

/**
 * The records a table part carries when somebody revised it and nobody has decided yet:
 * `w:tblPrChange`, `w:trPrChange` and `w:tcPrChange` (a property set from before the change),
 * `w:cellIns` and `w:cellDel` (a cell added or removed), and `w:cellMerge` (one cell's vertical merge
 * changed, with the merge it replaced recorded inside it).
 *
 * This module owns one thing: those records sit *inside* a property set the model carries as source text,
 * so deciding one is a rewrite of that text rather than a rebuild from typed fields. A rebuild would drop
 * every property the model has no field for — `w:tcMar`, `w:cnfStyle`, another reviewer's still-pending
 * `w:cellIns` — which is the loss this slice exists to prevent.
 *
 * The shapes are ECMA-376 §17.13.5. MS's own example of the same family for a paragraph wraps the recorded
 * set in an element of the set's own name (`<w:pPrChange …><w:pPr/></w:pPrChange>`, "the child element of
 * this element contains the complete set of paragraph properties which were applied … before this
 * revision"), so a record is read with that wrapper and, since no file here carries a table record at all,
 * also without it: the two spellings round-trip alike, and a reader that guessed wrong would quietly lose a
 * reviewer's properties. 未取证: no WPS-written table with a revision on it has been examined.
 */

/** The tags that are a revision's record rather than a property of the part that holds them. */
export const TABLE_RECORD_TAGS: ReadonlySet<string> = new Set([
  "tblPrChange",
  "trPrChange",
  "tcPrChange",
  "tblPrExChange",
  "cellIns",
  "cellDel",
  "cellMerge",
]);

/** Which element wraps the recorded set inside a record of this tag — the property set's own name. */
const WRAPPER: Record<string, string> = { tblPrChange: "tblPr", trPrChange: "trPr", tcPrChange: "tcPr", tblPrExChange: "tblPrEx" };

const ATTR_ESC = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * A record this family spells: the signature, then the set as it was wrapped in an element of the set's own
 * name (`<w:tblPrChange …><w:tblPr/></w:tblPrChange>`, §17.13.5). The prefix follows the part the record is
 * hung on, attributes included — a namespace declared as `word:` has no `w:` to borrow. An empty recorded
 * set is written self-closed: that is what "the part stated nothing before" amounts to, and the reader takes
 * the wrapped-empty and the unwrapped-empty spellings alike.
 */
export function changeRecordXml(
  tag: string,
  sign: { id: number; author: string; date?: string },
  recorded: string,
  prefix = "w",
): string {
  const setTag = WRAPPER[tag]!;
  const attrs = `${prefix}:id="${sign.id}" ${prefix}:author="${ATTR_ESC(sign.author)}"${sign.date ? ` ${prefix}:date="${ATTR_ESC(sign.date)}"` : ""}`;
  const set = recorded ? `<${prefix}:${setTag}>${recorded}</${prefix}:${setTag}>` : `<${prefix}:${setTag}/>`;
  return `<${prefix}:${tag} ${attrs}>${set}</${prefix}:${tag}>`;
}

/**
 * The properties a record holds, each as the source text it arrived with. `w:tblPrChange` and its kin
 * record a whole set; `w:cellMerge` records only the one property it touched. A record nested inside a
 * record is skipped: the model holds at most one pending record per set, and a snapshot of the properties
 * is not supposed to bring along somebody else's undecided change.
 */
export function recordedSlots(recordXml: string, tag: string): PropSlot[] {
  const kids = childSpans(elementContent(recordXml));
  const wrapped = WRAPPER[tag] ? kids.find((k) => k.tag === WRAPPER[tag]) : undefined;
  const inner = wrapped ? childSpans(elementContent(wrapped.xml)) : kids;
  return inner.filter((k) => !TABLE_RECORD_TAGS.has(k.tag)).map((k) => ({ tag: k.tag, xml: k.xml }));
}

/** Whether refusing this record puts a whole property set back or only the property the record names. */
export const coversWholeSet = (tag: string): boolean => tag !== "cellMerge";

/**
 * 这一枚记录说的是**属性**还是说的**这一格在不在**。`w:tcPrChange` 与 `w:cellMerge` 记的是属性套（拒绝或
 * 「显示原文」都按 `prev` 摆）；`w:cellIns` / `w:cellDel` 记的是格子的存在，套里没有属性可摆，把它当属性套
 * 摆回去等于抹掉一格从没说过的属性。判定只许有一处：`docx.ts` 的「拒绝」与「显示原文」两张脸问的是同一句话。
 */
export const isPropertyRecord = (tag: string): boolean => tag === "tcPrChange" || tag === "cellMerge";

/**
 * The children a set has after one of its records is decided. Accepting lets the set as it now stands stay
 * and takes the record out; refusing puts the recorded children back — in place of the whole set when the
 * record was a `w:*PrChange`, in place of the one property it names when it was a `w:cellMerge` — and keeps
 * the records of every other revision, each after the properties, which is where the schema puts them.
 */
export function decideSlots(
  slots: PropSlot[],
  rec: { tag: string; prevSlots: PropSlot[] },
  outcome: "accept" | "reject",
): PropSlot[] {
  const rest = slots.filter((s) => s.tag !== rec.tag);
  if (outcome === "accept") return rest;
  // A record's snapshot never carries the record itself back — that is the one being decided.
  const prev = rec.prevSlots.filter((s) => s.tag !== rec.tag);
  if (!coversWholeSet(rec.tag)) return replaceTags(rest, prev);
  // The recorded set first, then whatever other revision still waits: records sit after properties.
  return [...prev, ...rest.filter((s) => TABLE_RECORD_TAGS.has(s.tag))];
}

/** The set with each property the record names put back where the one it replaces sat. */
function replaceTags(slots: PropSlot[], prev: PropSlot[]): PropSlot[] {
  const left = [...prev];
  const out: PropSlot[] = slots.map((s) => {
    const at = left.findIndex((p) => p.tag === s.tag);
    return at < 0 ? s : left.splice(at, 1)[0]!;
  });
  return out.concat(left);
}

/**
 * The same element carrying different children. Its own opening tag is reused, so a prefix bound on the
 * element survives the rewrite; an element left with no children at all becomes nothing, which is what an
 * empty `w:tblPr` or `w:tcPr` amounts to.
 */
export function setChildren(xml: string, kids: PropSlot[]): string {
  const gt = xml.indexOf(">");
  if (gt < 0) return xml;
  const head = xml.slice(0, gt + 1);
  const open = head.endsWith("/>") ? `${head.slice(0, -2)}>` : head;
  if (!kids.length) return "";
  const name = /^<([^/\s>]+)/.exec(open)?.[1];
  return name ? `${open}${kids.map((k) => k.xml).join("")}</${name}>` : xml;
}

/** A carried property set with one of its records decided. */
export function decideSetXml(xml: string, rec: { tag: string; prevSlots: PropSlot[] }, outcome: "accept" | "reject"): string {
  return setChildren(xml, decideSlots(childSpans(elementContent(xml)).map((k) => ({ tag: k.tag, xml: k.xml })), rec, outcome));
}
