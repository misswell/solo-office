import { unzipSync } from "fflate";
import { XMLBuilder, XMLParser } from "fast-xml-parser";

export type DocumentKind = "xlsx" | "docx" | "pptx";

/** Key the parser groups attributes under, and the one `ordered()` unpacks again. */
export const ATTRS = ":@" as const;

/**
 * OOXML parts are namespace-prefixed (`w:p`, `a:tbl`, `xdr:sp`). Prefixes are
 * declared per part and carry no semantics here, so the owning boundary drops
 * them once and every reader sees one canonical key shape. Repeated siblings
 * already become arrays, so `asArray()` covers both one and many occurrences
 * without a per-tag array list.
 *
 * The parser is asked for its ordered output because the plain object output
 * groups siblings by tag name, which silently reorders a body — every table
 * would land after every paragraph. `ordered()` converts that back to the
 * object shape readers expect and keeps the sequence under `#order`.
 */
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  attributesGroupName: ATTRS,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  removeNSPrefix: true,
  preserveOrder: true,
});

export interface Rel {
  id: string;
  type: string;
  target: string;
  targetMode: string;
}

/** OPC (ISO 10060 / ECMA-376 Part 2) package: ZIP parts + content types + relationships. */
export class OpcPackage {
  readonly parts: Map<string, Uint8Array>;
  private readonly relsCache = new Map<string, Map<string, Rel>>();
  private readonly contentTypes: Map<string, string>;
  readonly defaultTypes = new Map<string, string>();
  readonly overrideTypes = new Map<string, string>();

  static from(data: Uint8Array): OpcPackage {
    return new OpcPackage(data);
  }

  private constructor(data: Uint8Array) {
    if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) {
      throw new Error("目标不是 ZIP 容器，无法按 OPC 包读取");
    }
    this.parts = new Map(Object.entries(unzipSync(data)));
    const ct = this.parts.get("[Content_Types].xml");
    this.contentTypes = new Map();
    if (ct) {
      const root = parseXml(text(ct)).Types ?? {};
      for (const d of asArray(root.Default)) {
        this.contentTypes.set(String(d["@_Extension"]).toLowerCase(), d["@_ContentType"]);
        this.defaultTypes.set(String(d["@_Extension"]), d["@_ContentType"]);
      }
      for (const o of asArray(root.Override)) {
        this.contentTypes.set(o["@_PartName"], o["@_ContentType"]);
        this.overrideTypes.set(o["@_PartName"], o["@_ContentType"]);
      }
    }
  }

  get kind(): DocumentKind | "unknown" {
    if (this.parts.has("word/document.xml")) return "docx";
    if (this.parts.has("xl/workbook.xml")) return "xlsx";
    if (this.parts.has("ppt/presentation.xml")) return "pptx";
    return "unknown";
  }

  /** Normalized part name as used by [Content_Types].xml overrides. */
  contentType(path: string): string | undefined {
    return this.contentTypes.get("/" + path) ?? this.contentTypes.get(ext(path));
  }

  text(path: string): string | undefined {
    const b = this.parts.get(path);
    return b ? text(b) : undefined;
  }

  bytes(path: string): Uint8Array | undefined {
    return this.parts.get(path);
  }

  xml(path: string): any {
    const s = this.text(path);
    if (s === undefined) throw new Error(`包内缺少部件 ${path}`);
    return parseXml(s);
  }

  /** Relationships for a part (or the package root when `fromPart` is ""). */
  rels(fromPart: string): Map<string, Rel> {
    const cached = this.relsCache.get(fromPart);
    if (cached) return cached;
    const dir = fromPart.includes("/") ? fromPart.slice(0, fromPart.lastIndexOf("/")) : "";
    const base = fromPart.includes("/") ? fromPart.slice(fromPart.lastIndexOf("/") + 1) : "";
    const path = dir ? `${dir}/_rels/${base}.rels` : "_rels/.rels";
    const map = new Map<string, Rel>();
    const raw = this.text(path);
    if (raw) {
      for (const r of asArray(parseXml(raw).Relationships?.Relationship)) {
        map.set(r["@_Id"], {
          id: r["@_Id"],
          type: r["@_Type"],
          target: r["@_Target"],
          targetMode: r["@_TargetMode"] ?? "Internal",
        });
      }
    }
    this.relsCache.set(fromPart, map);
    return map;
  }

  /** Resolve a relationship target against the source part's directory. */
  resolve(fromPart: string, rel: Rel): string {
    if (rel.targetMode === "External") return rel.target;
    const t = rel.target;
    // A leading slash makes the target a package-absolute part name; WPS writes
    // those, so joining it under the source directory would invent `xl/xl/…`.
    if (t.startsWith("/")) return normalize(t);
    if (!fromPart.includes("/")) return normalize(t);
    const dir = fromPart.slice(0, fromPart.lastIndexOf("/"));
    return normalize(`${dir}/${t}`);
  }
}

function ext(path: string): string {
  const i = path.lastIndexOf(".");
  return i < 0 ? path : path.slice(i + 1).toLowerCase();
}

function normalize(p: string): string {
  const out: string[] = [];
  for (const seg of p.split("/")) {
    if (seg === "." || seg === "") continue;
    if (seg === "..") out.pop();
    else out.push(seg);
  }
  return out.join("/");
}

export function text(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

const NAMED_ENTITIES: Record<string, string> = { lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" };

/**
 * Decodes the XML entities a text node still carries. Writers may emit numeric
 * character references for non-ASCII text (openpyxl does), and the parsers leave
 * them in place, so readers route extracted text through here.
 */
export function decodeEntities(s: string): string {
  return s.replace(/&#x([0-9a-fA-F]+);|&#(\d+);|&(lt|gt|quot|apos|amp);/g, (m, hex, dec, named) => {
    if (named) return NAMED_ENTITIES[named];
    const code = hex ? parseInt(hex, 16) : parseInt(dec, 10);
    return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : m;
  });
}

export const asArray = <T>(v: T | T[] | undefined): T[] =>
  v === undefined ? [] : Array.isArray(v) ? v : [v];

function parseXml(source: string): any {
  return ordered(parser.parse(source));
}

/** Parser for trees that will be serialized again: keeps namespace prefixes and xmlns declarations. */
const treeParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  attributesGroupName: ATTRS,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  removeNSPrefix: false,
  preserveOrder: true,
});

/** The parser's ordered nodes, kept raw so a writer can patch an element and serialize it back. */
export function parseOrderedXml(source: string): any[] {
  return treeParser.parse(source) as any[];
}

const builder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  attributesGroupName: ATTRS,
  preserveOrder: true,
  format: false,
  suppressEmptyNode: true,
});

/** Serializes ordered nodes (the shape `parseOrderedXml` returns) back to XML text. */
export function buildXml(nodes: any[]): string {
  return builder.build(nodes) as string;
}

// ------------------------------------------------------------- raw part trees

/** Element name of an ordered-node entry, or undefined for text/attribute entries. */
export function tagOf(entry: any): string | undefined {
  for (const key of Object.keys(entry)) {
    if (key !== "#text" && key !== ATTRS && !key.startsWith("@_")) return key;
  }
  return undefined;
}

/** An ordered element's children as a plain array. */
export function rawChildren(entry: any): any[] {
  const tag = tagOf(entry);
  return tag ? (entry[tag] as any[]) : [];
}

const localName = (tag: string): string => (tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag);

/** Local (unprefixed) name of an ordered-node entry, or undefined for text entries. */
export function tagLocalOf(entry: any): string | undefined {
  const tag = tagOf(entry);
  return tag ? localName(tag) : undefined;
}

/** First element with the given local name, depth-first across an ordered node list. */
export function rawFind(nodes: any[], local: string): any | undefined {
  for (const entry of nodes) {
    const tag = tagOf(entry);
    if (!tag) continue;
    if (localName(tag) === local) return entry;
    const hit = rawFind(rawChildren(entry), local);
    if (hit) return hit;
  }
  return undefined;
}

/** All elements with the given local name, depth-first across an ordered node list. */
export function rawFindAll(nodes: any[], local: string): any[] {
  const out: any[] = [];
  for (const entry of nodes) {
    const tag = tagOf(entry);
    if (!tag) continue;
    if (localName(tag) === local) out.push(entry);
    out.push(...rawFindAll(rawChildren(entry), local));
  }
  return out;
}

/** Attributes of an ordered element, created on demand so a patch can add one. */
export function rawAttrs(entry: any): Record<string, string> {
  if (!entry[ATTRS]) entry[ATTRS] = {};
  return entry[ATTRS];
}

/** A `_rels` part name alongside its part (package root included). */
export function relsPathOf(partPath: string): string {
  const dir = partPath.slice(0, partPath.lastIndexOf("/"));
  const base = partPath.slice(partPath.lastIndexOf("/") + 1);
  return dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
}

/**
 * Turns the parser's ordered output back into the object shape readers expect:
 * attributes as `@_name` keys and repeated siblings as arrays. Nodes with two or
 * more distinct children also carry `#order`, the child sequence, because tag
 * names alone cannot say whether a paragraph came before a table.
 */
function ordered(nodes: any[]): Record<string, any> {
  const out: Record<string, any> = {};
  const order: string[] = [];
  let distinct = 0;
  // In ordered output each entry is one element: `{ tag: children, ":@": attributes }`.
  for (const node of nodes) {
    for (const [name, value] of Object.entries(node)) {
      if (name === ATTRS) continue;
      if (name === "#text") {
        out["#text"] = value;
        continue;
      }
      const child = ordered(value as any[]);
      if (node[ATTRS]) Object.assign(child, node[ATTRS]);
      const existing = out[name];
      if (existing === undefined) {
        out[name] = child;
        distinct += 1;
      } else if (Array.isArray(existing)) existing.push(child);
      else out[name] = [existing, child];
      order.push(name);
    }
  }
  if (distinct > 1) out["#order"] = order;
  return out;
}

/** Direct children in document order, using `#order` when the parser kept it. */
export function children(node: any): { name: string; node: any }[] {
  if (!node || typeof node !== "object") return [];
  const order: string[] | undefined = node["#order"];
  const out: { name: string; node: any }[] = [];
  if (!order) {
    for (const [key, value] of Object.entries(node)) {
      if (key.startsWith("@_") || key === "#text" || key === ATTRS) continue;
      for (const item of asArray(value as any)) out.push({ name: key, node: item });
    }
    return out;
  }
  const seen = new Map<string, number>();
  for (const name of order) {
    const i = seen.get(name) ?? 0;
    seen.set(name, i + 1);
    out.push({ name, node: asArray(node[name])[i] });
  }
  return out;
}

// ------------------------------------------------------------- source spans

/**
 * The content between an element's own open and close tags, `""` when it is self-closing. Takes the
 * text `childSpans` returns, so the outermost tags are the element's.
 */
export function elementContent(xml: string): string {
  const gt = xml.indexOf(">");
  if (gt < 0 || xml[gt - 1] === "/") return "";
  const close = xml.lastIndexOf("</");
  return close > gt ? xml.slice(gt + 1, close) : "";
}

/**
 * Direct children of an element's content, in source order, each with its local tag name and its own
 * source text kept byte for byte. Nesting is tracked by tag pair, so a table inside a cell cannot be
 * mistaken for the cell's own properties; prefixes are honoured per element, so `w:tcPr` and
 * `ns3:tcPr` both read. Comments, processing instructions and CDATA are passed over. This is a scan,
 * not a parse: it exists to hand a writer markup it does not model without rebuilding it.
 */
export function childSpans(content: string): { tag: string; xml: string }[] {
  const out: { tag: string; xml: string }[] = [];
  const stack: string[] = [];
  let start = 0;
  let opened = "";
  let i = 0;
  const closeComment = (from: number, stop: string): number => {
    const at = content.indexOf(stop, from);
    return at < 0 ? content.length : at + stop.length;
  };
  while (i < content.length) {
    const lt = content.indexOf("<", i);
    if (lt < 0) break;
    if (content.startsWith("<!--", lt)) {
      i = closeComment(lt + 4, "-->");
      continue;
    }
    if (content.startsWith("<![CDATA[", lt)) {
      i = closeComment(lt + 9, "]]>");
      continue;
    }
    // A `>` written literally inside a quoted attribute value would end the token early; producers
    // escape it, and a part that does not is reported rather than guessed at.
    const gt = content.indexOf(">", lt + 1);
    if (gt < 0) break;
    const head = content.slice(lt + 1, gt);
    i = gt + 1;
    if (head.startsWith("?")) continue;
    const selfClosing = head.trimEnd().endsWith("/");
    const name = token(head);
    if (name.startsWith("/")) {
      if (stack.at(-1) !== name.slice(1).trim()) {
        stack.length = 0;
        continue;
      }
      stack.pop();
      if (!stack.length) out.push({ tag: localName(opened), xml: content.slice(start, i) });
      continue;
    }
    if (stack.length) {
      if (!selfClosing) stack.push(name);
      continue;
    }
    if (selfClosing) {
      out.push({ tag: localName(name), xml: content.slice(lt, i) });
      continue;
    }
    start = lt;
    opened = name;
    stack.push(name);
  }
  return out;
}

/** The element a token names, without the markup around it: `w:tcPr` for `<w:tcPr ...>` or `<w:tcPr/>`. */
function token(head: string): string {
  return head.trim().replace(/\/$/, "").split(/\s/)[0]!.trim();
}
