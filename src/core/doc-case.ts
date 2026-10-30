/**
 * # 更改大小写的那五档字面
 *
 * WPS 开始页 `RB_Font` 里那枚折叠钮 `Font1`（`CT_Home.kuip:121`，`referenceCommand="ChangeCaseRainbow"`、
 * `idMso="ChangeCaseGallery"`、`ksoCmd="ChangeCase"`、`rbHotKey="7"`；组内它在字号放大/缩小之后、
 * `ClearFormatsRainbow` 之前，`:120-123`）。中文界面字面（「更改大小写」「全部大写」「切换大小写」「词首字母大写」
 * 「小写」）在 `mui/zh_CN/wps.qm` 里都在——但 `.qm` 是一张扁平哈希池，**只证明这些字在包里有**，
 * 不证明下拉里就摆这五档、更不证明它们的次序。这里的档位与档位名按 Word 那五档摆，见 `CASE_OPTIONS`。
 *
 * 这一层只改**字面**：进来一段文本，出去一段**等长**的文本。位置（模型偏移）因此不动，选区、光标、run 的切分
 * 与它们身上的格式都不必重算。汉字没有大小写，五档走一遍一句都不改。
 *
 * 传进来的文本是**模型偏移空间**的那一条流（`flatOf`）：图/断行/制表/锚这些标记占一枚字符但没有字面，用
 * `NUL` 顶着位。标记在这里也是词与句的分界——它自己永远不被改写。
 */

/** 五档：句首字母大写 / 全部小写 / 全部大写 / 词首字母大写 / 切换大小写。 */
export type CaseMode = "sentence" | "lower" | "upper" | "title" | "toggle";

/**
 * ribbon 那一格摆的字面。第一档 `""` 是「这一格是选方向的，不是开关」那个占位（与「转换成文本」同一形状），
 * 落完笔格子回到它。
 */
export const CASE_OPTIONS: readonly [string, string][] = [
  ["", "更改大小写"],
  ["sentence", "句首字母大写"],
  ["lower", "小写"],
  ["upper", "全部大写"],
  ["title", "词首字母大写"],
  ["toggle", "切换大小写"],
];

/**
 * 换一个字的写面。**长度必须不变**，所以换了写面就换了字节数的少数几个字（`ß` 大写是 `SS`、`İ` 小写带一枚
 * 组合点）原样留着：这一层只改字面，位置、run 的切分、选区与光标全靠「进多少出多少」撑着。
 * Word 在那几个字上会展开，我们这一档不展开——一次全大写读到 `ß` 不改，是这一条的读数。
 */
const toCase = (ch: string, upper: boolean): string => {
  const next = upper ? ch.toUpperCase() : ch.toLowerCase();
  return next.length === ch.length ? next : ch;
};

/** 这一个字符有没有大小写可换（`A`↔`a`、`Ω`↔`ω`都算；汉字、数字、标点都不算）。 */
const isCased = (ch: string): boolean => ch.toLowerCase() !== ch.toUpperCase();

/** 有字母身份但没有大小写可换的字（汉字、`ª`以外的多数表意文字）：它不「变大写」，但会结束「这一句/词开头」。 */
const isLetter = (ch: string): boolean => /\p{L}/u.test(ch);

/**
 * 句界：句末标点后重新开始，一段的开头也算。`NUL`（断行/制表那类标记）也是一句之首——断行把上一句截断。
 * 逗号、顿号不算句界。`.` 算句界是 Word 的口径（`Mr. smith` 按句首字母大写会成 `Mr. Smith`），
 * 数字里的那枚点没有字母跟着，改不动谁。
 */
const endsSentence = (ch: string): boolean => "。！？；…!?.;".includes(ch) || ch === "\u0000";

/** 词界：非字母非数字处起新词。撇号不是字母，所以 `don't` 词首大写会落出 `Don'T`（Word 同此读数）。 */
const endsWord = (ch: string): boolean => !/\p{L}|\p{N}/u.test(ch);

/**
 * 按档改写一段文本（模型偏移空间那条流），长度不变。
 *
 * `sentence` 与 `title` 只把「本段开头的第一个有写字」当作句/词之首——Word 也是对**选中的那一段**动手，
 * 不会回头去看选区外面的上一句。
 */
export function caseText(text: string, mode: CaseMode): string {
  if (mode === "upper" || mode === "lower") {
    const upper = mode === "upper";
    let out = "";
    for (const ch of text) out += toCase(ch, upper);
    return out;
  }
  if (mode === "toggle") {
    let out = "";
    for (const ch of text) out += isCased(ch) ? toCase(ch, ch !== ch.toUpperCase()) : ch;
    return out;
  }
  const boundary = mode === "sentence" ? endsSentence : endsWord;
  let atStart = true;
  let out = "";
  for (const ch of text) {
    if (boundary(ch)) {
      atStart = true;
      out += ch;
      continue;
    }
    if (isCased(ch)) {
      out += toCase(ch, atStart);
      atStart = false;
      continue;
    }
    // 汉字那类「是字母但没大小写」的字把开头用完；数字、空格、引号不算开头也不结束开头，
    // 所以 `end. 123abc` 里 `abc` 仍按新一句起头（Word 在读句界时同样只认字与句末标点）。
    if (isLetter(ch)) atStart = false;
    out += ch;
  }
  return out;
}

/**
 * 改完之后**哪些位置**真的变了，按连续区间交回（偏移相对传进来的那段文本）。
 *
 * 逐区间落笔是为了开着修订时签名签得准：整段一起交给自己当量的那条 diff 会把中间**没变**的字也算进改动里
 * 划掉重写——`你好 WORLD hello` 全大写时那三个汉字会被划一道，件里就多了一句没人改过的修订。
 */
export function changedRanges(before: string, after: string): [number, number][] {
  const out: [number, number][] = [];
  let at = 0;
  while (at < before.length) {
    if (before[at] === after[at]) {
      at += 1;
      continue;
    }
    const from = at;
    while (at < before.length && before[at] !== after[at]) at += 1;
    out.push([from, at]);
  }
  return out;
}
