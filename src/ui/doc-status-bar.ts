/**
 * 文字页脚的**状态栏**：WPS 的 Writer 在窗口底下常驻一条 `KxWpsStatusCommand id="statusbar"`（`office6/res/wpscommon.kuip`，
 * 442–511 行，逐行数到 **61 个直接子项**——其中 **17 条是 `separator`**，余下 44 项才是读数格；这 61 项里
 * **12 项**带 `minimumWidth`，值只有 72/75/90/105 四种，
 * 我们这八格对应的两条——`WordCountInfoCommand`（457 行）与 `UnitCommand`（459 行）——各带 `minimumWidth="75"`，
 * 页码/页数/节/纵向位置/行/列 六条一条都不带）。`KxWpsStatusWordCountInfoCommand` 还带着 `aid="792"`，
 * detail 说的是「Word Count. Click to open "Word Count" dialog.」——ribbon 那条「字数统计」
 * 就写在**同一个文件的 267 行**、同一个 `aid="792"`，而这一条里其余 60 项一个 `aid` 都没有。所以这一格不是
 * 另一个功能，是同一条命令的另一种长相：报一个数，点一下开同一个对话框。
 *
 * 这一条上挂着**八格**，顺序照 kuip 的子项次序（页码→页数→节→纵向位置→行→列→字数→单位）：
 *
 * - `KxWpsStatusPageNumberInfoCommand id="PageNumberInfoCommand"`，detail「Page Number. Click to open "Go To" dialog.」
 *   ——**第 X 页**，X 是光标（选区活动端）落在的那一页。量法与「字数统计」的 页数 同一套（`doc-layout.ts`
 *   的 `pageAtCaret`：光标自己那枚矩形落进哪一行盒，就是哪一页）。光标不在纸面上（点了批注面板的输入框、
 *   点了页眉页脚带）时**保留上一次的读数**——WPS 真机上那种时候显示什么**未取证**，这一条是自己定的口径。
 * - `KxWpsStatusPageAreaInfoCommand id="PageAreaInfoCommand"`，detail「Page Area Number. Click to open "Go To" dialog.」
 *   ——**共 N 页**，N 与对话框的 页数 同一份测量；光标跑出纸面也照样量得出来。
 * - `KxWpsStatusSectionInfoCommand`（451 行，detail「@Page Section. Click to open "Go To" dialog.」）——**第 X 节**，
 *   X 是光标落在第几节。这一格**不量几何**：节的边界是文件自己声明的（带 `w:pPr/w:sectPr` 的段落结束它那一节），
 *   从模型数（`sectionEndsOf`），块进第几节、节首一块在哪（`sectionCaretAnchor`）的口径都写在 `doc-layout.ts`
 *   文件头。光标不在纸面上时与页码同一口径保留上一次读数。点一下开「定位」对话框、预选「节」目标；
 *   真机字面与点开时预选什么**未取证**。
 * - `KxWpsPagePositionInfoCommand id="PagePositionInfoCommand"`（452 行，detail「@Vertical Page Position. Click to
 *   Open "Go To" dialog.」，夹在 节 与 行/列 之间）——**纵向位置**：光标那行离**纸沿**多远，报当前
 *   **测量单位**下的数（默认厘米两位小数；四种单位与各自小数的口径在 `doc-units.ts`）。
 *   量法在 `doc-layout.ts` 的 `pagePositionAt`（纸沿 = 本页第一枚内容的上沿 − 这一节的上边距；取点与「第 X 页」
 *   同一条中线）。这一格是**按钮**（detail 带 "Click to Open"），点开同一个「定位」对话框；预选哪个目标
 *   未取证，这一条走默认的「页」。真机字面、真机默认哪种单位与几位小数、以及光标不在纸面上时的表现**都未取证**。
 * - `KxWpsStatusRowInfoCommand`（detail「@Line Status Bar」）与 `KxWpsStatusColumnInfoCommand`（detail
 *   「@Column Status Bar」，kuip 里靠它 `rightSibling` 与上一格 `leftSibling` 配成一对）——**行 X／列 X**：
 *   行是光标在**这一页里**排第几（每页重数，0 高的断点记号不数），列是它在这一行的第几个字。
 *   量法在 `doc-layout.ts` 的 `rowColumnAt`（取点与页码同一条中线）。光标不在纸面上时与页码同一口径
 *   保留上一次的读数；表格与只读容器里数不出列，列那一格保留上一次的读数（行照报）。这两格是 **span
 *   不是 button**：kuip 里这两条没有点击动作，没有可点开的东西。真机字面、每页是否重数行、行尾空白
 *   怎么计**都未取证**。
 * - `KxWpsStatusWordCountInfoCommand`（上段那条）——**字数**，点开对话框。
 * - `KxWpsStatusUnitCommand`（459 行，detail「@Measurement unit. Click to open "Options" dialog.」，夹在
 *   字数 与 编码 之间）——**测量单位**：这一格报选项里选的那一种（默认厘米），点一下开「选项」对话框
 *   （`doc-options.ts`，这一版只摆了「测量单位」一组）。选择只活在这个窗口里；真机上这格显示什么字面、
 *   选项里摆哪几种、默认哪种、存去哪里**都未取证**（口径见 `doc-units.ts` 头注）。
 *
 * 页码/页数/节/纵向位置/字数/测量单位 六格是**按钮**：kuip 的 detail 说 页码/页数/节/纵向位置 点一下开「定位」
 * 对话框（`doc-goto.ts`，与 ⌘G 同一条命令 `Id="GoTo"`；点节那一格预选「节」目标），字数那一格是我们先接上的
 * 同形状，单位那一格按 459 行的 detail 开「选项」对话框。
 * WPS 那几格点击的细节与中文界面字面都**未取证**（`.qm` 那堵墙与 `doc-word-count.ts` 头注记的是同一件），
 * 这里按 detail 接上这条腿：种子取光标所在页/节，「定位」把落点交回 `DocEditor.goToPage`／`goToSection`。
 *
 * 还没接的一格：`KxWpsTextEncodingListCommand`（编码）。
 *
 * 还有一条**没接**的腿：`res/wpsentries.kuip` 里 `idPath="Status bar Popup Menu.WordCountStatusBar"` 说明 WPS
 * 的状态栏可以右键弹出一串共 15 项的开关，逐格显示或隐藏。本轮不摆那个弹出菜单：逐格开关要先有「哪些格在
 * 显示」这件事，我们七格是写死的。
 *
 * 样式从这个文件注入；这一条挂在 `document.body` 上而不是编辑器的宿主里——`render()` 会把宿主整个换掉。
 * `body` 本身是 `flex-direction: column`，所以一个 `flex: none` 的直接子元素就落在窗口底下，与表格里那排
 * 页签同一种摆法；换开表格或演示时由外壳把它摘掉，不留到别人页脚上。
 */
import type { DocEditor } from "./doc-edit.js";
import { openOptions } from "./doc-options.js";
import { currentUnit, formatLength, onUnitChange, unitDef } from "./doc-units.js";
import { openWordCount } from "./doc-word-count.js";

const STYLE_ID = "solo-doc-status-styles";

const CSS = `
.solo-doc-status {
  flex: none; display: flex; align-items: center; gap: 10px;
  height: 26px; padding: 0 8px;
  background: #e6e9ed; border-top: 1px solid #cfd4da;
  font: 11.5px/1 -apple-system, "Helvetica Neue", "PingFang SC", sans-serif; color: #5b636d;
}
.solo-doc-status button {
  font: inherit; color: inherit; background: none; border: 0; padding: 2px 6px; border-radius: 4px; cursor: pointer;
  /* 数字用等宽数位，打字时那一格不跳宽。 */
  font-variant-numeric: tabular-nums;
}
.solo-doc-status button:hover { background: #d7dce2; color: #1f2328; }
/* 字数 与 测量单位 两格各按 WPS 声明的 minimumWidth="75" 撑成一条自己的格子
   （wpscommon.kuip:457/:459）。这两格在 kuip 里正好是紧邻的两条（中间只隔着 编码 那条没接的），家里都只按内容宽，
   于是「字数：2」与「厘米」贴在一起读成「字数：2 厘米」——数被读成带了单位。75px 那一档把两格分开；字仍靠左，
   因为内容在这 75px 里怎么摆 WPS **未取证**，靠左是不把数推进别人格子里的那种摆法。
   余下六格（页码/页数/节/纵向位置/行/列）kuip 里都**不带** minimumWidth，这里也不给。 */
.solo-doc-status .solo-doc-status-words, .solo-doc-status .solo-doc-status-unit { min-width: 75px; text-align: left; }
/* 分隔条：位置与条数是 kuip 声明的（见构造函数里那段），画成 1px 细线是我们自己的摆法。 */
.solo-doc-status .solo-doc-status-sep { flex: none; width: 1px; height: 12px; align-self: center; background: #cfd4da; }
.solo-doc-status .solo-doc-status-loc, .solo-doc-status .solo-doc-status-sec, .solo-doc-status .solo-doc-status-vpos, .solo-doc-status .solo-doc-status-rc {
  padding: 0 2px; color: #6b7480;
  /* 页码/总页/节/纵向位置/行/列 都是数，同样用等宽数位。 */
  font-variant-numeric: tabular-nums;
}
@media print { .solo-doc-status { display: none; } }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

export class DocStatusBar {
  private readonly root: HTMLDivElement;
  private readonly pageItem: HTMLButtonElement;
  private readonly pagesItem: HTMLButtonElement;
  private readonly sectionItem: HTMLButtonElement;
  private readonly posItem: HTMLButtonElement;
  private readonly lineItem: HTMLSpanElement;
  private readonly colItem: HTMLSpanElement;
  private readonly wordItem: HTMLButtonElement;
  private readonly unitItem: HTMLButtonElement;
  /** 换单位时状态栏要跟着重读：「纵向位置」那格的数与单位格自己的字面都变。 */
  private readonly offUnit: () => void;
  /** 光标跑出纸面那几格里，页码保留上一次的读数（口径见文件头）；页数不跟光标走，每回现量。 */
  private lastPage = 1;
  /** 节与页码同一口径：光标不在纸面上时保留上一次的读数。 */
  private lastSection = 1;
  private lastSections = 1;
  /** 纵向位置与页码同一口径：光标不在纸面上、或光标行量不出矩形时保留上一次的读数。 */
  private lastPos: number | undefined;
  /** 行/列 与页码同一口径：光标不在纸面上、或数不出列（表格/只读容器）时保留上一次的读数。 */
  private lastRow: number | undefined;
  private lastCol: number | undefined;

  /**
   * 数是从这一个编辑器**现读**的：这几格报的数与对话框里那几行必须出自同一次计算，所以都走
   * `editor.caretPage()`／`editor.caretLine()`（行/列，与「第 X 页」同一趟取点）与 `editor.counts()`，
   * 不在构造时把数抄下来。
   */
  constructor(private readonly editor: DocEditor) {
    injectStyles();
    this.root = document.createElement("div");
    this.root.className = "solo-doc-status";
    this.pageItem = document.createElement("button");
    this.pageItem.type = "button";
    this.pageItem.className = "solo-doc-status-loc";
    this.pageItem.title = "页码。点击打开「定位」对话框";
    this.pagesItem = document.createElement("button");
    this.pagesItem.type = "button";
    this.pagesItem.className = "solo-doc-status-loc";
    this.pagesItem.title = "页数。点击打开「定位」对话框";
    this.sectionItem = document.createElement("button");
    this.sectionItem.type = "button";
    this.sectionItem.className = "solo-doc-status-sec";
    this.sectionItem.title = "节号：光标在第几节。点击打开「定位」对话框";
    // 页码/页数/节/纵向位置 四格与 ⌘G 开的是同一个对话框，种子与范围都由编辑器现量（口径见 `doc-edit.ts` 的 `showGoTo`）；
    // 点哪一格开哪个目标：前两格「页」，节这一格「节」。纵向位置那条 detail 与页码那条同形（都带 "Click to Open"），
    // 真机点开预选什么未取证，这一格走默认的「页」。
    this.pageItem.addEventListener("click", () => this.editor.showGoTo());
    this.pagesItem.addEventListener("click", () => this.editor.showGoTo());
    this.sectionItem.addEventListener("click", () => this.editor.showGoTo("section"));
    this.posItem = document.createElement("button");
    this.posItem.type = "button";
    this.posItem.className = "solo-doc-status-vpos";
    this.posItem.title = "纵向位置：光标那行离纸沿多远。点击打开「定位」对话框";
    this.posItem.addEventListener("click", () => this.editor.showGoTo());
    // kuip 里 行/列 是两条没有点击动作的 status 条目（detail 里没有 "Click to open"），画成 span 而不是按钮。
    this.lineItem = document.createElement("span");
    this.lineItem.className = "solo-doc-status-rc";
    this.lineItem.title = "行号：光标在这一页里排第几行";
    this.colItem = document.createElement("span");
    this.colItem.className = "solo-doc-status-rc";
    this.colItem.title = "列号：光标在这一行的第几个字";
    this.wordItem = document.createElement("button");
    this.wordItem.type = "button";
    this.wordItem.className = "solo-doc-status-words";
    this.wordItem.title = "字数。点击打开「字数统计」对话框";
    this.wordItem.addEventListener("click", () => openWordCount(this.editor.counts(), this.editor.selectionCounts(), this.editor.layout()));
    this.unitItem = document.createElement("button");
    this.unitItem.type = "button";
    this.unitItem.className = "solo-doc-status-unit";
    this.unitItem.title = "测量单位。点击打开「选项」对话框";
    this.unitItem.addEventListener("click", () => openOptions());
    // 摆法照 kuip 声明的邻接：`statusbar` 的 61 条直接子项里有 **17 条 `separator`**，它们的位置是文件自己说的
    // （页码 与 页数 之间、页数 与 节 之间、纵向位置 与 行 之间、列 与 字数 之间、字数 与 单位 之间各一条），
    // 而 节·纵向位置 与 行·列 **中间不隔**——这两组在 kuip 里就是紧邻的两条（`RowInfoCommand`/`ColumnInfoCommand`
    // 还互相带 `rightSibling`/`leftSibling`），本来也是一起读的两句。家里原先只用一道 uniform `gap`：
    // 于是 `字数：2` 与 `厘米` 只隔 10px，读成「字数：2 厘米」（数被安上了单位）。分隔条**画成什么** WPS 未取证，
    // 这里给一条 1px 细线；隔几条、隔在哪两条之间是实据。
    const sep = (): HTMLElement => {
      const s = document.createElement("span");
      s.className = "solo-doc-status-sep";
      s.setAttribute("aria-hidden", "true");
      return s;
    };
    this.root.append(
      this.pageItem, sep(), this.pagesItem, sep(), this.sectionItem, this.posItem,
      sep(), this.lineItem, this.colItem, sep(), this.wordItem, sep(), this.unitItem,
    );
    document.body.append(this.root);
    this.offUnit = onUnitChange(() => this.refresh());
    this.refresh();
  }

  refresh(): void {
    // 一次量完：页码／页数／纵向位置／行／列五格出自同一趟铺排（`caretStatusOf` 为什么要合起来量写在那一支上）。
    const st = this.editor.caretStatus();
    const at = st.placement;
    if (at) this.lastPage = at.page;
    // 光标不在纸面上时几何那几格不给数，但总页数不跟着光标走：那一趟里已经现量过，量不出来才留空。
    const total = at?.pages ?? st.pages;
    // 页码/页面两格的字面照 WPS（#240）：「页码: 1」「页面: 1/5」——页面那一格把当前与总数合在一处，
    // 不再是「共 X 页」的单独一句；`lastPage` 在光标跑出纸面时保留上一次读数（文件头那条口径）。
    this.pageItem.textContent = total === undefined ? "" : `页码: ${this.lastPage}`;
    this.pagesItem.textContent = total === undefined ? "" : `页面: ${this.lastPage}/${total}`;
    const cs = this.editor.caretSection();
    if (cs) {
      this.lastSection = cs.section;
      this.lastSections = cs.sections;
    }
    // 节那格报「节 X/Y」（#239）：WPS 状态栏同格（`KxWpsStatusSectionInfoCommand`）显「节: 1/3」——
    // 光标在第几节、件里一共几节，`caretSection()` 本来就交这一对，状态栏此前只报了前一半。
    this.sectionItem.textContent = total === undefined ? "" : `节 ${this.lastSection}/${this.lastSections}`;
    const pp = st.positionPx;
    if (pp !== undefined) this.lastPos = pp;
    this.posItem.textContent =
      total === undefined || this.lastPos === undefined ? "" : `纵向位置 ${formatLength(this.lastPos)}`;
    const rc = st.line;
    if (rc) {
      this.lastRow = rc.row;
      if (rc.column !== undefined) this.lastCol = rc.column;
    }
    this.lineItem.textContent = total === undefined || this.lastRow === undefined ? "" : `行 ${this.lastRow}`;
    this.colItem.textContent = total === undefined || this.lastCol === undefined ? "" : `列 ${this.lastCol}`;
    const c = this.editor.counts();
    this.wordItem.textContent = `字数：${c.cjk + c.words}`;
    this.unitItem.textContent = unitDef(currentUnit()).name;
  }

  detach(): void {
    this.offUnit();
    this.root.remove();
  }
}
