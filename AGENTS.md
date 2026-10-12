# Solo Office 项目记忆

长期使命：用 `/Users/guofeng/Code/solo/rea`（Ghidra 引擎）只读分析本机 WPS Office，推导并验证其行为，
在本目录用 **100% TypeScript + Electron** 独立实现功能等价桌面办公套件。用户要求持续向 WPS 全功能一致推进。
单轮增量不能表述为全套已完全一致；区分观察、推导、未知和实际验证范围。

完整交接、语料和证据位置：`.scratch/HANDOFF.md`、`.scratch/rea-notes.md`。
2026-10-11 用户安排三路并行，文件所有权与交付规则见 `.scratch/PARALLEL.md`；动手前读取。
共享外壳与最终打包装机由主代理协调，其他代理不得并发覆盖。

硬约束：

- `/Applications` 中除自己的 `/Applications/Solo Office.app` 外一个字节不动；WPS 只读分析。
- 独立复现行为，不复制 WPS 的代码、图标、模板、字体或文案到产品。
- `.scratch/wps-corpus/` 111 个文件仅作测试输入，不发布到应用包。
- 不读取用户 Documents/Desktop 中的个人文档作语料。
- 文件读取先 wc -l，超过 200 行分段，每次输出受限。
- 构建、编译、打包、签名、发布及验证类操作委派 `luna_worker`，主代理协调验收。
- 可见 GUI 操作优先 AgentSpace；不可用须说明再回退。临时后台进程按身份核验并收尾。
- 不安装/升级无关软件，不修改 WPS、不安装或升级 Java/Ghidra/Node/Homebrew。

每轮门禁：tsc、vitest、真实 Electron SOLO_SMOKE、独立 openpyxl/python-pptx 反读、111 文件语料回归；
通过后打包、验证签名、安装自己的 app，并汇报最终产物绝对路径。
当前目录没有 git 仓库，未配置 GitHub remote；不得捏造 commit、tag 或 Release。
