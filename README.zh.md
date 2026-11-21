# 用 PDF 文献库为 Word 增补可编辑的 EndNote 引文

[English](README.md) · [裸跑与开源工具对照](docs/comparison.md) · [桌面评测](docs/evaluation.md) · [同类项目与借鉴](docs/landscape.md)

**v0.1.1 已通过真实换库和编辑验收。** 稿件里的旧记录 60 在目标库里属于另一篇论文。修复版将这处引文重新绑定到正确的 1 号记录，复用缺少 DOI、题名有误的 2 号记录，只新增一篇。重开 Word 删除并插回后，文末篇数为 3 → 2 → 3。原 v0.1.0 在共享桌面停机的记录保留在评测中。[换库试验与各组结果](docs/comparison.md)

[![主分支持续集成](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml/badge.svg?branch=main)](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml?query=branch%3Amain) · [v0.1.0 发布标签的测试](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/runs/44836288185)

合作者发来一份已有部分引文的 Word，你手里有旧 EndNote 库和一批 PDF。这个 DeepSeek Harness 插件从 PDF 找到支持具体句子的原文，在旧库副本中复用或补入条目，最后交付**带 EndNote 域的 Word 和与之配套的新库**。交稿后，合作者仍能在 Word 的 EndNote 选项卡中增删引文。

## 它解决哪个问题

旧稿里的“记录 8”可能指向另一套库中的论文。即便各方使用同一套库，同一论文也可能因为标题拼写略有差异而被导入两次。插件读取 Word 中嵌入的文献信息，与目标库的 DOI、PMID、题名、作者、年份及期刊信息比较。确认等价时优先复用稿件中已经引用的条目。若题名相同却对应不同 DOI，工作流给出需要复核的记录和字段。这样可以在多处引用同一论文，同时保留一个文末条目。

已有工具提供了很好的设计参考。我们借鉴 PubMate 的预演和逐条状态报告，借鉴 JabRef 和 Zotero 的字段对照，把这些方法放在 EndNote 实际导入之前。逐页核对引文证据和交稿后的原生增删测试也纳入了发布流程。[中文同类项目调研](docs/landscape.html)

## 安装与第一次使用

准备一台能交互运行桌面 Word、EndNote 及其 CWYW 插件的 Windows 电脑。DSH 和 Python 均在这台电脑运行。验收环境为 DeepSeek Harness `0.2.2-rc.5`、Node 24.13.0 和 EndNote 21。以下命令在 PowerShell 中运行，Python 使用你要交给 DSH 调用的 3.10 或更新环境。

```powershell
git clone --branch v0.1.1 https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw.git
cd dsh-plugin-endnote-cwyw
python -m pip install -r requirements.txt
dsh plugin --profile web add github:Spark-xY6/dsh-plugin-endnote-cwyw#v0.1.1
```

重新启动该 DSH profile。包内注册了 `cite-from-pdfs` 和 `endnote-cwyw`。也可以把仓库 `skills/` 下的两个目录相邻放入其他支持本地命令的技能宿主。[DSH 官方安装说明](https://deepseek-harness.github.io/deepseek-harness/en/develop/basic/publish)

先保存稿件，正常关闭 EndNote。让 `.enl` 与同名 `.Data` 目录保持相邻，另选一个新目录存放交付物。将以下提示中的路径替换为自己的路径。

> 使用 cite-from-pdfs。PDF 在「文献目录」，旧库是「库文件.enl」，稿件是「稿件.docx」。请逐句检查需要补引文的位置，核对原文页码，同一篇论文尽量复用已有记录。在副本上自动完成，交付可继续增删引文的 Word 和配套库，并列出仍需我判断的位置。

如果希望先看会新增哪些记录，可以说“**只做预演**”。私有 JSON 会列出复用记录、待新增论文、同篇别名和字段差异。正式运行时，插件会核对每条引文的短引句是否出现在指定的 PDF 物理页，调用 Word 中的 EndNote 插件重新排版，然后审查实际保存的域。收到结果后，先打开交付目录里的库，再打开 `manuscript_cwyw.docx`。给合作者转交时，Word、`.enl` 和同名 `.Data` 都要一起保留。[命令与计划示例](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md)

## 如何判断结果

在 EndNote 21 随附示例库上，同一篇公开论文以两个文件名被选中，库从 **190 条增至 191 条**。重开 Word 删除这篇、再插回，三个已保存版本的参考文献数为 **2 → 1 → 2**。改用交付库重跑后仍是 **191 条**，新增为零。

四个独立的 Codex 会话随后拿到同一份未标注的五段稿件、四份 PDF 和旧库。使用技能的两次和裸跑的两次均找对三处引文，给缺乏依据的试验数字留空，交付了四个可编辑引文和三条书目。技能组两次都把两篇新引论文的 PDF 随库交付，裸跑两次都没有附进库中。四次运行共用一篇稿件。DSH 宿主另行验证了插件加载。[对照结果与复现方法](docs/comparison.md) · [原生桌面评测](docs/evaluation.html)

在两个缺 DOI 和 PMID 的公开书目数据集上，本插件自动识别重复对的召回率分别为 93.1% 和 57.6%，专用工具 BibDedupe 分别为 99.6% 和 83.0%。按公开标签核对，两者自动判重均为零误合并。报告也记录了三种开源 EndNote 工具在原稿和人工标记稿上的实际输出。GitHub Actions 用合成材料验证包加载、去重和失败保护。[详细比较](docs/comparison.md) · [持续集成](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions)

自动流程适合正文中已有真正 EndNote 引文的 Word。手打编号、其他引文插件、脚注引文和定制的显示格式需要单独核对。PDF 索引最多读取前 12 页，有需要时可以直接查阅后续页。扫描件需要可提取的文字。软件能验证指定页确有这段话，论点与引文是否相符仍由代理和作者判断。

## 隐私与共建

将论文、文献库和工作目录放在仓库外。默认本地索引。公开文献可按需开启在线元数据查询，可能向 Crossref、OpenAlex 或 PubMed 发送题名或标识符。DSH 所配置的模型服务也可能收到会话里读取的稿件片段。请按稿件的保密要求选择模型和权限。

欢迎提交用合成条目可以复现的错误、其他语言界面的兼容案例或误合并样例。仓库源代码使用 MIT 许可；Word、EndNote 和第三方 Python 包沿用各自条款。[贡献说明](CONTRIBUTING.md) · [开源许可](LICENSE)
