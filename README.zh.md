# 给 Word 自动补 EndNote 真引文：PDF 原文找证据、防重复入库、交稿后可继续改

[English](README.md) · [裸跑与开源工具实测对照](docs/comparison.md) · [桌面实机验收](docs/evaluation.md) · [同类工具调研](docs/landscape.md)

[![主分支持续集成](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml/badge.svg?branch=main)](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml?query=branch%3Amain) · [v0.1.0 发布验证](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/runs/44836288185)

**一句话看懂它干嘛**：你把“半成品 Word + 一堆 PDF + 旧 EndNote 库”交给它，它先在 PDF 里逐页找支持每句话的原文，再自动查重对齐你的旧库、把缺的文献连同 PDF 一起补进库副本，最后用官方 EndNote 插件在 Word 里生成**合作者打开还能继续一键增删**的真引文域（不再是死文本）。

---

## 为什么做这个插件？

写论文或跟合作者来回改稿时，最怕三件事：
1. **换库就撞号**：合作者稿子里的“记录 60”是甲论文，到你的 EndNote 库里记录 60 却变成了乙论文；
2. **一篇文章导成两条**：同一个文献改名又下了一份 PDF，插出两个不同编号，文末参考文献重复；
3. **开源脚本只产假编号/临时括号**：很多工具直接往 Word 塞死文本或临时 `{Author, Year}`，不仅不附 PDF 全文到库里，也经常破坏原来引文，甚至误操作把你的主文献库弄乱。

---

## 先看 Input -> Output 真实桌面实测效果

![dsh-plugin-endnote-cwyw Input 到 Output 真实桌面实测效果](docs/images/endnote-workflow-showcase.png)

| 维度 | 直接让 Agent 裸跑 | 已有开源轮子（PubMate / endnote-insert 等） | 使用 `dsh-plugin-endnote-cwyw`（v0.1.1） |
| :--- | :--- | :--- | :--- |
| **真 EndNote 可编辑域 & 删插闭环** | 4/4（两轮均生成真域，但全靠现场写临时脚本） | 需人工预先打标 `[2]` 或只输出 ENW/RIS，缺少自动跨库重绑 | **交付即真 `EN.CITE` + `EN.REFLIST`**：Word 原生删 1 篇 -> 保存重开 -> 再插回，篇数 **3 -> 2 -> 3** |
| **新引论文全文 PDF 自动入库** | **0/2 篇**（两轮都没有把 PDF 附入交付库） | 不负责 PDF 原文页定位与全文入库关联 | **2/2 篇全部把对应全文 PDF 附进交付库**（SHA-256 逐字节核验一致） |
| **旧库防撞号 & 去重合并** | 需临时写代码比对 | 偏重已有编号/DOI 转换，未内置跨库冲突复核队列 | 跨库实测将旧 `#60` 自动重绑回正确 `#1`，复用错题名无 DOI 的 `#2`，两份同篇 PDF 只导入 1 条 |
| **防误伤主库安全锁** | 无进程窗口锁 | 多数脚本不检查同进程多库窗口 | 仅操作安全副本；检测到同一 EndNote 进程开着另一个库时**直接拒绝写入并保原库完好** |
| **一次运行留存临时代码** | 留存 **6–9 份临时脚本**（8.8–10.4 KB） | — | 第二轮 **0 份临时脚本**（第一轮 1 份自检脚本 2.1 KB） |

> **实测边界诚实说明**：上述对比基于同一篇 5 段合成稿件与 EndNote 21 示例库（`190` 条库与 `60` 条换库测试）；在两个缺 DOI/PMID 的公开书目去重数据集（`stroke` / `haematology`）上，本插件自动判重零误合并，召回率为 `93.1%` 和 `57.6%`（专用去重工具 BibDedupe 为 `99.6%` 和 `83.0%`）。详见 [完整实测对照](docs/comparison.md)。

---

## 1 分钟安装与直接套用话术

运行环境：安装有桌面 **Microsoft Word + EndNote（实测 EndNote 21）** 的 Windows 电脑、Python 3.10+ 与 DeepSeek Harness（实测 `0.2.2-rc.5`）。

```powershell
git clone --branch v0.1.1 https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw.git
cd dsh-plugin-endnote-cwyw
python -m pip install -r requirements.txt
dsh plugin --profile web add github:Spark-xY6/dsh-plugin-endnote-cwyw#v0.1.1
```

重启该 DSH profile 后，先保存 Word、关闭 EndNote，把输入文件放在仓库外的目录，直接复制下面这段话发给 DSH：

> 请使用 cite-from-pdfs。PDF 文件夹在「文献目录」，旧库是「库文件.enl」（同名 .Data 在同级目录），稿件是「稿件.docx」。请逐句核对 PDF 原文页码后再决定引用，缺乏依据的句子留空并标注，同一论文尽量复用已有记录。全程只在副本上执行，给我交付带 EndNote 真域的 Word、配套更新库和待复核清单。

如果只想先看不写入库，在句尾加上：**“先只做预演（preview only）。”**

收到交付物后，**先双击打开交付目录里的 `.enl` 库，再打开 `manuscript_cwyw.docx`**，即可在 Word 的 EndNote 选项卡自由删改。发给合作者时记得把 `.docx`、`.enl` 和同名 `.Data` 文件夹一起打包。

---

## 进阶文档与开源许可

- 📊 **[裸跑 vs 插件 vs 3 款开源 EndNote 工具 vs BibDedupe 详细实测](docs/comparison.md)**
- 🖥️ **[桌面端实机验收与换库复现指南](docs/evaluation.md)** · **[同类项目调研与借鉴来源](docs/landscape.md)**
- ⚙️ **[命令行与计划文件格式说明](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md)**
- **隐私与边界**：默认全本地处理；仅当显式开启在线元数据补全时才会按题名/DOI 查询 Crossref/OpenAlex/PubMed。自动流程适配正文原生 EndNote 域；手打编号、脚注引文与 Zotero 域需单独核对。源代码基于 [MIT 许可](LICENSE)，第三方 Python 依赖及 Word/EndNote 沿用各自许可。
