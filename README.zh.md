# 给 Word 自动补 EndNote 真引文：PDF 原文找证据、防重复入库、交稿后可继续改

[English](README.md) · [裸跑与开源工具实测对照](docs/comparison.md) · [桌面实机验收](docs/evaluation.md) · [同类工具调研](docs/landscape.md)

[![主分支持续集成](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml/badge.svg?branch=main)](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml?query=branch%3Amain) · [v0.1.0 发布验证](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/runs/44836288185)

写论文时经常遇到手头有一份**“半成品 Word + 一堆参考 PDF + 旧 EndNote 库”**：这个插件会先去 PDF 里逐页找支持每句话的原文依据，自动查重对齐你的旧库、把缺的文献连同全文 PDF 一起补进库副本，最后调官方 EndNote 在 Word 中插入**合作者打开还能一键增删联动**的真引文域（不再是死文本上标）。

---

## 为什么不直接让 AI 裸跑操作桌面端？

1. **裸跑极度依赖顶尖模型运气，换普通模型极易反复踩坑失败**：直接让 AI 现场写脚本控制电脑上的 EndNote 和 Word，窗口焦点、CWYW 域码编码、记录号撞号任何一步出错都会卡死或反复编写 6–9 套临时代码仍中断，效率极低。
2. **换库撞号与反复导入**：合作者稿件里的“记录 `#60`”到你的库里往往对应另一篇文献；同一篇文献改名保存两份 PDF 也会插出两个重复编号。
3. **本插件把整条暗坑封装成确定步骤**：普通模型按固定工具链即可一次跑稳——先核对本地 PDF 原文再引用，找不到证据的句子克制留白绝不瞎编，并且全程在库副本与安全锁保护下执行。

---

## 桌面端真实效果实拍（Word 16 + EndNote 21）

下图来自同一通用英文测试稿在真实桌面软件中的导出对比：左图展示前 4 句自动标出真实 `1–3` 上标、黑框第 5 句因本地 PDF 查无依据而主动留白；右图演示直接在 Word 中删除正文第 3 处引用后，文末参考文献列表实时由 **3 条缩减为 2 条**（`0` 手工修补）：

![dsh-plugin-endnote-cwyw 桌面端真联动与 Word 删除引用 3 变 2 实拍图](docs/images/endnote-workflow-showcase.png)

| 对比维度 | 直接让 Agent 裸跑 | 已有开源轮子（PubMate / endnote-insert 等） | 使用 `dsh-plugin-endnote-cwyw`（v0.1.1） |
| :--- | :--- | :--- | :--- |
| **对模型能力的依赖度与稳定性** | 普通模型极易在 COM/编码处卡死失败；即便强基线每轮仍残留 **6–9 份临时脚本** | 多数脚本缺少桌面弹窗/同进程多库安全保护 | 工具步骤封装固化，常规与强模型均稳过，第二轮 **0 份临时脚本残留** |
| **真 EndNote 可编辑域 & 删插闭环** | 强基线虽可写临时脚本勉强生成，但极其依赖临场发挥 | 需人工预先打标 `[2]` 或只输出 ENW/RIS，缺少自动重排 | **交付即真 `EN.CITE` + `EN.REFLIST`**：Word 原生删 1 篇 -> 保存重开 -> 再插回，篇数 **3 → 2 → 3** |
| **新引论文全文 PDF 自动入库** | **0/2 篇**（强基线两轮亦均漏附全文 PDF） | 不负责 PDF 原文页定位与全文入库关联 | **2/2 篇全部把对应全文 PDF 附进交付库**（SHA-256 逐字节一致） |
| **旧库防撞号 & 去重合并** | 需现场写临时代码比对 | 偏重已有编号/DOI 转换，缺少跨库重绑队列 | 跨库将旧 `#60` 自动重绑回正确 `#1`，复用错题名无 DOI 的 `#2`，同篇不同名 PDF 只导 1 条 |
| **防误伤主库安全锁** | 无同进程窗口保护锁 | 不检查同进程已打开的其他库窗口 | 仅操作安全副本；检测到同一 EndNote 进程开着另一库时**直接拒绝写入并保原库完好** |

> **诚实边界说明**：上述端到端对比基于同一 5 段英文测试稿与 EndNote 21 示例库（`190` 条主库与 `60` 条换库校验）；在两个缺 DOI/PMID 的公开书目去重数据集（`stroke` / `haematology`）上，本插件自动判重**精准率 100%**（零误合并），召回率为 `93.11%` 与 `57.58%`（专用去重工具 `BibDedupe` 为 `99.58%` 与 `83.03%`）。详见 [完整实测对照](docs/comparison.md)。

---

## 1 分钟安装与直接套用话术

运行环境：安装有桌面 **Microsoft Word + EndNote（实测 EndNote 21）** 的 Windows 电脑、Python 3.10+ 与 DeepSeek Harness（实测 `0.2.2-rc.5`）。

```powershell
git clone --branch v0.1.1 https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw.git
cd dsh-plugin-endnote-cwyw
python -m pip install -r requirements.txt
dsh plugin --profile web add github:Spark-xY6/dsh-plugin-endnote-cwyw#v0.1.1
```

重启该 DSH profile 后，先保存 Word、关闭 EndNote，把输入文件放在仓库外目录，复制下面这段话发给 DSH：

> 请使用 cite-from-pdfs。PDF 文件夹在「文献目录」，旧库是「库文件.enl」（同名 .Data 在同级目录），稿件是「稿件.docx」。请逐句核对 PDF 原文页码后再决定引用，缺乏依据的句子留空并标注，同一论文尽量复用已有记录。全程只在副本上执行，给我交付带 EndNote 真域的 Word、配套更新库和待复核清单。

如果只想先看匹配方案、暂不写库，在句尾加上：**“先只做预演（preview only）。”**

收到交付物后，**先双击打开交付目录里的 `.enl` 库，再打开 `manuscript_cwyw.docx`**，即可在 Word 的 EndNote 选项卡自由删改；发给合作者时记得把 `.docx`、`.enl` 和同名 `.Data` 文件夹一起打包。

---

## 进阶文档与开源许可

- 📊 **[裸跑 vs 插件 vs 3 款开源 EndNote 工具 vs BibDedupe 详细实测](docs/comparison.md)**
- 🖥️ **[桌面端实机验收与换库复现指南](docs/evaluation.md)** · **[同类项目调研与借鉴来源](docs/landscape.md)**
- ⚙️ **[命令行与计划文件格式说明](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md)**
- **隐私与边界**：默认全本地处理；仅当显式开启在线元数据补全时才会按题名/DOI 查询 Crossref/OpenAlex/PubMed。自动流程适配正文原生 EndNote 域；手打编号、脚注引文与 Zotero 域需单独核对。源代码基于 [MIT 许可](LICENSE)，第三方 Python 依赖及 Word/EndNote 沿用各自许可。
