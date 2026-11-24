# EndNote 引文助手：给 Word 增补可编辑引文 (`dsh-plugin-endnote-cwyw`)

[English](README.en.md) · [完整工作流](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md) · 原稿、原 `.enl` 文献库与 PDF 全程只读

**动库和改文档前，先摊开「要插在哪一句、依据来自 PDF 哪一页原文、会复用已有记录还是新增一条」让你核对；再把增量文献写进独立 `.enl + .Data` 副本并挂上全文 PDF，最终由 Word 官方 EndNote CWYW 渲染出可随时删引、补插、换样式的真动态引文。**

![核对 PDF 原页依据后再写入 Word 的 EndNote 真动态引文与删引实拍](docs/images/endnote-workflow-showcase.png)

## 直接这样对 AI 说

| 你的需要 | 发给助手（附稿件、PDF 目录和 `.enl` 路径） |
| --- | --- |
| 先核对插点位置与 PDF 原页证据 | “用 cite-from-pdfs 为《文稿.docx》预演加引文：PDF 在 `Papers`，EndNote 库是 `Library.enl`。先列目标句 `[+Cite]` 位置、对应 PDF 页码摘录和复用/新增明细，再执行写库。” |
| 确认后交付可编辑 Word + 副本库 | “按确认方案在独立副本上写库加引文，把可编辑 Word、配套 `.enl + .Data` 副本库与检查报告一起交给我。” |
| 静态引文转为真 CWYW 动态域 | “用 endnote-cwyw 检查《文稿.docx》的域代码，在副本上转换成可被 EndNote 实时刷新的动态引文。” |

## 为什么做、有什么用与实测效果

- **告别“看模型运气”且不碰主文献库**：让通用 AI 裸跑操作桌面端 Word 与 EndNote，往往在 COM 接口、域码转义或双库干扰上连写 6–9 套临时脚本仍失败。插件把整条工作链固化并严格隔离在时间戳副本中运行，原 `.enl`、`.Data`、PDF 与 Word 原稿全程只读不改。
- **逐句核对 PDF 原页后再动库，无据句克制留空**：离线解析无网 PDF（跳过页眉排版代号、首页提全作者、支持双栏与前 30 页扫描、常见刊名缩写及跨年 `±1` 匹配）；批量预演一次性标出正文已删改句子和缺摘录条目，PDF 查无依据的断言主动留空不瞎编。
- **期刊标点精确定位 + 真动态域联动**：英文 `Nature*` 样式下引文紧贴句号后零空格（`result.²`），中文置于全角句号前（`结果²。`），支持曲引号与批注段落；实测同篇或改名 PDF 自动合库去重并完整挂载全文附件（裸跑对照多次漏附全文），在 Word + EndNote 21 内直接删掉第 3 处引文并刷新，正文上标与文末参考文献列表随之 `3 → 2 → 3`（及 `2 → 1 → 2`）实时收缩恢复。
- **零误合并的大规模书目基准**：在 CoLRev 公开的卒中（1,292 条）与血液学（1,415 条）去重基准集上，自动判同精确率 **100%（0 误合并）**，模糊临界条目自动单列供人工复核。*注：写库与刷新 CWYW 动态域需在本地 Windows 桌面环境运行 Microsoft Word 与 EndNote。*

## 30 秒安装

关闭 EndNote 桌面窗口后执行（Python 3.10+，缺依赖时可运行 `conda run -n base python -m pip install -r requirements.txt`）：

```powershell
dsh plugin --profile web add (Resolve-Path '.\dsh-plugin-endnote-cwyw-0.3.0.tgz').Path
```

[桌面验证实录](docs/evaluation.md) · [基准对比实验](docs/comparison.md) · [同类工具对比](docs/landscape.md) · [MIT License](LICENSE)
