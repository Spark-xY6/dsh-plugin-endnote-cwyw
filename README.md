# EndNote 引文助手：给 Word 增补可编辑引文 (`dsh-plugin-endnote-cwyw`)

[English](README.en.md) · [排障与完整流程](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md) · 原稿、主 `.enl` 库与 PDF 全程只读

给 Word 论文补引文时，先把**插在哪一句、依据本地 PDF 哪一页哪句原话**摊开让你核对；确认后再写进独立 `.enl` 副本并生成 **Word 真 EndNote 动态引文**。

![核对 PDF 原页依据后再写入 Word 的 EndNote 真动态引文与删引实拍](docs/images/endnote-workflow-showcase.png)

## 它帮你省掉哪几步折腾

- **动库前先看证据，没依据绝不硬塞**：离线扫描双栏 PDF，逐句列出 `[+Cite]` 插点与 PDF 原页摘录；原稿已被删改找不到的句子一次标全，PDF 查无实据之处主动留空不瞎编。
- **完全不碰你的主文献库**：所有导入与去重都在时间戳 `.enl + .Data` 副本上完成（2,707 条公开书目去重测试 **0 误合并**），自动随库挂上全文 PDF 附件（裸跑对照屡次漏附），不怕写脏主库。
- **交付真正能删改的活引文**：调用本机官方 EndNote CWYW 写入动态域——英文 `Nature` 紧贴句号后零空格（`result.²`），中文置于全角句号前（`结果²。`）；在 Word 里删掉第 3 处引文，文末列表随之 `3 → 2` 实时收缩。

## 怎么用

1. 先**关闭 EndNote 桌面窗口**再安装（需本地 Windows 版 Word + EndNote 及 Python 3.10+）：

```powershell
dsh plugin --profile web add (Resolve-Path '.\dsh-plugin-endnote-cwyw-0.3.0.tgz').Path
```

2. 把稿件、PDF 文件夹和 `.enl` 路径直接发给 AI：

| 你要做的事 | 直接这样说 |
| --- | --- |
| 先核对插点与 PDF 原页 | “用 cite-from-pdfs 为《文稿.docx》预演加引文：PDF 在 `Papers`，EndNote 库是 `Library.enl`。先列目标句 `[+Cite]` 位置、对应 PDF 页码摘录和复用/新增明细，再动手写库。” |
| 确认后交付成稿与副本库 | “按确认方案在独立副本上写库加引文，把可编辑 Word、配套 `.enl + .Data` 副本库与检查报告交给我。” |
| 把静态引用转为活域码 | “用 endnote-cwyw 检查《文稿.docx》的域代码，在副本上转换成可被 EndNote 实时刷新的动态引文。” |

[桌面验证实录](docs/evaluation.md) · [基准对比](docs/comparison.md) · [同类方案对比](docs/landscape.md) · [MIT License](LICENSE)
