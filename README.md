# PDF-Backed Editable EndNote Citations for Word (`dsh-plugin-endnote-cwyw`)

[中文说明](README.zh.md) · [Controlled Comparison](docs/comparison.md) · [Desktop Evaluation](docs/evaluation.md) · [Related Projects](docs/landscape.md)

[![Community checks on main](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml/badge.svg?branch=main)](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml?query=branch%3Amain) · [Checks for v0.1.0](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/runs/44836288185)

Hand this plugin a partly cited Word draft, a folder of PDFs, and your EndNote library (`.enl` + `.Data`). It verifies exact supporting quotes on physical PDF pages, deduplicates and aligns against your target library without breaking existing record numbers, attaches full-text PDFs inside a safe library copy, and invokes the official Word EndNote CWYW add-in so collaborators receive a **genuinely editable Word + EndNote library pair**.

---

## Why Deterministic Tooling Beats Bare Desktop Scripting

- **Bare runs depend heavily on flagship model luck**: Asking an ordinary coding agent to ad-hoc script Windows COM interfaces, CWYW field encodings, PDF attachment parsing, and multi-library locks leads to repeated loops, errors, or dead superscript text—and even a flagship baseline (`gpt-6.2-astra`) left `6–9` throwaway scripts behind and attached `0/2` full-text PDFs.
- **Cross-library record ID collisions**: Record `#60` in a collaborator's document often maps to a completely different paper in your local EndNote library.
- **Duplicate imports**: Two renamed copies of the same PDF produce duplicate citation numbers and split bibliographies.

---

## Live Word 16 + EndNote 21 Desktop Result

Rendered from Word 16 on a shared neutral English benchmark: **Left** shows live superscripts `1–3` inserted from physical PDFs while the unsupported 5th claim stays intentionally blank; **Right** shows deleting citation `3` directly inside Microsoft Word and watching the reference list automatically shrink **`3 → 2`**:

![dsh-plugin-endnote-cwyw live Word 16 CWYW deletion 3 to 2 showcase](docs/images/endnote-workflow-showcase.png)

| Capability | Bare Agent (`gpt-6.2-astra`) | Open-Source Tools (PubMate / endnote-insert / etc.) | With `dsh-plugin-endnote-cwyw` (`v0.1.1`) |
| :--- | :--- | :--- | :--- |
| **Model stability & clean workspace** | Ordinary models stall on COM/encoding loops; strong baseline left **6–9 ad-hoc scripts (8.8–10.4 KB)** | Often lack same-PID dual-library safety locks | Deterministic pipeline; reliable without top-tier model luck (**0–1 helper script**) |
| **Live CWYW editable fields (`3 -> 2 -> 3`)** | Passed via fragile custom scripts | Require pre-marked IDs/DOIs or leave bibliography unrefreshed | **Verified live `EN.CITE` + `EN.REFLIST`**: delete in Word -> reopen -> reinsert -> **3 -> 2 -> 3** entries |
| **Full-text PDF attached in library** | **0/2 new papers** attached | Not coupled with physical PDF page verification | **2/2 newly cited PDFs attached** with verified SHA-256 bytes |
| **Cross-library ID collision rebinding (`v0.1.1`)** | Required manual scripting | Not supported out-of-the-box | Rebound stale `#60` -> `#1`, reused misspelled DOI-free `#2`, imported only 1 missing paper (`60 -> 61`) |
| **Dual-library safety interlock** | None | No same-PID multi-window guard | **Refuses writes and preserves both libraries** if another EndNote library is open in the same process |

*Honest scope note:* On two external DOI/PMID-free bibliographic deduplication datasets (`stroke` and `haematology`), the plugin achieved **100% precision** (zero false merges) with `93.11%` and `57.58%` recall (vs dedicated deduplicator `BibDedupe` at `99.58%` and `83.03%`). See [Controlled Comparison](docs/comparison.md) for all raw metrics and caveats.

---

## 1-Minute Install & Prompt

Requires Windows desktop with **Microsoft Word + EndNote** (verified on EndNote 21 + Word 16), Python 3.10+, and DeepSeek Harness (`0.2.2-rc.5` tested).

```powershell
git clone --branch v0.1.1 https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw.git
cd dsh-plugin-endnote-cwyw
python -m pip install -r requirements.txt
dsh plugin --profile web add github:Spark-xY6/dsh-plugin-endnote-cwyw#v0.1.1
```

Restart the DSH profile, save Word and close EndNote, keep inputs outside the repo, and ask:

> Use cite-from-pdfs with my PDFs folder, Library.enl (with sibling .Data), and Draft.docx. Cite only where a physical PDF page supports the sentence, leave unsupported claims blank for review, reuse existing library records, work strictly on copies, and deliver an editable Word + EndNote library pair.

Add **"Preview only"** to inspect proposed reuse/import decisions in JSON before touching any Word or library copy.

---

## Further Reading & License

- 📊 **[Controlled Comparison vs Bare Agent, 3 Open-Source Tools & BibDedupe](docs/comparison.md)**
- 🖥️ **[Desktop Reproduction & Acceptance Guide](docs/evaluation.md)** · **[Landscape & Adopted Ideas](docs/landscape.md)**
- ⚙️ **[CLI Commands & Plan Schema](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md)**
- **License & Privacy**: Local-first indexing (up to 12 pages per PDF automatically; agent can inspect later pages). Optional Crossref/OpenAlex/PubMed metadata lookups. Source code under [MIT](LICENSE); third-party Python dependencies and commercial Office/EndNote remain under their respective terms.
