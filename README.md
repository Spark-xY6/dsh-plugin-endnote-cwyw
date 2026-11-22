# PDF-Backed Editable EndNote Citations for Word (`dsh-plugin-endnote-cwyw`)

[中文说明](README.zh.md) · [Controlled Comparison](docs/comparison.md) · [Desktop Evaluation](docs/evaluation.md) · [Related Projects](docs/landscape.md)

[![Community checks on main](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml/badge.svg?branch=main)](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml?query=branch%3Amain) · [Checks for v0.1.0](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/runs/44836288185)

**In plain English:** Hand this plugin a partly cited Word draft, a folder of PDFs, and your EndNote library (`.enl` + `.Data`). It verifies exact supporting quotes on physical PDF pages, deduplicates and matches against your target library without breaking existing record numbers, attaches the new PDFs inside a safe library copy, and runs the official Word EndNote CWYW add-in so your collaborator receives a **fully editable Word + EndNote library pair**.

---

## Why build this?

- **Cross-library record ID collisions**: Record `#60` in a collaborator's draft often points to a completely different paper in your library.
- **Duplicate imports**: Two renamed PDFs of the same paper create duplicate numbers and split bibliography entries.
- **Open-source gaps & bare-agent model fragility**: Running EndNote/Word CWYW automation bare heavily depends on flagship model capability—most weaker models repeatedly fail, loop, or stall on COM window handling, PDF verification, and record ID rebinding with very low efficiency, while even flagship bare runs leave throwaway scripts behind and fail `0/2` to attach newly cited PDFs into the delivered EndNote library.

---

## Measured Input -> Output Result

![dsh-plugin-endnote-cwyw Input to Output desktop showcase](docs/images/endnote-workflow-showcase.png)

| Capability | Bare Agent (`gpt-6.2-astra`) | Open-Source Tools (PubMate / endnote-insert / etc.) | With `dsh-plugin-endnote-cwyw` (`v0.1.1`) |
| :--- | :--- | :--- | :--- |
| **Live CWYW editable fields + round-trip delete/reinsert** | Passed (`4->3->4`, `3->2->3`), wrote 6–9 helper scripts | Require pre-marked IDs/DOIs or emit raw fields without target-library rebinding | **Verified live `EN.CITE` + `EN.REFLIST`**: reopen in Word, delete, reopen, reinsert -> **3 -> 2 -> 3** entries |
| **Full-text PDF attached in delivered library** | **0/2 new papers** attached | Not automated with PDF page verification | **2/2 newly cited PDFs attached** with matching SHA-256 bytes |
| **Cross-library ID collision rebinding (`v0.1.1`)** | Required manual/ad-hoc scripting | Not supported out-of-the-box | Rebound stale `#60` -> `#1`, reused misspelled DOI-free `#2`, imported only 1 missing paper (`60 -> 61`) |
| **Dual-library safety interlock** | None | No same-PID multi-window guard | **Refuses import and preserves both libraries** if another EndNote library window is open in the same process |
| **Retained throwaway scripts** | **6–9 scripts (8.8–10.4 KB)** | — | **0–1 script (0–2.1 KB)** |

*Honest scope note:* On two external DOI/PMID-free bibliographic deduplication datasets (`stroke` and `haematology`), the plugin achieved 100% precision with 93.1% and 57.6% recall (vs specialized deduplicator BibDedupe at 99.6% and 83.0%, 0 false merges for both). See [Controlled Comparison](docs/comparison.md) for all raw metrics and caveats.

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
