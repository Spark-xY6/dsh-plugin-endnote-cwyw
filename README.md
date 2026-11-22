# PDF-Backed Editable EndNote Citations for Word · dsh-plugin-endnote-cwyw

[中文说明](README.zh.md) · [Controlled Comparison](docs/comparison.md) · [Desktop Evaluation](docs/evaluation.md) · [Related Tools](docs/landscape.md)

[![Community checks on main](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml/badge.svg?branch=main)](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/workflows/checks.yml?query=branch%3Amain) · [Checks for v0.1.0](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions/runs/44836288185)

Hand this plugin a partly cited Word draft, a folder of PDFs, and your EndNote library (`.enl` + `.Data`). It verifies exact supporting claims on physical PDF pages, deduplicates and realigns records against your target library without breaking existing numbers, attaches full-text PDFs inside a safe library copy, and runs the official Word EndNote CWYW add-in so collaborators receive a **genuinely editable Word + EndNote library pair**.

## Why Deterministic Tooling Beats Bare Desktop Scripting

Asking a general coding agent to ad-hoc script Windows COM automation, CWYW field encodings, PDF attachment parsing, and multi-library locks hinges heavily on flagship model luck. Outside top-tier models, agents routinely stall or loop through `6–9` broken helper scripts before giving up—and even a flagship baseline (`gpt-6.2-astra`) failed `0/2` times to attach newly cited full-text PDFs into its delivered EndNote library.

Packaging those fragile steps into tested tooling lets everyday models complete the full desktop cycle cleanly. On a shared neutral English draft rendered in Word 16 + EndNote 21, the left page shows verified `1–3` live superscripts while an unsupported 5th claim stays intentionally blank; the right page shows deleting citation `3` directly in Word and watching the reference list automatically shrink **`3 -> 2`**:

![Word 16 live CWYW render showing blank unsupported claim and 3-to-2 shrinking list after deletion](docs/images/endnote-workflow-showcase.png)

## Measured Results & Honest Scope

- **Live editable CWYW fields & PDF attachments**: Delivered Word documents carry real `EN.CITE` + `EN.REFLIST` fields that survived round-trip Word edits (`3 -> 2 -> 3` and `4 -> 3 -> 4`), with **`2/2` newly cited full-text PDFs** attached (SHA-256 verified) and `0–1` helper scripts vs `6–9` when bare.
- **Cross-library rebinding & dual-library lock**: Automatically rebound stale collaborator `#60` to `#1`, reused misspelled DOI-free `#2`, worked strictly on library copies, and refused writes whenever another `.enl` window was already open in the same EndNote process.
- **Deduplication caveat**: On two external DOI/PMID-free bibliographic benchmarks (`stroke` and `haematology`), conservative rules achieved **100% precision** (zero false merges) with `93.11%` and `57.58%` recall—for massive identifier-free legacy libraries, pre-clean with dedicated deduplicator `BibDedupe` (`99.58% / 83.03%` recall). Raw numbers live in [docs/comparison.md](docs/comparison.md).

## 1-Minute Setup & Prompt

Requires Windows desktop with **Microsoft Word + EndNote** (verified on EndNote 21 + Word 16), Python 3.10+, and DeepSeek Harness (`0.2.2-rc.5` tested):

```powershell
git clone --branch v0.1.1 https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw.git
cd dsh-plugin-endnote-cwyw
python -m pip install -r requirements.txt
dsh plugin --profile web add github:Spark-xY6/dsh-plugin-endnote-cwyw#v0.1.1
```

Restart the DSH profile, save Word and close EndNote, keep inputs outside the repo, and ask:

> Use cite-from-pdfs with my PDFs folder, Library.enl (with sibling .Data), and Draft.docx. Cite only where a physical PDF page supports the sentence, leave unsupported claims blank for review, reuse existing library records, work strictly on copies, and deliver an editable Word + EndNote library pair.

Append **"Preview only"** to inspect reuse and import decisions in JSON first. When opening the delivered bundle, **double-click the delivered `.enl` library before opening `manuscript_cwyw.docx`**, and keep `.docx`, `.enl`, and sibling `.Data` together when sharing with collaborators. Further guides: [docs/evaluation.md](docs/evaluation.md), [docs/landscape.md](docs/landscape.md), [workflow-and-troubleshooting.md](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md). Local-first by default; source code under the [MIT License](LICENSE).
