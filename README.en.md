# Editable EndNote Citations for Word (`dsh-plugin-endnote-cwyw`)

[简体中文](README.md) · English · [Workflow & Troubleshooting](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md) · Original `.enl`, PDFs & DOCX stay read-only

**Asking a general LLM to drive desktop Word and EndNote directly often burns 6–9 throwaway scripts on COM permissions, CWYW field escaping, and library focus clashes—while touching your main `.enl` risks corrupting years of references. This plugin separates evidence verification from desktop execution: preview every `[+Cite]` sentence anchor and PDF page quote first, merge into a timestamped copy, and deliver live Word EndNote CWYW citations you can delete or restyle anytime.**

![Verified PDF page-evidence preview and live Word EndNote CWYW delete/refresh workflow](docs/images/endnote-workflow-showcase.png)

## How it works in practice

- **Strict read-only isolation on originals**: Keeps your source `.enl`, `.Data`, PDFs, and `.docx` untouched; all imports and Word formatting run exclusively on timestamped copies with deterministic library-binding checks.
- **Offline page-grounded preview before writing anything**: Scans two-column PDFs across up to `--max-pages 30` without network calls, filters publisher header slugs, extracts all page-1 authors, matches journal abbreviations and `±1` Epub years, surfaces all edited/missing manuscript sentences in a single batch preview, and leaves unsupported claims blank rather than guessing references.
- **Exact journal punctuation & live CWYW bibliography updates**: Places English `Nature*` citations immediately after `.` with zero space (`result.²`), places Chinese citations before `。` (`结果²。`), deduplicates renamed PDFs on import with **`2/2` full-text attachments** preserved (vs `0/2` in bare runs), and updates live inside Word + EndNote 21 (`3 → 2 → 3` and `2 → 1 → 2` when citations are deleted and reinserted).
- **100% precision on public deduplication benchmarks**: Achieved **`100%` precision (`0` false merges)** across CoLRev’s 1,292-record stroke and 1,415-record haematology benchmarks (`93.1%` / `57.6%` recall, routing ambiguous pairs to manual review). *(Note: `.enl` import and CWYW field formatting require an interactive Windows desktop with Microsoft Word and EndNote.)*

## Example prompts

| Goal | Prompt (provide manuscript, PDF folder & `.enl` path) |
| --- | --- |
| Preview placements & PDF pages first | “Use cite-from-pdfs to preview citations for `Draft.docx` using `Papers/` and `Library.enl`. Show `[+Cite]` sentence previews, PDF pages, verbatim quotes, and reuse vs add decisions before editing the library.” |
| Deliver live Word + library copy | “Apply the approved plan on copies and deliver `manuscript_cwyw.docx`, its matching `.enl + .Data` copy, and the audit report.” |
| Convert or repair CWYW fields | “Inspect citation fields in `Draft.docx` and convert static citations into refreshable EndNote CWYW fields on a copy.” |

## Installation (Windows Word + EndNote, Python 3.10+)

Close EndNote cleanly before running:

```powershell
dsh plugin --profile web add (Resolve-Path '.\dsh-plugin-endnote-cwyw-0.3.0.tgz').Path
```

[Desktop Evaluation](docs/evaluation.md) · [Benchmarks](docs/comparison.md) · [Related Tools](docs/landscape.md) · [MIT License](LICENSE)
