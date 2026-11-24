# Native citation evaluation / 桌面端验收

On Windows with Microsoft Word (`16.0.17932.21140`) and EndNote 21 (`21.0.0 build 17096`), two copies of the same public PDF (`01-rules.pdf` and `02-renamed-copy.pdf`) produced **one deduplicated record** in a copy of Clarivate's 190-record Sample Library (taking active records from `190` to `191`). Word's official Cite While You Write (CWYW) add-in placed the formatted live citation with exact punctuation rules (zero space immediately after `.` under English `Nature` styles, or before `。` for Chinese text) and preserved the pre-existing citation. On reopening the delivered Word + library copy pair, removing the new citation with **Edit & Manage Citation(s)** shrank the bibliography `2 → 1`, and reinserting it restored `1 → 2`, with zero temporary citations remaining. [Machine-readable results](evaluation.json) · [中文说明](evaluation.html)

| Saved Word document | Citation positions | Bibliography entries | Records embedded in the citations |
| --- | ---: | ---: | --- |
| Partially cited input | 1 | 1 | Wiegrebe `#60` |
| Delivered file | 2 | 2 | Wiegrebe `#60`, Sandve `#2470` |
| After removing Sandve in Word | 1 | 1 | Wiegrebe `#60` |
| After reinserting Sandve in Word | 2 | 2 | Wiegrebe `#60`, Sandve `#2470` |
| Re-run against updated 191-record library | 2 | 2 | Wiegrebe `#60`, Sandve `#2470` |

## Pre-import Preview, Deduplication & Copy Isolation

* **Page-grounded preview first.** Read-only batch preview parses local PDFs offline across up to 30 two-column pages, extracts all page-1 authors, filters publisher header slugs, matches journal abbreviations and `±1` Epub years, displays the exact `[+Cite]` sentence anchor and verbatim PDF page excerpt for review, flags any missing/edited draft sentences together in one pass, and refuses to proceed (`REVIEW` / `FAIL`) if an excerpt is absent from the cited PDF page.
* **Copy-only library writes & rollback safety.** All source `.enl`, `.Data`, PDF, and input DOCX SHA-256 digests remain unchanged (`7/7` verified). Re-running with the updated 191-record copy produces `0` duplicate imports.
* **Host & automated suite coverage.** DeepSeek Harness `0.2.2-rc.5` loads both `cite-from-pdfs` and `endnote-cwyw` cleanly in an isolated profile; all `59/59` offline Python unit tests and `2/2` DSH plugin package tests pass.

## Reproduce the inputs

1. Copy EndNote 21's local `Sample_Library_21.enlx` into a new folder, restore `.enl` + `.Data` (`190` active, `2` trashed), confirm Wiegrebe (2008), then close EndNote.
2. Create an input DOCX containing the two sample paragraphs (`examples/seed.md`) with Wiegrebe formatted in `Nature` style via EndNote CWYW.
3. Download Sandve et al. (2013) *Ten Simple Rules for Reproducible Computational Research* as `01-rules.pdf` and duplicate it as `02-renamed-copy.pdf`.
4. Run preview and copy-isolated formatting (`examples/plan.json`, citing Rule 1 on physical page 2) to obtain the delivered Word + library copy pair, then test live deletion/reinsertion inside Word's EndNote tab and inspect with `audit_cwyw_docx.py`.

## Scope

Desktop CWYW formatting and library import require local Windows Microsoft Word and EndNote Desktop; non-interactive/sandboxed desktop sessions cleanly fail fast with an explicit interactive-desktop prompt (`WindowInterference`). Records whose DOI metadata embeds raw carriage returns (such as the sample Yovel 2008 entry) can split into two bibliography paragraphs under Nature style and are rejected by the strict reference-count gate.
