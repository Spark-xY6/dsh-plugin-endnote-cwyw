# Editable EndNote Citations for Word (`dsh-plugin-endnote-cwyw`)

[简体中文](README.md) · English · [Workflow & Troubleshooting](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md) · Original `.enl`, PDFs & DOCX stay read-only

Preview every **`[+Cite]` sentence anchor and verbatim PDF page excerpt** first, then merge papers into an isolated `.enl + .Data` copy and deliver **live Word EndNote CWYW citations** you can delete, insert, or restyle anytime.

![Verified PDF page-evidence preview and live Word EndNote CWYW delete/refresh workflow](docs/images/endnote-workflow-showcase.png)

## What it handles for you

- **Page-grounded preview before touching anything**: Scans local two-column PDFs offline (up to 30 pages), lists every matched `[+Cite]` sentence with its physical PDF page and quote, flags edited/missing draft lines in one pass, and leaves unsupported claims blank instead of guessing.
- **Your main EndNote library stays untouched**: All imports and deduplication run strictly on timestamped copies (`100%` precision / `0` false merges across 2,707 CoLRev benchmark records), attaching full-text PDFs cleanly (`2/2` vs `0/2` bare).
- **Real editable CWYW fields with exact punctuation**: Formats via installed Word + EndNote CWYW—zero space after `.` in English `Nature*` styles (`result.²`) and before `。` in Chinese (`结果²。`); deleting citation #3 in Word updates the bibliography live (`3 → 2`).

## How to use

1. Close EndNote Desktop cleanly, then install (Windows Word + EndNote, Python 3.10+ required for live CWYW refresh):

```powershell
dsh plugin --profile web add (Resolve-Path '.\dsh-plugin-endnote-cwyw-0.3.0.tgz').Path
```

2. Send your draft, PDF folder, and `.enl` path to your assistant:

| Goal | Prompt |
| --- | --- |
| Preview placements & PDF pages | “Use cite-from-pdfs to preview citations for `Draft.docx` using `Papers/` and `Library.enl`. Show `[+Cite]` sentence previews, PDF pages, and verbatim quotes before editing the library.” |
| Deliver live Word + library copy | “Apply the approved plan on copies and deliver `manuscript_cwyw.docx`, its matching `.enl + .Data` copy, and the audit report.” |
| Repair or convert static fields | “Inspect citation fields in `Draft.docx` and convert static citations into refreshable EndNote CWYW fields on a copy.” |

[Desktop Evaluation](docs/evaluation.md) · [Benchmarks](docs/comparison.md) · [Related Tools](docs/landscape.md) · [MIT License](LICENSE)
