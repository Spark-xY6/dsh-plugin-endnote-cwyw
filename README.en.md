# Editable EndNote Citations & Cross-Library Handoff for Word (`dsh-plugin-endnote-cwyw`)

[简体中文](README.md) · English · [Workflow & Troubleshooting](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md) · Original `.enl`, PDFs & DOCX stay read-only

Built for two high-friction scenarios: **continuing a collaborator's live-cited Word manuscript against an EndNote library with different record numbers (while keeping PDF attachments accessible)**, and **previewing `[+Cite]` sentence anchors + verbatim PDF page quotes before writing a copy-isolated `.enl + .Data` + CWYW DOCX pair**.

![Collaborator cross-library handoff mismatch interception, rebinding, and live editing loop](docs/images/collaborator-handoff-showcase.png)

## Where it adds real value (and what existing tools already cover)

Existing tools already serve single-library or marker-ready workflows well: Clarivate's `CWYW` & `Quote & Cite` handle manual edits cleanly; `felimet/endnote-for-word` and `Breathay/endnote-cwyw` automate single-library insertion and re-open checks; `PubMate`, `Gaoyuan`, and `wmyung` convert pre-tagged or structured drafts; `Zotero` and `Citavi` cover teams able to switch managers.

This plugin concentrates on **cross-library collaborator handoffs without broken links or missing PDFs** and **page-grounded pre-import verification**:

- **1. Safe collaborator handoff across libraries with mismatched record numbers**:
  - **Identity-based rebinding (`rebind_cwyw`)**: When a colleague's Word draft embeds `#60` for one paper while your local library assigns `#60` to an unrelated study, neither EndNote's Traveling Library refresh nor plain file-hash checks catch the mismatch. The plugin matches embedded metadata via DOI, PMID, and title/author/year (tolerant to journal abbreviations and `±1` yr)—in a 60-record cross-library collision benchmark, it rebound `#60 → #1`, reused a no-DOI spelling variant `#2`, added only Wilson as `#61` with `0` edits to existing records, refused writes when multiple libraries were open in EndNote, and passed live deletion/reinsertion (`4 → 3 → 4`, `3 → 2 → 3`) in Word 16 + EndNote 21.
  - **Full-text PDFs attached to the paired library copy**: In paired trials on the same manuscript, plugin sessions attached all new PDFs cleanly (`2/2`) inside `.Data/PDF`, whereas bare tool sessions delivered `0/2` attachments—so the recipient can open cited source PDFs directly inside EndNote.
- **2. Page-grounded preview before touching anything**:
  - Scans local two-column PDFs offline (up to 30 pages), shows every `[+Cite]` anchor with its physical PDF page and verbatim quote, flags missing/edited draft lines in one pass, and leaves unsupported claims blank instead of guessing.

![Verified PDF page-evidence preview and live Word EndNote CWYW delete/refresh workflow](docs/images/endnote-workflow-showcase.png)

- **3. Your main library stays untouched + exact CWYW punctuation**:
  - All imports and deduplication run strictly on timestamped copies (`100%` precision / `0` false merges across 2,707 CoLRev benchmark records) and format via official Word + EndNote CWYW—zero space after `.` in English `Nature*` styles (`result.²`) and before `。` in Chinese (`结果²。`).

## How to use

1. Close EndNote Desktop cleanly, then install (Windows Word + EndNote, Python 3.10+ required for live CWYW refresh):

```powershell
dsh plugin --profile web add (Resolve-Path '.\dsh-plugin-endnote-cwyw-0.3.0.tgz').Path
```

2. Send your draft, PDF folder, and `.enl` path to your assistant:

| Goal | Prompt |
| --- | --- |
| **Collaborator cross-library handoff** | “A collaborator sent live-cited `Handoff.docx` and `Papers/`; my library is `MyLib.enl` with different record IDs. Reconcile and rebind existing citations by bibliographic identity on copies, attach all missing PDFs, and deliver an editable Word + library pair.” |
| Preview placements & PDF pages | “Use cite-from-pdfs to preview citations for `Draft.docx` using `Papers/` and `Library.enl`. Show `[+Cite]` sentence previews, PDF pages, and verbatim quotes before editing the library.” |
| Deliver live Word + library copy | “Apply the approved plan on copies and deliver `manuscript_cwyw.docx`, its matching `.enl + .Data` copy, and the audit report.” |
| Repair or convert static fields | “Inspect citation fields in `Draft.docx` and convert static citations into refreshable EndNote CWYW fields on a copy.” |

[Desktop Evaluation](docs/evaluation.md) · [Benchmarks](docs/comparison.md) · [Related Tools](docs/landscape.md) · [MIT License](LICENSE)
