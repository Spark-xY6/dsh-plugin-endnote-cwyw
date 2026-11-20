---
name: cite-from-pdfs
description: "Use a PDF literature collection and a closed EndNote library to add evidence-backed references to a partially cited Word manuscript. Reuse equivalent records, build an updated library copy, and deliver an editable EndNote CWYW DOCX."
---

# Cite from PDFs

Given the PDF directory, old `.enl` with its paired `.Data`, and DOCX, complete the authorized citation update on copies. Provide the resulting Word and its matching updated library together. Existing genuine EndNote citations may be in the target library or in another EndNote library.

1. Use the active Windows Conda `base` interpreter with `-X utf8`. Run `scripts/doctor.py` if Word/EndNote readiness is uncertain. Require a closed EndNote library before snapshots and imports. Never edit either EndNote database via SQLite or close pre-existing Office processes.
2. Run `scripts/index_pdf_folder.py index` using the supplied PDF directory/library and a new work directory. Index locally by default. Enable `--enrich-missing-metadata` only for a collection approved for online DOI, arXiv ID or title lookups. Search its catalog for the manuscript's uncited claims, then read the actual source pages. Use the full-text support and metadata to choose citations. Titles or search scores alone do not establish evidence. Use an approved model provider for any manuscript content read in the agent session.
3. Build the citation plan from the DOCX yourself. Preserve existing citation positions. Add new groups where a page directly supports the statement. Record the exact anchor, PDF ID, page and rationale. Leave uncertain claims or ambiguous records unapplied and report them. If the user authorizes automatic completion, carry unambiguous decisions through without another approval round. If requested to review first, stop at the evidence table.
4. Run the prepare, native format and publish commands in [workflow and troubleshooting](references/workflow-and-troubleshooting.md). Execute the emitted formatter command directly from PowerShell as its own process. This creates a full candidate library, rematches identities, reconciles existing CWYW records, imports only missing works, formats via the official Word add-in and audits the saved outputs. It publishes the pair only after the audit passes.
5. Deliver `manuscript_cwyw.docx`, the paired `library/*.enl` and matching `.Data`, and `audit.json`. Report omitted placements and any document-specific verification limits. The user's source library stays available throughout recovery.

Refer to [workflow and troubleshooting](references/workflow-and-troubleshooting.md) for the plan, duplicate decisions and recovery. For particular existing field types and native Word verification, use [endnote-cwyw](../endnote-cwyw/SKILL.md).
