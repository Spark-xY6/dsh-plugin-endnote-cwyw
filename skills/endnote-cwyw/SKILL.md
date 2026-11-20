---
name: endnote-cwyw
description: Insert, update, recover, and validate genuine EndNote Cite While You Write (CWYW) citations in Microsoft Word DOCX files on Windows using the installed EndNote Word add-in. Use when converting temporary `{Author, Year #Record}` or static numbered citations into editable EndNote fields, adding references to an existing CWYW document, refreshing an EndNote bibliography or output style, verifying `EN.CITE`/`EN.CITE.DATA`/`EN.REFLIST` and Traveling Library metadata, or diagnosing Word/EndNote COM failures.
---

# EndNote CWYW

Create real, refreshable EndNote fields through the official Word add-in. Treat visible citation text as insufficient evidence: a genuine result must contain EndNote field codes and embedded record metadata.

## Non-negotiable rules

- Preserve the named source and backup. Freeze their SHA-256 hashes and work on a separately named candidate.
- Never synthesize `EN.CITE`, `EN.CITE.DATA`, `EN.REFLIST`, `EN.Layout`, or Traveling Library XML as the primary insertion method.
- Use EndNote-recognized temporary citations such as `{Example, 2024 #1}` or `{Example, 2024 #1; Sample, 2025 #2}`, then invoke EndNote's official `ENFormatBibliography` macro. These example numbers are illustrative.
- Read a frozen copy of `<library>.Data\sdb\sdb.eni` for record and group evidence. The `.enl` file is not the readable record database.
- Never terminate existing Word or EndNote processes without explicit user approval. In particular, treat invisible `/Automation -Embedding` Word processes as potentially holding unsaved state.
- Do not claim visual or pagination QA unless the final DOCX was actually rendered or opened and inspected.

Before live Word/EndNote automation or failure recovery, read [references/windows-cwyw-workflow.md](references/windows-cwyw-workflow.md).

## Choose the route

- **Audit an existing DOCX:** run `scripts/audit_cwyw_docx.py` only.
- **Convert static citations:** map every reference to a verified EndNote record, replace each static citation with an unformatted EndNote citation, remove the obsolete static bibliography only when performing a complete conversion, then format the isolated candidate.
- **Add citations to a genuine CWYW DOCX:** use [cite-from-pdfs](../cite-from-pdfs/SKILL.md) for PDF-backed additions and an updated library copy. Preserve existing fields when their embedded identities and record numbers resolve to that library. For a foreign library or duplicate record numbers for one work, `scripts/rebind_cwyw.py` checks embedded records and prepares verified temporary citations for native reformatting. Disclose any unsupported layouts or display overrides.
- **Refresh changed EndNote metadata:** correct the records in EndNote, use the official update command, and verify that the refreshed Traveling Library contains the new values. If EndNote retains stale embedded metadata, rebuild the affected citations through the temporary-citation route.
- **Repair display-only output after COM is blocked:** preserve the genuine fields and make only an explicitly bounded OOXML display/style correction. Label this as a fallback; it does not synchronize library metadata.

## Workflow

### 1. Freeze and inspect

Record the source, backup, EndNote `.enl`, and copied `sdb.eni` paths, sizes, modification times, and SHA-256 hashes. Verify bibliographic identity and target library ID before accepting a record number. Inspect the current document before deciding whether it is static, mixed, or genuine CWYW:

```powershell
conda run -n base python -X utf8 <skill-dir>\scripts\audit_cwyw_docx.py `
  --docx <source.docx> --report <audit-before.json>
```

If the document already has genuine fields, never flatten them to text.

### 2. Freeze and query the EndNote database

Close EndNote cleanly before copying the database. Copy `<library>.Data\sdb\sdb.eni` into the work directory, hash it, and query only that snapshot:

```powershell
conda run -n base python -X utf8 <skill-dir>\scripts\inspect_endnote_snapshot.py `
  --database <copied-sdb.eni> --group "Example group" `
  --record-ids 1,2 --report <endnote-snapshot.json>
```

Require exact title/author/year evidence for each record. Record numbers are library-specific and must never be reused from an older snapshot without revalidation.

### 3. Prepare unformatted citations

Insert plain temporary citations in this exact family of forms:

```text
{Author, 2024 #123}
{Author, 2024 #123; Other, 2021 #456}
```

The record number is the decisive identifier. Keep the author and year aligned with the mapped record so a human can audit the placeholder. Count citation groups and unique record IDs before invoking Word.

### 4. Format through the installed CWYW add-in

Open the exact EndNote library explicitly, confirm it is not an empty `New Library`, and run the formatter on the isolated candidate:

```powershell
conda run -n base python -X utf8 <skill-dir>\scripts\format_cwyw.py `
  --input <pre-cwyw.docx> --output <cwyw-candidate.docx> `
  --style "Nature" --expect-main <citation-group-count> `
  --expect-references <unique-reference-count> `
  --expect-record-ids 123,456 --require-superscript `
  --report <format-report.json>
```

Run `--preflight-only` first when process state or placeholders are uncertain. The formatter refuses to overwrite outputs. It uses Word's `ENFormatBibliography`, optionally selects a style through EndNote's own Configure Bibliography dialog, waits for stable fields, and attempts exact-document recovery if an RPC error occurs after formatting.

### 5. Prove the result independently

Audit the saved package without relying on the automation report:

```powershell
conda run -n base python -X utf8 <skill-dir>\scripts\audit_cwyw_docx.py `
  --docx <cwyw-candidate.docx> --base <frozen-source.docx> `
  --expect-main <citation-group-count> --expect-references <unique-reference-count> `
  --expect-record-ids 123,456 --expect-style "Nature" `
  --require-superscript --report <audit-after.json>
```

Treat `EN.CITE.DATA` as an implementation detail: grouped citations use it, single-record citations may not. Do not require its count to equal the main citation count.

### 6. Release through explicit gates

Require all of the following before replacing any named target:

- ZIP CRC and XML parsing pass; internal relationships resolve.
- Expected main `EN.CITE` count, exactly one `EN.REFLIST`, and zero temporary citations remain.
- Every main citation contains full EndNote metadata in its field code or `w:fldData`.
- Field-derived record IDs and `EN.Libraries` record IDs match the expected set.
- The requested output style is present in `EN.Layout`; superscript checks pass when required.
- Source and protected backup hashes remain frozen.
- Package/media differences are within the declared edit boundary. If Word rewrites media, rebuild from the frozen base using only the explicitly required OOXML members and re-audit.
- The final candidate hash equals the delivered target hash.

## Failure policy

- If Word waits for a library, stop the attempt, close only processes created by this attempt, and reopen the exact `.enl` explicitly.
- If Configure Bibliography returns before the bibliography changes, run the official update again and wait for stable counts/style instead of saving immediately.
- If COM returns an RPC/re-entrancy error, first inspect or recover the exact running document. A failed call can still leave a fully formatted in-memory document.
- If hidden Word instances block initialization, report their exact process state and request approval before terminating only the verified automation instances.
- If a custom style is absent from the CWYW selector, use an available official style for field creation, then fix style registration. Do not confuse successful field creation with successful style selection.

## Additive cross-machine options

- **Automatic library launch (`--enl`):** pass a closed exact library. `format_cwyw.py` refuses to reuse an already-running EndNote with unverified library identity, opens the explicit path, checks the library title for the launched PID, and closes only that instance. It verifies/connects `EndNote.WordAddin.Connect` and retries transient `RPC_E_CALL_REJECTED` errors.
- Install the Python packages listed in the repository's `requirements.txt` into the interpreter used by the Windows desktop session. Keep both skill folders together so sibling imports resolve.
- **Active-record filtering (`trash_state`):** `scripts/inspect_endnote_snapshot.py` excludes trashed EndNote records (`trash_state != 0`) by default and exposes normalized `doi`, `abstract`, and `url` fields when present. Pass `--include-trashed` to include Trash records or `--all-records` to dump all active records.
- **Track-Changes & media safety:** `scripts/audit_cwyw_docx.py` excludes Word Track Changes deleted text (`w:delText`) from visible temporary-citation checks, checks `word/footnotes.xml` and `word/endnotes.xml`, ignores optional bibliography heading paragraphs (`References`, `Bibliography`, `参考文献`), and supports `--restore-media-out <repaired.docx>` (paired with `--base`) to restore any `word/media/*` assets rewritten by Word.


