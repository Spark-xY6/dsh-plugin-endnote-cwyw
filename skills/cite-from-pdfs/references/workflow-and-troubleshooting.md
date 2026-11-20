# Workflow and troubleshooting

## Run the authorized update

Use Windows PowerShell and Conda `base`. These paths are illustrative.

```powershell
conda run -n base python -X utf8 scripts\index_pdf_folder.py index --pdf-dir <pdfs> --enl <old.enl> --work-dir <index-work> --output <catalog.json>
conda run -n base python -X utf8 scripts\index_pdf_folder.py search --catalog <catalog.json> --query "<claim>" --top-k 5
$ready = conda run -n base python -X utf8 scripts\apply_citations.py --prepare-only --enl <old.enl> --docx <source.docx> --catalog <catalog.json> --plan <plan.json> --work-dir <work> --output-dir <new-delivery-directory> | ConvertFrom-Json
if ($LASTEXITCODE -or $ready.status -ne 'READY') { throw 'Preparation failed' }
$formatter = $ready.format_command[0]
$formatterArgs = @($ready.format_command | Select-Object -Skip 1)
& $formatter @formatterArgs
if ($LASTEXITCODE) { throw 'Native CWYW formatting failed' }
conda run -n base python -X utf8 scripts\apply_citations.py --publish-prepared $ready.work --ready-sha256 $ready.ready_sha256
```

Run the formatter directly in the interactive Windows desktop session. This avoids the Word VBA crash observed when the combined Python wrapper launched the formatter as a child. The agent runs all three steps when authorized. Use the verified `base` interpreter directly if the Conda launcher is unavailable.

```json
{
  "style": "Nature",
  "insertions": [
    {
      "insertion_id": "C01",
      "anchor_text": "The exact sentence in the Word file.",
      "position": "before_terminal_punct",
      "pdf_ids": ["P002"],
      "rationale": "PDF p. 2 directly supports the statement."
    }
  ]
}
```

When a literal anchor repeats, specify a one-based `occurrence`. To append at the end of an anchor, use `after_anchor`. Preserve the source EndNote style by omitting `style` if the source defines one. For numeric superscript output, use `--require-superscript` when applicable. `space_before_citation` defaults to true for the prepared Word document. References selected for the same anchor go in one group.

## Bibliographic identity

- DOI and explicitly identified PMID have priority. Conflicting identifiers for otherwise similar works or contradictory metadata go to review. DOI/PMID are compared with bibliographic details; catalog record IDs never authorize reuse.
- When identifiers are unavailable, exact normalized titles require matching first-author family and year. Close title variants additionally require a matching journal or matching volume and first page. Short titles, changes in year and weak variants need corroboration. Resolve a preprint versus the published article explicitly when identifiers conflict.
- Within the same library, prefer an equivalent record already cited by the document. Multiple proven equivalent IDs collapse to one for this document, without deleting entries or disrupting other documents pointing to them. The same rule deduplicates newly selected articles in one batch and avoids reinsertion on rerun against the updated library.
- For genuine foreign CWYW citations, use embedded Traveling Library metadata to find the equivalent work in the target. Rebind verified citations to target IDs through temporary citations and the official Word add-in. Missing existing journal articles can be imported from their complete embedded metadata. When the Word uses unsupported display overrides, notes, other citation systems or static numbering, qualify what is present and resolve or convert those placements before treating them as editable EndNote fields.

## Candidate release and failure

The full `.enl` and paired `.Data` (including groups and attachments) are copied into a new directory. The original remains the recovery source. Import fails an isolated candidate if the database closes uncleanly, if new IDs are ambiguous, or if its record delta differs from the plan. The wrapper also checks bibliographic identity and library binding after Word has saved. Keep the diagnostic work directory as evidence. A failed run does not publish the delivery directory. Correct the input or close an attempt-owned stuck dialog normally, then retry using a new destination. The retired metadata-only in-place rollback command can leave mismatched files and should never be used.

For fully verified library-only import, `scripts/sync_to_endnote.py` accepts the read-only `--enl`, a fresh `--candidate-dir`, `--catalog`, `--select-pdf-ids`, `--work-dir` and `--output`. Use `apply_citations.py` for the coordinated Word and library release.
