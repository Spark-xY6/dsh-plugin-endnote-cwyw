# Workflow and troubleshooting

## Run the authorized update

Use Windows PowerShell in the Python environment where you installed the dependencies. Run from the `cite-from-pdfs` skill directory, using absolute paths for inputs and fresh destinations. Substitute real paths for each angle-bracket placeholder.

```powershell
python -X utf8 scripts\index_pdf_folder.py index --pdf-dir <pdfs> --enl <old.enl> --work-dir <index-work> --output <catalog.json>
python -X utf8 scripts\index_pdf_folder.py search --catalog <catalog.json> --query "<claim>" --top-k 5
python -X utf8 scripts\preview_import.py --enl <old.enl> --docx <source.docx> --catalog <catalog.json> --plan <plan.json> --work-dir <preview-work> --output <private-review.json>
if ($LASTEXITCODE) { throw 'Read the local review report before preparing the library' }
$ready = python -X utf8 scripts\apply_citations.py --prepare-only --enl <old.enl> --docx <source.docx> --catalog <catalog.json> --plan <plan.json> --work-dir <work> --output-dir <new-delivery-directory> | ConvertFrom-Json
if ($LASTEXITCODE -or $ready.status -ne 'READY') { throw 'Preparation failed' }
$formatter = $ready.format_command[0]
$formatterArgs = @($ready.format_command | Select-Object -Skip 1)
& $formatter @formatterArgs
if ($LASTEXITCODE) { throw 'Native CWYW formatting failed' }
python -X utf8 scripts\apply_citations.py --publish-prepared $ready.work --ready-sha256 $ready.ready_sha256
```

Run the formatter directly in the interactive Windows desktop session. This avoids the Word VBA crash observed when the combined Python wrapper launched the formatter as a child. The agent runs all three steps when authorized. For a collection approved for public bibliographic lookups, append `--enrich-missing-metadata` to the index command.

```json
{
  "style": "Nature",
  "insertions": [
    {
      "insertion_id": "C01",
      "anchor_text": "The exact sentence in the Word file.",
      "position": "before_terminal_punct",
      "pdf_ids": ["P002"],
      "sources": [{"pdf_id": "P002", "page": 2, "quote": "Short passage copied from this physical PDF page"}],
      "rationale": "Explain how this passage supports the specific sentence."
    }
  ]
}
```

When a literal anchor repeats, specify a one-based `occurrence`. To append at the end of an anchor, use `after_anchor`. Preserve the source EndNote style by omitting `style` if the source defines one. For numeric superscript output, use `--require-superscript` when applicable. `space_before_citation` defaults to true for the prepared Word document. References selected for the same anchor go in one group.

Every newly placed PDF needs a `sources` entry. Page 1 is the first page of the file, including covers. Supply a short quotation from selectable page text. Preparation checks the actual PDF again, before starting an import. This check establishes where the words occur. The agent and reader judge whether those words support the claim. The review JSON contains bibliographic details; retain it with the private working files.

## Bibliographic identity

- DOI and explicitly identified PMID have priority. Conflicting identifiers for otherwise similar works or contradictory metadata go to review. DOI/PMID are compared with bibliographic details; catalog record IDs never authorize reuse.
- When identifiers are unavailable, exact normalized titles require matching first-author family and year. Close title variants additionally require a matching journal or matching volume and first page. Short titles, changes in year and weak variants need corroboration. Resolve a preprint versus the published article explicitly when identifiers conflict.
- Within the same library, prefer an equivalent record already cited by the document. Multiple proven equivalent IDs collapse to one for this document, without deleting entries or disrupting other documents pointing to them. The same rule deduplicates newly selected articles in one batch and avoids reinsertion on rerun against the updated library.
- For genuine foreign CWYW citations, use embedded Traveling Library metadata to find the equivalent work in the target. Rebind verified citations to target IDs through temporary citations and the official Word add-in. Missing existing journal articles can be imported from their complete embedded metadata. When the Word uses unsupported display overrides, notes, other citation systems or static numbering, qualify what is present and resolve or convert those placements before treating them as editable EndNote fields.

## Candidate release and failure

Save and close the input library before starting. During native import, keep other EndNote libraries closed. On the tested English EndNote 21 desktop, the importer opens **File → Import → File** on the verified candidate, observes the native progress control, and waits for EndNote to finish closing before opening the database snapshot. A second library, an unfamiliar dialog or an unrecognized interface language can stop this step. Keep the failure receipt, identify the window, and retry on a fresh candidate when the desktop is ready.

The full `.enl` and paired `.Data` (including groups and attachments) are copied into a new directory. The original remains the recovery source. Import fails an isolated candidate if the database closes uncleanly, if new IDs are ambiguous, or if its record delta differs from the plan. The wrapper also checks bibliographic identity and library binding after Word has saved. Keep the diagnostic work directory as evidence. A failed run does not publish the delivery directory. Correct the input or close an attempt-owned stuck dialog normally, then retry using a new destination. The retired metadata-only in-place rollback command can leave mismatched files and should never be used.

For fully verified library-only import, `scripts/sync_to_endnote.py` accepts the read-only `--enl`, a fresh `--candidate-dir`, `--catalog`, `--select-pdf-ids`, `--work-dir` and `--output`. Use `apply_citations.py` for the coordinated Word and library release.
