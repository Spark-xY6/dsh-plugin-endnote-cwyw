# Windows EndNote CWYW: Proven Path and Failure Notes

## Contents

1. Proven insertion mechanism
2. Field topology and validation
3. EndNote database evidence
4. Word and EndNote process handling
5. Output-style handling
6. Recovery and fallback boundaries

## 1. Proven insertion mechanism

The reliable route is Word automation plus the macro installed by EndNote's CWYW add-in:

1. Put one or more unformatted citations in the DOCX, for example:

   ```text
   {Example, 2024 #1; Sample, 2025 #2}
   ```

2. Open the exact EndNote library before formatting.
3. Open the candidate in a dedicated Word instance.
4. Invoke:

   ```python
   application.Run("ENFormatBibliography")
   ```

5. Pump COM messages and wait until temporary citations disappear and the expected EndNote fields and bibliography stabilize.
6. Save to a new DOCX path.

An isolated two-record test of this route produced three Word fields: one main `EN.CITE`, one nested `EN.CITE.DATA`, and one `EN.REFLIST`. The main result displayed `1,2`, while the bibliography contained both mapped records. This proves that the official formatter—not handcrafted OOXML—created the Traveling Library data.

The lower-level `EndNote.WordAddin.Connect.OnButtonPressed` callback can be invoked with a fabricated Office `IRibbonControl`, but this was less stable and more prone to hidden-window and add-in-template conflicts. Prefer `Application.Run("ENFormatBibliography")`.

## 2. Field topology and validation

Do not equate the number of top-level citation positions, nested data fields, and unique references.

- A single-record citation often stores its complete `<EndNote>...</EndNote>` XML directly in the `EN.CITE` instruction.
- A grouped citation can use a top-level `EN.CITE` with base64-encoded EndNote XML in `w:fldData`, plus a nested `EN.CITE.DATA` field.
- One reference can appear at multiple main citation positions.
- One `EN.REFLIST` should span the generated bibliography.

Validate each count separately. Grouped and single citations can produce different nested field counts.

Evidence of a genuine citation includes:

- a Word field instruction beginning `ADDIN EN.CITE`;
- `<EndNote>`, `<RecNum>`, and full `<record>` metadata in the instruction or decoded `w:fldData`;
- corresponding record IDs in the `EN.Libraries` document variable;
- an `ADDIN EN.REFLIST` field;
- zero remaining `{Author, Year #Record}` placeholders.

Visible superscript numbers or a formatted reference list alone are not proof.

## 3. EndNote database evidence

EndNote stores the readable record database at:

```text
<library>.Data\sdb\sdb.eni
```

The `.enl` file is not the record table. Close EndNote cleanly, copy `sdb.eni` to the work area, freeze its hash, and use SQLite in read-only immutable mode. Relevant tables on the proven EndNote 21 library were `refs` and `groups`.

The `groups.members` BLOB layout was:

- bytes 0-3: non-member header data;
- bytes 4-7: little-endian member count;
- bytes 8 onward: little-endian 32-bit record IDs.

Require `len(blob) == 8 + 4 * count` and unique IDs. EndNote libraries can use proprietary collations such as `ENCIN_zh_CN`; a full SQLite integrity check may be unavailable. Report that limitation and instead verify schema, full table readability, counts, unique IDs, group membership, and the snapshot hash.

## 4. Word and EndNote process handling

Known blocking states:

- Word processes launched as `/Automation -Embedding` can be invisible, absent from the Running Object Table, and still hold `EndNote Cwyw.dotm` or EndNote initialization locks.
- Starting a second Word instance can trigger a hidden “file in use” dialog for `EndNote Cwyw.dotm`.
- EndNote can open an empty `New Library`, leaving Word waiting for library selection.
- Large documents can continue updating after the style dialog closes.

Safe handling:

- Enumerate processes and windows before live formatting.
- Open the exact `.enl` explicitly and verify its title/state.
- Use a dedicated candidate and a dedicated Word instance.
- Close only the document and processes created by the current attempt.
- Never bulk-kill invisible Word processes without explicit approval.
- Keep the Word window visible during fragile automation so prompts are discoverable.

## 5. Output-style handling

Use `ENFormatBibliographyDlg` to open EndNote's Configure Bibliography dialog. On the proven Windows build, the style combo had control ID `5049` and the OK button had ID `1`. Select the exact item by its text, send selection-change notifications, click OK, then invoke `ENFormatBibliography` again.

Do not save as soon as the dialog returns. Wait until:

- the requested style appears in the `EN.Layout` document variable;
- field and bibliography counts are stable for multiple polling cycles;
- temporary citations are zero;
- style-specific display markers have appeared when applicable.

A style file existing on disk does not prove it is registered in the CWYW selector. If a custom style is unavailable, distinguish these states:

- genuine field creation succeeded;
- requested style selection failed;
- display correction remains pending.

## 6. Recovery and fallback boundaries

An RPC or COM re-entrancy exception may occur after EndNote has already formatted the document. Before rerunning:

1. Use the running Word application and verify both the exact candidate path and the attempt-owned Word process ID before recovery.
2. Inspect main citations, nested fields, reference list, bibliography count, placeholders, and superscript state.
3. If every invariant passes, use `SaveAs2` to a new candidate and audit the saved package.
4. If invariants fail, close only the attempt-owned document without saving and diagnose the prompt/process state.

Package-level OOXML edits are a last-mile fallback only. They can preserve genuine fields while correcting displayed bibliography text or an `EN.Layout` style name, but they do not update the EndNote database or rebuild Traveling Library metadata. Never describe package-only field fabrication as “direct EndNote insertion.”

## 7. Additive robustness notes (trash filtering, Track Changes, media restore)

- **`inspect_endnote_snapshot.py`:** filters `COALESCE(trash_state, 0) = 0` by default (`--include-trashed` overrides) and selects optional `electronic_resource_number` (normalized as `doi`), `abstract`, and `url` columns when present.
- **`format_cwyw.py`:** accepts `--enl <target.enl>` to launch and cleanly close an attempt-owned EndNote window, skips opening `ENFormatBibliographyDlg` if `EN.Layout` already holds the requested style, falls back to child `ComboBox` discovery if control ID `5049` differs across EndNote versions, and records unexpected `#32770` modal dialogs that stay visible for over 3 seconds.
- **`audit_cwyw_docx.py`:** ignores `w:delText` deletions when scanning temporary citations, inspects footnotes/endnotes, excludes bibliography heading paragraphs from entry counts, and supports `--restore-media-out <repaired.docx>` with `--base` to restore original `word/media/*` bytes if Word SaveAs2 alters embedded images.


