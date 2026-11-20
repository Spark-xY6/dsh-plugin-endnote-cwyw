# PDF citations for EndNote and Word

[中文使用说明](README.zh.md) · [Desktop evaluation](docs/evaluation.md) · [Related projects](docs/landscape.md)

A DeepSeek Harness plugin for a familiar handoff: your collaborator sends a partly cited Word draft, you have an older EndNote library and a folder of PDFs, and both of you want to keep using **Edit & Manage Citation(s)** in Word. The plugin delivers an updated library copy and a DOCX containing live EndNote citations. One work can appear at several places in the text while retaining a single bibliography entry.

## Why build this?

A record number belongs to one EndNote library. The number from someone else's draft can point to another paper in yours. Importing a slightly different title for a paper you already cite can give it a second number. The library may also contain several plausible papers for a claim. This workflow treats these as separate decisions: find supporting words on a PDF page, identify the paper in the target library, then ask the official EndNote Word add-in to format the citation.

PubMate's read-only preview and per-identifier reports informed our import preview. JabRef and Zotero inspired showing the fields that support or contradict a proposed match. The public evaluation exercises what happens after delivery when a contributor removes a citation and inserts it again. [Comparison, sources and adopted ideas](docs/landscape.md)

## Install

Use an interactive Windows desktop with Word, EndNote and its working Cite While You Write (CWYW) add-in. The tested host is DSH `0.2.2-rc.5` with Node 24.13.0 and EndNote 21. Run DSH on the same desktop, with its agent allowed to read the three inputs and execute local PowerShell. [DSH plugin installation](https://deepseek-harness.github.io/deepseek-harness/en/develop/basic/publish)

```powershell
git clone --branch v0.1.0 https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw.git
cd dsh-plugin-endnote-cwyw
# Use a Python 3.10+ environment of your choice for the agent and this command.
python -m pip install -r requirements.txt
dsh plugin --profile web add github:Spark-xY6/dsh-plugin-endnote-cwyw#v0.1.0
```

Restart the profile after installation. The bundle registers **cite-from-pdfs** and **endnote-cwyw**. The repository also works with filesystem-skill hosts that can run local commands: copy both directories under `skills/` together. For a specific installation or contributor test, use the Git commit shown on the release. [Releases](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/releases)

## Use it with a draft

1. Save the draft and close the source EndNote library cleanly. Keep the `.enl` next to its matching `.Data` directory. Choose a new output folder outside the library. Keep all inputs and work files outside the Git checkout.
2. Give the agent local paths. Example: **“Use cite-from-pdfs with my PDFs folder, Library.enl and Draft.docx. Add citations only where a PDF page supports my sentence. Reuse existing records. Work on copies, deliver the updated Word and library together, and list claims or matches that need review.”**
3. The agent builds the plan, reads the supporting pages, previews import/reuse decisions, then prepares, formats and audits the result. To inspect decisions first, ask **“Preview only.”** In that mode the agent produces a private JSON review with the proposed record numbers and field differences.
4. Open the delivered library and then `manuscript_cwyw.docx`. Use the EndNote tab to insert, edit or remove citations. Keep the updated `.enl` **and** its `.Data` directory together when sending the pair to your collaborator. The local `audit.json` records the checks for that pair. [CLI commands and plan format](skills/cite-from-pdfs/references/workflow-and-troubleshooting.md)

## Evidence and limits

On the supplied EndNote 21 sample library, selecting two filenames for the same public paper added **one** record (190 → 191). Reopening the delivery, removing that paper in Word and reinserting it changed the saved bibliography counts **2 → 1 → 2**. Running again with the updated library added **zero** records (191 → 191).

GitHub Actions checks package registration and regression cases using synthetic inputs. The desktop evaluation runs the licensed applications with a fixed evidence plan, then tests the saved documents before and after edits. It also checks what happens when the same article is selected again. Model-driven selection of papers from a previously unseen collection needs separate evaluation. [Desktop results](docs/evaluation.md) · [GitHub Actions](https://github.com/Spark-xY6/dsh-plugin-endnote-cwyw/actions)

The automatic route is for genuine EndNote fields in the document body. Hand-typed numbers, Zotero or Mendeley fields, citations inside notes, and custom citation display overrides require separate review. The agent judges whether a passage supports a claim. The software verifies the quoted words on the specified physical PDF page. It indexes searchable text from up to the first twelve pages; the agent can read later pages directly. Image-only scans need readable text for quote verification. The project has no cross-corpus accuracy estimate.

## Privacy, licensing and contributions

The Python tools use local files. Online metadata enrichment is optional and can send a PDF-derived title, DOI or arXiv ID to Crossref, OpenAlex or PubMed. A DSH session may send manuscript text or tool output to its configured model provider. Choose a provider and permissions suitable for your documents. Raw reports can contain paths, titles and manuscript text.

The plugin source is MIT-licensed. Microsoft Word and Clarivate EndNote are separate products. Python dependencies retain their own licenses, including PyMuPDF's AGPL/commercial options. To contribute, start with a synthetic case or public DOI and the observed result. [License](LICENSE) · [PyMuPDF terms](https://pymupdf.readthedocs.io/en/latest/about.html#license-and-copyright) · [Contributor guide](CONTRIBUTING.md)
