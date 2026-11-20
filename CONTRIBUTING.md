# Contributing

Small, reproducible cases help most. Good first contributions include a citation layout from a different Word language, a false duplicate, or an EndNote style that produces several paragraphs per reference. Review existing issues before opening one.

1. Fork the repository and create a topic branch. Keep both skill directories adjacent.
2. Build a minimal synthetic example. For metadata, use invented records or a published DOI and your own example sentence. Describe the Word, EndNote, DSH and runtime versions, observed result, and expected result.
3. Run the checks below from the repository root on Windows. For changes to native import or formatting, also follow the licensed desktop recipe and report the citation and bibliography counts after each edit. Linux/macOS contributors can run the DSH package tests with Node.
4. Open a pull request with the regression case, fix and test results. State which desktop versions were actually exercised. Changes that relax record matching should include both a true duplicate and a similar-looking distinct publication.

```powershell
npm ci --ignore-scripts
npm test
python -m pip install -r requirements.txt
python -B -m unittest discover -s tests -p "test_*.py" -v
```

Work with a Python environment you control. Never attach libraries, manuscripts, full-text collections, screenshots with account details, or unredacted catalog/preview/audit JSON to public issues. An EndNote Traveling Library lives inside a DOCX and can disclose records even when visible text looks anonymous. An aggregate count or a short invented example usually reproduces the problem.

Code contributions use the project MIT license. Include original source and license notices for any new third-party code. See [desktop evaluation](docs/evaluation.md).
