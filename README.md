# Production Format Engine (feasibility lab)

An offline engine that edits Word files directly (Open XML) to set and fix proposal formatting standards,
and calls Word only for what Word alone knows: page numbers, field results, and the PDF. It is the
foundation for the production app that will sit beside Stationkeeper and replace the ProposalEngine macros.

Report: https://claude.ai/artifact/GxfYmABQQaetLV46dbhWVA

## Status (2026-10-28)
Feasible and reliable on everything tested, against real Word 16.0.17932:
- Full 8-step recipe on a 757-page, 140-chapter document: every setting read back by Word, 0 failures.
- Style merge matches Word's own CopyStylesFromTemplate on every property.
- 22 of 22 rogue edits refused, 26 of 26 puncture cases, 30 fuzzed files, 0 made worse.
- Live mode in open Word: same result as file mode; one Ctrl+Z restores.

Open: real volumes (FormatCheck), SharePoint co-authoring, security approval, Excel and PowerPoint.

## Layout
| Folder | What |
|---|---|
| `engine/` | Package layer, unit guard, fingerprint, safe write path, schema order |
| `engine/ops/` | Word operations, Quick Parts (glossary), style merge |
| `harness/` | Word-verified test suites: verify, puncture, rogue, m1_live, m5_merge, m6_open, glossary_test |
| `fieldcheck/` | FormatCheck read-only field kit and its PyInstaller spec |
| `fixtures/` | Synthetic proposal fixtures (generator, Word-written outputs, manifests). No real content. |
| `scripts/` | Chained run of every suite |

## Rules this code keeps
- Never write in place. Temp file, full re-check, dated snapshot, atomic replace.
- Never write a file that is open anywhere we can see. Any doubt means no write.
- Every operation marks what it changes. The unit guard refuses any other change to any paragraph,
  table property, style, list or Quick Part.
- Word judges layout, fields and the PDF.
- No network, no generative AI.
- Real proposal documents never go in this repository. They are checked only on the work PC with FormatCheck,
  whose report holds counts only.

## Setup (Windows, Word installed)
```
py -3.11 -m venv venv
venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Run the suites
```
venv\Scripts\python.exe harness\verify.py fixtures\small.docx all results\small_all.json
venv\Scripts\python.exe harness\verify.py fixtures\xl.docx recipe results\xl.json --sample 10 --pdf --manifest fixtures\manifest_xl.json
venv\Scripts\python.exe harness\rogue.py fixtures\small.docx results\rogue results\rogue.json
venv\Scripts\python.exe harness\puncture.py fixtures\small.docx results\puncture results\puncture.json
venv\Scripts\python.exe harness\m1_live.py fixtures\small.docx results\m1 results\m1.json 8
venv\Scripts\python.exe harness\m5_merge.py fixtures\small.docx results\m5 results\m5.json
```
Or all of them: `powershell -ExecutionPolicy Bypass -File scripts\run_all.ps1`

## Build FormatCheck
```
venv\Scripts\pyinstaller.exe --noconfirm --distpath dist --workpath build fieldcheck\FormatCheck.spec
```
Always build from the spec. The release zip is attached to the GitHub release, not committed.
