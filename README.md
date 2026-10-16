# pii_contracte

[Română](README.ro.md) · **English**

A Windows application that **detects and redacts personal identification data** in contracts
(**.docx**, **digital .pdf**, **.doc** via Microsoft Word). All processing is local; no network access.

Workflow: pick the files → **Scan** → review the findings in a table with tick boxes → **Generate** redacted copies
(`<name>_REDACTED.<ext>`) and a CSV report with masked values. Originals are never modified.

The interface is available in **English and Romanian** ("Language / Limba" selector in the window; defaults to the
Windows display language; the choice is remembered). Detection is built for **Romanian documents** (CNP personal
numeric code, Romanian ID cards, RO IBANs, wording used in Romanian contracts), whatever the interface language.

> ⚠️ **No warranty.** Automatic redaction is **not guaranteed to be complete**: the program can miss data (names
> without an anchor phrase, text inside images, unexpected wording) or flag ordinary text. Every redacted copy must
> be **checked by a person** before it is shared. Whoever uses the program remains responsible for processing the
> personal data (controller / processor under the GDPR). See also sections 15–16 of the [LICENSE](LICENSE).

> Status: version 0.4.0, used internally. Tested on Windows 11 with Python 3.14 and Microsoft Word 16.0.
> Scanned pages (images) are **not** analysed yet (OCR is planned). History: [CHANGELOG.md](CHANGELOG.md).

## What it detects

| Type | How | Ticked by default |
|---|---|---|
| CNP (Romanian personal numeric code) | check digit + date + county code (failed check next to "CNP" → low confidence) | yes |
| IBAN | mod-97 | yes |
| Email, Romanian phone numbers | pattern | yes |
| ID card series/number, passport number | anchor ("seria … nr.", "pașaport nr.") | yes |
| Driving licence number | anchor ("permis de conducere … nr.") | yes |
| Date of birth | anchor ("născut(ă) la …") | yes |
| Name | anchor ("subsemnatul", "domnul/doamna", "reprezentată prin", "…, domiciliat/CNP") | yes |
| Home address | after "domiciliat în / cu domiciliul / reședința" | yes |
| Property address | after "situat/amplasat/aflat în", "imobilul/apartamentul din" | yes |
| Land register / cadastral number | "Cartea Funciară … nr.", "CF nr.", "număr cadastral" | yes |
| Place of birth, document issuer, document dates | anchors ("născut … în", "eliberat de", "valabil până la") | **no** (indirect data) |
| Any value found once | searched across the whole document (e.g. the name in the signature block) | as the original |
| Manual text | typed in the interface | yes |

**Not detected:** text in images and scanned pages, PDF form fields, names without an anchor phrase, the bank name,
marital status. Use the "Text to redact manually" field for these. **Always check the redacted copy before sharing it.**

## What the redacted copy contains
- **PDF:** the text is removed from the page content (not just covered); links containing redacted values and the
  metadata are removed.
- **DOCX:** values are replaced with `[REDACTED:TYPE]` (`[REDACTAT:TIP]` in the Romanian interface); tracked changes
  are accepted (deleted text disappears); comments and metadata (author, company) are removed.
- **Automatic verification:** the copy (and any export to another format) is read back; if a ticked value is still
  present, the interface shows "VERIFICATION FOUND PROBLEMS".
- **Export:** optionally to PDF or DOCX via Word (the copy in the original format is kept as well). PDF → DOCX breaks
  the layout.
- **.doc:** converted with Word into a `Converted\` folder next to the original (an **unredacted** copy that stays
  there), then redacted.

## Using the .exe
Copy the whole `pii_contracte\` folder from the build and start `pii_contracte.exe`. Short user guide:
[README-USER.txt](README-USER.txt) ([CITESTE-MA.txt](CITESTE-MA.txt) in Romanian).
Self-test: `pii_contracte.exe --selftest`. Force a language: `pii_contracte.exe --lang en` (or `ro`).

The .exe is not code-signed, so Windows SmartScreen asks for confirmation the first time a downloaded copy is run
(More info → Run anyway). To avoid the prompt for a copy you trust, unblock the zip before extracting it:
`Unblock-File .\pii_contracte_<version>.zip`.

## Development

Requirements: Windows 10/11, **Python 3.14** (from python.org, with "tcl/tk and IDLE"), Microsoft Word for .doc
files and exports. Do not use Python 3.13.0: Tk does not start from a venv (CPython bug, fixed in 3.13.1).

```powershell
git clone <repo-url> pii_contracte
cd pii_contracte
py -3.14 -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass   # current window only
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
python -m pytest -q -rs tests        # on Windows: 2 skipped (U+2011 has no glyph in Arial; Linux locale test)
python -m pii_gui                    # run the interface from source
```

Build the `.exe` (PyInstaller 6.22.3, onedir, no console):
```powershell
.\build.ps1
```
The script creates a separate venv (`.venv-build`), runs the tests, builds `dist\pii_contracte\`, runs the self-test
from the `.exe` and creates `dist\pii_contracte_<version>.zip` with its SHA256.

Manual tests on Windows (Word, registry, .exe on an account without Python):
[docs/TESTING-WINDOWS.md](docs/TESTING-WINDOWS.md).

### Layout
```
pii_core/          core, no user interface
  detectors.py     detectors (regex + validation) for Romanian text
  engine.py        scan_file / redact_file / verify / write_report (masked report)
  docx_handler.py  DOCX directly on the XML (lxml, no XXE): body, tables, headers, notes, tracked changes
  pdf_handler.py   PDF with PyMuPDF: text with coordinates, true redaction
  doc_convert.py   Microsoft Word via COM (pywin32): .doc → docx/pdf, export, watchdog, registry
  i18n.py          user-visible texts in RO/EN (catalogue); internal type identifiers stay in Romanian
  models.py
pii_gui/           Tkinter interface
  app.py           main window; model.py = tick-box state; jobs.py = background operations
  runtime.py       file logging, settings (language), --selftest, --lang
tests/             pytest; 100% synthetic documents (fixtures.py)
pii_contracte_app.py   .exe entry point
pii_contracte.spec, build.ps1, requirements-build.txt   build
demo_cli.py, diag_iban.py   command-line utilities (they print full values to the console!)
```

### Dependencies (pinned in `requirements*.txt`)
| Package | Version | Purpose | Licence |
|---|---|---|---|
| pymupdf | 1.28.2 | PDF | AGPL-3.0 / commercial |
| lxml | 6.1.3 | DOCX | BSD-3-Clause |
| pywin32 | 312 (Windows only) | Word COM | PSF |
| pytest, python-docx | 9.1.1, 1.2.0 | tests only | MIT |
| pyinstaller | 6.22.3 | build only | GPLv2+ with bootloader exception |

## Security and privacy
- No network access. Originals are only read; output can never overwrite an original.
- Full values are shown only in the application window. The CSV report masks them (dates fully, numbers down to the
  last 2 characters). The log (`%LOCALAPPDATA%\pii_contracte\`) contains no values but may contain file names.
- XML is parsed without external entities or network access; zip-bomb and size limits apply.
- Word: a separate, invisible instance with **macros disabled** (`AutomationSecurity = ForceDisable`), documents
  opened read-only, no dialogs. Word's PDF conversion prompt is suppressed temporarily via
  `HKCU\Software\Microsoft\Office\<ver>\Word\Options\DisableConvertPdfWarning`; the original value is restored after
  Word has exited (and on the next start if the program was killed). A Word call blocked for more than 120 s stops
  only the program's own Word instance.
- **Never add real contracts, reports or redacted copies to the repository.** `.gitignore` blocks
  `.doc/.docx/.pdf/.csv`; check `git status` before every commit.

## Known limitations
- No OCR: scanned pages and images (e.g. an ID card pasted into a document) only produce a warning.
- Names without an anchor phrase, reversed names (Ion Popescu vs POPESCU ION), different spellings of one name.
- Files in `Converted\` (`Conversie\` in Romanian) are complete, unredacted copies.
- The `.exe` is not code-signed (see above).

## Reporting problems
A value that is not redacted, or a redaction error: see [SECURITY.md](SECURITY.md).
**Never attach real documents** to issues; reproduce the problem with made-up data.

## Licence
Copyright (C) 2026 CatalinFusea.

[GNU AGPL-3.0-or-later](LICENSE), chosen because of PyMuPDF (AGPL-3.0). Internal use: no additional obligations.
If the program (or the `.exe`) is given to someone else, the complete source code must be made available to them
under the AGPL-3.0. Third-party licences: [LICENSES.txt](LICENSES.txt).
