# Changelog

[Română](CHANGELOG.ro.md) · **English**

## 0.4.0 — English interface
- The interface, messages, CSV report, in-document placeholder (`[REDACTED:TYPE]`), file and folder names
  (`_REDACTED`, `redacted_…`, `report_…`, `Converted\`) and the self-test are available in Romanian and English.
- "Language / Limba" selector in the window (rebuilds the window, keeps files and options); the choice is saved in
  `%LOCALAPPDATA%\pii_contracte\settings.json`; defaults to the Windows display language; `--lang ro|en`.
- Single text catalogue (`pii_core/i18n.py`); internal type identifiers stay in Romanian, so the logic is unchanged.
- Documentation: `README.md` in English (main), `README.ro.md`, `SECURITY` and the user guide in both languages.
- The Word self-test uses the DOCX copy actually produced (it previously had a hard-coded `test_REDACTAT.docx` path).
- Tests: complete, consistent catalogue (same parameters in both languages, no Romanian diacritics in English),
  scan/redact/report/self-test in English, language switching in the interface.

## 0.3.2 — ready for publication
- Copyright and licence notice (AGPL-3.0-or-later, FSF recommended wording) in every source file.
- README: explicit warning that redaction is not guaranteed and must be checked by a person; reporting section.
- `SECURITY.md`: how to report a redaction leak without exposing real data.

## 0.3.1 — repository
- Restructured README, changelog, Windows test checklist, LICENSE (AGPL-3.0), `.gitignore` that blocks documents.
- pytest pinned to 9.1.1; LICENSE included in the .exe package.

## 0.3.0 — .exe package
- PyInstaller 6.22.3 "onedir" build, no console, no UPX (`pii_contracte.spec`, `build.ps1`).
- Log file in `%LOCALAPPDATA%\pii_contracte\pii_contracte.log`; interface errors shown; `--selftest`
  (PDF, DOCX, Tcl/Tk, Word) on synthetic data.
- CSV report: dates fully masked; CNP/IBAN/ID/phone/licence/land-register numbers show only the last 2 characters.
- Registry: the `HKCU\Software\pii_contracte` key is deleted when it becomes empty.
- `tasklist`/`taskkill` get an explicit stdin (standard handles may be missing in a console-less .exe).
- Verified on Windows 11: the self-test from the .exe reports "ALL OK", including Word via pywin32.

## 0.2.6 — registry restore
- After 0.2.5, `DisableConvertPdfWarning` stayed at 1 on Windows. Probable cause (not validated): Word writes its
  options on exit. Now the program waits for the Word process to exit (max. 15 s, then taskkill), restores the value
  and verifies it by reading it back.
- The original value is recorded in `HKCU\Software\pii_contracte\RegistryRestore` and restored on the next start if
  the program was killed. Verified on Windows 11.

## 0.2.5 — PDF → DOCX export hang
- Cause: Word's "Word will now convert your PDF…" prompt, not suppressed by `DisplayAlerts`; with Word invisible it
  waits forever. Suppressed temporarily via `DisableConvertPdfWarning = 1` (source: Blue Prism community, not
  Microsoft documentation).
- Watchdog: a Word call longer than 120 s stops only the program's own instance; Word restarts for the next file.

## 0.2.4 — export format
- "Redacted copies as: original format / PDF / DOCX". Redaction happens in the source format; Word exports the
  already redacted copy; both are kept; the export is verified (`engine.verify_file`).

## 0.2.3 — new detectors (after the first real contract)
- Ticked by default: driving licence number, land register / cadastral number, property address.
- Unticked by default (indirect data, not propagated): place of birth, document issuer, document dates.

## 0.2.2 — robust tests on Windows
- Tests normalise extracted text (NBSP with Arial); U+2011 test skipped when the font lacks the glyph; the GUI smoke
  test fails (instead of skipping) if Tk does not start on Windows.
- Diagnosis: Tk does not start from a venv with CPython 3.13.0 (`sys.prefix` instead of `sys.base_prefix`, fixed in
  3.13.1).

## 0.2.1 — hyphens
- PDF: U+00AD/U+2010/U+2011 → `-`, U+00A0/U+202F → space (1:1 normalisation). With Arial the hyphen is extracted as
  U+00AD, so hyphenated names were not detected.
- DOCX: `<w:noBreakHyphen/>` treated as `-`, `<w:softHyphen/>` as invisible.

## 0.2.0 — user interface
- Tkinter interface: files, background scan, tick-box table (per occurrence and per value), file filter, manual
  terms with rescan (unticks are kept), warnings, generation + report.
- `.doc` via Word COM: separate instance, macros disabled, read-only, no dialogs.

## 0.1.0 — core
- Detection and redaction for DOCX (lxml, directly on the XML) and digital PDF (PyMuPDF, true redaction).
- Post-redaction verification, masked CSV report, XXE / zip-bomb protections.
