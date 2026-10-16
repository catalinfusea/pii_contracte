# Manual tests on Windows

[Română](TESTARE-WINDOWS.md) · **English**

The automated tests (`python -m pytest -q -rs tests`) simulate Word and the registry. The real behaviour of Word,
the registry and the `.exe` is checked manually with the list below. Use **synthetic documents only**.

Status: ✅ verified · ⬜ to verify · last updated: version 0.4.0.

## The application (from source or .exe)
| # | Test | Expected | Status |
|---|---|---|---|
| 1 | `python -m pytest -q -rs tests` | all pass, 2 skipped (U+2011 in Arial; Linux locale test) | ⬜ 0.4.0 |
| 2 | .docx + .pdf: scan, generate | copies + report; no "VERIFICATION FOUND PROBLEMS" | ✅ |
| 3 | Redacted DOCX copy opened in Word | no "unreadable content / repair" prompt | ✅ opens (absence of the repair prompt not explicitly confirmed) |
| 4 | PDF copy: select the text under the black boxes; Ctrl+F for a CNP | nothing selectable / found | ⬜ |
| 5 | CSV report in Excel | correct diacritics, separate columns, masked values | ⬜ |
| 6 | .doc → DOCX and .doc → PDF | file in `Converted\`; original unchanged (`certutil -hashfile x.doc SHA256` before/after) | ⬜ |
| 7 | .doc with an `AutoOpen` macro showing a MsgBox | no MsgBox; conversion succeeds | ⬜ |
| 8 | Password-protected .doc | error in the list, no password prompt, no hang | ⬜ |
| 9 | .txt renamed to .doc / corrupted .doc | error or conversion, no dialog, no hang | ⬜ |
| 10 | Same .doc already open in Word | conversion works; the user's Word is unaffected | ⬜ |
| 11 | After any Word operation: Task Manager | no leftover WINWORD.EXE | ✅ |
| 12 | .doc downloaded from email/browser (Mark-of-the-Web) | conversion or a clear error | ⬜ |
| 13 | Read-only original folder / network share | "Cannot create the Converted folder" error | ⬜ |
| 14 | Names/paths with diacritics and spaces | works | ⬜ |
| 15 | Untick one occurrence of a name, generate | no verification problem; the unticked occurrence remains | ⬜ |
| 16 | Close the window during a scan with .doc files | closes after the current file; Word does not stay running | ⬜ |
| 17 | Windows scaling 125% / 150% | sharp text, table rows do not overlap | ⬜ |
| 19 | PDF export of a .docx | `_REDACTED.docx` + `_REDACTED.pdf`; no verification problems | ✅ |
| 20 | DOCX export of a .pdf | both copies; no hang (DOCX layout is approximate) | ✅ 0.2.5 |
| 22 | After 19/20: `reg query "HKCU\Software\Microsoft\Office\16.0\Word\Options" /v DisableConvertPdfWarning` and `reg query "HKCU\Software\pii_contracte"` | "unable to find" for both | ✅ 0.2.6 (empty key remained; deleted since 0.3.0) |

## The .exe package
| # | Test | Expected | Status |
|---|---|---|---|
| B1 | `.\build.ps1` | tests pass, build succeeds, self-test "ALL OK" including Word | ✅ 0.3.0 |
| B2 | `dist\pii_contracte` copied to an account/PC **without Python**; double-click | the window opens (after SmartScreen) | ⬜ |
| B3 | On B2: `pii_contracte.exe --selftest` | "ALL OK" (Word only if Office is installed) | ⬜ |
| B4 | On B2: .docx and .pdf, PDF and DOCX export | as above; no leftover WINWORD.EXE | ⬜ |
| B8 | After B4: the registry checks from #22 | "unable to find" for both | ⬜ |
| B9 | Defender when copying/running | no quarantine; otherwise note the detection name | ⬜ |
| B10 | CSV report | dates as `**.**.****`, numbers only with the last 2 characters | ⬜ |

## Interface language (0.4.0)
| # | Test | Expected | Status |
|---|---|---|---|
| L1 | First start, Windows in Romanian / in English | interface in the Windows language | ⬜ |
| L2 | "Language / Limba" → Română, with files and terms added | the window is rebuilt in Romanian, files and terms remain | ⬜ |
| L3 | Close and restart | same language as when closed | ⬜ |
| L4 | In English: scan + generate on .docx and .pdf | `_REDACTED`, `[REDACTED:…]`, `report_…csv` with an English header | ⬜ |
| L5 | `pii_contracte.exe --lang ro --selftest` / `--lang en --selftest` | self-test window in the requested language, "TOTUL OK" / "ALL OK" | ⬜ |
