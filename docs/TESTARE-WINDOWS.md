# Teste manuale pe Windows

**Română** · [English](TESTING-WINDOWS.md)

Testele automate (`python -m pytest -q -rs tests`) simulează Word și registry-ul. Comportamentul real al Word,
al registry-ului și al `.exe`-ului se verifică manual, cu lista de mai jos. Folosiți **documente sintetice**.

Stare: ✅ verificat · ⬜ de verificat · ultima actualizare: versiunea 0.4.0.

## Aplicația (din sursă sau .exe)
| # | Test | Așteptat | Stare |
|---|---|---|---|
| 1 | `python -m pytest -q -rs tests` | toate trec, 2 sărite (U+2011 în Arial; test de locale Linux) | ⬜ 0.4.0 |
| 2 | .docx + .pdf: scanare, generare | copii + raport; fără „VERIFICAREA A GĂSIT PROBLEME” | ✅ |
| 3 | Copia DOCX redactată deschisă în Word | fără „conținut ilizibil / reparare” | ✅ se deschide (absența dialogului de reparare neconfirmată explicit) |
| 4 | Copia PDF: selectați textul din zonele negre; Ctrl+F după un CNP | nimic selectabil / găsit | ⬜ |
| 5 | Raportul CSV în Excel | diacritice corecte, coloane separate, valori mascate | ⬜ |
| 6 | .doc → DOCX și .doc → PDF | fișier în `Conversie\`; originalul neschimbat (`certutil -hashfile x.doc SHA256` înainte/după) | ⬜ |
| 7 | .doc cu macro `AutoOpen` care afișează un MsgBox | niciun MsgBox; conversia reușește | ⬜ |
| 8 | .doc protejat cu parolă | eroare în listă, fără cerere de parolă, fără blocaj | ⬜ |
| 9 | .txt redenumit .doc / .doc corupt | eroare sau conversie, fără dialog, fără blocaj | ⬜ |
| 10 | Același .doc deschis deja în Word | conversia merge; Word-ul utilizatorului neafectat | ⬜ |
| 11 | După orice operație cu Word: Task Manager | niciun WINWORD.EXE rămas | ✅ |
| 12 | .doc descărcat din e-mail/browser (Mark-of-the-Web) | conversie sau eroare clară | ⬜ |
| 13 | Folder original read-only / share de rețea | eroare „Nu pot crea folderul Conversie” | ⬜ |
| 14 | Nume/căi cu diacritice și spații | funcționează | ⬜ |
| 15 | Debifați o apariție a unui nume, generați | fără problemă la verificare; apariția debifată rămâne | ⬜ |
| 16 | Închideți fereastra în timpul unei scanări cu .doc | se închide după fișierul curent; Word nu rămâne pornit | ⬜ |
| 17 | Scalare Windows 125% / 150% | text clar, rândurile tabelului nu se suprapun | ⬜ |
| 19 | Export PDF pe un .docx | `_REDACTAT.docx` + `_REDACTAT.pdf`; fără probleme la verificare | ✅ |
| 20 | Export DOCX pe un .pdf | ambele copii; fără blocaj (aspectul DOCX e aproximativ) | ✅ 0.2.5 |
| 22 | După 19/20: `reg query "HKCU\Software\Microsoft\Office\16.0\Word\Options" /v DisableConvertPdfWarning` și `reg query "HKCU\Software\pii_contracte"` | „unable to find” la ambele | ✅ 0.2.6 (cheia goală rămânea; ștearsă din 0.3.0) |

## Pachetul .exe
| # | Test | Așteptat | Stare |
|---|---|---|---|
| B1 | `.\build.ps1` | teste trecute, build fără erori, autotest „TOTUL OK” inclusiv Word | ✅ 0.3.0 |
| B2 | `dist\pii_contracte` copiat pe un cont/PC **fără Python**; dublu-click | fereastra se deschide (după SmartScreen) | ⬜ |
| B3 | Pe B2: `pii_contracte.exe --selftest` | „TOTUL OK” (Word doar dacă e instalat Office) | ⬜ |
| B4 | Pe B2: .docx și .pdf, export PDF și DOCX | ca mai sus; niciun WINWORD.EXE rămas | ⬜ |
| B8 | După B4: verificările de registry de la #22 | „unable to find” la ambele | ⬜ |
| B9 | Defender la copiere/rulare | fără carantină; altfel notați numele detecției | ⬜ |
| B10 | Raportul CSV | date `**.**.****`, numere doar cu ultimele 2 caractere | ⬜ |

## Limba interfeței (0.4.0)
| # | Test | Așteptat | Stare |
|---|---|---|---|
| L1 | Prima pornire, Windows în română / în engleză | interfața în limba Windows | ⬜ |
| L2 | „Limba / Language” → English, cu fișiere și termeni adăugați | fereastra se reface în engleză, fișierele și termenii rămân | ⬜ |
| L3 | Închide și repornește | aceeași limbă ca la închidere | ⬜ |
| L4 | În engleză: scanare + generare pe .docx și .pdf | `_REDACTED`, `[REDACTED:…]`, raport `report_…csv` cu antet în engleză | ⬜ |
| L5 | `pii_contracte.exe --lang ro --selftest` / `--lang en --selftest` | fereastra autotestului în limba cerută, „TOTUL OK” / „ALL OK” | ⬜ |
