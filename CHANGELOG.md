# Istoric versiuni

## 0.3.2 — pregătire pentru publicare
- Notă de copyright și de licență (AGPL-3.0-or-later, textul recomandat de FSF) în fiecare fișier sursă.
- README: avertisment explicit că redactarea nu e garantată și trebuie verificată de un om; secțiune de raportare.
- `SECURITY.md`: cum se raportează o scăpare de redactare fără a expune date reale.

## 0.3.1 — repository
- README restructurat, CHANGELOG, docs/TESTARE-WINDOWS.md, LICENSE (AGPL-3.0), .gitignore care blochează documentele.
- pytest fixat la 9.1.1; LICENSE inclus în pachetul .exe.

## 0.3.0 — etapa 3: pachet .exe
- Build PyInstaller 6.22.3 „onedir”, fără consolă, fără UPX (`pii_contracte.spec`, `build.ps1`).
- Jurnal în `%LOCALAPPDATA%\pii_contracte\pii_contracte.log`; erorile din interfață afișate; `--selftest`
  (PDF, DOCX, Tcl/Tk, Word) pe date sintetice.
- Raport CSV: datele calendaristice complet mascate; CNP/IBAN/acte/telefon/permis/CF doar ultimele 2 caractere.
- Nota „dată indirectă” (fără „nebifată implicit”, înșelătoare pe rândurile bifate).
- Registry: cheia `HKCU\Software\pii_contracte` e ștearsă când rămâne goală.
- `tasklist`/`taskkill` primesc stdin explicit (în .exe fără consolă handle-urile standard pot lipsi).
- Licență AGPL-3.0, `LICENSES.txt` + licențele dependențelor copiate în pachet.
- Verificat pe Windows 11: autotestul din .exe „TOTUL OK”, inclusiv Word prin pywin32.

## 0.2.6 — restaurarea registry-ului
- Pe Windows, după 0.2.5 `DisableConvertPdfWarning` rămânea 1. Cauza probabilă (nevalidată): Word își scrie
  opțiunile la închidere. Acum: după `Quit()` se așteaptă ieșirea procesului (max. 15 s, apoi taskkill), apoi se
  restaurează valoarea și se verifică prin recitire.
- Valoarea inițială e notată în `HKCU\Software\pii_contracte\RegistryRestore` și restaurată la pornirea următoare
  dacă programul a fost oprit forțat.
- Verificat pe Windows 11: după export PDF → DOCX, valoarea nu mai există.

## 0.2.5 — blocaj la export PDF → DOCX
- Cauza: dialogul Word „Word will now convert your PDF…”, nesuprimat de `DisplayAlerts`; cu Word invizibil,
  așteaptă la nesfârșit. Suprimat temporar prin `DisableConvertPdfWarning = 1` (sursă: comunitatea Blue Prism,
  nu documentație Microsoft).
- Watchdog: un apel Word mai lung de 120 s oprește doar instanța programului; Word repornește pentru fișierul următor.

## 0.2.4 — format de export
- „Copii redactate ca: formatul original / PDF / DOCX”. Redactarea se face în formatul sursei; Word exportă copia
  deja redactată; ambele se păstrează; exportul e verificat (`engine.verify_file`).

## 0.2.3 — detectoare noi (după primul contract real)
- Bifate implicit: număr permis de conducere, carte funciară / cadastral, adresa imobilului.
- Nebifate implicit (date indirecte, nepropagate): loc naștere, emitent act, date ale actelor.
- Adresele se opresc și la „domiciliat”, „înregistrat/înscris”, „compus”, „având”.

## 0.2.2 — teste robuste pe Windows
- Testele normalizează textul extras (NBSP la Arial); varianta U+2011 se sare dacă fontul nu are glifa;
  smoke test-ul GUI pică (nu se sare) dacă Tk nu pornește pe Windows.
- Diagnostic: Tk nu pornește din venv cu CPython 3.13.0 (`sys.prefix` în loc de `sys.base_prefix`, corectat în 3.13.1).

## 0.2.1 — cratime
- PDF: U+00AD/U+2010/U+2011 → `-`, U+00A0/U+202F → spațiu (normalizare 1:1). Cu Arial, cratima e extrasă ca
  U+00AD, iar numele cu cratimă nu erau detectate.
- DOCX: `<w:noBreakHyphen/>` tratat ca `-`, `<w:softHyphen/>` ca invizibil.

## 0.2.0 — etapa 2: interfață
- Interfață Tkinter: fișiere, scanare în thread separat, tabel cu bife (per apariție și per valoare), filtru pe fișier,
  termeni manuali cu rescanare (debifările se păstrează), avertismente, generare + raport.
- `.doc` prin Word COM: instanță separată, macro-uri dezactivate, read-only, fără dialoguri; conversia în
  `Conversie\` lângă original.
- `verify()` ignoră valorile debifate intenționat; nume de ieșire unice; teste cross-platform.

## 0.1.0 — etapa 1: nucleu
- Detecție și redactare pentru DOCX (lxml, direct pe XML) și PDF digital (PyMuPDF, redactare reală).
- Verificare post-redactare, raport CSV mascat, protecții XXE / zip-bomb.
