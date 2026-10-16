# pii_contracte

Aplicație Windows care **detectează și redactează datele de identificare** din contracte
(**.docx**, **.pdf digital**, **.doc** prin Microsoft Word). Procesarea e locală, fără rețea.

Flux: alegi fișierele → **Scanează** → verifici constatările într-un tabel cu bife → **Generează** copii redactate
(`<nume>_REDACTAT.<ext>`) și un raport CSV cu valori mascate. Originalele nu sunt modificate.

> Stare: versiunea 0.3.1, folosită intern. Testată pe Windows 11 cu Python 3.14 și Microsoft Word 16.0.
> Paginile scanate (imagini) **nu** sunt analizate încă (OCR = etapa 4). Istoric: [CHANGELOG.md](CHANGELOG.md).

## Ce detectează

| Tip | Cum | Bifat implicit |
|---|---|---|
| CNP | cifră de control + dată + cod județ (validare eșuată lângă „CNP” → încredere mică) | da |
| IBAN | mod-97 | da |
| Email, telefon RO | tipar | da |
| Serie/număr CI, număr pașaport | ancoră („seria … nr.”, „pașaport nr.”) | da |
| Număr permis de conducere | ancoră („permis de conducere … nr.”) | da |
| Data nașterii | ancoră („născut(ă) la …”) | da |
| Nume | ancoră („subsemnatul”, „domnul/doamna”, „reprezentată prin”, „…, domiciliat/CNP”) | da |
| Adresă de domiciliu | după „domiciliat în / cu domiciliul / reședința” | da |
| Adresa imobilului | după „situat/amplasat/aflat în”, „imobilul/apartamentul din” | da |
| Carte funciară / nr. cadastral | „Cartea Funciară … nr.”, „CF nr.”, „număr cadastral” | da |
| Loc naștere, emitent act, date ale actelor | ancore („născut … în”, „eliberat de”, „valabil până la”) | **nu** (date indirecte) |
| Orice valoare găsită o dată | căutată în tot documentul (ex. numele din semnătură) | ca originalul |
| Text manual | introdus în interfață | da |

**Nu detectează:** text din imagini și pagini scanate, câmpuri de formular PDF, nume fără ancoră, banca, starea civilă.
Pentru acestea există câmpul „Text de redactat manual”. **Copia redactată trebuie verificată înainte de a fi trimisă.**

## Ce face copia redactată
- **PDF:** textul e șters din conținutul paginii (nu doar acoperit); linkurile cu valori redactate și metadatele sunt
  eliminate.
- **DOCX:** înlocuire cu `[REDACTAT:TIP]`; modificările urmărite sunt acceptate (textul șters dispare); comentariile
  și metadatele (autor, companie) sunt eliminate.
- **Verificare automată:** copia (și exportul în alt format) e recitită; dacă o valoare bifată mai apare, interfața
  afișează „VERIFICAREA A GĂSIT PROBLEME”.
- **Export:** opțional în PDF sau DOCX prin Word (se păstrează și copia în formatul original). PDF → DOCX strică aspectul.
- **.doc:** convertit cu Word în `Conversie\` lângă original (copie **neredactată**, rămâne acolo), apoi redactat.

## Utilizare (.exe)
Folderul `pii_contracte\` din build se copiază întreg; se pornește `pii_contracte.exe`. Instrucțiuni scurte pentru
utilizator: [CITESTE-MA.txt](CITESTE-MA.txt). Autotest: `pii_contracte.exe --selftest`.

## Dezvoltare

Cerințe: Windows 10/11, **Python 3.14** (de pe python.org, cu „tcl/tk and IDLE”), Microsoft Word pentru .doc și export.
Nu folosiți Python 3.13.0: Tk nu pornește din venv (bug CPython, rezolvat în 3.13.1).

```powershell
git clone <url-repo> pii_contracte
cd pii_contracte
py -3.14 -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass   # doar pentru fereastra curentă
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
python -m pytest -q -rs tests        # 87 trecute, 1 sărit pe Windows (U+2011 fără glifă în Arial)
python -m pii_gui                    # interfața din sursă
```

Build `.exe` (PyInstaller 6.22.3, onedir, fără consolă):
```powershell
.\build.ps1
```
Scriptul face un venv separat (`.venv-build`), rulează testele, construiește `dist\pii_contracte\`, rulează
autotestul din `.exe` și creează `dist\pii_contracte_<versiune>.zip` cu SHA256.

Teste manuale pe Windows (Word, registry, .exe pe un cont fără Python): [docs/TESTARE-WINDOWS.md](docs/TESTARE-WINDOWS.md).

### Structură
```
pii_core/          nucleu fără interfață
  detectors.py     detectoare (regex + validări) pe text românesc
  engine.py        scan_file / redact_file / verify / write_report (raport mascat)
  docx_handler.py  DOCX direct pe XML (lxml, fără XXE): corp, tabele, antete, note, track changes
  pdf_handler.py   PDF cu PyMuPDF: text cu coordonate, redactare reală
  doc_convert.py   Microsoft Word prin COM (pywin32): .doc → docx/pdf, export, watchdog, registry
  models.py
pii_gui/           interfața Tkinter
  app.py           fereastra; model.py = starea bifelor; jobs.py = operații în thread separat
  runtime.py       log în fișier, --selftest
tests/             pytest; documente 100% sintetice (fixtures.py)
pii_contracte_app.py   punctul de intrare pentru .exe
pii_contracte.spec, build.ps1, requirements-build.txt   build
demo_cli.py, diag_iban.py   utilitare de linie de comandă (afișează valori complete în consolă!)
```

### Dependențe (versiuni fixate în `requirements*.txt`)
| Pachet | Versiune | Rol | Licență |
|---|---|---|---|
| pymupdf | 1.28.2 | PDF | AGPL-3.0 / comercial |
| lxml | 6.1.3 | DOCX | BSD-3-Clause |
| pywin32 | 312 (doar Windows) | Word COM | PSF |
| pytest, python-docx | 9.1.1, 1.2.0 | doar teste | MIT |
| pyinstaller | 6.22.3 | doar build | GPLv2+ cu excepție pentru bootloader |

## Securitate și confidențialitate
- Fără acces la rețea. Originalele sunt doar citite; ieșirea nu poate suprascrie un original.
- Valorile complete apar doar în fereastra aplicației. Raportul CSV le maschează (datele calendaristice complet,
  numerele până la ultimele 2 caractere). Jurnalul (`%LOCALAPPDATA%\pii_contracte\`) nu conține valori,
  dar poate conține nume de fișiere.
- XML parsat fără entități externe și fără rețea; limite anti zip-bomb și de mărime.
- Word: instanță separată, invizibilă, **macro-uri dezactivate** (`AutomationSecurity = ForceDisable`), documente
  deschise read-only, fără dialoguri. Dialogul de conversie PDF al Word e suprimat temporar prin
  `HKCU\Software\Microsoft\Office\<ver>\Word\Options\DisableConvertPdfWarning`; valoarea inițială e restaurată după
  închiderea Word (și la pornirea următoare, dacă programul a fost oprit forțat). Un apel Word blocat peste 120 s
  duce la oprirea doar a instanței programului.
- **Nu adăugați în repository contracte reale, rapoarte sau copii redactate.** `.gitignore` blochează `.doc/.docx/.pdf/.csv`
  în afara testelor; verificați `git status` înainte de fiecare commit.

## Limitări cunoscute
- Fără OCR: pagini scanate și imagini (ex. buletin lipit în document) — doar avertisment.
- Nume fără ancoră, nume inversate (Ion Popescu vs POPESCU ION), grafii diferite ale aceluiași nume.
- Fișierele din `Conversie\` sunt copii complete, neredactate.
- `.exe`-ul nu e semnat digital: SmartScreen cere confirmare la prima rulare.

## Licență
[GNU AGPL-3.0](LICENSE) — impusă de PyMuPDF (AGPL-3.0). Uz intern: fără obligații suplimentare. Dacă programul
(sau `.exe`-ul) e dat altcuiva, codul sursă complet trebuie pus la dispoziția acelei persoane, sub AGPL-3.0.
Licențele componentelor terțe: [LICENSES.txt](LICENSES.txt).
