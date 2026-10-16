# pii_contracte - detectarea și redactarea datelor de identificare din contracte.
# Copyright (C) 2026  CatalinFusea
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published
# by the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
# SPDX-License-Identifier: AGPL-3.0-or-later

"""Texte vizibile pentru utilizator, în română (implicit) și engleză.

- Identificatorii interni (tipurile de date, ex. "Nume", "CNP") NU se traduc; doar etichetele lor (kind_label).
- Mesajele se generează în limba activă în momentul în care sunt create (ex. avertismentele la scanare).
- Fiecare cheie are ambele limbi, cu aceiași parametri {..} (verificat de tests/test_i18n.py).
"""
from __future__ import annotations

LANGS = ("ro", "en")
_lang = "ro"


def set_lang(lang: str) -> None:
    global _lang
    if lang not in LANGS:
        raise ValueError(f"Limbă necunoscută: {lang}")
    _lang = lang


def get_lang() -> str:
    return _lang


def t(key: str, **kw) -> str:
    ro, en = CATALOG[key]
    return (en if _lang == "en" else ro).format(**kw) if kw else (en if _lang == "en" else ro)


def kind_label(kind: str) -> str:
    """Eticheta afișată a unui tip intern (ex. "Nume" -> "Name"). Tipuri necunoscute: neschimbate."""
    pair = KINDS.get(kind)
    return kind if pair is None else (pair[1] if _lang == "en" else pair[0])


def conf_label(conf_value: str) -> str:
    return t("conf." + {"mare": "high", "medie": "medium", "mică": "low"}.get(conf_value, "medium"))


# tip intern -> (etichetă RO, etichetă EN)
KINDS: dict[str, tuple[str, str]] = {
    "Nume": ("Nume", "Name"),
    "Adresă": ("Adresă", "Address"),
    "CNP": ("CNP", "CNP"),
    "IBAN": ("IBAN", "IBAN"),
    "Email": ("Email", "Email"),
    "Telefon": ("Telefon", "Phone"),
    "Serie act identitate": ("Serie act identitate", "ID card series"),
    "Număr act identitate": ("Număr act identitate", "ID card number"),
    "Număr pașaport": ("Număr pașaport", "Passport number"),
    "Data nașterii": ("Data nașterii", "Date of birth"),
    "Adresă imobil": ("Adresă imobil", "Property address"),
    "Carte funciară / cadastral": ("Carte funciară / cadastral", "Land register / cadastral no."),
    "Număr permis conducere": ("Număr permis conducere", "Driving licence number"),
    "Loc naștere": ("Loc naștere", "Place of birth"),
    "Emitent act": ("Emitent act", "Document issuer"),
    "Dată act": ("Dată act", "Document date"),
    "Manual": ("Manual", "Manual"),
}

CATALOG: dict[str, tuple[str, str]] = {
    # ---------------------------------------------------------------- nume de fișiere / foldere
    "file.redacted_suffix": ("_REDACTAT", "_REDACTED"),
    "file.out_dir_prefix": ("redactat_", "redacted_"),
    "file.report_prefix": ("raport_", "report_"),
    "file.conversion_dir": ("Conversie", "Converted"),
    "redact.placeholder": ("[REDACTAT:{kind}]", "[REDACTED:{kind}]"),

    # ---------------------------------------------------------------- încredere / localizare
    "conf.high": ("mare", "high"),
    "conf.medium": ("medie", "medium"),
    "conf.low": ("mică", "low"),
    "loc.page": ("pagina {n}", "page {n}"),
    "loc.paragraph": ("{part}, paragraful {n}", "{part}, paragraph {n}"),
    "loc.hyperlink": ("hyperlink", "hyperlink"),
    "part.document": ("corp", "body"),
    "part.header": ("antet", "header"),
    "part.footer": ("subsol", "footer"),
    "part.footnotes": ("note de subsol", "footnotes"),
    "part.endnotes": ("note de final", "endnotes"),
    "part.comments": ("comentarii", "comments"),

    # ---------------------------------------------------------------- note pe constatări
    "note.cnp_county": ("cod județ {code} necunoscut", "unknown county code {code}"),
    "note.cnp_invalid": ("cifra de control sau data nu se validează",
                         "check digit or date does not validate"),
    "note.iban_invalid": ("cifra de control nu se validează", "check digit does not validate"),
    "note.uncertain_end": ("sfârșitul e incert - verificați", "end of value is uncertain - check it"),
    "note.indirect": ("dată indirectă", "indirect data"),
    "note.propagated": ("aceeași valoare apare și în altă parte", "same value found elsewhere"),
    "note.manual": ("adăugat manual", "added manually"),

    # ---------------------------------------------------------------- erori / avertismente la citire
    "err.unsupported": ("Format nesuportat: {ext}", "Unsupported format: {ext}"),
    "err.no_ext": ("(fără extensie)", "(no extension)"),
    "err.too_big": ("Fișier prea mare (peste 100 MB).", "File too large (over 100 MB)."),
    "err.unexpected_read": ("Eroare neașteptată la citire: {name}", "Unexpected read error: {name}"),
    "err.docx_invalid": ("Fișierul nu e un .docx valid (poate e protejat cu parolă sau corupt).",
                         "The file is not a valid .docx (it may be password-protected or corrupted)."),
    "err.docx_suspect": ("Fișier .docx suspect (prea mare la dezarhivare).",
                         "Suspicious .docx file (too large when decompressed)."),
    "err.docx_no_document": ("Fișierul nu conține word/document.xml.",
                             "The file does not contain word/document.xml."),
    "err.pdf_open": ("Fișierul nu poate fi deschis ca PDF (corupt sau alt format).",
                     "The file cannot be opened as a PDF (corrupted or another format)."),
    "err.pdf_password": ("PDF protejat cu parolă - deschideți-l și salvați-l fără parolă.",
                         "Password-protected PDF - open it and save it without a password."),
    "err.pdf_too_many_pages": ("PDF prea mare ({n} pagini).", "PDF too large ({n} pages)."),
    "err.overwrite_original": ("Ieșirea ar suprascrie originalul.", "The output would overwrite the original."),
    "warn.track_changes": ("Conține modificări urmărite (track changes). În copia redactată sunt acceptate automat, "
                           "iar textul șters dispare.",
                           "Contains tracked changes. They are accepted automatically in the redacted copy "
                           "and deleted text disappears."),
    "warn.comments": ("Conține comentarii. În copia redactată sunt eliminate.",
                      "Contains comments. They are removed from the redacted copy."),
    "warn.docx_images": ("Conține {n} imagine(i). Datele din imagini (ex. buletin scanat lipit în document) "
                         "NU sunt detectate - verificați manual.",
                         "Contains {n} image(s). Data inside images (e.g. a scanned ID card pasted into the "
                         "document) is NOT detected - check manually."),
    "warn.docx_embedded": ("Conține obiecte încorporate (ex. Excel). Nu sunt analizate - verificați manual.",
                           "Contains embedded objects (e.g. Excel). They are not analysed - check manually."),
    "warn.pdf_scanned": ("Pagini fără text, cu imagini (probabil scanate): {pages}. NU sunt analizate "
                         "(fără OCR deocamdată) - verificați manual.",
                         "Pages without text, with images (probably scanned): {pages}. They are NOT analysed "
                         "(no OCR yet) - check manually."),
    "warn.pdf_forms": ("PDF cu câmpuri de formular. Valorile din câmpuri NU sunt analizate - verificați manual.",
                       "PDF with form fields. Field values are NOT analysed - check manually."),
    "warn.pdf_images": ("Conține imagini. Datele din imagini (ex. semnătură, buletin) NU sunt detectate.",
                        "Contains images. Data inside images (e.g. signature, ID card) is NOT detected."),
    "verify.still_present": ("{kind} {masked} apare încă în copia redactată",
                             "{kind} {masked} is still present in the redacted copy"),
    "verify.unknown_format": ("Format necunoscut pentru verificare: {ext}", "Unknown format for verification: {ext}"),

    # ---------------------------------------------------------------- raport CSV
    "report.columns": ("Fișier;Tip;Valoare (mascată);Încredere;Localizare;Notă;Redactat",
                       "File;Type;Value (masked);Confidence;Location;Note;Redacted"),
    "report.error": ("EROARE", "ERROR"),
    "report.warning": ("AVERTISMENT", "WARNING"),
    "report.yes": ("da", "yes"),
    "report.no": ("nu", "no"),

    # ---------------------------------------------------------------- Word / conversie
    "word.no_pywin32": ("Lipsește pywin32 - conversia .doc nu e disponibilă.",
                        "pywin32 is missing - .doc conversion is not available."),
    "word.cannot_start": ("Microsoft Word nu poate fi pornit (nu e instalat sau e blocat).",
                          "Microsoft Word cannot be started (not installed or blocked)."),
    "word.settings_refused": ("Word nu a acceptat setările de siguranță; conversia e oprită.",
                              "Word did not accept the safety settings; conversion stopped."),
    "word.only_doc": ("Se convertesc doar fișiere .doc.", "Only .doc files are converted."),
    "word.cannot_create_dir": ("Nu pot crea folderul „{dir}” lângă original ({err}).",
                               "Cannot create the \"{dir}\" folder next to the original ({err})."),
    "word.export_unsupported": ("Export nesuportat: {ext} -> {fmt}.", "Unsupported export: {ext} -> {fmt}."),
    "word.not_started": ("Convertorul nu e pornit.", "The converter is not started."),
    "word.timeout": ("Word nu a răspuns în {s} s (probabil un dialog ascuns) și a fost oprit.",
                     "Word did not respond within {s} s (probably a hidden dialog) and was stopped."),
    "word.unknown_format": ("Format necunoscut: {fmt}", "Unknown format: {fmt}"),
    "word.overwrite_source": ("Conversia ar suprascrie sursa.", "The conversion would overwrite the source."),
    "word.cannot_open": ("Word nu a putut deschide fișierul (corupt, protejat cu parolă sau deschis în alt program).",
                         "Word could not open the file (corrupted, password-protected or open in another "
                         "program)."),
    "word.cannot_save": ("Word nu a putut salva conversia (folder fără drept de scriere sau fișier destinație "
                         "deschis).",
                         "Word could not save the conversion (no write permission or the target file is open)."),
    "word.no_output": ("Word nu a produs fișierul convertit.", "Word did not produce the converted file."),
    "word.cannot_inspect": ("Nu am putut verifica dacă documentul .doc are modificări urmărite sau comentarii.",
                            "Could not check whether the .doc document has tracked changes or comments."),
    "word.revisions_pdf": ("Documentul .doc are {revs} modificări urmărite și {comments} comentarii. În PDF-ul "
                           "convertit pot apărea ca marcaje (inclusiv text șters) - verificați PDF-ul redactat "
                           "sau alegeți conversia în DOCX.",
                           "The .doc document has {revs} tracked changes and {comments} comments. They may appear "
                           "as markup in the converted PDF (including deleted text) - check the redacted PDF or "
                           "choose conversion to DOCX."),

    # ---------------------------------------------------------------- operații (jobs)
    "stage.scan": ("scanare", "scanning"),
    "stage.convert": ("conversie Word", "Word conversion"),
    "stage.redact": ("redactare", "redacting"),
    "stage.export": ("export {fmt} cu Word", "{fmt} export with Word"),
    "job.converted_doc": ("Convertit din .doc cu Word. Copia convertită (NEREDACTATĂ) e în folderul „{dir}” "
                          "de lângă original.",
                          "Converted from .doc with Word. The converted (UNREDACTED) copy is in the \"{dir}\" "
                          "folder next to the original."),
    "job.scan_fatal": ("Eroare neașteptată la scanare: {name}", "Unexpected error while scanning: {name}"),
    "job.redact_failed": ("Copia nu a putut fi generată ({name}).", "The copy could not be generated ({name})."),
    "job.export_impossible": ("Export {fmt} imposibil: {err} Copia redactată a rămas în formatul inițial.",
                              "{fmt} export not possible: {err} The redacted copy remains in its original format."),
    "job.export_failed": ("Export {fmt} eșuat: {err} Copia redactată a rămas în formatul inițial.",
                          "{fmt} export failed: {err} The redacted copy remains in its original format."),
    "job.generate_fatal": ("Eroare la generare: {name}", "Error while generating: {name}"),

    # ---------------------------------------------------------------- interfață
    "app.title": ("Redactare date de identificare - contracte", "Personal data redaction - contracts"),
    "ft.contracts": ("Contracte (Word, PDF)", "Contracts (Word, PDF)"),
    "ft.all": ("Toate fișierele", "All files"),
    "col.file": ("Fișier", "File"),
    "col.kind": ("Tip", "Type"),
    "col.value": ("Valoare", "Value"),
    "col.conf": ("Încredere", "Confidence"),
    "col.loc": ("Localizare", "Location"),
    "col.ctx": ("Context", "Context"),
    "col.note": ("Notă", "Note"),
    "ui.files": ("1. Fișiere", "1. Files"),
    "ui.add": ("Adaugă…", "Add…"),
    "ui.remove_sel": ("Elimină selectate", "Remove selected"),
    "ui.clear": ("Golește lista", "Clear list"),
    "ui.options": ("Opțiuni", "Options"),
    "ui.language": ("Limba / Language:", "Language / Limba:"),
    "ui.doc_to": ("Fișierele .doc se convertesc în:", ".doc files are converted to:"),
    "ui.doc_note": ("Conversia (neredactată) se salvează în „{dir}” lângă original și rămâne acolo.",
                    "The (unredacted) conversion is saved in \"{dir}\" next to the original and stays there."),
    "ui.doc_unavailable": ("Conversia .doc nu e disponibilă (lipsește pywin32 sau nu e Windows). "
                           "Fișierele .doc vor apărea cu eroare.",
                           ".doc conversion is not available (pywin32 missing or not Windows). "
                           ".doc files will show an error."),
    "ui.manual_terms": ("Text de redactat manual (ex. un nume ratat):", "Text to redact manually (e.g. a missed name):"),
    "ui.term_add": ("Adaugă", "Add"),
    "ui.term_del": ("Șterge", "Delete"),
    "ui.scan": ("Scanează", "Scan"),
    "ui.stop": ("Oprește", "Stop"),
    "ui.review": ("2. Verificare - debifați ce NU trebuie redactat (click pe căsuță sau tasta Spațiu)",
                  "2. Review - untick what must NOT be redacted (click the box or press Space)"),
    "ui.filter_file": ("Fișier:", "File:"),
    "ui.val_on": ("Bifează valoarea peste tot", "Tick this value everywhere"),
    "ui.val_off": ("Debifează valoarea peste tot", "Untick this value everywhere"),
    "ui.all_on": ("Bifează tot", "Tick all"),
    "ui.all_off": ("Debifează tot", "Untick all"),
    "ui.warnings": ("Avertismente și erori (de verificat manual):", "Warnings and errors (check manually):"),
    "ui.generate_frame": ("3. Generare", "3. Generate"),
    "ui.out_dir": ("Folder ieșire:", "Output folder:"),
    "ui.choose": ("Alege…", "Browse…"),
    "ui.out_as": ("Copii redactate ca:", "Redacted copies as:"),
    "ui.fmt_original": ("formatul original", "original format"),
    "ui.generate": ("Generează copii redactate și raport", "Generate redacted copies and report"),
    "ui.open_folder": ("Deschide folderul", "Open folder"),
    "ui.footer": ("Originalele nu se modifică. Valorile complete apar doar în această fereastră; "
                  "raportul CSV conține valori mascate.",
                  "Originals are never modified. Full values are shown only in this window; "
                  "the CSV report contains masked values."),
    "ui.fmt_note_unavailable": ("Exportul în alt format folosește Word și nu e disponibil (lipsește pywin32 sau "
                                "nu e Windows).",
                                "Export to another format uses Word and is not available (pywin32 missing or not "
                                "Windows)."),
    "ui.fmt_note_pdf": ("Documentele Word se redactează în DOCX, apoi Word le exportă în PDF. Se păstrează ambele "
                        "copii; ambele sunt verificate.",
                        "Word documents are redacted as DOCX, then Word exports them to PDF. Both copies are kept "
                        "and both are verified."),
    "ui.fmt_note_docx": ("PDF-urile se redactează ca PDF, apoi Word le convertește în DOCX (PDF Reflow) - aspectul "
                         "se poate strica. Se păstrează ambele copii; ambele sunt verificate.",
                         "PDFs are redacted as PDF, then Word converts them to DOCX (PDF Reflow) - the layout may "
                         "break. Both copies are kept and both are verified."),
    "ui.status_start": ("Adăugați fișiere (.docx, .doc, .pdf), apoi apăsați „Scanează”.",
                        "Add files (.docx, .doc, .pdf), then press \"Scan\"."),
    "ui.pick_files": ("Alegeți contractele", "Choose the contracts"),
    "ui.unsupported_skipped": ("Format nesuportat, ignorate:", "Unsupported format, skipped:"),
    "ui.dirty": ("Lista de fișiere, termenii sau formatul .doc s-au schimbat - apăsați din nou „Scanează” "
                 "(debifările existente se păstrează).",
                 "The file list, terms or .doc format changed - press \"Scan\" again (existing unticks are kept)."),
    "ui.scanning": ("Scanare…", "Scanning…"),
    "ui.stopping": ("Se oprește după fișierul curent…", "Stopping after the current file…"),
    "ui.progress": ("{stage} {i}/{n}: {name}", "{stage} {i}/{n}: {name}"),
    "ui.verify_problems": ("{name}: VERIFICAREA A GĂSIT PROBLEME în copia redactată - {problems}",
                           "{name}: VERIFICATION FOUND PROBLEMS in the redacted copy - {problems}"),
    "ui.all_files": ("Toate fișierele", "All files"),
    "ui.file_error": ("{name}: EROARE - {err}", "{name}: ERROR - {err}"),
    "ui.link_rows": ("Rândurile cu localizarea „hyperlink”: dacă valoarea e bifată oriunde în document, linkul e "
                     "eliminat chiar dacă rândul acesta e debifat.",
                     "Rows with location \"hyperlink\": if the value is ticked anywhere in the document, the link "
                     "is removed even if this row is unticked."),
    "ui.link_note": ("link (vezi nota de jos)", "link (see note below)"),
    "ui.scan_done": ("Scanare terminată: {files} fișiere, {total} constatări", "Scan finished: {files} files, "
                     "{total} findings"),
    "ui.scan_errors": (", {n} fișiere cu eroare (vezi mai jos)", ", {n} files with errors (see below)"),
    "ui.stopped_f": (" - OPRITĂ înainte de final", " - STOPPED before the end"),
    "ui.stopped_m": (" - OPRIT înainte de final", " - STOPPED before the end"),
    "ui.scan_next": (". Verificați tabelul, apoi generați copiile.", ". Review the table, then generate the copies."),
    "ui.ticked": ("Bifate: {sel}/{total}", "Ticked: {sel}/{total}"),
    "ui.choose_out": ("Folder pentru copiile redactate", "Folder for the redacted copies"),
    "ui.need_out": ("Alegeți folderul de ieșire.", "Choose the output folder."),
    "ui.same_folder": ("Folderul de ieșire e același cu al originalelor. Copiile redactate vor sta lângă originale "
                       "și pot fi confundate.\n\nContinuați?",
                       "The output folder is the same as the originals' folder. Redacted copies will sit next to "
                       "the originals and may be confused with them.\n\nContinue?"),
    "ui.fmt_txt_original": ("în formatul original", "in the original format"),
    "ui.fmt_txt_pdf": ("în PDF (plus copia redactată originală)", "as PDF (plus the redacted original-format copy)"),
    "ui.fmt_txt_docx": ("în DOCX (plus copia redactată originală)", "as DOCX (plus the redacted original-format copy)"),
    "ui.confirm_generate": ("Se generează {n} copii redactate {fmt} ({sel} din {total} constatări bifate) și "
                            "raportul CSV în:\n{out}",
                            "Generating {n} redacted copies {fmt} ({sel} of {total} findings ticked) and the CSV "
                            "report in:\n{out}"),
    "ui.confirm_skipped": ("\n\n{n} fișiere cu eroare NU vor avea copie redactată.",
                           "\n\n{n} files with errors will NOT get a redacted copy."),
    "ui.continue_q": ("\n\nContinuați?", "\n\nContinue?"),
    "ui.generating": ("Generare…", "Generating…"),
    "ui.gen_done": ("Generate: {n} copii redactate", "Generated: {n} redacted copies"),
    "ui.gen_failed": (", {n} eșuate", ", {n} failed"),
    "ui.gen_problems": (", {n} cu probleme la verificare (vezi lista de avertismente)",
                        ", {n} with verification problems (see the warnings list)"),
    "ui.gen_report": (".\nRaport: {name}", ".\nReport: {name}"),
    "ui.busy_close": ("O operație e în curs. Se oprește după fișierul curent și fereastra se închide. Continuați?",
                      "An operation is in progress. It will stop after the current file and the window will "
                      "close. Continue?"),
    "ui.unexpected": ("Eroare neașteptată ({name}).\nDetalii în:\n{path}", "Unexpected error ({name}).\nDetails in:\n"
                      "{path}"),
    "ui.lang_rescan": ("Rezultatele scanării se pierd la schimbarea limbii și trebuie rescanat. Continuați?",
                       "Scan results are cleared when the language changes and you will need to scan again. "
                       "Continue?"),

    # ---------------------------------------------------------------- autotest
    "st.title": ("Autotest pii_contracte", "pii_contracte self-test"),
    "st.frozen": ("pachet .exe", ".exe package"),
    "st.source": ("din sursă", "from source"),
    "st.pdf": ("PDF: detecție + redactare + verificare", "PDF: detection + redaction + verification"),
    "st.docx": ("DOCX: detecție + redactare + verificare", "DOCX: detection + redaction + verification"),
    "st.tk": ("Interfață (Tcl/Tk)", "User interface (Tcl/Tk)"),
    "st.word": ("Microsoft Word (pywin32/COM)", "Microsoft Word (pywin32/COM)"),
    "st.word_ok": ("Word pornit, export DOCX→PDF, închis", "Word started, DOCX→PDF export, closed"),
    "st.word_skipped": ("[SĂRIT]  Microsoft Word: doar pe Windows", "[SKIPPED] Microsoft Word: Windows only"),
    "st.ok": ("[OK]    ", "[OK]    "),
    "st.error": ("[EROARE]", "[ERROR] "),
    "st.all_ok": ("REZULTAT: TOTUL OK", "RESULT: ALL OK"),
    "st.has_errors": ("REZULTAT: EXISTĂ ERORI - vezi liniile [EROARE] și logul",
                      "RESULT: ERRORS FOUND - see the [ERROR] lines and the log"),
    "st.report_saved": ("Raport salvat în:", "Report saved in:"),
    "st.cnp_missing": ("CNP nedetectat (eroare: {err})", "CNP not detected (error: {err})"),
    "st.incomplete": ("redactare incompletă: {problems}", "incomplete redaction: {problems}"),
    "st.no_pywin32": ("pywin32 lipsește - .doc și exportul prin Word nu vor merge",
                      "pywin32 is missing - .doc and export via Word will not work"),
    "st.no_pdf": ("Word nu a produs PDF-ul", "Word did not produce the PDF"),
}
