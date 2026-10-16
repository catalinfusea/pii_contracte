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

"""Conversie .doc -> .docx / .pdf prin Microsoft Word (COM, pywin32). Doar Windows.

Siguranță:
- Instanță Word separată (DispatchEx): nu atinge documentele deschise de utilizator.
- Macro-uri dezactivate pentru fișierele deschise programatic: AutomationSecurity = ForceDisable (3).
  Implicit Word folosește „Low” (macro-uri ACTIVATE) la deschiderea programatică.
- Fără dialoguri (DisplayAlerts = wdAlertsNone), document deschis read-only, invizibil,
  neadăugat în lista de fișiere recente.
- Parolă falsă la deschidere: un document protejat eșuează cu eroare în loc să ceară parola
  (comportament de verificat pe Windows - vezi checklist).
- Dialogul „Word will now convert your PDF…” (PDF Reflow) NU e suprimat de DisplayAlerts și, cu Word
  invizibil, blochează conversia. Pe durata instanței noastre se setează
  HKCU\\Software\\Microsoft\\Office\\<ver>\\Word\\Options\\DisableConvertPdfWarning = 1 (DWORD),
  apoi valoarea anterioară e restaurată DUPĂ ieșirea procesului Word (Word își poate rescrie opțiunile la
  închidere) și verificată prin recitire; valoarea inițială e notată și în HKCU\\Software\\pii_contracte, ca să
  fie restaurată la pornirea următoare dacă programul e oprit forțat.
  (Sursă pentru valoare: comunitate Blue Prism, nu documentație Microsoft.)
- Watchdog: dacă un apel către Word depășește TIMEOUT_S, instanța NOASTRĂ de Word (PID-ul apărut la pornire)
  e oprită cu taskkill; fișierul primește eroare, aplicația nu rămâne blocată.

Constante (learn.microsoft.com, Word VBA reference):
  WdSaveFormat: wdFormatXMLDocument = 12, wdFormatPDF = 17
  WdAlertLevel: wdAlertsNone = 0
  MsoAutomationSecurity: msoAutomationSecurityForceDisable = 3
  WdSaveOptions: wdDoNotSaveChanges = 0
"""
from __future__ import annotations

import logging
import os
import subprocess
import sys
import threading

log = logging.getLogger("pii_core")

CONVERSION_DIR = "Conversie"
FORMATS = {"docx": (12, ".docx"), "pdf": (17, ".pdf")}
_WD_ALERTS_NONE = 0
_MSO_AUTOSEC_FORCE_DISABLE = 3
_WD_DO_NOT_SAVE = 0
_DUMMY_PASSWORD = "pii-contracte-fara-parola-7731"


TIMEOUT_S = 120


def _sleep(seconds: float) -> None:
    __import__("time").sleep(seconds)


QUIT_WAIT_S = 15
_OFFICE_KEY = r"Software\Microsoft\Office"
_PDF_WARN_VALUE = "DisableConvertPdfWarning"


class ConversionError(Exception):
    pass


# ------------------------------------------------------------------ utilitare Windows
def _no_window() -> dict:
    # .exe fără consolă: fără fereastră nouă și cu stdin explicit (handle-urile standard pot lipsi)
    kw = {"stdin": subprocess.DEVNULL}
    if sys.platform == "win32":
        kw["creationflags"] = 0x08000000   # CREATE_NO_WINDOW
    return kw


def list_word_pids() -> set[int]:
    """PID-urile WINWORD.EXE curente (tasklist; fără dependențe în plus)."""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, timeout=15, **_no_window()).stdout
    except (OSError, subprocess.SubprocessError):
        return set()
    pids = set()
    for line in out.splitlines():
        parts = [x.strip('"') for x in line.split('","')]
        if len(parts) > 1 and parts[0].lower() == "winword.exe" and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def kill_pid(pid: int) -> None:
    subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, timeout=15,
                   **_no_window())


_MARKER_KEY = r"Software\pii_contracte\RegistryRestore"   # valoarea inițială, pentru recuperare după crash
_ABSENT = "<absent>"


class PdfWarningSuppressed:
    """Setează DisableConvertPdfWarning=1 pentru versiunile Office cu cheie Word în HKCU și restaurează la ieșire.

    Valoarea inițială e notată întâi în HKCU\\Software\\pii_contracte\\RegistryRestore; dacă programul e oprit
    forțat, `recover()` (apelat la următoarea pornire a convertorului) o restaurează.
    """

    def __init__(self, versions=None, winreg_mod=None, sleep=None):
        self._versions = versions
        self._wr = winreg_mod
        self._sleep = sleep or __import__("time").sleep
        self._saved: list[tuple[str, object]] = []   # (cale, valoare anterioară sau None = nu exista)

    def _winreg(self):
        if self._wr is None:
            try:
                import winreg
            except ImportError:
                return None
            self._wr = winreg
        return self._wr

    def _office_versions(self, winreg) -> list[str]:
        vers = set(self._versions or ())
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _OFFICE_KEY) as k:
                i = 0
                while True:
                    try:
                        name = winreg.EnumKey(k, i)
                    except OSError:
                        break
                    i += 1
                    try:
                        winreg.OpenKey(k, name + r"\Word").Close()
                    except OSError:
                        continue
                    if name.replace(".", "").isdigit():
                        vers.add(name)
        except OSError:
            pass
        return sorted(vers or {"16.0"})

    # ---- citire / scriere elementare
    @staticmethod
    def _read(winreg, path):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
                return winreg.QueryValueEx(k, _PDF_WARN_VALUE)[0]
        except OSError:
            return None

    def _write(self, winreg, path, value) -> bool:
        """value None = șterge. Întoarce True dacă recitirea confirmă starea."""
        for _ in range(10):
            try:
                with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0,
                                        winreg.KEY_READ | winreg.KEY_SET_VALUE) as k:
                    if value is None:
                        try:
                            winreg.DeleteValue(k, _PDF_WARN_VALUE)
                        except OSError:
                            pass
                    else:
                        winreg.SetValueEx(k, _PDF_WARN_VALUE, 0, winreg.REG_DWORD, value)
            except OSError:
                pass
            if self._read(winreg, path) == value:
                return True
            self._sleep(0.3)
        return False

    def _marker_set(self, winreg, path, prev) -> None:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _MARKER_KEY, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, path, 0, winreg.REG_SZ, _ABSENT if prev is None else str(prev))

    def _marker_clear(self, winreg, path) -> None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _MARKER_KEY, 0, winreg.KEY_SET_VALUE) as k:
                winreg.DeleteValue(k, path)
        except OSError:
            return
        # cheile programului, dacă au rămas goale: nu lăsăm urme în registry
        for key in (_MARKER_KEY, _MARKER_KEY.rsplit("\\", 1)[0]):
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
                    try:
                        winreg.EnumValue(k, 0)
                        return                      # mai are valori
                    except OSError:
                        pass
                    try:
                        winreg.EnumKey(k, 0)
                        return                      # mai are subchei
                    except OSError:
                        pass
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key)
            except OSError:
                return

    def recover(self) -> int:
        """Restaurează valorile rămase de la o rulare întreruptă. Întoarce câte a restaurat."""
        winreg = self._winreg()
        if winreg is None:
            return 0
        pending = []
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _MARKER_KEY) as k:
                i = 0
                while True:
                    try:
                        name, data, _ = winreg.EnumValue(k, i)
                    except OSError:
                        break
                    pending.append((name, data))
                    i += 1
        except OSError:
            return 0
        n = 0
        for path, data in pending:
            prev = None if data == _ABSENT else int(data)
            if self._write(winreg, path, prev):
                self._marker_clear(winreg, path)
                n += 1
        if n:
            log.warning("Restaurat %s după o rulare întreruptă (%d chei)", _PDF_WARN_VALUE, n)
        return n

    def __enter__(self):
        winreg = self._winreg()
        if winreg is None:
            return self
        self.recover()
        for ver in self._office_versions(winreg):
            path = rf"{_OFFICE_KEY}\{ver}\Word\Options"
            try:
                prev = self._read(winreg, path)
                self._marker_set(winreg, path, prev)
                self._saved.append((path, prev))
                if not self._write(winreg, path, 1):
                    log.warning("Nu pot seta %s pentru Office %s", _PDF_WARN_VALUE, ver)
            except OSError:
                log.warning("Nu pot seta %s pentru Office %s", _PDF_WARN_VALUE, ver)
        return self

    def __exit__(self, *exc):
        winreg = self._winreg()
        if winreg is None:
            return
        for path, prev in self._saved:
            if self._write(winreg, path, prev):
                self._marker_clear(winreg, path)
            else:
                log.warning("Nu pot restaura %s (rămâne notat pentru următoarea pornire)", _PDF_WARN_VALUE)
        self._saved = []


def available() -> bool:
    """True dacă pywin32 e instalat (nu garantează că Word e instalat)."""
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def target_path(src: str, fmt: str) -> str:
    """<folderul originalului>\\Conversie\\<nume>.<docx|pdf>"""
    if fmt not in FORMATS:
        raise ValueError(f"Format necunoscut: {fmt}")
    folder = os.path.join(os.path.dirname(os.path.abspath(src)), CONVERSION_DIR)
    stem = os.path.splitext(os.path.basename(src))[0]
    return os.path.join(folder, stem + FORMATS[fmt][1])


class WordConverter:
    """Folosire (în thread-ul care face conversiile):

        with WordConverter() as wc:
            out = wc.convert(r"C:\\...\\contract.doc", "pdf")
    """

    def __init__(self, timeout: float = TIMEOUT_S, list_pids=None, kill=None, registry=None, wait=None):
        self._word = None
        self._pid: int | None = None
        self._com_init = False
        self._reg = None
        self._timeout = timeout
        self._list_pids = list_pids or list_word_pids
        self._kill = kill or kill_pid
        self._registry = registry or PdfWarningSuppressed
        self._wait = wait or _sleep
        self.last_notes: list[str] = []   # observații despre ultimul fișier convertit

    def __enter__(self) -> "WordConverter":
        try:
            import pythoncom
            import win32com.client  # noqa: F401
        except ImportError as e:
            raise ConversionError("Lipsește pywin32 - conversia .doc nu e disponibilă.") from e
        pythoncom.CoInitialize()          # obligatoriu în orice thread care folosește COM
        self._com_init = True
        self._reg = self._registry().__enter__()
        try:
            self._start_word()
        except ConversionError:
            self.__exit__(None, None, None)
            raise
        return self

    def _start_word(self) -> None:
        import win32com.client
        before = self._list_pids()
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as e:  # noqa: BLE001 - com_error / altele
            raise ConversionError("Microsoft Word nu poate fi pornit (nu e instalat sau e blocat).") from e
        new = self._list_pids() - before
        self._pid = next(iter(new)) if len(new) == 1 else None
        if self._pid is None:
            log.warning("Nu am identificat PID-ul instanței Word; watchdog-ul nu poate opri Word")
        self._word = word
        try:
            word.Visible = False
            word.DisplayAlerts = _WD_ALERTS_NONE
            word.AutomationSecurity = _MSO_AUTOSEC_FORCE_DISABLE
        except Exception as e:  # noqa: BLE001
            self._quit()
            raise ConversionError("Word nu a acceptat setările de siguranță; conversia e oprită.") from e

    def convert(self, src: str, fmt: str) -> str:
        """.doc original -> <folderul originalului>\\Conversie\\<nume>.<fmt> (copie NEREDACTATĂ)."""
        if os.path.splitext(src)[1].lower() != ".doc":
            raise ConversionError("Se convertesc doar fișiere .doc.")
        src = os.path.abspath(src)
        dst = target_path(src, fmt)
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
        except OSError as e:
            raise ConversionError(f"Nu pot crea folderul „{CONVERSION_DIR}” lângă original "
                                  f"({type(e).__name__}).") from e
        self._run(src, dst, fmt, inspect=True)
        log.info("Convertit .doc -> %s", fmt)
        return dst

    def export(self, src: str, fmt: str, dst: str) -> str:
        """Copie REDACTATĂ .docx -> .pdf sau .pdf -> .docx (PDF Reflow din Word 2013+)."""
        ext = os.path.splitext(src)[1].lower()
        if ext not in (".docx", ".pdf") or FORMATS[fmt][1] == ext:
            raise ConversionError(f"Export nesuportat: {ext} -> {fmt}.")
        self._run(os.path.abspath(src), os.path.abspath(dst), fmt, inspect=False)
        log.info("Exportat %s -> %s", ext, fmt)
        return dst

    def _run(self, src: str, dst: str, fmt: str, inspect: bool) -> None:
        if not self._com_init:
            raise ConversionError("Convertorul nu e pornit.")
        if self._word is None:            # instanța anterioară a fost oprită de watchdog
            self._start_word()
        killed = threading.Event()

        def on_timeout():
            if self._pid is not None:
                log.warning("Word nu răspunde de %ss; opresc instanța programului", self._timeout)
                killed.set()
                self._kill(self._pid)

        timer = threading.Timer(self._timeout, on_timeout)
        timer.daemon = True
        timer.start()
        try:
            self._run_inner(src, dst, fmt, inspect)
        except ConversionError as e:
            if killed.is_set():
                self._word, self._pid = None, None
                raise ConversionError(f"Word nu a răspuns în {int(self._timeout)} s (probabil un dialog "
                                      f"ascuns) și a fost oprit.") from e
            raise
        finally:
            timer.cancel()

    def _run_inner(self, src: str, dst: str, fmt: str, inspect: bool) -> None:
        if fmt not in FORMATS:
            raise ConversionError(f"Format necunoscut: {fmt}")
        if os.path.normcase(dst) == os.path.normcase(src):
            raise ConversionError("Conversia ar suprascrie sursa.")
        try:
            doc = self._word.Documents.Open(
                FileName=src, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
                PasswordDocument=_DUMMY_PASSWORD, Visible=False, NoEncodingDialog=True)
        except Exception as e:  # noqa: BLE001
            raise ConversionError("Word nu a putut deschide fișierul (corupt, protejat cu parolă "
                                  "sau deschis în alt program).") from e
        self.last_notes = self._inspect(doc, fmt) if inspect else []
        try:
            doc.SaveAs2(FileName=dst, FileFormat=FORMATS[fmt][0], AddToRecentFiles=False)
        except Exception as e:  # noqa: BLE001
            raise ConversionError("Word nu a putut salva conversia (folder fără drept de scriere "
                                  "sau fișier destinație deschis).") from e
        finally:
            try:
                doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
            except Exception:  # noqa: BLE001
                log.warning("Închiderea documentului în Word a eșuat")
        if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
            raise ConversionError("Word nu a produs fișierul convertit.")

    @staticmethod
    def _inspect(doc, fmt: str) -> list[str]:
        """Doar citire. La PDF, Word poate exporta modificările urmărite/comentariile ca marcaje."""
        notes = []
        try:
            revs, comments = int(doc.Revisions.Count), int(doc.Comments.Count)
        except Exception:  # noqa: BLE001
            return ["Nu am putut verifica dacă documentul .doc are modificări urmărite sau comentarii."]
        if fmt == "pdf" and (revs or comments):
            notes.append(f"Documentul .doc are {revs} modificări urmărite și {comments} comentarii. În PDF-ul "
                         "convertit pot apărea ca marcaje (inclusiv text șters) - verificați PDF-ul redactat "
                         "sau alegeți conversia în DOCX.")
        return notes

    def _quit(self) -> None:
        """Închide Word și AȘTEAPTĂ ieșirea procesului: Word își poate scrie opțiunile în registry la închidere,
        deci restaurarea registry-ului trebuie făcută după."""
        if self._word is None:
            return
        pid = self._pid
        try:
            self._word.Quit(SaveChanges=_WD_DO_NOT_SAVE)
        except Exception:  # noqa: BLE001
            log.warning("Word nu s-a închis normal")
        self._word, self._pid = None, None
        if pid is None:
            self._wait(2.0)                    # fără PID: doar o pauză scurtă
            return
        waited = 0.0
        while pid in self._list_pids() and waited < QUIT_WAIT_S:
            self._wait(0.25)
            waited += 0.25
        if pid in self._list_pids():
            log.warning("Word nu s-a închis în %ss; îl opresc", QUIT_WAIT_S)
            self._kill(pid)

    def __exit__(self, *exc) -> None:
        self._quit()
        if self._reg is not None:
            self._reg.__exit__(None, None, None)
            self._reg = None
        self._uninit()

    def _uninit(self) -> None:
        if self._com_init:
            import pythoncom
            pythoncom.CoUninitialize()
            self._com_init = False
