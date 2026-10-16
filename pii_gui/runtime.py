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

"""Infrastructură pentru rularea ca .exe (PyInstaller, mod fără consolă).

- În modul fără consolă sys.stdout/sys.stderr pot fi None: logul merge într-un fișier rotativ
  în %LOCALAPPDATA%\\pii_contracte\\ (fără valori detectate; poate conține căi de fișiere).
- --selftest verifică pe mașina țintă că toate componentele din pachet funcționează, pe date SINTETICE.
"""
from __future__ import annotations

import logging
import logging.handlers
import os
import platform
import sys
import tempfile
import traceback
import zipfile

import json

from pii_core.i18n import LANGS, set_lang, t

APP_DIR_NAME = "pii_contracte"
LOG_NAME = "pii_contracte.log"
SELFTEST_NAME = "autotest.txt"
SETTINGS_NAME = "settings.json"
_LANG_ROMANIAN = 0x18          # PRIMARYLANGID pentru română (ro-RO = LCID 0x0418)


def data_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".local", "share")
    d = os.path.join(base, APP_DIR_NAME)
    os.makedirs(d, exist_ok=True)
    return d


# ------------------------------------------------------------------ setări și limbă
def load_settings() -> dict:
    try:
        with open(os.path.join(data_dir(), SETTINGS_NAME), encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(settings: dict) -> None:
    try:
        with open(os.path.join(data_dir(), SETTINGS_NAME), "w", encoding="utf-8") as fh:
            json.dump(settings, fh, ensure_ascii=False, indent=1)
    except OSError:
        logging.getLogger("pii_core").warning("Nu pot salva setările")


def system_lang() -> str:
    """Limba interfeței Windows (română -> "ro", altfel "en"). În afara Windows: variabilele de locale."""
    try:
        if sys.platform == "win32":
            import ctypes
            langid = ctypes.windll.kernel32.GetUserDefaultUILanguage()
            return "ro" if (langid & 0x3FF) == _LANG_ROMANIAN else "en"
        loc = os.environ.get("LC_ALL") or os.environ.get("LC_MESSAGES") or os.environ.get("LANG") or ""
        return "ro" if loc.lower().startswith("ro") else "en"
    except Exception:  # noqa: BLE001
        return "ro"


def _lang_arg(argv: list[str]) -> str | None:
    for i, a in enumerate(argv):
        if a.startswith("--lang="):
            return a.split("=", 1)[1]
        if a == "--lang" and i + 1 < len(argv):
            return argv[i + 1]
    return None


def resolve_lang(argv: list[str]) -> str:
    """Ordinea: --lang ro|en, apoi setarea salvată, apoi limba Windows."""
    for cand in (_lang_arg(argv), load_settings().get("lang"), system_lang()):
        if cand in LANGS:
            return cand
    return "ro"


def file_args(argv: list[str]) -> list[str]:
    """Argumentele care sunt fișiere (fără opțiuni --x și fără valoarea lui --lang)."""
    out, skip = [], False
    for a in argv:
        if skip:
            skip = False
            continue
        if a == "--lang":
            skip = True
            continue
        if not a.startswith("--"):
            out.append(a)
    return out


def setup_logging(level: int = logging.INFO) -> str:
    """Log în fișier (1 MB x 3) + consolă dacă există. Întoarce calea logului."""
    path = os.path.join(data_dir(), LOG_NAME)
    root = logging.getLogger()
    root.setLevel(level)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    if not any(isinstance(h, logging.handlers.RotatingFileHandler) for h in root.handlers):
        fh = logging.handlers.RotatingFileHandler(path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
        fh.setFormatter(fmt)
        root.addHandler(fh)
    if sys.stderr is not None and not any(type(h) is logging.StreamHandler for h in root.handlers):
        sh = logging.StreamHandler(sys.stderr)
        sh.setLevel(logging.WARNING)
        sh.setFormatter(fmt)
        root.addHandler(sh)

    def excepthook(t, v, tb):
        logging.getLogger("pii_core").critical("Excepție netratată:\n%s", "".join(traceback.format_exception(t, v, tb)))
    sys.excepthook = excepthook
    return path


# ------------------------------------------------------------------ autotest
def _cnp(s: int, yy: int, mm: int, dd: int, county: int, nnn: int) -> str:
    body = f"{s}{yy:02d}{mm:02d}{dd:02d}{county:02d}{nnn:03d}"
    t = sum(int(a) * int(b) for a, b in zip(body, "279146358279")) % 11
    return body + str(1 if t == 10 else t)


_DOCX_CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document.main+xml"/></Types>')
_DOCX_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
              'officeDocument" Target="word/document.xml"/></Relationships>')


def _make_docx(path: str, text: str) -> None:
    from xml.sax.saxutils import escape
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
           f'<w:p><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p></w:body></w:document>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _DOCX_CT)
        z.writestr("_rels/.rels", _DOCX_RELS)
        z.writestr("word/document.xml", doc)


def selftest(with_word: bool = True, with_tk: bool = True) -> tuple[bool, list[str]]:
    """Rulează verificările și întoarce (totul_ok, linii_raport). Date 100% sintetice."""
    lines: list[str] = []
    ok = True

    def check(name, fn):
        nonlocal ok
        try:
            detail = fn()
            lines.append(f"{t('st.ok')} {name}" + (f": {detail}" if detail else ""))
        except Exception as e:  # noqa: BLE001
            ok = False
            lines.append(f"{t('st.error')} {name}: {type(e).__name__}: {e}")
            logging.getLogger("pii_core").exception("Autotest: %s", name)

    from pii_core import __version__
    lines.append(f"pii_contracte {__version__} | Python {platform.python_version()} | "
                 f"{platform.system()} {platform.release()} | "
                 f"{t('st.frozen') if getattr(sys, 'frozen', False) else t('st.source')}")
    cnp = _cnp(1, 80, 1, 15, 40, 123)
    text = f"Subsemnatul POPESCU ION, domiciliat în Mun. București, Str. Testului nr. 1, CNP {cnp}."
    tmp = tempfile.mkdtemp(prefix="pii_autotest_")
    produced: dict[str, str] = {}

    def t_pdf():
        import pymupdf
        from pii_core import engine
        src = os.path.join(tmp, "test.pdf")
        d = pymupdf.open()
        d.new_page().insert_text((50, 80), text.replace("ă", "a").replace("ș", "s"), fontsize=9)
        d.save(src)
        d.close()
        res = engine.scan_file(src)
        if res.error or not any(f.kind == "CNP" for f in res.findings):
            raise RuntimeError(t("st.cnp_missing", err=res.error))
        out, problems = engine.redact_file(res, os.path.join(tmp, "out"))
        if problems or cnp in engine.pd.all_text(out):
            raise RuntimeError(t("st.incomplete", problems=problems))
        return f"pymupdf {pymupdf.__version__}"

    def t_docx():
        import lxml.etree
        from pii_core import engine
        src = os.path.join(tmp, "test.docx")
        _make_docx(src, text)
        res = engine.scan_file(src)
        if res.error or not any(f.kind == "CNP" for f in res.findings):
            raise RuntimeError(t("st.cnp_missing", err=res.error))
        out, problems = engine.redact_file(res, os.path.join(tmp, "out"))
        if problems or cnp in engine.dx.all_text(out):
            raise RuntimeError(t("st.incomplete", problems=problems))
        produced["docx"] = out
        return "lxml " + ".".join(map(str, lxml.etree.LXML_VERSION[:3]))

    def t_tk():
        import tkinter
        r = tkinter.Tk()
        r.withdraw()
        v = r.tk.call("info", "patchlevel")
        r.destroy()
        return f"Tcl/Tk {v}"

    def t_word():
        from pii_core import doc_convert
        if not doc_convert.available():
            raise RuntimeError(t("st.no_pywin32"))
        src = produced["docx"]                       # copia DOCX redactată de testul anterior
        dst = os.path.join(tmp, "out", "test_export.pdf")
        with doc_convert.WordConverter(timeout=60) as wc:
            wc.export(src, "pdf", dst)
        if not os.path.getsize(dst):
            raise RuntimeError(t("st.no_pdf"))
        return t("st.word_ok")

    check(t("st.pdf"), t_pdf)
    check(t("st.docx"), t_docx)
    if with_tk:
        check(t("st.tk"), t_tk)
    if with_word and sys.platform == "win32":
        check(t("st.word"), t_word)
    elif with_word:
        lines.append(t("st.word_skipped"))
    import shutil
    shutil.rmtree(tmp, ignore_errors=True)
    lines.append(t("st.all_ok") if ok else t("st.has_errors"))
    return ok, lines


def run_selftest_cli() -> int:
    """Pentru `pii_contracte.exe --selftest`: scrie raportul în %LOCALAPPDATA% și îl afișează."""
    setup_logging()
    set_lang(resolve_lang(sys.argv[1:]))
    ok, lines = selftest()
    path = os.path.join(data_dir(), SELFTEST_NAME)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    if sys.stdout is not None:
        print("\n".join(lines))
        print(f"\n{t('st.report_saved')} {path}")
    else:                                            # .exe fără consolă: fereastră cu rezultatul
        try:
            import tkinter
            from tkinter import messagebox
            r = tkinter.Tk()
            r.withdraw()
            (messagebox.showinfo if ok else messagebox.showerror)(
                t("st.title"), "\n".join(lines) + f"\n\n{t('st.report_saved')}\n{path}")
            r.destroy()
        except Exception:  # noqa: BLE001
            pass
    return 0 if ok else 1
