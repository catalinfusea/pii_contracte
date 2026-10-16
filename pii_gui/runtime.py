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

APP_DIR_NAME = "pii_contracte"
LOG_NAME = "pii_contracte.log"
SELFTEST_NAME = "autotest.txt"


def data_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".local", "share")
    d = os.path.join(base, APP_DIR_NAME)
    os.makedirs(d, exist_ok=True)
    return d


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
            lines.append(f"[OK]     {name}" + (f": {detail}" if detail else ""))
        except Exception as e:  # noqa: BLE001
            ok = False
            lines.append(f"[EROARE] {name}: {type(e).__name__}: {e}")
            logging.getLogger("pii_core").exception("Autotest: %s", name)

    from pii_core import __version__
    lines.append(f"pii_contracte {__version__} | Python {platform.python_version()} | "
                 f"{platform.system()} {platform.release()} | "
                 f"{'pachet .exe' if getattr(sys, 'frozen', False) else 'din sursă'}")
    cnp = _cnp(1, 80, 1, 15, 40, 123)
    text = f"Subsemnatul POPESCU ION, domiciliat în Mun. București, Str. Testului nr. 1, CNP {cnp}."
    tmp = tempfile.mkdtemp(prefix="pii_autotest_")

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
            raise RuntimeError(f"CNP nedetectat (eroare: {res.error})")
        out, problems = engine.redact_file(res, os.path.join(tmp, "out"))
        if problems or cnp in engine.pd.all_text(out):
            raise RuntimeError(f"redactare incompletă: {problems}")
        return f"pymupdf {pymupdf.__version__}"

    def t_docx():
        import lxml.etree
        from pii_core import engine
        src = os.path.join(tmp, "test.docx")
        _make_docx(src, text)
        res = engine.scan_file(src)
        if res.error or not any(f.kind == "CNP" for f in res.findings):
            raise RuntimeError(f"CNP nedetectat (eroare: {res.error})")
        out, problems = engine.redact_file(res, os.path.join(tmp, "out"))
        if problems or cnp in engine.dx.all_text(out):
            raise RuntimeError(f"redactare incompletă: {problems}")
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
            raise RuntimeError("pywin32 lipsește - .doc și exportul prin Word nu vor merge")
        src = os.path.join(tmp, "out", "test_REDACTAT.docx")
        dst = os.path.join(tmp, "out", "test_export.pdf")
        with doc_convert.WordConverter(timeout=60) as wc:
            wc.export(src, "pdf", dst)
        if not os.path.getsize(dst):
            raise RuntimeError("Word nu a produs PDF-ul")
        return "Word pornit, export DOCX→PDF, închis"

    check("PDF: detecție + redactare + verificare", t_pdf)
    check("DOCX: detecție + redactare + verificare", t_docx)
    if with_tk:
        check("Interfață (Tcl/Tk)", t_tk)
    if with_word and sys.platform == "win32":
        check("Microsoft Word (pywin32/COM)", t_word)
    elif with_word:
        lines.append("[SĂRIT]  Microsoft Word: doar pe Windows")
    import shutil
    shutil.rmtree(tmp, ignore_errors=True)
    lines.append("REZULTAT: " + ("TOTUL OK" if ok else "EXISTĂ ERORI - vezi liniile [EROARE] și logul"))
    return ok, lines


def run_selftest_cli() -> int:
    """Pentru `pii_contracte.exe --selftest`: scrie raportul în %LOCALAPPDATA% și îl afișează."""
    setup_logging()
    ok, lines = selftest()
    path = os.path.join(data_dir(), SELFTEST_NAME)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    if sys.stdout is not None:
        print("\n".join(lines))
        print(f"\nRaport: {path}")
    else:                                            # .exe fără consolă: fereastră cu rezultatul
        try:
            import tkinter
            from tkinter import messagebox
            r = tkinter.Tk()
            r.withdraw()
            (messagebox.showinfo if ok else messagebox.showerror)(
                "Autotest pii_contracte", "\n".join(lines) + f"\n\nRaport salvat în:\n{path}")
            r.destroy()
        except Exception:  # noqa: BLE001
            pass
    return 0 if ok else 1
