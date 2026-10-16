"""Convertorul Word cu pythoncom/win32com SIMULATE: verifică secvența de apeluri și setările de siguranță.
Comportamentul real al Word se verifică pe Windows (checklist din README)."""
import os
import sys
import threading
import types

import pytest

from pii_core import doc_convert as dc


class FakeDoc:
    def __init__(self, log, fail_save=False, revs=0):
        self.log, self.fail_save = log, fail_save
        self.Revisions = types.SimpleNamespace(Count=revs)
        self.Comments = types.SimpleNamespace(Count=0)

    def SaveAs2(self, **kw):
        self.log.append(("SaveAs2", kw))
        if self.fail_save:
            raise RuntimeError("com_error")
        with open(kw["FileName"], "wb") as fh:
            fh.write(b"PK-fake")

    def Close(self, **kw):
        self.log.append(("Close", kw))


class FakeWord:
    def __init__(self, log, fail_open=False, fail_save=False):
        object.__setattr__(self, "log", log)
        object.__setattr__(self, "Documents", types.SimpleNamespace(Open=self._open))
        object.__setattr__(self, "_fo", fail_open)
        object.__setattr__(self, "_fs", fail_save)

    def __setattr__(self, k, v):
        self.log.append(("set", k, v))

    def _open(self, **kw):
        self.log.append(("Open", kw))
        if OPTS.get("hang"):                       # dialog ascuns: apelul nu se întoarce până la kill
            if not OPTS["killed"].wait(5):
                raise AssertionError("watchdog-ul nu a oprit Word")
            raise RuntimeError("RPC server unavailable")
        if self._fo:
            raise RuntimeError("com_error")
        return FakeDoc(self.log, self._fs, OPTS.get("revs", 0))

    def Quit(self, **kw):
        self.log.append(("Quit", kw))


OPTS = {}


@pytest.fixture
def com(monkeypatch):
    log, opts = [], OPTS
    opts.clear()
    pc = types.ModuleType("pythoncom")
    pc.CoInitialize = lambda: log.append(("CoInitialize",))
    pc.CoUninitialize = lambda: log.append(("CoUninitialize",))
    client = types.ModuleType("win32com.client")

    def dispatch_ex(progid):
        log.append(("DispatchEx", progid))
        if opts.get("no_word"):
            raise RuntimeError("Invalid class string")
        return FakeWord(log, opts.get("fail_open"), opts.get("fail_save"))
    client.DispatchEx = dispatch_ex
    pkg = types.ModuleType("win32com")
    pkg.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pc)
    monkeypatch.setitem(sys.modules, "win32com", pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setattr(dc, "_sleep", lambda s: None)
    monkeypatch.setattr(dc, "list_word_pids", lambda: set())
    return log, opts


def names(log):
    return [e[0] if e[0] != "set" else f"set:{e[1]}" for e in log]


@pytest.mark.parametrize("fmt,code,ext", [("docx", 12, ".docx"), ("pdf", 17, ".pdf")])
def test_convert_ok(com, tmp_path, fmt, code, ext):
    log, _ = com
    src = tmp_path / "Contract vechi.doc"
    src.write_bytes(b"x")
    with dc.WordConverter() as wc:
        out = wc.convert(str(src), fmt)
    assert out == str(tmp_path / "Conversie" / ("Contract vechi" + ext)) and os.path.isfile(out)
    sets = {e[1]: e[2] for e in log if e[0] == "set"}
    assert sets == {"Visible": False, "DisplayAlerts": 0, "AutomationSecurity": 3}
    # setările de siguranță înainte de orice Open
    assert names(log).index("set:AutomationSecurity") < names(log).index("Open")
    open_kw = next(e[1] for e in log if e[0] == "Open")
    assert open_kw["ReadOnly"] is True and open_kw["AddToRecentFiles"] is False
    assert open_kw["ConfirmConversions"] is False and open_kw["Visible"] is False
    assert open_kw["PasswordDocument"]                     # parolă falsă -> eroare, nu prompt
    save_kw = next(e[1] for e in log if e[0] == "SaveAs2")
    assert save_kw["FileFormat"] == code
    assert names(log)[-3:] == ["Close", "Quit", "CoUninitialize"]
    assert src.read_bytes() == b"x"


def test_open_fails_still_quits(com, tmp_path):
    log, opts = com
    opts["fail_open"] = True
    src = tmp_path / "a.doc"
    src.write_bytes(b"x")
    with pytest.raises(dc.ConversionError, match="nu a putut deschide"):
        with dc.WordConverter() as wc:
            wc.convert(str(src), "docx")
    assert names(log)[-2:] == ["Quit", "CoUninitialize"]


def test_save_fails_closes_doc(com, tmp_path):
    log, opts = com
    opts["fail_save"] = True
    src = tmp_path / "a.doc"
    src.write_bytes(b"x")
    with pytest.raises(dc.ConversionError, match="nu a putut salva"):
        with dc.WordConverter() as wc:
            wc.convert(str(src), "pdf")
    assert names(log)[-3:] == ["Close", "Quit", "CoUninitialize"]


def test_word_missing(com, tmp_path):
    log, opts = com
    opts["no_word"] = True
    with pytest.raises(dc.ConversionError, match="Word nu poate fi pornit"):
        dc.WordConverter().__enter__()
    assert names(log) == ["CoInitialize", "DispatchEx", "CoUninitialize"]


def test_only_doc_accepted(com, tmp_path):
    with dc.WordConverter() as wc:
        with pytest.raises(dc.ConversionError):
            wc.convert(str(tmp_path / "a.docx"), "pdf")


def test_no_pywin32(monkeypatch):
    monkeypatch.setitem(sys.modules, "pythoncom", None)        # import -> ImportError
    assert dc.available() is False
    with pytest.raises(dc.ConversionError, match="pywin32"):
        dc.WordConverter().__enter__()


def test_revisions_warning_pdf_only(com, tmp_path):
    log, opts = com
    opts["revs"] = 3
    src = tmp_path / "a.doc"
    src.write_bytes(b"x")
    with dc.WordConverter() as wc:
        wc.convert(str(src), "docx")
        assert wc.last_notes == []
        wc.convert(str(src), "pdf")
        assert "3 modificări urmărite" in wc.last_notes[0]
    assert not [e for e in log if e[0] in ("AcceptAll", "set") and e[1:2] == ("Revisions",)]


@pytest.mark.parametrize("src_ext,fmt,code", [(".docx", "pdf", 17), (".pdf", "docx", 12)])
def test_export(com, tmp_path, src_ext, fmt, code):
    log, _ = com
    src = tmp_path / ("x_REDACTAT" + src_ext)
    src.write_bytes(b"x")
    dst = tmp_path / ("x_REDACTAT_2." + fmt)
    with dc.WordConverter() as wc:
        assert wc.export(str(src), fmt, str(dst)) == str(dst)
        assert wc.last_notes == []
    assert next(e[1] for e in log if e[0] == "SaveAs2")["FileFormat"] == code
    assert next(e[1] for e in log if e[0] == "Open")["ReadOnly"] is True
    assert not (tmp_path / "Conversie").exists()


def test_export_rejects_same_format(com, tmp_path):
    with dc.WordConverter() as wc:
        with pytest.raises(dc.ConversionError):
            wc.export(str(tmp_path / "a.pdf"), "pdf", str(tmp_path / "b.pdf"))


def test_tasklist_parsing(monkeypatch):
    out = '"WINWORD.EXE","1234","Console","1","150,000 K"\n"WINWORD.EXE","88","Console","1","9 K"\n'
    monkeypatch.setattr(dc.subprocess, "run", lambda *a, **k: types.SimpleNamespace(stdout=out))
    assert dc.list_word_pids() == {1234, 88}
    monkeypatch.setattr(dc.subprocess, "run",
                        lambda *a, **k: types.SimpleNamespace(stdout="INFO: No tasks are running."))
    assert dc.list_word_pids() == set()
