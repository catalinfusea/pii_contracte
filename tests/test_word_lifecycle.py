"""Watchdog, închiderea Word și restaurarea registry-ului (COM și winreg SIMULATE)."""
import sys
import threading

import pytest

from pii_core import doc_convert as dc
from test_doc_convert import OPTS, com  # noqa: F401 - fixture reutilizat

OPT = r"Software\Microsoft\Office\16.0\Word\Options"
VAL = "DisableConvertPdfWarning"


class FakeProcs:
    """PID-uri WINWORD: fiecare DispatchEx pornește unul nou; după Quit procesul mai trăiește `quit_polls` interogări."""
    def __init__(self):
        self.alive, self.next, self.quit_polls = set(), 4242, 0

    def list_pids(self):
        if self.quit_polls:
            self.quit_polls -= 1
            if not self.quit_polls:
                self.alive.clear()
        return set(self.alive)


@pytest.fixture
def procs(com, monkeypatch):  # noqa: F811
    p = FakeProcs()
    client = sys.modules["win32com.client"]
    orig = client.DispatchEx

    def dispatch(progid):
        w = orig(progid)
        p.alive.add(p.next)
        p.next += 1
        return w
    monkeypatch.setattr(client, "DispatchEx", dispatch)
    return p


class NoReg:
    def __enter__(self):
        return self

    def __exit__(self, *a):
        pass


class FakeWinreg:
    HKEY_CURRENT_USER, KEY_READ, KEY_SET_VALUE, REG_DWORD, REG_SZ = 1, 1, 2, 4, 1

    def __init__(self):
        self.values = {}                                     # (cale cheie, nume) -> valoare
        self.deleted_keys = []
        self.created = set()
        self.subkeys = {r"Software\Microsoft\Office": ["16.0", "Common"],
                        r"Software\Microsoft\Office\16.0": ["Word"]}

    class Key:
        def __init__(self, path):
            self.path = path

        def __enter__(self):
            return self

        def __exit__(self, *a):
            pass

        def Close(self):
            pass

    def _path(self, root, sub):
        return root.path + "\\" + sub if isinstance(root, self.Key) else sub

    def OpenKey(self, root, sub, *a):
        path = self._path(root, sub)
        parent, _, leaf = path.rpartition("\\")
        if not (leaf in self.subkeys.get(parent, []) or path in self.subkeys or path in self.created
                or any(k[0] == path for k in self.values)):
            raise OSError(path)
        return self.Key(path)

    def CreateKeyEx(self, root, sub, *a):
        path = self._path(root, sub)
        while path:
            self.created.add(path)
            path = path.rpartition("\\")[0]
        return self.Key(self._path(root, sub))

    def EnumKey(self, k, i):
        kids = list(self.subkeys.get(k.path, []))
        kids += sorted({c[len(k.path) + 1:].split("\\")[0] for c in self.created
                        if c.startswith(k.path + "\\")} - set(kids))
        try:
            return kids[i]
        except IndexError:
            raise OSError

    def EnumValue(self, k, i):
        items = [(n, v) for (p, n), v in self.values.items() if p == k.path]
        if i >= len(items):
            raise OSError
        return items[i][0], items[i][1], 1

    def QueryValueEx(self, k, name):
        if (k.path, name) not in self.values:
            raise OSError
        return self.values[(k.path, name)], 4

    def SetValueEx(self, k, name, r, t, v):
        self.values[(k.path, name)] = v

    def DeleteKey(self, root, sub):
        if any(p == sub for p, _ in self.values) or any(c.startswith(sub + "\\") for c in self.created):
            raise OSError("cheie ne-goală")
        self.created.discard(sub)
        self.deleted_keys.append(sub)

    def DeleteValue(self, k, name):
        if (k.path, name) not in self.values:
            raise OSError
        del self.values[(k.path, name)]


def sup(reg):
    return dc.PdfWarningSuppressed(winreg_mod=reg, sleep=lambda s: None)


def test_watchdog_kills_hung_word_and_restarts(com, procs, tmp_path):  # noqa: F811
    OPTS["hang"], OPTS["killed"] = True, threading.Event()
    killed = []

    def kill(pid):
        killed.append(pid)
        procs.alive.discard(pid)
        OPTS["killed"].set()
    src = tmp_path / "x_REDACTAT.pdf"
    src.write_bytes(b"x")
    with dc.WordConverter(timeout=0.2, list_pids=procs.list_pids, kill=kill, registry=NoReg,
                          wait=lambda s: None) as wc:
        with pytest.raises(dc.ConversionError, match="nu a răspuns"):
            wc.export(str(src), "docx", str(tmp_path / "o.docx"))
        assert killed == [4242]
        OPTS["hang"] = False                                 # următorul fișier: Word repornit
        wc.export(str(src), "docx", str(tmp_path / "o.docx"))
        procs.quit_polls = 3                                 # se închide singur după Quit
    log = com[0]
    assert len([e for e in log if e[0] == "DispatchEx"]) == 2
    assert killed == [4242]


def test_quit_waits_then_kills_if_word_stays(com, procs, tmp_path):  # noqa: F811
    killed = []
    with dc.WordConverter(list_pids=procs.list_pids, kill=killed.append, registry=NoReg, wait=lambda s: None):
        pass                                                 # Word nu iese niciodată singur
    assert killed == [4242]


def test_registry_restored_only_after_word_exits(com, procs):  # noqa: F811
    """Word își scrie opțiunile (inclusiv valoarea =1) în momentul în care procesul se închide."""
    reg = FakeWinreg()

    def list_pids():
        before = bool(procs.alive)
        pids = procs.list_pids()
        if before and not pids:
            reg.values[(OPT, VAL)] = 1                       # scrierea de la ieșirea lui Word
        return pids
    with dc.WordConverter(list_pids=list_pids, kill=lambda p: None, wait=lambda s: None,
                          registry=lambda: sup(reg)):
        assert reg.values[(OPT, VAL)] == 1
        procs.quit_polls = 4
    for _ in range(10):                                      # „trece timpul”: Word termină de ieșit
        list_pids()
    assert reg.values == {}                                  # nu exista înainte -> șters, marker șters


def test_registry_set_and_restored():
    reg = FakeWinreg()
    with sup(reg):
        assert reg.values[(OPT, VAL)] == 1
        assert reg.values[(dc._MARKER_KEY, OPT)] == "<absent>"
    assert reg.values == {}
    assert reg.deleted_keys == [dc._MARKER_KEY, r"Software\pii_contracte"]   # fără urme în registry
    reg.values[(OPT, VAL)] = 0
    with sup(reg):
        assert reg.values[(OPT, VAL)] == 1
    assert reg.values == {(OPT, VAL): 0}                     # valoarea utilizatorului restaurată


def test_registry_recovery_after_crash():
    reg = FakeWinreg()
    sup(reg).__enter__()                                     # „crash”: fără __exit__
    assert reg.values[(OPT, VAL)] == 1
    assert sup(reg).recover() == 1
    assert reg.values == {}


def test_registry_recovery_runs_on_next_start():
    reg = FakeWinreg()
    reg.values[(OPT, VAL)] = 1                               # rămas de la versiunea 2.5
    reg.values[(dc._MARKER_KEY, OPT)] = "<absent>"
    with sup(reg):
        assert reg.values[(OPT, VAL)] == 1
    assert reg.values == {}


def test_registry_ignores_common_key():
    reg = FakeWinreg()
    with sup(reg):
        assert all("Common" not in p for p, _ in reg.values)
