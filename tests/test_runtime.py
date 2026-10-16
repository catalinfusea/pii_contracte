import logging

from pii_core import detectors
from pii_gui import runtime


def test_selftest_ok(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    ok, lines = runtime.selftest(with_word=False, with_tk=False)
    assert ok, lines
    assert lines[-1].endswith("TOTUL OK")


def test_selftest_reports_broken_detection(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.setattr(detectors, "ALL_DETECTORS", tuple(d for d in detectors.ALL_DETECTORS
                                                          if d is not detectors.find_cnp))
    ok, lines = runtime.selftest(with_word=False, with_tk=False)
    assert not ok
    assert sum(line.startswith("[EROARE]") for line in lines) == 2     # PDF și DOCX


def test_log_file_in_localappdata(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    root = logging.getLogger()
    before = list(root.handlers)
    try:
        path = runtime.setup_logging()
        logging.getLogger("pii_core").warning("mesaj de test")
        for h in root.handlers:
            h.flush()
        assert path.startswith(str(tmp_path)) and "mesaj de test" in open(path, encoding="utf-8").read()
    finally:
        for h in root.handlers[:]:
            if h not in before:
                root.removeHandler(h)
                h.close()
