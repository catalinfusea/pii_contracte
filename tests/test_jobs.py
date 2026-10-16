import os
import queue
import shutil
import threading

import pytest

import fixtures as fx
from pii_core import doc_convert, engine
from pii_gui import jobs


def drain(q):
    out = []
    while True:
        try:
            out.append(q.get_nowait())
        except queue.Empty:
            return out


class FakeConverter:
    """Simulează Word: „convertește” copiind un .docx sintetic în Conversie\\."""
    entered = exited = 0

    def __init__(self, template):
        self.template = template

    def __enter__(self):
        FakeConverter.entered += 1
        return self

    def convert(self, src, fmt):
        dst = doc_convert.target_path(src, fmt)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy(self.template, dst)
        return dst

    def __exit__(self, *a):
        FakeConverter.exited += 1


class BrokenConverter:
    def __enter__(self):
        raise doc_convert.ConversionError("Microsoft Word nu poate fi pornit (test).")

    def __exit__(self, *a):
        raise AssertionError("nu trebuie apelat")


@pytest.fixture
def files(tmp_path):
    d = tmp_path / "in"
    d.mkdir()
    fx.make_docx(str(d / "a.docx"))
    fx.make_pdf(str(d / "b.pdf"))
    (d / "c.doc").write_bytes(b"\xd0\xcf\x11\xe0 binar vechi")   # conținutul nu contează pentru fake
    return d


def test_scan_job_with_doc(files, tmp_path):
    q, cancel = queue.Queue(), threading.Event()
    FakeConverter.entered = FakeConverter.exited = 0
    paths = [str(files / n) for n in ("a.docx", "b.pdf", "c.doc")]
    jobs.scan_job(paths, [], "docx", q, cancel, converter_factory=lambda: FakeConverter(paths[0]))
    msgs = drain(q)
    results = [m[1] for m in msgs if m[0] == "file"]
    assert msgs[-1][0] == "done" and msgs[-1][1]["files"] == 3
    assert FakeConverter.entered == FakeConverter.exited == 1
    doc = results[2]
    assert doc.error is None and doc.source.endswith("c.doc")
    assert doc.file == str(files / "Conversie" / "c.docx")
    assert doc.display_name == "c.doc"
    assert "Conversie" in doc.warnings[0]
    assert os.path.exists(files / "c.doc")                     # originalul rămâne


def test_scan_job_word_unavailable(files):
    q, cancel = queue.Queue(), threading.Event()
    paths = [str(files / "a.docx"), str(files / "c.doc")]
    jobs.scan_job(paths, [], "pdf", q, cancel, converter_factory=BrokenConverter)
    results = [m[1] for m in drain(q) if m[0] == "file"]
    assert results[0].error is None
    assert "Word nu poate fi pornit" in results[1].error


def test_no_converter_without_doc(files):
    q, cancel = queue.Queue(), threading.Event()
    jobs.scan_job([str(files / "a.docx")], [], "docx", q, cancel, converter_factory=BrokenConverter)
    assert drain(q)[-1][0] == "done"


def test_cancel(files):
    q, cancel = queue.Queue(), threading.Event()
    cancel.set()
    jobs.scan_job([str(files / "a.docx"), str(files / "b.pdf")], [], "docx", q, cancel)
    msgs = drain(q)
    assert not [m for m in msgs if m[0] == "file"]
    assert msgs[-1][1]["cancelled"] is True


def test_redact_job(files, tmp_path):
    other = tmp_path / "alt"
    other.mkdir()
    shutil.copy(files / "a.docx", other / "a.docx")              # același nume, alt folder
    paths = [str(files / "a.docx"), str(other / "a.docx"), str(files / "b.pdf")]
    results = [engine.scan_file(p) for p in paths]
    bad = tmp_path / "x.pdf"
    bad.write_bytes(b"nu e pdf")
    results.append(engine.scan_file(str(bad)))
    out = tmp_path / "out"
    q, cancel = queue.Queue(), threading.Event()
    jobs.redact_job(results, str(out), q, cancel)
    msgs = drain(q)
    done = msgs[-1][1]
    assert done["done"] == 3 and done["failed"] == 0 and done["problems"] == 0
    names = sorted(os.listdir(out))
    assert "a_REDACTAT.docx" in names and "a_REDACTAT_2.docx" in names and "b_REDACTAT.pdf" in names
    assert os.path.basename(done["report"]) in names
    for m in msgs:                                               # mesajele nu conțin valori
        assert fx.CNP_A not in repr(m) and fx.NAME_A not in repr(m)


def test_default_out_dir(tmp_path):
    import datetime as dt
    d = jobs.default_out_dir(str(tmp_path / "x.pdf"), dt.datetime(2026, 10, 6, 14, 5))
    assert d == os.path.join(str(tmp_path), "redactat_2026-10-06_1405")
