"""Etapa 2: verificare cu apariții debifate, nume de ieșire unice, raport pentru .doc."""
import csv
import os

import pytest

import fixtures as fx
from pii_core import engine
from pii_core.models import Confidence, Finding


@pytest.fixture
def docx_file(tmp_path):
    p = tmp_path / "contract.docx"
    fx.make_docx(str(p))
    return str(p)


def test_partial_deselect_no_false_alarm(docx_file, tmp_path):
    res = engine.scan_file(docx_file)
    occ = [f for f in res.findings if f.value == fx.NAME_A]
    assert len(occ) == 2
    occ[1].selected = False                       # semnătura rămâne, preambulul se redactează
    assert engine.norm_value(fx.NAME_A) in engine.mixed_values(res)
    out, problems = engine.redact_file(res, str(tmp_path / "out"))
    assert problems == []                         # înainte: alarmă falsă
    import docx
    paras = [p.text for p in docx.Document(out).paragraphs]
    assert sum(t.count(fx.NAME_A) for t in paras) == 1
    assert next(t for t in paras if t.startswith("Între:")).startswith("Între: [REDACTAT:NUME]")


def test_verify_still_catches_real_leak():
    f = Finding(id=1, file="x", kind="CNP", value=fx.CNP_A, confidence=Confidence.MARE,
                       location="p", context="")
    assert engine.verify(f"text {fx.CNP_A} text", [f], set())
    assert engine.verify(f"text {fx.CNP_A} text", [f], {"altceva"})
    assert engine.verify(f"text {fx.CNP_A} text", [f], {engine.norm_value(fx.CNP_A)}) == []


def test_output_names_unique(tmp_path):
    taken = set()
    a = engine.output_path(str(tmp_path / "d1" / "c.docx"), str(tmp_path / "o"), taken)
    b = engine.output_path(str(tmp_path / "d2" / "c.docx"), str(tmp_path / "o"), taken)
    c = engine.output_path(str(tmp_path / "d3" / "c.docx"), str(tmp_path / "o"), taken)
    assert len({a, b, c}) == 3
    assert b.endswith("c_REDACTAT_2.docx") and c.endswith("c_REDACTAT_3.docx")


def test_report_uses_source_name(docx_file, tmp_path):
    res = engine.scan_file(docx_file)
    res.source = str(tmp_path / "vechi.doc")
    rep = tmp_path / "r.csv"
    engine.write_report([res], str(rep))
    with open(rep, encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh, delimiter=";"))[1:]
    assert rows and all(r[0] == "vechi.doc" for r in rows)
    assert os.path.basename(docx_file) not in rep.read_text(encoding="utf-8-sig")
