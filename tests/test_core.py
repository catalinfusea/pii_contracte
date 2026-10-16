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

import csv
import hashlib
import os
import sys
import zipfile

import pymupdf
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.dirname(__file__))

import fixtures as fx  # noqa: E402
from pii_core import engine  # noqa: E402
from pii_core import detectors as det  # noqa: E402
from pii_core.models import Confidence  # noqa: E402


def kinds_values(res):
    return {(f.kind, f.value) for f in res.findings}


def sha(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()


# ------------------------------------------------------------- detectoare
def test_cnp_valid_and_invalid():
    assert det.cnp_checksum_ok(fx.CNP_A)
    bad = fx.CNP_A[:-1] + str((int(fx.CNP_A[-1]) + 1) % 10)
    assert not det.cnp_checksum_ok(bad)
    ms = det.find_cnp(f"CNP {fx.CNP_A}")
    assert ms[0].confidence == Confidence.MARE
    # cifră de control greșită lângă ancoră -> raportat cu încredere mică
    assert det.find_cnp(f"CNP {bad}")[0].confidence == Confidence.MICA
    # fără ancoră și invalid -> ignorat
    assert det.find_cnp(f"Contract nr. {bad}") == []


def test_cnp_invalid_date_and_unknown_county():
    d = fx.gen_cnp(1, 80, 2, 31, 40, 1)        # 31 februarie
    assert det.find_cnp(d) == []
    u = fx.gen_cnp(1, 80, 2, 10, 60, 1)        # cod județ necunoscut
    assert det.find_cnp(u)[0].confidence == Confidence.MEDIE


def test_contract_number_not_cnp():
    assert not det.cnp_checksum_ok(fx.CONTRACT_NO)
    assert det.find_cnp(f"Contract nr. {fx.CONTRACT_NO}") == []


def test_iban():
    assert det.iban_ok(fx.IBAN_A)
    spaced = " ".join(fx.IBAN_A[i:i + 4] for i in range(0, 24, 4))
    assert det.find_iban(f"cont {spaced}, deschis")[0].value == spaced
    assert det.find_iban("RO00BTRL0000123456789012") == []


def test_iban_variants():
    v = fx.IBAN_A
    nbsp = "\u00a0".join(v[i:i + 4] for i in range(0, 24, 4))
    dash = "-".join(v[i:i + 4] for i in range(0, 24, 4))
    for form in [nbsp, dash, v.lower(), "ro" + v[2:].lower()]:
        ms = det.find_iban(f"IBAN: {form}, deschis")
        assert ms and ms[0].value == form and ms[0].confidence == Confidence.MARE, form
    bad = v[:-1] + ("0" if v[-1] != "0" else "1")
    assert det.find_iban(f"cont IBAN {bad}")[0].confidence == Confidence.MICA
    assert det.find_iban(f"cod {bad}") == []          # fără ancoră: ignorat


def test_contact_and_id():
    t = f"tel. {fx.PHONE_A}, +40 21 123 4567, email {fx.EMAIL_A}, CI seria RX nr. 123456"
    kv = {(m.kind, m.value) for m in det.detect(t)}
    assert ("Telefon", fx.PHONE_A) in kv
    assert ("Telefon", "+40 21 123 4567") in kv
    assert ("Email", fx.EMAIL_A) in kv
    assert ("Serie act identitate", "RX") in kv and ("Număr act identitate", "123456") in kv


def test_names_and_address():
    kv = {(m.kind, m.value) for m in det.detect(fx.preamble() + " " + fx.preamble_b())}
    assert ("Nume", fx.NAME_A) in kv
    assert ("Nume", fx.NAME_B) in kv
    assert ("Adresă", "Mun. București, Str. Florilor nr. 5, bl. A2, sc. 1, ap. 14, Sector 3") in kv
    assert ("Adresă", "Jud. Ilfov, Com. Otopeni, Str. Mare nr. 2") in kv


def test_no_false_positives_on_clean_text():
    t = ("Prețul este de 150.000 lei, plătibil în 12 rate. Art. 5 alin. (2). Contract nr. "
         f"{fx.CONTRACT_NO}. Părțile au convenit. Vânzătorul garantează.")
    assert det.detect(t) == []


# ------------------------------------------------------------- DOCX
@pytest.fixture
def docx_file(tmp_path):
    p = tmp_path / "contract.docx"
    fx.make_docx(str(p))
    return str(p)


def test_docx_scan(docx_file):
    res = engine.scan_file(docx_file)
    assert res.error is None
    kv = kinds_values(res)
    for exp in [("CNP", fx.CNP_A), ("CNP", fx.CNP_B), ("IBAN", fx.IBAN_A), ("Email", fx.EMAIL_A),
                ("Email", fx.EMAIL_B), ("Telefon", fx.PHONE_A), ("Telefon", "0733 444 555"),
                ("Nume", fx.NAME_A), ("Nume", fx.NAME_B)]:
        assert exp in kv, exp
    # CNP-ul din textul șters (track changes) nu e o constatare - dispare oricum
    assert ("CNP", fx.CNP_OLD) not in kv
    # propagare: numele apare de 2 ori (preambul + semnătură)
    assert sum(1 for f in res.findings if f.value == fx.NAME_A) == 2
    # linkul mailto
    assert any(f.location == "hyperlink" for f in res.findings)
    assert any("comentarii" in w for w in res.warnings)
    assert any("modificări urmărite" in w for w in res.warnings)


def test_docx_redact(docx_file, tmp_path):
    before = sha(docx_file)
    res = engine.scan_file(docx_file)
    out, problems = engine.redact_file(res, str(tmp_path / "out"))
    assert problems == []
    assert sha(docx_file) == before, "originalul a fost modificat"
    with zipfile.ZipFile(out) as z:
        blob = b"".join(z.read(n) for n in z.namelist()).decode("utf-8", "ignore")
    for v in [fx.CNP_A, fx.CNP_B, fx.CNP_OLD, fx.IBAN_A, fx.EMAIL_A, fx.EMAIL_B, fx.NAME_A,
              "0733 444 555", "Popescu Ion", "Florilor"]:
        assert v not in blob, v
    assert "[REDACTAT:CNP]" in blob
    assert "Clauză fără date personale" in blob       # textul obișnuit rămâne
    assert "Contract nr. " + fx.CONTRACT_NO in blob    # nr. de contract nu e redactat
    assert "w:del " not in blob and "<w:del>" not in blob
    # se deschide în python-docx (structură validă)
    import docx
    docx.Document(out)


def test_docx_deselect_and_manual(docx_file, tmp_path):
    res = engine.scan_file(docx_file, manual_terms=["1.000 lei"])
    assert any(f.kind == "Manual" for f in res.findings)
    for f in res.findings:
        if f.kind == "IBAN":
            f.selected = False
    out, problems = engine.redact_file(res, str(tmp_path / "out"))
    assert problems == []
    text = engine.dx.all_text(out)
    assert fx.IBAN_A in text           # debifat -> rămâne
    assert "1.000 lei" not in text     # adăugat manual -> redactat


def test_docx_xxe_not_resolved(docx_file, tmp_path):
    secret = tmp_path / "secret.txt"
    secret.write_text("MARKERXXE7731", encoding="utf-8")
    bad = str(tmp_path / "xxe.docx")
    fx.make_xxe_docx(docx_file, bad, str(secret))
    res = engine.scan_file(bad)
    assert res.error is None
    assert all("MARKERXXE7731" not in f.context for f in res.findings)
    assert "MARKERXXE7731" not in engine.dx.all_text(bad)


def test_bad_docx(tmp_path):
    p = tmp_path / "x.docx"
    p.write_bytes(b"not a zip")
    assert "nu e un .docx valid" in engine.scan_file(str(p)).error


# ------------------------------------------------------------- PDF
@pytest.fixture
def pdf_file(tmp_path):
    p = tmp_path / "contract.pdf"
    fx.make_pdf(str(p))
    return str(p)


def test_pdf_scan(pdf_file):
    res = engine.scan_file(pdf_file)
    assert res.error is None
    kv = kinds_values(res)
    for exp in [("CNP", fx.CNP_A), ("CNP", fx.CNP_B), ("IBAN", fx.IBAN_A), ("Email", fx.EMAIL_B),
                ("Nume", fx.NAME_A), ("Nume", fx.NAME_B), ("Telefon", fx.PHONE_A)]:
        assert exp in kv, exp
    assert any("probabil scanate): 2" in w for w in res.warnings)


def test_pdf_redact(pdf_file, tmp_path):
    before = sha(pdf_file)
    res = engine.scan_file(pdf_file)
    out, problems = engine.redact_file(res, str(tmp_path / "out"))
    assert problems == []
    assert sha(pdf_file) == before
    doc = pymupdf.open(out)
    text = engine.pd.normalize_text("".join(p.get_text() for p in doc))   # NBSP/U+00AD (Arial pe Windows)
    links = [lk.get("uri") for p in doc for lk in p.get_links()]
    meta = doc.metadata
    doc.close()
    for v in [fx.CNP_A, fx.CNP_B, fx.IBAN_A, fx.EMAIL_B, "POPESCU", "Florilor", "0722"]:
        assert v not in text, v
    assert not any(fx.EMAIL_B in (u or "") for u in links)
    assert not meta.get("author") and not meta.get("title")
    assert "Clauză fără date personale" in text
    assert fx.CONTRACT_NO in text


def test_pdf_encrypted(tmp_path):
    p = str(tmp_path / "enc.pdf")
    fx.make_encrypted_pdf(p)
    assert "parolă" in engine.scan_file(p).error


def test_unsupported(tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("x")
    assert "nesuportat" in engine.scan_file(str(p)).error


# ------------------------------------------------------------- raport
def test_report_masked(docx_file, pdf_file, tmp_path):
    results = [engine.scan_file(docx_file), engine.scan_file(pdf_file)]
    rp = tmp_path / "raport.csv"
    engine.write_report(results, str(rp))
    content = rp.read_text(encoding="utf-8-sig")
    for v in [fx.CNP_A, fx.IBAN_A, fx.EMAIL_A, fx.NAME_A, "Florilor"]:
        assert v not in content, v
    rows = list(csv.reader(content.splitlines(), delimiter=";"))
    assert rows[0] == engine.report_columns()
    assert len(rows) > 10


def test_never_overwrites_original(tmp_path):
    src = str(tmp_path / "a.pdf")
    assert engine.output_path(src, str(tmp_path)) != src
