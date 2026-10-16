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

"""Varianta în engleză: catalog complet și consecvent; scanare, redactare, raport, autotest, limbă."""
import csv
import re
import string

import docx
import pytest

import fixtures as fx
from pii_core import engine, i18n
from pii_core.i18n import CATALOG, KINDS
from pii_gui import jobs, runtime

DIACRITICS = re.compile("[ăâîșțşţĂÂÎȘȚŞŢ]")


def fields(s):
    return sorted({f for _, f, _, _ in string.Formatter().parse(s) if f})


@pytest.mark.parametrize("key", sorted(CATALOG))
def test_catalog_pairs_consistent(key):
    ro, en = CATALOG[key]
    assert ro and en
    assert fields(ro) == fields(en), key
    assert not DIACRITICS.search(en), f"diacritice românești în textul englez: {key}"


def test_kind_labels_cover_all_detected_kinds(tmp_path):
    p = tmp_path / "c.docx"
    fx.make_docx(str(p))
    import test_detectors_v23 as v23
    from pii_core import detectors
    kinds = {f.kind for f in engine.scan_file(str(p), ["x"]).findings}
    for text in (v23.P1, v23.P2, v23.A1):
        kinds |= {m.kind for m in detectors.detect(text)}
    assert kinds - set(KINDS) == set()
    assert {"Loc naștere", "Număr permis conducere", "Carte funciară / cadastral"} <= kinds
    assert all(not DIACRITICS.search(en) for _, en in KINDS.values())


def test_scan_redact_report_in_english(tmp_path):
    i18n.set_lang("en")
    src = tmp_path / "contract.docx"
    fx.make_docx(str(src))
    res = engine.scan_file(str(src))
    assert any("tracked changes" in w for w in res.warnings)
    assert any(f.location.startswith("body, paragraph") for f in res.findings)
    assert any(f.note == "same value found elsewhere" for f in res.findings)
    out, problems = engine.redact_file(res, str(tmp_path / "o"))
    assert problems == [] and out.endswith("contract_REDACTED.docx")
    text = "\n".join(p.text for p in docx.Document(out).paragraphs)
    assert "[REDACTED:CNP]" in text and "[REDACTED:NAME]" in text and "REDACTAT" not in text
    rep = tmp_path / "r.csv"
    engine.write_report([res], str(rep))
    rows = list(csv.reader(rep.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))
    assert rows[0] == ["File", "Type", "Value (masked)", "Confidence", "Location", "Note", "Redacted"]
    body = rows[1:]
    assert any(r[1] == "WARNING" for r in body)
    assert any(r[1] == "Name" and r[3] == "medium" and r[6] == "yes" for r in body)
    assert fx.CNP_A not in rep.read_text(encoding="utf-8-sig")


def test_verify_message_in_english():
    from pii_core.models import Confidence, Finding
    i18n.set_lang("en")
    f = Finding(id=1, file="x", kind="Nume", value="Ion Pop", confidence=Confidence.MEDIE, location="", context="")
    assert engine.verify("text Ion Pop", [f]) == ["Name I** P** is still present in the redacted copy"]


def test_errors_and_folders_in_english(tmp_path):
    i18n.set_lang("en")
    bad = tmp_path / "x.docx"
    bad.write_bytes(b"no zip")
    assert "not a valid .docx" in engine.scan_file(str(bad)).error
    assert "Unsupported format" in engine.scan_file(str(tmp_path / "a.txt")).error
    assert jobs.default_out_dir(str(tmp_path / "a.pdf")).split("/")[-1].split("\\")[-1].startswith("redacted_")
    from pii_core import doc_convert
    assert doc_convert.target_path(str(tmp_path / "a.doc"), "pdf").replace("\\", "/").endswith("/Converted/a.pdf")


def test_selftest_in_english(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    i18n.set_lang("en")
    ok, lines = runtime.selftest(with_word=False, with_tk=False)
    assert ok and lines[-1] == "RESULT: ALL OK"
    assert any("PDF: detection + redaction + verification" in line for line in lines)


def test_resolve_lang_order(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.setattr(runtime, "system_lang", lambda: "en")
    assert runtime.resolve_lang([]) == "en"                       # limba sistemului
    runtime.save_settings({"lang": "ro"})
    assert runtime.resolve_lang([]) == "ro"                       # setarea salvată
    assert runtime.resolve_lang(["--lang", "en"]) == "en"         # argumentul are prioritate
    assert runtime.resolve_lang(["--lang=en", "a.pdf"]) == "en"
    assert runtime.resolve_lang(["--lang", "xx"]) == "ro"         # valoare invalidă -> următoarea sursă
    assert runtime.file_args(["--lang", "en", "a.pdf", "--selftest", "b.docx"]) == ["a.pdf", "b.docx"]


def test_system_lang_from_env(monkeypatch):
    if runtime.sys.platform == "win32":
        pytest.skip("pe Windows se citește limba interfeței, nu variabilele de mediu")
    for var in ("LC_ALL", "LC_MESSAGES"):
        monkeypatch.delenv(var, raising=False)
    monkeypatch.setenv("LANG", "ro_RO.UTF-8")
    assert runtime.system_lang() == "ro"
    monkeypatch.setenv("LANG", "de_DE.UTF-8")
    assert runtime.system_lang() == "en"
