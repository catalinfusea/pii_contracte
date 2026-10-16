"""Caractere care la extragere din PDF nu sunt cele „vizibile” (ex. Arial pe Windows: cratima -> U+00AD)."""
import pymupdf
import pytest

import fixtures as fx
from pii_core import engine
from pii_core import pdf_handler as pd


def _pdf(path, text):
    d = pymupdf.open()
    d.new_page().insert_text((50, 80), text, fontname="f0", fontfile=fx.FONT, fontsize=10)
    d.save(path)
    d.close()


def test_normalization_is_one_to_one():
    for k in pd.CHAR_NORMALIZE:
        assert len(chr(k).translate(pd.CHAR_NORMALIZE)) == 1


@pytest.mark.parametrize("sep", ["\u00ad", "\u2011", "-"])
def test_hyphenated_name_detected_and_redacted(tmp_path, sep):
    if not pymupdf.Font(fontfile=fx.FONT).has_glyph(ord(sep)):
        pytest.skip(f"fontul de test ({fx.FONT}) nu are glifă pentru U+{ord(sep):04X}")
    src = str(tmp_path / "c.pdf")
    _pdf(src, f"Între: POPESCU ION{sep}MARIAN, domiciliat în Mun. București, Str. Florilor nr. 5, CNP {fx.CNP_A}")
    res = engine.scan_file(src)
    extracted = pd.all_text(src)
    assert "POPESCU ION-MARIAN" in extracted, f"text extras: {extracted!r}"
    assert ("Nume", "POPESCU ION-MARIAN") in {(f.kind, f.value) for f in res.findings}
    out, problems = engine.redact_file(res, str(tmp_path / "o"))
    assert problems == []
    doc = pymupdf.open(out)
    raw = "".join(p.get_text() for p in doc)
    doc.close()
    assert "POPESCU" not in raw and "MARIAN" not in raw and fx.CNP_A not in raw


def test_verify_sees_through_soft_hyphen(tmp_path):
    """Dacă redactarea ar rata valoarea, verificarea trebuie să o vadă și cu U+00AD."""
    src = str(tmp_path / "c.pdf")
    _pdf(src, "Text POPESCU ION\u00adMARIAN aici")
    assert "POPESCU ION-MARIAN" in pd.all_text(src)
