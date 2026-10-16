import pytest

from pii_core.engine import mask


@pytest.mark.parametrize("kind,value,expected", [
    ("Data nașterii", "05.04.1975", "**.**.****"),
    ("Dată act", "10/05/2015", "**/**/****"),
    ("CNP", "2850412080319", "***********19"),
    ("IBAN", "RO49 AAAA 1B31 0075 9384 0000", "*" * 22 + "00"),
    ("Telefon", "0722 123 456", "********56"),
    ("Număr act identitate", "123456", "****56"),
    ("Serie act identitate", "RX", "**"),
    ("Nume", "POPESCU ION-MARIAN", "P****** I*********"),
    ("Email", "ion.popescu@example.com", "i***@example.com"),
])
def test_mask(kind, value, expected):
    assert mask(value, kind) == expected


def test_dates_reveal_no_digits():
    assert not any(c.isdigit() for c in mask("12.04.1985", "Data nașterii"))
