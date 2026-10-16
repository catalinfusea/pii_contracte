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
