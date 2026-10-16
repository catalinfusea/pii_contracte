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

import fixtures as fx
from pii_core import engine
from pii_gui.model import ReviewModel, is_link


@pytest.fixture
def model(tmp_path):
    p = tmp_path / "contract.docx"
    fx.make_docx(str(p))
    m = ReviewModel()
    m.add_result(engine.scan_file(str(p)))
    bad = tmp_path / "x.docx"
    bad.write_bytes(b"nu e zip")
    m.add_result(engine.scan_file(str(bad)))
    return m


def test_rows_and_counts(model):
    rows = model.rows()
    assert rows and all(r.file_index == 0 for r in rows)
    sel, total = model.counts()
    assert sel == total == len(rows)
    assert len(model.processable()) == 1


def test_toggle_and_get(model):
    iid = model.rows()[0].iid
    model.toggle(iid)
    assert model.get(iid).finding.selected is False
    model.toggle(iid)
    assert model.get(iid).finding.selected is True


def test_same_value(model):
    iid = next(r.iid for r in model.rows() if r.finding.value == fx.NAME_A)
    changed = model.set_same_value(iid, False)
    assert len(changed) == 2
    assert all(not model.get(i).finding.selected for i in changed)
    assert all(r.finding.selected for r in model.rows() if r.finding.value != fx.NAME_A)


def test_set_all_per_file(model):
    model.set_all(False, file_index=0)
    assert model.counts()[0] == 0
    model.set_all(True)
    assert model.counts()[0] == model.counts()[1]


def test_rescan_keeps_deselections(tmp_path):
    p = tmp_path / "c.docx"
    fx.make_docx(str(p))
    old = ReviewModel()
    old.add_result(engine.scan_file(str(p)))
    target = next(r for r in old.rows() if r.finding.kind == "IBAN")
    old.toggle(target.iid)
    new = ReviewModel()
    new.add_result(engine.scan_file(str(p), ["1.000 lei"]), old)   # rescanare cu termen manual
    iban = [r.finding for r in new.rows() if r.finding.kind == "IBAN"]
    assert iban and all(not f.selected for f in iban)
    assert all(f.selected for r in new.rows() for f in [r.finding] if f.kind != "IBAN")
    assert any(r.finding.kind == "Manual" for r in new.rows())


def test_link_rows_flagged(model):
    assert any(is_link(r.finding) for r in model.rows())
