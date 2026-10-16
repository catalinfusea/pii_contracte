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
    assert sum(line.startswith("[EROARE]") for line in lines) == 2  # noqa: E501     # PDF și DOCX


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
