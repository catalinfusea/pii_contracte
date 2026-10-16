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

"""Fiecare fișier sursă are nota de copyright și de licență (AGPL-3.0-or-later)."""
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent


def test_every_source_file_has_license_notice():
    missing = []
    for p in list(ROOT.rglob("*.py")) + list(ROOT.rglob("*.pyw")) + list(ROOT.glob("*.spec")) + [ROOT / "build.ps1"]:
        if any(part.startswith(".venv") or part in ("build", "dist", "__pycache__") for part in p.parts):
            continue
        head = p.read_text(encoding="utf-8")[:1500]
        if "Copyright (C) 2026" not in head or "SPDX-License-Identifier: AGPL-3.0-or-later" not in head:
            missing.append(str(p.relative_to(ROOT)))
    assert not missing, missing
