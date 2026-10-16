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

"""Punctul de intrare pentru .exe (PyInstaller) și pentru rularea din sursă.

    pii_contracte.exe                 interfața
    pii_contracte.exe f1.docx f2.pdf  interfața, cu fișierele deja adăugate
    pii_contracte.exe --selftest      autotest pe date sintetice (raport în %LOCALAPPDATA%\\pii_contracte)
"""
from pii_gui.app import main

if __name__ == "__main__":
    main()
