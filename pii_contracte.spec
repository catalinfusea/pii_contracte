# -*- mode: python ; coding: utf-8 -*-
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

# PyInstaller 6.22.3 - build "onedir", fără consolă. Rulați prin build.ps1 (Windows).
import sys
from importlib import metadata

# fișierele de licență ale dependențelor, copiate din pachetele instalate (nu rescrise de mână)
license_files = []
for dist_name in ("pymupdf", "lxml", "pywin32", "pyinstaller"):
    try:
        dist = metadata.distribution(dist_name)
    except metadata.PackageNotFoundError:
        continue
    for f in dist.files or []:
        low = str(f).lower()
        if ".dist-info" in low and ("licen" in low or "copying" in low):
            license_files.append((str(f.locate()), f"licente/{dist_name}"))

hidden = []
if sys.platform == "win32":
    hidden += ["win32com.client", "pythoncom", "pywintypes", "winreg"]   # importate în funcții (doc_convert)

a = Analysis(
    ["pii_contracte_app.py"],
    pathex=[],
    binaries=[],
    datas=[("LICENSE", "."), ("LICENSES.txt", "."), ("CITESTE-MA.txt", "."), ("README-USER.txt", ".")]
    + license_files,
    hiddenimports=hidden,
    hookspath=[],
    runtime_hooks=[],
    excludes=["pytest", "docx", "numpy", "PIL", "matplotlib", "IPython", "setuptools", "pip"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="pii_contracte",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                  # UPX crește rata de fals-pozitive la antivirus
    console=False,
    disable_windowed_traceback=False,
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, upx_exclude=[], name="pii_contracte")
