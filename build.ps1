# Copyright (C) 2026 CatalinFusea - parte din pii_contracte
# SPDX-License-Identifier: AGPL-3.0-or-later
<#
Build pii_contracte.exe (onedir) pe Windows. Testat: NU (scris pe Linux; spec-ul validat cu build pe Linux).
Cerințe: Python 3.14 (py launcher), Word instalat doar pentru autotestul complet.

Pași: venv separat .venv-build -> instalare versiuni fixate -> teste -> PyInstaller -> autotest din .exe -> arhivă.
Uz (PowerShell, din folderul proiectului):
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    .\build.ps1
#>
$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

$py = ".\.venv-build\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Host "== Creez .venv-build (Python 3.14)"
    py -3.14 -m venv .venv-build
    if ($LASTEXITCODE -ne 0) { throw "Nu pot crea venv-ul (e instalat Python 3.14?)" }
}
& $py -VV

Write-Host "== Instalez dependențele fixate"
& $py -m pip install --disable-pip-version-check -q -r requirements-build.txt
if ($LASTEXITCODE -ne 0) { throw "pip install a eșuat" }

Write-Host "== Rulez testele"
& $py -m pytest -q -rs tests
if ($LASTEXITCODE -ne 0) { throw "Testele au eșuat - nu fac build" }

Write-Host "== PyInstaller"
if (Test-Path .\build) { Remove-Item -Recurse -Force .\build }
if (Test-Path .\dist\pii_contracte) { Remove-Item -Recurse -Force .\dist\pii_contracte }
& $py -m PyInstaller --noconfirm --clean pii_contracte.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller a eșuat" }

$exe = ".\dist\pii_contracte\pii_contracte.exe"
Write-Host "== Autotest din .exe (apare o fereastră cu rezultatul)"
$p = Start-Process -FilePath $exe -ArgumentList "--selftest" -Wait -PassThru
$report = Join-Path $env:LOCALAPPDATA "pii_contracte\autotest.txt"
if (Test-Path $report) { Get-Content $report }
if ($p.ExitCode -ne 0) { throw "Autotestul .exe a eșuat (cod $($p.ExitCode)) - vezi $report" }

$ver = & $py -c "import pii_core; print(pii_core.__version__)"
$zip = ".\dist\pii_contracte_$ver.zip"
if (Test-Path $zip) { Remove-Item $zip }
Compress-Archive -Path .\dist\pii_contracte -DestinationPath $zip
$hash = (Get-FileHash $zip -Algorithm SHA256).Hash
Write-Host "== Gata: $zip"
Write-Host "   SHA256: $hash"
