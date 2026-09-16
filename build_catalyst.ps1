# Assembles a clean Catalyst AppSail build folder containing ONLY the runtime
# files (no .venv, .git, sample docs). Point the AppSail "build path" at ./catalyst_build.
#
# Usage:
#   .\build_catalyst.ps1              # copy app files only (managed runtime installs requirements.txt)
#   .\build_catalyst.ps1 -Vendor      # ALSO pip-install deps into the folder (fallback if the
#                                      # build does not auto-install requirements.txt)

param([switch]$Vendor)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$dest = Join-Path $root "catalyst_build"

if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
New-Item -ItemType Directory -Path $dest | Out-Null

# Runtime files the app actually needs.
Copy-Item (Join-Path $root "app.py")           $dest
Copy-Item (Join-Path $root "url_scraper.py")   $dest
Copy-Item (Join-Path $root "requirements.txt") $dest
Copy-Item (Join-Path $root "templates") (Join-Path $dest "templates") -Recurse

if ($Vendor) {
    Write-Host "Vendoring dependencies into catalyst_build ..."
    $py = Join-Path $root ".venv\Scripts\python.exe"
    if (-not (Test-Path $py)) { $py = "python" }
    & $py -m pip install -r (Join-Path $dest "requirements.txt") -t $dest --upgrade
}

Write-Host "Build folder ready: $dest"
Get-ChildItem $dest | Select-Object Name
