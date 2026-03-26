$ErrorActionPreference = "Stop"

$python = "C:\Users\ibrahimsaed\AppData\Local\Programs\Python\Python314\python.exe"
$entry = Get-ChildItem -File -Filter "*13 app.py" | Select-Object -First 1
$distPath = Join-Path $PSScriptRoot "release"
$workPath = Join-Path $PSScriptRoot "release_build"
$specPath = Join-Path $PSScriptRoot "release_spec"

if (-not (Test-Path $python)) {
    throw "Python 3.14 interpreter not found at: $python"
}

if (-not $entry) {
    throw "Could not find the latest app entry matching *13 app.py"
}

New-Item -ItemType Directory -Force -Path $distPath | Out-Null
New-Item -ItemType Directory -Force -Path $workPath | Out-Null
New-Item -ItemType Directory -Force -Path $specPath | Out-Null

Get-ChildItem -Force -Path $distPath -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name bedayti `
    --collect-all openpyxl `
    --collect-all webdriver_manager `
    --collect-submodules selenium `
    --distpath "$distPath" `
    --workpath "$workPath" `
    --specpath "$specPath" `
    "$($entry.FullName)"
