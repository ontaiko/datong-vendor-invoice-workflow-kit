param(
    [string]$WorkspaceRoot = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($WorkspaceRoot)) {
    $workspacePath = (Get-Item -LiteralPath (Join-Path $PSScriptRoot "..\..")).FullName
} else {
    $workspacePath = (Get-Item -LiteralPath $WorkspaceRoot).FullName
}
$toolDir = $PSScriptRoot
$pythonCandidates = @()
if (-not [string]::IsNullOrWhiteSpace($env:DATONG_PYTHON_EXE)) {
    $pythonCandidates += $env:DATONG_PYTHON_EXE
}
$pythonCandidates += (Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe")
$pathPython = Get-Command python.exe -ErrorAction SilentlyContinue
if ($null -ne $pathPython) {
    $pythonCandidates += $pathPython.Source
}
$python = $pythonCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
$entry = Join-Path $toolDir "invoice_ocr_excel_gui.py"
$distDir = Join-Path $toolDir "dist"
$buildDir = Join-Path $toolDir "build"
$specDir = $buildDir
$asciiName = "invoice_ocr_excel"
$targetName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("5aSn57Wx6YCy6LKo5Yqp5omLLmV4ZQ=="))
$targetPath = Join-Path $toolDir $targetName

if ([string]::IsNullOrWhiteSpace($python) -or -not (Test-Path -LiteralPath $python)) {
    throw "System OCR Python was not found. Set DATONG_PYTHON_EXE or install Python 3.12."
}
if (-not (Test-Path -LiteralPath $entry)) {
    throw "GUI entry was not found: $entry"
}

& $python -X utf8 (Join-Path $toolDir "scripts\generate-package-manifest.py")
if ($LASTEXITCODE -ne 0) {
    throw "Package manifest generation failed."
}

$pipShowExit = 1
try {
    & $python -X utf8 -m pip show pyinstaller 1>$null 2>$null
    $pipShowExit = $LASTEXITCODE
} catch {
    $pipShowExit = 1
}
if ($pipShowExit -ne 0) {
    & $python -X utf8 -m pip install pyinstaller
}

New-Item -ItemType Directory -Force -Path $distDir, $buildDir | Out-Null

& $python -X utf8 -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name $asciiName `
    --distpath $distDir `
    --workpath $buildDir `
    --specpath $specDir `
    $entry

if ($LASTEXITCODE -ne 0) {
    throw "EXE build failed."
}

$builtExe = Join-Path $distDir "$($asciiName).exe"
if (-not (Test-Path -LiteralPath $builtExe)) {
    throw "Built EXE was not found: $builtExe"
}

Copy-Item -LiteralPath $builtExe -Destination $targetPath -Force

& $targetPath --self-test
if ($LASTEXITCODE -ne 0) {
    throw "Built EXE self-test failed."
}

Write-Host "Done: $targetPath"
