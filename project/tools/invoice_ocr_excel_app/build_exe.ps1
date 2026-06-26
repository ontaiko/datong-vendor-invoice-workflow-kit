param(
    [string]$WorkspaceRoot = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($WorkspaceRoot)) {
    $workspacePath = (Get-Item -LiteralPath (Join-Path $PSScriptRoot "..\..")).FullName
} else {
    $workspacePath = (Get-Item -LiteralPath $WorkspaceRoot).FullName
}
$toolDir = Join-Path $workspacePath "tools\invoice_ocr_excel_app"
$python = "C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe"
$entry = Join-Path $toolDir "invoice_ocr_excel_gui.py"
$distDir = Join-Path $toolDir "dist"
$buildDir = Join-Path $toolDir "build"
$specDir = $buildDir
$asciiName = "invoice_ocr_excel"
$targetName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("6YCy6LKo5Zau5ZyW54mH6L2JRXhjZWwuZXhl"))
$targetPath = Join-Path $toolDir $targetName

if (-not (Test-Path -LiteralPath $python)) {
    throw "System OCR Python was not found: $python"
}
if (-not (Test-Path -LiteralPath $entry)) {
    throw "GUI entry was not found: $entry"
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

Write-Host "Done: $targetPath"
