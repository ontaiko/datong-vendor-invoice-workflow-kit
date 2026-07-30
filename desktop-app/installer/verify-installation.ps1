param(
    [string]$InstallDir = ""
)

$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT = "0"

if ([string]::IsNullOrWhiteSpace($InstallDir)) {
    $appName = [Text.Encoding]::UTF8.GetString(
        [Convert]::FromBase64String("5aSn57Wx6YCy6LKo5Yqp5omL")
    )
    $InstallDir = Join-Path (Join-Path $env:LOCALAPPDATA "Programs") $appName
}
$InstallDir = [System.IO.Path]::GetFullPath($InstallDir)
$exeName = [Text.Encoding]::UTF8.GetString(
    [Convert]::FromBase64String("5aSn57Wx6YCy6LKo5Yqp5omLLmV4ZQ==")
)
$appExe = Join-Path $InstallDir $exeName
$venvPython = Join-Path $InstallDir "engine\.venv\Scripts\python.exe"
$manifestPath = Join-Path $InstallDir "package-manifest.json"

$required = @(
    $appExe,
    $venvPython,
    $manifestPath,
    (Join-Path $InstallDir "scripts\local_paddleocr_invoice_to_xlsx.py"),
    (Join-Path $InstallDir "scripts\match-existing-products.py"),
    (Join-Path $InstallDir "scripts\review-invoice-product-check.py"),
    (Join-Path $InstallDir "scripts\fill-import-templates.ps1")
)

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($file in $manifest.files) {
    if ([string]$file.path -like "reference_data/*") {
        $relative = ([string]$file.path).Replace("/", [System.IO.Path]::DirectorySeparatorChar)
        $required += Join-Path $InstallDir $relative
    }
}

$missing = @($required | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
if ($missing.Count -gt 0) {
    $missing | ForEach-Object { Write-Host "Missing: $_" }
    throw "Installed APP files are incomplete."
}

& $venvPython -X utf8 -m pip check
if ($LASTEXITCODE -ne 0) {
    throw "OCR environment dependency check failed."
}
& $venvPython -X utf8 -c "import cv2, numpy, openpyxl, paddle, paddleocr, rapidfuzz; from PIL import Image; print('OCR_RUNTIME_OK')"
if ($LASTEXITCODE -ne 0) {
    throw "OCR runtime import test failed."
}
& $appExe --self-test
if ($LASTEXITCODE -ne 0) {
    throw "APP self-test failed."
}

Write-Host "Installation verification complete: $InstallDir"

