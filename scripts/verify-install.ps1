param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [Parameter(Mandatory = $true)]
    [string]$CodexHome,

    [switch]$SkipEngine,
    [switch]$SkipExcelCheck
)

$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT = "0"

function Decode-Name {
    param([Parameter(Mandatory = $true)][string]$Base64)
    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Base64))
}

$skillNames = @(
    "convert-vendor-invoice-image",
    "extract-vendor-invoice-image",
    "match-product-catalog",
    "review-invoice-product-check",
    "build-inventory-import-files"
)

$referenceFolder = Decode-Name "5Y+D6ICD6LOH5paZ"
$requiredReferenceNames = @(
    (Decode-Name "T0NS6Kit5a6aLmpzb24="),
    (Decode-Name "5bu65qqU55SoLnhscw=="),
    (Decode-Name "5o6h6LO85Zau5Yyv5YWl56+E5L6LLnhscw=="),
    (Decode-Name "55Si5ZOB5q+U5bCN6Lqr5Lu96Zec6Y216KmeLmNzdg=="),
    (Decode-Name "55Si5ZOB6LOH5paZ6Ly45Ye6LkNTVg=="),
    (Decode-Name "5bug5ZWG5Luj6JmfLnhsc3g=")
)

$requiredFiles = @(
    (Join-Path $ProjectRoot "AGENTS.md"),
    ($requiredReferenceNames | ForEach-Object { Join-Path (Join-Path $ProjectRoot $referenceFolder) $_ })
)

foreach ($skillName in $skillNames) {
    $requiredFiles += Join-Path $CodexHome "skills\$skillName\SKILL.md"
}
$requiredFiles += Join-Path $CodexHome "skills\convert-vendor-invoice-image\scripts\check-product-csv-date.py"
$requiredFiles += Join-Path $CodexHome "skills\extract-vendor-invoice-image\scripts\local_paddleocr_invoice_to_xlsx.py"
$requiredFiles += Join-Path $CodexHome "skills\match-product-catalog\scripts\match-existing-products.py"
$requiredFiles += Join-Path $CodexHome "skills\review-invoice-product-check\scripts\review-invoice-product-check.py"
$requiredFiles += Join-Path $CodexHome "skills\build-inventory-import-files\scripts\fill-import-templates.ps1"

$missing = @($requiredFiles | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
if ($missing.Count -gt 0) {
    $missing | ForEach-Object { Write-Host "Missing: $_" }
    throw "Workflow package validation failed: required files are missing."
}

if (-not $SkipEngine) {
    $venvPython = Join-Path $ProjectRoot ".venv-paddleocr\Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
        throw "OCR virtual environment is missing: $venvPython"
    }
    & $venvPython -X utf8 -c "import cv2; import numpy; import openpyxl; import paddle; import paddleocr; import rapidfuzz; from PIL import Image; print('Python engines: OK')"
    if ($LASTEXITCODE -ne 0) {
        throw "Python engine validation failed."
    }
    & $venvPython -X utf8 -m pip check
    if ($LASTEXITCODE -ne 0) {
        throw "Python dependency validation failed."
    }
}

if (-not $SkipExcelCheck) {
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.DisplayAlerts = $false
        Write-Host "Excel COM: OK"
    } catch {
        Write-Warning "Excel COM is unavailable. OCR and matching can run, but final .xls import files require desktop Microsoft Excel."
    } finally {
        if ($null -ne $excel) {
            $excel.Quit()
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        }
    }
}

Write-Host "Five workflow skills: OK"
Write-Host "Project references: OK"
Write-Host "Installation verification complete."
