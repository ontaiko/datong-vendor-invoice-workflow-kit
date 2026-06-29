param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [Parameter(Mandatory = $true)]
    [string]$PackageRoot,

    [switch]$SkipModelWarmup
)

$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT = "0"

function Get-Python312 {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"),
        "C:\Program Files\Python312\python.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Get-Item -LiteralPath $candidate).FullName
        }
    }

    $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        $resolved = & py -3.12 -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($resolved)) {
            return $resolved.Trim()
        }
    }

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $version = & $pythonCmd.Source -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>$null
        if ($LASTEXITCODE -eq 0 -and $version.Trim() -eq "3.12") {
            return $pythonCmd.Source
        }
    }

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "Python 3.12 was not found, and winget is not available. Install Python 3.12 first."
    }

    & winget install --id Python.Python.3.12 -e --scope user --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        throw "Python 3.12 installation failed."
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Get-Item -LiteralPath $candidate).FullName
        }
    }

    throw "Python 3.12 was installed, but python.exe could not be resolved."
}

$venvDir = Join-Path $ProjectRoot ".venv-paddleocr"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$requirements = Join-Path $PackageRoot "engine\requirements-ocr.txt"

if (-not (Test-Path -LiteralPath $requirements)) {
    throw "Missing requirements file: $requirements"
}

if (-not (Test-Path -LiteralPath $venvPython)) {
    $python = Get-Python312
    & $python -X utf8 -m venv $venvDir
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to create OCR virtual environment."
    }
}

& $venvPython -X utf8 -m pip install --upgrade pip setuptools wheel
if ($LASTEXITCODE -ne 0) {
    throw "Failed to upgrade pip."
}

& $venvPython -X utf8 -m pip install -r $requirements
if ($LASTEXITCODE -ne 0) {
    throw "Failed to install OCR requirements."
}

$modelZip = Join-Path $PackageRoot "engine\official_models.zip"
if (Test-Path -LiteralPath $modelZip) {
    $modelRoot = Join-Path $env:USERPROFILE ".paddlex\official_models"
    New-Item -ItemType Directory -Force -Path $modelRoot | Out-Null
    Expand-Archive -LiteralPath $modelZip -DestinationPath $modelRoot -Force
}

& $venvPython -X utf8 -m pip check
if ($LASTEXITCODE -ne 0) {
    throw "OCR environment dependency check failed."
}

& $venvPython -X utf8 -c "import cv2; import numpy; import openpyxl; import paddle; import paddleocr; import rapidfuzz; from PIL import Image; print('Local engines OK')"
if ($LASTEXITCODE -ne 0) {
    throw "OCR engine validation failed."
}

if (-not $SkipModelWarmup) {
    & $venvPython -X utf8 -c "from paddleocr import PaddleOCR; PaddleOCR(lang='ch', use_doc_orientation_classify=False, use_doc_unwarping=False, use_textline_orientation=False); print('OCR models ready')"
    if ($LASTEXITCODE -ne 0) {
        throw "OCR model download or initialization failed."
    }
}

Write-Host "OCR engine installed: $venvPython"
