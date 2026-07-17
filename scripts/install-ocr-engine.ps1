param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [Parameter(Mandatory = $true)]
    [string]$PackageRoot,

    [string]$ModelRoot = "",

    [switch]$SkipModelWarmup
)

$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT = "0"

if ([string]::IsNullOrWhiteSpace($ModelRoot)) {
    $ModelRoot = Join-Path $env:USERPROFILE ".paddlex\official_models"
}
$ModelRoot = [System.IO.Path]::GetFullPath($ModelRoot)
$env:PADDLE_PDX_CACHE_HOME = Split-Path -Parent $ModelRoot

function Assert-ModelFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)]$Manifest,
        [Parameter(Mandatory = $true)][string]$Label
    )

    foreach ($file in $Manifest.files) {
        $relative = ([string]$file.path).Replace("/", [System.IO.Path]::DirectorySeparatorChar)
        $path = Join-Path $Root $relative
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "$Label model file is missing: $path"
        }
        $item = Get-Item -LiteralPath $path
        if ($item.Length -ne [long]$file.bytes) {
            throw "$Label model file size mismatch: $path"
        }
        $actualHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actualHash -ne ([string]$file.sha256).ToLowerInvariant()) {
            throw "$Label model file hash mismatch: $path"
        }
    }
}

function Install-BundledModels {
    $modelSourceRoot = Join-Path $PackageRoot "engine\official_models"
    $manifestPath = Join-Path $PackageRoot "engine\model-manifest.json"
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Offline model manifest is missing: $manifestPath"
    }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-ModelFiles -Root $modelSourceRoot -Manifest $manifest -Label "Package"

    foreach ($file in $manifest.files) {
        $relative = ([string]$file.path).Replace("/", [System.IO.Path]::DirectorySeparatorChar)
        $source = Join-Path $modelSourceRoot $relative
        $destination = Join-Path $ModelRoot $relative
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
        Copy-Item -LiteralPath $source -Destination $destination -Force
    }
    Assert-ModelFiles -Root $ModelRoot -Manifest $manifest -Label "Installed"
    return $manifest
}

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

$modelManifest = Install-BundledModels

& $venvPython -X utf8 -m pip check
if ($LASTEXITCODE -ne 0) {
    throw "OCR environment dependency check failed."
}

& $venvPython -X utf8 -c "import cv2; import numpy; import openpyxl; import paddle; import paddleocr; import rapidfuzz; from PIL import Image; print('Local engines OK')"
if ($LASTEXITCODE -ne 0) {
    throw "OCR engine validation failed."
}

if (-not $SkipModelWarmup) {
    $modelNames = ($modelManifest.models | ConvertTo-Json -Compress)
    & $venvPython -X utf8 -c "import json; from paddlex import create_model; names=json.loads(r'''$modelNames'''); [create_model(model_name=name) for name in names]; from paddleocr import PaddleOCR; PaddleOCR(lang='ch', use_doc_orientation_classify=False, use_doc_unwarping=False, use_textline_orientation=False); print('Offline OCR models ready:', ', '.join(names))"
    if ($LASTEXITCODE -ne 0) {
        throw "Offline OCR model initialization failed."
    }
}

Write-Host "OCR engine installed: $venvPython"
Write-Host "Offline OCR models installed: $ModelRoot"
