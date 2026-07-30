param(
    [Parameter(Mandatory = $true)]
    [string]$InstallDir,

    [string]$ModelRoot = ""
)

$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
$env:PIP_DISABLE_PIP_VERSION_CHECK = "1"
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT = "0"

$InstallDir = [System.IO.Path]::GetFullPath($InstallDir)
$engineDir = Join-Path $InstallDir "engine"
$venvDir = Join-Path $engineDir ".venv"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$requirements = Join-Path $engineDir "requirements-ocr.txt"
$manifestPath = Join-Path $engineDir "model-manifest.json"
$modelSourceRoot = Join-Path $engineDir "official_models"
$logPath = Join-Path $InstallDir "install-runtime.log"
$settingsPath = Join-Path $InstallDir "app_settings.json"

trap {
    $message = $_.Exception.ToString()
    try {
        Add-Content -LiteralPath $logPath -Value ("FATAL: " + $message) -Encoding UTF8
    } catch {
    }
    Write-Error $message
    exit 1
}

if ([string]::IsNullOrWhiteSpace($ModelRoot)) {
    $ModelRoot = Join-Path $env:USERPROFILE ".paddlex\official_models"
}
$ModelRoot = [System.IO.Path]::GetFullPath($ModelRoot)
$env:PADDLE_PDX_CACHE_HOME = Split-Path -Parent $ModelRoot

function Write-InstallLog {
    param([Parameter(Mandatory = $true)][string]$Message)
    $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$FailureMessage
    )
    Write-InstallLog ("Run: {0} {1}" -f $FilePath, ($Arguments -join " "))
    $quotedArguments = @(
        foreach ($argument in $Arguments) {
            if ($argument -match '[\s"]') {
                '"' + $argument.Replace('"', '\"') + '"'
            } else {
                $argument
            }
        }
    )
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FilePath
    $startInfo.Arguments = $quotedArguments -join " "
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    [void]$process.Start()
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        Add-Content -LiteralPath $logPath -Value $stdout.TrimEnd() -Encoding UTF8
    }
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        Add-Content -LiteralPath $logPath -Value $stderr.TrimEnd() -Encoding UTF8
    }
    if ($process.ExitCode -ne 0) {
        throw "$FailureMessage (exit code $($process.ExitCode)). See $logPath"
    }
}

function Get-Python312 {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"),
        "C:\Program Files\Python312\python.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Get-Item -LiteralPath $candidate).FullName
        }
    }

    $pyLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        $resolved = & $pyLauncher.Source -3.12 -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($resolved)) {
            return $resolved.Trim()
        }
    }

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($winget) {
        Invoke-Checked -FilePath $winget.Source -Arguments @(
            "install",
            "--id", "Python.Python.3.12",
            "-e",
            "--scope", "user",
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements"
        ) -FailureMessage "Python 3.12 installation failed"

        foreach ($candidate in $candidates) {
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return (Get-Item -LiteralPath $candidate).FullName
            }
        }
    }

    $downloadPath = Join-Path $env:TEMP "python-3.12.10-amd64.exe"
    $downloadUrl = "https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"
    Write-InstallLog "winget is unavailable; downloading the signed Python installer."
    Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -OutFile $downloadPath
    $signature = Get-AuthenticodeSignature -LiteralPath $downloadPath
    if ($signature.Status -ne "Valid" -or $signature.SignerCertificate.Subject -notmatch "Python Software Foundation") {
        throw "The downloaded Python installer signature is not valid."
    }
    $arguments = @(
        "/quiet",
        "InstallAllUsers=0",
        "PrependPath=0",
        "Include_test=0",
        "Include_launcher=1",
        "SimpleInstall=1"
    )
    $process = Start-Process -FilePath $downloadPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) {
        throw "Python 3.12 installation failed (exit code $($process.ExitCode))."
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Get-Item -LiteralPath $candidate).FullName
        }
    }
    throw "Python 3.12 was installed, but python.exe could not be resolved."
}

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
        $actualHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($item.Length -ne [long]$file.bytes -or $actualHash -ne ([string]$file.sha256).ToLowerInvariant()) {
            throw "$Label model file validation failed: $path"
        }
    }
}

function Backup-UserData {
    $settingsExists = Test-Path -LiteralPath $settingsPath -PathType Leaf
    $referenceDir = Join-Path $InstallDir "reference_data"
    $referenceExists = Test-Path -LiteralPath $referenceDir -PathType Container
    if (-not $settingsExists) {
        return
    }
    $backupDir = Join-Path $InstallDir ("user-data-backups\pre-update-" + (Get-Date -Format "yyyyMMdd-HHmmss"))
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
    if ($settingsExists) {
        Copy-Item -LiteralPath $settingsPath -Destination $backupDir -Force
    }
    if ($referenceExists) {
        Copy-Item -LiteralPath $referenceDir -Destination $backupDir -Recurse -Force
    }
    Write-InstallLog "User settings and reference data were backed up."
}

function Save-AppSettings {
    $settings = @{}
    if (Test-Path -LiteralPath $settingsPath -PathType Leaf) {
        try {
            $loaded = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($property in $loaded.PSObject.Properties) {
                $settings[$property.Name] = $property.Value
            }
        } catch {
            $settings = @{}
        }
    }

    $workspaceName = [Text.Encoding]::UTF8.GetString(
        [Convert]::FromBase64String("5aSn57Wx5bel5L2c5Yqp5omL")
    )
    $productCsvName = [Text.Encoding]::UTF8.GetString(
        [Convert]::FromBase64String("55Si5ZOB6LOH5paZ6Ly45Ye6LkNTVg==")
    )
    if (-not $settings.ContainsKey("workspace_root") -or [string]::IsNullOrWhiteSpace([string]$settings["workspace_root"])) {
        $settings["workspace_root"] = Join-Path (Join-Path $env:USERPROFILE "Documents") $workspaceName
    }
    $settings["python_exe"] = $venvPython
    if (-not $settings.ContainsKey("product_csv_path") -or [string]::IsNullOrWhiteSpace([string]$settings["product_csv_path"])) {
        $settings["product_csv_path"] = Join-Path (Join-Path $InstallDir "reference_data") $productCsvName
    }
    $settings | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $settingsPath -Encoding UTF8
}

New-Item -ItemType Directory -Force -Path $InstallDir, $engineDir, $ModelRoot | Out-Null
Set-Content -LiteralPath $logPath -Value ("Runtime setup started: " + (Get-Date -Format "o")) -Encoding UTF8

if (-not (Test-Path -LiteralPath $requirements -PathType Leaf)) {
    throw "Missing OCR requirements: $requirements"
}
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Missing OCR model manifest: $manifestPath"
}

Backup-UserData

$python = Get-Python312
Write-InstallLog "Python 3.12 found: $python"
if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
    Invoke-Checked -FilePath $python -Arguments @("-X", "utf8", "-m", "venv", $venvDir) -FailureMessage "OCR virtual environment creation failed"
}

Invoke-Checked -FilePath $venvPython -Arguments @("-X", "utf8", "-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel") -FailureMessage "Python package bootstrap failed"
Invoke-Checked -FilePath $venvPython -Arguments @("-X", "utf8", "-m", "pip", "install", "-r", $requirements) -FailureMessage "OCR package installation failed"
Invoke-Checked -FilePath $venvPython -Arguments @("-X", "utf8", "-m", "pip", "check") -FailureMessage "OCR dependency validation failed"

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-ModelFiles -Root $modelSourceRoot -Manifest $manifest -Label "Bundled"
foreach ($file in $manifest.files) {
    $relative = ([string]$file.path).Replace("/", [System.IO.Path]::DirectorySeparatorChar)
    $source = Join-Path $modelSourceRoot $relative
    $destination = Join-Path $ModelRoot $relative
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $destination) | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination -Force
}
Assert-ModelFiles -Root $ModelRoot -Manifest $manifest -Label "Installed"

Invoke-Checked -FilePath $venvPython -Arguments @(
    "-X", "utf8", "-c",
    "import cv2, numpy, openpyxl, paddle, paddleocr, rapidfuzz; from PIL import Image; print('OCR_RUNTIME_OK')"
) -FailureMessage "OCR runtime import test failed"

Save-AppSettings

$exeName = [Text.Encoding]::UTF8.GetString(
    [Convert]::FromBase64String("5aSn57Wx6YCy6LKo5Yqp5omLLmV4ZQ==")
)
$appExe = Join-Path $InstallDir $exeName
Invoke-Checked -FilePath $appExe -Arguments @("--self-test") -FailureMessage "Installed APP self-test failed"
Write-InstallLog "Runtime setup completed successfully."
