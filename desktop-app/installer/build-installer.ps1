param(
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"
$installerDir = $PSScriptRoot
$appDir = Split-Path -Parent $installerDir
$outputDir = Join-Path $appDir "dist-installer"
$appBaseName = [Text.Encoding]::UTF8.GetString(
    [Convert]::FromBase64String("5aSn57Wx6YCy6LKo5Yqp5omL")
)
$issPath = Join-Path $installerDir ($appBaseName + ".iss")

$isccCandidates = @(
    (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"),
    (Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe")
)
$isccCommand = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if ($isccCommand) {
    $isccCandidates = @($isccCommand.Source) + $isccCandidates
}
$iscc = $isccCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($iscc)) {
    throw "Inno Setup 6 was not found. Install JRSoftware.InnoSetup first."
}

$pythonCandidates = @(
    (Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"),
    (Get-Command python.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue)
)
$python = $pythonCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($python)) {
    throw "Python was not found."
}

& $python -X utf8 (Join-Path $appDir "scripts\generate-package-manifest.py")
if ($LASTEXITCODE -ne 0) {
    throw "Package manifest generation failed."
}

$appExe = Join-Path $appDir ($appBaseName + ".exe")
$selfTest = Start-Process -FilePath $appExe -ArgumentList @("--self-test") -Wait -PassThru -WindowStyle Hidden
if ($selfTest.ExitCode -ne 0) {
    throw "APP self-test failed before installer build."
}

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
& $iscc "/DMyAppVersion=$Version" $issPath
if ($LASTEXITCODE -ne 0) {
    throw "Installer compilation failed."
}

$installerPath = Join-Path $outputDir ("Datong-Invoice-Assistant-Setup-v{0}.exe" -f $Version)
if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
    throw "Built installer was not found: $installerPath"
}

$hash = (Get-FileHash -LiteralPath $installerPath -Algorithm SHA256).Hash
$hashLine = "{0}  {1}" -f $hash.ToLowerInvariant(), (Split-Path -Leaf $installerPath)
$hashPath = "$installerPath.sha256"
Set-Content -LiteralPath $hashPath -Value $hashLine -Encoding ASCII

$releaseManifest = [ordered]@{
    schema_version = 1
    app_name = "Datong Invoice Assistant"
    version = $Version
    generated_at = (Get-Date).ToString("o")
    installer = Split-Path -Leaf $installerPath
    bytes = (Get-Item -LiteralPath $installerPath).Length
    sha256 = $hash.ToLowerInvariant()
    requires_network_on_first_install = $true
    preserves_user_settings_and_reference_data = $true
}
$releaseManifestPath = Join-Path $outputDir ("release-manifest-v{0}.json" -f $Version)
$releaseManifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $releaseManifestPath -Encoding UTF8

Write-Host "Installer: $installerPath"
Write-Host "SHA256: $hash"
Write-Host "Release manifest: $releaseManifestPath"
