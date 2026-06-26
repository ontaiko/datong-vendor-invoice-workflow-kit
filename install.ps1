param(
    [string]$ProjectRoot = "",
    [string]$CodexHome = "",
    [switch]$SkipEngine,
    [switch]$NoEnvUpdate
)

$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectFolderName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("5aSn57Wx5bel5L2c5Yqp5omL"))
$inventoryFolderName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("5bu65qqU6YCy6LKo55So"))
$ocrFolderName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("6YCy6LKo5ZyW54mH6L2J6Kmm566X6KGo"))
$invoiceImageFolderName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("5bug5ZWG6YCy6LKo5Zau"))

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Join-Path ([Environment]::GetFolderPath("MyDocuments")) $projectFolderName
}
if ([string]::IsNullOrWhiteSpace($CodexHome)) {
    $CodexHome = Join-Path $env:USERPROFILE ".codex"
}

$ProjectRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
$CodexHome = [System.IO.Path]::GetFullPath($CodexHome)
$skillsTarget = Join-Path $CodexHome "skills"

function Copy-DirectoryContents {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )
    if (-not (Test-Path -LiteralPath $Source)) {
        throw "Missing source directory: $Source"
    }
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    Copy-Item -LiteralPath (Join-Path $Source "*") -Destination $Destination -Recurse -Force
}

function Update-TextPaths {
    param([string[]]$Roots)

    $extensions = @(".md", ".py", ".ps1", ".json", ".yaml", ".yml", ".txt")
    foreach ($root in $Roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object {
            $extensions -contains $_.Extension.ToLowerInvariant()
        } | ForEach-Object {
            $path = $_.FullName
            $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
            $updated = $text.Replace("C:\Users\user\Documents\大統工作助手", $ProjectRoot)
            $updated = $updated.Replace("C:\Users\user\.codex", $CodexHome)
            $updated = $updated.Replace("C:\Users\user", $env:USERPROFILE)
            if ($updated -ne $text) {
                Set-Content -LiteralPath $path -Value $updated -Encoding UTF8
            }
        }
    }
}

New-Item -ItemType Directory -Force -Path $ProjectRoot, $CodexHome, $skillsTarget | Out-Null

Copy-DirectoryContents -Source (Join-Path $packageRoot "skills") -Destination $skillsTarget
Copy-DirectoryContents -Source (Join-Path $packageRoot "project") -Destination $ProjectRoot

New-Item -ItemType Directory -Force -Path `
    (Join-Path $ProjectRoot ".codex-tmp"), `
    (Join-Path $ProjectRoot $inventoryFolderName), `
    (Join-Path (Join-Path $ProjectRoot $inventoryFolderName) $ocrFolderName), `
    (Join-Path $ProjectRoot $invoiceImageFolderName) | Out-Null

Update-TextPaths -Roots @($skillsTarget, $ProjectRoot)

if (-not $NoEnvUpdate) {
    [Environment]::SetEnvironmentVariable("DATONG_WORKSPACE", $ProjectRoot, "User")
    [Environment]::SetEnvironmentVariable("CODEX_HOME", $CodexHome, "User")
    $env:DATONG_WORKSPACE = $ProjectRoot
    $env:CODEX_HOME = $CodexHome
}

if (-not $SkipEngine) {
    & (Join-Path $packageRoot "scripts\install-ocr-engine.ps1") -ProjectRoot $ProjectRoot -PackageRoot $packageRoot
}

Write-Host ""
Write-Host "Install complete."
Write-Host "Project: $ProjectRoot"
Write-Host "Codex skills: $skillsTarget"
Write-Host "Restart Codex before using the installed skills."
