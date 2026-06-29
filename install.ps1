param(
    [string]$ProjectRoot = "",
    [string]$CodexHome = "",
    [switch]$SkipEngine,
    [switch]$SkipModelWarmup,
    [switch]$SkipExcelCheck,
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
$skillNames = @(
    "convert-vendor-invoice-image",
    "extract-vendor-invoice-image",
    "match-product-catalog",
    "review-invoice-product-check",
    "build-inventory-import-files"
)
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupRoot = Join-Path $CodexHome "install-backups\datong-vendor-invoice-workflow-kit\$timestamp"

function Copy-DirectoryContents {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )
    if (-not (Test-Path -LiteralPath $Source)) {
        throw "Missing source directory: $Source"
    }
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $Destination -Recurse -Force
    }
}

function Update-TextPaths {
    param([string[]]$Files)

    foreach ($path in ($Files | Sort-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $text = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        $updated = $text.Replace("C:\Users\user\Documents\大統工作助手", $ProjectRoot)
        $updated = $updated.Replace("C:\Users\user\.codex", $CodexHome)
        $updated = $updated.Replace("C:\Users\user", $env:USERPROFILE)
        if ($updated -ne $text) {
            Set-Content -LiteralPath $path -Value $updated -Encoding UTF8
        }
    }
}

function Backup-ExistingItem {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$RelativeDestination
    )
    if (-not (Test-Path -LiteralPath $Source)) { return }
    $destination = Join-Path $backupRoot $RelativeDestination
    New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
    Copy-Item -LiteralPath $Source -Destination $destination -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $ProjectRoot, $CodexHome, $skillsTarget | Out-Null

foreach ($skillName in $skillNames) {
    $sourceSkill = Join-Path (Join-Path $packageRoot "skills") $skillName
    $targetSkill = Join-Path $skillsTarget $skillName
    if (-not (Test-Path -LiteralPath (Join-Path $sourceSkill "SKILL.md") -PathType Leaf)) {
        throw "Package skill is incomplete: $sourceSkill"
    }
    Backup-ExistingItem -Source $targetSkill -RelativeDestination (Join-Path "skills" $skillName)
    if (Test-Path -LiteralPath $targetSkill) {
        Remove-Item -LiteralPath $targetSkill -Recurse -Force
    }
    Copy-Item -LiteralPath $sourceSkill -Destination $skillsTarget -Recurse -Force
}

$projectPackageRoot = Join-Path $packageRoot "project"
$memorySeedSource = Join-Path $packageRoot "memory-seed\datong-project-memory.md"
$memoryTarget = Join-Path $ProjectRoot "PROJECT_MEMORY.md"
if (-not (Test-Path -LiteralPath $memorySeedSource -PathType Leaf)) {
    throw "Project memory seed is missing: $memorySeedSource"
}
Get-ChildItem -LiteralPath $projectPackageRoot -Recurse -File | ForEach-Object {
    $relative = $_.FullName.Substring($projectPackageRoot.Length + 1)
    Backup-ExistingItem -Source (Join-Path $ProjectRoot $relative) -RelativeDestination (Join-Path "project" $relative)
}
Backup-ExistingItem -Source $memoryTarget -RelativeDestination "project\PROJECT_MEMORY.md"
Copy-DirectoryContents -Source (Join-Path $packageRoot "project") -Destination $ProjectRoot
Copy-Item -LiteralPath $memorySeedSource -Destination $memoryTarget -Force

New-Item -ItemType Directory -Force -Path `
    (Join-Path $ProjectRoot ".codex-tmp"), `
    (Join-Path $ProjectRoot $inventoryFolderName), `
    (Join-Path (Join-Path $ProjectRoot $inventoryFolderName) $ocrFolderName), `
    (Join-Path $ProjectRoot $invoiceImageFolderName) | Out-Null

$rewriteExtensions = @(".md", ".py", ".ps1", ".json", ".yaml", ".yml", ".txt")
$rewriteFiles = @()
foreach ($skillName in $skillNames) {
    $rewriteFiles += Get-ChildItem -LiteralPath (Join-Path $skillsTarget $skillName) -Recurse -File |
        Where-Object { $rewriteExtensions -contains $_.Extension.ToLowerInvariant() } |
        Select-Object -ExpandProperty FullName
}
$rewriteFiles += Get-ChildItem -LiteralPath $projectPackageRoot -Recurse -File |
    Where-Object { $rewriteExtensions -contains $_.Extension.ToLowerInvariant() } |
    ForEach-Object { Join-Path $ProjectRoot $_.FullName.Substring($projectPackageRoot.Length + 1) }
$rewriteFiles += $memoryTarget
Update-TextPaths -Files $rewriteFiles

if (-not $NoEnvUpdate) {
    [Environment]::SetEnvironmentVariable("DATONG_WORKSPACE", $ProjectRoot, "User")
    [Environment]::SetEnvironmentVariable("CODEX_HOME", $CodexHome, "User")
    $env:DATONG_WORKSPACE = $ProjectRoot
    $env:CODEX_HOME = $CodexHome
}

if (-not $SkipEngine) {
    & (Join-Path $packageRoot "scripts\install-ocr-engine.ps1") `
        -ProjectRoot $ProjectRoot `
        -PackageRoot $packageRoot `
        -SkipModelWarmup:$SkipModelWarmup
}

& (Join-Path $packageRoot "scripts\verify-install.ps1") `
    -ProjectRoot $ProjectRoot `
    -CodexHome $CodexHome `
    -SkipEngine:$SkipEngine `
    -SkipExcelCheck:$SkipExcelCheck

Write-Host ""
Write-Host "Install complete."
Write-Host "Project: $ProjectRoot"
Write-Host "Codex skills: $skillsTarget"
Write-Host "Project memory: $memoryTarget"
if (Test-Path -LiteralPath $backupRoot) {
    Write-Host "Previous files backup: $backupRoot"
}
Write-Host "Restart Codex before using the installed skills."
