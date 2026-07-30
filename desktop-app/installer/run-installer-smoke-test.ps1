param(
    [Parameter(Mandatory = $true)]
    [string]$InstallerPath,

    [Parameter(Mandatory = $true)]
    [string]$TestRoot,

    [switch]$Launch
)

$ErrorActionPreference = "Stop"
$InstallerPath = [System.IO.Path]::GetFullPath($InstallerPath)
$TestRoot = [System.IO.Path]::GetFullPath($TestRoot)
$statusPath = "$TestRoot.smoke-status.json"
$logPath = "$TestRoot.setup.log"

if ($Launch) {
    $scriptPath = $MyInvocation.MyCommand.Path
    $arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -InstallerPath "{1}" -TestRoot "{2}"' -f `
        $scriptPath, $InstallerPath, $TestRoot
    Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WindowStyle Hidden | Out-Null
    Write-Host "Smoke test launched."
    Write-Host "Status: $statusPath"
    exit 0
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $TestRoot) | Out-Null
$installerArguments = '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /TASKS="" /DIR="{0}" /LOG="{1}"' -f `
    $TestRoot, $logPath
$startedAt = Get-Date
$process = Start-Process -FilePath $InstallerPath -ArgumentList $installerArguments -Wait -PassThru -WindowStyle Hidden
$result = [ordered]@{
    completed = $true
    exit_code = $process.ExitCode
    started_at = $startedAt.ToString("o")
    completed_at = (Get-Date).ToString("o")
    install_root = $TestRoot
    setup_log = $logPath
}
$result | ConvertTo-Json | Set-Content -LiteralPath $statusPath -Encoding UTF8
exit $process.ExitCode

