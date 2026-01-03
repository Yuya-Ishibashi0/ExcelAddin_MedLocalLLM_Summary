$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    Write-Host "Run as administrator."
    exit 1
}

$manifestDir = "C:\\OfficeAddinManifests"
$manifestDest = Join-Path $manifestDir "ExcelLocalLLM.xml"
$shareName = "OfficeAddinManifests"
$catalogId = "{B5E7B94E-51B3-4F97-A2E8-2DF8D3B3D9F3}"
$catalogKey = "HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\$catalogId"
$taskName = "ExcelLocalLLM Watcher"
$installRoot = Join-Path $env:LOCALAPPDATA "ExcelLocalLLM"

schtasks /Delete /TN "$taskName" /F | Out-Null
Remove-Item -Path $catalogKey -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path $manifestDest -Force -ErrorAction SilentlyContinue

$smbCmd = Get-Command Get-SmbShare -ErrorAction SilentlyContinue
if ($smbCmd) {
    $share = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
    if ($share) {
        Remove-SmbShare -Name $shareName -Force
    }
} else {
    cmd /c "net share $shareName /delete" > $null
}

if (Test-Path $manifestDir) {
    $remaining = Get-ChildItem -Path $manifestDir -Force -ErrorAction SilentlyContinue
    if (-not $remaining) {
        Remove-Item -Path $manifestDir -Force
    }
}

$installMarkers = @(
    Join-Path $installRoot "scripts\\start.ps1",
    Join-Path $installRoot "app\\server.py"
)
$canRemoveInstall = $true
foreach ($marker in $installMarkers) {
    if (-not (Test-Path $marker)) {
        $canRemoveInstall = $false
        break
    }
}
if ($canRemoveInstall -and (Test-Path $installRoot)) {
    Remove-Item -Path $installRoot -Recurse -Force
}
