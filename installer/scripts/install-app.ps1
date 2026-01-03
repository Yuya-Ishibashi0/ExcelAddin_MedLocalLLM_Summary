$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)

$manifestSource = Join-Path $repoRoot "add-in\\manifest.xml"
$manifestDir = "C:\\OfficeAddinManifests"
$manifestDest = Join-Path $manifestDir "ExcelLocalLLM.xml"
$shareName = "OfficeAddinManifests"
$catalogId = "{B5E7B94E-51B3-4F97-A2E8-2DF8D3B3D9F3}"
$catalogKey = "HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\$catalogId"
$taskName = "ExcelLocalLLM Watcher"
$watchScript = Join-Path $repoRoot "scripts\\excel-watch.ps1"

function Get-WindowsAppsDir {
    if (-not $env:LOCALAPPDATA) {
        return $null
    }
    $dir = Join-Path $env:LOCALAPPDATA "Microsoft\\WindowsApps"
    return [IO.Path]::GetFullPath($dir)
}

function Test-IsWindowsAppsPath($path) {
    if (-not $path) {
        return $false
    }
    $windowsApps = Get-WindowsAppsDir
    if (-not $windowsApps) {
        return $false
    }
    $full = [IO.Path]::GetFullPath($path)
    $prefix = $windowsApps.TrimEnd('\') + '\'
    return $full.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Find-PythonExe {
    $commands = Get-Command python -All -ErrorAction SilentlyContinue
    foreach ($cmd in $commands) {
        if ($cmd -and $cmd.Source -and -not (Test-IsWindowsAppsPath $cmd.Source)) {
            return $cmd.Source
        }
    }
    $patterns = @(
        (Join-Path $env:LOCALAPPDATA "Programs\\Python\\Python*\\python.exe"),
        (Join-Path ${env:ProgramFiles} "Python*\\python.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Python*\\python.exe")
    )
    foreach ($pattern in $patterns) {
        if (-not $pattern) {
            continue
        }
        $match = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
            Sort-Object -Property FullName -Descending |
            Select-Object -First 1
        if ($match) {
            return $match.FullName
        }
    }
    return $null
}

if (-not (Test-Path $manifestSource)) {
    throw "manifest.xml not found: $manifestSource"
}

New-Item -Path $manifestDir -ItemType Directory -Force | Out-Null
Copy-Item -Path $manifestSource -Destination $manifestDest -Force

$resolvedManifestDir = (Resolve-Path -Path $manifestDir).ProviderPath

function Ensure-ServerService {
    $serverSvc = Get-Service -Name "LanmanServer" -ErrorAction SilentlyContinue
    if (-not $serverSvc) {
        throw "LanmanServer service not available. SMB sharing is required."
    }
    if ($serverSvc.Status -ne "Running") {
        Start-Service -Name "LanmanServer" -ErrorAction Stop
    }
}

$shareReady = $false
$smbCmd = Get-Command Get-SmbShare -ErrorAction SilentlyContinue
if ($smbCmd) {
    Ensure-ServerService
    try {
        $share = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
        if ($share) {
            $shareReady = $true
        }
    } catch {
        $shareReady = $false
    }
    if (-not $shareReady) {
        try {
            New-SmbShare -Name $shareName -Path $resolvedManifestDir -ReadAccess "BUILTIN\\Users" -ErrorAction Stop | Out-Null
            $shareReady = $true
        } catch {
            $shareReady = $false
        }
    }
}
if (-not $shareReady) {
    cmd /c "net share $shareName" > $null 2>&1
    if ($LASTEXITCODE -ne 0) {
        cmd /c "net share $shareName=$resolvedManifestDir /GRANT:Users,READ" > $null
    }
    cmd /c "net share $shareName" > $null 2>&1
    if ($LASTEXITCODE -eq 0) {
        $shareReady = $true
    }
}
if (-not $shareReady) {
    throw "Failed to create SMB share. Run: net share $shareName=$resolvedManifestDir /GRANT:Users,READ"
}

New-Item -Path $catalogKey -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name "Id" -Value $catalogId -PropertyType String -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name "Url" -Value "\\\\localhost\\OfficeAddinManifests" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null

$taskAction = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$watchScript`""
schtasks /Create /F /SC ONLOGON /RL HIGHEST /TN "$taskName" /TR "$taskAction" /RU "$env:USERNAME" | Out-Null

$pythonExe = Find-PythonExe
if ($pythonExe) {
    $requirements = Join-Path $repoRoot "app\\requirements.txt"
    if (Test-Path $requirements) {
        try {
            & $pythonExe -m pip install -r $requirements | Out-Null
        } catch {
            # Ignore pip failures; dependencies can be installed separately.
        }
    }
}
