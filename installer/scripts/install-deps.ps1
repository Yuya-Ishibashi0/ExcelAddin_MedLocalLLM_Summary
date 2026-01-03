param(
    [string]$PythonVersion = "3.13.0",
    [string]$PythonInstaller = "",
    [string]$OllamaInstaller = "",
    [switch]$Offline
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$assetsDir = Join-Path $scriptRoot "..\\assets"

Write-Host "install-deps.ps1 path: $PSCommandPath"
if (-not (Test-Path $assetsDir)) {
    New-Item -Path $assetsDir -ItemType Directory -Force | Out-Null
}

function Test-Command($name) {
    return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

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

function Ensure-File($path, $url, $offlineMode) {
    if (Test-Path $path) {
        return $path
    }
    if ($offlineMode -or -not $url) {
        return $null
    }
    Invoke-WebRequest -Uri $url -OutFile $path
    return $path
}

function Get-ArchitectureTag {
    $arch = $env:PROCESSOR_ARCHITECTURE
    if ($env:PROCESSOR_ARCHITEW6432) {
        $arch = $env:PROCESSOR_ARCHITEW6432
    }
    switch ($arch) {
        "ARM64" { return "arm64" }
        "AMD64" { return "amd64" }
        default { return "" }
    }
}

$pythonExe = Find-PythonExe
$pythonOk = [bool]$pythonExe
$pythonCommands = Get-Command python -All -ErrorAction SilentlyContinue
foreach ($cmd in $pythonCommands) {
    if (-not $cmd -or -not $cmd.Source) {
        continue
    }
    if (Test-IsWindowsAppsPath $cmd.Source) {
        Write-Host "Python alias detected (ignored): $($cmd.Source)"
    } else {
        Write-Host "Python command candidate: $($cmd.Source)"
    }
}
if ($pythonOk) {
    Write-Host "Python found: $pythonExe"
} else {
    Write-Host "Python not found. Will install."
}

$archTag = Get-ArchitectureTag
if (-not $archTag) {
    throw "Unsupported architecture: $env:PROCESSOR_ARCHITECTURE"
}

if (-not $PythonInstaller) {
    $pattern = "python-*-$archTag.exe"
    $match = Get-ChildItem -Path $assetsDir -Filter $pattern -ErrorAction SilentlyContinue |
        Sort-Object -Property Name -Descending |
        Select-Object -First 1
    if ($match) {
        $PythonInstaller = $match.FullName
    } else {
        $PythonInstaller = Join-Path $assetsDir "python-$PythonVersion-$archTag.exe"
    }
}

if (-not $pythonOk) {
    $pythonUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-$archTag.exe"
    $pythonPath = Ensure-File $PythonInstaller $pythonUrl $Offline
    if (-not $pythonPath) {
        throw "Python installer not found. Place python-$PythonVersion-$archTag.exe in $assetsDir."
    }
    Write-Host "Running Python installer: $pythonPath"
    $process = Start-Process -FilePath $pythonPath -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_test=0" -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Python installer failed with exit code $($process.ExitCode)."
    }
    Write-Host "Python installer completed."
    $pythonExe = Find-PythonExe
    if ($pythonExe) {
        Write-Host "Python detected after install: $pythonExe"
    } else {
        Write-Host "Python still not detected in this session."
    }
}

if (-not $OllamaInstaller) {
    $OllamaInstaller = Join-Path $assetsDir "OllamaSetup.exe"
}

$ollamaInstalled = Test-Command "ollama"
if (-not $ollamaInstalled) {
    $ollamaPath = Join-Path ${env:ProgramFiles} "Ollama\\ollama.exe"
    if (Test-Path $ollamaPath) {
        $ollamaInstalled = $true
    }
}

if (-not $ollamaInstalled) {
    $ollamaUrl = "https://ollama.com/download/OllamaSetup.exe"
    $ollamaPath = Ensure-File $OllamaInstaller $ollamaUrl $Offline
    if (-not $ollamaPath) {
        throw "Ollama installer not found. Place OllamaSetup.exe in $assetsDir."
    }
    Write-Host "Running Ollama installer: $ollamaPath"
    Start-Process -FilePath $ollamaPath -ArgumentList "/S" -Wait
    Write-Host "Ollama installer completed."
}
