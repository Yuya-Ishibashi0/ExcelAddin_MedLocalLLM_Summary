param(
    [switch]$Offline,
    [switch]$SkipModel,
    [string]$Model = ""
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$depsScript = Join-Path $scriptRoot "install-deps.ps1"
$appScript = Join-Path $scriptRoot "install-app.ps1"

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-DefaultModel($rootPath) {
    if ($env:OLLAMA_MODEL) {
        return $env:OLLAMA_MODEL
    }
    $settingsPath = Join-Path $rootPath "app\\settings.json"
    if (Test-Path $settingsPath) {
        try {
            $settings = Get-Content -Raw -Encoding UTF8 $settingsPath | ConvertFrom-Json
            if ($settings -and $settings.ollama_model) {
                return [string]$settings.ollama_model
            }
        } catch {
            # Ignore invalid JSON; fall back to default.
        }
    }
    return "gemma3:1b-it-qat"
}

function Get-OllamaCommand {
    $cmd = Get-Command ollama -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }
    $candidates = @(
        Join-Path ${env:ProgramFiles} "Ollama\\ollama.exe",
        Join-Path ${env:LOCALAPPDATA} "Programs\\Ollama\\ollama.exe",
        Join-Path ${env:LOCALAPPDATA} "Ollama\\ollama.exe"
    )
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }
    return $null
}

function Test-OllamaModel {
    param(
        [string]$OllamaExe,
        [string]$ModelName
    )
    try {
        $output = & $OllamaExe list 2>$null
    } catch {
        return $false
    }
    if (-not $output) {
        return $false
    }
    $pattern = "(?m)^$([regex]::Escape($ModelName))\\s"
    return ($output -match $pattern)
}

$logDir = Join-Path $repoRoot "installer\\logs"
New-Item -Path $logDir -ItemType Directory -Force | Out-Null
$logPath = Join-Path $logDir ("install-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
$transcriptStarted = $false
try {
    Start-Transcript -Path $logPath -Append | Out-Null
    $transcriptStarted = $true
} catch {
}

try {
    if (-not (Test-IsAdmin)) {
        throw "Run as administrator."
    }

    $depsArgs = @{}
    if ($Offline) {
        $depsArgs.Offline = $true
    }
    & $depsScript @depsArgs
    & $appScript

    if (-not $SkipModel) {
        $ollamaExe = Get-OllamaCommand
        if (-not $ollamaExe) {
            Write-Host "Ollama not found; skip model download."
            return
        }
        $targetModel = $Model
        if (-not $targetModel) {
            $targetModel = Get-DefaultModel $repoRoot
        }
        if (-not $targetModel) {
            return
        }
        $alreadyExists = Test-OllamaModel -OllamaExe $ollamaExe -ModelName $targetModel
        if ($alreadyExists) {
            return
        }
        if ($Offline) {
            Write-Host "Model '$targetModel' not found. Run 'ollama pull $targetModel' when online."
            return
        }
        Write-Host "Downloading model: $targetModel"
        try {
            & $ollamaExe pull $targetModel
            if ($LASTEXITCODE -ne 0) {
                Write-Host "Model download failed. Run 'ollama pull $targetModel' later."
            }
        } catch {
            Write-Host "Model download failed. Run 'ollama pull $targetModel' later."
        }
    }
} finally {
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
    if ($logPath) {
        Write-Host "Log: $logPath"
    }
}
