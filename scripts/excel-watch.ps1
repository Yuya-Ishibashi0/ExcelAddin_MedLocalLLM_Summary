param(
    [int]$PollSeconds = 2,
    [string]$LogPath = ""
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
$startScript = Join-Path $repoRoot "scripts\\start.ps1"

if (-not $LogPath) {
    $baseDir = $env:LOCALAPPDATA
    if (-not $baseDir) {
        $baseDir = $env:TEMP
    }
    if ($baseDir) {
        $logDir = Join-Path $baseDir "ExcelLocalLLM\\logs"
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        $LogPath = Join-Path $logDir "excel-watch.log"
    }
}

function Write-Log($message) {
    if (-not $LogPath) {
        return
    }
    $line = ("{0:yyyy-MM-dd HH:mm:ss} {1}" -f (Get-Date), $message)
    try {
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    } catch {
    }
}

function Invoke-Start {
    Write-Log "Starting services via start.ps1"
    $args = @(
        "-NoProfile",
        "-ExecutionPolicy","Bypass",
        "-WindowStyle","Hidden",
        "-File", $startScript
    )
    Start-Process -WindowStyle Hidden -FilePath "powershell.exe" -ArgumentList $args
}

function Test-ExcelRunning {
    return [bool](Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
}

Write-Log "excel-watch start"

if (Test-ExcelRunning) {
    Write-Log "Excel already running; trigger start."
    Invoke-Start
}

try {
    Register-WmiEvent -Query "SELECT * FROM Win32_ProcessStartTrace WHERE ProcessName='EXCEL.EXE'" -SourceIdentifier "ExcelStart" -ErrorAction Stop | Out-Null
    Write-Log "WMI event registered."
    while ($true) {
        Wait-Event -SourceIdentifier "ExcelStart" | Out-Null
        Remove-Event -SourceIdentifier "ExcelStart" -ErrorAction SilentlyContinue
        Write-Log "Excel start detected."
        Invoke-Start
    }
} catch {
    Write-Log ("WMI register failed: {0}" -f $_.Exception.Message)
    $excelWasRunning = $false
    while ($true) {
        $excelRunning = Test-ExcelRunning
        if ($excelRunning -and -not $excelWasRunning) {
            Write-Log "Excel start detected (polling)."
            Invoke-Start
        }
        $excelWasRunning = $excelRunning
        Start-Sleep -Seconds $PollSeconds
    }
}
