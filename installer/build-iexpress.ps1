param(
    [switch]$Interactive,
    [switch]$RunIExpress,
    [switch]$IncludeAssets
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot

$buildRoot = Join-Path $repoRoot "installer\\build\\iexpress"
$distRoot = Join-Path $repoRoot "installer\\dist"
$payloadZip = Join-Path $buildRoot "payload.zip"
$runCmd = Join-Path $buildRoot "run.cmd"
$sedPath = Join-Path $buildRoot "package.sed"
$outputExe = Join-Path $distRoot "ExcelLocalLLM-Setup.exe"
$licenseFileName = "LICENSE.txt"
$finishMessage = "Setup complete. You can close this window."
$sourceRoot = [IO.Path]::GetFullPath($buildRoot).TrimEnd('\') + '\'
$tempRoot = $null
if ($env:TEMP) {
    $tempRoot = Join-Path $env:TEMP "ExcelLocalLLM-IExpress"
}
if (-not $tempRoot) {
    $tempRoot = Join-Path $buildRoot "_temp_output"
}
New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null
$tempExe = Join-Path $tempRoot "ExcelLocalLLM-Setup.exe"

Write-Host "Preparing build folder..."
if (Test-Path $buildRoot) {
    Remove-Item -Path $buildRoot -Recurse -Force
}
New-Item -Path $buildRoot -ItemType Directory -Force | Out-Null
New-Item -Path $distRoot -ItemType Directory -Force | Out-Null
$licensePath = Join-Path $repoRoot $licenseFileName
$displayLicense = ""
if (Test-Path $licensePath) {
    $licenseDest = Join-Path $buildRoot $licenseFileName
    try {
        $licenseText = Get-Content -Path $licensePath -Raw
        Set-Content -Path $licenseDest -Value $licenseText -Encoding Default
    } catch {
        Copy-Item -Path $licensePath -Destination $licenseDest -Force
    }
    Write-Host "LICENSE.txt copied to build folder."
    $displayLicense = [IO.Path]::GetFullPath($licenseDest)
} else {
    Write-Host "LICENSE.txt not found; skipping license prompt."
}

$excludeMarkers = @(
    "\\.git\\",
    "\\installer\\build\\",
    "\\installer\\dist\\",
    "\\installer\\Output\\",
    "\\installer\\logs\\",
    "\\__pycache__\\",
    "\\.vscode\\"
)
if (-not $IncludeAssets) {
    $excludeMarkers += "\\installer\\assets\\"
}

$files = Get-ChildItem -Path $repoRoot -Recurse -File -Force
$included = New-Object System.Collections.Generic.List[System.IO.FileInfo]
foreach ($file in $files) {
    $pathLower = $file.FullName.ToLowerInvariant()
    $skip = $false
    foreach ($marker in $excludeMarkers) {
        if ($pathLower.Contains($marker.ToLowerInvariant())) {
            $skip = $true
            break
        }
    }
    if ($skip) {
        continue
    }
    if ($file.Extension -in @(".pyc", ".pyo")) {
        continue
    }
    $included.Add($file) | Out-Null
}

if (Test-Path $payloadZip) {
    Remove-Item -Path $payloadZip -Force
}
Write-Host "Creating payload.zip..."
if ($IncludeAssets) {
    Write-Host "Including installer/assets (offline-capable, larger package)."
} else {
    Write-Host "Excluding installer/assets (online download, smaller package)."
}
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($payloadZip, [System.IO.Compression.ZipArchiveMode]::Create)
$total = $included.Count
$index = 0
foreach ($file in $included) {
    $relative = $file.FullName.Substring($repoRoot.Length + 1)
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
        $zip,
        $file.FullName,
        $relative,
        [System.IO.Compression.CompressionLevel]::NoCompression
    ) | Out-Null
    $index += 1
    if ($index % 200 -eq 0) {
        Write-Host "Zipping $index / $total"
    }
}
$zip.Dispose()
$payloadSize = (Get-Item $payloadZip).Length
if ($payloadSize -gt 1900000000) {
    Write-Host "Warning: payload.zip is over ~1.9GB; IExpress may fail."
}

$runCmdContent = @'
@echo off
setlocal
set "ROOT=%TEMP%\ExcelLocalLLM-Install"
if exist "%ROOT%" rmdir /s /q "%ROOT%"
mkdir "%ROOT%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Path '%~dp0payload.zip' -DestinationPath '%ROOT%' -Force"
call "%ROOT%\installer\install-all.cmd"
'@
Set-Content -Path $runCmd -Value $runCmdContent -Encoding ASCII

$sourceFiles0 = @(
    "%FILE0%=",
    "%FILE1%="
)
$strings = @(
    "FILE0=""payload.zip""",
    "FILE1=""run.cmd"""
)
if ($displayLicense) {
    $sourceFiles0 += "%FILE2%="
    $strings += "FILE2=""$licenseFileName"""
}
$sedLines = @(
    "[Version]",
    "Class=IEXPRESS",
    "SEDVersion=3",
    "",
    "[Options]",
    "PackagePurpose=InstallApp",
    "ShowInstallProgramWindow=1",
    "HideExtractAnimation=1",
    "UseLongFileName=1",
    "InsideCompressed=1",
    "CAB_FixedSize=0",
    "CAB_ResvCodeSigning=0",
    "RebootMode=N",
    "InstallPrompt=",
    "DisplayLicense=$displayLicense",
    "FinishMessage=$finishMessage",
    "TargetName=$tempExe",
    "FriendlyName=Excel Local LLM Installer",
    "AppLaunched=run.cmd",
    "PostInstallCmd=<None>",
    "AdminQuietInstCmd=",
    "UserQuietInstCmd=",
    "SourceFiles=SourceFiles",
    "",
    "[SourceFiles]",
    "SourceFiles0=$sourceRoot",
    "",
    "[SourceFiles0]"
) + $sourceFiles0 + @(
    "",
    "[Strings]"
) + $strings
$sedContent = ($sedLines -join "`r`n")
Set-Content -Path $sedPath -Value $sedContent -Encoding ASCII

$payloadSize = (Get-Item $payloadZip).Length
Write-Host "Payload prepared: $payloadZip ($payloadSize bytes)"
Write-Host "Run IExpress GUI and select payload.zip + run.cmd to build the EXE."
if (-not $RunIExpress) {
    return
}

$iexpress = Get-Command iexpress.exe -ErrorAction Stop
Write-Host "Running IExpress..."
$tempReport = Join-Path $tempRoot "~ExcelLocalLLM-Setup.RPT"
if (Test-Path $tempExe) {
    Remove-Item -Path $tempExe -Force
}
if (Test-Path $tempReport) {
    Remove-Item -Path $tempReport -Force
}
$args = @("/N", "/M", $sedPath)
if (-not $Interactive) {
    $args = @("/N", "/Q", "/M", $sedPath)
}
$process = Start-Process -FilePath $iexpress.Source -ArgumentList $args -WorkingDirectory $buildRoot -Wait -PassThru
if ($process.ExitCode -ne 0) {
    $logDirs = @($env:TEMP, "$env:WINDIR\Temp") | Where-Object { $_ -and (Test-Path $_) }
    foreach ($dir in $logDirs) {
        $log = Get-ChildItem -Path $dir -Filter "IEXPRESS*.LOG" -ErrorAction SilentlyContinue |
            Sort-Object -Property LastWriteTime -Descending |
            Select-Object -First 1
        if ($log) {
            Write-Host "IExpress log: $($log.FullName)"
            Get-Content -Path $log.FullName -Tail 50
            break
        }
    }
    throw "IExpress failed with exit code $($process.ExitCode)."
}
if (-not (Test-Path $tempExe)) {
    throw "ExcelLocalLLM-Setup.exe was not created in $tempRoot."
}
Copy-Item -Path $tempExe -Destination $outputExe -Force
if (Test-Path $tempReport) {
    Copy-Item -Path $tempReport -Destination (Join-Path $distRoot "~ExcelLocalLLM-Setup.RPT") -Force
}
Write-Host "Done: $outputExe"
