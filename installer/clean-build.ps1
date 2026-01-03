$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot

$buildRoot = Join-Path $repoRoot "installer\\build"
if (Test-Path $buildRoot) {
    Remove-Item -Path $buildRoot -Recurse -Force
}

$installerRoot = Join-Path $repoRoot "installer"
Get-ChildItem -Path $installerRoot -Filter "*.RPT" -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -Path $installerRoot -Filter "*.DDF" -ErrorAction SilentlyContinue | Remove-Item -Force
