$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
$appRoot = Join-Path $repoRoot "app"
$addInRoot = Join-Path $repoRoot "add-in"

$pythonProcesses = @(Get-CimInstance Win32_Process -Filter "Name = 'python.exe'" -ErrorAction SilentlyContinue)

$ollamaRunning = Get-Process -Name "ollama" -ErrorAction SilentlyContinue
if (-not $ollamaRunning) {
    Start-Process -WindowStyle Hidden -FilePath "ollama" -ArgumentList "serve"
}

$uvicornRunning = $pythonProcesses | Where-Object {
    $_.CommandLine -like "*uvicorn*server:app*" -and $_.CommandLine -like "*--port 8787*"
}
if (-not $uvicornRunning) {
    Start-Process -WindowStyle Hidden -FilePath "python" -ArgumentList "-m","uvicorn","server:app","--host","127.0.0.1","--port","8787" -WorkingDirectory $appRoot
}

$httpRunning = $pythonProcesses | Where-Object {
    $_.CommandLine -like "*-m http.server 3000*" -or $_.CommandLine -like "*http.server 3000*"
}
if (-not $httpRunning) {
    Start-Process -WindowStyle Hidden -FilePath "python" -ArgumentList "-m","http.server","3000" -WorkingDirectory $addInRoot
}
