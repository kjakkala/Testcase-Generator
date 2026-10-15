param(
  [int]$Port = $(if ($env:QA_DASHBOARD_PORT) { [int]$env:QA_DASHBOARD_PORT } else { 3377 }),
  [string]$DataDir = $(if ($env:QA_DASHBOARD_DATA_DIR) { $env:QA_DASHBOARD_DATA_DIR } else { Join-Path $env:ProgramData 'Upland\TestPlanGenerator' })
)

$ErrorActionPreference = 'Stop'
$InstallDir = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Node = Get-Command node.exe -ErrorAction Stop
$LogDir = Join-Path $DataDir 'logs'
$PidFile = Join-Path $DataDir 'server.pid'

New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$listener = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($listener) {
  throw "Port $Port is already in use by PID $($listener.OwningProcess). No process was stopped."
}

$env:HOST = '127.0.0.1'
$env:PORT = [string]$Port
$env:COOKIE_SECURE = 'true'
$env:APP_USER_STORE = Join-Path $DataDir 'local-store.json'
$env:APP_TOKEN_HISTORY_FILE = Join-Path $DataDir 'token-history.json'
$env:APP_CONFIG_FILE = Join-Path $DataDir 'config.json'

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stdout = Join-Path $LogDir "server-$stamp.out.log"
$stderr = Join-Path $LogDir "server-$stamp.err.log"
$process = Start-Process -FilePath $Node.Source -ArgumentList 'server.js' -WorkingDirectory $InstallDir -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru
Set-Content -LiteralPath $PidFile -Value $process.Id -Encoding ASCII

Write-Output "Started QA Dashboard PID $($process.Id) on loopback port $Port."
Write-Output "Persistent data: $DataDir"
Write-Output "Logs: $LogDir"