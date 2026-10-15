param(
  [string]$DataDir = $(if ($env:QA_DASHBOARD_DATA_DIR) { $env:QA_DASHBOARD_DATA_DIR } else { Join-Path $env:ProgramData 'Upland\TestPlanGenerator' })
)

$ErrorActionPreference = 'Stop'
$PidFile = Join-Path $DataDir 'server.pid'
if (-not (Test-Path -LiteralPath $PidFile)) {
  Write-Output 'No QA Dashboard PID file found; nothing was stopped.'
  exit 0
}

$serverPid = [int](Get-Content -LiteralPath $PidFile -Raw).Trim()
$process = Get-CimInstance Win32_Process -Filter "ProcessId = $serverPid" -ErrorAction SilentlyContinue
if (-not $process) {
  Remove-Item -LiteralPath $PidFile -Force
  Write-Output 'The recorded process was not running; removed the stale PID file.'
  exit 0
}

if ($process.Name -notin @('node.exe', 'node')) {
  throw "PID $serverPid is not Node.js. Refusing to stop it."
}
if ($process.CommandLine -notmatch 'server\.js') {
  throw "PID $serverPid is not running server.js. Refusing to stop it."
}

Stop-Process -Id $serverPid -Force
Remove-Item -LiteralPath $PidFile -Force
Write-Output "Stopped QA Dashboard PID $serverPid."