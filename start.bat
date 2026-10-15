@echo off
cd /d "%~dp0"
title Test Plan Generator
color 0A
echo.
echo  =========================================
echo   Test Plan Generator - Starting...
echo  =========================================
echo.

:: Check if Node.js is installed
node --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo  ERROR: Node.js is not installed!
    echo.
    echo  Please install Node.js from: https://nodejs.org
    echo  Download the LTS version and run the installer.
    echo.
    pause
    exit /b 1
)

echo  Node.js found: 
node --version
echo.

:: Install dependencies if node_modules missing
if not exist "node_modules" (
    echo  Installing dependencies...
    npm install
    echo.
)

:: Kill any existing process on port 3377
echo  Checking for existing server on port 3377...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3377 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue; Write-Host ' Stopped PID' $_ }"
echo.

echo  Starting server on http://localhost:3377
echo.

:: Launch server in a NEW visible window (stays open, shows logs, close it to stop)
start "Test Plan Generator - Server" cmd /k "cd /d "%~dp0" && color 0A && node server.js"

:: Wait 3 seconds for server to initialise
timeout /t 3 /nobreak >nul

:: Open the dashboard in Chrome incognito (falls back to Edge InPrivate if Chrome not found)
echo  Opening dashboard in incognito window...
set "CHROME_PATH="
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe"      set "CHROME_PATH=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "CHROME_PATH=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe"       set "CHROME_PATH=%LocalAppData%\Google\Chrome\Application\chrome.exe"

if defined CHROME_PATH (
  start "" "%CHROME_PATH%" --incognito "http://localhost:3377"
) else (
  start "" "msedge" --inprivate "http://localhost:3377"
)

echo  Dashboard opened. Close the Server window to stop.
echo.
:: This launcher window can close now — server runs in its own window
exit
exit
