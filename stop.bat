@echo off
title Stop Test Plan Generator
color 0C
echo.
echo  =========================================
echo   Test Plan Generator - Stopping...
echo  =========================================
echo.

set FOUND=0
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":3377 " ^| findstr "LISTENING"') do (
    echo  Killing server process (PID %%a)...
    taskkill /PID %%a /F >nul 2>&1
    set FOUND=1
)

if "%FOUND%"=="1" (
    echo  Server stopped.
) else (
    echo  No server running on port 3377.
)

echo.
pause
