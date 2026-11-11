@echo off
chcp 65001 >nul
echo =========================================
echo     Fix Python PATH (One-Click Setup)
echo =========================================
echo.

:: 1. Already working?
where py >nul 2>&1 && (py --version >nul 2>&1 && goto :ok)
python --version >nul 2>&1 && goto :ok

:: 2. Search common install folders for python.exe
echo Searching for installed Python...
set "PYDIR="
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*" "%ProgramFiles%\Python3*" "%ProgramFiles(x86)%\Python3*" "C:\Python3*") do (
    if exist "%%~D\python.exe" set "PYDIR=%%~D"
)

if not defined PYDIR goto :install

echo Found Python at: %PYDIR%
echo Adding to user PATH...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$d='%PYDIR%'; $s=Join-Path $d 'Scripts'; $p=[Environment]::GetEnvironmentVariable('Path','User'); if(-not $p){$p=''}; $parts=$p.Split(';') | Where-Object { $_ -ne '' }; foreach($x in @($d,$s)){ if($parts -notcontains $x){ $parts += $x } }; [Environment]::SetEnvironmentVariable('Path', ($parts -join ';'), 'User')"
set "PATH=%PYDIR%;%PYDIR%\Scripts;%PATH%"
goto :ok

:install
echo Python not found. Installing Python 3.11 via winget...
winget install -e --id Python.Python.3.11 --scope user --override "/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1" --accept-package-agreements --accept-source-agreements
if errorlevel 1 (
    echo.
    echo ERROR: Automatic install failed.
    echo Please download Python from https://www.python.org/downloads/
    echo and CHECK "Add python.exe to PATH" on the first screen.
    pause
    exit /b 1
)
echo.
echo Python installed. Please CLOSE this window and run start.bat again.
pause
exit /b 0

:ok
echo.
python --version 2>nul || py --version
echo.
echo SUCCESS: Python is ready.
echo Please CLOSE all cmd windows, then double-click start.bat
pause
