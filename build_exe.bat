@echo off
:: ============================================================
::  ExtraDataCleaner — Windows .exe builder
::  Double-click this file once to produce ExtraDataCleaner.exe
:: ============================================================

setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo  ExtraDataCleaner -- build script
echo  ==================================
echo.

:: ── 1. Find Python ────────────────────────────────────────────────────────────
set PY=

python --version >nul 2>&1
if not errorlevel 1 ( set PY=python & goto :found_python )

py --version >nul 2>&1
if not errorlevel 1 ( set PY=py & goto :found_python )

python3 --version >nul 2>&1
if not errorlevel 1 ( set PY=python3 & goto :found_python )

for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Program Files\Python313\python.exe"
    "C:\Program Files\Python312\python.exe"
) do (
    if exist %%P ( set PY=%%P & goto :found_python )
)

echo  [ERROR] Python not found.
echo.
echo  Python 3.13 appears installed but Windows cannot see it.
echo  Please close this window, open a NEW Command Prompt, and try again.
echo  (PATH changes only apply to new windows)
pause & exit /b 1

:found_python
echo  [OK] Python found: %PY%
%PY% --version

:: ── 2. Install packages ───────────────────────────────────────────────────────
echo.
echo  Installing required packages...
%PY% -m pip install --upgrade pip --quiet
%PY% -m pip install --upgrade ^
    pandas numpy openpyxl chardet python-dateutil ^
    pillow pyinstaller ^
    --quiet

if errorlevel 1 (
    echo  [ERROR] pip install failed. Check your internet connection.
    pause & exit /b 1
)
echo  [OK] Packages installed.

:: ── 3. Generate icon ──────────────────────────────────────────────────────────
echo.
echo  Generating icon.ico...
%PY% make_icon.py
if errorlevel 1 (
    echo  [WARN] Icon generation failed - building without custom icon.
)

:: ── 4. Build with PyInstaller ─────────────────────────────────────────────────
echo.
echo  Building ExtraDataCleaner.exe  (1-3 minutes, please wait)...
echo.
%PY% -m PyInstaller ExtraDataCleaner.spec --noconfirm --clean

if errorlevel 1 (
    echo.
    echo  [ERROR] Build failed. See output above.
    pause & exit /b 1
)

:: ── 5. Move .exe into this folder ────────────────────────────────────────────
if exist "dist\ExtraDataCleaner.exe" (
    move /Y "dist\ExtraDataCleaner.exe" "ExtraDataCleaner.exe" >nul
    echo.
    echo  ============================================================
    echo   SUCCESS!  ExtraDataCleaner.exe is ready in this folder.
    echo   Double-click it any time — no Python needed to run it.
    echo  ============================================================
) else (
    echo  [ERROR] ExtraDataCleaner.exe not found in dist\
)

:: ── 6. Clean up build folders ────────────────────────────────────────────────
set /p CLEAN="  Delete build/ and dist/ folders? (y/n) [y]: "
if /i "%CLEAN%"=="" set CLEAN=y
if /i "%CLEAN%"=="y" (
    rmdir /s /q build 2>nul
    rmdir /s /q dist  2>nul
    echo  Build folders removed.
)

echo.
pause
