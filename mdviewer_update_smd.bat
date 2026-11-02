@echo off
rem ============================================================
rem  get_latest_smd.bat
rem  Windows (cmd) version of get_latest_smd.sh
rem
rem  Fetches the latest somarkdown version and saves
rem  somarkdown.umd.min.js / somarkdown.css into lib\somarkdown,
rem  then writes the version number to lib\somarkdown\VERSION.
rem
rem  Usage:
rem      get_latest_smd.bat          normal (verbose) mode
rem      get_latest_smd.bat -k      quiet mode (same as -k in the .sh)
rem
rem  Requires: Windows 10 1803+ (built-in curl.exe) and PowerShell.
rem ============================================================

setlocal ENABLEEXTENSIONS

set "TARGET_DIR=lib\somarkdown"
set "VERSION_FILE=%TARGET_DIR%\VERSION"
set "PACKAGE_API=https://data.jsdelivr.com/v1/package/npm/somarkdown"
rem  Same npm package as cdn.jsdelivr.net; unpkg is used first because it
rem  is faster in some networks, jsDelivr remains as the fallback.
set "CDN_PRIMARY=https://unpkg.com/somarkdown"
set "CDN_FALLBACK=https://cdn.jsdelivr.net/npm/somarkdown"
set "MAX_TIME=60"
set "QUIET=0"
set "LATEST_VERSION="

rem ---------- parse options ----------
:parse_args
if "%~1"=="" goto args_done
if "%~1"=="-k" (
    set "QUIET=1"
) else (
    echo Unknown option: %~1
    exit /b 1
)
shift
goto parse_args
:args_done

rem ---------- prepare directory ----------
if "%QUIET%"=="1" goto do_mkdir
echo [get_latest_smd] Preparing directory: %TARGET_DIR%
:do_mkdir
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

rem ---------- fetch latest version ----------
if "%QUIET%"=="1" goto do_fetch
echo [get_latest_smd] Fetching latest somarkdown version
:do_fetch
set "TMP_VER=%TEMP%\get_latest_smd_ver_%RANDOM%.txt"
powershell -NoProfile -Command "(Invoke-RestMethod -Uri '%PACKAGE_API%').tags.latest | Out-File -FilePath '%TMP_VER%' -Encoding ascii"
set /p LATEST_VERSION=<"%TMP_VER%"
del "%TMP_VER%" 2>nul
if not defined LATEST_VERSION (
    if "%QUIET%"=="1" goto fetch_fail
    echo [get_latest_smd] ERROR: failed to determine the latest version
    goto fail
)
if "%QUIET%"=="1" goto show_version
echo [get_latest_smd] Latest version: %LATEST_VERSION%
:show_version

rem ---------- download files ----------
rem  Each file: try the primary CDN, fall back to jsDelivr.
if "%QUIET%"=="1" goto dl_js
echo [get_latest_smd] Downloading somarkdown.umd.min.js
:dl_js
curl.exe -fsSL --max-time %MAX_TIME% "%CDN_PRIMARY%@%LATEST_VERSION%/dist/somarkdown.umd.min.js" -o "%TARGET_DIR%\somarkdown.umd.min.js" || curl.exe -fsSL --max-time %MAX_TIME% "%CDN_FALLBACK%@%LATEST_VERSION%/dist/somarkdown.umd.min.js" -o "%TARGET_DIR%\somarkdown.umd.min.js"
if errorlevel 1 goto fail

if "%QUIET%"=="1" goto dl_css
echo [get_latest_smd] Downloading somarkdown.css
:dl_css
curl.exe -fsSL --max-time %MAX_TIME% "%CDN_PRIMARY%@%LATEST_VERSION%/dist/somarkdown.css" -o "%TARGET_DIR%\somarkdown.css" || curl.exe -fsSL --max-time %MAX_TIME% "%CDN_FALLBACK%@%LATEST_VERSION%/dist/somarkdown.css" -o "%TARGET_DIR%\somarkdown.css"
if errorlevel 1 goto fail

rem ---------- write version file ----------
if "%QUIET%"=="1" goto write_ver
echo [get_latest_smd] Writing version file: %VERSION_FILE%
:write_ver
> "%VERSION_FILE%" echo %LATEST_VERSION%
if errorlevel 1 goto fail

if "%QUIET%"=="1" goto done
echo [get_latest_smd] Done
:done
exit /b 0

:fetch_fail
:fail
if "%QUIET%"=="1" exit /b 1
echo [get_latest_smd] FAILED
exit /b 1
