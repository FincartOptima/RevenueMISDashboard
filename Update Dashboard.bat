@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

set LOGFILE=update_log.txt

echo ================================================
echo   Fincart Revenue Dashboard - Update
echo ================================================
echo.
echo Make sure the latest MIS Excel file (.xlsb or .xlsx)
echo has already been copied into this folder before
echo continuing.
echo.
echo A full log of this run is saved to:
echo   %LOGFILE%
echo.
pause

for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm:ss'"') do set NOW=%%i
echo. >> "%LOGFILE%"
echo ================================================ >> "%LOGFILE%"
echo Update run started %NOW% >> "%LOGFILE%"
echo ================================================ >> "%LOGFILE%"

echo.
echo [1/3] Reading the Excel file and rebuilding data.js...
echo       (Excel runs in the background - this can take
echo        several minutes, please wait)
echo.

powershell -NoProfile -Command "& { python -u extract.py 2>&1 | ForEach-Object { $_; $_ | Out-File -FilePath '%LOGFILE%' -Append -Encoding utf8 }; exit $LASTEXITCODE }"
if errorlevel 1 (
    echo.
    echo [FAILED] extract.py hit an error - see above, or check %LOGFILE%. Dashboard NOT updated.
    echo [FAILED] extract.py hit an error - see log above this line. >> "%LOGFILE%"
    pause
    exit /b 1
)

echo.
echo [2/3] Saving changes to Git...
echo.

git add data.js
git add -A -- "*.xlsb" "*.xlsx"
git reset -- "~$*.xlsb" "~$*.xlsx" >nul 2>nul

git diff --cached --quiet
if not errorlevel 1 (
    echo No changes detected - dashboard data is already up to date.
    echo No changes detected - nothing to commit. >> "%LOGFILE%"
    pause
    exit /b 0
)

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set TODAY=%%i

git commit -m "Update dashboard data (%TODAY%)" > "%TEMP%\dashupdate_commit.txt" 2>&1
set COMMIT_RESULT=!errorlevel!
type "%TEMP%\dashupdate_commit.txt"
type "%TEMP%\dashupdate_commit.txt" >> "%LOGFILE%"
del "%TEMP%\dashupdate_commit.txt" >nul 2>nul
if !COMMIT_RESULT! neq 0 (
    echo.
    echo [FAILED] git commit failed - see above.
    pause
    exit /b 1
)

echo.
echo [3/3] Pushing to GitHub...
echo.

git push origin main > "%TEMP%\dashupdate_push.txt" 2>&1
set PUSH_RESULT=!errorlevel!
type "%TEMP%\dashupdate_push.txt"
type "%TEMP%\dashupdate_push.txt" >> "%LOGFILE%"
del "%TEMP%\dashupdate_push.txt" >nul 2>nul
if !PUSH_RESULT! neq 0 (
    echo.
    echo [FAILED] git push failed - check your internet connection / GitHub sign-in and try again.
    echo          Your changes are safely committed locally - just re-run this script to retry the push.
    pause
    exit /b 1
)

echo.
echo ================================================ >> "%LOGFILE%"
echo Update run finished successfully. >> "%LOGFILE%"
echo ================================================
echo   Done! Dashboard updated and live on GitHub.
echo ================================================
pause
