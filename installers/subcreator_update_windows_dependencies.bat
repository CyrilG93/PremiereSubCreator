@echo off
setlocal

REM // Launch the signed-hash PowerShell updater from the same folder as this beginner-friendly entry point.
title Sub Creator - Update dependencies
echo.
echo Sub Creator dependency updater
echo ==============================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0subcreator_update_windows_dependencies.ps1"
set "SUBCREATOR_EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%SUBCREATOR_EXIT_CODE%"=="0" (
  echo Update failed. Read the message above, then run this file again.
) else (
  echo Dependencies updated successfully.
)
echo.
pause
exit /b %SUBCREATOR_EXIT_CODE%
