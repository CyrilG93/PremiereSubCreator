@echo off
setlocal enabledelayedexpansion

REM // Resolve current directory and expected build output.
set "SUBCREATOR_SCRIPT_DIR=%~dp0"
set "SUBCREATOR_PROJECT_DIR=%SUBCREATOR_SCRIPT_DIR%.."
set "SUBCREATOR_SOURCE_DIR=%SUBCREATOR_PROJECT_DIR%\dist\com.cyrilg93.subcreator"
set "SUBCREATOR_DEST_DIR=%APPDATA%\Adobe\CEP\extensions\com.cyrilg93.subcreator"

REM // Validate build output is present before installing.
if not exist "%SUBCREATOR_SOURCE_DIR%" (
  echo Build missing: %SUBCREATOR_SOURCE_DIR%
  echo Run: npm run subcreator:build
  exit /b 1
)

REM // Ensure destination parent exists and refresh extension files.
if not exist "%APPDATA%\Adobe\CEP\extensions" mkdir "%APPDATA%\Adobe\CEP\extensions"
if exist "%SUBCREATOR_DEST_DIR%" rmdir /s /q "%SUBCREATOR_DEST_DIR%"
xcopy "%SUBCREATOR_SOURCE_DIR%" "%SUBCREATOR_DEST_DIR%" /e /i /h /y >nul

echo Sub Creator installed to %SUBCREATOR_DEST_DIR%
echo If needed, enable CEP debug mode and restart Premiere Pro.
endlocal
