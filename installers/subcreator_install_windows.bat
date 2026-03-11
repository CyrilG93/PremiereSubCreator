@echo off
setlocal enabledelayedexpansion

REM // Resolve current directory and expected build output.
set "SUBCREATOR_SCRIPT_DIR=%~dp0"
set "SUBCREATOR_PROJECT_DIR=%SUBCREATOR_SCRIPT_DIR%.."
set "SUBCREATOR_SOURCE_DIR=%SUBCREATOR_PROJECT_DIR%\dist\com.cyrilg93.subcreator"
set "SUBCREATOR_DEST_DIR=%APPDATA%\Adobe\CEP\extensions\com.cyrilg93.subcreator"
set "SUBCREATOR_PYTHON_CMD="
set "SUBCREATOR_PYTHON_LABEL="
set "SUBCREATOR_PYTHON_VERSION_LINE="
set "SUBCREATOR_PYTHON_MAJOR="
set "SUBCREATOR_PYTHON_MINOR="
set "SUBCREATOR_PYTHON_SEEN="

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
call :subcreator_enable_cep_debug_mode

REM // Detect Python launcher; if missing we skip Whisper setup as requested.
call :subcreator_detect_python
if not defined SUBCREATOR_PYTHON_CMD (
  if defined SUBCREATOR_PYTHON_SEEN (
    echo Whisper setup skipped: no supported Python version found ^(need 3.8 to 3.13^). Detected: !SUBCREATOR_PYTHON_SEEN!
  ) else (
    echo Whisper setup skipped: Python not found on this machine.
  )
  echo Whisper source will be hidden in the panel.
  echo If needed, enable CEP debug mode and restart Premiere Pro.
  goto :subcreator_done
)

REM // Parse Python version to avoid unsupported 3.14+ auto-install.
call :subcreator_parse_python_version "%SUBCREATOR_PYTHON_VERSION_LINE%"
if not defined SUBCREATOR_PYTHON_MAJOR (
  echo Whisper setup skipped: unable to parse Python version from "%SUBCREATOR_PYTHON_VERSION_LINE%".
  echo If needed, enable CEP debug mode and restart Premiere Pro.
  goto :subcreator_done
)

if !SUBCREATOR_PYTHON_MAJOR! GTR 3 goto :subcreator_python_unsupported
if !SUBCREATOR_PYTHON_MAJOR! EQU 3 if !SUBCREATOR_PYTHON_MINOR! GEQ 14 goto :subcreator_python_unsupported
goto :subcreator_python_supported

:subcreator_python_unsupported
echo Whisper setup skipped: Python !SUBCREATOR_PYTHON_MAJOR!.!SUBCREATOR_PYTHON_MINOR! detected ^(openai-whisper currently targets Python ^<= 3.13^).
echo Whisper source will be hidden in the panel.
echo If needed, enable CEP debug mode and restart Premiere Pro.
goto :subcreator_done

:subcreator_python_supported
echo Installing Whisper with !SUBCREATOR_PYTHON_LABEL!...

REM // Ensure pip exists, then install openai-whisper in user site-packages.
call !SUBCREATOR_PYTHON_CMD! -m pip --version >nul 2>nul
if errorlevel 1 (
  call !SUBCREATOR_PYTHON_CMD! -m ensurepip --upgrade >nul 2>nul
)

call !SUBCREATOR_PYTHON_CMD! -m pip install --user --upgrade openai-whisper
if errorlevel 1 (
  echo Whisper package install failed. You can run manually:
  echo   !SUBCREATOR_PYTHON_LABEL! -m pip install --user --upgrade openai-whisper
) else (
  echo Whisper Python package installed successfully.
)

REM // Install ffmpeg via winget when available; otherwise keep setup non-blocking.
where ffmpeg >nul 2>nul
if not errorlevel 1 (
  echo ffmpeg already available.
  goto :subcreator_finish
)

where winget >nul 2>nul
if errorlevel 1 (
  echo ffmpeg not found and winget unavailable. Install ffmpeg manually if Whisper transcription fails.
  goto :subcreator_finish
)

echo Installing ffmpeg via winget...
winget install -e --id Gyan.FFmpeg --accept-source-agreements --accept-package-agreements >nul 2>nul
where ffmpeg >nul 2>nul
if errorlevel 1 (
  echo ffmpeg install failed. Install manually with winget or another package manager.
) else (
  echo ffmpeg installed successfully.
)

:subcreator_finish
echo If needed, enable CEP debug mode and restart Premiere Pro.

:subcreator_done
endlocal
exit /b 0

:subcreator_enable_cep_debug_mode
REM // Enable CEP debug mode for multiple CSXS versions to maximize Adobe host compatibility.
for %%v in (7 8 9 10 11 12) do (
  reg add "HKCU\Software\Adobe\CSXS.%%v" /v PlayerDebugMode /t REG_SZ /d 1 /f >nul 2>nul
)
echo CEP debug mode enabled for CSXS.7 to CSXS.12
goto :eof

:subcreator_detect_python
REM // Prefer explicit supported Python minor versions first to avoid defaulting to unsupported 3.14+.
for %%m in (13 12 11 10 9 8) do (
  for /f "tokens=* delims=" %%v in ('py -3.%%m --version 2^>nul') do (
    set "SUBCREATOR_PYTHON_CMD=py -3.%%m"
    set "SUBCREATOR_PYTHON_LABEL=py -3.%%m"
    set "SUBCREATOR_PYTHON_VERSION_LINE=%%v"
    goto :eof
  )
)

for /f "tokens=* delims=" %%v in ('py -3 --version 2^>nul') do (
  if defined SUBCREATOR_PYTHON_SEEN (
    set "SUBCREATOR_PYTHON_SEEN=!SUBCREATOR_PYTHON_SEEN!, "
  )
  set "SUBCREATOR_PYTHON_SEEN=!SUBCREATOR_PYTHON_SEEN!py -3=%%v"
)

REM // Fallback to generic python executable if it points to a supported version.
for /f "tokens=* delims=" %%v in ('python --version 2^>nul') do (
  set "SUBCREATOR_PYTHON_VERSION_LINE=%%v"
  call :subcreator_parse_python_version "%%v"
  if defined SUBCREATOR_PYTHON_MAJOR (
    if !SUBCREATOR_PYTHON_MAJOR! EQU 3 if !SUBCREATOR_PYTHON_MINOR! GEQ 8 if !SUBCREATOR_PYTHON_MINOR! LEQ 13 (
      set "SUBCREATOR_PYTHON_CMD=python"
      set "SUBCREATOR_PYTHON_LABEL=python"
      goto :eof
    )
  )
  if defined SUBCREATOR_PYTHON_SEEN (
    set "SUBCREATOR_PYTHON_SEEN=!SUBCREATOR_PYTHON_SEEN!, "
  )
  set "SUBCREATOR_PYTHON_SEEN=!SUBCREATOR_PYTHON_SEEN!python=%%v"
)

REM // Last fallback to py -3 only when it resolves to a supported version.
for /f "tokens=* delims=" %%v in ('py -3 --version 2^>nul') do (
  set "SUBCREATOR_PYTHON_VERSION_LINE=%%v"
  call :subcreator_parse_python_version "%%v"
  if defined SUBCREATOR_PYTHON_MAJOR (
    if !SUBCREATOR_PYTHON_MAJOR! EQU 3 if !SUBCREATOR_PYTHON_MINOR! GEQ 8 if !SUBCREATOR_PYTHON_MINOR! LEQ 13 (
      set "SUBCREATOR_PYTHON_CMD=py -3"
      set "SUBCREATOR_PYTHON_LABEL=py -3"
      goto :eof
    )
  )
)
goto :eof

:subcreator_parse_python_version
REM // Parse "Python X.Y.Z" into major/minor numeric values.
set "SUBCREATOR_PYTHON_MAJOR="
set "SUBCREATOR_PYTHON_MINOR="
set "SUBCREATOR_VERSION_VALUE="
for /f "tokens=2 delims= " %%v in ("%~1") do set "SUBCREATOR_VERSION_VALUE=%%v"
for /f "tokens=1,2 delims=." %%a in ("!SUBCREATOR_VERSION_VALUE!") do (
  set "SUBCREATOR_PYTHON_MAJOR=%%a"
  set "SUBCREATOR_PYTHON_MINOR=%%b"
)
goto :eof
