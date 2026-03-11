@echo off
setlocal enabledelayedexpansion

REM // Resolve current directory and expected build output.
set "SUBCREATOR_SCRIPT_DIR=%~dp0"
set "SUBCREATOR_PROJECT_DIR=%SUBCREATOR_SCRIPT_DIR%.."
set "SUBCREATOR_SOURCE_DIR=%SUBCREATOR_PROJECT_DIR%\dist\com.cyrilg93.subcreator"
set "SUBCREATOR_DEST_DIR=%APPDATA%\Adobe\CEP\extensions\com.cyrilg93.subcreator"
set "SUBCREATOR_RUNTIME_DIR=%APPDATA%\SubCreator"
set "SUBCREATOR_RUNTIME_FILE=%SUBCREATOR_RUNTIME_DIR%\subcreator-runtime.json"
set "SUBCREATOR_PYTHON_CMD="
set "SUBCREATOR_PYTHON_LABEL="
set "SUBCREATOR_PYTHON_VERSION_LINE="
set "SUBCREATOR_PYTHON_MAJOR="
set "SUBCREATOR_PYTHON_MINOR="
set "SUBCREATOR_PYTHON_PATH="
set "SUBCREATOR_PYTHON_SEEN="
set "SUBCREATOR_WHISPER_PATH="
set "SUBCREATOR_FFMPEG_PATH="
set "SUBCREATOR_PATH_HINTS="

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
  goto :subcreator_after_whisper_setup
)

REM // Parse Python version to avoid unsupported 3.14+ auto-install.
call :subcreator_parse_python_version "%SUBCREATOR_PYTHON_VERSION_LINE%"
if not defined SUBCREATOR_PYTHON_MAJOR (
  echo Whisper setup skipped: unable to parse Python version from "%SUBCREATOR_PYTHON_VERSION_LINE%".
  goto :subcreator_after_whisper_setup
)

if !SUBCREATOR_PYTHON_MAJOR! GTR 3 goto :subcreator_python_unsupported
if !SUBCREATOR_PYTHON_MAJOR! EQU 3 if !SUBCREATOR_PYTHON_MINOR! GEQ 14 goto :subcreator_python_unsupported
goto :subcreator_python_supported

:subcreator_python_unsupported
echo Whisper setup skipped: Python !SUBCREATOR_PYTHON_MAJOR!.!SUBCREATOR_PYTHON_MINOR! detected ^(openai-whisper currently targets Python ^<= 3.13^).
echo Whisper source will be hidden in the panel.
goto :subcreator_after_whisper_setup

:subcreator_python_supported
call :subcreator_detect_python_executable_path
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

:subcreator_after_whisper_setup
REM // Install ffmpeg via winget when available; otherwise keep setup non-blocking.
where ffmpeg >nul 2>nul
if not errorlevel 1 (
  echo ffmpeg already available.
  goto :subcreator_collect_runtime
)

where winget >nul 2>nul
if errorlevel 1 (
  echo ffmpeg not found and winget unavailable. Install ffmpeg manually if Whisper transcription fails.
  goto :subcreator_collect_runtime
)

echo Installing ffmpeg via winget...
winget install -e --id Gyan.FFmpeg --accept-source-agreements --accept-package-agreements >nul 2>nul
where ffmpeg >nul 2>nul
if errorlevel 1 (
  echo ffmpeg install failed. Install manually with winget or another package manager.
) else (
  echo ffmpeg installed successfully.
)

:subcreator_collect_runtime
call :subcreator_detect_python_executable_path
call :subcreator_detect_whisper_path
call :subcreator_detect_ffmpeg_path
call :subcreator_write_runtime_config

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

:subcreator_detect_python_executable_path
REM // Resolve concrete Python interpreter path from selected launcher.
set "SUBCREATOR_PYTHON_PATH="
if not defined SUBCREATOR_PYTHON_CMD goto :eof
for /f "tokens=* delims=" %%p in ('!SUBCREATOR_PYTHON_CMD! -c "import sys; print(sys.executable)" 2^>nul') do (
  set "SUBCREATOR_PYTHON_PATH=%%p"
  goto :subcreator_detect_python_executable_path_done
)
:subcreator_detect_python_executable_path_done
if defined SUBCREATOR_PYTHON_PATH (
  for %%d in ("!SUBCREATOR_PYTHON_PATH!") do call :subcreator_add_path_hint "%%~dpd"
)
goto :eof

:subcreator_detect_whisper_path
REM // Detect whisper executable path from user installs or PATH fallback.
set "SUBCREATOR_WHISPER_PATH="
if defined SUBCREATOR_PYTHON_MAJOR if defined SUBCREATOR_PYTHON_MINOR (
  if exist "%APPDATA%\Python\Python!SUBCREATOR_PYTHON_MAJOR!!SUBCREATOR_PYTHON_MINOR!\Scripts\whisper.exe" (
    set "SUBCREATOR_WHISPER_PATH=%APPDATA%\Python\Python!SUBCREATOR_PYTHON_MAJOR!!SUBCREATOR_PYTHON_MINOR!\Scripts\whisper.exe"
  )
)
if not defined SUBCREATOR_WHISPER_PATH if defined SUBCREATOR_PYTHON_MAJOR if defined SUBCREATOR_PYTHON_MINOR (
  if exist "%LOCALAPPDATA%\Programs\Python\Python!SUBCREATOR_PYTHON_MAJOR!!SUBCREATOR_PYTHON_MINOR!\Scripts\whisper.exe" (
    set "SUBCREATOR_WHISPER_PATH=%LOCALAPPDATA%\Programs\Python\Python!SUBCREATOR_PYTHON_MAJOR!!SUBCREATOR_PYTHON_MINOR!\Scripts\whisper.exe"
  )
)
if not defined SUBCREATOR_WHISPER_PATH (
  for /f "tokens=* delims=" %%w in ('where whisper 2^>nul') do (
    set "SUBCREATOR_WHISPER_PATH=%%w"
    goto :subcreator_detect_whisper_path_done
  )
)
:subcreator_detect_whisper_path_done
if defined SUBCREATOR_WHISPER_PATH (
  for %%d in ("!SUBCREATOR_WHISPER_PATH!") do call :subcreator_add_path_hint "%%~dpd"
)
goto :eof

:subcreator_detect_ffmpeg_path
REM // Detect ffmpeg binary path for runtime config generation.
set "SUBCREATOR_FFMPEG_PATH="
for /f "tokens=* delims=" %%f in ('where ffmpeg 2^>nul') do (
  set "SUBCREATOR_FFMPEG_PATH=%%f"
  goto :subcreator_detect_ffmpeg_path_done
)
if not defined SUBCREATOR_FFMPEG_PATH if exist "C:\ffmpeg\bin\ffmpeg.exe" set "SUBCREATOR_FFMPEG_PATH=C:\ffmpeg\bin\ffmpeg.exe"
if not defined SUBCREATOR_FFMPEG_PATH if exist "C:\Program Files\ffmpeg\bin\ffmpeg.exe" set "SUBCREATOR_FFMPEG_PATH=C:\Program Files\ffmpeg\bin\ffmpeg.exe"
:subcreator_detect_ffmpeg_path_done
if defined SUBCREATOR_FFMPEG_PATH (
  for %%d in ("!SUBCREATOR_FFMPEG_PATH!") do call :subcreator_add_path_hint "%%~dpd"
)
goto :eof

:subcreator_add_path_hint
REM // Keep PATH hints unique for runtime config.
set "SUBCREATOR_HINT_VALUE=%~1"
if not defined SUBCREATOR_HINT_VALUE goto :eof
if defined SUBCREATOR_PATH_HINTS (
  echo ;!SUBCREATOR_PATH_HINTS!; | find /I ";!SUBCREATOR_HINT_VALUE!;" >nul
  if not errorlevel 1 goto :eof
  set "SUBCREATOR_PATH_HINTS=!SUBCREATOR_PATH_HINTS!;!SUBCREATOR_HINT_VALUE!"
) else (
  set "SUBCREATOR_PATH_HINTS=!SUBCREATOR_HINT_VALUE!"
)
goto :eof

:subcreator_append_hint_json
REM // Append one escaped path hint to JSON array content.
set "SUBCREATOR_HINT_JSON_RAW=%~1"
if not defined SUBCREATOR_HINT_JSON_RAW goto :eof
set "SUBCREATOR_HINT_JSON_ESC=!SUBCREATOR_HINT_JSON_RAW:\=\\!"
set "SUBCREATOR_HINT_JSON_ESC=!SUBCREATOR_HINT_JSON_ESC:"=\"!"
if defined SUBCREATOR_PATH_HINTS_JSON (
  set "SUBCREATOR_PATH_HINTS_JSON=!SUBCREATOR_PATH_HINTS_JSON!, \"!SUBCREATOR_HINT_JSON_ESC!\""
) else (
  set "SUBCREATOR_PATH_HINTS_JSON=\"!SUBCREATOR_HINT_JSON_ESC!\""
)
goto :eof

:subcreator_write_runtime_config
REM // Persist installer-detected runtime config under user AppData.
if not exist "%SUBCREATOR_RUNTIME_DIR%" mkdir "%SUBCREATOR_RUNTIME_DIR%"

call :subcreator_add_path_hint "C:\Program Files\ffmpeg\bin"
call :subcreator_add_path_hint "C:\ffmpeg\bin"
call :subcreator_add_path_hint "%ProgramFiles%\Python"
call :subcreator_add_path_hint "%SystemRoot%\System32"

set "SUBCREATOR_PATH_HINTS_JSON="
call :subcreator_build_path_hints_json
if not defined SUBCREATOR_PATH_HINTS_JSON set "SUBCREATOR_PATH_HINTS_JSON="

set "JSON_PYTHON_CMD=!SUBCREATOR_PYTHON_CMD:\=\\!"
set "JSON_PYTHON_LABEL=!SUBCREATOR_PYTHON_LABEL:\=\\!"
set "JSON_PYTHON_PATH=!SUBCREATOR_PYTHON_PATH:\=\\!"
set "JSON_PYTHON_VERSION=!SUBCREATOR_PYTHON_VERSION_LINE:\=\\!"
set "JSON_WHISPER_PATH=!SUBCREATOR_WHISPER_PATH:\=\\!"
set "JSON_FFMPEG_PATH=!SUBCREATOR_FFMPEG_PATH:\=\\!"

set "JSON_PYTHON_CMD=!JSON_PYTHON_CMD:"=\"!"
set "JSON_PYTHON_LABEL=!JSON_PYTHON_LABEL:"=\"!"
set "JSON_PYTHON_PATH=!JSON_PYTHON_PATH:"=\"!"
set "JSON_PYTHON_VERSION=!JSON_PYTHON_VERSION:"=\"!"
set "JSON_WHISPER_PATH=!JSON_WHISPER_PATH:"=\"!"
set "JSON_FFMPEG_PATH=!JSON_FFMPEG_PATH:"=\"!"

set "SUBCREATOR_GENERATED_AT="
for /f "tokens=* delims=" %%t in ('powershell -NoProfile -Command "(Get-Date).ToUniversalTime().ToString(\"yyyy-MM-ddTHH:mm:ssZ\")" 2^>nul') do (
  set "SUBCREATOR_GENERATED_AT=%%t"
  goto :subcreator_generated_at_done
)
:subcreator_generated_at_done
if not defined SUBCREATOR_GENERATED_AT set "SUBCREATOR_GENERATED_AT=unknown"

(
echo {
echo   "version": 1,
echo   "generatedBy": "subcreator_install_windows.bat",
echo   "generatedAtUtc": "!SUBCREATOR_GENERATED_AT!",
echo   "pythonCommand": "!JSON_PYTHON_CMD!",
echo   "pythonLabel": "!JSON_PYTHON_LABEL!",
echo   "pythonPath": "!JSON_PYTHON_PATH!",
echo   "pythonVersion": "!JSON_PYTHON_VERSION!",
echo   "whisperPath": "!JSON_WHISPER_PATH!",
echo   "ffmpegPath": "!JSON_FFMPEG_PATH!",
echo   "pathHints": [!SUBCREATOR_PATH_HINTS_JSON!]
echo }
) > "%SUBCREATOR_RUNTIME_FILE%"

echo Runtime config written: %SUBCREATOR_RUNTIME_FILE%
echo   pythonPath=!SUBCREATOR_PYTHON_PATH!
echo   whisperPath=!SUBCREATOR_WHISPER_PATH!
echo   ffmpegPath=!SUBCREATOR_FFMPEG_PATH!
goto :eof

:subcreator_build_path_hints_json
REM // Convert semicolon-separated path hints into a JSON array payload.
if not defined SUBCREATOR_PATH_HINTS goto :eof
set "SUBCREATOR_PATH_HINTS_WORK=!SUBCREATOR_PATH_HINTS!"
:subcreator_build_path_hints_json_loop
if not defined SUBCREATOR_PATH_HINTS_WORK goto :eof
for /f "tokens=1* delims=;" %%a in ("!SUBCREATOR_PATH_HINTS_WORK!") do (
  call :subcreator_append_hint_json "%%~a"
  set "SUBCREATOR_PATH_HINTS_WORK=%%~b"
)
goto :subcreator_build_path_hints_json_loop
