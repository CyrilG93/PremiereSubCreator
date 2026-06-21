param()

$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# // Resolve the versioned runtime manifest shipped beside this script in release archives.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestPath = Join-Path $scriptDir "windows-runtime.json"
$runtimeDir = Join-Path $env:LOCALAPPDATA "SubCreator\runtime"
$runtimeConfigDir = Join-Path $env:APPDATA "SubCreator"
$runtimeConfigFile = Join-Path $runtimeConfigDir "subcreator-runtime.json"
$runtimeVersionFile = Join-Path $runtimeDir ".subcreator-runtime-version"
$downloadRoot = Join-Path $env:TEMP ("SubCreator-dependencies-" + [System.Guid]::NewGuid().ToString("N"))

function Write-SubCreatorStep {
  param([string]$Message)
  # // Keep progress readable in the .bat console for non-technical users.
  Write-Host ("[Sub Creator] " + $Message) -ForegroundColor Cyan
}

function Read-SubCreatorRuntimeManifest {
  # // Reject incomplete or malformed metadata before downloading or executing anything.
  if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "Missing runtime manifest: $manifestPath"
  }

  $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
  $version = [string]$manifest.version
  $releaseTag = [string]$manifest.releaseTag
  $assetName = [string]$manifest.assetName
  $sha256 = ([string]$manifest.sha256).Trim().ToLowerInvariant()
  if (-not $version -or -not $releaseTag -or -not $assetName -or $sha256 -notmatch "^[a-f0-9]{64}$") {
    throw "The Windows runtime manifest is incomplete or invalid."
  }

  return [pscustomobject]@{
    Version = $version
    AssetName = $assetName
    Sha256 = $sha256
    Url = "https://github.com/CyrilG93/PremiereSubCreator/releases/download/$releaseTag/$assetName"
  }
}

function Get-SubCreatorRuntimeFile {
  param(
    [string]$Url,
    [string]$Destination
  )

  # // Use native Windows HTTPS support first, then BITS as a resilient fallback on older PowerShell hosts.
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  try {
    Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing
    return
  } catch {
    Write-SubCreatorStep "Standard download failed; retrying with Windows BITS..."
  }

  $bitsCommand = Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue
  if (-not $bitsCommand) {
    throw "The dependency runtime could not be downloaded from $Url"
  }
  Start-BitsTransfer -Source $Url -Destination $Destination -DisplayName "Sub Creator dependencies"
}

function Test-SubCreatorRuntime {
  # // Validate the exact executables and Python imports consumed by the extension.
  $pythonPath = Join-Path $runtimeDir "python\python.exe"
  $whisperPath = Join-Path $runtimeDir "python\Scripts\whisper.exe"
  $ffmpegPath = Join-Path $runtimeDir "ffmpeg\bin\ffmpeg.exe"
  foreach ($requiredPath in @($pythonPath, $whisperPath, $ffmpegPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
      throw "A required dependency is missing after installation: $requiredPath"
    }
  }

  & $pythonPath -c "import whisper; import whisperx; print('Whisper and WhisperX: OK')" | Write-Host
  if ($LASTEXITCODE -ne 0) {
    throw "The private Python runtime could not import Whisper and WhisperX."
  }

  & $ffmpegPath -version | Select-Object -First 1 | Write-Host
  if ($LASTEXITCODE -ne 0) {
    throw "The private FFmpeg runtime did not start correctly."
  }

  return [pscustomobject]@{
    PythonPath = $pythonPath
    WhisperPath = $whisperPath
    FfmpegPath = $ffmpegPath
    PythonVersion = (& $pythonPath --version 2>&1 | Select-Object -First 1)
  }
}

function Write-SubCreatorRuntimeConfig {
  param($RuntimePaths)

  # // Regenerate the UTF-8-without-BOM config consumed by CEP after replacing the private runtime.
  $pathHints = @(
    (Split-Path -Parent $RuntimePaths.PythonPath),
    (Split-Path -Parent $RuntimePaths.WhisperPath),
    (Split-Path -Parent $RuntimePaths.FfmpegPath),
    (Join-Path $env:SystemRoot "System32")
  ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique

  $config = [ordered]@{
    version = 1
    generatedBy = "subcreator_update_windows_dependencies.ps1"
    generatedAtUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    pythonCommand = $RuntimePaths.PythonPath
    pythonLabel = "Sub Creator private Python"
    pythonPath = $RuntimePaths.PythonPath
    pythonVersion = [string]$RuntimePaths.PythonVersion
    whisperPath = $RuntimePaths.WhisperPath
    ffmpegPath = $RuntimePaths.FfmpegPath
    pathHints = @($pathHints)
  }

  New-Item -ItemType Directory -Path $runtimeConfigDir -Force | Out-Null
  $configJson = $config | ConvertTo-Json -Depth 4
  [System.IO.File]::WriteAllText($runtimeConfigFile, $configJson, (New-Object System.Text.UTF8Encoding($false)))
}

try {
  # // Premiere keeps runtime files loaded, so fail early with one actionable instruction instead of a partial update.
  if (Get-Process | Where-Object { $_.ProcessName -match "Adobe Premiere Pro" } | Select-Object -First 1) {
    throw "Close Premiere Pro before updating Sub Creator dependencies."
  }

  $manifest = Read-SubCreatorRuntimeManifest
  New-Item -ItemType Directory -Path $downloadRoot -Force | Out-Null
  $runtimeInstallerPath = Join-Path $downloadRoot $manifest.AssetName

  Write-SubCreatorStep "Downloading the tested dependency runtime (about 333 MB)..."
  Get-SubCreatorRuntimeFile -Url $manifest.Url -Destination $runtimeInstallerPath

  Write-SubCreatorStep "Verifying the download integrity..."
  $downloadHash = (Get-FileHash -LiteralPath $runtimeInstallerPath -Algorithm SHA256).Hash.ToLowerInvariant()
  if ($downloadHash -ne $manifest.Sha256) {
    throw "Downloaded runtime SHA-256 mismatch. The file will not be executed."
  }

  Write-SubCreatorStep "Installing the private Python, WhisperX, and FFmpeg runtime..."
  $installerArguments = @("/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/CURRENTUSER")
  $installerProcess = Start-Process -FilePath $runtimeInstallerPath -ArgumentList $installerArguments -Wait -PassThru
  if ($installerProcess.ExitCode -ne 0) {
    throw "The dependency installer failed with exit code $($installerProcess.ExitCode)."
  }

  Write-SubCreatorStep "Validating dependencies..."
  $runtimePaths = Test-SubCreatorRuntime
  if (-not (Test-Path -LiteralPath $runtimeVersionFile -PathType Leaf)) {
    throw "The installed runtime version marker is missing."
  }
  $installedVersion = (Get-Content -LiteralPath $runtimeVersionFile -Raw).Trim()
  if ($installedVersion -ne $manifest.Version) {
    throw "Installed runtime version $installedVersion does not match expected version $($manifest.Version)."
  }

  Write-SubCreatorRuntimeConfig -RuntimePaths $runtimePaths
  Write-Host ""
  Write-Host "Sub Creator dependencies are ready. You can reopen Premiere Pro." -ForegroundColor Green
  exit 0
} catch {
  Write-Host ""
  Write-Host ("ERROR: " + $_.Exception.Message) -ForegroundColor Red
  exit 1
} finally {
  # // Remove only this updater's uniquely named temporary download directory.
  if (Test-Path -LiteralPath $downloadRoot) {
    Remove-Item -LiteralPath $downloadRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
