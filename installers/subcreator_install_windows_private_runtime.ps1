param(
  [string]$PayloadRoot = "",
  [switch]$SkipRuntimeInstall,
  [string]$RuntimeVersion = "1"
)

$ErrorActionPreference = "Stop"

# // Resolve the payload folder whether the installer is launched from a release folder or an extracted self-extracting EXE payload.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $PayloadRoot) {
  $PayloadRoot = Split-Path -Parent $scriptDir
}
$PayloadRoot = [System.IO.Path]::GetFullPath($PayloadRoot)

$sourceDir = Join-Path $PayloadRoot "dist\com.cyrilplugin.subcreator"
$destDir = Join-Path $env:APPDATA "Adobe\CEP\extensions\com.cyrilplugin.subcreator"
$legacyDestDir = Join-Path $env:APPDATA "Adobe\CEP\extensions\com.cyrilg93.subcreator"
$runtimeDir = Join-Path $env:LOCALAPPDATA "SubCreator\runtime"
$runtimeConfigDir = Join-Path $env:APPDATA "SubCreator"
$runtimeConfigFile = Join-Path $runtimeConfigDir "subcreator-runtime.json"
$bundledModelsDir = Join-Path $PayloadRoot "Models"
$bundledFontsDir = Join-Path $PayloadRoot "Fonts"
$payloadRuntimeDir = Join-Path $PayloadRoot "runtime"
$whisperCacheDir = Join-Path $env:USERPROFILE ".cache\whisper"
$runtimeVersionFile = Join-Path $runtimeDir ".subcreator-runtime-version"

function Write-SubCreatorInfo {
  param([string]$Message)
  # // Keep install logs readable in both console and self-extracting installer windows.
  Write-Host $Message
}

function Copy-SubCreatorDirectoryFresh {
  param(
    [string]$Source,
    [string]$Destination
  )

  # // Replace generated installation folders atomically enough for a user-level CEP extension install.
  if (Test-Path -LiteralPath $Destination) {
    Remove-Item -LiteralPath $Destination -Recurse -Force
  }
  New-Item -ItemType Directory -Path (Split-Path -Parent $Destination) -Force | Out-Null
  Copy-Item -LiteralPath $Source -Destination $Destination -Recurse -Force
}

function Backup-SubCreatorUserTemplates {
  param(
    [string]$ExtensionDir,
    [string]$BackupDir
  )

  # // Preserve only user-added MOGRT files by comparing against the catalog bundled with the installed extension.
  $templatesDir = Join-Path $ExtensionDir "templates\mogrt"
  if (-not (Test-Path -LiteralPath $templatesDir)) {
    return
  }

  $bundled = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  $catalogPath = Join-Path $ExtensionDir "assets\mogrt-catalog.json"
  if (Test-Path -LiteralPath $catalogPath) {
    try {
      $catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
      foreach ($template in @($catalog.templates)) {
        $relative = [string]$template.relativePath
        if ($relative) {
          [void]$bundled.Add(($relative -replace "\\", "/").TrimStart("/"))
        }
      }
    } catch {
      Write-SubCreatorInfo "Template catalog could not be read; preserving all existing templates from $ExtensionDir."
    }
  }

  Get-ChildItem -LiteralPath $templatesDir -Recurse -File | ForEach-Object {
    $relative = ($_.FullName.Substring($templatesDir.Length).TrimStart("\") -replace "\\", "/")
    if ($bundled.Contains($relative)) {
      return
    }

    $target = Join-Path $BackupDir $relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    Copy-Item -LiteralPath $_.FullName -Destination $target -Force
  }
}

function Restore-SubCreatorUserTemplates {
  param(
    [string]$BackupDir,
    [string]$ExtensionDir
  )

  # // Merge preserved user MOGRTs without overwriting templates from the freshly installed bundle.
  if (-not (Test-Path -LiteralPath $BackupDir)) {
    return
  }

  $templatesDir = Join-Path $ExtensionDir "templates\mogrt"
  New-Item -ItemType Directory -Path $templatesDir -Force | Out-Null
  Get-ChildItem -LiteralPath $BackupDir -Recurse -File | ForEach-Object {
    $relative = $_.FullName.Substring($BackupDir.Length).TrimStart("\")
    $target = Join-Path $templatesDir $relative
    if (Test-Path -LiteralPath $target) {
      return
    }

    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    Copy-Item -LiteralPath $_.FullName -Destination $target -Force
  }
}

function Enable-SubCreatorCepDebugMode {
  # // Enable unsigned CEP extensions for current-user Adobe hosts across recent CSXS versions.
  $writes = 0
  for ($version = 7; $version -le 20; $version += 1) {
    $key = "HKCU:\Software\Adobe\CSXS.$version"
    try {
      New-Item -Path $key -Force | Out-Null
      New-ItemProperty -Path $key -Name "PlayerDebugMode" -Value "1" -PropertyType String -Force | Out-Null
      $writes += 1
    } catch {
      Write-SubCreatorInfo "WARNING: unable to enable CEP debug mode for CSXS.$version."
    }
  }

  if ($writes -gt 0) {
    Write-SubCreatorInfo "CEP debug mode enabled for CSXS.7 to CSXS.20."
  }
}

function Copy-SubCreatorBundledModels {
  # // Copy the bundled Whisper starter model into the cache path already used by Whisper and the panel.
  if (-not (Test-Path -LiteralPath $bundledModelsDir)) {
    return
  }

  New-Item -ItemType Directory -Path $whisperCacheDir -Force | Out-Null
  $copied = 0

  Get-ChildItem -LiteralPath $bundledModelsDir -File -Filter "*.pt" | ForEach-Object {
    $target = Join-Path $whisperCacheDir $_.Name
    if (-not (Test-Path -LiteralPath $target)) {
      Copy-Item -LiteralPath $_.FullName -Destination $target -Force
      $copied += 1
    }
  }

  Get-ChildItem -LiteralPath $bundledModelsDir -File | Where-Object { $_.Name -match "\.pt\.part-\d+$" } | Group-Object { $_.Name -replace "\.part-\d+$", "" } | ForEach-Object {
    $target = Join-Path $whisperCacheDir $_.Name
    if (Test-Path -LiteralPath $target) {
      return
    }

    $stream = [System.IO.File]::Open($target, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
      $_.Group | Sort-Object Name | ForEach-Object {
        $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
        $stream.Write($bytes, 0, $bytes.Length)
      }
    } finally {
      $stream.Dispose()
    }
    $copied += 1
  }

  if ($copied -gt 0) {
    Write-SubCreatorInfo "Copied $copied bundled Whisper model(s) to $whisperCacheDir."
  }
}

function Install-SubCreatorBundledFonts {
  # // Install bundled fonts per user so the included templates render without administrator rights.
  if (-not (Test-Path -LiteralPath $bundledFontsDir)) {
    return
  }

  $targetDir = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"
  $registryPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
  New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
  New-Item -Path $registryPath -Force | Out-Null

  $installed = 0
  $skipped = 0
  $failed = 0
  Get-ChildItem -LiteralPath $bundledFontsDir -Recurse -File | Where-Object { $_.Extension -match "^\.(ttf|otf|ttc)$" } | ForEach-Object {
    $destination = Join-Path $targetDir $_.Name
    $copyNeeded = $true
    if (Test-Path -LiteralPath $destination) {
      try {
        # // Avoid overwriting an identical font because Windows may keep installed fonts memory-mapped while Adobe apps are open.
        $sourceHash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        $destinationHash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash
        if ($sourceHash -eq $destinationHash) {
          $copyNeeded = $false
          $skipped += 1
        }
      } catch {
        Write-SubCreatorInfo "WARNING: unable to compare installed font $destination."
      }
    }

    if ($copyNeeded) {
      try {
        Copy-Item -LiteralPath $_.FullName -Destination $destination -Force -ErrorAction Stop
        $installed += 1
      } catch {
        # // Keep the extension/runtime installation successful when Windows has a previously installed font locked in memory.
        $failed += 1
        Write-SubCreatorInfo "WARNING: font is currently in use and could not be updated: $destination"
        return
      }
    }

    $kind = if ($_.Extension -ieq ".otf") { "OpenType" } else { "TrueType" }
    $displayName = ([System.IO.Path]::GetFileNameWithoutExtension($_.Name) -replace "[-_]+", " ")
    New-ItemProperty -Path $registryPath -Name "$displayName ($kind)" -Value $destination -PropertyType String -Force | Out-Null
  }

  if ($installed -gt 0) {
    Write-SubCreatorInfo "Installed $installed bundled font(s) for the current Windows user."
  }
  if ($skipped -gt 0) {
    Write-SubCreatorInfo "Kept $skipped identical bundled font(s) already installed."
  }
  if ($failed -gt 0) {
    Write-SubCreatorInfo "WARNING: kept $failed locked font file(s); close Adobe applications before reinstalling to update them."
  }
}

function Install-SubCreatorPrivateRuntime {
  # // Copy the packaged private runtime into LocalAppData so Sub Creator does not depend on system Python or system FFmpeg.
  if (-not (Test-Path -LiteralPath $payloadRuntimeDir)) {
    throw "Private runtime payload is missing: $payloadRuntimeDir"
  }

  Copy-SubCreatorDirectoryFresh -Source $payloadRuntimeDir -Destination $runtimeDir
  Write-SubCreatorInfo "Private runtime installed to $runtimeDir."
}

function Write-SubCreatorRuntimeConfig {
  # // Persist exact private runtime paths for the CEP panel and host scripts.
  $pythonPath = Join-Path $runtimeDir "python\python.exe"
  $whisperPath = Join-Path $runtimeDir "python\Scripts\whisper.exe"
  $ffmpegPath = Join-Path $runtimeDir "ffmpeg\bin\ffmpeg.exe"
  $pathHints = @(
    (Split-Path -Parent $pythonPath),
    (Split-Path -Parent $whisperPath),
    (Split-Path -Parent $ffmpegPath),
    (Join-Path $env:SystemRoot "System32")
  ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique

  $pythonVersion = ""
  if (Test-Path -LiteralPath $pythonPath) {
    $pythonVersion = (& $pythonPath --version 2>&1 | Select-Object -First 1)
  }

  New-Item -ItemType Directory -Path $runtimeConfigDir -Force | Out-Null
  $config = [ordered]@{
    version = 1
    generatedBy = "subcreator_install_windows_private_runtime.ps1"
    generatedAtUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    pythonCommand = $pythonPath
    pythonLabel = "Sub Creator private Python"
    pythonPath = $pythonPath
    pythonVersion = $pythonVersion
    whisperPath = $(if (Test-Path -LiteralPath $whisperPath) { $whisperPath } else { "" })
    ffmpegPath = $(if (Test-Path -LiteralPath $ffmpegPath) { $ffmpegPath } else { "" })
    pathHints = @($pathHints)
  }

  # // Write UTF-8 without BOM so every CEP/ExtendScript JSON parser can read the runtime config.
  $configJson = $config | ConvertTo-Json -Depth 4
  [System.IO.File]::WriteAllText($runtimeConfigFile, $configJson, (New-Object System.Text.UTF8Encoding($false)))
  Write-SubCreatorInfo "Runtime config written: $runtimeConfigFile"
}

function Test-SubCreatorPrivateRuntime {
  # // Validate imports after installation so packaging problems are visible immediately.
  $pythonPath = Join-Path $runtimeDir "python\python.exe"
  $ffmpegPath = Join-Path $runtimeDir "ffmpeg\bin\ffmpeg.exe"

  if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw "Private Python is missing after install: $pythonPath"
  }
  if (-not (Test-Path -LiteralPath $ffmpegPath)) {
    throw "Private FFmpeg is missing after install: $ffmpegPath"
  }

  & $pythonPath -c "import whisper; import whisperx; print('Whisper runtime validation OK')" | Write-Host
  if ($LASTEXITCODE -ne 0) {
    throw "Private Python could not import whisper and whisperx."
  }
}

function Write-SubCreatorRuntimeVersion {
  # // Mark the validated runtime so future connected installers can skip the large runtime download.
  Set-Content -LiteralPath $runtimeVersionFile -Value $RuntimeVersion -Encoding ASCII
  Write-SubCreatorInfo "Private runtime version $RuntimeVersion is ready."
}

if (-not (Test-Path -LiteralPath $sourceDir)) {
  throw "Build missing: $sourceDir"
}

Write-SubCreatorInfo "Installing Sub Creator from $PayloadRoot"

$backupRoot = Join-Path $env:TEMP ("subcreator-mogrt-backup-" + [System.Guid]::NewGuid().ToString("N"))
$backupDir = Join-Path $backupRoot "mogrt"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

Backup-SubCreatorUserTemplates -ExtensionDir $destDir -BackupDir $backupDir
Backup-SubCreatorUserTemplates -ExtensionDir $legacyDestDir -BackupDir $backupDir
Copy-SubCreatorDirectoryFresh -Source $sourceDir -Destination $destDir
if (Test-Path -LiteralPath $legacyDestDir) {
  Remove-Item -LiteralPath $legacyDestDir -Recurse -Force
}
Restore-SubCreatorUserTemplates -BackupDir $backupDir -ExtensionDir $destDir
Remove-Item -LiteralPath $backupRoot -Recurse -Force -ErrorAction SilentlyContinue

Write-SubCreatorInfo "Sub Creator installed to $destDir."
Enable-SubCreatorCepDebugMode
Copy-SubCreatorBundledModels
Install-SubCreatorBundledFonts
if (-not $SkipRuntimeInstall) {
  Install-SubCreatorPrivateRuntime
} else {
  Write-SubCreatorInfo "Keeping the compatible private runtime already installed."
}
Write-SubCreatorRuntimeConfig
Test-SubCreatorPrivateRuntime
Write-SubCreatorRuntimeVersion

Write-SubCreatorInfo "Installation complete. Restart Premiere Pro, then open Window > Extensions > Sub Creator."
