param(
  [string]$Destination = "",
  [switch]$SkipBuild,
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"

# // Resolve the repository root from this script location so the updater works from npm, cmd, or PowerShell.
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptDir ".."))
$distDir = Join-Path $repoRoot "dist\com.cyrilplugin.subcreator"
$legacyDestination = Join-Path $env:APPDATA "Adobe\CEP\extensions\com.cyrilg93.subcreator"

# // Use the same current-user CEP folder as the modern Windows installer.
if (-not $Destination) {
  $Destination = Join-Path $env:APPDATA "Adobe\CEP\extensions\com.cyrilplugin.subcreator"
}
$Destination = [System.IO.Path]::GetFullPath($Destination)

function Write-SubCreatorInfo {
  param([string]$Message)
  # // Keep the quick-update output readable when launched from npm or the .bat helper.
  Write-Host "[Sub Creator] $Message"
}

function Invoke-SubCreatorBuild {
  # // Rebuild the TypeScript panel before copying so local tests include the latest source edits.
  if ($SkipBuild) {
    Write-SubCreatorInfo "Skipping build because -SkipBuild was provided."
    return
  }

  if ($DryRun) {
    Write-SubCreatorInfo "Would run npm.cmd run subcreator:build"
    return
  }

  Push-Location $repoRoot
  try {
    & npm.cmd run subcreator:build
    if ($LASTEXITCODE -ne 0) {
      throw "subcreator:build failed with code $LASTEXITCODE."
    }
  } finally {
    Pop-Location
  }
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

  # // Merge preserved user MOGRTs without overwriting templates from the freshly built bundle.
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
  if ($DryRun) {
    Write-SubCreatorInfo "Would enable CEP debug mode for CSXS.7 to CSXS.20"
    return
  }

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

function Copy-SubCreatorLocalBuild {
  # // Replace the installed CEP extension with the freshly built dist folder while keeping user templates.
  if (-not (Test-Path -LiteralPath $distDir -PathType Container)) {
    throw "Build missing: $distDir"
  }

  $backupRoot = Join-Path $env:TEMP ("subcreator-local-update-" + [System.Guid]::NewGuid().ToString("N"))
  $backupDir = Join-Path $backupRoot "mogrt"

  if ($DryRun) {
    Write-SubCreatorInfo "Would copy $distDir -> $Destination"
    Write-SubCreatorInfo "Would remove legacy folder $legacyDestination if present"
    return
  }

  New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
  try {
    Backup-SubCreatorUserTemplates -ExtensionDir $Destination -BackupDir $backupDir
    Backup-SubCreatorUserTemplates -ExtensionDir $legacyDestination -BackupDir $backupDir

    New-Item -ItemType Directory -Path (Split-Path -Parent $Destination) -Force | Out-Null
    if (Test-Path -LiteralPath $Destination) {
      Remove-Item -LiteralPath $Destination -Recurse -Force
    }
    Copy-Item -LiteralPath $distDir -Destination $Destination -Recurse -Force

    if (Test-Path -LiteralPath $legacyDestination) {
      Remove-Item -LiteralPath $legacyDestination -Recurse -Force
    }

    Restore-SubCreatorUserTemplates -BackupDir $backupDir -ExtensionDir $Destination
  } finally {
    Remove-Item -LiteralPath $backupRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}

Write-SubCreatorInfo "Updating local CEP plugin from $repoRoot"
Write-SubCreatorInfo "Destination: $Destination"
Invoke-SubCreatorBuild
Copy-SubCreatorLocalBuild
Enable-SubCreatorCepDebugMode
Write-SubCreatorInfo "Local update complete. Restart Premiere Pro, then open Window > Extensions > Sub Creator."
