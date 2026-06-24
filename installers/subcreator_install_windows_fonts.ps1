param(
  [Parameter(Mandatory = $true)]
  [string]$FontsDir
)

$ErrorActionPreference = "Stop"

# // Install per-user fonts without overwriting files that Adobe or Windows may have memory-mapped.
$sourceRoot = [System.IO.Path]::GetFullPath($FontsDir)
$targetRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"
$registryPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
$machineRegistryPath = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
$systemFontsRoot = Join-Path $env:WINDIR "Fonts"

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class SubCreatorFontApi {
  [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  public static extern int AddFontResourceEx(string name, uint flags, IntPtr reserved);

  [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
  public static extern IntPtr SendMessageTimeout(
    IntPtr window,
    uint message,
    IntPtr wordParameter,
    IntPtr longParameter,
    uint flags,
    uint timeout,
    out IntPtr result
  );
}
"@

function Get-SubCreatorFontTitle {
  param(
    [System.IO.FileInfo]$FontFile,
    $ShellApplication
  )

  # // Read the font's internal full name instead of guessing it from a sometimes unrelated filename.
  try {
    $shellFolder = $ShellApplication.Namespace($FontFile.DirectoryName)
    $shellItem = $shellFolder.ParseName($FontFile.Name)
    $title = [string]$shellItem.ExtendedProperty("System.Title")
    if ($title.Trim()) {
      return $title.Trim()
    }
  } catch {}

  return ([System.IO.Path]::GetFileNameWithoutExtension($FontFile.Name) -replace "[-_]+", " ").Trim()
}

function Resolve-SubCreatorFontRegistryValue {
  param(
    [string]$RegistryValue,
    [string]$DefaultRoot
  )

  # // Normalize registry values so existing user/system fonts can be recognized without changing their registration.
  if (-not $RegistryValue) {
    return ""
  }

  $expandedValue = [System.Environment]::ExpandEnvironmentVariables($RegistryValue)
  if ([System.IO.Path]::IsPathRooted($expandedValue)) {
    return [System.IO.Path]::GetFullPath($expandedValue)
  }

  return [System.IO.Path]::GetFullPath((Join-Path $DefaultRoot $expandedValue))
}

function Test-SubCreatorManagedFontPath {
  param([string]$FontPath)

  # // Treat only content-addressed Sub Creator files as installer-owned; every other font path belongs to the user or Windows.
  if (-not $FontPath) {
    return $false
  }

  try {
    $fullPath = [System.IO.Path]::GetFullPath($FontPath)
    $root = [System.IO.Path]::GetFullPath($targetRoot).TrimEnd("\") + "\"
    $fileName = [System.IO.Path]::GetFileName($fullPath)
    return $fullPath.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) -and
      $fileName.StartsWith("SubCreator-", [System.StringComparison]::OrdinalIgnoreCase)
  } catch {
    return $false
  }
}

function Get-SubCreatorExistingFontRegistration {
  param([string]$RegistryName)

  # // Check both per-user and system-wide registrations before adding a Sub Creator copy with the same internal font name.
  $locations = @(
    @{ Scope = "current-user"; Path = $registryPath; DefaultRoot = $targetRoot },
    @{ Scope = "system"; Path = $machineRegistryPath; DefaultRoot = $systemFontsRoot }
  )

  foreach ($location in $locations) {
    $properties = Get-ItemProperty -Path $location.Path -Name $RegistryName -ErrorAction SilentlyContinue
    if (-not $properties) {
      continue
    }

    $value = [string]$properties.$RegistryName
    if (-not $value) {
      continue
    }

    $resolvedPath = Resolve-SubCreatorFontRegistryValue -RegistryValue $value -DefaultRoot $location.DefaultRoot
    return [PSCustomObject]@{
      Scope = $location.Scope
      RegistryPath = $location.Path
      Value = $value
      ResolvedPath = $resolvedPath
      IsManaged = Test-SubCreatorManagedFontPath -FontPath $resolvedPath
    }
  }

  return $null
}

function Remove-SubCreatorLegacyFontRegistration {
  param(
    [System.IO.FileInfo]$SourceFile,
    [string]$CorrectRegistryName,
    [string]$FontKind
  )

  # // Preserve legacy entries unless they clearly point to a content-addressed Sub Creator file; never touch user-owned font registrations.
  $legacyDisplayName = ([System.IO.Path]::GetFileNameWithoutExtension($SourceFile.Name) -replace "[-_]+", " ")
  $legacyRegistryName = "$legacyDisplayName ($FontKind)"
  if ($legacyRegistryName -eq $CorrectRegistryName) {
    return
  }

  $legacyValue = (Get-ItemProperty -Path $registryPath -Name $legacyRegistryName -ErrorAction SilentlyContinue).$legacyRegistryName
  if (-not $legacyValue) {
    return
  }

  try {
    $legacyPath = Resolve-SubCreatorFontRegistryValue -RegistryValue ([string]$legacyValue) -DefaultRoot $targetRoot
    if (Test-SubCreatorManagedFontPath -FontPath $legacyPath) {
      Remove-ItemProperty -Path $registryPath -Name $legacyRegistryName -Force
    }
  } catch {}
}

if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
  throw "Bundled font folder is missing: $sourceRoot"
}

New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null
New-Item -Path $registryPath -Force | Out-Null
$shellApplication = New-Object -ComObject Shell.Application
$installed = 0
$kept = 0
$preserved = 0
$failed = 0
$registeredNames = New-Object System.Collections.Generic.List[string]

Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
  Where-Object { $_.Extension -match "^\.(ttf|otf|ttc)$" } |
  ForEach-Object {
    $sourceFile = $_
    try {
      $fontTitle = Get-SubCreatorFontTitle -FontFile $sourceFile -ShellApplication $shellApplication
      $fontKind = if ($sourceFile.Extension -ieq ".otf") { "OpenType" } else { "TrueType" }
      $registryName = "$fontTitle ($fontKind)"
      $existingRegistration = Get-SubCreatorExistingFontRegistration -RegistryName $registryName

      if ($existingRegistration -and -not $existingRegistration.IsManaged) {
        # // A user or Windows font with the same internal name already exists; keep it as the authoritative registration.
        $preserved += 1
        if (Test-Path -LiteralPath $existingRegistration.ResolvedPath -PathType Leaf) {
          [void][SubCreatorFontApi]::AddFontResourceEx($existingRegistration.ResolvedPath, 0, [IntPtr]::Zero)
        }
        [void]$registeredNames.Add($fontTitle)
        Write-Host ("SUBCREATOR_FONT_PRESERVED=" + $fontTitle + " from " + $existingRegistration.Scope)
        return
      }

      # // Content-addressed filenames avoid replacing or deleting a Sub Creator font that is already loaded by Premiere.
      $sourceHash = (Get-FileHash -LiteralPath $sourceFile.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
      $safeBaseName = ([System.IO.Path]::GetFileNameWithoutExtension($sourceFile.Name) -replace "[^A-Za-z0-9._-]", "-")
      $targetName = "SubCreator-$safeBaseName-$($sourceHash.Substring(0, 12))$($sourceFile.Extension.ToLowerInvariant())"
      $destination = Join-Path $targetRoot $targetName
      if (Test-Path -LiteralPath $destination -PathType Leaf) {
        $destinationHash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($destinationHash -ne $sourceHash) {
          throw "Existing managed font has an unexpected hash: $destination"
        }
        $kept += 1
      } else {
        Copy-Item -LiteralPath $sourceFile.FullName -Destination $destination -ErrorAction Stop
        $installed += 1
      }

      Remove-SubCreatorLegacyFontRegistration -SourceFile $sourceFile -CorrectRegistryName $registryName -FontKind $fontKind
      New-ItemProperty -Path $registryPath -Name $registryName -Value $destination -PropertyType String -Force | Out-Null

      # // Register the persistent per-user font in the current logon session so newly opened Adobe apps can enumerate it immediately.
      $fontCount = [SubCreatorFontApi]::AddFontResourceEx($destination, 0, [IntPtr]::Zero)
      if ($fontCount -lt 1) {
        throw "Windows could not load the font into the current session: $fontTitle"
      }
      [void]$registeredNames.Add($fontTitle)
    } catch {
      $failed += 1
      Write-Warning ("Font installation failed for " + $sourceFile.FullName + ": " + $_.Exception.Message)
    }
  }

# // Tell desktop applications that the session font table changed; Premiere must still be restarted if it was already open.
$broadcastResult = [IntPtr]::Zero
[void][SubCreatorFontApi]::SendMessageTimeout(
  [IntPtr]0xFFFF,
  0x001D,
  [IntPtr]::Zero,
  [IntPtr]::Zero,
  0x0002,
  5000,
  [ref]$broadcastResult
)

Write-Host "SUBCREATOR_FONTS_INSTALLED=$installed"
Write-Host "SUBCREATOR_FONTS_SKIPPED=$kept"
Write-Host "SUBCREATOR_FONTS_PRESERVED=$preserved"
Write-Host "SUBCREATOR_FONTS_FAILED=$failed"
Write-Host ("SUBCREATOR_FONT_NAMES=" + (($registeredNames | Sort-Object -Unique) -join ", "))
if ($failed -gt 0) {
  exit 1
}
exit 0
