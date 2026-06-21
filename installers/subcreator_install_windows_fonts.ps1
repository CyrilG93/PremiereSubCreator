param(
  [Parameter(Mandatory = $true)]
  [string]$FontsDir
)

$ErrorActionPreference = "Stop"

# // Install per-user fonts without overwriting files that Adobe or Windows may have memory-mapped.
$sourceRoot = [System.IO.Path]::GetFullPath($FontsDir)
$targetRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"
$registryPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"

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

function Remove-SubCreatorLegacyFontRegistration {
  param(
    [System.IO.FileInfo]$SourceFile,
    [string]$CorrectRegistryName,
    [string]$FontKind
  )

  # // Remove only the old Sub Creator entry inferred from this exact filename; never delete font files or unrelated registrations.
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
    $legacyPath = [System.IO.Path]::GetFullPath([string]$legacyValue)
    $expectedLegacyPath = Join-Path $targetRoot $SourceFile.Name
    if ($legacyPath -ieq $expectedLegacyPath) {
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
$failed = 0
$registeredNames = New-Object System.Collections.Generic.List[string]

Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
  Where-Object { $_.Extension -match "^\.(ttf|otf|ttc)$" } |
  ForEach-Object {
    $sourceFile = $_
    try {
      # // Content-addressed filenames avoid replacing or deleting a font that is already loaded by Premiere.
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

      $fontTitle = Get-SubCreatorFontTitle -FontFile $sourceFile -ShellApplication $shellApplication
      $fontKind = if ($sourceFile.Extension -ieq ".otf") { "OpenType" } else { "TrueType" }
      $registryName = "$fontTitle ($fontKind)"
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
Write-Host "SUBCREATOR_FONTS_FAILED=$failed"
Write-Host ("SUBCREATOR_FONT_NAMES=" + (($registeredNames | Sort-Object -Unique) -join ", "))
if ($failed -gt 0) {
  exit 1
}
exit 0
