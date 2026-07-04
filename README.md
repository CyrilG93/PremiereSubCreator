# Sub Creator

Sub Creator is an Adobe Premiere Pro extension for creating subtitles from an `SRT` file, Whisper transcription, WhisperX transcription, or corrected text.
It can generate animated `MOGRT` subtitles or native Premiere subtitle tracks.

Compatible with Premiere Pro `2025+` on Windows and macOS. It's not compatible with Premiere 24 or earlier.

## Main Features

- `Creation`: generate subtitles from `SRT`, `Whisper`, `WhisperX`, or `Whisper + SRT`, with `MOGRT` or `Premiere subtitles` output.
- MOGRT subtitle timing is aligned to sequence frames so adjacent clips share one exact boundary without one-frame visual gaps.
- `Visual editor`: read exposed controls from a selected MOGRT and apply the same look to other selected clips.
- `Text editor`: review, move, split, merge, and rebuild generated subtitle blocks.

## Installation

### Dependencies

Sub Creator works without extra dependencies if you only use `SRT` mode.

On Windows, use the full `.exe` for a first installation or the light `.exe` for a smaller connected installation. On macOS, use the unified `.pkg` installer. These installers configure the private Python, FFmpeg, Whisper tools and all dependencies automatically.

If you use the manual `.zip` installer, `Whisper`, `WhisperX`, and `Whisper + SRT` also need:

- Python `3.11` or `3.12`: [python.org/downloads](https://www.python.org/downloads/)
- FFmpeg: [ffmpeg.org/download.html](https://ffmpeg.org/download.html)

The Sub Creator installer then configures the required Python tools when possible.

Quick tips:

- Windows: during Python installation, enable `Add python.exe to PATH`.
- Windows: keep the `py launcher` if the Python installer offers it.
- macOS ZIP installer: install Python and FFmpeg before running the Sub Creator installer again.
- If Python or FFmpeg is installed after Sub Creator, run the Sub Creator installer again.

### Install on Windows with the `.exe`

1. Download `SubCreator-v...-Windows-Full-Installer.exe` for a first installation, or the smaller `Light` installer if the computer is online.
2. Close Premiere Pro.
3. Run the installer.
4. Choose the Whisper models you want to keep available offline.
5. Wait for the installation summary.
6. Reopen Premiere Pro.
7. Open the extension from `Window > Extensions > Sub Creator`.

The full installer embeds the private runtime and the recommended `base` model, so a first installation works without Python, FFmpeg, or an internet connection. The light installer downloads the private runtime and selected models when they are missing, so it needs an internet connection for a first installation. Later updates keep a compatible runtime and existing Whisper models. The private Python and LGPL FFmpeg runtime stays inside the current Windows user profile and does not modify the system Python installation.

After clicking `Install`, Setup can appear frozen for a few seconds while Windows checks the files. During installation, leave any Command Prompt or PowerShell windows open until Setup finishes.

### Install on Windows with the `.zip`

1. Download the `.zip` from the latest version.
2. Unzip it.
3. Close Premiere Pro.
4. Run `subcreator_install_windows.bat`.
5. Wait for the installation summary in the console.
6. Reopen Premiere Pro.
7. Open the extension from `Window > Extensions > Sub Creator`.

The Windows installer automatically enables the CEP debug mode required for unsigned extensions. If the release includes a separate `Fonts` folder, install those fonts manually when you want the bundled templates to match their intended typography.

### Repair or update Windows dependencies

1. Close Premiere Pro.
2. Run `subcreator_update_windows_dependencies.bat`.
3. Wait for the download, integrity check, and validation to finish.
4. Reopen Premiere Pro.

This tool reinstalls the tested private Python, WhisperX, and FFmpeg runtime used by Sub Creator. It preserves downloaded Whisper models and does not modify the system Python installation. An internet connection is required for the approximately 333 MB download.

### Quick local Windows update for testing

To test panel changes without rebuilding the `.exe`, run:

```powershell
npm.cmd run subcreator:update:local:windows
```

You can also double-click `UPDATE_LOCAL_WINDOWS.bat`. Restart Premiere Pro after the copy.

### Install on macOS with the `.pkg`

1. Download `SubCreator-v...-macOS-Installer-arm64.pkg` for Apple Silicon or `SubCreator-v...-macOS-Installer-x86_64.pkg` for an Intel Mac.
2. Close Premiere Pro.
3. Open the `.pkg`.
4. Choose the Whisper models you want to keep available offline.
5. Complete the installation.
6. Reopen Premiere Pro.
7. Open the extension from `Window > Extensions > Sub Creator`.

The `.pkg` includes the private Python and LGPL FFmpeg runtime for the selected Mac architecture, so it does not depend on a separate runtime download. Existing compatible runtimes and Whisper models are preserved. The installer also enables CEP debug mode and installs the bundled fonts for the current macOS user.

### Update an existing macOS installation

1. Download `SubCreator-v...-macOS-Update.pkg`.
2. Close Premiere Pro.
3. Open the `.pkg` and complete the installation.
4. Reopen Premiere Pro.

The update package works on Apple Silicon and Intel Macs. It updates only the extension and bundled fonts, while preserving the private runtime, downloaded Whisper models, and custom MOGRT files. Use the full architecture-specific installer for the first installation.

### Install on macOS with the `.zip`

1. Download and unzip the latest `.zip`.
2. Close Premiere Pro.
3. Run `subcreator_install_mac.sh`.
4. If needed, drag the `.sh` file into Terminal and press Enter.
5. Reopen Premiere Pro.
6. Open the extension from `Window > Extensions > Sub Creator`.

The manual ZIP installer uses compatible Python and FFmpeg installations already available on the Mac.

## Whisper

The Windows `.exe` and macOS `.pkg` offer `tiny`, `base`, `small`, `medium`, and `large-v3` during installation. The Windows Full installer already includes `base`; other selected models are downloaded only when missing or damaged. Existing models are never removed.

The `.zip` installers include the `base` model. It is enough for testing and many simple projects.

Sub Creator renders the temporary transcription audio directly in Premiere Pro for faster generation. Adobe Media Encoder is used automatically only if Premiere cannot complete the direct export.

To add other models, download a `.pt` file and place it in the Whisper cache folder:

- macOS: `~/.cache/whisper/`
- Windows: `%USERPROFILE%\.cache\whisper\`

Downloads:

- `tiny.pt`: [Download](https://openaipublic.azureedge.net/main/whisper/models/65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9/tiny.pt)
- `base.pt`: [Download](https://openaipublic.azureedge.net/main/whisper/models/ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e/base.pt)
- `small.pt`: [Download](https://openaipublic.azureedge.net/main/whisper/models/9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794/small.pt)
- `medium.pt`: [Download](https://openaipublic.azureedge.net/main/whisper/models/345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1/medium.pt)
- `large-v3.pt`: [Download](https://openaipublic.azureedge.net/main/whisper/models/e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb/large-v3.pt)

Which model to choose:

- `tiny`: very fast, less accurate.
- `base`: good starting point.
- `small`: best quality/speed balance.
- `medium`: recommended for difficult audio, accents, or background noise.
- `large-v3`: best accuracy, slower and heavier.
- `turbo`: fast option when speed matters most.

If transcription is unstable, manually select the language in `Whisper language` instead of using `Auto detect`.

The log panel keeps a timestamped history of generation stages. If `Generate` cannot start in a Whisper mode, it also shows the detected runtime, installed models, and cache paths. Share this diagnostic block when requesting support.

## Usage

### Basic Workflow

1. Open `Sub Creator` in Premiere Pro.
2. In `Creation`, choose the source: `SRT`, `Whisper`, `WhisperX`, or `Whisper + SRT`.
3. Choose the range: `Entire sequence` or `In/Out points`.
4. Choose the output type: `MOGRT` or `Premiere subtitles`.
5. If you use `MOGRT`, choose a template from the gallery.
6. Adjust text limits if needed.
7. Click `Generate subtitles`.

Sub Creator remembers separate generation settings for `MOGRT` output and `Premiere subtitles` output.

### Available Sources

- `SRT`: uses an existing subtitle file.
- `Whisper (fast)`: transcribes the active sequence audio.
- `WhisperX (precise)`: transcribes with Whisper, then improves subtitle timing.
- `Whisper + SRT (corrected)`: uses corrected text and aligns it with the audio.

### Available Outputs

- `MOGRT`: creates animated graphic clips using a template.
- `Premiere subtitles`: creates a native Premiere subtitle track.

When `Premiere subtitles` creates an SRT source file, Sub Creator saves it in an `SRT` folder next to the current `.prproj` file.

### After Generation

- Use `Visual editor` to inspect the selected MOGRT automatically, edit its visual settings live, and copy styles between MOGRT clips.
- Use `Text editor` to correct, move, split, or merge generated subtitles.
- After Effects MOGRT clips keep the template's available source duration and leave Time Remapping disabled, so their right edge can be extended manually when subtitle timing needs correction.
- Add your own `.mogrt` files from `Open MOGRT folder`, then click `Refresh gallery`.

## Limitations

- Premiere does not expose every MOGRT setting through the API used by the extension.
- Some settings are still more reliable in Premiere's `Properties` panel.
- MOGRTs created in After Effects are usually more predictable than MOGRTs created directly in Premiere.
- Whisper can make mistakes, especially with noisy audio, unclear voices, strong accents, or less supported languages.
- WhisperX may need to download an alignment model the first time it is used.

## Support

- Website: https://www.cyrilplugin.com/
- Discord: https://www.cyrilplugin.com/website/social/discord

## Troubleshooting

If generation fails with an `EvalScript error`, open the debug log in Sub Creator and share the `Host result` details. The log now includes the Premiere host function name and response details, which helps identify whether Premiere needs a restart, the extension was installed while Premiere was open, or the host call failed inside Premiere.

## Changelog

### 1.1.41 - 2026-06-29

- Text editor merges now keep the subtitles exactly as edited, even when the text is longer than the creation word limit.

### 1.1.40 - 2026-06-26

- Windows installers no longer install bundled fonts automatically.
- Bundled fonts are kept as a separate release folder for manual installation.

### 1.1.39 - 2026-06-24

- Visual editor buttons now use the full panel width.
- Color live updates now rebuild the Premiere selection refresh more reliably.

### 1.1.38 - 2026-06-24

- Visual editor is simpler and keeps live updates always enabled.
- Color changes now force an extra Premiere refresh so timeline previews update faster.

### 1.1.37 - 2026-06-24

- Visual editor now refreshes automatically when selecting MOGRT clips in the timeline.

### 1.1.36 - 2026-06-24

- The generation progress bar now appears immediately after clicking Generate subtitles.

### 1.1.35 - 2026-06-24

- Windows installers now preserve fonts that were already installed on the user's machine.

### 1.1.34 - 2026-06-22

- Fixed host calls on Premiere installations where ExtendScript does not provide `JSON`.
- Whisper sequence export diagnostics now work before the audio export starts.

### 1.1.33 - 2026-06-22

- Active sequence audio export keeps Premiere direct export first and improves AME fallback recovery.
- Export errors now include extra Premiere capability diagnostics.

### 1.1.32 - 2026-06-22

- Active sequence audio export is more reliable when Premiere rejects a host response.
- Export failures now include a detailed debug payload for troubleshooting.

### 1.1.25 - 2026-06-18

- Fixed Whisper generation from the active sequence on Premiere Pro 26.x.

### 1.1.19 - 2026-06-12

- Whisper sequence export now includes its own WAV preset and no longer depends on localized Adobe preset names.

### 1.1.18 - 2026-06-12

- Windows updates no longer fail when an included font is already loaded by another application.

### 1.1.17 - 2026-06-12

- Fixed Whisper detection after a first Windows installation.
- Improved recovery of the bundled private runtime and Windows model cache path.

### 1.1.5 - 2026-06-07

- Added a lightweight Windows installer that reuses the existing private runtime during updates.
- Added optional Whisper model downloads directly in the Windows installer.

### 1.1.4 - 2026-06-03

- Switched the Windows private-runtime installer to an LGPL FFmpeg build.

### 1.1.3 - 2026-06-03

- Added a Windows `.exe` installer option with a private Python and FFmpeg runtime.

### 1.1.1 - 2026-05-12

- Improved Premiere host error messages during Whisper sequence export.
- Fixed a Windows installer message about bundled Whisper model validation.

### 1.1.0 - 2026-05-11

- The generation of MOGRT created in Premiere now keep keyframes.
- Windows and macOS installers now install bundled template fonts automatically for the user.
- Added Montserrat font family in the Fonts folder.

### 1.0.0 - 2026-05-07

- First stable release of Sub Creator.
- Includes MOGRT generation, native Premiere subtitles, Whisper, WhisperX, and Whisper + SRT workflows.
- Improved installers, project-side SRT organization, visual editing, text editing, and subtitle timing behavior.

### 0.16.0 - 2026-05-05

- Native `Premiere subtitles` output now saves generated SRT files in an `SRT` folder next to the Premiere project and imports them into an `SRT` bin in Premiere.
- WhisperX startup is more reliable on Windows when multiple Python launchers are present.
- Very small gaps between generated subtitles are closed for smoother playback.

### 0.15.4 - 2026-05-04

- Added native `Premiere subtitles` output.
- Generation settings are now separated between `MOGRT` and native subtitle output.
- Base templates were updated with longer versions to avoid subtitle cutoffs.

### 0.14.0 - 2026-04-27

- Added `WhisperX` mode for more precise timing.
- WhisperX uses the selected local Whisper model before alignment.
- WhisperX subtitles are better protected against clips that are too short.

### 0.13.0 - 2026-04-24

- More reliable MOGRT generation on long sequences.
- Better progress display during Whisper transcription.
- Migrated to the `com.cyrilplugin.subcreator` extension ID.

### 0.12.0 - 2026-04-22

- Improved the `Visual editor`.
- Added the `Copy properties` workflow.
- More reliable font copying from Premiere settings.

### 0.11.0 - 2026-04-19

- Improved punctuation and apostrophe cleanup.
- More stable `Text editor` on Windows.
- More robust Windows installer.

### 0.10.0 - 2026-04-17

- Cleaner MOGRT gallery.
- Separate limits for letters, words, and lines.
- More reliable macOS and Windows installation.

### 0.9.0 - 2026-04-13

- New base template pack.
- Better support for After Effects templates with `Clip Duration`.
- Stability improvements for Whisper and the editors.

### 0.8.0 - 2026-03-24

- Better support for MOGRTs created in Premiere.
- Clearer `Visual editor`.
- More usable lists and menus inside the panel.

### 0.7.0 - 2026-03-20

- Added support for MOGRT templates created in Premiere Pro.
- Better Premiere text reading in the `Text editor`.

### 0.6.0 - 2026-03-19

- Added `Whisper + SRT`.
- Improved timing and editing workflows.

### 0.5.0 - 2026-03-18

- Safer `Text editor` rebuild.
- Better reuse of word-level timings.

### 0.4.0 - 2026-03-17

- First `Text editor`.
- Added split, merge, and retiming tools.

### 0.3.0 - 2026-03-16

- Added Whisper workflow from the active sequence.
- Added `In/Out points` support.

### 0.2.0 - 2026-03-16

- Added the MOGRT gallery.
- Preserved custom templates during updates.

### 0.1.0 - 2026-03-13

- First public version for Premiere Pro 2025+.
