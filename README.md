# Sub Creator

Sub Creator is an Adobe Premiere Pro extension for creating subtitles from an `SRT` file, Whisper transcription, WhisperX transcription, or corrected text. It can generate animated `MOGRT` subtitles or native Premiere subtitle tracks.

Compatible with Premiere Pro `2025+` on Windows and Apple Silicon macOS. Premiere Pro 24 and earlier, and Intel Macs, are not supported.

## Main features

- Generate subtitles from `SRT`, `Whisper`, `WhisperX`, or `Whisper + SRT`.
- Create animated `MOGRT` clips or native Premiere subtitle tracks.
- Align subtitle timing to sequence frames to avoid one-frame gaps.
- Inspect and copy exposed MOGRT settings with the `Visual editor`.
- Review, move, split, merge, and rebuild subtitles with the `Text editor`.
- Follow Premiere Pro's Light, Dark, and Darkest appearance settings.

## Installation

Sub Creator is distributed only as complete installers. Python, FFmpeg, Whisper, WhisperX, and the required dependencies are kept in a private runtime and do not need to be installed separately.

### Windows

1. Download `SubCreator-v...-Windows-Full-Installer.exe`.
2. Close Premiere Pro.
3. Run the installer.
4. Choose the Whisper models you want to keep available offline.
5. Wait for the installation summary.
6. Reopen Premiere Pro.
7. Open `Window > Extensions > Sub Creator`.

The Full installer embeds the private runtime and the recommended `base` model, so the default installation works without Python, FFmpeg, or an internet connection. Extra selected models are downloaded only when they are missing or damaged. Existing models and custom MOGRT files are preserved when the Full installer is run again.

The installer is intentionally unsigned, so Windows may display a SmartScreen warning. After clicking `Install`, Setup can also appear frozen briefly while Windows checks the files. Leave any Command Prompt or PowerShell windows open until Setup finishes.

The installer automatically enables the CEP debug mode required for unsigned extensions. If the release includes a separate `Fonts` folder, install those fonts manually when you want the bundled templates to match their intended typography.

### Apple Silicon macOS

1. Download `SubCreator-v...-macOS-Installer-arm64.pkg`.
2. Close Premiere Pro.
3. Open the `.pkg`.
4. Choose the Whisper models you want to keep available offline.
5. Complete the installation.
6. Reopen Premiere Pro.
7. Open `Window > Extensions > Sub Creator`.

The Full `.pkg` embeds the Apple Silicon private Python and LGPL FFmpeg runtime. Selected Whisper models may need an internet connection the first time they are installed. Running a newer Full installer preserves compatible runtime files, downloaded Whisper models, and custom MOGRT files.

The installer is intentionally unsigned. If macOS blocks it, Control-click the package, choose `Open`, then confirm the installation. Intel Macs are not supported.

## Whisper models

The installers offer `tiny`, `base`, `small`, `medium`, and `large-v3`. Existing valid models are never removed.

- `tiny`: very fast, less accurate.
- `base`: good starting point and embedded in the Windows Full installer.
- `small`: best quality and speed balance.
- `medium`: recommended for difficult audio, accents, or background noise.
- `large-v3`: best accuracy, slower and heavier.

If transcription is unstable, manually select the language in `Whisper language` instead of using `Auto detect`.

Sub Creator renders temporary transcription audio directly in Premiere Pro. Adobe Media Encoder is used automatically only if Premiere cannot complete the direct export.

The log panel keeps a timestamped history of generation stages. If `Generate` cannot start in a Whisper mode, it shows the detected runtime, installed models, and cache paths. Share this diagnostic block when requesting support.

## Usage

### Basic workflow

1. Open `Sub Creator` in Premiere Pro.
2. In `Creation`, choose `SRT`, `Whisper`, `WhisperX`, or `Whisper + SRT`.
3. Choose `Entire sequence` or `In/Out points`.
4. Choose `MOGRT` or `Premiere subtitles`.
5. If you use `MOGRT`, choose a template from the gallery.
6. Adjust the text limits if needed.
7. Click `Generate subtitles`.

Sub Creator remembers separate generation settings for `MOGRT` and `Premiere subtitles` output.

### Sources

- `SRT`: uses an existing subtitle file.
- `Whisper (fast)`: transcribes the active sequence audio.
- `WhisperX (precise)`: transcribes with Whisper, then improves subtitle timing.
- `Whisper + SRT (corrected)`: aligns corrected text with the audio.

### Outputs

- `MOGRT`: creates animated graphic clips using a template.
- `Premiere subtitles`: creates a native Premiere subtitle track.

When native subtitles create an SRT source file, Sub Creator saves it in an `SRT` folder next to the current `.prproj` file.

### After generation

- Use `Visual editor` to inspect a selected MOGRT or Premiere graphic, edit exposed settings live, and copy styles between clips. Standard video and audio clips are ignored by its automatic selection refresh.
- Use `Text editor` to correct, move, split, or merge generated subtitles.
- Add your own `.mogrt` files from `Open MOGRT folder`, then click `Refresh gallery`.

## Limitations

- Premiere does not expose every MOGRT setting through the extension API.
- Some settings remain more reliable in Premiere's `Properties` panel.
- MOGRTs created in After Effects are usually more predictable than MOGRTs created directly in Premiere.
- Whisper can make mistakes with noisy audio, unclear voices, strong accents, or less supported languages.
- WhisperX may need to download an alignment model the first time it is used.

## Troubleshooting

If generation fails with an `EvalScript error`, open the debug log in Sub Creator and share the `Host result` details. The log includes the Premiere host function name and response details to help identify the failure.

## Support

- Website: https://www.cyrilplugin.com/
- Discord: https://www.cyrilplugin.com/website/social/discord

## Changelog

### 1.1.53 - 2026-07-31 (Pre-release)

- Added complete Windows and Apple Silicon macOS installers with the private transcription runtime included.
- Improved MOGRT timing, frame alignment, subtitle rebuilding, and long-generation reliability.
- Added automatic Visual editor refresh, Premiere theme matching, clearer progress, and better diagnostics.
- Reduced background activity by ignoring regular media selections and pausing Visual editor monitoring while the panel is hidden.

### 1.1.0 - 2026-05-11

- Improved MOGRT generation and preserved template keyframes.
- Added the Montserrat font family to the provided resources.

### 1.0.0 - 2026-05-07

- First stable release of Sub Creator.
- Includes MOGRT generation, native Premiere subtitles, Whisper, WhisperX, visual editing, and text editing.
