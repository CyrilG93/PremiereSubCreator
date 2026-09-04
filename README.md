# Sub Creator

Sub Creator is an Adobe Premiere Pro extension for creating subtitles from an `SRT` file, Whisper transcription, WhisperX transcription, or corrected text. It can generate animated `MOGRT` subtitles or native Premiere subtitle tracks.

Compatible with Premiere Pro `2025+` on Windows and Apple Silicon macOS. Premiere Pro 24 and earlier, and Intel Macs, are not supported.

## Main features

- Generate subtitles from `SRT`, `Whisper`, `WhisperX`, or `Whisper + SRT`.
- Translate selected Sub Creator MOGRT subtitles or the SRT generated for native Premiere subtitles with a personal DeepL API Free key, then duplicate them on a new synchronized track.
- Keep names, brands, and technical terms consistent with a global Whisper dictionary shared by all Premiere projects.
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

The Full installer embeds the private runtime and the recommended `base` model, so the default installation works without Python, FFmpeg, or an internet connection. Valid models already in the local cache are preselected and marked `already installed`. Extra selected models are downloaded only when they are missing or damaged. Existing models and custom MOGRT files are preserved when the Full installer is run again.

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

The Full `.pkg` embeds the Apple Silicon private Python and LGPL FFmpeg runtime. Valid models already in the current user's Whisper cache are preselected. Selected Whisper models may need an internet connection the first time they are installed. Running a newer Full installer preserves compatible runtime files, downloaded Whisper models, and custom MOGRT files.

Signed macOS installers are signed and notarized by Apple, so they should open normally. If you use an older unsigned installer and macOS blocks it, Control-click the package, choose `Open`, then confirm the installation. Intel Macs are not supported.

## Whisper models

The installers offer `tiny`, `base`, `small`, `medium`, and `large-v3`. Existing valid models are never removed.

- `tiny`: very fast, less accurate.
- `base`: good starting point and embedded in the Windows Full installer.
- `small`: best quality and speed balance.
- `medium`: recommended for difficult audio, accents, or background noise.
- `large-v3`: best accuracy, slower and heavier.

If transcription is unstable, manually select the language in `Whisper language` instead of using `Auto detect`.

## Global Whisper dictionary

When using `Whisper` or `WhisperX`, enable `Use global Whisper dictionary` to guide transcription and enforce exact spellings. Add one canonical spelling per line:

```text
Adobe Premiere Pro
Cyril Plugin
WhisperX
```

If Whisper commonly produces a specific mistake, map the heard or incorrect variant to the exact spelling:

```text
adobe premiere => Adobe Premiere Pro
serial plugin | cyril plug-in => Cyril Plugin
```

Separate multiple variants with `|`. Sub Creator first gives the canonical spellings to Whisper, then applies the explicit corrections while preserving subtitle timing. The dictionary is stored in the current user profile and remains available when switching Premiere projects or updating the extension.

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
7. For Whisper sources, optionally enable the global dictionary and add names or exact spellings.
8. Optionally enable `Remove punctuation` to remove commas, periods, and other punctuation while keeping apostrophes in words such as `J'aime` or `I'm`.
9. Click `Generate subtitles`.

Sub Creator keeps punctuation by default and remembers separate generation settings for `MOGRT` and `Premiere subtitles` output.

### Sources

- `SRT`: uses an existing subtitle file.
- `Whisper (fast)`: transcribes the active sequence audio.
- `WhisperX (precise)`: transcribes with Whisper, then improves subtitle timing.
- `Whisper + SRT (corrected)`: aligns corrected text with the audio.

### Outputs

- `MOGRT`: creates animated graphic clips using a template.
- `Premiere subtitles`: creates a native Premiere subtitle track.

### Translation

The `Translation` tab translates either selected Sub Creator MOGRT subtitle clips or a source SRT for native Premiere subtitles, without changing the original timing. Choose source and target languages, enter your personal DeepL API Free key, then review or correct each translated subtitle before creating the translated track.

The key is saved locally in Sub Creator's CEP profile on your computer and is sent only to DeepL for translation. Once a key is entered, the source and target language lists are loaded from DeepL; use `Refresh DeepL languages` to update them on demand. The translated text is sent to DeepL; the subtitle timings and visual style stay local. When Sub Creator creates `Premiere subtitles`, it automatically loads the exact generated SRT into `Translation`, ready for DeepL. For an existing native Premiere caption track created outside Sub Creator, choose `SRT file / native Premiere subtitles` and select its original SRT, because Premiere's CEP API does not reliably expose native caption text or export the selected track.

When native subtitles create an SRT source file, Sub Creator saves it in an `SRT` folder next to the current `.prproj` file.

### After generation

- Use `Visual editor` to inspect a selected MOGRT or Premiere graphic, edit exposed settings live, and copy styles between clips. Large selections that share the same source graphic are read and applied together while remaining selected for consecutive edits. Color changes are verified when applied so the chosen color is used on the first action. `Clip Duration` stays specific to each clip and is not copied with visual styles. Standard video and audio clips are ignored by its automatic selection refresh.
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

### 1.3.0 - 2026-09-04

- New: Translation tab for translate mogrt or native subtiles with DeepL ([free API key required](https://www.deepl.com/your-account/keys)).
- New: Introduces a persistent Whisper dictionary shared across Premiere projects.
- Added signed and Apple-notarized Apple Silicon macOS installer builds.
- Applying visual editor settings no longer change the animation duration of the selected MOGRTs.
- Better and faster reading for batch mogrt in visual editor.

### 1.2.0 - 2026-08-10

- Added complete Windows and Apple Silicon macOS installers, so Python, FFmpeg, Whisper, and WhisperX do not need to be installed separately.
- Improved subtitle generation with clearer live progress, richer diagnostics, separate Whisper language selection, and optional punctuation removal that preserves apostrophes in words such as `J'aime` and `I'm`.
- Improved animated MOGRT workflows with more reliable frame-aligned timing, subtitle rebuilding, and preservation of text styling and animation data.
- Expanded the Visual editor with automatic selection refresh, safer copying of exposed properties between subtitle clips, and faster batch application on large selections.
- Improved the Text editor, Premiere theme matching, and overall panel responsiveness during long subtitle generations.

### 1.1.0 - 2026-05-11

- Improved MOGRT generation and preserved template keyframes.
- Added the Montserrat font family to the provided resources.

### 1.0.0 - 2026-05-07

- First stable release of Sub Creator.
- Includes MOGRT generation, native Premiere subtitles, Whisper, WhisperX, visual editing, and text editing.
