# Sub Creator

Sub Creator is an Adobe Premiere extension for generating subtitle MOGRTs from an `SRT` file, a `Whisper` transcription, or `Whisper + SRT` alignment.

It is built for a simple workflow inside Premiere:
- choose a subtitle source
- choose a MOGRT in the gallery
- generate subtitle clips on the timeline
- adjust exposed controls in the `Visual editor`
- rebalance text and timing in the `Text editor`

## Included features

- `SRT`, `Whisper`, and `Whisper + SRT` creation modes
- generation on `Entire sequence` or `In/Out points`, with automatic fallback to the full sequence when no valid In/Out range exists
- bundled base MOGRT templates with gallery support for your own templates
- searchable MOGRT gallery with remembered folder/view state between sessions
- generation controls for max letters, max words, and max lines per subtitle
- better word grouping for apostrophes and hyphenated words such as `peut-etre`
- support for custom `.mogrt` files from after effect and premiere
- `Visual editor` for exposed controls
- `Text editor` for subtitle text redistribution and timing rebuild
- `Stop current job` for active Whisper and Whisper + SRT processing

## Platforms

- Adobe Premiere Pro `2025+`
- macOS
- Windows

## Dependencies

Sub Creator does not need the same dependencies for every mode.

| Mode | Extra dependencies |
| --- | --- |
| `SRT only` | none beyond Premiere |
| `Whisper` | Python `3.8` to `3.13`, `ffmpeg`, `openai-whisper`, and at least one local Whisper model |
| `Whisper + SRT` | Python `3.10` to `3.13`, `ffmpeg`, `whisperx`, and a corrected `.srt` or `.txt` transcript |

The installers try to configure the Python tools automatically when possible.
If Whisper cannot be configured, `SRT` mode still works.
If `Whisper + SRT` still looks unavailable on Windows after install, rerun the Windows installer once Python `3.11` or `3.12` is confirmed and let it finish until the runtime summary is printed.

## Installation

### What you need

If you only use `SRT` mode:
- just install the extension and use it

If you want `Whisper` or `Whisper + SRT`:
- install Python `3.11` or `3.12`
- install `ffmpeg`

Download links:
- Python: [python.org/downloads](https://www.python.org/downloads/)
- FFmpeg: [ffmpeg.org/download.html](https://ffmpeg.org/download.html)
- Python for Windows: [python.org/downloads/windows](https://www.python.org/downloads/windows/)
- Python for macOS: [python.org/downloads/macos](https://www.python.org/downloads/macos/)

### Important things to check

Windows:
- during Python installation, enable `Add python.exe to PATH`
- if the installer offers the `py launcher`, keep it enabled
- if `python` opens the Microsoft Store instead of Python, disable the Windows `App execution aliases` for Python
- if you install Python or `ffmpeg` after Sub Creator was already installed, run the Sub Creator installer again

macOS:
- install Python first if you want to use `Whisper` or `Whisper + SRT`
- if `ffmpeg` is not already available on your Mac, install it before rerunning the Sub Creator installer

For `ffmpeg`, use the download page above and choose the package for your system.
If the `ffmpeg` installer asks to add it to `PATH`, accept it.

1. Download the latest release `.zip` from GitHub.
2. Close Premiere Pro.
3. Run the included installer for your OS.
4. Let the installer finish and print its runtime summary.
5. Reopen Premiere Pro.
6. Open `Sub Creator` from the Premiere extensions panel.

The installers copy the bundled `base` Whisper model into the local Whisper cache and write a config file that helps the panel find Python, Whisper, WhisperX, and `ffmpeg`.
If an older Sub Creator install folder is found, the installer migrates it to the current `com.cyrilplugin.subcreator` folder name while preserving user-added MOGRT templates.
The release package can also include a `Fonts` folder with the font archives used by the bundled templates, so you can install matching fonts manually when needed.
The GitHub repository can also contain an `AE` folder with After Effects source projects for bundled templates, but that folder is not included in release `.zip` files.

### macOS

Use the included installer script:

```bash
./subcreator_install_mac.sh
```

You can also drag and drop `subcreator_install_mac.sh` into a Terminal window and press Enter.

### Windows

Run the included batch installer:

```bat
subcreator_install_windows.bat
```

On Windows, the installer now opens in a dedicated console so the final setup summary stays visible instead of disappearing immediately.
It also enables `PlayerDebugMode` for a wide `CSXS` range and checks that the registry keys were written, which helps when Premiere would otherwise keep `Window > Extensions` unavailable.

### Simple recommended order for beginners

If you want `Whisper` or `Whisper + SRT`, this is the easiest order:

1. Install Python from the official Python website.
2. Install `ffmpeg`.
3. Run the Sub Creator installer.
4. Reopen Premiere Pro.

If `SRT` is the only mode you need, you can skip Python and `ffmpeg`.

## Basic workflows

### SRT

Use this when you already have a subtitle file.

1. Open `Creation`.
2. Choose `SRT`.
3. Select your `.srt` file.
4. Choose `Entire sequence` or `In/Out points`.
5. Choose a MOGRT.
6. Click `Generate subtitles`.

### Whisper

Use this when you want Sub Creator to transcribe the active sequence audio.

1. Open `Creation`.
2. Choose `Whisper`.
3. Select a local Whisper model.
4. Choose a `Whisper language`, or leave `Auto detect`.
5. Choose `Entire sequence` or `In/Out points`.
6. Choose a MOGRT.
7. Click `Generate subtitles`.

Sub Creator exports a temporary WAV from the active sequence, runs Whisper, then creates subtitle MOGRTs on the timeline.

If you start the wrong analysis, use `Stop current job`.

### Whisper + SRT

Use this when you already have corrected text and want better timings than a plain imported SRT.

1. Open `Creation`.
2. Choose `Whisper + SRT`.
3. Select a corrected `.srt` or `.txt`.
4. Choose the explicit language used in the corrected transcript.
5. Choose `Entire sequence` or `In/Out points`.
6. Choose a MOGRT.
7. Click `Generate subtitles`.

Sub Creator exports a temporary WAV from the active sequence, aligns the corrected text with WhisperX, then creates the subtitle MOGRTs.

If you start the wrong analysis, use `Stop current job`.

## Editing generated subtitles

### Visual editor

Use the `Visual editor` to read exposed MOGRT properties from the selected clips and apply the same changes to the rest of the selection.

If you want a safer one-to-many transfer, select a source clip, click `Copy properties`, select the destination clip(s), then click `Apply changes`. `Copy properties` now rereads the current selection automatically before storing the snapshot. This is especially useful for fonts changed directly in Premiere's `Properties` panel, because Sub Creator now reuses the exact source font token when Premiere exposes it and also tries to keep the matching text style in sync.

Typical examples:
- sizes
- positions
- sliders
- checkboxes
- colors
- many exposed template controls

Important:
- not every MOGRT exposes every control to Premiere
- some controls are more reliable in Premiere's `Properties` panel than in the `Visual editor`
- for fonts, colors, and some template-specific controls, changing them directly in `Properties` is often the safest and fastest option
- faux-style text toggles such as `Bold`, `Italic`, `All Caps`, and `Small Caps` only appear when the MOGRT actually exposes them as editable controls
- when a template exposes 2D box or offset controls, Sub Creator aims to keep the same pixel units shown in Premiere's `Properties` panel
- for some After Effects MOGRTs with duplicated internal controls, Sub Creator avoids writing the same values twice so generated animations stay stable after reopening Premiere
- for animated After Effects MOGRTs, Sub Creator writes animation/layout controls first and text last to better preserve the final saved animation state
- for animated After Effects MOGRTs, Sub Creator also keeps common highlight dropdown values aligned with the template menu so generated clips export more reliably

Sub Creator can still read some of those values and apply them to other selected clips when the template exposes them correctly.

### Text editor

Use the `Text editor` to adjust subtitle text after generation.

You can:
- move words between neighboring subtitles
- split a subtitle block
- merge subtitle blocks
- rebuild timing after text changes

When the template exposes those controls cleanly, the rebuilt clips also keep the source clip visual settings instead of going back to the raw template defaults.

## Custom MOGRTs

Sub Creator ships with bundled base templates, and you can also add your own `.mogrt` files.

Basic flow:
1. In the panel, click `Open MOGRT folder`.
2. Add your `.mogrt` files at the root or in subfolders.
3. Click `Refresh gallery`.

Sub Creator keeps your original `.mogrt` filenames in the installed template folder, so updates do not silently rename bundled or custom templates.

For simple After Effects test templates, you can expose a slider named `Clip Duration`.
Sub Creator will fill it automatically with each generated subtitle clip duration, which helps AE expressions adapt to the real timeline length without breaking older custom templates that already work.

## Add more Whisper models manually

If you want more models than the ones bundled with the installer, download them manually and place them in the Whisper cache.

## Which Whisper model should you choose?

If you are not sure, start with `base` or `small`.

- `tiny`: fastest option, but lowest accuracy
- `base`: best simple starting point for most users
- `small`: better quality than `base`, still reasonable on a good computer
- `medium`: better for harder audio, accents, noise, or more demanding projects
- `large-v3`: best accuracy, but much slower and heavier
- `turbo`: optimized for speed, useful when you want faster results and can accept that maximum accuracy is not the priority

Simple recommendations:

- quick tests or weak computer: `tiny` or `base`
- everyday subtitle work: `base` or `small`
- difficult audio or better quality needed: `medium`
- best possible transcription quality: `large-v3`
- speed first: `turbo`

Language tips:

- For English-only audio, English-only Whisper models can be more efficient.
- For multilingual audio, use the normal multilingual models.
- If results are inconsistent, set the `Whisper language` manually instead of leaving `Auto detect`.

Important:

- Bigger models are usually more accurate, but they need more time and more memory.
- Whisper can still hallucinate or make mistakes, especially on low-resource languages, noisy audio, or unclear speech.
- The best model depends on your audio quality, your language, and your computer speed.

This guidance is based on:

- OpenAI Whisper model card: [https://github.com/openai/whisper/blob/main/model-card.md](https://github.com/openai/whisper/blob/main/model-card.md)
- Whisper model overview: [https://whisper-api.com/blog/models/](https://whisper-api.com/blog/models/)

Model links:
- `tiny.pt`: [https://openaipublic.azureedge.net/main/whisper/models/65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9/tiny.pt](https://openaipublic.azureedge.net/main/whisper/models/65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9/tiny.pt)
- `base.pt`: [https://openaipublic.azureedge.net/main/whisper/models/ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e/base.pt](https://openaipublic.azureedge.net/main/whisper/models/ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e/base.pt)
- `small.pt`: [https://openaipublic.azureedge.net/main/whisper/models/9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794/small.pt](https://openaipublic.azureedge.net/main/whisper/models/9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794/small.pt)
- `medium.pt`: [https://openaipublic.azureedge.net/main/whisper/models/345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1/medium.pt](https://openaipublic.azureedge.net/main/whisper/models/345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1/medium.pt)
- `large-v3.pt`: [https://openaipublic.azureedge.net/main/whisper/models/e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb/large-v3.pt](https://openaipublic.azureedge.net/main/whisper/models/e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb/large-v3.pt)

Cache locations:
- macOS: `~/.cache/whisper/`
- Windows: `%USERPROFILE%\.cache\whisper\`

### Compatibility notes

- After Effects-authored MOGRTs are the most predictable option.
- Premiere-authored MOGRTs are supported for subtitle generation and text editing, but some visual controls may still be better adjusted from Premiere's `Properties` panel.
- Some template-specific controls are not exposed by Adobe's API at all. If a control does not appear in the `Visual editor`, change it directly in Premiere's `Properties` panel.

If you want to author an After Effects template specifically for Sub Creator, read:

- [After Effects MOGRT compatibility guide](docs/AE_MOGRT_COMPATIBILITY_GUIDE.md)

## Known limitations

- `Whisper` and `Whisper + SRT` require local Python tools.
- Some MOGRT controls are not exposed by Premiere's API.
- Some controls may read correctly but still be more practical to change in Premiere's `Properties` panel.
- Premiere-authored MOGRTs can expose fewer reliable style controls than After Effects-authored MOGRTs.
- The `Visual editor` is best used for exposed controls that your template clearly maps to Premiere.

## Troubleshooting

### No Whisper model appears in the dropdown menu

Open `Open Whisper models folder` and check that at least one `.pt` model is present in the local Whisper cache.

### Whisper or Whisper + SRT is unavailable

Check that Python, `ffmpeg`, and the required package are installed:
- `openai-whisper` for `Whisper`
- `whisperx` for `Whisper + SRT`

If Whisper transcription quality is poor, check the `Whisper language` selector in the panel. It is separate from the UI language.

### `Window > Extensions` is greyed out on Windows

Close Premiere Pro completely, rerun `subcreator_install_windows.bat`, let it finish, then reopen Premiere.
The Windows installer enables CEP debug mode automatically for recent `CSXS` versions, which is required for unsigned CEP panels to appear.

### A control is missing or behaves differently from Premiere

Use Premiere's `Properties` panel for that template. Some controls are not exposed or are only partially reliable through the CEP API.

## Support

- My website: https://www.cyrilplugin.com/
- Join my discord for any help: https://www.cyrilplugin.com/website/social/discord

## Changelog

### 0.12.1 - 2026-04-23

- The Windows installer is more robust when detecting Python, which avoids false `Unable to parse Python version` errors on some PCs where `Python 3.11.8` was already installed correctly.
- Whisper setup no longer depends on the fragile batch-label flow that could skip installation even when a supported Python version was available.

### 0.12.0 - 2026-04-22

- The Visual Editor can now copy properties more safely between clips, with a faster `Copy properties` workflow and more reliable font transfer from Premiere's `Properties` panel.
- Font copying is more accurate for templates that expose exact font tokens, which helps avoid unwanted fallback fonts when reusing subtitle styles.
- Subtitle timing is safer when `In/Out points` are not usable, and hyphenated words stay cleaner when subtitles are generated or rebuilt.
- The bundled Base template set has been refreshed with updated `Mr Beast Style` and `TikTok Style` templates.

### 0.11.0 - 2026-04-19

- Apostrophes and punctuation are normalized more cleanly, so tokens like `c'est`, `d'accord`, `Salut!`, and `quoi?` stay attached as expected.
- The Text Editor merge/apply flow is more reliable on Windows, with creation limits kept in sync when editor changes are applied back to the timeline.
- Visual and Text Editor actions now show clearer busy states while reading or applying changes, which makes background work visible on slower machines.
- The Windows installer flow is more robust, and release archives now keep installer files at the root of the extracted folder.

### 0.10.5 - 2026-04-17

- The Visual Editor is more reliable with cleaner AE control detection, including proper dropdowns for subtitle animation options like `Highlight Mode`.
- The MOGRT gallery now better remembers the selected folder and search state between sessions.
- After Effects text-style handling is cleaner, with non-editable faux-style controls hidden when the template does not actually expose them.
- Visual refresh after apply is safer, with transient internal controls filtered out instead of appearing in the panel unexpectedly.

### 0.10.0 - 2026-04-17

- The bundled Base templates, gallery search, and remembered gallery folder/view state are more polished for daily use.
- Subtitle generation is more flexible with separate limits for max letters, max words, and max lines per subtitle.
- Windows and macOS installs are more reliable for updates, bundled models, and custom MOGRT preservation.
- Bundled and custom MOGRT filenames now stay unchanged after install, which avoids confusing renamed templates in the MOGRT folder.

### 0.9.0 - 2026-04-13

- Sub Creator now ships with a single bundled base subtitle MOGRT instead of the previous template pack.
- Custom After Effects templates are easier to support with the `Clip Duration` workflow for timeline-aware animation.
- Whisper stop handling, installer behavior, and Windows runtime detection are more reliable.
- Visual Editor and Text Editor are more stable with exposed MOGRT controls, duplicated AE parameters, and preserved layout values.
