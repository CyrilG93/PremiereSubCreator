# Sub Creator

Sub Creator is an Adobe Premiere extension for generating subtitle MOGRTs from an `SRT` file, a `Whisper` transcription, or `Whisper + SRT` alignment.

It is built for a simple workflow inside Premiere:
- choose a subtitle source
- choose a MOGRT style
- generate subtitle clips on the timeline
- adjust exposed controls in the `Visual editor`
- rebalance text and timing in the `Text editor`

## Included features

- `SRT`, `Whisper`, and `Whisper + SRT` creation modes
- generation on `Entire sequence` or `In/Out points`
- built-in MOGRT gallery with refresh support
- support for custom `.mogrt` files
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

The installers copy the bundled `base` Whisper model into the local Whisper cache and write a runtime config that helps the panel find Python, Whisper, WhisperX, and `ffmpeg`.

### macOS

Use the included installer script:

```bash
./installers/subcreator_install_mac.sh
```

You can also drag and drop the installer.sh in a terminal window and press enter.

### Windows

Run the included batch installer:

```bat
installers\subcreator_install_windows.bat
```

On Windows, the installer now opens in a dedicated console so the final setup summary stays visible instead of disappearing immediately.

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

Sub Creator can still read some of those values and apply them to other selected clips when the template exposes them correctly.

### Text editor

Use the `Text editor` to adjust subtitle text after generation.

You can:
- move words between neighboring subtitles
- split a subtitle block
- merge subtitle blocks
- rebuild timing after text changes

## Custom MOGRTs

Sub Creator includes a gallery, but you can also add your own `.mogrt` files.

Basic flow:
1. In the panel, click `Open MOGRT folder`.
2. Add your `.mogrt` files at the root or in subfolders.
3. Click `Refresh gallery`.

## Add more Whisper models manually

If you want more models than the ones bundled with the installer, download them manually and place them in the Whisper cache.

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

### A control is missing or behaves differently from Premiere

Use Premiere's `Properties` panel for that template. Some controls are not exposed or are only partially reliable through the CEP API.

## Support

- My website: https://www.cyrilplugin.com/
- Join my discord for any help: https://www.cyrilplugin.com/website/social/discord
