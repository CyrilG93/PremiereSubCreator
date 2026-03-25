# Sub Creator

Sub Creator is an Adobe Premiere Pro extension for generating subtitle MOGRTs from an `SRT` file, a `Whisper` transcription, or `Whisper + SRT` alignment.

It is built for a simple workflow inside Premiere:
- choose a subtitle source
- choose a MOGRT style
- generate subtitle clips on the timeline
- adjust exposed controls in the `Visual editor`
- rebalance text and timing in the `Text editor`

## Platforms

- Adobe Premiere Pro `2025+`
- macOS
- Windows

## Dependencies

Sub Creator does not need the same dependencies for every mode.

| Mode | Extra dependencies |
| --- | --- |
| `SRT` | none beyond Premiere |
| `Whisper` | Python `3.8` to `3.13`, `ffmpeg`, `openai-whisper`, and at least one local Whisper model |
| `Whisper + SRT` | Python `3.10` to `3.13`, `ffmpeg`, `whisperx`, and a corrected `.srt` or `.txt` transcript |

The installers try to configure the Python tools automatically when possible.
If Whisper cannot be configured, `SRT` mode still works.

## Installation

1. Download the latest release `.zip` from GitHub.
2. Close Premiere Pro.
3. Run the included installer for your OS.
4. Reopen Premiere Pro.
5. Open `Sub Creator` from the Premiere extensions panel.

### macOS

Use the included installer script:

```bash
./installers/subcreator_install_mac.sh
```

### Windows

Run the included batch installer:

```bat
installers\subcreator_install_windows.bat
```

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
4. Choose `Entire sequence` or `In/Out points`.
5. Choose a MOGRT.
6. Click `Generate subtitles`.

Sub Creator exports a temporary WAV from the active sequence, runs Whisper, then creates subtitle MOGRTs on the timeline.

If you start the wrong analysis, use `Stop current job`.

### Whisper + SRT

Use this when you already have corrected text and want better timings than a plain imported SRT.

1. Open `Creation`.
2. Choose `Whisper + SRT`.
3. Select a corrected `.srt` or `.txt`.
4. Choose `Entire sequence` or `In/Out points`.
5. Choose a MOGRT.
6. Click `Generate subtitles`.

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
2. Add your `.mogrt` files.
3. Click `Refresh gallery`.

### Compatibility notes

- After Effects-authored MOGRTs are the most predictable option.
- Premiere-authored MOGRTs are supported for subtitle generation and text editing, but some visual controls may still be better adjusted from Premiere's `Properties` panel.
- Some template-specific controls are not exposed by Adobe's CEP API at all. If a control does not appear in the `Visual editor`, change it directly in Premiere.

If you want to author an After Effects template specifically for Sub Creator, read:

- [After Effects MOGRT compatibility guide](docs/AE_MOGRT_COMPATIBILITY_GUIDE.md)

## Included features

- `SRT`, `Whisper`, and `Whisper + SRT` creation modes
- generation on `Entire sequence` or `In/Out points`
- built-in MOGRT gallery with refresh support
- support for custom `.mogrt` files
- `Visual editor` for exposed controls
- `Text editor` for subtitle text redistribution and timing rebuild
- `Stop current job` for active Whisper and Whisper + SRT processing

## Known limitations

- `Whisper` and `Whisper + SRT` require local Python tools.
- Some MOGRT controls are not exposed by Premiere's API.
- Some controls may read correctly but still be more practical to change in Premiere's `Properties` panel.
- Premiere-authored MOGRTs can expose fewer reliable style controls than After Effects-authored MOGRTs.
- The `Visual editor` is best used for exposed controls that your template clearly maps to Premiere.

## Troubleshooting

### No Whisper model appears in the dropdown

Open `Open Whisper models folder` and check that at least one `.pt` model is present in the local Whisper cache.

### Whisper or Whisper + SRT is unavailable

Check that Python, `ffmpeg`, and the required package are installed:
- `openai-whisper` for `Whisper`
- `whisperx` for `Whisper + SRT`

### A control is missing or behaves differently from Premiere

Use Premiere's `Properties` panel for that template. Some controls are not exposed or are only partially reliable through the CEP API.

## Support

- GitHub releases: `https://github.com/CyrilG93/PremiereSubCreator/releases`
- GitHub repository: `https://github.com/CyrilG93/PremiereSubCreator`
