# Sub Creator (Premiere Pro 2025+)

Sub Creator is a CEP panel extension for Adobe Premiere Pro 2025+ focused on dynamic/design subtitles.

It supports:
- Two source workflows:
  - SRT import via native file picker.
  - Whisper local transcription from an audio/video file (CEP Node runtime first, ExtendScript fallback).
- Caption planning with max letters, max lines, font size, and animation mode metadata.
- MOGRT gallery with real template previews extracted from each `.mogrt` thumbnail.
- Two UI tabs:
  - `Creation`: source -> planning -> MOGRT insertion.
  - `Visual editor`: read and apply editable Essential Graphics style parameters on selected MOGRT clips with grouped compact controls (sliders, checkboxes, colors, vectors).
- Premiere timeline insertion via ExtendScript:
  - Insert MOGRT per cue (selected from integrated gallery).
  - Apply text controls recursively (including grouped Essential Graphics properties) and map animation mode to controls like `Highlight Based On`.
  - Apply style layout controls (characters per line / max lines / font size) when exposed by the selected template.
  - Fallback to timeline markers when no MOGRT is provided.
- Interface localization (French + English, easy to extend).
- Version label in panel header + automatic GitHub release update banner.
- Persisted panel settings (source, style, limits, language, and selected MOGRT).
- Mac + Windows installers.
- Visual editor does not expose subtitle text content editing (style-only) to avoid overriding generated captions.
- Visual editor apply sends only modified controls, preserving untouched MOGRT parameters.
- Vector controls such as `Offset` and `Size` are normalized to sequence dimensions for readable values in the panel.
- Size vectors now include 1920/1080 compatibility scaling so common subtitle templates display `100%`-style values in editor.
- Known menu-like controls (for example alignment/paragraph/based-on) are rendered as dropdowns when detected.
- Visual editor includes richer host debug payloads and a `Copy logs` button for troubleshooting.
- `Copy logs` falls back to CEP runtime clipboard APIs when browser clipboard permission is denied.
- Color controls are detected with stricter rules to avoid rendering numeric sliders/dropdowns as color pickers.
- Packed numeric color payloads are decoded/encoded using Premiere BRG channel order for consistent read/apply in visual editor.
- Color arrays returned as `[A,R,G,B]` by Premiere are now interpreted and applied correctly in visual editor.
- Visual editor color controls use `HEX` + `RGB` fields and open the native CEP/browser color palette when clicking the swatch.
- Color layout calibration now keeps read-layout and write-layout caches separate to improve consistency on controls like `Stroke Color`.
- Slider fallback ranges now better distinguish `0..100` controls from true signed offset/position sliders.

## Important product choices

### Do we need prebuilt MOGRT files?
Yes, for premium animated design styles you should prepare MOGRT templates.

MOGRT files placed under `templates/mogrt` are auto-discovered and shown in the panel gallery (Landscape/Portrait/Square filters).

If a `.mogrt` contains `thumb.png` or `thumb.mp4`, Sub Creator extracts it during build and uses it as the gallery preview.

Without MOGRT, the panel still works but inserts markers as a safe fallback.

### Do we need an SRT file?
SRT works immediately.

Whisper local can generate SRT on the fly from an audio/video file.
If `whisper` is not available in PATH, Sub Creator also tries common fallbacks (`python3 -m whisper`, `python -m whisper`, and user-local Whisper binaries).
If no local Whisper runtime is detected at panel startup, the `Whisper local (audio)` source option is hidden automatically.
Installers also write a user-local runtime config (`subcreator-runtime.json`) with detected `python` / `whisper` / `ffmpeg` paths so CEP can run reliably even when host PATH is incomplete.

## Project structure

- `src/panel` CEP UI (HTML/CSS/TS).
- `src/core` subtitle parsing/planning logic.
- `src/host/SubCreatorHost.jsx` ExtendScript host bridge.
- `src/host/manifest.xml` CEP manifest.
- `src/locales` language dictionaries.
- `templates/mogrt` local MOGRT library auto-packaged into extension.
- `scripts` prefixed project commands.
- `installers` macOS + Windows install scripts.
- `Releases` local zip output folder.

## Local development

```bash
npm install
npm run subcreator:verify
npm run subcreator:install:dev
```

## Whisper local setup

Install Whisper CLI once on your machine:

```bash
pip install -U openai-whisper
```

Then verify:

```bash
whisper --help
```

Note: first Whisper transcription downloads the selected model. In enterprise/proxy environments, Python SSL trust issues can block this download (`CERTIFICATE_VERIFY_FAILED`).

### Manual Whisper model download (offline/proxy workaround)

If model download is blocked by SSL/proxy, download the model file manually and place it in Whisper cache:

- `tiny.pt`: `https://openaipublic.azureedge.net/main/whisper/models/65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9/tiny.pt`
- `base.pt`: `https://openaipublic.azureedge.net/main/whisper/models/ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e/base.pt`
- `small.pt`: `https://openaipublic.azureedge.net/main/whisper/models/9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794/small.pt`
- `medium.pt`: `https://openaipublic.azureedge.net/main/whisper/models/345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1/medium.pt`
- `large-v3.pt` (`large`): `https://openaipublic.azureedge.net/main/whisper/models/e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb/large-v3.pt`

Cache locations:

- macOS: `~/.cache/whisper/` (example: `~/.cache/whisper/base.pt`)
- Windows: `%USERPROFILE%\\.cache\\whisper\\` (example: `C:\\Users\\<you>\\.cache\\whisper\\base.pt`)

The filename must match the selected model (for example `base` -> `base.pt`).

## Commands

- `npm run subcreator:build` Build extension to `dist/com.cyrilg93.subcreator`.
- `npm run subcreator:lint` Run ESLint.
- `npm run subcreator:test` Run unit tests.
- `npm run subcreator:verify` Run lint + tests + build.
- `npm run subcreator:install:dev` Install build in CEP extensions directory.
- `npm run subcreator:package` Build local release zip in `Releases/`.

## Installers

### macOS

```bash
./installers/subcreator_install_mac.sh
```

Installer behavior:
- Installs extension files to CEP.
- Enables CEP debug mode by default for CSXS.7 -> CSXS.12.
- If multiple Python versions are installed, selects the highest compatible one (3.13 -> 3.8).
- Tries to auto-install `openai-whisper` with local Python when Python is available and version is <= 3.13.
- Adds `~/Library/Python/<version>/bin` to `~/.zprofile` and `~/.zshrc` when needed so `whisper` is in PATH.
- Skips Whisper auto-install when Python is missing or when Python version is 3.14+ (unsupported target for current package metadata).
- Tries to install `ffmpeg` via Homebrew when available.
- Writes runtime config to `~/Library/Application Support/SubCreator/subcreator-runtime.json`.

### Windows

```bat
installers\subcreator_install_windows.bat
```

Installer behavior:
- Installs extension files to CEP.
- Enables CEP debug mode by default for CSXS.7 -> CSXS.12.
- If multiple Python versions are installed, selects the highest compatible one (3.13 -> 3.8).
- Tries to auto-install `openai-whisper` with `py -3` or `python` when available and version is <= 3.13.
- Skips Whisper auto-install when Python is missing or when Python version is 3.14+ (unsupported target for current package metadata).
- Tries to install `ffmpeg` via `winget` when available.
- Writes runtime config to `%APPDATA%\\SubCreator\\subcreator-runtime.json`.

## Release packaging

The release command creates a zip in `Releases/` and includes only mandatory files:
- `README.md`
- `installers/subcreator_install_mac.sh`
- `installers/subcreator_install_windows.bat`
- `dist/com.cyrilg93.subcreator/*`

```bash
npm run subcreator:package
```

## CEP notes

- Extension id: `com.cyrilg93.subcreator`
- Host: Premiere Pro `PPRO [25.0,99.9]`
- Runtime: CSXS 11

Track behavior in panel:
- MOGRT subtitles are inserted on an empty top video target (reuse existing empty top track, otherwise create one).
- Track selection avoids signature-based ambiguity and always targets the highest empty video track after creation.
- Audio track index is handled internally for Premiere `importMGT` compatibility.
- Update banner checks `https://api.github.com/repos/CyrilG93/PremiereSubCreator/releases/latest` and displays only when a newer version exists.

Caption planning behavior:
- Long cues are split by contiguous word groups (not arbitrary character cuts).
- Chunk timing follows word timing boundaries when available, or proportional word distribution otherwise.
- Boundary rebalancing favors readable punctuation grouping (for example avoids starting a chunk with `time,` when a better split exists).
- Boundary rebalancing also avoids weak connector endings (for example ending a chunk with `since` when next words can absorb it).

If panel loading is blocked in development, enable CEP debug mode and restart Premiere.

## Add another language

1. Add `src/locales/<code>.json`.
2. Add the language option in `src/panel/index.html`.
3. Rebuild with `npm run subcreator:build`.

## Next recommended milestone

- Ship curated MOGRT packs for each preset (`clean`, `punch`, `minimal`).
- Add per-word visual emphasis controls (scale/color/blur) in UI and MOGRT parameters.
