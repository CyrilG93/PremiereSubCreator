# Sub Creator (Premiere Pro 2025+)

Sub Creator is a CEP panel extension for Adobe Premiere Pro 2025+ focused on dynamic/design subtitles.

It supports:
- Three source workflows:
  - SRT import via native file picker.
  - Premiere active caption track extraction (text + timing) when API exposes it.
  - Whisper local transcription from an audio/video file (CEP Node runtime first, ExtendScript fallback).
- Caption planning with max letters, max lines, style presets, uppercase, and animation mode metadata.
- MOGRT gallery with real template previews extracted from each `.mogrt` thumbnail.
- Premiere timeline insertion via ExtendScript:
  - Insert MOGRT per cue (selected from integrated gallery).
  - Apply text controls recursively (including grouped Essential Graphics properties) and map animation mode to controls like `Highlight Based On`.
  - Apply style layout controls (characters per line / max lines / font size) when exposed by the selected template.
  - Fallback to timeline markers when no MOGRT is provided.
- Interface localization (French + English, easy to extend).
- Version label in panel header + automatic GitHub release update banner.
- Persisted panel settings (source, style, limits, language, and selected MOGRT).
- Mac + Windows installers.

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

For Premiere native auto-transcription, CEP scripting still has API gaps depending on version; this extension tries to read the active caption track directly and falls back with a clear message if unavailable.
On some Premiere builds, CEP returns only `SyntheticCaption` placeholders for caption clips; in that case, use SRT or Whisper source.

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
- If multiple Python versions are installed, selects the highest compatible one (3.13 -> 3.8).
- Tries to auto-install `openai-whisper` with local Python when Python is available and version is <= 3.13.
- Adds `~/Library/Python/<version>/bin` to `~/.zprofile` and `~/.zshrc` when needed so `whisper` is in PATH.
- Skips Whisper auto-install when Python is missing or when Python version is 3.14+ (unsupported target for current package metadata).
- Tries to install `ffmpeg` via Homebrew when available.

### Windows

```bat
installers\subcreator_install_windows.bat
```

Installer behavior:
- Installs extension files to CEP.
- If multiple Python versions are installed, selects the highest compatible one (3.13 -> 3.8).
- Tries to auto-install `openai-whisper` with `py -3` or `python` when available and version is <= 3.13.
- Skips Whisper auto-install when Python is missing or when Python version is 3.14+ (unsupported target for current package metadata).
- Tries to install `ffmpeg` via `winget` when available.

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
- Single synthetic placeholders from Premiere APIs (`SyntheticCaption`) are filtered so invalid caption-track reads do not silently generate wrong subtitles.

If panel loading is blocked in development, enable CEP debug mode and restart Premiere.

## Add another language

1. Add `src/locales/<code>.json`.
2. Add the language option in `src/panel/index.html`.
3. Rebuild with `npm run subcreator:build`.

## Next recommended milestone

- Harden active caption-track extraction across Premiere versions using UXP APIs when available.
- Ship curated MOGRT packs for each preset (`clean`, `punch`, `minimal`).
- Add per-word visual emphasis controls (scale/color/blur) in UI and MOGRT parameters.
