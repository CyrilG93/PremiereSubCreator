# Sub Creator (Premiere Pro 2025+)

Sub Creator is a CEP panel extension for Adobe Premiere Pro 2025+ focused on dynamic/design subtitles.

It supports:
- Three source workflows:
  - SRT import.
  - Premiere caption files already imported in project bins (select one when multiple exist).
  - Whisper local transcription from an audio/video file.
- Caption planning with max letters, max lines, style presets, uppercase, and animation mode metadata.
- Premiere timeline insertion via ExtendScript:
  - Insert MOGRT per cue (selected from integrated gallery).
  - Fallback to timeline markers when no MOGRT is provided.
- Interface localization (French + English, easy to extend).
- Mac + Windows installers.

## Important product choices

### Do we need prebuilt MOGRT files?
Yes, for premium animated design styles you should prepare MOGRT templates.

MOGRT files placed under `templates/mogrt` are auto-discovered and shown in the panel gallery (Landscape/Portrait/Square filters).

Without MOGRT, the panel still works but inserts markers as a safe fallback.

### Do we need an SRT file?
SRT works immediately.

Whisper local can generate SRT on the fly from an audio/video file.

For Premiere native auto-transcription, CEP scripting still cannot directly trigger/read the Text panel transcript. Current Premiere workflow in this extension is: export/import captions as `.srt`, then select the desired source from the project list.

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

### Windows

```bat
installers\subcreator_install_windows.bat
```

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

Track settings in panel:
- `Video track`: destination track for inserted MOGRT (`0 = V1`, `1 = V2`, ...).
- `Audio track`: required by Premiere `importMGT` API even when template has no audio (`0 = A1`, ...).

If panel loading is blocked in development, enable CEP debug mode and restart Premiere.

## Add another language

1. Add `src/locales/<code>.json`.
2. Add the language option in `src/panel/index.html`.
3. Rebuild with `npm run subcreator:build`.

## Next recommended milestone

- Implement real active-sequence transcription provider and map words to precise cue timings.
- Ship curated MOGRT packs for each preset (`clean`, `punch`, `minimal`).
- Add per-word visual emphasis controls (scale/color/blur) in UI and MOGRT parameters.
