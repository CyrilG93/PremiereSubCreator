# Sub Creator (Premiere Pro 2025+)

Sub Creator is a CEP panel extension for Adobe Premiere Pro 2025+ focused on dynamic/design subtitles.

It supports:
- Three source workflows:
  - SRT import via native file picker.
  - Whisper local transcription from an audio/video file (CEP Node runtime first, ExtendScript fallback).
  - Whisper transcription from the active Premiere sequence via temporary WAV export of the current audible mix.
- Caption planning with max letters, max lines, font size, and animation mode metadata.
- MOGRT gallery with real template previews extracted from each `.mogrt` thumbnail.
- MOGRT gallery now reads installed templates dynamically from the extension `templates/mogrt` folder, including manually added `.mogrt` files and custom top-level folders.
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
- Persisted panel settings (source, style, limits, language, selected gallery folder, and selected MOGRT).
- Mac + Windows installers.
- Visual editor reads style fields from text-document payloads when available (`Font Family`, `Font Style`, `Font Size`, faux style toggles) without exposing editable caption text.
- Font controls are rendered as dropdowns when style options are discoverable in MOGRT payloads, and `Font Style` options are filtered by selected `Font Family` when family/style mapping is available.
- Long font dropdowns are constrained to available panel height to avoid clipped lists near the bottom of the UI.
- Gallery filter options now come from real installed MOGRT folder names instead of hardcoded aspect presets, and the panel includes buttons to open the installed MOGRT folder and refresh the gallery without restarting Premiere.
- Gallery refresh now reuses a filesystem signature cache so focus changes can skip full installed-template rescans when nothing changed.
- Manually added MOGRTs reuse embedded `.mogrt` thumbnails at runtime when present, and can also provide sidecar preview files (`<name>.png/.jpg/.webp/.mp4` or `thumb.*`) inside the same folder.
- Embedded runtime previews now reuse cached extracted thumbnail files until the source `.mogrt` changes, which reduces repeated archive extraction cost.
- Font-family apply now retries multiple token variants (`family-style`, `family`, common style aliases) and validates readback to reduce fallback-to-wrong-font behavior on some MOGRTs.
- Font-style apply now preserves the chosen family token and retries compatible style aliases when Premiere falls back to another family.
- Font-style dropdowns are now stricter for the currently selected family to reduce invalid family/style combinations.
- Startup restores panel state first, then defers heavier tasks such as Whisper runtime probing, installed MOGRT refresh, and release checks so the extension becomes interactive faster.
- The generate action now sits directly above the MOGRT gallery, and the gallery folder filter is restored after reopening Premiere instead of resetting to `All formats`.
- Subtitle generation now shows a dedicated progress bar; Whisper workflows expose stage progress (sequence export, Whisper analysis, timing parse, plan/apply), and CEP Node Whisper runs stream percentage updates when the CLI reports them.
- The panel theme now derives its neutral surface colors from Premiere CEP `appSkinInfo`, so it follows host appearance variants more closely instead of using a fixed dark-blue skin.
- On macOS, OS font fallback now uses `system_profiler` metadata (`family`, `style`, exact internal token) via an absolute system path and an enlarged CEP buffer, instead of filename guesses, which improves matching for collection fonts like `Al Bayan`, `Avenir`, and `Futura`.
- Font family/style matching now normalizes aliases such as `Al Bayan` / `AlBayan` and `Plain` / `Regular`, so cached dropdown values still resolve to the correct system font token.
- System font catalog loading is now lazy and starts when the visual editor needs font-family/style expansion, instead of blocking panel startup.
- Font token apply now retries canonical variants (`Family-Style`, compact internal ids, and cached aliases) and keeps text style dropdowns visible even when host readback is incomplete after a failed font change.
- When `Font Family` changes, the visual editor now resets `Font Style` toward neutral family defaults (`Regular`, `Book`, `Roman`, `Plain`, `Medium`, `Semibold`) instead of silently reusing the previous family style, which reduces wrong-token fallbacks such as inherited `Bold`.
- The visual editor now keeps a `Font Style` control available whenever a `Font Family` control exists, using host data first and local system styles as fallback so family changes also write an explicit style.
- When CEP Node is available, the visual editor augments font dropdowns with local OS-installed font families/styles (macOS/Windows font directories) as a fallback when MOGRT options are limited.
- Visual property reads now load the OS font catalog only when the current selection actually exposes font family/style controls.
- Faux style toggles enforce Premiere-like exclusivity for `All Caps` and `Small Caps`.
- Faux `Bold` / `Italic` checkboxes remain independent from `Font Style`, because Premiere exposes them as separate text-style parameters and mixing both can produce wrong font tokens.
- Visual editor apply sends current style controls so the same setup can be pushed to newly selected MOGRT clips.
- Visual editor includes optional `Live update` mode (disabled by default, persisted) to push edits while tweaking controls.
- Manual apply on multi-selection now runs with a visible progress bar (`done/total/remaining`) so long updates are trackable, and the bar stays hidden outside active multi-clip updates.
- Vector controls such as `Offset` and `Size` are normalized to sequence dimensions for readable values in the panel.
- Normalized `Position` vectors (`0..1` style values) are automatically shown in sequence pixels (for example `1920 / 1080` on 4K timelines).
- `Scale` vectors now follow the same sequence-axis basis as `Position` within a group when that MOGRT uses normalized sequence units.
- Vector values shown in visual editor are rounded to one decimal for cleaner `Position` / `Scale` editing.
- Size vectors now include 1920/1080 compatibility scaling so common subtitle templates display `100%`-style values in editor.
- Known menu-like controls (for example alignment/paragraph/based-on) are rendered as dropdowns when detected.
- Visual editor includes richer host debug payloads in the panel log for troubleshooting.
- The log panel can now be collapsed and switched between compact/full payload views without losing the raw debug entry.
- Selecting a MOGRT in the gallery no longer rebuilds all cards, and video previews only play on hover/focus instead of autoplaying everywhere.
- Color controls are detected with stricter rules to avoid rendering numeric sliders/dropdowns as color pickers.
- Packed numeric color payloads are decoded/encoded using Premiere BRG channel order for consistent read/apply in visual editor.
- Color arrays returned as `[A,R,G,B]` by Premiere are now interpreted and applied correctly in visual editor.
- Ambiguous 4-channel arrays (alpha markers on first and last slot) now default to `ARGB`, which fixes common `Stroke Color` mismatches.
- Visual editor color controls use swatch + `HEX` only and open the native CEP/browser color palette when clicking the swatch.
- Visual editor now reapplies current style values even when unchanged locally, so the same settings can be pushed to newly selected MOGRT clips.
- After visual apply, Sub Creator nudges/restores the playhead to force an immediate Program Monitor refresh for color updates.
- Slider fallback ranges now better distinguish `0..100` controls from true signed offset/position sliders.
- Opacity-like controls (`opacity` / `opacité`) are now normalized to `0..100` in visual editor even when host metadata reports a wider range.

## Important product choices

### Do we need prebuilt MOGRT files?
Yes, for premium animated design styles you should prepare MOGRT templates.

MOGRT files placed under `templates/mogrt` are auto-discovered and shown in the panel gallery.

Top-level folders under `templates/mogrt` become the gallery filter values as-is.

If a `.mogrt` contains `thumb.png` or `thumb.mp4`, Sub Creator extracts it during build and uses it as the gallery preview.

For manually added installed templates, the panel first tries to extract embedded thumbnail assets from the `.mogrt` itself, then falls back to sidecar preview files placed next to the `.mogrt` (`<same-name>.png/.jpg/.webp/.mp4` or `thumb.*`).

Without MOGRT, the panel still works but inserts markers as a safe fallback.

### Can Sub Creator list every installed font like Premiere Properties?
Not reliably with current CEP/ExtendScript APIs.

Sub Creator can apply explicit family/style names, and can show dropdown options when they are exposed by the selected MOGRT text payload, but it cannot query the full Properties font browser list directly.

### Do we need an SRT file?
SRT works immediately.

Whisper local can generate subtitles on the fly from an audio/video file.
Whisper active sequence can export the current sequence audible mix to a temporary WAV, then transcribe it automatically.
If `whisper` is not available in PATH, Sub Creator also tries common fallbacks (`python3 -m whisper`, `python -m whisper`, and user-local Whisper binaries).
If no local Whisper runtime is detected at panel startup, both Whisper source options are hidden automatically.
Installers also write a user-local runtime config (`subcreator-runtime.json`) with detected `python` / `whisper` / `ffmpeg` paths so CEP can run reliably even when host PATH is incomplete.
For temporary sequence export, Sub Creator looks for Adobe's built-in WAV system preset in either Premiere Pro or Adobe Media Encoder, so users do not need to install a custom export preset.

Whisper integration now requests `json + srt` output with `--word_timestamps True`, so caption planning can reuse precise word timings whenever Whisper provides them and only fall back to synthetic timing when needed.

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
npm run subcreator:package
```

`npm run subcreator:package` now rebuilds first and refuses to zip a stale `dist` version.

## Whisper local setup

Install Whisper CLI once on your machine:

```bash
pip install -U openai-whisper
```

Then verify:

```bash
whisper --help
```

Sub Creator uses the local Whisper CLI with:

- `--output_format all`
- `--word_timestamps True`

This allows the panel to consume word-level timing data from Whisper JSON for more accurate chunk timing.

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
- Preserves previously added files under installed `templates/mogrt` when reinstalling/updating the extension.
- Restores preserved MOGRT files without aborting when duplicate files are skipped during reinstall/update.
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
- Preserves previously added files under installed `templates/mogrt` when reinstalling/updating the extension.
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
- macOS metadata files (`.DS_Store`, `__MACOSX`, AppleDouble `._*`) are stripped from the archive.

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
- Update-banner clicks are opened through the CEP browser API instead of relying on native HTML link behavior inside Premiere.

Caption planning behavior:
- Long cues are split by contiguous word groups (not arbitrary character cuts).
- Chunk timing follows word timing boundaries when available, or proportional word distribution otherwise.
- When only cue-level timing is available, synthetic word timing is now weighted by word length and punctuation pauses instead of using flat per-word slices.
- Boundary rebalancing favors readable punctuation grouping (for example avoids starting a chunk with `time,` when a better split exists).
- Boundary rebalancing also avoids weak connector endings (for example ending a chunk with `since` when next words can absorb it).
- Boundary rebalancing also keeps short connectors after commas attached to the previous chunk when possible (for example avoids starting a chunk with `puis` / `et` alone).

If panel loading is blocked in development, enable CEP debug mode and restart Premiere.

## Add another language

1. Add `src/locales/<code>.json`.
2. Add the language option in `src/panel/index.html`.
3. Rebuild with `npm run subcreator:build`.

## Next recommended milestone

- Ship curated MOGRT packs for each preset (`clean`, `punch`, `minimal`).
- Add per-word visual emphasis controls (scale/color/blur) in UI and MOGRT parameters.
