# Sub Creator

Sub Creator is an Adobe Premiere Pro extension for creating animated subtitles from an `SRT` file, a `Whisper` transcription, or `Whisper + SRT` alignment against the active sequence audio.

It is designed to be simple to use inside Premiere:
- choose a subtitle source
- choose a MOGRT style
- generate subtitles on the timeline
- optionally fine-tune the selected MOGRTs in the built-in visual editor
- rebalance subtitle text between selected MOGRTs in the built-in text editor

## What Sub Creator can do

- Import subtitles from an `SRT` file.
- Transcribe the active Premiere sequence with `Whisper`.
- Align a corrected `.srt` or `.txt` transcript against the active Premiere sequence audio with `WhisperX` in `Whisper + SRT` mode.
- Analyze either:
  - the full active sequence
  - only the current `In/Out` range
- Insert one subtitle MOGRT per subtitle block.
- Use the included MOGRT gallery with preview thumbnails/videos.
- Let you add your own `.mogrt` files manually.
- Preserve your custom MOGRTs when you update the extension.
- Edit many exposed MOGRT parameters after generation:
  - colors
  - sliders
  - checkboxes
  - vectors/position/scale
  - font family/style/size when the template exposes them
- Edit subtitle text after generation in the `Text editor` tab:
  - move words between neighboring subtitles with drag and drop
  - split one subtitle at a selected word
  - merge subtitle blocks
  - retime rebuilt subtitles automatically

## Requirements

### Required

- Adobe Premiere Pro `2025+`
- macOS or Windows

### Required only for Whisper mode

- Python `3.8` to `3.13`
- `openai-whisper`
- `ffmpeg`
- at least one Whisper model available in the local cache

### Required only for Whisper + SRT mode

- Python `3.10` to `3.13`
- `whisperx`
- `ffmpeg`
- a corrected `.srt` or `.txt` transcript file

The installers try to configure Whisper automatically when possible.
If that fails, `SRT` mode still works normally.

## Installation

### 1. Download the release

Download the latest release `.zip` from GitHub and extract it.

The package contains:
- the extension itself
- the macOS installer
- the Windows installer
- one bundled Whisper model
- this README

### 2. Close Premiere Pro

Quit Premiere before running the installer.

### 3. Run the installer

#### macOS

Easiest method (recommended): drag and drop install_mac.sh into the Terminal window, then press Enter.

Manual method (command line): run:

```bash
./installers/subcreator_install_mac.sh
```

#### Windows

Run:

```bat
installers\subcreator_install_windows.bat
```

### 4. Reopen Premiere Pro

Then open Sub Creator from the Premiere extensions menu.

## What the installer does

The installer:
- installs the extension into Adobe CEP extensions
- enables CEP debug mode
- copies bundled Whisper models into the local Whisper cache when the release includes them
- tries to detect Python
- tries to install `openai-whisper`
- tries to install `whisperx` when the detected Python version is compatible
- tries to detect or install `ffmpeg`
- writes a local runtime config used by the extension
- preserves your custom MOGRT files during reinstall/update

If Whisper cannot be configured, the extension still installs.
In that case, `SRT` still works, while `Whisper` and `Whisper + SRT` stay unavailable until the runtime is fixed.

## First use

### SRT workflow

Use `SRT` if you already have a subtitle file.

Basic flow:
1. Open the `Creation` tab.
2. Choose `SRT` as source.
3. Select your `.srt` file.
4. Choose whether to generate:
   - `Entire sequence`
   - `In/Out points`
5. Choose a MOGRT in the gallery.
6. Click `Generate subtitles`.

When `In/Out points` is selected, only cues intersecting the current sequence range are generated.

### Whisper workflow

Use `Whisper` if you want the extension to transcribe the active sequence automatically.

Basic flow:
1. Open the `Creation` tab.
2. Choose `Whisper` as source.
3. Choose one of the Whisper models already available locally.
   If needed, use `Open Whisper models folder` to inspect or add `.pt` files.
4. Choose whether to analyze:
   - `Entire sequence`
   - `In/Out points`
5. Choose a MOGRT in the gallery.
6. Click `Generate subtitles`.

Sub Creator will:
- export the active sequence audio temporarily as `.wav`
- send it to Whisper
- use the resulting timings to generate subtitles

When `In/Out points` is selected, subtitle timing is placed back at the sequence `In` point.

### Whisper + SRT workflow

Use `Whisper + SRT` if you already have a corrected transcript and want better timings than a plain imported SRT.

Basic flow:
1. Open the `Creation` tab.
2. Choose `Whisper + SRT` as source.
3. Select a corrected `.srt` or `.txt` file.
4. Choose whether to analyze:
   - `Entire sequence`
   - `In/Out points`
5. Choose a MOGRT in the gallery.
6. Click `Generate subtitles`.

Sub Creator will:
- export the active sequence audio temporarily as `.wav`
- align the corrected text with `WhisperX`
- reuse the aligned timings to generate subtitles

Notes:
- a corrected `.srt` is recommended because it already preserves your intended subtitle blocks
- a corrected `.txt` is also supported, but it is pre-segmented heuristically before alignment
- the first corrected-align run may download one WhisperX alignment model if it is not already cached
- Whisper + SRT does not require downloading extra NLTK `punkt` data
- when `In/Out points` is selected with a corrected `.srt`, only cues intersecting that range are aligned

## Whisper models

Sub Creator only shows Whisper models already present in the local Whisper cache.

If the release package includes bundled models, the installer copies them there automatically.
That gives you at least one working model without a separate download step.

If the model dropdown is empty, read the manual section below and add a model yourself.

Useful panel behavior:
- the model dropdown shows only models already detected locally
- `Open Whisper models folder` opens the cache location used by the panel

`Whisper + SRT` does not use the Whisper model dropdown.
It depends on `whisperx` instead of a selected `.pt` model.

Common choices:
- `tiny`: fastest, lowest quality
- `base`: good default
- `small`: better quality, slower
- `medium`: heavier, better
- `large-v3`: best quality, slowest
- `turbo`: faster large-class option when available

### Add more models manually

If you want more models than the ones bundled with the installer, download them manually and place them in the Whisper cache.

Model links:
- `tiny.pt`: `https://openaipublic.azureedge.net/main/whisper/models/65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9/tiny.pt`
- `base.pt`: `https://openaipublic.azureedge.net/main/whisper/models/ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e/base.pt`
- `small.pt`: `https://openaipublic.azureedge.net/main/whisper/models/9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794/small.pt`
- `medium.pt`: `https://openaipublic.azureedge.net/main/whisper/models/345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1/medium.pt`
- `large-v3.pt`: `https://openaipublic.azureedge.net/main/whisper/models/e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb/large-v3.pt`

Cache locations:
- macOS: `~/.cache/whisper/`
- Windows: `%USERPROFILE%\.cache\whisper\`

Example:
- `base` -> `base.pt`
- `large-v3` -> `large-v3.pt`

## MOGRT gallery

Sub Creator ships with a built-in MOGRT library.

You can:
- choose styles directly from the gallery
- filter by folder/category
- open the installed MOGRT folder from the panel
- refresh the gallery without restarting Premiere

### Add your own MOGRTs

You can manually add `.mogrt` files to the installed gallery folder.

Recommended flow:
1. In the panel, click `Open MOGRT folder`.
2. Create a folder if needed.
3. Drop your `.mogrt` files inside.
4. Click `Refresh gallery`.

Top-level folders become gallery categories automatically.
Premiere-authored `.mogrt` files that expose standard `TextLayer` / `Source Text` controls can also be used for subtitle text replacement, including style-preserving text updates on common Premiere text-document payloads.

### Add preview images/videos

If your custom MOGRT has no embedded preview, you can add one next to the file.

Supported sidecar preview names:
- `<same-name>.png`
- `<same-name>.jpg`
- `<same-name>.webp`
- `<same-name>.mp4`
- `thumb.png`
- `thumb.jpg`
- `thumb.webp`
- `thumb.mp4`

### Updates and reinstalls

Your manually added MOGRT files are preserved when you reinstall or update the extension.

## Visual editor

After generation, you can switch to the `Visual editor` tab to edit selected subtitle MOGRTs already placed in the timeline.

Typical editable parameters include:
- colors
- stroke/highlight settings
- sliders
- checkboxes
- alignment
- offsets and size
- some text style fields when exposed by the template

Useful notes:
- `Live update` can apply changes while you tweak values
- multi-MOGRT apply shows progress
- logs can be collapsed or switched between compact/full modes
- long dropdowns use a custom popover that can open below, above, or centered in the viewport when the native menu would be clipped, and it stays open until you choose a value or click outside without disappearing during its own scroll
- Premiere-authored `.mogrt` files are scanned across all exposed components, so the `Visual editor` can recover more real controls than a first-component-only read
- when multiple Premiere components are exposed, the `Visual editor` rebuilds a layered hierarchy instead of flattening everything into `General`
- duplicate Premiere components are kept separate and numbered (`Group 02`, `Group 01`, `Text 02`, `Text 01`, etc.) to stay closer to Premiere's own reading order
- group components can contain nested component sections (`Text`, `Shape`, effects, extra subsections) so large Premiere templates are easier to browse
- synthetic `Settings` wrappers are collapsed, so single-subsection components open directly on their real controls
- clip-level sections such as `Motion`, `Vector Motion`, and `Opacity` are grouped together under one `Clip` section instead of being scattered through the layer tree
- when a Premiere component uses sequence-normalized coordinates, matching `Anchor Point` values are converted to the same pixel space as `Position` so the panel stays consistent
- when Premiere exposes a `Shape` component plus one low-signal nested subsection such as `Align and Transform`, Sub Creator folds that extra subsection into the parent `Shape` block
- low-signal internal controls such as generic `Property ...` and raw `Align` toggles are hidden to keep Premiere-authored templates more usable
- Premiere-only `Responsive Design` pins and internal effect metadata such as `Controls`, `Applied Version`, or sequence-size bookkeeping are hidden
- clip-level `Opacity > Blend Mode` is currently hidden because CEP does not expose a reliable label/value mapping for it and writes do not stick consistently
- Premiere internal `Parent Width / Height / Rotation` fields are hidden when they do not match the visible `Properties` panel values, so Sub Creator does not push misleading size values back onto other clips
- clip-level sections are still rendered after the layer/effect sections to better match Premiere's reading order
- on a single selected clip, `Apply changes` sends only controls you actually changed in the panel, so ambiguous Premiere-only fields are less likely to be rewritten accidentally
- some Premiere text layers still expose only an opaque runtime `Source Text` placeholder to CEP, so `font / fill / stroke` controls cannot be shown reliably until Adobe exposes more than that placeholder
- some Premiere enum labels still rely on inferred mappings because CEP exposes the numeric value but not the official option labels; when a control proves unreliable, Sub Creator hides it instead of exposing a misleading field

## Text editor

After generation, you can switch to the `Text editor` tab to fix subtitle wording and block boundaries without doing manual copy/paste and retiming on the timeline.

Current `V1` workflow:
1. Select the subtitle MOGRT clips you want to edit in the timeline.
2. Open the `Text editor` tab.
3. Click `Read selected subtitles`.
4. Edit the text directly, or drag one word to another subtitle block.
5. Use:
   - `Split at selected word`
   - `Merge previous`
   - `Merge next`
6. Click `Apply text changes`.

UI note:
- hovering `Split` or `Merge` previews the subtitle block(s) affected by that action

Important notes:
- `Text editor` currently works on selected subtitle MOGRTs from one video track at a time.
- the `Text editor` ignores selected MOGRTs that do not expose a real editable subtitle text payload
- for some Premiere-authored `.mogrt` files, the host may read back only an opaque placeholder glyph instead of the visible subtitle text; in that case, the `Text editor` reuses the generated timing metadata to recover the real text
- when those Premiere-authored templates are rebuilt from the `Text editor`, Sub Creator imports temporary baked `.mogrt` files so the visible text stays correct instead of being rewritten destructively after import
- the extension rebuilds and retimes the selected subtitle clips automatically when you apply changes
- the `Text editor` now tries to resolve the `.mogrt` file from the subtitle clips actually selected on the timeline, instead of depending only on the gallery selection
- when only one contiguous region was changed, Sub Creator rebuilds only that changed subtitle slice instead of recreating the full selection
- when several disjoint regions were changed, Sub Creator rebuilds one safe combined span that covers the edited areas instead of risking a partial multi-pass apply
- the rebuilt clips keep the full original time span of the selected subtitle range, even after merges reduce the number of blocks
- if clips remain later on the source track, Sub Creator rebuilds on a safe track above instead of risking a transient MOGRT insert that could trim or overwrite later media
- if the rebuilt subtitle timing would create one real new overlap with non-selected clips on the same track, Sub Creator rebuilds on the first empty video track above
- if no empty video track exists above, Sub Creator creates a new top video track and rebuilds there instead
- the rebuilt clips try to preserve the original MOGRT visual/text style from the selected subtitle clips
- subtitles generated by Sub Creator now keep local word-timing metadata when available, so later text edits can reuse more precise timings
- when only part of the edited range still has exact word timings, Sub Creator keeps those precise blocks anchored and only re-estimates the uncertain blocks between them
- persisted word timings can now still be reused after a safe rebuild on another track, as long as the rebuilt subtitle text and timing still match closely
- timing redistribution is heuristic in this `V1`, so it is designed for practical fixes rather than perfect word-level retiming

## Dependencies summary

### SRT only

No extra dependency is required beyond Premiere Pro.

### Whisper

Needed:
- Python `3.8` to `3.13`
- `openai-whisper`
- `ffmpeg`

Useful checks:

```bash
python3 --version
whisper --help
ffmpeg -version
```

If `whisper` is not available in your shell `PATH`, the extension also tries common fallbacks automatically.

## Troubleshooting

### Whisper option is missing in the panel

Possible reasons:
- Python is not installed
- Python version is unsupported
- Whisper is not installed
- `ffmpeg` is missing
- runtime detection failed

What to do:
1. rerun the installer
2. check Python version
3. check `whisper --help`
4. check `ffmpeg -version`

### No Whisper model appears in the panel

Possible reasons:
- no bundled model was copied during install
- you deleted the local Whisper cache files
- you installed the extension manually without the full release package

What to do:
1. rerun the installer from the full release package
2. read the `Whisper models` section above
3. copy one or more `.pt` model files into the Whisper cache

### Whisper transcription is slow

Normal.
Transcription itself can still take time depending on the model and sequence length.

### My custom MOGRT does not appear

Check:
- the file is really inside the installed MOGRT folder
- it is a valid `.mogrt`
- you clicked `Refresh gallery`
- the file is not inside an unexpected subfolder structure

### My custom MOGRT appears but only shows PREVIEW

That means no usable embedded preview was found.
Add a sidecar preview file next to the `.mogrt`.

### The extension was updated and I want to keep my own MOGRTs

That is already handled by the installer.
Your custom templates should be restored automatically after reinstall/update.

## Recommended usage

### Best quality workflow

1. Use `Whisper`
2. Analyze the active sequence or the `In/Out` range
3. Choose a MOGRT style
4. Generate subtitles
5. Fine-tune the result in the `Visual editor`

### Fastest workflow

1. Prepare a clean `SRT`
2. Import it with `SRT` mode
3. Choose a MOGRT
4. Generate subtitles

## Notes

- The extension is focused on animated/design subtitles, so the final result also depends on the selected MOGRT.
- Some MOGRTs expose more editable parameters than others.
- Some advanced font behaviors depend on how the MOGRT itself was built.

## Releases

Latest releases:
- https://github.com/CyrilG93/PremiereSubCreator/releases
