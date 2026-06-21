# Windows installer with GitHub Actions

The `Build Windows Installer` workflow creates the Windows `.exe` files on a temporary GitHub-hosted Windows machine.

- `Windows-Light-Installer.exe` contains the extension and downloads the private runtime only when it is missing.
- `Windows-Full-Installer.exe` contains the extension, private runtime, and verified `base` model for a fully offline first installation.
- Both installers offer the same Whisper model choices; the Full installer does not download its included `base` model.
- The extension can recover the standard private runtime directly if its generated configuration file is missing or unreadable.
- The extension includes its own WAV export preset, so Whisper sequence export does not depend on localized Adobe preset names.

## Run the workflow

1. Open the repository on GitHub.
2. Open `Actions`.
3. Select `Build Windows Installer`.
4. Click `Run workflow`.
5. Leave `Publish the generated EXE files` disabled for a test build, or enable it to add the files to the stable release matching `package.json`.

The generated installer remains downloadable from the workflow run for 14 days.

## First build

The first run can take much longer because the workflow checks the private Windows runtime referenced by `installers/windows-runtime.json`.

If that runtime is missing from GitHub, the workflow:

1. Builds the private Python and LGPL FFmpeg runtime.
2. Builds the lightweight connected installer and the complete installer.
3. Uploads the user installers to the matching product release and the reusable runtime to the dedicated `windows-runtime-v1` dependency release when publication is enabled.
4. Commits the generated runtime tag and SHA-256 back to the selected branch.

Later runs reuse the published runtime and normally rebuild only the lightweight installer.

For a local Windows build that must generate only the connected installer, set `SUBCREATOR_LIGHT_ONLY=1` before running `npm run subcreator:package:windows-exe`.

## Signing

The workflow currently creates unsigned `.exe` files. They work, but Windows may display a stronger SmartScreen warning. Code signing can be added later after a Windows signing certificate is available.
