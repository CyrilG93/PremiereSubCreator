# Windows Full installer with GitHub Actions

The `Build Windows Installer` workflow creates one complete Windows installer on a temporary GitHub-hosted Windows machine:

- `SubCreator-vX.Y.Z-Windows-Full-Installer.exe` contains the extension, private Python runtime, LGPL FFmpeg runtime, and the verified Whisper `base` model.
- The installer offers `tiny`, `base`, `small`, `medium`, and `large-v3`; `base` is already embedded and other selected models are downloaded only when missing or damaged.
- The installed private runtime stays inside the current Windows user profile and does not modify the system Python installation.
- The workflow does not build or publish Light installers, dependency updaters, or standalone runtime assets.

## Run the workflow

1. Open the repository on GitHub.
2. Open `Actions`.
3. Select `Build Windows Installer`.
4. Click `Run workflow`.
5. Leave `Publish the Full EXE` disabled for a test build, or enable it to publish the installer to the stable release matching `package.json`.

The generated Full installer remains downloadable from the workflow run for 14 days.

## Build behavior

Each clean GitHub runner:

1. Installs the exact Node.js dependencies from `package-lock.json`.
2. Runs lint, typecheck, tests, and the extension build.
3. Builds and validates the private Python and LGPL FFmpeg runtime.
4. Verifies and embeds the Whisper `base` model.
5. Compiles the Full installer with Inno Setup and verifies its product version and SHA-256.

When publication is enabled, lower semantic prereleases and their tags are removed before the stable release is created. A prerelease with the same version is promoted and reused, even if its tag uses a different letter case, so its installers are not duplicated. Technical non-semantic tags are left untouched for compatibility with previously distributed installers.

## Signing

The workflow currently creates an unsigned `.exe`, which is intentional for this release. Windows may display a SmartScreen warning before installation.
