# macOS Full installer with GitHub Actions

The `Build macOS Installer` workflow creates one complete Apple Silicon installer on a temporary GitHub-hosted ARM64 Mac:

- `SubCreator-vX.Y.Z-macOS-Installer-arm64.pkg` contains the extension, private Python runtime, and LGPL FFmpeg runtime.
- The installer offers `tiny`, `base`, `small`, `medium`, and `large-v3`; selected models are downloaded only when missing or damaged.
- The private runtime stays inside the current macOS user profile and does not modify the system Python installation.
- The workflow does not build or publish Light installers, dependency updaters, ZIP files, or Intel packages.

## Run the workflow

1. Open the repository on GitHub.
2. Open `Actions`.
3. Select `Build macOS Installer`.
4. Click `Run workflow`.
5. Leave `Publish the Full ARM64 PKG` disabled for a test build, or enable it to publish the installer to the stable release matching `package.json`.

The generated Full installer remains downloadable from the workflow run for 14 days.

## Build behavior

Each GitHub runner:

1. Confirms that its processor architecture is ARM64.
2. Installs the exact Node.js dependencies from `package-lock.json`.
3. Runs lint, typecheck, tests, and the extension build.
4. Rebuilds or reuses the validated private Python and LGPL FFmpeg runtime.
5. Creates the Full PKG, expands it to verify its structure, checks the embedded version and ARM64 declaration, and calculates its SHA-256.

When publication is enabled, lower semantic prereleases and their tags are removed before the stable release is created. A prerelease with the same version is promoted and reused, even if its tag uses a different letter case, so its installers are not duplicated. Technical non-semantic tags are left untouched.

## Signing

The workflow currently creates an unsigned `.pkg`, which is intentional for this release. macOS may require the package to be opened from its context menu.
