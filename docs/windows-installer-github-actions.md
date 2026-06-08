# Windows installer with GitHub Actions

The `Build Windows Installer` workflow creates the Windows `.exe` on a temporary GitHub-hosted Windows machine.

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
2. Builds the lightweight connected installer.
3. Uploads both files when release publication is enabled.
4. Commits the generated runtime tag and SHA-256 back to the selected branch.

Later runs reuse the published runtime and normally rebuild only the lightweight installer.

## Signing

The workflow currently creates unsigned `.exe` files. They work, but Windows may display a stronger SmartScreen warning. Code signing can be added later after a Windows signing certificate is available.
