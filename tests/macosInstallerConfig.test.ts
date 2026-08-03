import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const installerScriptPath = fileURLToPath(
  new URL("../installers/subcreator_install_macos_private_runtime.sh", import.meta.url)
);
const packagingSourcePath = fileURLToPath(new URL("../scripts/subcreator-package-macos-pkg.mjs", import.meta.url));
const runtimeManifestPath = fileURLToPath(new URL("../installers/macos-runtime.json", import.meta.url));
const workflowPath = fileURLToPath(new URL("../.github/workflows/build-macos-installer.yml", import.meta.url));

describe("macOS installer runtime recovery", () => {
  it("validates Whisper imports and FFmpeg before keeping an installed runtime", () => {
    // // A matching version marker alone must not preserve a damaged private runtime.
    const installerScript = readFileSync(installerScriptPath, "utf8");

    expect(installerScript).toContain('-c "import whisper; import whisperx"');
    expect(installerScript).toContain('/ffmpeg/bin/ffmpeg" -version');
    expect(installerScript).toContain('subcreator_install_bundled_runtime "${SUBCREATOR_RUNTIME_DIR}"');
    expect(installerScript).toContain("Bundled private runtime is missing");
    expect(installerScript).not.toContain("SUBCREATOR_RUNTIME_URL");
  });
});

describe("macOS Full package architecture", () => {
  it("supports Apple Silicon arm64 only", () => {
    // // Public macOS packaging must not expose Intel or connected-update variants.
    const packagingSource = readFileSync(packagingSourcePath, "utf8");
    const runtimeManifest = JSON.parse(readFileSync(runtimeManifestPath, "utf8")) as {
      assets?: Record<string, unknown>;
    };

    expect(packagingSource).toContain('const macArch = "arm64"');
    expect(packagingSource).toContain("macOS PKG packaging requires an Apple Silicon arm64 Mac.");
    expect(packagingSource).not.toMatch(/x86_64|SUBCREATOR_MAC_CONNECTED_ONLY|createConnectedPackage/);
    expect(Object.keys(runtimeManifest.assets ?? {})).toEqual(["arm64"]);
  });

  it("builds only the Full ARM64 PKG in GitHub Actions", () => {
    // // Keep the CI entry point aligned with the supported Apple Silicon-only public package.
    const workflow = readFileSync(workflowPath, "utf8");

    expect(workflow).toContain("runs-on: macos-15");
    expect(workflow).toContain('test "$(uname -m)" = "arm64"');
    expect(workflow).toContain("SUBCREATOR_REBUILD_RUNTIME: \"1\"");
    expect(workflow).toContain("SubCreator-v${{ steps.release.outputs.version }}-macOS-Installer-arm64");
    expect(workflow).not.toMatch(/Light|x86_64-Installer|Intel-Installer|Updater|\.zip/);
  });
});

describe("macOS installer fonts", () => {
  it("keeps identical fonts instead of rewriting them", () => {
    // // Hash comparison avoids unnecessary font replacement during repeated Full installations.
    const installerScript = readFileSync(installerScriptPath, "utf8");

    expect(installerScript).toContain('shasum -a 256 "${font_path}"');
    expect(installerScript).toContain('shasum -a 256 "${target_path}"');
    expect(installerScript).toContain("Kept ${skipped} identical bundled font(s) already installed.");
  });
});

describe("macOS installer Whisper model selection", () => {
  it("preselects models found in the current user's Whisper cache", () => {
    // // The Installer can resolve the graphical console user before the customization screen opens.
    const packagingSource = readFileSync(packagingSourcePath, "utf8");

    expect(packagingSource).toContain('selected="subcreatorSelectPreviouslyInstalledModel');
    expect(packagingSource).toContain('require-scripts="true"');
    expect(packagingSource).toContain("system.ioregistry.fromPath('IOService:/')");
    expect(packagingSource).toContain("system.env && system.env.USER");
    expect(packagingSource).toContain("kCGSSessionUserNameKey");
    expect(packagingSource).toContain("/.cache/whisper/");
    expect(packagingSource).toContain("system.files.fileExistsAtPath");
    expect(packagingSource).toContain("my.choice.packageUpgradeAction");
    expect(packagingSource).toContain("return true;");
  });
});
