import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const installerScriptPath = fileURLToPath(
  new URL("../installers/subcreator_install_macos_private_runtime.sh", import.meta.url)
);
const packagingSourcePath = fileURLToPath(new URL("../scripts/subcreator-package-macos-pkg.mjs", import.meta.url));
const runtimeManifestPath = fileURLToPath(new URL("../installers/macos-runtime.json", import.meta.url));

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
