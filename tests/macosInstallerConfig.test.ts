import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const installerScriptPath = fileURLToPath(
  new URL("../installers/subcreator_install_macos_private_runtime.sh", import.meta.url)
);
const updateScriptPath = fileURLToPath(new URL("../installers/subcreator_update_macos.sh", import.meta.url));

describe("macOS installer runtime recovery", () => {
  it("validates Whisper imports and FFmpeg before keeping an installed runtime", () => {
    // // A matching version marker alone must not preserve a damaged private runtime.
    const installerScript = readFileSync(installerScriptPath, "utf8");

    expect(installerScript).toContain('-c "import whisper; import whisperx"');
    expect(installerScript).toContain('/ffmpeg/bin/ffmpeg" -version');
    expect(installerScript).toContain('subcreator_download_runtime "${SUBCREATOR_RUNTIME_DIR}"');
  });
});

describe("macOS installer font updates", () => {
  it.each([installerScriptPath, updateScriptPath])("keeps identical fonts instead of rewriting them in %s", (scriptPath) => {
    // // Hash comparison avoids unnecessary font replacement during full installs and lightweight updates.
    const installerScript = readFileSync(scriptPath, "utf8");

    expect(installerScript).toContain('shasum -a 256 "${font_path}"');
    expect(installerScript).toContain('shasum -a 256 "${target_path}"');
    expect(installerScript).toContain("Kept ${skipped} identical bundled font(s) already installed.");
  });
});
