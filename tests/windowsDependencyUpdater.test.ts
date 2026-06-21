import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const batchPath = fileURLToPath(new URL("../installers/subcreator_update_windows_dependencies.bat", import.meta.url));
const updaterPath = fileURLToPath(new URL("../installers/subcreator_update_windows_dependencies.ps1", import.meta.url));
const batchSource = readFileSync(batchPath, "utf8");
const updaterSource = readFileSync(updaterPath, "utf8");

describe("Windows dependency updater", () => {
  it("offers a one-click batch launcher for the PowerShell updater", () => {
    // // End users should not need to enter a PowerShell command manually.
    expect(batchSource).toContain("subcreator_update_windows_dependencies.ps1");
    expect(batchSource).toContain("-ExecutionPolicy Bypass");
    expect(batchSource).toContain("exit /b %SUBCREATOR_EXIT_CODE%");
  });

  it("verifies the published runtime before executing it", () => {
    // // The SHA-256 comparison must occur before Start-Process launches the downloaded EXE.
    const hashCheckIndex = updaterSource.indexOf("$downloadHash -ne $manifest.Sha256");
    const executeIndex = updaterSource.indexOf("Start-Process -FilePath $runtimeInstallerPath");

    expect(hashCheckIndex).toBeGreaterThan(0);
    expect(executeIndex).toBeGreaterThan(hashCheckIndex);
  });

  it("validates runtime imports and rewrites CEP config without a BOM", () => {
    // // A successful update must prove the runtime works and leave CEP with exact executable paths.
    expect(updaterSource).toContain('import whisper; import whisperx');
    expect(updaterSource).toContain("ffmpeg\\bin\\ffmpeg.exe");
    expect(updaterSource).toContain("New-Object System.Text.UTF8Encoding($false)");
  });
});
